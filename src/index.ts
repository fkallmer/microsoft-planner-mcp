import { FastMCP } from "fastmcp";
import { z } from "zod";
import { execSync } from "child_process";
import { randomUUID } from "crypto";
import { writeFileSync, unlinkSync, readFileSync } from "fs";
import { tmpdir } from "os";
import { join, basename } from "path";

const mcp = new FastMCP({
  name: "microsoft-planner-mcp",
  version: "1.0.0",
});

// Helper to execute az rest commands
function azRest(method: string, url: string, body?: object): string {
  const args = [`az rest --method ${method} --url "${url}"`];
  if (body) {
    const bodyJson = JSON.stringify(body).replace(/"/g, '\\"');
    args.push(`--headers "Content-Type=application/json" --body "${bodyJson}"`);
  }
  try {
    const result = execSync(args.join(" "), { encoding: "utf-8", maxBuffer: 10 * 1024 * 1024 });
    return result;
  } catch (error: any) {
    throw new Error(`az rest failed: ${error.message}`);
  }
}

// Helper to get ETag for update/delete operations (supports multiple resource types)
type ResourceType = "task" | "taskDetails" | "bucket";

function getETag(resourceType: ResourceType, resourceId: string): string {
  const urlMap: Record<ResourceType, string> = {
    task: `https://graph.microsoft.com/v1.0/planner/tasks/${resourceId}`,
    taskDetails: `https://graph.microsoft.com/v1.0/planner/tasks/${resourceId}/details`,
    bucket: `https://graph.microsoft.com/v1.0/planner/buckets/${resourceId}`,
  };
  const result = JSON.parse(azRest("GET", urlMap[resourceType]));
  return result["@odata.etag"];
}

// Helper to get groupId from a plan (required for comments and group member listing)
function getGroupIdFromPlan(planId: string): string {
  const url = `https://graph.microsoft.com/v1.0/planner/plans/${planId}`;
  const result = JSON.parse(azRest("GET", url));
  return result.container.containerId;
}

// Helper to encode URL for reference keys (Graph API requires specific encoding)
// Per Microsoft docs: encode colons and dots, but NOT forward slashes
// Example: https://github.com → https%3A//github%2Ecom
function encodeUrlForReference(url: string): string {
  return url
    .replace(/:/g, "%3A")     // Encode colons
    .replace(/\./g, "%2E")    // Encode dots
    .replace(/ /g, "%20")     // Encode spaces
    .replace(/#/g, "%23")     // Encode hash
    .replace(/\?/g, "%3F")    // Encode question mark
    .replace(/&/g, "%26")     // Encode ampersand
    .replace(/=/g, "%3D");    // Encode equals
}

// Helper for PATCH requests with complex JSON bodies (uses temp file to avoid shell escaping issues)
function azRestPatchWithEtag(url: string, etag: string, body: object): string {
  const tmpFile = join(tmpdir(), `mcp-planner-${randomUUID()}.json`);
  try {
    writeFileSync(tmpFile, JSON.stringify(body));
    const escapedEtag = etag.replace(/"/g, '\\"');
    const args = [
      `az rest --method PATCH --url "${url}"`,
      `--headers "Content-Type=application/json" "If-Match=${escapedEtag}"`,
      `--body @${tmpFile}`,
    ];
    const result = execSync(args.join(" "), { encoding: "utf-8" });
    return result;
  } finally {
    try { unlinkSync(tmpFile); } catch {}
  }
}

// Tool: List tasks for a plan
mcp.addTool({
  name: "list-tasks",
  description: "List all tasks in a Planner plan",
  parameters: z.object({
    planId: z.string().describe("The Planner plan ID"),
  }),
  execute: async ({ planId }) => {
    const url = `https://graph.microsoft.com/v1.0/planner/plans/${planId}/tasks`;
    const result = JSON.parse(azRest("GET", url));
    return JSON.stringify(result.value, null, 2);
  },
});

// Tool: Get single task
mcp.addTool({
  name: "get-task",
  description: "Get details of a specific Planner task",
  parameters: z.object({
    taskId: z.string().describe("The task ID"),
  }),
  execute: async ({ taskId }) => {
    const url = `https://graph.microsoft.com/v1.0/planner/tasks/${taskId}`;
    const result = azRest("GET", url);
    return result;
  },
});

// Tool: Get task details (description, checklist, references)
mcp.addTool({
  name: "get-task-details",
  description: "Get extended task details including description and checklist",
  parameters: z.object({
    taskId: z.string().describe("The task ID"),
  }),
  execute: async ({ taskId }) => {
    const url = `https://graph.microsoft.com/v1.0/planner/tasks/${taskId}/details`;
    const result = azRest("GET", url);
    return result;
  },
});

// Tool: Create task
mcp.addTool({
  name: "create-task",
  description: "Create a new task in a Planner plan",
  parameters: z.object({
    planId: z.string().describe("The plan ID"),
    bucketId: z.string().describe("The bucket ID"),
    title: z.string().describe("Task title"),
  }),
  execute: async ({ planId, bucketId, title }) => {
    const url = "https://graph.microsoft.com/v1.0/planner/tasks";
    const body = { planId, bucketId, title };
    const result = azRest("POST", url, body);
    return result;
  },
});

// Tool: Update task (title, percentComplete, assignments, categories, priority)
mcp.addTool({
  name: "update-task",
  description: "Update task properties (title, progress, assignments, categories, due date, priority). Auto-fetches ETag.",
  parameters: z.object({
    taskId: z.string().describe("The task ID"),
    title: z.string().optional().describe("New title"),
    percentComplete: z.number().min(0).max(100).optional().describe("Progress 0-100"),
    assignUserId: z.string().optional().describe("User ID to assign"),
    category: z.string().optional().describe("Category to apply (category1-category25)"),
    removeCategory: z.string().optional().describe("Category to remove (category1-category25)"),
    dueDateTime: z.string().optional().describe("Due date (ISO 8601 format, e.g., '2024-12-31' or '2024-12-31T17:00:00Z'). Use 'clear' to remove due date."),
    priority: z.number().min(0).max(10).optional().describe("Priority: 1=urgent, 3=important, 5=medium, 9=low"),
  }),
  execute: async ({ taskId, title, percentComplete, assignUserId, category, removeCategory, dueDateTime, priority }) => {
    const etag = getETag("task", taskId);
    const body: Record<string, any> = {};
    if (title !== undefined) body.title = title;
    if (percentComplete !== undefined) body.percentComplete = percentComplete;
    if (priority !== undefined) body.priority = priority;
    if (assignUserId) {
      body.assignments = {
        [assignUserId]: {
          "@odata.type": "#microsoft.graph.plannerAssignment",
          orderHint: " !",
        },
      };
    }
    if (category || removeCategory) {
      body.appliedCategories = {};
      if (category) body.appliedCategories[category] = true;
      if (removeCategory) body.appliedCategories[removeCategory] = null;
    }
    if (dueDateTime) {
      if (dueDateTime.toLowerCase() === "clear") {
        body.dueDateTime = null;
      } else {
        // If just a date is provided (YYYY-MM-DD), append time for end of day UTC
        body.dueDateTime = dueDateTime.includes("T") ? dueDateTime : `${dueDateTime}T23:59:59Z`;
      }
    }

    const url = `https://graph.microsoft.com/v1.0/planner/tasks/${taskId}`;
    // Escape the inner quotes in the ETag for shell: W/"..." -> W/\"...\"
    const escapedEtag = etag.replace(/"/g, '\\"');
    const args = [
      `az rest --method PATCH --url "${url}"`,
      `--headers "Content-Type=application/json" "If-Match=${escapedEtag}"`,
      `--body '${JSON.stringify(body)}'`,
    ];
    try {
      const result = execSync(args.join(" "), { encoding: "utf-8" });
      return result || "Task updated successfully";
    } catch (error: any) {
      throw new Error(`Update failed: ${error.message}`);
    }
  },
});

// Tool: Update task details (description with GitHub links)
mcp.addTool({
  name: "update-task-details",
  description: "Update task description (use for GitHub links). Auto-fetches ETag.",
  parameters: z.object({
    taskId: z.string().describe("The task ID"),
    description: z.string().describe("Task description (supports markdown, include GitHub URLs)"),
  }),
  execute: async ({ taskId, description }) => {
    const etag = getETag("taskDetails", taskId);
    const url = `https://graph.microsoft.com/v1.0/planner/tasks/${taskId}/details`;
    // Escape the inner quotes in the ETag for shell: W/"..." -> W/\"...\"
    const escapedEtag = etag.replace(/"/g, '\\"');
    const args = [
      `az rest --method PATCH --url "${url}"`,
      `--headers "Content-Type=application/json" "If-Match=${escapedEtag}"`,
      `--body '${JSON.stringify({ description })}'`,
    ];
    try {
      const result = execSync(args.join(" "), { encoding: "utf-8" });
      return result || "Task details updated successfully";
    } catch (error: any) {
      throw new Error(`Update details failed: ${error.message}`);
    }
  },
});

// Tool: Add checklist item
mcp.addTool({
  name: "add-checklist-item",
  description: "Add a checklist item (subtask) to a Planner task",
  parameters: z.object({
    taskId: z.string().describe("The task ID"),
    title: z.string().describe("Checklist item title"),
    isChecked: z.boolean().optional().default(false).describe("Whether the item is checked"),
  }),
  execute: async ({ taskId, title, isChecked }) => {
    const etag = getETag("taskDetails", taskId);
    const itemId = randomUUID();
    const body = {
      checklist: {
        [itemId]: {
          "@odata.type": "#microsoft.graph.plannerChecklistItem",
          title,
          isChecked,
        },
      },
    };

    const url = `https://graph.microsoft.com/v1.0/planner/tasks/${taskId}/details`;
    const escapedEtag = etag.replace(/"/g, '\\"');
    const args = [
      `az rest --method PATCH --url "${url}"`,
      `--headers "Content-Type=application/json" "If-Match=${escapedEtag}"`,
      `--body '${JSON.stringify(body)}'`,
    ];
    try {
      const result = execSync(args.join(" "), { encoding: "utf-8" });
      return result || JSON.stringify({ success: true, itemId, title });
    } catch (error: any) {
      throw new Error(`Add checklist item failed: ${error.message}`);
    }
  },
});

// Tool: Add multiple checklist items at once
mcp.addTool({
  name: "add-checklist-items",
  description: "Add multiple checklist items (subtasks) to a Planner task in one operation",
  parameters: z.object({
    taskId: z.string().describe("The task ID"),
    items: z.array(z.string()).describe("Array of checklist item titles"),
  }),
  execute: async ({ taskId, items }) => {
    const etag = getETag("taskDetails", taskId);
    const checklist: Record<string, any> = {};

    for (const title of items) {
      const itemId = randomUUID();
      checklist[itemId] = {
        "@odata.type": "#microsoft.graph.plannerChecklistItem",
        title,
        isChecked: false,
      };
    }

    const body = { checklist };
    const url = `https://graph.microsoft.com/v1.0/planner/tasks/${taskId}/details`;
    const escapedEtag = etag.replace(/"/g, '\\"');
    const args = [
      `az rest --method PATCH --url "${url}"`,
      `--headers "Content-Type=application/json" "If-Match=${escapedEtag}"`,
      `--body '${JSON.stringify(body)}'`,
    ];
    try {
      const result = execSync(args.join(" "), { encoding: "utf-8" });
      return result || JSON.stringify({ success: true, itemCount: items.length });
    } catch (error: any) {
      throw new Error(`Add checklist items failed: ${error.message}`);
    }
  },
});

// Tool: Update checklist item (toggle or rename)
mcp.addTool({
  name: "update-checklist-item",
  description: "Update a checklist item (toggle checked state or rename)",
  parameters: z.object({
    taskId: z.string().describe("The task ID"),
    itemId: z.string().describe("The checklist item ID (from get-task-details)"),
    title: z.string().optional().describe("New title for the item"),
    isChecked: z.boolean().optional().describe("Set checked state"),
  }),
  execute: async ({ taskId, itemId, title, isChecked }) => {
    const etag = getETag("taskDetails", taskId);
    const itemUpdate: Record<string, any> = {
      "@odata.type": "#microsoft.graph.plannerChecklistItem",
    };
    if (title !== undefined) itemUpdate.title = title;
    if (isChecked !== undefined) itemUpdate.isChecked = isChecked;

    const body = {
      checklist: {
        [itemId]: itemUpdate,
      },
    };

    const url = `https://graph.microsoft.com/v1.0/planner/tasks/${taskId}/details`;
    const escapedEtag = etag.replace(/"/g, '\\"');
    const args = [
      `az rest --method PATCH --url "${url}"`,
      `--headers "Content-Type=application/json" "If-Match=${escapedEtag}"`,
      `--body '${JSON.stringify(body)}'`,
    ];
    try {
      const result = execSync(args.join(" "), { encoding: "utf-8" });
      return result || "Checklist item updated successfully";
    } catch (error: any) {
      throw new Error(`Update checklist item failed: ${error.message}`);
    }
  },
});

// Tool: Delete checklist item
mcp.addTool({
  name: "delete-checklist-item",
  description: "Delete a checklist item from a Planner task",
  parameters: z.object({
    taskId: z.string().describe("The task ID"),
    itemId: z.string().describe("The checklist item ID to delete"),
  }),
  execute: async ({ taskId, itemId }) => {
    const etag = getETag("taskDetails", taskId);
    const body = {
      checklist: {
        [itemId]: null,
      },
    };

    const url = `https://graph.microsoft.com/v1.0/planner/tasks/${taskId}/details`;
    const escapedEtag = etag.replace(/"/g, '\\"');
    const args = [
      `az rest --method PATCH --url "${url}"`,
      `--headers "Content-Type=application/json" "If-Match=${escapedEtag}"`,
      `--body '${JSON.stringify(body)}'`,
    ];
    try {
      const result = execSync(args.join(" "), { encoding: "utf-8" });
      return result || "Checklist item deleted successfully";
    } catch (error: any) {
      throw new Error(`Delete checklist item failed: ${error.message}`);
    }
  },
});

// Tool: Delete task
mcp.addTool({
  name: "delete-task",
  description: "Delete a Planner task. Auto-fetches ETag.",
  parameters: z.object({
    taskId: z.string().describe("The task ID to delete"),
  }),
  execute: async ({ taskId }) => {
    const etag = getETag("task", taskId);
    const url = `https://graph.microsoft.com/v1.0/planner/tasks/${taskId}`;
    // Escape the inner quotes in the ETag for shell: W/"..." -> W/\"...\"
    const escapedEtag = etag.replace(/"/g, '\\"');
    const args = [
      `az rest --method DELETE --url "${url}"`,
      `--headers "If-Match=${escapedEtag}"`,
    ];
    try {
      execSync(args.join(" "), { encoding: "utf-8" });
      return "Task deleted successfully";
    } catch (error: any) {
      throw new Error(`Delete failed: ${error.message}`);
    }
  },
});

// Tool: List buckets for a plan
mcp.addTool({
  name: "list-buckets",
  description: "List all buckets in a Planner plan",
  parameters: z.object({
    planId: z.string().describe("The Planner plan ID"),
  }),
  execute: async ({ planId }) => {
    const url = `https://graph.microsoft.com/v1.0/planner/plans/${planId}/buckets`;
    const result = JSON.parse(azRest("GET", url));
    return JSON.stringify(result.value, null, 2);
  },
});

// Tool: List plans for current user
mcp.addTool({
  name: "list-plans",
  description: "List all Planner plans accessible to the current user",
  parameters: z.object({}),
  execute: async () => {
    const url = "https://graph.microsoft.com/v1.0/me/planner/plans";
    const result = JSON.parse(azRest("GET", url));
    return JSON.stringify(result.value, null, 2);
  },
});

// Tool: Get plan details (includes category labels)
mcp.addTool({
  name: "get-plan-details",
  description: "Get plan details including category label names (what category1-25 mean)",
  parameters: z.object({
    planId: z.string().describe("The Planner plan ID"),
  }),
  execute: async ({ planId }) => {
    const url = `https://graph.microsoft.com/v1.0/planner/plans/${planId}/details`;
    const result = azRest("GET", url);
    return result;
  },
});

// Tool: Get all tasks assigned to current user across all plans
mcp.addTool({
  name: "get-my-tasks",
  description: "Get all tasks assigned to the current user across all plans",
  parameters: z.object({}),
  execute: async () => {
    const url = "https://graph.microsoft.com/v1.0/me/planner/tasks";
    const result = JSON.parse(azRest("GET", url));
    return JSON.stringify(result.value, null, 2);
  },
});

// Tool: List group members (for finding user IDs for assignment)
mcp.addTool({
  name: "list-group-members",
  description: "List all members of the group that owns a plan (returns user IDs for task assignment)",
  parameters: z.object({
    planId: z.string().describe("The Planner plan ID (will resolve to the group that owns it)"),
  }),
  execute: async ({ planId }) => {
    const groupId = getGroupIdFromPlan(planId);
    const url = `https://graph.microsoft.com/v1.0/groups/${groupId}/members`;
    const result = JSON.parse(azRest("GET", url));
    // Return simplified list with id and displayName
    const members = result.value.map((m: any) => ({
      id: m.id,
      displayName: m.displayName,
      userPrincipalName: m.userPrincipalName,
    }));
    return JSON.stringify(members, null, 2);
  },
});

// Tool: Get task comments
mcp.addTool({
  name: "get-task-comments",
  description: "Get all comments on a Planner task",
  parameters: z.object({
    taskId: z.string().describe("The task ID"),
  }),
  execute: async ({ taskId }) => {
    // Get task to find conversationThreadId and planId
    const taskUrl = `https://graph.microsoft.com/v1.0/planner/tasks/${taskId}`;
    const task = JSON.parse(azRest("GET", taskUrl));

    if (!task.conversationThreadId) {
      return JSON.stringify({ comments: [], message: "No comments on this task" });
    }

    const groupId = getGroupIdFromPlan(task.planId);

    // Planner's conversationThreadId is a conversation ID - get threads from it, then posts
    const threadsUrl = `https://graph.microsoft.com/v1.0/groups/${groupId}/conversations/${task.conversationThreadId}/threads`;
    const threadsResult = JSON.parse(azRest("GET", threadsUrl));

    // Collect posts from all threads
    const comments: any[] = [];
    for (const thread of threadsResult.value) {
      const postsUrl = `https://graph.microsoft.com/v1.0/groups/${groupId}/conversations/${task.conversationThreadId}/threads/${thread.id}/posts`;
      const postsResult = JSON.parse(azRest("GET", postsUrl));

      for (const post of postsResult.value) {
        comments.push({
          id: post.id,
          threadId: thread.id,
          content: post.body?.content,
          contentType: post.body?.contentType,
          createdDateTime: post.createdDateTime,
          from: post.from?.emailAddress?.name || post.from?.emailAddress?.address,
        });
      }
    }

    return JSON.stringify(comments, null, 2);
  },
});

// Tool: Add task comment
mcp.addTool({
  name: "add-task-comment",
  description: "Add a comment to a Planner task",
  parameters: z.object({
    taskId: z.string().describe("The task ID"),
    comment: z.string().describe("The comment text to add"),
  }),
  execute: async ({ taskId, comment }) => {
    // Get task to find conversationThreadId and planId
    const taskUrl = `https://graph.microsoft.com/v1.0/planner/tasks/${taskId}`;
    const task = JSON.parse(azRest("GET", taskUrl));
    const groupId = getGroupIdFromPlan(task.planId);

    if (task.conversationThreadId) {
      // Reply to existing conversation - need to get the thread ID first
      const threadsUrl = `https://graph.microsoft.com/v1.0/groups/${groupId}/conversations/${task.conversationThreadId}/threads`;
      const threadsResult = JSON.parse(azRest("GET", threadsUrl));

      if (!threadsResult.value || threadsResult.value.length === 0) {
        throw new Error("Conversation exists but has no threads");
      }

      const threadId = threadsResult.value[0].id;
      const replyUrl = `https://graph.microsoft.com/v1.0/groups/${groupId}/conversations/${task.conversationThreadId}/threads/${threadId}/reply`;
      const body = {
        post: {
          body: {
            contentType: "text",
            content: comment,
          },
        },
      };
      azRest("POST", replyUrl, body);
      return JSON.stringify({ success: true, message: "Comment added to existing thread" });
    } else {
      // Create new conversation (POST to threads creates a conversation with initial thread)
      const threadsUrl = `https://graph.microsoft.com/v1.0/groups/${groupId}/threads`;
      const threadBody = {
        topic: task.title,
        posts: [
          {
            body: {
              contentType: "text",
              content: comment,
            },
          },
        ],
      };
      const threadResult = JSON.parse(azRest("POST", threadsUrl, threadBody));
      // Graph API returns conversationId for the parent conversation - that's what Planner needs
      const conversationId = threadResult.conversationId || threadResult.id;

      // Update task with the conversation ID
      const etag = getETag("task", taskId);
      const escapedEtag = etag.replace(/"/g, '\\"');
      const updateUrl = `https://graph.microsoft.com/v1.0/planner/tasks/${taskId}`;
      const args = [
        `az rest --method PATCH --url "${updateUrl}"`,
        `--headers "Content-Type=application/json" "If-Match=${escapedEtag}"`,
        `--body '${JSON.stringify({ conversationThreadId: conversationId })}'`,
      ];
      try {
        execSync(args.join(" "), { encoding: "utf-8" });
      } catch (error: any) {
        // Thread was created but task update may fail - comment still exists
        return JSON.stringify({
          success: true,
          warning: "Thread created but task link may have failed",
          conversationId
        });
      }
      return JSON.stringify({ success: true, message: "New conversation created", conversationId });
    }
  },
});

// Tool: Move task to different bucket
mcp.addTool({
  name: "move-task",
  description: "Move a task to a different bucket",
  parameters: z.object({
    taskId: z.string().describe("The task ID"),
    bucketId: z.string().describe("The target bucket ID"),
  }),
  execute: async ({ taskId, bucketId }) => {
    const etag = getETag("task", taskId);
    const url = `https://graph.microsoft.com/v1.0/planner/tasks/${taskId}`;
    const escapedEtag = etag.replace(/"/g, '\\"');
    const args = [
      `az rest --method PATCH --url "${url}"`,
      `--headers "Content-Type=application/json" "If-Match=${escapedEtag}"`,
      `--body '${JSON.stringify({ bucketId })}'`,
    ];
    try {
      const result = execSync(args.join(" "), { encoding: "utf-8" });
      return result || "Task moved successfully";
    } catch (error: any) {
      throw new Error(`Move task failed: ${error.message}`);
    }
  },
});

// Tool: Create bucket
mcp.addTool({
  name: "create-bucket",
  description: "Create a new bucket in a Planner plan",
  parameters: z.object({
    planId: z.string().describe("The plan ID"),
    name: z.string().describe("The bucket name"),
  }),
  execute: async ({ planId, name }) => {
    const url = "https://graph.microsoft.com/v1.0/planner/buckets";
    const body = { planId, name, orderHint: " !" };
    const result = azRest("POST", url, body);
    return result;
  },
});

// Tool: Update bucket
mcp.addTool({
  name: "update-bucket",
  description: "Update a bucket's name",
  parameters: z.object({
    bucketId: z.string().describe("The bucket ID"),
    name: z.string().describe("The new bucket name"),
  }),
  execute: async ({ bucketId, name }) => {
    const etag = getETag("bucket", bucketId);
    const url = `https://graph.microsoft.com/v1.0/planner/buckets/${bucketId}`;
    const escapedEtag = etag.replace(/"/g, '\\"');
    const args = [
      `az rest --method PATCH --url "${url}"`,
      `--headers "Content-Type=application/json" "If-Match=${escapedEtag}"`,
      `--body '${JSON.stringify({ name })}'`,
    ];
    try {
      const result = execSync(args.join(" "), { encoding: "utf-8" });
      return result || "Bucket updated successfully";
    } catch (error: any) {
      throw new Error(`Update bucket failed: ${error.message}`);
    }
  },
});

// Tool: Delete bucket
mcp.addTool({
  name: "delete-bucket",
  description: "Delete a bucket from a Planner plan",
  parameters: z.object({
    bucketId: z.string().describe("The bucket ID to delete"),
  }),
  execute: async ({ bucketId }) => {
    const etag = getETag("bucket", bucketId);
    const url = `https://graph.microsoft.com/v1.0/planner/buckets/${bucketId}`;
    const escapedEtag = etag.replace(/"/g, '\\"');
    const args = [
      `az rest --method DELETE --url "${url}"`,
      `--headers "If-Match=${escapedEtag}"`,
    ];
    try {
      execSync(args.join(" "), { encoding: "utf-8" });
      return "Bucket deleted successfully";
    } catch (error: any) {
      throw new Error(`Delete bucket failed: ${error.message}`);
    }
  },
});

// Tool: Add reference (attachment link)
mcp.addTool({
  name: "add-reference",
  description: "Add a reference (URL attachment) to a Planner task",
  parameters: z.object({
    taskId: z.string().describe("The task ID"),
    url: z.string().describe("The URL to attach"),
    alias: z.string().optional().describe("Display name for the reference"),
    type: z.string().optional().describe("Reference type (e.g., 'Other', 'PowerPoint', 'Excel', 'Word', 'Pdf')"),
  }),
  execute: async ({ taskId, url: refUrl, alias, type }) => {
    const etag = getETag("taskDetails", taskId);
    const encodedUrl = encodeUrlForReference(refUrl);

    const referenceData: Record<string, any> = {
      "@odata.type": "#microsoft.graph.plannerExternalReference",
    };
    if (alias) referenceData.alias = alias;
    if (type) referenceData.type = type;

    const body = {
      references: {
        [encodedUrl]: referenceData,
      },
    };

    const url = `https://graph.microsoft.com/v1.0/planner/tasks/${taskId}/details`;
    try {
      const result = azRestPatchWithEtag(url, etag, body);
      return result || "Reference added successfully";
    } catch (error: any) {
      throw new Error(`Add reference failed: ${error.message}`);
    }
  },
});

// Tool: Delete reference
mcp.addTool({
  name: "delete-reference",
  description: "Delete a reference (URL attachment) from a Planner task",
  parameters: z.object({
    taskId: z.string().describe("The task ID"),
    url: z.string().describe("The URL of the reference to delete"),
  }),
  execute: async ({ taskId, url: refUrl }) => {
    const etag = getETag("taskDetails", taskId);
    const encodedUrl = encodeUrlForReference(refUrl);

    const body = {
      references: {
        [encodedUrl]: null,
      },
    };

    const url = `https://graph.microsoft.com/v1.0/planner/tasks/${taskId}/details`;
    try {
      const result = azRestPatchWithEtag(url, etag, body);
      return result || "Reference deleted successfully";
    } catch (error: any) {
      throw new Error(`Delete reference failed: ${error.message}`);
    }
  },
});

// Tool: Upload file attachment to SharePoint and attach to task
mcp.addTool({
  name: "upload-attachment",
  description: "Upload a local file to the plan's SharePoint and attach it to a task (max 4MB)",
  parameters: z.object({
    taskId: z.string().describe("The task ID"),
    filePath: z.string().describe("Local path to the file to upload"),
    alias: z.string().optional().describe("Display name for the attachment (defaults to filename)"),
  }),
  execute: async ({ taskId, filePath, alias }) => {
    // Read file
    const fileBuffer = readFileSync(filePath);
    const fileName = basename(filePath);
    const fileSize = fileBuffer.length;

    if (fileSize > 4 * 1024 * 1024) {
      throw new Error("File too large. Maximum size is 4MB. For larger files, upload manually to SharePoint.");
    }

    // Get task to find planId
    const taskUrl = `https://graph.microsoft.com/v1.0/planner/tasks/${taskId}`;
    const task = JSON.parse(azRest("GET", taskUrl));

    // Get groupId from plan
    const groupId = getGroupIdFromPlan(task.planId);

    // Get the group's drive (SharePoint document library)
    const driveUrl = `https://graph.microsoft.com/v1.0/groups/${groupId}/drive`;
    const drive = JSON.parse(azRest("GET", driveUrl));
    const driveId = drive.id;

    // Upload file to SharePoint (root of document library, in a "Planner Attachments" folder)
    const folderPath = "Planner Attachments";
    const uploadUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/root:/${folderPath}/${fileName}:/content`;

    // Write file to temp location for az rest to read
    const tmpFile = join(tmpdir(), `mcp-upload-${randomUUID()}-${fileName}`);
    try {
      writeFileSync(tmpFile, fileBuffer);

      // Upload using az rest with binary file
      const args = [
        `az rest --method PUT --url "${uploadUrl}"`,
        `--headers "Content-Type=application/octet-stream"`,
        `--body @${tmpFile}`,
      ];
      const uploadResult = JSON.parse(execSync(args.join(" "), { encoding: "utf-8", maxBuffer: 10 * 1024 * 1024 }));
      const fileId = uploadResult.id;

      // Create a sharing link for the file (cleaner URL than webUrl)
      const shareLinkUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${fileId}/createLink`;
      const shareLinkBody = { type: "view", scope: "organization" };
      const shareLinkResult = JSON.parse(azRest("POST", shareLinkUrl, shareLinkBody));
      const fileWebUrl = shareLinkResult.link.webUrl;

      // Add reference to task
      const etag = getETag("taskDetails", taskId);
      const encodedUrl = encodeUrlForReference(fileWebUrl);
      const referenceData: Record<string, any> = {
        "@odata.type": "#microsoft.graph.plannerExternalReference",
        alias: alias || fileName,
      };
      const refBody = {
        references: {
          [encodedUrl]: referenceData,
        },
      };
      const detailsUrl = `https://graph.microsoft.com/v1.0/planner/tasks/${taskId}/details`;
      azRestPatchWithEtag(detailsUrl, etag, refBody);

      return JSON.stringify({
        success: true,
        fileName,
        fileSize,
        sharePointUrl: fileWebUrl,
        message: "File uploaded to SharePoint and attached to task",
      });
    } finally {
      try { unlinkSync(tmpFile); } catch {}
    }
  },
});

mcp.start({ transportType: "stdio" });
