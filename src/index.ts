import { FastMCP } from "fastmcp";
import { z } from "zod";
import { execFileSync } from "child_process";
import { randomUUID } from "crypto";
import { readFileSync } from "fs";
import { basename } from "path";

const mcp = new FastMCP({
  name: "microsoft-planner-mcp",
  version: "1.0.0",
});

// IDs and names go into URL paths: encode them so a value can never change the endpoint
// (e.g. "abc/../../users" or "?$select=...").
const seg = (value: string) => encodeURIComponent(value);

interface GraphOptions {
  body?: object;            // JSON body
  raw?: Buffer;             // raw body (e.g. file upload)
  etag?: string;            // sent as If-Match
  contentType?: string;
  prefer?: string;          // e.g. "return=representation": PATCH returns the updated object
}

// Graph token from the Azure CLI (same login as `az rest`), fetched once and reused until shortly
// before it expires. Starting `az` for every request cost ~0.4 s; a direct request takes ~0.15 s.
// No shell is involved: execFileSync gets an argument array, fetch gets the payload as data.
let tokenCache: { token: string; ablauf: number } | null = null;

function graphToken(): string {
  if (tokenCache && Date.now() < tokenCache.ablauf - 5 * 60_000) return tokenCache.token;
  let out: string;
  try {
    out = execFileSync("az", ["account", "get-access-token", "--resource", "https://graph.microsoft.com", "-o", "json"],
      { encoding: "utf-8" });
  } catch (error: any) {
    throw new Error(`No Graph token from the Azure CLI (run az login): ${error.stderr?.toString().trim() || error.message}`);
  }
  const t = JSON.parse(out);
  const ablauf = t.expires_on ? Number(t.expires_on) * 1000 : Date.parse(t.expiresOn);
  tokenCache = { token: t.accessToken, ablauf };
  return t.accessToken;
}

async function graph(method: string, url: string, opts: GraphOptions = {}): Promise<string> {
  for (let versuch = 0; ; versuch++) {
    const headers: Record<string, string> = { Authorization: `Bearer ${graphToken()}` };
    let body: string | Blob | undefined;
    if (opts.body !== undefined) {
      headers["Content-Type"] = opts.contentType ?? "application/json";
      body = JSON.stringify(opts.body);
    } else if (opts.raw) {
      headers["Content-Type"] = opts.contentType ?? "application/octet-stream";
      body = new Blob([new Uint8Array(opts.raw)]);
    }
    if (opts.etag) headers["If-Match"] = opts.etag;
    if (opts.prefer) headers["Prefer"] = opts.prefer;
    const res = await fetch(url, { method, headers, body });
    // Token revoked or expired early: fetch a fresh one once
    if (res.status === 401 && versuch === 0) { tokenCache = null; continue; }
    // Throttling: honour Retry-After a few times (matters for parallel requests)
    if ((res.status === 429 || res.status === 503) && versuch < 3) {
      const warte = Math.min(Number(res.headers.get("Retry-After") ?? 1) || 1, 10);
      await new Promise(r => setTimeout(r, warte * 1000));
      continue;
    }
    const text = await res.text();
    if (!res.ok) throw new Error(`Graph ${method} failed (${res.status}): ${text.slice(0, 500)}`);
    return text;
  }
}

// Kept for the existing call sites
async function azRest(method: string, url: string, body?: object): Promise<string> {
  return graph(method, url, { body });
}

// GET a Graph collection and follow @odata.nextLink until every page is loaded.
// Graph pages large collections; reading only result.value silently drops the rest.
// (Idea from camtarn/microsoft-planner-mcp; no shell here, so $skiptoken needs no escaping.)
async function azRestGetAll(url: string): Promise<any[]> {
  const items: any[] = [];
  let next: string | undefined = url;
  while (next) {
    const result = JSON.parse(await azRest("GET", next));
    items.push(...(result.value ?? []));
    next = result["@odata.nextLink"];
  }
  return items;
}

// Helper to get ETag for update/delete operations (supports multiple resource types)
type ResourceType = "task" | "taskDetails" | "bucket";

async function getETag(resourceType: ResourceType, resourceId: string): Promise<string> {
  const urlMap: Record<ResourceType, string> = {
    task: `https://graph.microsoft.com/v1.0/planner/tasks/${seg(resourceId)}`,
    taskDetails: `https://graph.microsoft.com/v1.0/planner/tasks/${seg(resourceId)}/details`,
    bucket: `https://graph.microsoft.com/v1.0/planner/buckets/${seg(resourceId)}`,
  };
  const result = JSON.parse(await azRest("GET", urlMap[resourceType]));
  return result["@odata.etag"];
}

// Helper to get groupId from a plan (required for comments and group member listing)
async function getGroupIdFromPlan(planId: string): Promise<string> {
  const url = `https://graph.microsoft.com/v1.0/planner/plans/${seg(planId)}`;
  const result = JSON.parse(await azRest("GET", url));
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

// Helper for PATCH requests with an ETag
async function azRestPatchWithEtag(url: string, etag: string, body: object): Promise<string> {
  return graph("PATCH", url, { etag, body });
}

// Tool: List tasks for a plan
mcp.addTool({
  name: "list-tasks",
  description: "List all tasks in a Planner plan",
  parameters: z.object({
    planId: z.string().describe("The Planner plan ID"),
  }),
  execute: async ({ planId }) => {
    const url = `https://graph.microsoft.com/v1.0/planner/plans/${seg(planId)}/tasks`;
    return JSON.stringify(await azRestGetAll(url), null, 2);
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
    const url = `https://graph.microsoft.com/v1.0/planner/tasks/${seg(taskId)}`;
    const result = await azRest("GET", url);
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
    const url = `https://graph.microsoft.com/v1.0/planner/tasks/${seg(taskId)}/details`;
    const result = await azRest("GET", url);
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
    const result = await azRest("POST", url, body);
    return result;
  },
});

// Fields shared by update-task and update-tasks
const TaskUpdateFields = {
  title: z.string().optional().describe("New title"),
  percentComplete: z.number().min(0).max(100).optional().describe("Progress 0-100"),
  assignUserId: z.string().optional().describe("User ID to assign"),
  category: z.string().optional().describe("Category to apply (category1-category25)"),
  removeCategory: z.string().optional().describe("Category to remove (category1-category25)"),
  dueDateTime: z.string().optional().describe("Due date (ISO 8601 format, e.g., '2024-12-31' or '2024-12-31T17:00:00Z'). Use 'clear' to remove due date."),
  priority: z.number().min(0).max(10).optional().describe("Priority: 1=urgent, 3=important, 5=medium, 9=low"),
  bucketId: z.string().optional().describe("Move the task to this bucket"),
  etag: z.string().optional().describe("ETag from a previous read (list-tasks, get-plan-overview); saves one request"),
};

type TaskUpdate = { taskId: string } & { [K in keyof typeof TaskUpdateFields]?: any };

function taskUpdateBody(u: TaskUpdate): Record<string, any> {
  const body: Record<string, any> = {};
  if (u.title !== undefined) body.title = u.title;
  if (u.percentComplete !== undefined) body.percentComplete = u.percentComplete;
  if (u.priority !== undefined) body.priority = u.priority;
  if (u.bucketId) body.bucketId = u.bucketId;
  if (u.assignUserId) {
    body.assignments = { [u.assignUserId]: { "@odata.type": "#microsoft.graph.plannerAssignment", orderHint: " !" } };
  }
  if (u.category || u.removeCategory) {
    body.appliedCategories = {};
    if (u.category) body.appliedCategories[u.category] = true;
    if (u.removeCategory) body.appliedCategories[u.removeCategory] = false;  // Graph erwartet false, null ist ungültig
  }
  if (u.dueDateTime) {
    // A plain date (YYYY-MM-DD) becomes end of day UTC
    body.dueDateTime = u.dueDateTime.toLowerCase() === "clear" ? null
      : u.dueDateTime.includes("T") ? u.dueDateTime : `${u.dueDateTime}T23:59:59Z`;
  }
  return body;
}

// PATCH a task: uses the given ETag (or fetches one), retries once with a fresh ETag if it was
// stale (412), and returns the updated task (Prefer: return=representation, no extra read).
async function patchTask(u: TaskUpdate): Promise<any> {
  const url = `https://graph.microsoft.com/v1.0/planner/tasks/${seg(u.taskId)}`;
  const body = taskUpdateBody(u);
  let etag = u.etag ?? await getETag("task", u.taskId);
  for (let versuch = 0; ; versuch++) {
    try {
      const text = await graph("PATCH", url, { etag, body, prefer: "return=representation" });
      return text ? JSON.parse(text) : { id: u.taskId };
    } catch (error: any) {
      if (versuch === 0 && String(error.message).includes("(412)")) { etag = await getETag("task", u.taskId); continue; }
      throw error;
    }
  }
}

// Tool: Update task (title, percentComplete, assignments, categories, priority, bucket)
mcp.addTool({
  name: "update-task",
  description: "Update task properties (title, progress, assignments, categories, due date, priority, bucket). Pass etag from a previous read to save a request. Returns the updated task.",
  parameters: z.object({ taskId: z.string().describe("The task ID"), ...TaskUpdateFields }),
  execute: async (u) => {
    try {
      return JSON.stringify(await patchTask(u), null, 2);
    } catch (error: any) {
      throw new Error(`Update failed: ${error.message}`);
    }
  },
});

// Tool: Update many tasks at once (in parallel, at most 5 at a time)
mcp.addTool({
  name: "update-tasks",
  description: "Update several tasks in one call (e.g. mark many as done). Runs up to 5 updates in parallel and reports the result per task; one failure does not stop the others.",
  parameters: z.object({
    updates: z.array(z.object({ taskId: z.string().describe("The task ID"), ...TaskUpdateFields })).min(1).max(100),
  }),
  execute: async ({ updates }) => {
    const ergebnisse: any[] = new Array(updates.length);
    let naechster = 0;
    const arbeiter = async () => {
      while (naechster < updates.length) {
        const i = naechster++;
        try {
          const t = await patchTask(updates[i]);
          ergebnisse[i] = { taskId: updates[i].taskId, ok: true, title: t.title, percentComplete: t.percentComplete };
        } catch (error: any) {
          ergebnisse[i] = { taskId: updates[i].taskId, ok: false, error: String(error.message).slice(0, 300) };
        }
      }
    };
    await Promise.all(Array.from({ length: Math.min(5, updates.length) }, arbeiter));
    return JSON.stringify({ ok: ergebnisse.filter(e => e.ok).length, failed: ergebnisse.filter(e => !e.ok).length, results: ergebnisse }, null, 2);
  },
});

// Tool: Whole plan in one call (plan categories, buckets, tasks, optional details), loaded in parallel
mcp.addTool({
  name: "get-plan-overview",
  description: "Get a whole plan in one call: category names, buckets in board order, and all tasks per bucket (with ETags for fast updates). With includeDetails the description and checklist of every task are included (one request via $expand=details).",
  parameters: z.object({
    planId: z.string().describe("The Planner plan ID"),
    includeDetails: z.boolean().optional().default(false).describe("Include description and checklist of each task"),
    includeCompleted: z.boolean().optional().default(true).describe("Include completed tasks"),
  }),
  execute: async ({ planId, includeDetails, includeCompleted }) => {
    const basis = `https://graph.microsoft.com/v1.0/planner/plans/${seg(planId)}`;
    const [plan, details, buckets, tasks] = await Promise.all([
      azRest("GET", basis).then(JSON.parse),
      azRest("GET", `${basis}/details`).then(JSON.parse),
      azRestGetAll(`${basis}/buckets`),
      azRestGetAll(`${basis}/tasks${includeDetails ? "?$expand=details" : ""}`),
    ]);
    const kategorien: Record<string, string> = details.categoryDescriptions ?? {};
    // Planner sorts ascending by orderHint (ordinal string comparison), like the board
    const reihenfolge = (a: any, b: any) => (a.orderHint < b.orderHint ? -1 : a.orderHint > b.orderHint ? 1 : 0);
    const kurz = (t: any) => ({
      id: t.id, etag: t["@odata.etag"], title: t.title, percentComplete: t.percentComplete, priority: t.priority,
      dueDateTime: t.dueDateTime, assignees: Object.keys(t.assignments ?? {}).length,
      categories: Object.entries(t.appliedCategories ?? {}).filter(([, v]) => v).map(([k]) => kategorien[k] || k),
      ...(includeDetails && t.details ? {
        description: t.details.description || undefined,
        checklist: Object.values(t.details.checklist ?? {}).map((c: any) => ({ title: c.title, done: c.isChecked })),
      } : {}),
    });
    const sichtbar = tasks.filter(t => includeCompleted || t.percentComplete < 100);
    return JSON.stringify({
      plan: { id: plan.id, title: plan.title, categories: kategorien },
      total: tasks.length, open: tasks.filter(t => t.percentComplete < 100).length,
      buckets: buckets.sort(reihenfolge).map(b => ({
        id: b.id, name: b.name,
        tasks: sichtbar.filter(t => t.bucketId === b.id).sort(reihenfolge).map(kurz),
      })),
    }, null, 2);
  },
});

// Tool: Update task details (description with GitHub links)
mcp.addTool({
  name: "update-task-details",
  description: "Update task description (use for GitHub links). Auto-fetches ETag.",
  parameters: z.object({
    taskId: z.string().describe("The task ID"),
    description: z.string().describe("Task description (supports markdown, include GitHub URLs)"),
    etag: z.string().optional().describe("ETag of the task details from a previous read; saves one request"),
  }),
  execute: async ({ taskId, description, etag: bekannt }) => {
    const etag = bekannt ?? await getETag("taskDetails", taskId);
    const url = `https://graph.microsoft.com/v1.0/planner/tasks/${seg(taskId)}/details`;
    try {
      const result = await graph("PATCH", url, { etag, body: { description } });
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
    const etag = await getETag("taskDetails", taskId);
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

    const url = `https://graph.microsoft.com/v1.0/planner/tasks/${seg(taskId)}/details`;
    try {
      const result = await graph("PATCH", url, { etag, body: body });
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
    const etag = await getETag("taskDetails", taskId);
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
    const url = `https://graph.microsoft.com/v1.0/planner/tasks/${seg(taskId)}/details`;
    try {
      const result = await graph("PATCH", url, { etag, body: body });
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
    const etag = await getETag("taskDetails", taskId);
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

    const url = `https://graph.microsoft.com/v1.0/planner/tasks/${seg(taskId)}/details`;
    try {
      const result = await graph("PATCH", url, { etag, body: body });
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
    const etag = await getETag("taskDetails", taskId);
    const body = {
      checklist: {
        [itemId]: null,
      },
    };

    const url = `https://graph.microsoft.com/v1.0/planner/tasks/${seg(taskId)}/details`;
    try {
      const result = await graph("PATCH", url, { etag, body: body });
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
    const etag = await getETag("task", taskId);
    const url = `https://graph.microsoft.com/v1.0/planner/tasks/${seg(taskId)}`;
    try {
      await graph("DELETE", url, { etag });
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
    const url = `https://graph.microsoft.com/v1.0/planner/plans/${seg(planId)}/buckets`;
    return JSON.stringify(await azRestGetAll(url), null, 2);
  },
});

// Tool: List plans for current user
mcp.addTool({
  name: "list-plans",
  description: "List all Planner plans accessible to the current user",
  parameters: z.object({}),
  execute: async () => {
    const url = "https://graph.microsoft.com/v1.0/me/planner/plans";
    return JSON.stringify(await azRestGetAll(url), null, 2);
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
    const url = `https://graph.microsoft.com/v1.0/planner/plans/${seg(planId)}/details`;
    const result = await azRest("GET", url);
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
    return JSON.stringify(await azRestGetAll(url), null, 2);
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
    const groupId = await getGroupIdFromPlan(planId);
    const url = `https://graph.microsoft.com/v1.0/groups/${seg(groupId)}/members`;
    // Return simplified list with id and displayName
    const members = (await azRestGetAll(url)).map((m: any) => ({
      id: m.id,
      displayName: m.displayName,
      userPrincipalName: m.userPrincipalName,
    }));
    return JSON.stringify(members, null, 2);
  },
});

// Tool: Get task comments from the new Planner "Task Chat" (beta API).
// The legacy comments lived in the plan's Microsoft 365 group mailbox; writing one sent a mail
// to the whole group, so there is deliberately no tool to add comments here.
mcp.addTool({
  name: "get-task-comments",
  description: "Read the Task Chat comments of a Planner task (Microsoft Graph beta). Returns an empty list if the task has no chat yet. Adding comments is not supported.",
  parameters: z.object({
    taskId: z.string().describe("The task ID"),
  }),
  execute: async ({ taskId }) => {
    const url = `https://graph.microsoft.com/beta/planner/tasks/${seg(taskId)}/messages`;
    try {
      return JSON.stringify(await azRestGetAll(url), null, 2);
    } catch (error: any) {
      if (String(error.message).includes("does not have an associated chat thread")) {
        return JSON.stringify([], null, 2);
      }
      throw error;
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
    const etag = await getETag("task", taskId);
    const url = `https://graph.microsoft.com/v1.0/planner/tasks/${seg(taskId)}`;
    try {
      const result = await graph("PATCH", url, { etag, body: { bucketId } });
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
    const result = await azRest("POST", url, body);
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
    const etag = await getETag("bucket", bucketId);
    const url = `https://graph.microsoft.com/v1.0/planner/buckets/${seg(bucketId)}`;
    try {
      const result = await graph("PATCH", url, { etag, body: { name } });
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
    const etag = await getETag("bucket", bucketId);
    const url = `https://graph.microsoft.com/v1.0/planner/buckets/${seg(bucketId)}`;
    try {
      await graph("DELETE", url, { etag });
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
    const etag = await getETag("taskDetails", taskId);
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

    const url = `https://graph.microsoft.com/v1.0/planner/tasks/${seg(taskId)}/details`;
    try {
      const result = await azRestPatchWithEtag(url, etag, body);
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
    const etag = await getETag("taskDetails", taskId);
    const encodedUrl = encodeUrlForReference(refUrl);

    const body = {
      references: {
        [encodedUrl]: null,
      },
    };

    const url = `https://graph.microsoft.com/v1.0/planner/tasks/${seg(taskId)}/details`;
    try {
      const result = await azRestPatchWithEtag(url, etag, body);
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
    const taskUrl = `https://graph.microsoft.com/v1.0/planner/tasks/${seg(taskId)}`;
    const task = JSON.parse(await azRest("GET", taskUrl));

    // Get groupId from plan
    const groupId = await getGroupIdFromPlan(task.planId);

    // Get the group's drive (SharePoint document library)
    const driveUrl = `https://graph.microsoft.com/v1.0/groups/${seg(groupId)}/drive`;
    const drive = JSON.parse(await azRest("GET", driveUrl));
    const driveId = drive.id;

    // Upload file to SharePoint (root of document library, in a "Planner Attachments" folder)
    const folderPath = "Planner Attachments";
    const uploadUrl = `https://graph.microsoft.com/v1.0/drives/${seg(driveId)}/items/root:/${seg(folderPath)}/${seg(fileName)}:/content`;

    try {
      const uploadResult = JSON.parse(await graph("PUT", uploadUrl, { raw: fileBuffer }));
      const fileId = uploadResult.id;

      // Create a sharing link for the file (cleaner URL than webUrl)
      const shareLinkUrl = `https://graph.microsoft.com/v1.0/drives/${seg(driveId)}/items/${seg(fileId)}/createLink`;
      const shareLinkBody = { type: "view", scope: "organization" };
      const shareLinkResult = JSON.parse(await azRest("POST", shareLinkUrl, shareLinkBody));
      const fileWebUrl = shareLinkResult.link.webUrl;

      // Add reference to task
      const etag = await getETag("taskDetails", taskId);
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
      const detailsUrl = `https://graph.microsoft.com/v1.0/planner/tasks/${seg(taskId)}/details`;
      await azRestPatchWithEtag(detailsUrl, etag, refBody);

      return JSON.stringify({
        success: true,
        fileName,
        fileSize,
        sharePointUrl: fileWebUrl,
        message: "File uploaded to SharePoint and attached to task",
      });
    } catch (error: any) {
      throw new Error(`Upload failed: ${error.message}`);
    }
  },
});

mcp.start({ transportType: "stdio" });
