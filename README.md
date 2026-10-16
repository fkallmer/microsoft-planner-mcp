# microsoft-planner-mcp

Lightweight MCP (Model Context Protocol) server for Microsoft Planner that uses `az rest` for authentication instead of complex OAuth flows.

## What This Is

A minimal MCP server that enables Claude Code to interact with Microsoft Planner tasks directly. It leverages Azure CLI's existing authentication (`az login`) to make Graph API calls, eliminating the need for app registrations or token management.

## Prerequisites

- **Node.js** (v18 or later)
- **Azure CLI** installed and authenticated:
  ```bash
  az login
  ```

## Installation

```bash
git clone https://github.com/jbeineke-gcw/microsoft-planner-mcp.git
cd microsoft-planner-mcp
npm install
npm run build
```

## Usage with Claude Code

Add the MCP server to Claude Code:

```bash
claude mcp add microsoft-planner-mcp node /path/to/microsoft-planner-mcp/dist/index.js
```

Or add directly to your MCP settings file:

```json
{
  "mcpServers": {
    "microsoft-planner-mcp": {
      "command": "node",
      "args": ["/path/to/microsoft-planner-mcp/dist/index.js"]
    }
  }
}
```

## Available Tools

### Plans & Discovery
| Tool | Description |
|------|-------------|
| `list-plans` | List all Planner plans accessible to the current user |
| `get-plan-details` | Get plan details including category label names (what category1-25 mean) |
| `get-my-tasks` | Get all tasks assigned to the current user across all plans |
| `list-group-members` | List group members (returns user IDs for task assignment) |

### Buckets
| Tool | Description |
|------|-------------|
| `list-buckets` | List all buckets in a plan |
| `create-bucket` | Create a new bucket in a plan |
| `update-bucket` | Rename a bucket |
| `delete-bucket` | Delete a bucket from a plan |

### Tasks
| Tool | Description |
|------|-------------|
| `list-tasks` | List all tasks in a Planner plan |
| `get-task` | Get details of a specific task |
| `get-task-details` | Get extended task details (description, checklist, references) |
| `create-task` | Create a new task in a plan |
| `update-task` | Update task properties (title, progress, assignments, categories, due date) |
| `update-task-details` | Update task description (supports GitHub links) |
| `move-task` | Move a task to a different bucket |
| `delete-task` | Delete a Planner task |

### Checklist Items
| Tool | Description |
|------|-------------|
| `add-checklist-item` | Add a single checklist item (subtask) to a task |
| `add-checklist-items` | Add multiple checklist items in one operation |
| `update-checklist-item` | Update a checklist item (toggle checked or rename) |
| `delete-checklist-item` | Remove a checklist item from a task |

### Comments
| Tool | Description |
|------|-------------|
| `get-plan-overview` | Whole plan in one call: categories, buckets, tasks with ETags, optional details |
| `update-tasks` | Update many tasks in parallel (max 5 at a time), result per task |
| `get-task-comments` | Read Task Chat comments of a task (Graph beta, read-only) |

### References & Attachments
| Tool | Description |
|------|-------------|
| `add-reference` | Add a URL reference/attachment to a task |
| `delete-reference` | Remove a reference from a task |
| `upload-attachment` | Upload a local file to SharePoint and attach to a task (max 4MB) |

## Claude Code Agent (Optional)

For enhanced automation, create a Claude Code agent at `~/.claude/agents/microsoft-planner.md` with:

- **Auto-assignment**: Automatically assign tasks to a default user
- **Auto-labeling**: Apply default categories/labels to tasks
- **Status Intelligence**: Infer task bucket and progress from conversation context
- **GitHub Integration**: Include repository links in task descriptions

See the [agent template](https://github.com/vyente-ruffin/microsoft-planner-mcp/wiki/Agent-Configuration) for configuration details.

## Finding Your IDs

To use the MCP tools, you'll need your Planner Plan ID and Bucket IDs:

1. Use `list-plans` to discover all accessible plans and their IDs
2. Use `list-buckets` with your Plan ID to get bucket IDs
3. Or find your Plan ID from the Planner web URL: `https://tasks.office.com/...planId=YOUR_PLAN_ID`

## Finding User IDs for Assignment

To assign tasks to users, you need their Microsoft Graph user IDs:

1. Use `list-group-members` with your Plan ID to get all group members
2. This returns each user's `id`, `displayName`, and `userPrincipalName`
3. Use the `id` value with `update-task`'s `assignUserId` parameter

Example workflow:
```
"Who can I assign tasks to?" → list-group-members
"Assign this task to Joe" → update-task with Joe's user ID
```

## Understanding Category Labels

Planner uses generic category names (`category1` through `category25`) internally, but plans can define custom display labels:

1. Use `get-plan-details` to see the category label mappings
2. The `categoryDescriptions` object shows what each category means in that plan
3. Example: `category1` might be labeled "Bug", `category2` might be "Feature"

This helps when:
- You see a task has `appliedCategories: { category6: true }` and need to know what that means
- You want to apply the correct category based on its label rather than guessing the number

## Example Usage

Once configured, use natural language with Claude Code:

```
"List all my planner plans"
"List all tasks in the planner"
"Create a task called 'Review PR #123' in the backlog bucket"
"Mark task XYZ as complete"
"Set this task's due date to next Friday"
"Add a description with the GitHub PR link to the task"
"Add a checklist with: design, implement, test, document"
"Check off the 'design' item on that task"
"What do the category labels mean in this plan?"
"Who can I assign tasks to?"
"Assign this task to Joe"
"Move this task to the 'In Progress' bucket"
"Show me the comments on this task"
"Attach this GitHub PR link to the task"
"Upload /path/to/data-export.csv to this task"
"Create a new bucket called 'Blocked'"
"Show me all my assigned tasks across all plans"
```

## How It Works

This server uses `az rest` to make Microsoft Graph API calls. The Azure CLI handles all authentication, so as long as you're logged in with `az login`, the server can access Planner data your account has permissions for.

All Planner operations that require ETags (update, delete) automatically fetch the current ETag before making changes.

## License

ISC
