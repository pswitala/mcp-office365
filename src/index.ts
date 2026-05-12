#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";

import { getMsalApp, acquireTokenForAlias } from "./auth/msal.js";
import { performLogin } from "./auth/login.js";
import { listAliases, removeAccount } from "./auth/accounts.js";
import { saveConfig, loadConfig } from "./config.js";
import { mailToolDefinitions } from "./tools/mail.js";
import { calendarToolDefinitions } from "./tools/calendar.js";
import {
  INBOX_RESOURCE_URI,
  readInboxResource,
} from "./resources/inbox.js";
import {
  CALENDAR_TODAY_URI,
  CALENDAR_WEEK_URI,
  readCalendarTodayResource,
  readCalendarWeekResource,
} from "./resources/calendar.js";

const server = new McpServer({
  name: "mcp-office365",
  version: "0.1.0",
});

// ---------------------------------------------------------------------------
// Auth tools
// ---------------------------------------------------------------------------

server.tool(
  "check_login",
  "Verify that a saved mailbox alias has a valid token (does not trigger login — use the login CLI script for that).",
  {
    alias: z.string().describe("Mailbox alias to check"),
  },
  async ({ alias }) => {
    try {
      await acquireTokenForAlias(alias);
      return { content: [{ type: "text", text: `Alias "${alias}" is authenticated and ready.` }] };
    } catch (err) {
      return {
        content: [
          {
            type: "text",
            text: `Alias "${alias}" is not authenticated.\nRun: use the 'configure' tool to set a client ID, then 'login' with alias "${alias}".`,
          },
        ],
      };
    }
  }
);

server.tool(
  "list_accounts",
  "List all configured mailbox aliases.",
  {},
  async () => {
    const aliases = await listAliases();
    return {
      content: [
        {
          type: "text",
          text: aliases.length
            ? `Configured aliases:\n${aliases.map((a) => `  - ${a}`).join("\n")}`
            : "No mailboxes configured. Use the login tool to add one.",
        },
      ],
    };
  }
);

server.tool(
  "remove_account",
  "Remove a mailbox alias.",
  {
    alias: z.string().describe("Alias to remove"),
  },
  async ({ alias }) => {
    await removeAccount(alias);
    return { content: [{ type: "text", text: `Alias "${alias}" removed.` }] };
  }
);

server.tool(
  "configure",
  "Configure the Azure AD application client ID. This persists to ~/.mcp-o365/config.json so the env var is not needed on subsequent runs.",
  {
    clientId: z.string().describe("Azure AD Application (client) ID from portal.azure.com"),
  },
  async ({ clientId }) => {
    await saveConfig(clientId);
    return {
      content: [
        {
          type: "text",
          text: [
            `Client ID saved.`,
            "",
            "Next steps:",
            "  1. Use the 'login' tool to authenticate a mailbox (device code flow).",
            "  2. Use 'status' to verify configuration.",
            "",
            "If you don't have a client ID yet, register one at https://portal.azure.com:",
            "  1. New registration -> 'Accounts in any organizational directory",
            "     (multi-tenant) and personal Microsoft accounts'.",
            "  2. Authentication -> Add platform -> Mobile and desktop ->",
            "     enable 'Allow public client flows'.",
            "  3. API permissions -> Microsoft Graph -> Delegated:",
            "     Mail.Read, Mail.Send, Mail.ReadWrite, Calendars.ReadWrite, User.Read",
          ].join("\n"),
        },
      ],
    };
  }
);

server.tool(
  "login",
  "Authenticate a mailbox using device code flow. Opens a browser dialog for the user to sign in and authorize permissions.",
  {
    alias: z.string().describe("Alias name for this mailbox (e.g. 'work', 'personal')"),
  },
  async ({ alias }) => {
    try {
      const app = await getMsalApp();
      const result = await performLogin(app, alias);
      return {
        content: [{ type: "text", text: result.message }],
        isError: !result.success,
      };
    } catch (err) {
      return {
        content: [
          {
            type: "text",
            text: `Login failed: ${err instanceof Error ? err.message : String(err)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

server.tool(
  "status",
  "Show current configuration status: whether a client ID is configured, list of mailbox aliases, and token validity.",
  {},
  async () => {
    const cfg = await loadConfig();
    const aliases = await listAliases();
    const lines: string[] = [];

    if (cfg?.clientId) {
      lines.push(`Client ID: configured (${cfg.clientId.slice(0, 8)}...${cfg.clientId.slice(-4)})`);
    } else {
      lines.push("Client ID: NOT configured — use the 'configure' tool to set it");
    }

    lines.push(`Aliases: ${aliases.length} configured`);
    for (const a of aliases) lines.push(`  - ${a}`);

    // Check token validity for each alias
    for (const alias of aliases) {
      try {
        await acquireTokenForAlias(alias);
        lines.push(`  [${alias}] token: valid`);
      } catch {
        lines.push(`  [${alias}] token: expired/invalid — use 'login' to re-authenticate`);
      }
    }

    return {
      content: [{ type: "text", text: lines.join("\n") }],
    };
  }
);

// ---------------------------------------------------------------------------
// Mail tools
// ---------------------------------------------------------------------------

for (const def of mailToolDefinitions) {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (server.tool as any)(def.name, def.description, def.schema.shape, def.handler);
}

// ---------------------------------------------------------------------------
// Calendar tools
// ---------------------------------------------------------------------------

for (const def of calendarToolDefinitions) {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (server.tool as any)(def.name, def.description, def.schema.shape, def.handler);
}

// ---------------------------------------------------------------------------
// Resources
// ---------------------------------------------------------------------------

server.resource(
  "inbox",
  INBOX_RESOURCE_URI,
  { description: "Last 10 unread messages across all configured mailboxes" },
  async () => ({
    contents: [{ uri: INBOX_RESOURCE_URI, mimeType: "application/json", text: await readInboxResource() }],
  })
);

server.resource(
  "calendar-today",
  CALENDAR_TODAY_URI,
  { description: "Today's calendar events across all configured mailboxes" },
  async () => ({
    contents: [{ uri: CALENDAR_TODAY_URI, mimeType: "application/json", text: await readCalendarTodayResource() }],
  })
);

server.resource(
  "calendar-week",
  CALENDAR_WEEK_URI,
  { description: "Calendar events for the next 7 days across all configured mailboxes" },
  async () => ({
    contents: [{ uri: CALENDAR_WEEK_URI, mimeType: "application/json", text: await readCalendarWeekResource() }],
  })
);

// ---------------------------------------------------------------------------
// Start
// ---------------------------------------------------------------------------

const transport = new StdioServerTransport();
await server.connect(transport);
