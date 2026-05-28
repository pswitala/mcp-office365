import { z } from "zod";
import {
  listMessages,
  getMessage,
  sendMessage,
  replyMessage,
  searchMessages,
  createDraft,
  moveMessage,
  listFolders,
  listFolderTree,
  findFolder,
} from "../graph/mail.js";

export const mailToolDefinitions = [
  {
    name: "list_messages",
    description: "List email messages from a mailbox folder (default: inbox). Supports pagination.",
    schema: z.object({
      alias: z.string().describe("Mailbox alias (from login)"),
      folder: z.string().optional().default("inbox").describe("Folder name or ID (default: inbox)"),
      search: z.string().optional().describe("Search query"),
      top: z.number().int().min(1).max(100).optional().default(25).describe("Number of results (max 100)"),
      pageToken: z.string().optional().describe("nextPageToken from a previous call"),
    }),
    handler: async (args: {
      alias: string;
      folder?: string;
      search?: string;
      top?: number;
      pageToken?: string;
    }) => {
      const result = await listMessages(args.alias, {
        folder: args.folder,
        search: args.search,
        top: args.top,
        pageToken: args.pageToken,
      });
      return { content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }] };
    },
  },

  {
    name: "get_message",
    description: "Get the full content of an email message by ID.",
    schema: z.object({
      alias: z.string().describe("Mailbox alias"),
      id: z.string().describe("Message ID"),
    }),
    handler: async (args: { alias: string; id: string }) => {
      const msg = await getMessage(args.alias, args.id);
      return { content: [{ type: "text" as const, text: JSON.stringify(msg, null, 2) }] };
    },
  },

  {
    name: "search_messages",
    description: "Search emails using a keyword query.",
    schema: z.object({
      alias: z.string().describe("Mailbox alias"),
      query: z.string().describe("Search keywords"),
      top: z.number().int().min(1).max(100).optional().default(25),
      pageToken: z.string().optional(),
    }),
    handler: async (args: { alias: string; query: string; top?: number; pageToken?: string }) => {
      const result = await searchMessages(args.alias, args.query, {
        top: args.top,
        pageToken: args.pageToken,
      });
      return { content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }] };
    },
  },

  {
    name: "send_message",
    description: "Send an email message.",
    schema: z.object({
      alias: z.string().describe("Mailbox alias"),
      to: z.array(z.string()).describe("Recipient email addresses"),
      subject: z.string(),
      body: z.string().describe("Message body"),
      bodyType: z.enum(["text", "html"]).optional().default("text"),
      cc: z.array(z.string()).optional(),
    }),
    handler: async (args: {
      alias: string;
      to: string[];
      subject: string;
      body: string;
      bodyType?: "text" | "html";
      cc?: string[];
    }) => {
      await sendMessage(args.alias, {
        to: args.to,
        subject: args.subject,
        body: args.body,
        bodyType: args.bodyType,
        cc: args.cc,
      });
      return { content: [{ type: "text" as const, text: "Message sent." }] };
    },
  },

  {
    name: "reply_message",
    description: "Reply to an email message.",
    schema: z.object({
      alias: z.string().describe("Mailbox alias"),
      messageId: z.string().describe("ID of the message to reply to"),
      body: z.string().describe("Reply text"),
      bodyType: z.enum(["text", "html"]).optional().default("text"),
    }),
    handler: async (args: { alias: string; messageId: string; body: string; bodyType?: "text" | "html" }) => {
      await replyMessage(args.alias, args.messageId, args.body, args.bodyType);
      return { content: [{ type: "text" as const, text: "Reply sent." }] };
    },
  },

  {
    name: "create_draft",
    description: "Create a draft email without sending it.",
    schema: z.object({
      alias: z.string(),
      to: z.array(z.string()),
      subject: z.string(),
      body: z.string(),
      bodyType: z.enum(["text", "html"]).optional().default("text"),
      cc: z.array(z.string()).optional(),
    }),
    handler: async (args: {
      alias: string;
      to: string[];
      subject: string;
      body: string;
      bodyType?: "text" | "html";
      cc?: string[];
    }) => {
      const id = await createDraft(args.alias, {
        to: args.to,
        subject: args.subject,
        body: args.body,
        bodyType: args.bodyType,
        cc: args.cc,
      });
      return { content: [{ type: "text" as const, text: JSON.stringify({ draftId: id }) }] };
    },
  },

  {
    name: "move_message",
    description: "Move an email to a different folder.",
    schema: z.object({
      alias: z.string(),
      messageId: z.string(),
      destinationFolderId: z.string().describe("Folder ID or well-known name (e.g. deleteditems, junkemail)"),
    }),
    handler: async (args: { alias: string; messageId: string; destinationFolderId: string }) => {
      await moveMessage(args.alias, args.messageId, args.destinationFolderId);
      return { content: [{ type: "text" as const, text: "Message moved." }] };
    },
  },

  {
    name: "list_folders",
    description: "List mail folders in a mailbox.",
    schema: z.object({
      alias: z.string(),
    }),
    handler: async (args: { alias: string }) => {
      const folders = await listFolders(args.alias);
      return { content: [{ type: "text" as const, text: JSON.stringify(folders, null, 2) }] };
    },
  },

  {
    name: "list_folders_tree",
    description:
      "List mail folders as a nested tree showing subfolders. Use depth=1 for top-level folders plus their immediate children, depth=2 to also include grandchildren. Each folder includes id, displayName, unreadItemCount, totalItemCount, and a children array.",
    schema: z.object({
      alias: z.string().describe("Mailbox alias"),
      depth: z
        .number()
        .int()
        .min(1)
        .max(2)
        .optional()
        .default(1)
        .describe("Subfolder depth to fetch: 1 = immediate children, 2 = two levels deep (default: 1)"),
    }),
    handler: async (args: { alias: string; depth?: number }) => {
      const depth = (args.depth ?? 1) as 1 | 2;
      const tree = await listFolderTree(args.alias, depth);
      return { content: [{ type: "text" as const, text: JSON.stringify(tree, null, 2) }] };
    },
  },

  {
    name: "find_folder",
    description:
      "Search for a mail folder by name using case-insensitive partial matching. Returns matching folders with their full path (e.g. 'Inbox > LinkedIn'). Use the returned folder id with list_messages to read its emails.",
    schema: z.object({
      alias: z.string().describe("Mailbox alias"),
      name: z.string().describe("Folder name to search for (partial match, case-insensitive)"),
      depth: z
        .number()
        .int()
        .min(1)
        .max(2)
        .optional()
        .default(2)
        .describe("How many folder levels to search: 1 = top-level + immediate children, 2 = two levels deep (default: 2)"),
    }),
    handler: async (args: { alias: string; name: string; depth?: number }) => {
      const depth = (args.depth ?? 2) as 1 | 2;
      const matches = await findFolder(args.alias, args.name, depth);
      const text =
        matches.length > 0
          ? JSON.stringify(matches, null, 2)
          : `No folders matching "${args.name}" found within depth ${depth}. Try a broader search term or use list_folders_tree to browse the full hierarchy.`;
      return { content: [{ type: "text" as const, text }] };
    },
  },
] as const;
