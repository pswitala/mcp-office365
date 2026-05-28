import { graphGet, graphPost, getGraphClient } from "./client.js";

export interface MessageSummary {
  id: string;
  subject: string;
  from: string;
  receivedAt: string;
  isRead: boolean;
  preview: string;
  conversationId: string;
}

export interface MessageFull extends MessageSummary {
  to: string[];
  cc: string[];
  bodyContent: string;
  bodyContentType: string;
  hasAttachments: boolean;
}

export interface ListMessagesResult {
  messages: MessageSummary[];
  nextPageToken?: string;
}

const MESSAGE_SELECT =
  "id,subject,from,receivedDateTime,isRead,bodyPreview,conversationId";
const MESSAGE_FULL_SELECT =
  "id,subject,from,toRecipients,ccRecipients,receivedDateTime,isRead,bodyPreview,conversationId,body,hasAttachments";

function toSummary(m: Record<string, unknown>): MessageSummary {
  const from = m["from"] as { emailAddress?: { address?: string } } | undefined;
  return {
    id: m["id"] as string,
    subject: (m["subject"] as string) ?? "(no subject)",
    from: from?.emailAddress?.address ?? "",
    receivedAt: m["receivedDateTime"] as string,
    isRead: m["isRead"] as boolean,
    preview: m["bodyPreview"] as string,
    conversationId: m["conversationId"] as string,
  };
}

function toFull(m: Record<string, unknown>): MessageFull {
  const to = (m["toRecipients"] as Array<{ emailAddress?: { address?: string } }> ?? [])
    .map((r) => r.emailAddress?.address ?? "");
  const cc = (m["ccRecipients"] as Array<{ emailAddress?: { address?: string } }> ?? [])
    .map((r) => r.emailAddress?.address ?? "");
  const body = m["body"] as { content?: string; contentType?: string } | undefined;

  return {
    ...toSummary(m),
    to,
    cc,
    bodyContent: body?.content ?? "",
    bodyContentType: body?.contentType ?? "text",
    hasAttachments: m["hasAttachments"] as boolean,
  };
}

export async function listMessages(
  alias: string,
  opts: { folder?: string; search?: string; top?: number; pageToken?: string } = {}
): Promise<ListMessagesResult> {
  const { folder = "inbox", search, top = 25, pageToken } = opts;
  const limit = Math.min(top, 100);

  let url = `/me/mailFolders/${folder}/messages?$select=${MESSAGE_SELECT}&$top=${limit}&$orderby=receivedDateTime desc`;
  if (search) url += `&$search="${encodeURIComponent(search)}"`;
  if (pageToken) url = pageToken;

  const res = await graphGet<{ value: Record<string, unknown>[]; "@odata.nextLink"?: string }>(
    alias,
    url
  );

  return {
    messages: res.value.map(toSummary),
    nextPageToken: res["@odata.nextLink"],
  };
}

export async function getMessage(alias: string, id: string): Promise<MessageFull> {
  const m = await graphGet<Record<string, unknown>>(
    alias,
    `/me/messages/${id}?$select=${MESSAGE_FULL_SELECT}`
  );
  return toFull(m);
}

export async function searchMessages(
  alias: string,
  query: string,
  opts: { top?: number; pageToken?: string } = {}
): Promise<ListMessagesResult> {
  return listMessages(alias, { search: query, top: opts.top, pageToken: opts.pageToken });
}

export interface SendMessageOpts {
  to: string[];
  subject: string;
  body: string;
  bodyType?: "text" | "html";
  cc?: string[];
}

export async function sendMessage(alias: string, opts: SendMessageOpts): Promise<void> {
  const toRecipients = opts.to.map((a) => ({ emailAddress: { address: a } }));
  const ccRecipients = (opts.cc ?? []).map((a) => ({ emailAddress: { address: a } }));

  await graphPost(alias, "/me/sendMail", {
    message: {
      subject: opts.subject,
      body: { contentType: opts.bodyType ?? "text", content: opts.body },
      toRecipients,
      ccRecipients,
    },
  });
}

export async function replyMessage(
  alias: string,
  messageId: string,
  body: string,
  bodyType: "text" | "html" = "text"
): Promise<void> {
  await graphPost(alias, `/me/messages/${messageId}/reply`, {
    message: {},
    comment: body,
  });
}

export async function createDraft(alias: string, opts: SendMessageOpts): Promise<string> {
  const toRecipients = opts.to.map((a) => ({ emailAddress: { address: a } }));
  const ccRecipients = (opts.cc ?? []).map((a) => ({ emailAddress: { address: a } }));

  const res = await graphPost<{ id: string }>(alias, "/me/messages", {
    subject: opts.subject,
    body: { contentType: opts.bodyType ?? "text", content: opts.body },
    toRecipients,
    ccRecipients,
  });
  return res.id;
}

export async function moveMessage(alias: string, messageId: string, destinationFolderId: string): Promise<void> {
  await graphPost(alias, `/me/messages/${messageId}/move`, {
    destinationId: destinationFolderId,
  });
}

export async function listFolders(alias: string): Promise<Array<{ id: string; displayName: string; unreadItemCount: number; totalItemCount: number }>> {
  const res = await graphGet<{ value: Array<{ id: string; displayName: string; unreadItemCount: number; totalItemCount: number }> }>(
    alias,
    "/me/mailFolders?$select=id,displayName,unreadItemCount,totalItemCount&$top=50"
  );
  return res.value;
}

export interface FolderNode {
  id: string;
  displayName: string;
  unreadItemCount: number;
  totalItemCount: number;
  children: FolderNode[];
}

export interface FolderMatch {
  id: string;
  displayName: string;
  unreadItemCount: number;
  totalItemCount: number;
  path: string;
}

const FOLDER_SELECT = "id,displayName,unreadItemCount,totalItemCount";

async function getChildFolders(alias: string, folderId: string): Promise<Omit<FolderNode, "children">[]> {
  const res = await graphGet<{ value: Omit<FolderNode, "children">[] }>(
    alias,
    `/me/mailFolders/${folderId}/childFolders?$select=${FOLDER_SELECT}&$top=50`
  );
  return res.value;
}

export async function listFolderTree(alias: string, depth: 1 | 2 = 1): Promise<FolderNode[]> {
  const res = await graphGet<{ value: Omit<FolderNode, "children">[] }>(
    alias,
    `/me/mailFolders?$select=${FOLDER_SELECT}&$top=50`
  );

  return Promise.all(
    res.value.map(async (folder): Promise<FolderNode> => {
      const depth1 = await getChildFolders(alias, folder.id);
      let childNodes: FolderNode[];

      if (depth >= 2) {
        childNodes = await Promise.all(
          depth1.map(async (child): Promise<FolderNode> => {
            const depth2 = await getChildFolders(alias, child.id);
            return { ...child, children: depth2.map((c) => ({ ...c, children: [] })) };
          })
        );
      } else {
        childNodes = depth1.map((c) => ({ ...c, children: [] }));
      }

      return { ...folder, children: childNodes };
    })
  );
}

export async function findFolder(alias: string, name: string, depth: 1 | 2 = 2): Promise<FolderMatch[]> {
  const tree = await listFolderTree(alias, depth);
  const needle = name.toLowerCase();
  const matches: FolderMatch[] = [];

  function walk(nodes: FolderNode[], parentPath: string) {
    for (const node of nodes) {
      const path = parentPath ? `${parentPath} > ${node.displayName}` : node.displayName;
      if (node.displayName.toLowerCase().includes(needle)) {
        matches.push({ id: node.id, displayName: node.displayName, unreadItemCount: node.unreadItemCount, totalItemCount: node.totalItemCount, path });
      }
      if (node.children.length > 0) walk(node.children, path);
    }
  }

  walk(tree, "");
  return matches;
}
