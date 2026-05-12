import { listAliases } from "../auth/accounts.js";
import { listMessages } from "../graph/mail.js";

export const INBOX_RESOURCE_URI = "o365://inbox";

export async function readInboxResource(): Promise<string> {
  const aliases = await listAliases();
  if (aliases.length === 0) {
    return JSON.stringify({ error: "No mailboxes configured. Use the login tool to add one." });
  }

  const results: Record<string, unknown> = {};

  await Promise.all(
    aliases.map(async (alias) => {
      try {
        const { messages } = await listMessages(alias, {
          folder: "inbox",
          top: 10,
          search: "isRead:false",
        });
        results[alias] = messages.map((m) => ({
          id: m.id,
          subject: m.subject,
          from: m.from,
          receivedAt: m.receivedAt,
          preview: m.preview,
        }));
      } catch (err) {
        results[alias] = { error: String(err) };
      }
    })
  );

  return JSON.stringify(results, null, 2);
}
