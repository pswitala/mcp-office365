import { listAliases } from "../auth/accounts.js";
import { listEvents } from "../graph/calendar.js";

export const CALENDAR_TODAY_URI = "o365://calendar/today";
export const CALENDAR_WEEK_URI = "o365://calendar/week";

function todayRange(): { start: string; end: string } {
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0);
  const end = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59);
  return {
    start: start.toISOString(),
    end: end.toISOString(),
  };
}

function weekRange(): { start: string; end: string } {
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0);
  const end = new Date(start.getTime() + 7 * 24 * 60 * 60 * 1000);
  return {
    start: start.toISOString(),
    end: end.toISOString(),
  };
}

async function readCalendarRange(range: { start: string; end: string }): Promise<string> {
  const aliases = await listAliases();
  if (aliases.length === 0) {
    return JSON.stringify({ error: "No mailboxes configured. Use the login tool to add one." });
  }

  const results: Record<string, unknown> = {};

  await Promise.all(
    aliases.map(async (alias) => {
      try {
        const { events } = await listEvents(alias, { start: range.start, end: range.end, top: 50 });
        results[alias] = events.map((e) => ({
          id: e.id,
          subject: e.subject,
          start: e.start,
          end: e.end,
          location: e.location,
          organizer: e.organizer,
          attendees: e.attendees,
        }));
      } catch (err) {
        results[alias] = { error: String(err) };
      }
    })
  );

  return JSON.stringify(results, null, 2);
}

export async function readCalendarTodayResource(): Promise<string> {
  return readCalendarRange(todayRange());
}

export async function readCalendarWeekResource(): Promise<string> {
  return readCalendarRange(weekRange());
}
