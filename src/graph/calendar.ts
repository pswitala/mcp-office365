import { graphGet, graphPost, graphPatch, graphDelete } from "./client.js";

export interface EventSummary {
  id: string;
  subject: string;
  start: string;
  end: string;
  location: string;
  isAllDay: boolean;
  organizer: string;
  attendees: string[];
  isCancelled: boolean;
  webLink: string;
}

export interface ListEventsResult {
  events: EventSummary[];
  nextPageToken?: string;
}

const EVENT_SELECT =
  "id,subject,start,end,location,isAllDay,organizer,attendees,isCancelled,webLink";

function toSummary(e: Record<string, unknown>): EventSummary {
  const start = e["start"] as { dateTime?: string; date?: string } | undefined;
  const end = e["end"] as { dateTime?: string; date?: string } | undefined;
  const loc = e["location"] as { displayName?: string } | undefined;
  const org = e["organizer"] as { emailAddress?: { address?: string } } | undefined;
  const attendees = (e["attendees"] as Array<{ emailAddress?: { address?: string } }> ?? [])
    .map((a) => a.emailAddress?.address ?? "");

  return {
    id: e["id"] as string,
    subject: (e["subject"] as string) ?? "(no subject)",
    start: start?.dateTime ?? start?.date ?? "",
    end: end?.dateTime ?? end?.date ?? "",
    location: loc?.displayName ?? "",
    isAllDay: (e["isAllDay"] as boolean) ?? false,
    organizer: org?.emailAddress?.address ?? "",
    attendees,
    isCancelled: (e["isCancelled"] as boolean) ?? false,
    webLink: (e["webLink"] as string) ?? "",
  };
}

export async function listEvents(
  alias: string,
  opts: { start: string; end: string; calendarId?: string; top?: number; pageToken?: string }
): Promise<ListEventsResult> {
  const { start, end, calendarId, top = 25, pageToken } = opts;
  const limit = Math.min(top, 100);

  const base = calendarId ? `/me/calendars/${calendarId}/events` : "/me/events";
  let url = pageToken ?? `${base}?$select=${EVENT_SELECT}&$top=${limit}&$orderby=start/dateTime asc&$filter=start/dateTime ge '${start}' and end/dateTime le '${end}'`;

  const res = await graphGet<{ value: Record<string, unknown>[]; "@odata.nextLink"?: string }>(
    alias,
    url
  );

  return {
    events: res.value.map(toSummary),
    nextPageToken: res["@odata.nextLink"],
  };
}

export async function getEvent(alias: string, id: string): Promise<EventSummary & { body: string; bodyType: string }> {
  const e = await graphGet<Record<string, unknown>>(
    alias,
    `/me/events/${id}?$select=${EVENT_SELECT},body`
  );
  const body = e["body"] as { content?: string; contentType?: string } | undefined;
  return {
    ...toSummary(e),
    body: body?.content ?? "",
    bodyType: body?.contentType ?? "text",
  };
}

export interface CreateEventOpts {
  subject: string;
  start: string;
  end: string;
  timeZone?: string;
  attendees?: string[];
  body?: string;
  bodyType?: "text" | "html";
  location?: string;
  isAllDay?: boolean;
}

export async function createEvent(alias: string, opts: CreateEventOpts): Promise<EventSummary> {
  const tz = opts.timeZone ?? "UTC";
  const payload: Record<string, unknown> = {
    subject: opts.subject,
    start: { dateTime: opts.start, timeZone: tz },
    end: { dateTime: opts.end, timeZone: tz },
    isAllDay: opts.isAllDay ?? false,
  };
  if (opts.body) payload["body"] = { contentType: opts.bodyType ?? "text", content: opts.body };
  if (opts.location) payload["location"] = { displayName: opts.location };
  if (opts.attendees?.length) {
    payload["attendees"] = opts.attendees.map((a) => ({
      emailAddress: { address: a },
      type: "required",
    }));
  }
  const res = await graphPost<Record<string, unknown>>(alias, "/me/events", payload);
  return toSummary(res);
}

export async function updateEvent(
  alias: string,
  id: string,
  opts: Partial<CreateEventOpts>
): Promise<EventSummary> {
  const tz = opts.timeZone ?? "UTC";
  const payload: Record<string, unknown> = {};
  if (opts.subject !== undefined) payload["subject"] = opts.subject;
  if (opts.start !== undefined) payload["start"] = { dateTime: opts.start, timeZone: tz };
  if (opts.end !== undefined) payload["end"] = { dateTime: opts.end, timeZone: tz };
  if (opts.body !== undefined) payload["body"] = { contentType: opts.bodyType ?? "text", content: opts.body };
  if (opts.location !== undefined) payload["location"] = { displayName: opts.location };
  if (opts.attendees !== undefined) {
    payload["attendees"] = opts.attendees.map((a) => ({
      emailAddress: { address: a },
      type: "required",
    }));
  }
  const res = await graphPatch<Record<string, unknown>>(alias, `/me/events/${id}`, payload);
  return toSummary(res);
}

export async function deleteEvent(alias: string, id: string): Promise<void> {
  await graphDelete(alias, `/me/events/${id}`);
}

export interface FindMeetingTimesOpts {
  attendees: string[];
  durationMinutes: number;
  startHint: string;
  endHint: string;
  timeZone?: string;
}

export async function findMeetingTimes(
  alias: string,
  opts: FindMeetingTimesOpts
): Promise<Array<{ start: string; end: string; confidence: number }>> {
  const res = await graphPost<{
    meetingTimeSuggestions: Array<{
      meetingTimeSlot: { start: { dateTime: string }; end: { dateTime: string } };
      confidence: number;
    }>;
  }>(alias, "/me/findMeetingTimes", {
    attendees: opts.attendees.map((a) => ({
      emailAddress: { address: a },
      type: "required",
    })),
    timeConstraint: {
      timeslots: [
        {
          start: { dateTime: opts.startHint, timeZone: opts.timeZone ?? "UTC" },
          end: { dateTime: opts.endHint, timeZone: opts.timeZone ?? "UTC" },
        },
      ],
    },
    meetingDuration: `PT${opts.durationMinutes}M`,
  });

  return (res.meetingTimeSuggestions ?? []).map((s) => ({
    start: s.meetingTimeSlot.start.dateTime,
    end: s.meetingTimeSlot.end.dateTime,
    confidence: s.confidence,
  }));
}

export async function listCalendars(alias: string): Promise<Array<{ id: string; name: string; isDefault: boolean; color: string }>> {
  const res = await graphGet<{ value: Array<{ id: string; name: string; isDefaultCalendar: boolean; color: string }> }>(
    alias,
    "/me/calendars?$select=id,name,isDefaultCalendar,color"
  );
  return res.value.map((c) => ({
    id: c.id,
    name: c.name,
    isDefault: c.isDefaultCalendar,
    color: c.color,
  }));
}

export async function respondToEvent(
  alias: string,
  id: string,
  response: "accept" | "tentativelyAccept" | "decline",
  comment?: string
): Promise<void> {
  await graphPost(alias, `/me/events/${id}/${response}`, { comment: comment ?? "" });
}
