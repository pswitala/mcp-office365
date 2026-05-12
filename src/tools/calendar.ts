import { z } from "zod";
import {
  listEvents,
  getEvent,
  createEvent,
  updateEvent,
  deleteEvent,
  findMeetingTimes,
  listCalendars,
  respondToEvent,
} from "../graph/calendar.js";

export const calendarToolDefinitions = [
  {
    name: "list_events",
    description: "List calendar events in a date range.",
    schema: z.object({
      alias: z.string().describe("Mailbox alias"),
      start: z.string().describe("Start datetime in ISO 8601 (e.g. 2024-01-01T00:00:00)"),
      end: z.string().describe("End datetime in ISO 8601"),
      calendarId: z.string().optional().describe("Calendar ID (omit for default calendar)"),
      top: z.number().int().min(1).max(100).optional().default(25),
      pageToken: z.string().optional(),
    }),
    handler: async (args: {
      alias: string;
      start: string;
      end: string;
      calendarId?: string;
      top?: number;
      pageToken?: string;
    }) => {
      const result = await listEvents(args.alias, {
        start: args.start,
        end: args.end,
        calendarId: args.calendarId,
        top: args.top,
        pageToken: args.pageToken,
      });
      return { content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }] };
    },
  },

  {
    name: "get_event",
    description: "Get full details of a calendar event by ID.",
    schema: z.object({
      alias: z.string(),
      id: z.string().describe("Event ID"),
    }),
    handler: async (args: { alias: string; id: string }) => {
      const event = await getEvent(args.alias, args.id);
      return { content: [{ type: "text" as const, text: JSON.stringify(event, null, 2) }] };
    },
  },

  {
    name: "create_event",
    description: "Create a new calendar event.",
    schema: z.object({
      alias: z.string(),
      subject: z.string(),
      start: z.string().describe("ISO 8601 datetime"),
      end: z.string().describe("ISO 8601 datetime"),
      timeZone: z.string().optional().default("UTC").describe("IANA timezone (e.g. America/New_York)"),
      attendees: z.array(z.string()).optional().describe("Attendee email addresses"),
      body: z.string().optional(),
      bodyType: z.enum(["text", "html"]).optional().default("text"),
      location: z.string().optional(),
      isAllDay: z.boolean().optional().default(false),
    }),
    handler: async (args: {
      alias: string;
      subject: string;
      start: string;
      end: string;
      timeZone?: string;
      attendees?: string[];
      body?: string;
      bodyType?: "text" | "html";
      location?: string;
      isAllDay?: boolean;
    }) => {
      const event = await createEvent(args.alias, {
        subject: args.subject,
        start: args.start,
        end: args.end,
        timeZone: args.timeZone,
        attendees: args.attendees,
        body: args.body,
        bodyType: args.bodyType,
        location: args.location,
        isAllDay: args.isAllDay,
      });
      return { content: [{ type: "text" as const, text: JSON.stringify(event, null, 2) }] };
    },
  },

  {
    name: "update_event",
    description: "Update an existing calendar event (partial update — only supply fields to change).",
    schema: z.object({
      alias: z.string(),
      id: z.string(),
      subject: z.string().optional(),
      start: z.string().optional(),
      end: z.string().optional(),
      timeZone: z.string().optional(),
      attendees: z.array(z.string()).optional(),
      body: z.string().optional(),
      bodyType: z.enum(["text", "html"]).optional(),
      location: z.string().optional(),
    }),
    handler: async (args: {
      alias: string;
      id: string;
      subject?: string;
      start?: string;
      end?: string;
      timeZone?: string;
      attendees?: string[];
      body?: string;
      bodyType?: "text" | "html";
      location?: string;
    }) => {
      const { alias, id, ...opts } = args;
      const event = await updateEvent(alias, id, opts);
      return { content: [{ type: "text" as const, text: JSON.stringify(event, null, 2) }] };
    },
  },

  {
    name: "delete_event",
    description: "Delete (cancel) a calendar event.",
    schema: z.object({
      alias: z.string(),
      id: z.string(),
    }),
    handler: async (args: { alias: string; id: string }) => {
      await deleteEvent(args.alias, args.id);
      return { content: [{ type: "text" as const, text: "Event deleted." }] };
    },
  },

  {
    name: "find_meeting_times",
    description: "Find available meeting times for a set of attendees.",
    schema: z.object({
      alias: z.string(),
      attendees: z.array(z.string()).describe("Attendee email addresses"),
      durationMinutes: z.number().int().min(1).describe("Meeting duration in minutes"),
      startHint: z.string().describe("Earliest possible start (ISO 8601)"),
      endHint: z.string().describe("Latest possible end (ISO 8601)"),
      timeZone: z.string().optional().default("UTC"),
    }),
    handler: async (args: {
      alias: string;
      attendees: string[];
      durationMinutes: number;
      startHint: string;
      endHint: string;
      timeZone?: string;
    }) => {
      const suggestions = await findMeetingTimes(args.alias, {
        attendees: args.attendees,
        durationMinutes: args.durationMinutes,
        startHint: args.startHint,
        endHint: args.endHint,
        timeZone: args.timeZone,
      });
      return { content: [{ type: "text" as const, text: JSON.stringify(suggestions, null, 2) }] };
    },
  },

  {
    name: "list_calendars",
    description: "List all calendars for a mailbox.",
    schema: z.object({
      alias: z.string(),
    }),
    handler: async (args: { alias: string }) => {
      const calendars = await listCalendars(args.alias);
      return { content: [{ type: "text" as const, text: JSON.stringify(calendars, null, 2) }] };
    },
  },

  {
    name: "respond_to_event",
    description: "Accept, tentatively accept, or decline a calendar invitation.",
    schema: z.object({
      alias: z.string(),
      id: z.string().describe("Event ID"),
      response: z.enum(["accept", "tentativelyAccept", "decline"]),
      comment: z.string().optional(),
    }),
    handler: async (args: { alias: string; id: string; response: "accept" | "tentativelyAccept" | "decline"; comment?: string }) => {
      await respondToEvent(args.alias, args.id, args.response, args.comment);
      return { content: [{ type: "text" as const, text: `Response "${args.response}" sent.` }] };
    },
  },
] as const;
