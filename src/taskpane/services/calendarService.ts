import { MappedCalendarEvent } from "../types/calendar";
import { normalizeEvent } from "../utils/eventNormalizer";
import { matchTechnician } from "../utils/technicianMatcher";
import { getGraphClient } from "./graphClient";

type GraphAttendee = {
  type?: string;
  emailAddress?: {
    name?: string;
    address?: string;
  };
};

type GraphEvent = {
  id?: string;
  subject?: string;
  categories?: string[];
  webLink?: string;
  isAllDay?: boolean;
  attendees?: GraphAttendee[];
  start?: {
    dateTime?: string;
    timeZone?: string;
  };
  end?: {
    dateTime?: string;
    timeZone?: string;
  };
  location?: {
    displayName?: string;
    locationUri?: string;
    uniqueId?: string;
    uniqueIdType?: string;
  };
  locations?: Array<{
    displayName?: string;
    locationUri?: string;
    uniqueId?: string;
    uniqueIdType?: string;
  }>;
  onlineMeeting?: {
    joinUrl?: string;
  };
  onlineMeetingUrl?: string;
};

type GraphCalendarViewResponse = {
  value?: GraphEvent[];
  "@odata.nextLink"?: string;
};

const PAGE_SIZE = 100;
const MAX_PAGES = 50;
const MAX_EVENTS = 5000;

function safeToIso(dateTime?: string): string {
  if (!dateTime) {
    return "";
  }

  const parsed = new Date(dateTime);
  return Number.isNaN(parsed.getTime()) ? "" : parsed.toISOString();
}

function firstNonEmptyString(...values: Array<string | undefined | null>): string | null {
  for (const value of values) {
    if (typeof value === "string" && value.trim().length > 0) {
      return value.trim();
    }
  }
  return null;
}

function extractLocation(event: GraphEvent): string | null {
  const primary = firstNonEmptyString(
    event.location?.displayName,
    event.location?.locationUri,
    event.location?.uniqueId
  );

  if (primary) {
    return primary;
  }

  if (Array.isArray(event.locations)) {
    for (const location of event.locations) {
      const candidate = firstNonEmptyString(
        location.displayName,
        location.locationUri,
        location.uniqueId
      );

      if (candidate) {
        return candidate;
      }
    }
  }

  return null;
}

function extractRequiredTechnicians(event: GraphEvent): string[] {
  if (!Array.isArray(event.attendees)) {
    return [];
  }

  const seen = new Set<string>();
  const technicians: string[] = [];

  event.attendees.forEach((attendee) => {
    if ((attendee.type || "").toLowerCase() !== "required") {
      return;
    }

    const matched = matchTechnician(
      attendee.emailAddress?.name ?? null,
      attendee.emailAddress?.address ?? null
    );

    if (!matched) {
      return;
    }

    const key = matched.toLowerCase();
    if (seen.has(key)) {
      return;
    }

    seen.add(key);
    technicians.push(matched);
  });

  return technicians;
}

function normalizeGraphEvent(event: GraphEvent): MappedCalendarEvent {
  return normalizeEvent(
    {
      id: event.id,
      subject: event.subject,
      start: safeToIso(event.start?.dateTime),
      end: safeToIso(event.end?.dateTime),
      location: extractLocation(event),
      categories: Array.isArray(event.categories) ? event.categories : [],
      technicians: extractRequiredTechnicians(event),
      webLink: event.webLink,
    },
    crypto.randomUUID()
  );
}

function compareEventsByStart(a: MappedCalendarEvent, b: MappedCalendarEvent): number {
  const aTime = a.startIso ? new Date(a.startIso).getTime() : Number.MAX_SAFE_INTEGER;
  const bTime = b.startIso ? new Date(b.startIso).getTime() : Number.MAX_SAFE_INTEGER;

  if (aTime !== bTime) {
    return aTime - bTime;
  }

  return a.subject.localeCompare(b.subject);
}

async function getCalendarViewPage(
  startIso: string,
  endIso: string,
  nextLink?: string
): Promise<GraphCalendarViewResponse> {
  const client = await getGraphClient();

  if (nextLink) {
    return (await client.api(nextLink).get()) as GraphCalendarViewResponse;
  }

  return (await client
    .api("/me/calendarView")
    .query({
      startDateTime: startIso,
      endDateTime: endIso,
      $top: String(PAGE_SIZE),
      $select:
        "id,subject,start,end,location,locations,categories,webLink,isAllDay,onlineMeeting,onlineMeetingUrl,attendees",
      $orderby: "start/dateTime",
    })
    .get()) as GraphCalendarViewResponse;
}

export async function getCalendarEventsForRange(
  startIso: string,
  endIso: string
): Promise<MappedCalendarEvent[]> {
  const allEvents: GraphEvent[] = [];
  const seenIds = new Set<string>();

  let nextLink: string | undefined;
  let pageCount = 0;

  do {
    const response = await getCalendarViewPage(startIso, endIso, nextLink);
    const pageEvents = Array.isArray(response?.value) ? response.value : [];

    for (const event of pageEvents) {
      const id = typeof event.id === "string" ? event.id : "";

      if (id && seenIds.has(id)) {
        continue;
      }

      if (id) {
        seenIds.add(id);
      }

      allEvents.push(event);

      if (allEvents.length >= MAX_EVENTS) {
        break;
      }
    }

    pageCount += 1;
    nextLink = response?.["@odata.nextLink"];

    if (allEvents.length >= MAX_EVENTS) {
      break;
    }

    if (pageCount >= MAX_PAGES) {
      break;
    }
  } while (nextLink);

  return allEvents.map(normalizeGraphEvent).sort(compareEventsByStart);
}