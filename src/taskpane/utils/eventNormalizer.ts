import { MappedCalendarEvent } from "../types/calendar";

type RawEventLike = {
  id?: string;
  subject?: string;
  start?: string;
  end?: string;
  location?: string | null;
  categories?: string[];
  technicians?: string[];
  webLink?: string;
  calendarId?: string;
  calendarName?: string;
};

const NON_MAPPABLE_LOCATION_PATTERNS: RegExp[] = [
  /microsoft teams/i,
  /teams meeting/i,
  /zoom/i,
  /google meet/i,
  /webex/i,
  /conference call/i,
  /phone call/i,
  /call/i,
  /virtual/i,
  /online/i,
  /remote/i,
  /teleconference/i,
  /tbd/i,
  /n\/a/i,
  /^none$/i,
];

function cleanWhitespace(value: string): string {
  return value.replace(/\s+/g, " ").trim();
}

function looksMappableLocation(value: string): boolean {
  const cleaned = cleanWhitespace(value);

  if (!cleaned || cleaned.length < 6) {
    return false;
  }

  return !NON_MAPPABLE_LOCATION_PATTERNS.some((pattern) => pattern.test(cleaned));
}

function normalizeLocation(value?: string | null): string | null {
  if (typeof value !== "string") {
    return null;
  }

  const cleaned = cleanWhitespace(value);

  if (!looksMappableLocation(cleaned)) {
    return null;
  }

  return cleaned;
}

function normalizeStringArray(value?: string[]): string[] {
  if (!Array.isArray(value)) {
    return [];
  }

  const seen = new Set<string>();

  return value
    .filter((item): item is string => typeof item === "string")
    .map((item) => cleanWhitespace(item))
    .filter((item) => item.length > 0)
    .filter((item) => {
      const key = item.toLowerCase();
      if (seen.has(key)) {
        return false;
      }
      seen.add(key);
      return true;
    });
}

export function normalizeEvent(raw: RawEventLike, fallbackId: string): MappedCalendarEvent {
  return {
    id: typeof raw.id === "string" && raw.id.trim().length > 0 ? raw.id : fallbackId,
    subject:
      typeof raw.subject === "string" && raw.subject.trim().length > 0
        ? cleanWhitespace(raw.subject)
        : "(No subject)",
    startIso: typeof raw.start === "string" ? raw.start : "",
    endIso: typeof raw.end === "string" ? raw.end : "",
    categories: normalizeStringArray(raw.categories),
    technicians: normalizeStringArray(raw.technicians),
    addressText: normalizeLocation(raw.location),
    latitude: null,
    longitude: null,
    webLink: typeof raw.webLink === "string" ? raw.webLink : "",
    calendarId:
      typeof raw.calendarId === "string" && raw.calendarId.trim().length > 0
        ? raw.calendarId
        : "unknown-calendar",
    calendarName:
      typeof raw.calendarName === "string" && raw.calendarName.trim().length > 0
        ? cleanWhitespace(raw.calendarName)
        : "Unknown Calendar",
  };
}