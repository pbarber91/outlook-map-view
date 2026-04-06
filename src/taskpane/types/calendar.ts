export type CalendarSource = {
  id: string;
  name: string;
  color?: string;
  isDefaultCalendar?: boolean;
};

export type MappedCalendarEvent = {
  id: string;
  subject: string;
  startIso: string;
  endIso: string;
  categories: string[];
  technicians: string[];
  addressText: string | null;
  latitude: number | null;
  longitude: number | null;
  webLink?: string;
  calendarId: string;
  calendarName: string;
};