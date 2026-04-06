import { CalendarSource } from "../types/calendar";
import { getGraphClient } from "./graphClient";

type GraphCalendar = {
  id?: string;
  name?: string;
  color?: string;
  isDefaultCalendar?: boolean;
};

type GraphCalendarsResponse = {
  value?: GraphCalendar[];
};

export async function getAvailableCalendars(): Promise<CalendarSource[]> {
  const client = await getGraphClient();

  const response = (await client.api("/me/calendars").get()) as GraphCalendarsResponse;
  const calendars = Array.isArray(response?.value) ? response.value : [];

  return calendars
    .filter((calendar) => typeof calendar.id === "string" && typeof calendar.name === "string")
    .map((calendar) => ({
      id: calendar.id as string,
      name: calendar.name as string,
      color: typeof calendar.color === "string" ? calendar.color : undefined,
      isDefaultCalendar: Boolean(calendar.isDefaultCalendar),
    }))
    .sort((a, b) => {
      if (a.isDefaultCalendar && !b.isDefaultCalendar) return -1;
      if (!a.isDefaultCalendar && b.isDefaultCalendar) return 1;
      return a.name.localeCompare(b.name);
    });
}