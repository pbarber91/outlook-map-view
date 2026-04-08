import * as React from "react";
import { CalendarSource, MappedCalendarEvent } from "../types/calendar";
import { getCalendarEventsForRange } from "../services/calendarService";
import { getDateRange } from "../utils/dateRanges";
import { FilterState } from "../types/filters";
import { getAvailableCalendars } from "../services/calendarSourceService";

export function useCalendarEvents(filters: FilterState, refreshKey: number) {
  const [events, setEvents] = React.useState<MappedCalendarEvent[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string | null>(null);
  const [availableCalendars, setAvailableCalendars] = React.useState<CalendarSource[]>([]);
  const [calendarsLoading, setCalendarsLoading] = React.useState<boolean>(true);
  const [calendarsError, setCalendarsError] = React.useState<string | null>(null);

  React.useEffect(() => {
    let active = true;

    async function loadCalendars() {
      setCalendarsLoading(true);
      setCalendarsError(null);

      try {
        const calendars = await getAvailableCalendars();
        if (!active) return;
        setAvailableCalendars(calendars);
      } catch (err) {
        if (!active) return;
        setAvailableCalendars([]);
        setCalendarsError(
          err instanceof Error ? err.message : "Failed to load calendars."
        );
      } finally {
        if (active) {
          setCalendarsLoading(false);
        }
      }
    }

    loadCalendars();

    return () => {
      active = false;
    };
  }, [refreshKey]);

  React.useEffect(() => {
    let active = true;

    async function load() {
      setLoading(true);
      setError(null);

      try {
        const range = getDateRange(filters.preset, filters.startDate, filters.endDate);

        const result = await getCalendarEventsForRange(
          range.startIso,
          range.endIso,
          filters.calendarIds
        );

        if (!active) return;
        setEvents(result);
      } catch (err) {
        if (!active) return;
        setError(err instanceof Error ? err.message : "Failed to load calendar events.");
        setEvents([]);
      } finally {
        if (active) {
          setLoading(false);
        }
      }
    }

    if (!calendarsLoading) {
      load();
    }

    return () => {
      active = false;
    };
  }, [
    filters.preset,
    filters.startDate,
    filters.endDate,
    filters.calendarIds,
    refreshKey,
    calendarsLoading,
  ]);

  const availableCategories = React.useMemo(() => {
    const seen = new Set<string>();

    events.forEach((event) => {
      event.categories.forEach((category) => {
        seen.add(category);
      });
    });

    return Array.from(seen).sort((a, b) => a.localeCompare(b));
  }, [events]);

  const availableTechnicians = React.useMemo(() => {
    const seen = new Set<string>();

    events.forEach((event) => {
      event.technicians.forEach((technician) => {
        seen.add(technician);
      });
    });

    return Array.from(seen).sort((a, b) => a.localeCompare(b));
  }, [events]);

  const filteredEvents = React.useMemo(() => {
    return events.filter((event) => {
      const categoryMatch =
        filters.categories.length === 0 ||
        event.categories.some((category) => filters.categories.includes(category));

      const technicianMatch =
        filters.technicians.length === 0 ||
        event.technicians.some((technician) => filters.technicians.includes(technician));

      return categoryMatch && technicianMatch;
    });
  }, [events, filters.categories, filters.technicians]);

  const mappableEvents = React.useMemo(() => {
    return filteredEvents.filter((event) => !!event.addressText);
  }, [filteredEvents]);

  return {
    events,
    filteredEvents,
    mappableEvents,
    availableCategories,
    availableTechnicians,
    availableCalendars,
    calendarsLoading,
    calendarsError,
    loading,
    error,
  };
}