import * as React from "react";
import FilterBar from "./components/FilterBar";
import EventList from "./components/EventList";
import EventMap from "./components/EventMap";
import { FilterState } from "./types/filters";
import { useCalendarEvents } from "./hooks/useCalendarEvents";
import {
  geocodeEvents,
  GeocodedCalendarEvent,
  getGeocodeCacheSize,
} from "./services/geocodeService";
import {
  getRoutePreviewsByDay,
  RouteAnchor,
  RoutePreview,
} from "./services/routeService";
import {
  completeRedirectIfNeeded,
  ensureGraphAccessInteractiveRedirect,
} from "./services/graphClient";

declare const __MAPBOX_ACCESS_TOKEN__: string;

type MapMode = "routeDay" | "allFiltered";
type StartMode = "homeOffice" | "lastInspection";
type EndMode = "returnOffice" | "lastStop";

type DayRouteSetting = {
  startMode: StartMode;
  endMode: EndMode;
};

type TechnicianDaySummary = {
  technician: string;
  dayKey: string;
  dayLabel: string;
  stops: number;
  eventSeconds: number;
  driveSeconds: number;
  fieldSeconds: number;
  driveMeters: number;
  calendars: string[];
  categories: string[];
};

type AdjustedRoutePreview = RoutePreview & {
  startMode: StartMode;
  endMode: EndMode;
  adjustedDriveSeconds: number;
  adjustedDistanceMeters: number;
  adjustedFieldSeconds: number;
  firstLegSeconds: number;
  firstLegMeters: number;
  lastLegSeconds: number;
  lastLegMeters: number;
  firstEventId: string | null;
  lastEventId: string | null;
};

type ThemeMode = "light" | "dark";

function getTheme(mode: ThemeMode) {
  if (mode === "dark") {
    return {
      appBg: "#020617",
      shellBg: "linear-gradient(135deg, #0f172a 0%, #111827 100%)",
      shellBorder: "#334155",
      panelBg: "#111827",
      panelAltBg: "#0b1220",
      panelSoftBg: "#0f172a",
      panelSelectedBg: "rgba(37,99,235,0.18)",
      panelSelectedBorder: "#60a5fa",
      border: "#334155",
      borderSoft: "#475569",
      text: "#f8fafc",
      textMuted: "#94a3b8",
      textSoft: "#cbd5e1",
      controlBg: "#111827",
      chipBg: "#0f172a",
      chipBorder: "#334155",
      emptyBg: "#0f172a",
      shadow: "0 16px 36px rgba(2, 6, 23, 0.32)",
    };
  }

  return {
    appBg: "#f8fafc",
    shellBg: "linear-gradient(135deg, #ffffff 0%, #f8fbff 100%)",
    shellBorder: "#dbe2ea",
    panelBg: "#ffffff",
    panelAltBg: "#ffffff",
    panelSoftBg: "#f8fafc",
    panelSelectedBg: "#eff6ff",
    panelSelectedBorder: "#2563eb",
    border: "#d1d5db",
    borderSoft: "#cbd5e1",
    text: "#0f172a",
    textMuted: "#64748b",
    textSoft: "#475569",
    controlBg: "#ffffff",
    chipBg: "#ffffff",
    chipBorder: "#dbe2ea",
    emptyBg: "#f8fafc",
    shadow: "0 2px 8px rgba(15,23,42,0.04)",
  };
}

const initialFilters: FilterState = {
  preset: "thisWeek",
  startDate: "",
  endDate: "",
  categories: [],
  technicians: [],
  calendarIds: [],
};

const HOME_OFFICE: RouteAnchor = {
  id: "home-office",
  label: "Home Office",
  latitude: 26.895257,
  longitude: -82.00761,
};

const STATIC_MAP_STYLE = "mapbox/streets-v12";

function getDayKeyFromIso(iso: string): string | null {
  const date = new Date(iso);
  if (Number.isNaN(date.getTime())) {
    return null;
  }

  const year = date.getFullYear();
  const month = `${date.getMonth() + 1}`.padStart(2, "0");
  const day = `${date.getDate()}`.padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function getDefaultDayRouteSetting(): DayRouteSetting {
  return {
    startMode: "homeOffice",
    endMode: "returnOffice",
  };
}

function formatDuration(totalSeconds: number): string {
  const roundedMinutes = Math.round(totalSeconds / 60);
  const hours = Math.floor(roundedMinutes / 60);
  const minutes = roundedMinutes % 60;

  if (hours <= 0) {
    return `${minutes}m`;
  }

  return `${hours}h ${minutes}m`;
}

function formatMiles(totalMeters: number): string {
  const miles = totalMeters / 1609.344;
  return `${miles.toFixed(1)} mi`;
}

function formatDateForCsv(iso: string): string {
  const date = new Date(iso);
  if (Number.isNaN(date.getTime())) {
    return "";
  }

  return new Intl.DateTimeFormat("en-US", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(date);
}

function formatTimeForCsv(iso: string): string {
  const date = new Date(iso);
  if (Number.isNaN(date.getTime())) {
    return "";
  }

  return new Intl.DateTimeFormat("en-US", {
    hour: "numeric",
    minute: "2-digit",
  }).format(date);
}

function formatDateLong(iso: string): string {
  const date = new Date(iso);
  if (Number.isNaN(date.getTime())) {
    return "";
  }

  return new Intl.DateTimeFormat("en-US", {
    weekday: "short",
    month: "short",
    day: "numeric",
    year: "numeric",
  }).format(date);
}

function getEventDurationSeconds(startIso: string, endIso: string): number {
  const start = new Date(startIso);
  const end = new Date(endIso);

  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) {
    return 0;
  }

  const durationSeconds = Math.max(0, Math.round((end.getTime() - start.getTime()) / 1000));

  const isAllDay =
    start.getHours() === 0 &&
    start.getMinutes() === 0 &&
    end.getHours() === 0 &&
    end.getMinutes() === 0 &&
    durationSeconds >= 23 * 60 * 60;

  if (isAllDay) {
    return 0;
  }

  return durationSeconds;
}

function csvEscape(value: string | number | boolean | null | undefined): string {
  const stringValue = value == null ? "" : String(value);
  return `"${stringValue.replace(/"/g, '""')}"`;
}

function downloadCsv(
  filename: string,
  headers: string[],
  rows: Array<Array<string | number | boolean | null | undefined>>
) {
  const csv = [
    headers.map(csvEscape).join(","),
    ...rows.map((row) => row.map(csvEscape).join(",")),
  ].join("\n");

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);

  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);

  URL.revokeObjectURL(url);
}

function getTodayStamp(): string {
  const now = new Date();
  const year = now.getFullYear();
  const month = `${now.getMonth() + 1}`.padStart(2, "0");
  const day = `${now.getDate()}`.padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function getMapboxToken(): string {
  return typeof __MAPBOX_ACCESS_TOKEN__ === "string" ? __MAPBOX_ACCESS_TOKEN__ : "";
}

function buildStaticMapUrl(
  events: GeocodedCalendarEvent[],
  _route: RoutePreview | null
): string | null {
  const token = getMapboxToken();
  if (!token || events.length === 0) {
    return null;
  }

  const markerOverlay = events
    .slice(0, 40)
    .map((event) => `pin-s+2563eb(${event.longitude},${event.latitude})`)
    .join(",");

  return `https://api.mapbox.com/styles/v1/${STATIC_MAP_STYLE}/static/${markerOverlay}/auto/760x320?padding=36&access_token=${encodeURIComponent(
    token
  )}`;
}

function uniqueSorted(values: string[]): string[] {
  return Array.from(new Set(values.filter(Boolean))).sort((a, b) => a.localeCompare(b));
}

function getStartModeLabel(value: StartMode): string {
  return value === "homeOffice" ? "Home office" : "Last inspection";
}

function getEndModeLabel(value: EndMode): string {
  return value === "returnOffice" ? "Return home/office" : "End at last stop";
}

export default function App() {
  const standalone =
    typeof window !== "undefined" &&
    new URLSearchParams(window.location.search).get("standalone") === "1";

  const [filters, setFilters] = React.useState<FilterState>(initialFilters);
  const [selectedEventId, setSelectedEventId] = React.useState<string | null>(null);
  const [refreshKey, setRefreshKey] = React.useState<number>(0);
  const [geocodedEvents, setGeocodedEvents] = React.useState<GeocodedCalendarEvent[]>([]);
  const [mapLoading, setMapLoading] = React.useState<boolean>(false);
  const [mapError, setMapError] = React.useState<string | null>(null);
  const [showOnlyMappable, setShowOnlyMappable] = React.useState<boolean>(false);
  const [showRouteOverlay, setShowRouteOverlay] = React.useState<boolean>(true);
  const [routeLoading, setRouteLoading] = React.useState<boolean>(false);
  const [routeError, setRouteError] = React.useState<string | null>(null);
  const [routePreviews, setRoutePreviews] = React.useState<RoutePreview[]>([]);
  const [authWorking, setAuthWorking] = React.useState<boolean>(false);
  const [mapMode, setMapMode] = React.useState<MapMode>("routeDay");
  const [showTechnicianDashboard, setShowTechnicianDashboard] = React.useState<boolean>(false);
  const [dayRouteSettings, setDayRouteSettings] = React.useState<Record<string, DayRouteSetting>>({});

  const [themeMode, setThemeMode] = React.useState<ThemeMode>("light");

  React.useEffect(() => {
    try {
      const saved = window.localStorage.getItem("outlook-map-view-theme");
      if (saved === "light" || saved === "dark") {
        setThemeMode(saved);
        return;
      }
    } catch {}

    if (window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches) {
      setThemeMode("dark");
    }
  }, []);

  React.useEffect(() => {
    try {
      window.localStorage.setItem("outlook-map-view-theme", themeMode);
    } catch {}
  }, [themeMode]);

  const theme = React.useMemo(() => getTheme(themeMode), [themeMode]);

  const {
    filteredEvents,
    mappableEvents,
    availableCategories,
    availableTechnicians,
    availableCalendars,
    calendarsError,
    loading,
    error,
  } = useCalendarEvents(filters, refreshKey);

  React.useEffect(() => {
    if (availableCalendars.length === 0) {
      return;
    }

    if (filters.calendarIds.length > 0) {
      return;
    }

    const defaultCalendars = availableCalendars
      .filter((calendar) => calendar.isDefaultCalendar)
      .map((calendar) => calendar.id);

    const fallbackCalendars =
      defaultCalendars.length > 0 ? defaultCalendars : [availableCalendars[0].id];

    setFilters((prev) => ({
      ...prev,
      calendarIds: fallbackCalendars,
    }));
  }, [availableCalendars, filters.calendarIds.length]);

  const visibleEvents = React.useMemo(() => {
    return showOnlyMappable
      ? filteredEvents.filter((event) => !!event.addressText)
      : filteredEvents;
  }, [filteredEvents, showOnlyMappable]);

  React.useEffect(() => {
    if (selectedEventId && visibleEvents.some((event) => event.id === selectedEventId)) {
      return;
    }

    setSelectedEventId(visibleEvents[0]?.id ?? null);
  }, [visibleEvents, selectedEventId]);

  React.useEffect(() => {
    let active = true;

    async function loadMapEvents() {
      setMapLoading(true);
      setMapError(null);

      try {
        const sourceEvents = showOnlyMappable
          ? visibleEvents.filter((event) => !!event.addressText)
          : mappableEvents;

        const result = await geocodeEvents(sourceEvents);

        if (!active) return;
        setGeocodedEvents(result);
      } catch (err) {
        if (!active) return;
        setMapError(err instanceof Error ? err.message : "Failed to geocode map events.");
        setGeocodedEvents([]);
      } finally {
        if (active) {
          setMapLoading(false);
        }
      }
    }

    loadMapEvents();

    return () => {
      active = false;
    };
  }, [mappableEvents, visibleEvents, showOnlyMappable]);

  React.useEffect(() => {
    let active = true;

    async function loadRoutes() {
      if (!showRouteOverlay) {
        setRoutePreviews([]);
        setRouteError(null);
        setRouteLoading(false);
        return;
      }

      setRouteLoading(true);
      setRouteError(null);

      try {
        const result = await getRoutePreviewsByDay(
          geocodedEvents,
          HOME_OFFICE,
          HOME_OFFICE
        );

        if (!active) return;
        setRoutePreviews(result);
      } catch (err) {
        if (!active) return;
        setRouteError(err instanceof Error ? err.message : "Failed to build route preview.");
        setRoutePreviews([]);
      } finally {
        if (active) {
          setRouteLoading(false);
        }
      }
    }

    loadRoutes();

    return () => {
      active = false;
    };
  }, [geocodedEvents, showRouteOverlay]);

  const selectedEvent = visibleEvents.find((event) => event.id === selectedEventId) ?? null;

  const adjustedRoutePreviews = React.useMemo<AdjustedRoutePreview[]>(() => {
    return routePreviews.map((route) => {
      const settings = dayRouteSettings[route.dayKey] ?? getDefaultDayRouteSetting();
      const firstEventId = route.orderedEventIds[0] ?? null;
      const lastEventId = route.orderedEventIds[route.orderedEventIds.length - 1] ?? null;

      const firstLeg =
        firstEventId != null ? route.legs.find((leg) => leg.toEventId === firstEventId) : undefined;
      const lastLeg = route.legs.find((leg) => leg.toEventId === HOME_OFFICE.id);

      const firstLegSeconds = firstLeg?.durationSeconds ?? 0;
      const firstLegMeters = firstLeg?.distanceMeters ?? 0;
      const lastLegSeconds = lastLeg?.durationSeconds ?? 0;
      const lastLegMeters = lastLeg?.distanceMeters ?? 0;

      const adjustedDriveSeconds = Math.max(
        0,
        route.totalDurationSeconds -
          (settings.startMode === "lastInspection" ? firstLegSeconds : 0) -
          (settings.endMode === "lastStop" ? lastLegSeconds : 0)
      );

      const adjustedDistanceMeters = Math.max(
        0,
        route.totalDistanceMeters -
          (settings.startMode === "lastInspection" ? firstLegMeters : 0) -
          (settings.endMode === "lastStop" ? lastLegMeters : 0)
      );

      return {
        ...route,
        startMode: settings.startMode,
        endMode: settings.endMode,
        adjustedDriveSeconds,
        adjustedDistanceMeters,
        adjustedFieldSeconds: route.totalEventDurationSeconds + adjustedDriveSeconds,
        firstLegSeconds,
        firstLegMeters,
        lastLegSeconds,
        lastLegMeters,
        firstEventId,
        lastEventId,
      };
    });
  }, [routePreviews, dayRouteSettings]);

  const selectedDayKey = React.useMemo(() => {
    if (selectedEvent?.startIso) {
      return getDayKeyFromIso(selectedEvent.startIso);
    }

    if (adjustedRoutePreviews.length > 0) {
      return adjustedRoutePreviews[0].dayKey;
    }

    return null;
  }, [selectedEvent, adjustedRoutePreviews]);

  const activeRoutePreview = React.useMemo(() => {
    if (!selectedDayKey) {
      return null;
    }

    return adjustedRoutePreviews.find((route) => route.dayKey === selectedDayKey) ?? null;
  }, [adjustedRoutePreviews, selectedDayKey]);

  const driveTimeByEventId = React.useMemo(() => {
    if (mapMode !== "routeDay") {
      return {};
    }

    const result: Record<string, number> = {};

    if (!activeRoutePreview) {
      return result;
    }

    activeRoutePreview.legs.forEach((leg) => {
      if (leg.toEventId === HOME_OFFICE.id) {
        return;
      }
      if (
        activeRoutePreview.startMode === "lastInspection" &&
        activeRoutePreview.firstEventId &&
        leg.toEventId === activeRoutePreview.firstEventId
      ) {
        result[leg.toEventId] = 0;
        return;
      }
      result[leg.toEventId] = leg.durationSeconds;
    });

    return result;
  }, [activeRoutePreview, mapMode]);

  const mapEvents = React.useMemo(() => {
    if (mapMode === "allFiltered") {
      return geocodedEvents;
    }

    return geocodedEvents.filter((event) => {
      const dayKey = getDayKeyFromIso(event.startIso);
      return dayKey === activeRoutePreview?.dayKey;
    });
  }, [geocodedEvents, mapMode, activeRoutePreview]);

  const technicianDaySummaries = React.useMemo(() => {
    const routeByDay = new Map<string, AdjustedRoutePreview>();
    adjustedRoutePreviews.forEach((route) => {
      routeByDay.set(route.dayKey, route);
    });

    const summaryMap = new Map<string, TechnicianDaySummary>();

    visibleEvents.forEach((event) => {
      if (!event.technicians || event.technicians.length === 0) {
        return;
      }

      const dayKey = getDayKeyFromIso(event.startIso);
      if (!dayKey) {
        return;
      }

      const route = routeByDay.get(dayKey);
      const incomingLeg = route?.legs.find((leg) => leg.toEventId === event.id);

      const isExcludedFirstLeg =
        route?.startMode === "lastInspection" &&
        route.firstEventId != null &&
        route.firstEventId === event.id;

      const driveSeconds = isExcludedFirstLeg ? 0 : incomingLeg?.durationSeconds ?? 0;
      const driveMeters = isExcludedFirstLeg ? 0 : incomingLeg?.distanceMeters ?? 0;
      const eventSeconds = getEventDurationSeconds(event.startIso, event.endIso);
      const dayLabel = route?.dayLabel ?? formatDateLong(event.startIso);

      event.technicians.forEach((technician) => {
        const key = `${technician}__${dayKey}`;
        const existing = summaryMap.get(key);

        if (existing) {
          existing.stops += 1;
          existing.eventSeconds += eventSeconds;
          existing.driveSeconds += driveSeconds;
          existing.fieldSeconds += eventSeconds + driveSeconds;
          existing.driveMeters += driveMeters;
          existing.calendars = uniqueSorted([...existing.calendars, event.calendarName]);
          existing.categories = uniqueSorted([...existing.categories, ...event.categories]);
        } else {
          summaryMap.set(key, {
            technician,
            dayKey,
            dayLabel,
            stops: 1,
            eventSeconds,
            driveSeconds,
            fieldSeconds: eventSeconds + driveSeconds,
            driveMeters,
            calendars: uniqueSorted([event.calendarName]),
            categories: uniqueSorted(event.categories),
          });
        }
      });
    });

    adjustedRoutePreviews.forEach((route) => {
      if (route.endMode !== "returnOffice" || !route.lastEventId) {
        return;
      }

      const lastEvent = visibleEvents.find((item) => item.id === route.lastEventId);
      if (!lastEvent || !lastEvent.technicians || lastEvent.technicians.length === 0) {
        return;
      }

      lastEvent.technicians.forEach((technician) => {
        const key = `${technician}__${route.dayKey}`;
        const existing = summaryMap.get(key);

        if (existing) {
          existing.driveSeconds += route.lastLegSeconds;
          existing.driveMeters += route.lastLegMeters;
          existing.fieldSeconds += route.lastLegSeconds;
        } else {
          summaryMap.set(key, {
            technician,
            dayKey: route.dayKey,
            dayLabel: route.dayLabel,
            stops: 0,
            eventSeconds: 0,
            driveSeconds: route.lastLegSeconds,
            fieldSeconds: route.lastLegSeconds,
            driveMeters: route.lastLegMeters,
            calendars: uniqueSorted([lastEvent.calendarName]),
            categories: uniqueSorted(lastEvent.categories),
          });
        }
      });
    });

    return Array.from(summaryMap.values()).sort((a, b) => {
      if (a.technician !== b.technician) {
        return a.technician.localeCompare(b.technician);
      }
      return a.dayKey.localeCompare(b.dayKey);
    });
  }, [visibleEvents, adjustedRoutePreviews]);

  const technicianDashboardStats = React.useMemo(() => {
    const techs = uniqueSorted(technicianDaySummaries.map((item) => item.technician));

    return {
      technicians: techs.length,
      dayRows: technicianDaySummaries.length,
      totalStops: technicianDaySummaries.reduce((sum, item) => sum + item.stops, 0),
      totalFieldSeconds: technicianDaySummaries.reduce((sum, item) => sum + item.fieldSeconds, 0),
    };
  }, [technicianDaySummaries]);

  const effectiveShowRouteOverlay = showRouteOverlay && mapMode === "routeDay";
  const cacheSize = getGeocodeCacheSize();

  function updateDayRouteSetting(dayKey: string, patch: Partial<DayRouteSetting>) {
    setDayRouteSettings((prev) => ({
      ...prev,
      [dayKey]: {
        ...(prev[dayKey] ?? getDefaultDayRouteSetting()),
        ...patch,
      },
    }));
  }

  function handlePopOut() {
    const url = `${window.location.origin}/taskpane.html?standalone=1`;
    window.open(url, "_blank", "width=1700,height=1100,resizable=yes,scrollbars=yes");
  }

  async function handleStandaloneSignIn() {
    try {
      setAuthWorking(true);
      await ensureGraphAccessInteractiveRedirect();
    } finally {
      setAuthWorking(false);
    }
  }

  function handleExportFilteredCsv() {
    const filename = `outlook-map-view_filtered_${getTodayStamp()}.csv`;

    const headers = [
      "Date",
      "Start Time",
      "End Time",
      "Event Duration",
      "Subject",
      "Calendar",
      "Technicians",
      "Categories",
      "Address",
      "Latitude",
      "Longitude",
      "Mappable",
      "Map Mode",
    ];

    const rows = visibleEvents.map((event) => {
      const geocoded = geocodedEvents.find((item) => item.id === event.id);

      return [
        formatDateForCsv(event.startIso),
        formatTimeForCsv(event.startIso),
        formatTimeForCsv(event.endIso),
        formatDuration(getEventDurationSeconds(event.startIso, event.endIso)),
        event.subject,
        event.calendarName,
        event.technicians.join(", "),
        event.categories.join(", "),
        event.addressText ?? "",
        geocoded?.latitude ?? "",
        geocoded?.longitude ?? "",
        !!event.addressText,
        mapMode === "routeDay" ? "Route day" : "All filtered events",
      ];
    });

    downloadCsv(filename, headers, rows);
  }

  function handleExportRouteCsv() {
    const filename = `outlook-map-view_routes_${getTodayStamp()}.csv`;

    const headers = [
      "Route Day",
      "Start Mode",
      "End Mode",
      "Stop Order",
      "Subject",
      "Date",
      "Start Time",
      "End Time",
      "Event Duration",
      "Drive Time From Previous Stop",
      "Drive Distance From Previous Stop",
      "Reported Day Drive Time",
      "Day Total Event Time",
      "Reported Day Field Time",
      "Reported Day Drive Distance",
      "Calendar",
      "Technicians",
      "Categories",
      "Address",
      "Latitude",
      "Longitude",
      "Map Mode",
    ];

    const rows: Array<Array<string | number | boolean | null | undefined>> = [];

    adjustedRoutePreviews.forEach((route) => {
      route.orderedEventIds.forEach((eventId, index) => {
        const event = visibleEvents.find((item) => item.id === eventId);
        if (!event) {
          return;
        }

        const geocoded = geocodedEvents.find((item) => item.id === eventId);
        const incomingLeg = route.legs.find((leg) => leg.toEventId === eventId);
        const adjustedIncomingSeconds =
          route.startMode === "lastInspection" && route.firstEventId === eventId
            ? 0
            : incomingLeg?.durationSeconds ?? 0;
        const adjustedIncomingMeters =
          route.startMode === "lastInspection" && route.firstEventId === eventId
            ? 0
            : incomingLeg?.distanceMeters ?? 0;

        rows.push([
          route.dayLabel,
          getStartModeLabel(route.startMode),
          getEndModeLabel(route.endMode),
          index + 1,
          event.subject,
          formatDateForCsv(event.startIso),
          formatTimeForCsv(event.startIso),
          formatTimeForCsv(event.endIso),
          formatDuration(getEventDurationSeconds(event.startIso, event.endIso)),
          formatDuration(adjustedIncomingSeconds),
          formatMiles(adjustedIncomingMeters),
          formatDuration(route.adjustedDriveSeconds),
          formatDuration(route.totalEventDurationSeconds),
          formatDuration(route.adjustedFieldSeconds),
          formatMiles(route.adjustedDistanceMeters),
          event.calendarName,
          event.technicians.join(", "),
          event.categories.join(", "),
          event.addressText ?? "",
          geocoded?.latitude ?? "",
          geocoded?.longitude ?? "",
          mapMode === "routeDay" ? "Route day" : "All filtered events",
        ]);
      });
    });

    downloadCsv(filename, headers, rows);
  }

  function handlePrintRouteSummary() {
    const reportTitle = `Outlook Map View Route Summary - ${getTodayStamp()}`;

    const selectedCalendars = availableCalendars
      .filter((calendar) => filters.calendarIds.includes(calendar.id))
      .map((calendar) => calendar.name)
      .join(", ");

    const daySections = adjustedRoutePreviews
      .map((route) => {
        const routeEvents = route.orderedEventIds
          .map((eventId) => geocodedEvents.find((item) => item.id === eventId))
          .filter((item): item is GeocodedCalendarEvent => !!item);

        const staticMapUrl = buildStaticMapUrl(routeEvents, route);

        const rows = route.orderedEventIds
          .map((eventId, index) => {
            const event = visibleEvents.find((item) => item.id === eventId);
            if (!event) {
              return "";
            }

            const incomingLeg = route.legs.find((leg) => leg.toEventId === eventId);
            const adjustedIncomingSeconds =
              route.startMode === "lastInspection" && route.firstEventId === eventId
                ? 0
                : incomingLeg?.durationSeconds ?? 0;
            const adjustedIncomingMeters =
              route.startMode === "lastInspection" && route.firstEventId === eventId
                ? 0
                : incomingLeg?.distanceMeters ?? 0;

            return `
              <tr>
                <td>${index + 1}</td>
                <td>${escapeHtml(event.subject)}</td>
                <td>${escapeHtml(formatDateLong(event.startIso))}</td>
                <td>${escapeHtml(formatTimeForCsv(event.startIso))}</td>
                <td>${escapeHtml(formatTimeForCsv(event.endIso))}</td>
                <td>${escapeHtml(formatDuration(getEventDurationSeconds(event.startIso, event.endIso)))}</td>
                <td>${escapeHtml(formatDuration(adjustedIncomingSeconds))}</td>
                <td>${escapeHtml(formatMiles(adjustedIncomingMeters))}</td>
                <td>${escapeHtml(event.calendarName)}</td>
                <td>${escapeHtml(event.technicians.join(", "))}</td>
                <td>${escapeHtml(event.categories.join(", "))}</td>
                <td>${escapeHtml(event.addressText ?? "")}</td>
              </tr>
            `;
          })
          .join("");

        return `
          <section class="day-section">
            <div class="day-header">
              <h2>${escapeHtml(route.dayLabel)}</h2>
              <div class="day-summary">
                <div><strong>Start Mode:</strong> ${escapeHtml(getStartModeLabel(route.startMode))}</div>
                <div><strong>End Mode:</strong> ${escapeHtml(getEndModeLabel(route.endMode))}</div>
                <div><strong>Stops:</strong> ${route.orderedEventIds.length}</div>
                <div><strong>Reported Drive Time:</strong> ${escapeHtml(formatDuration(route.adjustedDriveSeconds))}</div>
                <div><strong>Event Time:</strong> ${escapeHtml(formatDuration(route.totalEventDurationSeconds))}</div>
                <div><strong>Reported Field Time:</strong> ${escapeHtml(formatDuration(route.adjustedFieldSeconds))}</div>
                <div><strong>Reported Drive Distance:</strong> ${escapeHtml(formatMiles(route.adjustedDistanceMeters))}</div>
              </div>
            </div>

            ${
              staticMapUrl
                ? `
              <div class="snapshot-wrap">
                <img class="snapshot" src="${staticMapUrl}" alt="Route map snapshot for ${escapeHtml(
                    route.dayLabel
                  )}" />
                <div class="snapshot-caption">Map snapshot for ${escapeHtml(route.dayLabel)}</div>
              </div>
            `
                : ""
            }

            <table>
              <thead>
                <tr>
                  <th>Stop</th>
                  <th>Subject</th>
                  <th>Date</th>
                  <th>Start</th>
                  <th>End</th>
                  <th>Event Time</th>
                  <th>Drive In</th>
                  <th>Drive Miles</th>
                  <th>Calendar</th>
                  <th>Technicians</th>
                  <th>Categories</th>
                  <th>Address</th>
                </tr>
              </thead>
              <tbody>
                ${rows}
              </tbody>
            </table>
          </section>
        `;
      })
      .join("");

    const html = `
      <!doctype html>
      <html>
        <head>
          <meta charset="utf-8" />
          <title>${escapeHtml(reportTitle)}</title>
          <style>
            body {
              font-family: Arial, Helvetica, sans-serif;
              color: #0f172a;
              margin: 24px;
              background: #ffffff;
            }
            h1 {
              margin: 0 0 8px 0;
              font-size: 28px;
            }
            .meta {
              margin-bottom: 18px;
              color: #475569;
              font-size: 14px;
              line-height: 1.6;
            }
            .day-section {
              margin-top: 28px;
              page-break-inside: avoid;
              border: 1px solid #dbe2ea;
              border-radius: 14px;
              padding: 18px;
            }
            .day-header {
              margin-bottom: 12px;
            }
            .day-header h2 {
              margin: 0 0 10px 0;
              font-size: 20px;
            }
            .day-summary {
              display: grid;
              grid-template-columns: repeat(2, minmax(0, 1fr));
              gap: 6px 18px;
              font-size: 14px;
              margin-bottom: 10px;
            }
            .snapshot-wrap {
              margin: 16px 0 18px 0;
              page-break-inside: avoid;
            }
            .snapshot {
              width: 100%;
              max-width: 760px;
              max-height: 320px;
              object-fit: cover;
              border: 1px solid #dbe2ea;
              border-radius: 12px;
              display: block;
              margin: 0 auto;
              box-shadow: 0 1px 4px rgba(15, 23, 42, 0.08);
            }
            .snapshot-caption {
              margin-top: 8px;
              color: #64748b;
              font-size: 12px;
              text-align: center;
            }
            table {
              width: 100%;
              border-collapse: collapse;
              font-size: 12px;
            }
            th, td {
              border: 1px solid #dbe2ea;
              padding: 8px;
              vertical-align: top;
              text-align: left;
            }
            th {
              background: #f8fafc;
              font-weight: 700;
            }
            tbody tr:nth-child(even) {
              background: #fcfdff;
            }
          </style>
        </head>
        <body>
          <h1>Route Summary</h1>
          <div class="meta">
            <div><strong>Generated:</strong> ${escapeHtml(new Date().toLocaleString())}</div>
            <div><strong>Calendars:</strong> ${escapeHtml(selectedCalendars || "Default selection")}</div>
            <div><strong>Technicians:</strong> ${escapeHtml(filters.technicians.join(", ") || "All")}</div>
            <div><strong>Categories:</strong> ${escapeHtml(filters.categories.join(", ") || "All")}</div>
            <div><strong>Map Mode:</strong> ${escapeHtml(mapMode === "routeDay" ? "Route day" : "All filtered events")}</div>
          </div>
          ${daySections || "<p>No route data available to print.</p>"}
        </body>
      </html>
    `;

    const printWindow = window.open("", "_blank", "width=1280,height=960,scrollbars=yes");
    if (!printWindow) {
      return;
    }

    printWindow.document.open();
    printWindow.document.write(html);
    printWindow.document.close();
    printWindow.focus();

    setTimeout(() => {
      printWindow.print();
    }, 500);
  }

  const stickyTop = standalone ? 20 : 16;

  return (
    <div
      style={{
        padding: standalone ? 20 : 16,
        fontFamily: "Arial, Helvetica, sans-serif",
        background: theme.appBg,
        minHeight: "100vh",
        color: theme.text,
        ["--omv-border" as any]: theme.border,
        ["--omv-border-soft" as any]: theme.borderSoft,
        ["--omv-panel-soft" as any]: theme.panelSoftBg,
        ["--omv-text" as any]: theme.text,
        ["--omv-text-soft" as any]: theme.textSoft,
      }}
    >
      <div
        style={{
          marginBottom: 16,
          background: theme.shellBg,
          border: `1px solid ${theme.shellBorder}`,
          borderRadius: 18,
          padding: 18,
          boxShadow: theme.shadow,
        }}
      >
        <div
          style={{
            display: "flex",
            alignItems: "flex-start",
            justifyContent: "space-between",
            gap: 16,
            flexWrap: "wrap",
            marginBottom: 14,
          }}
        >
          <div>
            <h1
              style={{
                margin: "0 0 6px 0",
                fontSize: standalone ? 32 : 28,
                lineHeight: 1.1,
                color: theme.text,
              }}
            >
              Outlook Map View
            </h1>
            <p style={{ margin: 0, color: theme.textSoft, fontSize: 15 }}>
              Visualize calendar events by date, category, technician, calendar, and location.
            </p>
          </div>

          <div
            style={{
              display: "flex",
              gap: 8,
              flexWrap: "wrap",
              justifyContent: "flex-end",
              alignItems: "flex-start",
            }}
          >
            {!standalone ? (
              <button
                type="button"
                onClick={handlePopOut}
                style={{
                  height: 40,
                  padding: "0 14px",
                  borderRadius: 10,
                  border: `1px solid ${theme.borderSoft}`,
                  background: theme.panelBg,
                  color: theme.text,
                  fontSize: 14,
                  fontWeight: 700,
                  cursor: "pointer",
                }}
              >
                Pop out view
              </button>
            ) : null}

            <button
              type="button"
              onClick={() => setThemeMode((prev) => (prev === "dark" ? "light" : "dark"))}
              style={{
                height: 40,
                padding: "0 14px",
                borderRadius: 10,
                border: `1px solid ${theme.borderSoft}`,
                background: theme.panelBg,
                color: theme.text,
                fontSize: 14,
                fontWeight: 700,
                cursor: "pointer",
              }}
            >
              {themeMode === "dark" ? "Light mode" : "Dark mode"}
            </button>

            <StatChip label="Visible" value={visibleEvents.length} themeMode={themeMode} />
            <StatChip label="Mapped" value={geocodedEvents.length} themeMode={themeMode} />
            <StatChip label="Cached" value={cacheSize} themeMode={themeMode} />
          </div>
        </div>

        {standalone && calendarsError ? (
          <div
            style={{
              marginBottom: 14,
              padding: 12,
              borderRadius: 12,
              border: "1px solid #fecaca",
              background: "#fef2f2",
              color: "#991b1b",
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              gap: 12,
              flexWrap: "wrap",
            }}
          >
            <span>Standalone view needs Microsoft 365 sign-in to load calendars.</span>
            <button
              type="button"
              onClick={handleStandaloneSignIn}
              disabled={authWorking}
              style={{
                height: 38,
                padding: "0 14px",
                borderRadius: 10,
                border: "1px solid #2563eb",
                background: "#2563eb",
                color: "#ffffff",
                fontSize: 14,
                fontWeight: 700,
                cursor: authWorking ? "default" : "pointer",
                opacity: authWorking ? 0.7 : 1,
              }}
            >
              {authWorking ? "Signing in..." : "Sign in to Microsoft 365"}
            </button>
          </div>
        ) : null}

        <FilterBar
          themeMode={themeMode}
          filters={filters}
          availableCategories={availableCategories}
          availableTechnicians={availableTechnicians}
          availableCalendars={availableCalendars}
          onPresetChange={(preset) => setFilters((prev) => ({ ...prev, preset }))}
          onCustomDateChange={(field, value) =>
            setFilters((prev) => ({ ...prev, [field]: value }))
          }
          onCategoriesChange={(categories) =>
            setFilters((prev) => ({ ...prev, categories }))
          }
          onTechniciansChange={(technicians) =>
            setFilters((prev) => ({ ...prev, technicians }))
          }
          onCalendarIdsChange={(calendarIds) =>
            setFilters((prev) => ({ ...prev, calendarIds }))
          }
          onRefresh={() => setRefreshKey((prev) => prev + 1)}
        />
      </div>

      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 16,
          marginBottom: 16,
          padding: 12,
          background: theme.controlBg,
          border: `1px solid ${theme.border}`,
          borderRadius: 12,
          flexWrap: "wrap",
        }}
      >
        <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 14, cursor: "pointer", color: theme.textSoft }}>
          <input
            id="showOnlyMappable"
            type="checkbox"
            checked={showOnlyMappable}
            onChange={(e) => setShowOnlyMappable(e.target.checked)}
          />
          Show only events with usable map addresses
        </label>

        <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 14, cursor: "pointer", color: theme.textSoft }}>
          <input
            id="showRouteOverlay"
            type="checkbox"
            checked={showRouteOverlay}
            onChange={(e) => setShowRouteOverlay(e.target.checked)}
            disabled={mapMode === "allFiltered"}
          />
          Show route overlay
        </label>

        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <span
            style={{
              fontSize: 12,
              fontWeight: 700,
              color: theme.textMuted,
              textTransform: "uppercase",
              letterSpacing: "0.04em",
            }}
          >
            Map Mode
          </span>

          <div
            style={{
              display: "inline-flex",
              border: `1px solid ${theme.borderSoft}`,
              borderRadius: 999,
              overflow: "hidden",
              background: theme.panelBg,
            }}
          >
            <button
              type="button"
              onClick={() => setMapMode("routeDay")}
              style={{
                border: "none",
                padding: "8px 12px",
                background: mapMode === "routeDay" ? (themeMode === "dark" ? "rgba(37,99,235,0.18)" : "#dbeafe") : theme.panelBg,
                color: mapMode === "routeDay" ? "#2563eb" : theme.textSoft,
                fontSize: 13,
                fontWeight: 700,
                cursor: "pointer",
              }}
            >
              Route day
            </button>
            <button
              type="button"
              onClick={() => setMapMode("allFiltered")}
              style={{
                border: "none",
                borderLeft: "1px solid #cbd5e1",
                padding: "8px 12px",
                background: mapMode === "allFiltered" ? (themeMode === "dark" ? "rgba(37,99,235,0.18)" : "#dbeafe") : theme.panelBg,
                color: mapMode === "allFiltered" ? "#2563eb" : theme.textSoft,
                fontSize: 13,
                fontWeight: 700,
                cursor: "pointer",
              }}
            >
              All filtered events
            </button>
          </div>
        </div>

        <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
          <span
            style={{
              fontSize: 12,
              fontWeight: 700,
              color: theme.textMuted,
              textTransform: "uppercase",
              letterSpacing: "0.04em",
            }}
          >
            Export
          </span>

          <button
            type="button"
            onClick={handleExportFilteredCsv}
            style={{
              height: 36,
              padding: "0 12px",
              borderRadius: 10,
              border: `1px solid ${theme.borderSoft}`,
              background: theme.panelBg,
              color: theme.text,
              fontSize: 13,
              fontWeight: 700,
              cursor: "pointer",
            }}
          >
            Export filtered CSV
          </button>

          <button
            type="button"
            onClick={handleExportRouteCsv}
            disabled={adjustedRoutePreviews.length === 0}
            style={{
              height: 36,
              padding: "0 12px",
              borderRadius: 10,
              border: `1px solid ${theme.borderSoft}`,
              background: adjustedRoutePreviews.length === 0 ? theme.panelSoftBg : theme.panelBg,
              color: adjustedRoutePreviews.length === 0 ? theme.textMuted : theme.text,
              fontSize: 13,
              fontWeight: 700,
              cursor: adjustedRoutePreviews.length === 0 ? "default" : "pointer",
            }}
          >
            Export route CSV
          </button>

          <button
            type="button"
            onClick={handlePrintRouteSummary}
            disabled={adjustedRoutePreviews.length === 0}
            style={{
              height: 36,
              padding: "0 12px",
              borderRadius: 10,
              border: `1px solid ${theme.borderSoft}`,
              background: adjustedRoutePreviews.length === 0 ? theme.panelSoftBg : theme.panelBg,
              color: adjustedRoutePreviews.length === 0 ? theme.textMuted : theme.text,
              fontSize: 13,
              fontWeight: 700,
              cursor: adjustedRoutePreviews.length === 0 ? "default" : "pointer",
            }}
          >
            Print route summary
          </button>
        </div>

        <span style={{ fontSize: 12, color: theme.textMuted }}>
          Day-level route assumptions now drive reporting totals.
        </span>

        <span style={{ marginLeft: "auto", fontSize: 12, color: theme.textMuted }}>
          Showing {visibleEvents.length} event{visibleEvents.length === 1 ? "" : "s"} •{" "}
          {geocodedEvents.length} mapped
        </span>
      </div>

      <div
        style={{
          marginBottom: 16,
          border: `1px solid ${theme.border}`,
          padding: 16,
          borderRadius: 12,
          background: theme.panelBg,
        }}
      >
        <div
          style={{
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            gap: 16,
            flexWrap: "wrap",
            marginBottom: showTechnicianDashboard ? 14 : 0,
          }}
        >
          <div>
            <h2 style={{ margin: "0 0 4px 0", fontSize: 20, color: theme.text }}>Technician Day Totals</h2>
            <p style={{ margin: 0, color: theme.textMuted, fontSize: 14 }}>
              Totals reflect current filters and per-day start/end assumptions.
            </p>
          </div>

          <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
            <MiniStatChip label="Techs" value={technicianDashboardStats.technicians} themeMode={themeMode} />
            <MiniStatChip label="Day Rows" value={technicianDashboardStats.dayRows} themeMode={themeMode} />
            <MiniStatChip label="Stops" value={technicianDashboardStats.totalStops} themeMode={themeMode} />
            <MiniStatChip
              label="Field Time"
              value={formatDuration(technicianDashboardStats.totalFieldSeconds)}
              themeMode={themeMode}
            />

            <button
              type="button"
              onClick={() => setShowTechnicianDashboard((prev) => !prev)}
              style={{
                height: 38,
                padding: "0 14px",
                borderRadius: 10,
                border: `1px solid ${theme.borderSoft}`,
                background: theme.panelBg,
                color: theme.text,
                fontSize: 14,
                fontWeight: 700,
                cursor: "pointer",
              }}
            >
              {showTechnicianDashboard ? "Hide dashboard" : "Show dashboard"}
            </button>
          </div>
        </div>

        {showTechnicianDashboard ? (
          technicianDaySummaries.length === 0 ? (
            <div
              style={{
                border: "1px dashed #cbd5e1",
                borderRadius: 10,
                padding: 14,
                color: theme.textMuted,
                background: theme.appBg,
              }}
            >
              No technician totals available for the current filters.
            </div>
          ) : (
            <>
              <div
                style={{
                  display: "grid",
                  gridTemplateColumns: "repeat(auto-fit, minmax(260px, 1fr))",
                  gap: 12,
                  marginBottom: 16,
                }}
              >
                {uniqueSorted(technicianDaySummaries.map((item) => item.technician)).map((technician) => {
                  const rows = technicianDaySummaries.filter((item) => item.technician === technician);
                  const stops = rows.reduce((sum, item) => sum + item.stops, 0);
                  const driveSeconds = rows.reduce((sum, item) => sum + item.driveSeconds, 0);
                  const fieldSeconds = rows.reduce((sum, item) => sum + item.fieldSeconds, 0);
                  const miles = rows.reduce((sum, item) => sum + item.driveMeters, 0);

                  return (
                    <div
                      key={technician}
                      style={{
                        border: `1px solid ${theme.shellBorder}`,
                        borderRadius: 12,
                        padding: 14,
                        background: theme.panelSoftBg,
                      }}
                    >
                      <div style={{ fontWeight: 700, color: theme.text, marginBottom: 8 }}>{technician}</div>
                      <div style={{ display: "grid", gap: 6, fontSize: 14, color: theme.textSoft }}>
                        <div>
                          <strong style={{ color: theme.text }}>Days:</strong> {rows.length}
                        </div>
                        <div>
                          <strong style={{ color: theme.text }}>Stops:</strong> {stops}
                        </div>
                        <div>
                          <strong style={{ color: theme.text }}>Drive:</strong> {formatDuration(driveSeconds)}
                        </div>
                        <div>
                          <strong style={{ color: theme.text }}>Field:</strong> {formatDuration(fieldSeconds)}
                        </div>
                        <div>
                          <strong style={{ color: theme.text }}>Miles:</strong> {formatMiles(miles)}
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>

              <div
                style={{
                  marginBottom: 12,
                  padding: 10,
                  borderRadius: 10,
                  background: theme.appBg,
                  border: `1px solid ${theme.shellBorder}`,
                  color: theme.textMuted,
                  fontSize: 12,
                }}
              >
                “Start from last inspection” currently removes the office departure leg for that day.
                It does not yet calculate the actual cross-day travel from the prior day’s last stop.
              </div>

              <div style={{ overflowX: "auto" }}>
                <table
                  style={{
                    width: "100%",
                    borderCollapse: "collapse",
                    fontSize: 13,
                  }}
                >
                  <thead>
                    <tr>
                      <th style={tableHeaderStyle}>Technician</th>
                      <th style={tableHeaderStyle}>Day</th>
                      <th style={tableHeaderStyle}>Stops</th>
                      <th style={tableHeaderStyle}>Event Time</th>
                      <th style={tableHeaderStyle}>Drive Time</th>
                      <th style={tableHeaderStyle}>Field Time</th>
                      <th style={tableHeaderStyle}>Drive Miles</th>
                      <th style={tableHeaderStyle}>Calendars</th>
                      <th style={tableHeaderStyle}>Categories</th>
                    </tr>
                  </thead>
                  <tbody>
                    {technicianDaySummaries.map((item) => (
                      <tr key={`${item.technician}-${item.dayKey}`}>
                        <td style={tableCellStyle}>{item.technician}</td>
                        <td style={tableCellStyle}>{item.dayLabel}</td>
                        <td style={tableCellStyle}>{item.stops}</td>
                        <td style={tableCellStyle}>{formatDuration(item.eventSeconds)}</td>
                        <td style={tableCellStyle}>{formatDuration(item.driveSeconds)}</td>
                        <td style={tableCellStyle}>{formatDuration(item.fieldSeconds)}</td>
                        <td style={tableCellStyle}>{formatMiles(item.driveMeters)}</td>
                        <td style={tableCellStyle}>{item.calendars.join(", ")}</td>
                        <td style={tableCellStyle}>{item.categories.join(", ")}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </>
          )
        ) : null}
      </div>

      {loading ? <p style={{ color: theme.textSoft }}>Loading calendar events...</p> : null}
      {error ? <p style={{ color: "#ef4444" }}>{error}</p> : null}
      {mapLoading ? <p style={{ color: theme.textSoft }}>Geocoding mappable events...</p> : null}
      {mapError ? <p style={{ color: "#ef4444" }}>{mapError}</p> : null}
      {routeLoading ? <p style={{ color: theme.textSoft }}>Building route preview...</p> : null}
      {routeError ? <p style={{ color: "#ef4444" }}>{routeError}</p> : null}

      <div
        style={{
          display: "grid",
          gridTemplateColumns: standalone
            ? "minmax(340px, 420px) minmax(0, 1fr)"
            : "minmax(300px, 360px) minmax(0, 1fr)",
          gap: 16,
          alignItems: "start",
        }}
      >
        <div
          style={{
            display: "flex",
            flexDirection: "column",
            gap: 16,
            position: "sticky",
            top: stickyTop,
            alignSelf: "start",
          }}
        >
          <div
            style={{
              maxHeight: "42vh",
              overflowY: "auto",
              paddingRight: 4,
            }}
          >
            <EventList
              events={visibleEvents}
              selectedEventId={selectedEventId}
              onSelectEvent={setSelectedEventId}
              driveTimeByEventId={driveTimeByEventId}
              themeMode={themeMode}
            />
          </div>

          <div
            style={{
              border: `1px solid ${theme.border}`,
              padding: 16,
              borderRadius: 12,
              background: theme.panelBg,
            }}
          >
            <h2 style={{ marginTop: 0, fontSize: 18, color: theme.text }}>Route Preview</h2>

            <div
              style={{
                maxHeight: "28vh",
                overflowY: "auto",
                paddingRight: 4,
              }}
            >
              {mapMode === "allFiltered" ? (
                <p style={{ margin: 0, color: theme.textMuted }}>
                  Route cards stay available below, but the map is currently showing all filtered mappable events.
                </p>
              ) : effectiveShowRouteOverlay && adjustedRoutePreviews.length > 0 ? (
                <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
                  {adjustedRoutePreviews.map((route) => {
                    const isActive = route.dayKey === activeRoutePreview?.dayKey;

                    return (
                      <div
                        key={route.dayKey}
                        style={{
                          border: isActive ? `2px solid ${theme.panelSelectedBorder}` : `1px solid ${theme.border}`,
                          background: isActive ? theme.panelSelectedBg : theme.panelBg,
                          borderRadius: 12,
                          padding: 12,
                        }}
                      >
                        <div
                          style={{
                            display: "flex",
                            justifyContent: "space-between",
                            alignItems: "center",
                            gap: 12,
                            flexWrap: "wrap",
                            marginBottom: 8,
                          }}
                        >
                          <div style={{ fontWeight: 700, color: theme.text }}>{route.dayLabel}</div>
                          <div style={{ fontSize: 12, color: theme.textMuted }}>
                            Reporting uses the settings below
                          </div>
                        </div>

                        <div
                          style={{
                            display: "grid",
                            gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))",
                            gap: 10,
                            marginBottom: 12,
                          }}
                        >
                          <label style={{ display: "grid", gap: 4, fontSize: 13, color: theme.textSoft }}>
                            <span style={{ fontWeight: 700, color: theme.text }}>Start</span>
                            <select
                              value={route.startMode}
                              onChange={(e) =>
                                updateDayRouteSetting(route.dayKey, {
                                  startMode: e.target.value as StartMode,
                                })
                              }
                              style={{
                                height: 36,
                                borderRadius: 8,
                                border: `1px solid ${theme.borderSoft}`,
                                padding: "0 10px",
                                background: theme.panelBg,
                                color: theme.text,
                              }}
                            >
                              <option value="homeOffice">Home office</option>
                              <option value="lastInspection">Start from last inspection</option>
                            </select>
                          </label>

                          <label style={{ display: "grid", gap: 4, fontSize: 13, color: theme.textSoft }}>
                            <span style={{ fontWeight: 700, color: theme.text }}>End</span>
                            <select
                              value={route.endMode}
                              onChange={(e) =>
                                updateDayRouteSetting(route.dayKey, {
                                  endMode: e.target.value as EndMode,
                                })
                              }
                              style={{
                                height: 36,
                                borderRadius: 8,
                                border: `1px solid ${theme.borderSoft}`,
                                padding: "0 10px",
                                background: theme.panelBg,
                                color: theme.text,
                              }}
                            >
                              <option value="lastStop">End at last inspection</option>
                              <option value="returnOffice">Return home/office</option>
                            </select>
                          </label>
                        </div>

                        <div style={{ color: theme.textSoft, fontSize: 14, display: "grid", gap: 6 }}>
                          <div>
                            Start mode:{" "}
                            <strong style={{ color: theme.text }}>
                              {getStartModeLabel(route.startMode)}
                            </strong>
                          </div>
                          <div>
                            End mode:{" "}
                            <strong style={{ color: theme.text }}>
                              {getEndModeLabel(route.endMode)}
                            </strong>
                          </div>
                          <div>
                            Stops:{" "}
                            <strong style={{ color: theme.text }}>{route.orderedEventIds.length}</strong>
                          </div>
                          <div>
                            Reported drive time:{" "}
                            <strong style={{ color: theme.text }}>
                              {formatDuration(route.adjustedDriveSeconds)}
                            </strong>
                          </div>
                          <div>
                            Event time:{" "}
                            <strong style={{ color: theme.text }}>
                              {formatDuration(route.totalEventDurationSeconds)}
                            </strong>
                          </div>
                          <div>
                            Reported field time:{" "}
                            <strong style={{ color: theme.text }}>
                              {formatDuration(route.adjustedFieldSeconds)}
                            </strong>
                          </div>
                          <div>
                            Reported drive distance:{" "}
                            <strong style={{ color: theme.text }}>
                              {formatMiles(route.adjustedDistanceMeters)}
                            </strong>
                          </div>
                        </div>

                        {route.startMode === "lastInspection" ? (
                          <div
                            style={{
                              marginTop: 10,
                              fontSize: 12,
                              color: theme.textMuted,
                              background: theme.panelSoftBg,
                              border: `1px solid ${theme.shellBorder}`,
                              borderRadius: 8,
                              padding: 8,
                            }}
                          >
                            Start-from-last-inspection currently excludes the office departure leg. Cross-day reposition travel is not yet added.
                          </div>
                        ) : null}
                      </div>
                    );
                  })}
                </div>
              ) : (
                <p style={{ margin: 0, color: theme.textMuted }}>
                  {effectiveShowRouteOverlay
                    ? "Need at least one mapped event to preview a route from Home Office."
                    : "Route overlay is turned off."}
                </p>
              )}
            </div>
          </div>

          <div
            style={{
              border: `1px solid ${theme.border}`,
              padding: 16,
              borderRadius: 12,
              background: theme.panelBg,
            }}
          >
            <h2 style={{ marginTop: 0, fontSize: 18 }}>Selected Event</h2>

            <div
              style={{
                maxHeight: "18vh",
                overflowY: "auto",
                paddingRight: 4,
              }}
            >
              {selectedEvent ? (
                <div>
                  <h3 style={{ marginTop: 0, marginBottom: 8, color: theme.text }}>{selectedEvent.subject}</h3>
                  <p style={{ margin: "0 0 8px 0", color: theme.textSoft }}>
                    {selectedEvent.addressText || "No usable address"}
                  </p>
                  <p style={{ margin: "0 0 8px 0", color: theme.textMuted }}>
                    Calendar: {selectedEvent.calendarName}
                  </p>
                  <p style={{ margin: "0 0 8px 0", color: theme.textMuted }}>
                    Technicians:{" "}
                    {selectedEvent.technicians.length > 0
                      ? selectedEvent.technicians.join(", ")
                      : "None"}
                  </p>
                  <p style={{ margin: "0 0 12px 0", color: theme.textMuted }}>
                    Categories:{" "}
                    {selectedEvent.categories.length > 0
                      ? selectedEvent.categories.join(", ")
                      : "None"}
                  </p>

                  {selectedEvent.webLink ? (
                    <a
                      href={selectedEvent.webLink}
                      target="_blank"
                      rel="noreferrer"
                      style={{
                        color: "#2563eb",
                        fontWeight: 600,
                        textDecoration: "none",
                      }}
                    >
                      Open in Outlook
                    </a>
                  ) : null}
                </div>
              ) : (
                <p style={{ margin: 0, color: theme.textMuted }}>No event selected.</p>
              )}
            </div>
          </div>
        </div>

        <div>
          <EventMap
            events={mapEvents}
            selectedEventId={selectedEventId}
            onSelectEvent={setSelectedEventId}
            routeGeometry={activeRoutePreview?.geometry ?? null}
            showRouteOverlay={effectiveShowRouteOverlay}
            standalone={standalone}
          />
        </div>
      </div>
    </div>
  );
}

function StatChip({ label, value, themeMode }: { label: string; value: string | number; themeMode: ThemeMode }) {
  const theme = getTheme(themeMode);
  return (
    <div
      style={{
        background: theme.chipBg,
        border: `1px solid ${theme.chipBorder}`,
        borderRadius: 12,
        padding: "8px 10px",
        minWidth: 78,
      }}
    >
      <div
        style={{
          fontSize: 11,
          fontWeight: 700,
          color: theme.textMuted,
          textTransform: "uppercase",
          letterSpacing: "0.04em",
          marginBottom: 2,
        }}
      >
        {label}
      </div>
      <div style={{ fontSize: 18, fontWeight: 700, color: theme.text }}>{value}</div>
    </div>
  );
}

function MiniStatChip({ label, value, themeMode }: { label: string; value: string | number; themeMode: ThemeMode }) {
  const theme = getTheme(themeMode);
  return (
    <div
      style={{
        background: theme.panelSoftBg,
        border: `1px solid ${theme.chipBorder}`,
        borderRadius: 10,
        padding: "8px 10px",
        minWidth: 80,
      }}
    >
      <div
        style={{
          fontSize: 10,
          fontWeight: 700,
          color: theme.textMuted,
          textTransform: "uppercase",
          letterSpacing: "0.04em",
          marginBottom: 2,
        }}
      >
        {label}
      </div>
      <div style={{ fontSize: 15, fontWeight: 700, color: theme.text }}>{value}</div>
    </div>
  );
}

const tableHeaderStyle: React.CSSProperties = {
  border: "1px solid var(--omv-border-soft)",
  background: "var(--omv-panel-soft)",
  padding: "8px",
  textAlign: "left",
  fontWeight: 700,
  color: "var(--omv-text)",
};

const tableCellStyle: React.CSSProperties = {
  border: "1px solid var(--omv-border-soft)",
  padding: "8px",
  verticalAlign: "top",
  color: "var(--omv-text-soft)",
};