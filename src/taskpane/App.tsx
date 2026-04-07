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

type ThemeMode = "light" | "dark";

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

function getTheme(mode: ThemeMode) {
  if (mode === "dark") {
    return {
      appBg: "#020617",
      shellBg: "linear-gradient(135deg, #0f172a 0%, #111827 100%)",
      shellBorder: "#334155",
      panelBg: "#111827",
      panelAltBg: "#0f172a",
      panelSelectedBg: "rgba(37,99,235,0.16)",
      panelSelectedBorder: "#60a5fa",
      border: "#334155",
      borderSoft: "#475569",
      text: "#f8fafc",
      mutedText: "#94a3b8",
      subtleText: "#cbd5e1",
      accent: "#60a5fa",
      accentBg: "rgba(37,99,235,0.18)",
      danger: "#fca5a5",
      statusBg: "#0f172a",
      statusBorder: "#334155",
      controlBg: "#111827",
      chipBg: "#0f172a",
      chipBorder: "#334155",
      shadow: "0 18px 38px rgba(2, 6, 23, 0.35)",
    };
  }

  return {
    appBg: "#f8fafc",
    shellBg: "linear-gradient(135deg, #ffffff 0%, #f8fbff 100%)",
    shellBorder: "#dbe2ea",
    panelBg: "#ffffff",
    panelAltBg: "#f8fafc",
    panelSelectedBg: "#eff6ff",
    panelSelectedBorder: "#2563eb",
    border: "#d1d5db",
    borderSoft: "#cbd5e1",
    text: "#0f172a",
    mutedText: "#64748b",
    subtleText: "#475569",
    accent: "#2563eb",
    accentBg: "#dbeafe",
    danger: "crimson",
    statusBg: "#ffffff",
    statusBorder: "#d1d5db",
    controlBg: "#ffffff",
    chipBg: "#ffffff",
    chipBorder: "#dbe2ea",
    shadow: "0 2px 8px rgba(15,23,42,0.04)",
  };
}

export default function App() {
  const [themeMode, setThemeMode] = React.useState<ThemeMode>("light");
  const [filters, setFilters] = React.useState<FilterState>(initialFilters);
  const [selectedEventId, setSelectedEventId] = React.useState<string | null>(null);
  const [refreshKey, setRefreshKey] = React.useState<number>(0);
  const [geocodedEvents, setGeocodedEvents] = React.useState<GeocodedCalendarEvent[]>([]);
  const [mapLoading, setMapLoading] = React.useState<boolean>(false);
  const [mapError, setMapError] = React.useState<string | null>(null);
  const [showOnlyMappable, setShowOnlyMappable] = React.useState<boolean>(false);
  const [showRouteOverlay, setShowRouteOverlay] = React.useState<boolean>(true);
  const [returnToOffice, setReturnToOffice] = React.useState<boolean>(false);
  const [routeLoading, setRouteLoading] = React.useState<boolean>(false);
  const [routeError, setRouteError] = React.useState<string | null>(null);
  const [routePreviews, setRoutePreviews] = React.useState<RoutePreview[]>([]);

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
    loading,
    error,
  } = useCalendarEvents(filters, refreshKey);

  React.useEffect(() => {
    if (availableCalendars.length === 0 || (filters.calendarIds ?? []).length > 0) {
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
  }, [availableCalendars, filters.calendarIds]);

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
          returnToOffice ? HOME_OFFICE : null
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
  }, [geocodedEvents, showRouteOverlay, returnToOffice]);

  const selectedEvent = visibleEvents.find((event) => event.id === selectedEventId) ?? null;

  const selectedDayKey = React.useMemo(() => {
    if (selectedEvent?.startIso) {
      return getDayKeyFromIso(selectedEvent.startIso);
    }
    if (routePreviews.length > 0) {
      return routePreviews[0].dayKey;
    }
    return null;
  }, [selectedEvent, routePreviews]);

  const activeRoutePreview = React.useMemo(() => {
    if (!selectedDayKey) {
      return null;
    }
    return routePreviews.find((route) => route.dayKey === selectedDayKey) ?? null;
  }, [routePreviews, selectedDayKey]);

  const activeDayEvents = React.useMemo(() => {
    if (!selectedDayKey) {
      return visibleEvents;
    }

    return visibleEvents.filter((event) => getDayKeyFromIso(event.startIso) === selectedDayKey);
  }, [visibleEvents, selectedDayKey]);

  const driveTimeByEventId = React.useMemo(() => {
    const result: Record<string, number> = {};

    if (!activeRoutePreview) {
      return result;
    }

    activeRoutePreview.legs.forEach((leg) => {
      if (leg.toEventId !== HOME_OFFICE.id) {
        result[leg.toEventId] = leg.durationSeconds;
      }
    });

    return result;
  }, [activeRoutePreview]);

  const cacheSize = getGeocodeCacheSize();

  return (
    <div
      style={{
        padding: 16,
        fontFamily: "Arial, Helvetica, sans-serif",
        background: theme.appBg,
        minHeight: "100vh",
        color: theme.text,
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
                fontSize: 28,
                lineHeight: 1.1,
                color: theme.text,
              }}
            >
              Outlook Map View
            </h1>
            <p style={{ margin: 0, color: theme.mutedText, fontSize: 15 }}>
              Visualize calendar events by date, category, technician, and location.
            </p>
          </div>

          <div
            style={{
              display: "flex",
              gap: 8,
              flexWrap: "wrap",
              justifyContent: "flex-end",
              alignItems: "center",
            }}
          >
            <button
              type="button"
              onClick={() => setThemeMode((prev) => (prev === "dark" ? "light" : "dark"))}
              style={{
                border: `1px solid ${theme.borderSoft}`,
                background: theme.panelBg,
                color: theme.text,
                borderRadius: 12,
                padding: "8px 12px",
                fontSize: 13,
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

        <FilterBar
          filters={filters}
          availableCategories={availableCategories}
          availableTechnicians={availableTechnicians}
          availableCalendars={availableCalendars}
          onPresetChange={(preset) =>
            setFilters((prev) => ({
              ...prev,
              preset,
            }))
          }
          onCustomDateChange={(field, value) =>
            setFilters((prev) => ({
              ...prev,
              [field]: value,
            }))
          }
          onCategoriesChange={(categories) =>
            setFilters((prev) => ({
              ...prev,
              categories,
            }))
          }
          onTechniciansChange={(technicians) =>
            setFilters((prev) => ({
              ...prev,
              technicians,
            }))
          }
          onCalendarIdsChange={(calendarIds) =>
            setFilters((prev) => ({
              ...prev,
              calendarIds,
            }))
          }
          onRefresh={() => setRefreshKey((prev) => prev + 1)}
          themeMode={themeMode}
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
          boxShadow: themeMode === "dark" ? "0 8px 24px rgba(2, 6, 23, 0.18)" : "none",
        }}
      >
        <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 14, cursor: "pointer", color: theme.subtleText }}>
          <input
            id="showOnlyMappable"
            type="checkbox"
            checked={showOnlyMappable}
            onChange={(e) => setShowOnlyMappable(e.target.checked)}
          />
          Show only events with usable map addresses
        </label>

        <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 14, cursor: "pointer", color: theme.subtleText }}>
          <input
            id="showRouteOverlay"
            type="checkbox"
            checked={showRouteOverlay}
            onChange={(e) => setShowRouteOverlay(e.target.checked)}
          />
          Show route overlay
        </label>

        <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 14, cursor: "pointer", color: theme.subtleText }}>
          <input
            id="returnToOffice"
            type="checkbox"
            checked={returnToOffice}
            onChange={(e) => setReturnToOffice(e.target.checked)}
            disabled={!showRouteOverlay}
          />
          Return to office
        </label>

        <span style={{ fontSize: 12, color: theme.mutedText }}>
          Route starts at <strong style={{ color: theme.text }}>Home Office</strong>
        </span>

        <span style={{ marginLeft: "auto", fontSize: 12, color: theme.mutedText }}>
          Showing {visibleEvents.length} event{visibleEvents.length === 1 ? "" : "s"} • {geocodedEvents.length} mapped
        </span>
      </div>

      {loading ? <StatusMessage text="Loading calendar events..." themeMode={themeMode} /> : null}
      {error ? <StatusMessage text={error} themeMode={themeMode} danger /> : null}
      {mapLoading ? <StatusMessage text="Geocoding mappable events..." themeMode={themeMode} /> : null}
      {mapError ? <StatusMessage text={mapError} themeMode={themeMode} danger /> : null}
      {routeLoading ? <StatusMessage text="Building route preview..." themeMode={themeMode} /> : null}
      {routeError ? <StatusMessage text={routeError} themeMode={themeMode} danger /> : null}

      <div
        style={{
          display: "grid",
          gridTemplateColumns: "minmax(320px, 420px) 1fr",
          gap: 16,
          alignItems: "start",
        }}
      >
        <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
          <EventList
            events={activeDayEvents}
            selectedEventId={selectedEventId}
            onSelectEvent={setSelectedEventId}
            driveTimeByEventId={driveTimeByEventId}
            themeMode={themeMode}
          />

          <div
            style={{
              border: `1px solid ${theme.border}`,
              padding: 16,
              borderRadius: 12,
              background: theme.panelBg,
              boxShadow: themeMode === "dark" ? "0 10px 30px rgba(2, 6, 23, 0.22)" : "none",
            }}
          >
            <h2 style={{ marginTop: 0, fontSize: 18, color: theme.text }}>Route Preview</h2>

            {showRouteOverlay && routePreviews.length > 0 ? (
              <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
                {routePreviews.map((route) => {
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
                          fontWeight: 700,
                          color: theme.text,
                          marginBottom: 8,
                        }}
                      >
                        {route.dayLabel}
                      </div>
                      <div style={{ color: theme.mutedText, fontSize: 14, display: "grid", gap: 6 }}>
                        <div>
                          Start: <strong style={{ color: theme.text }}>Home Office</strong>
                        </div>
                        <div>
                          End: <strong style={{ color: theme.text }}>{returnToOffice ? "Home Office" : "Last stop"}</strong>
                        </div>
                        <div>
                          Stops: <strong style={{ color: theme.text }}>{route.orderedEventIds.length}</strong>
                        </div>
                        <div>
                          Drive time: <strong style={{ color: theme.text }}>{formatDuration(route.totalDurationSeconds)}</strong>
                        </div>
                        <div>
                          Event time: <strong style={{ color: theme.text }}>{formatDuration(route.totalEventDurationSeconds)}</strong>
                        </div>
                        <div>
                          Total field time: <strong style={{ color: theme.text }}>{formatDuration(route.totalFieldTimeSeconds)}</strong>
                        </div>
                        <div>
                          Drive distance: <strong style={{ color: theme.text }}>{formatMiles(route.totalDistanceMeters)}</strong>
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
            ) : (
              <p style={{ margin: 0, color: theme.mutedText }}>
                {showRouteOverlay
                  ? "Need at least one mapped event to preview a route from Home Office."
                  : "Route overlay is turned off."}
              </p>
            )}
          </div>

          <div
            style={{
              border: `1px solid ${theme.border}`,
              padding: 16,
              borderRadius: 12,
              background: theme.panelBg,
              boxShadow: themeMode === "dark" ? "0 10px 30px rgba(2, 6, 23, 0.22)" : "none",
            }}
          >
            <h2 style={{ marginTop: 0, fontSize: 18, color: theme.text }}>Selected Event</h2>

            {selectedEvent ? (
              <div>
                <h3 style={{ marginTop: 0, marginBottom: 8, color: theme.text }}>{selectedEvent.subject}</h3>
                <p style={{ margin: "0 0 8px 0", color: theme.subtleText }}>
                  {selectedEvent.addressText || "No usable address"}
                </p>
                <p style={{ margin: "0 0 8px 0", color: theme.mutedText }}>
                  Technicians: {selectedEvent.technicians.length > 0 ? selectedEvent.technicians.join(", ") : "None"}
                </p>
                <p style={{ margin: "0 0 12px 0", color: theme.mutedText }}>
                  Categories: {selectedEvent.categories.length > 0 ? selectedEvent.categories.join(", ") : "None"}
                </p>

                {selectedEvent.webLink ? (
                  <a
                    href={selectedEvent.webLink}
                    target="_blank"
                    rel="noreferrer"
                    style={{
                      color: theme.accent,
                      fontWeight: 600,
                      textDecoration: "none",
                    }}
                  >
                    Open in Outlook
                  </a>
                ) : null}
              </div>
            ) : (
              <p style={{ margin: 0, color: theme.mutedText }}>No event selected.</p>
            )}
          </div>
        </div>

        <EventMap
          events={geocodedEvents.filter((event) => {
            const dayKey = getDayKeyFromIso(event.startIso);
            return dayKey === activeRoutePreview?.dayKey || !showRouteOverlay;
          })}
          selectedEventId={selectedEventId}
          onSelectEvent={setSelectedEventId}
          routeGeometry={activeRoutePreview?.geometry ?? null}
          showRouteOverlay={showRouteOverlay}
        />
      </div>
    </div>
  );
}

function StatusMessage({
  text,
  themeMode,
  danger = false,
}: {
  text: string;
  themeMode: ThemeMode;
  danger?: boolean;
}) {
  const theme = getTheme(themeMode);

  return (
    <p
      style={{
        margin: "0 0 12px 0",
        padding: "10px 12px",
        borderRadius: 12,
        border: `1px solid ${danger ? (themeMode === "dark" ? "rgba(248,113,113,0.35)" : "rgba(220,38,38,0.18)") : theme.statusBorder}`,
        background: theme.statusBg,
        color: danger ? theme.danger : theme.subtleText,
      }}
    >
      {text}
    </p>
  );
}

function StatChip({
  label,
  value,
  themeMode,
}: {
  label: string;
  value: number;
  themeMode: ThemeMode;
}) {
  const theme = getTheme(themeMode);

  return (
    <div
      style={{
        background: theme.panelBg,
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
          color: theme.mutedText,
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
