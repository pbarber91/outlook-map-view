import * as React from "react";
import { MappedCalendarEvent } from "../types/calendar";
import { getCategoryColor } from "../utils/categoryColors";

type ThemeMode = "light" | "dark";

type EventListProps = {
  events: MappedCalendarEvent[];
  selectedEventId: string | null;
  onSelectEvent: (eventId: string) => void;
  driveTimeByEventId?: Record<string, number>;
  themeMode?: ThemeMode;
};

function formatEventDateRange(startIso: string, endIso: string): string {
  const start = new Date(startIso);
  const end = new Date(endIso);

  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) {
    return "Invalid date";
  }

  const sameDay = start.toDateString() === end.toDateString();
  const allDay =
    start.getHours() === 0 &&
    start.getMinutes() === 0 &&
    end.getHours() === 0 &&
    end.getMinutes() === 0 &&
    Math.abs(end.getTime() - start.getTime()) >= 23 * 60 * 60 * 1000;

  const dateFormatter = new Intl.DateTimeFormat(undefined, {
    month: "short",
    day: "numeric",
    year: "numeric",
  });

  const timeFormatter = new Intl.DateTimeFormat(undefined, {
    hour: "numeric",
    minute: "2-digit",
  });

  if (allDay) {
    return `${dateFormatter.format(start)} • All day`;
  }

  if (sameDay) {
    return `${dateFormatter.format(start)} • ${timeFormatter.format(start)} - ${timeFormatter.format(end)}`;
  }

  return `${dateFormatter.format(start)} ${timeFormatter.format(start)} - ${dateFormatter.format(end)} ${timeFormatter.format(end)}`;
}

function formatDuration(totalSeconds: number): string {
  const roundedMinutes = Math.round(totalSeconds / 60);
  const hours = Math.floor(roundedMinutes / 60);
  const minutes = roundedMinutes % 60;

  if (hours <= 0) {
    return `${minutes}m drive`;
  }

  return `${hours}h ${minutes}m drive`;
}

function getTheme(mode: ThemeMode) {
  if (mode === "dark") {
    return {
      panelBg: "#111827",
      border: "#334155",
      text: "#f8fafc",
      muted: "#94a3b8",
      countBg: "#1f2937",
      countText: "#cbd5e1",
      emptyBg: "#0f172a",
      emptyText: "#94a3b8",
      itemBg: "#111827",
      itemSelectedBg: "rgba(37,99,235,0.16)",
      badgeBg: "#0f172a",
      badgeText: "#e2e8f0",
      noCategories: "#64748b",
      mapYesBg: "rgba(16,185,129,0.18)",
      mapYesText: "#a7f3d0",
      mapNoBg: "rgba(245,158,11,0.18)",
      mapNoText: "#fde68a",
    };
  }

  return {
    panelBg: "#ffffff",
    border: "#d1d5db",
    text: "#111827",
    muted: "#6b7280",
    countBg: "#f3f4f6",
    countText: "#4b5563",
    emptyBg: "#f9fafb",
    emptyText: "#6b7280",
    itemBg: "#ffffff",
    itemSelectedBg: "#eff6ff",
    badgeBg: "#f8fafc",
    badgeText: "#0f172a",
    noCategories: "#9ca3af",
    mapYesBg: "#d1fae5",
    mapYesText: "#065f46",
    mapNoBg: "#fef3c7",
    mapNoText: "#92400e",
  };
}

export default function EventList({
  events,
  selectedEventId,
  onSelectEvent,
  driveTimeByEventId = {},
  themeMode = "light",
}: EventListProps) {
  const theme = getTheme(themeMode);

  return (
    <div
      style={{
        border: `1px solid ${theme.border}`,
        padding: 16,
        borderRadius: 12,
        background: theme.panelBg,
        boxShadow: themeMode === "dark" ? "0 10px 30px rgba(2, 6, 23, 0.22)" : "none",
      }}
    >
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          marginBottom: 12,
        }}
      >
        <h2 style={{ margin: 0, fontSize: 18, color: theme.text }}>Events</h2>
        <span
          style={{
            fontSize: 12,
            color: theme.countText,
            background: theme.countBg,
            borderRadius: 999,
            padding: "4px 8px",
          }}
        >
          {events.length}
        </span>
      </div>

      {events.length === 0 ? (
        <div
          style={{
            border: `1px dashed ${theme.border}`,
            borderRadius: 10,
            padding: 16,
            color: theme.emptyText,
            background: theme.emptyBg,
          }}
        >
          No events to display.
        </div>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {events.map((event) => {
            const isSelected = selectedEventId === event.id;
            const primaryCategory = event.categories[0];
            const primaryColor = getCategoryColor(primaryCategory);
            const driveTime = driveTimeByEventId[event.id] ?? 0;

            return (
              <button
                key={event.id}
                type="button"
                onClick={() => onSelectEvent(event.id)}
                style={{
                  textAlign: "left",
                  border: isSelected ? "2px solid #2563eb" : `1px solid ${theme.border}`,
                  background: isSelected ? theme.itemSelectedBg : theme.itemBg,
                  borderRadius: 12,
                  padding: 14,
                  cursor: "pointer",
                  boxShadow: isSelected ? "0 0 0 3px rgba(37,99,235,0.08)" : "none",
                }}
              >
                <div
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "flex-start",
                    gap: 12,
                    marginBottom: 8,
                  }}
                >
                  <div style={{ display: "flex", alignItems: "center", gap: 10, minWidth: 0 }}>
                    <span
                      style={{
                        width: 12,
                        height: 12,
                        borderRadius: 999,
                        background: primaryColor.marker,
                        flex: "0 0 auto",
                      }}
                    />
                    <h3 style={{ margin: 0, fontSize: 15, lineHeight: 1.35, color: theme.text }}>
                      {event.subject || "(No subject)"}
                    </h3>
                  </div>

                  {event.addressText ? (
                    <span
                      style={{
                        fontSize: 11,
                        fontWeight: 700,
                        color: theme.mapYesText,
                        background: theme.mapYesBg,
                        borderRadius: 999,
                        padding: "4px 8px",
                        whiteSpace: "nowrap",
                      }}
                    >
                      Mappable
                    </span>
                  ) : (
                    <span
                      style={{
                        fontSize: 11,
                        fontWeight: 700,
                        color: theme.mapNoText,
                        background: theme.mapNoBg,
                        borderRadius: 999,
                        padding: "4px 8px",
                        whiteSpace: "nowrap",
                      }}
                    >
                      No map
                    </span>
                  )}
                </div>

                <div style={{ fontSize: 13, color: theme.text, opacity: 0.88, marginBottom: 8 }}>
                  {formatEventDateRange(event.startIso, event.endIso)}
                </div>

                {driveTime > 0 ? (
                  <div
                    style={{
                      fontSize: 12,
                      fontWeight: 700,
                      color: theme.badgeText,
                      background: theme.badgeBg,
                      border: `1px solid ${theme.border}`,
                      borderRadius: 999,
                      display: "inline-block",
                      padding: "4px 8px",
                      marginBottom: 8,
                    }}
                  >
                    {formatDuration(driveTime)}
                  </div>
                ) : null}

                <div style={{ fontSize: 13, color: theme.muted, marginBottom: 10 }}>
                  {event.addressText || "No usable address"}
                </div>

                <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                  {event.categories.length > 0 ? (
                    event.categories.map((category) => {
                      const color = getCategoryColor(category);

                      return (
                        <span
                          key={category}
                          style={{
                            fontSize: 12,
                            borderRadius: 999,
                            background: color.background,
                            color: color.text,
                            border: `1px solid ${color.border}`,
                            padding: "4px 8px",
                          }}
                        >
                          {category}
                        </span>
                      );
                    })
                  ) : (
                    <span style={{ fontSize: 12, color: theme.noCategories }}>No categories</span>
                  )}
                </div>
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
}
