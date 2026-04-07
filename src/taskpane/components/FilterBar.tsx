import * as React from "react";
import { DatePreset, FilterState } from "../types/filters";
import { CalendarSource } from "../types/calendar";
import { getCategoryColor } from "../utils/categoryColors";

type ThemeMode = "light" | "dark";

type FilterBarProps = {
  themeMode?: ThemeMode;
  filters: FilterState;
  availableCategories: string[];
  availableTechnicians: string[];
  availableCalendars: CalendarSource[];
  onPresetChange: (preset: DatePreset) => void;
  onCustomDateChange: (field: "startDate" | "endDate", value: string) => void;
  onCategoriesChange: (categories: string[]) => void;
  onTechniciansChange: (technicians: string[]) => void;
  onCalendarIdsChange: (calendarIds: string[]) => void;
  onRefresh: () => void;
};

const presetOptions: Array<{ value: DatePreset; label: string }> = [
  { value: "today", label: "Today" },
  { value: "tomorrow", label: "Tomorrow" },
  { value: "thisWeek", label: "This Week" },
  { value: "custom", label: "Custom" },
];

export default function FilterBar({
  filters,
  availableCategories,
  availableTechnicians,
  availableCalendars,
  onPresetChange,
  onCustomDateChange,
  onCategoriesChange,
  onTechniciansChange,
  onCalendarIdsChange,
  onRefresh,
  themeMode = "light",
}: FilterBarProps) {
  const isDark = themeMode === "dark";
  const theme = isDark
    ? {
        panelBg: "#111827",
        border: "#334155",
        borderSoft: "#475569",
        text: "#f8fafc",
        textSoft: "#cbd5e1",
        textMuted: "#94a3b8",
        controlBg: "#0f172a",
        chipBg: "#111827",
        chipInactiveBg: "#0f172a",
        chipInactiveText: "#cbd5e1",
        emptyBg: "#0f172a",
        shadow: "0 12px 28px rgba(2, 6, 23, 0.24)",
      }
    : {
        panelBg: "#ffffff",
        border: "#dbe2ea",
        borderSoft: "#cbd5e1",
        text: "#0f172a",
        textSoft: "#475569",
        textMuted: "#64748b",
        controlBg: "#ffffff",
        chipBg: "#ffffff",
        chipInactiveBg: "#ffffff",
        chipInactiveText: "#334155",
        emptyBg: "#f8fafc",
        shadow: "0 1px 2px rgba(0,0,0,0.04)",
      };
  const handleCategoryToggle = (category: string) => {
    const exists = filters.categories.includes(category);

    if (exists) {
      onCategoriesChange(filters.categories.filter((c) => c !== category));
      return;
    }

    onCategoriesChange([...filters.categories, category]);
  };

  const handleTechnicianToggle = (technician: string) => {
    const exists = filters.technicians.includes(technician);

    if (exists) {
      onTechniciansChange(filters.technicians.filter((t) => t !== technician));
      return;
    }

    onTechniciansChange([...filters.technicians, technician]);
  };

  const handleCalendarToggle = (calendarId: string) => {
    const exists = filters.calendarIds.includes(calendarId);

    if (exists) {
      onCalendarIdsChange(filters.calendarIds.filter((id) => id !== calendarId));
      return;
    }

    onCalendarIdsChange([...filters.calendarIds, calendarId]);
  };

  return (
    <div
      style={{
        background: theme.panelBg,
        border: `1px solid ${theme.border}`,
        borderRadius: 16,
        padding: 16,
        boxShadow: theme.shadow,
      }}
    >
      <div
        style={{
          display: "grid",
          gridTemplateColumns: "minmax(180px, 220px) auto auto 1fr",
          gap: 12,
          alignItems: "end",
        }}
      >
        <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
          <label
            htmlFor="datePreset"
            style={{
              fontSize: 12,
              fontWeight: 700,
              color: theme.textSoft,
              textTransform: "uppercase",
              letterSpacing: "0.04em",
            }}
          >
            Date Range
          </label>
          <select
            id="datePreset"
            value={filters.preset}
            onChange={(e) => onPresetChange(e.target.value as DatePreset)}
            style={{
              height: 40,
              border: `1px solid ${theme.borderSoft}`,
              borderRadius: 10,
              padding: "0 12px",
              background: theme.panelBg,
              color: theme.text,
              fontSize: 14,
            }}
          >
            {presetOptions.map((option) => (
              <option key={option.value} value={option.value}>
                {option.label}
              </option>
            ))}
          </select>
        </div>

        {filters.preset === "custom" && (
          <>
            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              <label
                htmlFor="startDate"
                style={{
                  fontSize: 12,
                  fontWeight: 700,
                  color: theme.textSoft,
                  textTransform: "uppercase",
                  letterSpacing: "0.04em",
                }}
              >
                Start
              </label>
              <input
                id="startDate"
                type="date"
                value={filters.startDate ?? ""}
                onChange={(e) => onCustomDateChange("startDate", e.target.value)}
                style={{
                  height: 40,
                  border: `1px solid ${theme.borderSoft}`,
                  borderRadius: 10,
                  padding: "0 12px",
                  background: theme.panelBg,
                  color: theme.text,
                  fontSize: 14,
                }}
              />
            </div>

            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              <label
                htmlFor="endDate"
                style={{
                  fontSize: 12,
                  fontWeight: 700,
                  color: theme.textSoft,
                  textTransform: "uppercase",
                  letterSpacing: "0.04em",
                }}
              >
                End
              </label>
              <input
                id="endDate"
                type="date"
                value={filters.endDate ?? ""}
                onChange={(e) => onCustomDateChange("endDate", e.target.value)}
                style={{
                  height: 40,
                  border: `1px solid ${theme.borderSoft}`,
                  borderRadius: 10,
                  padding: "0 12px",
                  background: theme.panelBg,
                  color: theme.text,
                  fontSize: 14,
                }}
              />
            </div>
          </>
        )}

        <div
          style={{
            display: "flex",
            justifyContent: filters.preset === "custom" ? "flex-start" : "flex-end",
            alignItems: "end",
            minWidth: 120,
          }}
        >
          <button
            type="button"
            onClick={onRefresh}
            style={{
              height: 40,
              padding: "0 14px",
              borderRadius: 10,
              border: "1px solid #2563eb",
              background: "#2563eb",
              color: "#ffffff",
              fontSize: 14,
              fontWeight: 700,
              cursor: "pointer",
              boxShadow: "0 1px 2px rgba(37,99,235,0.18)",
            }}
          >
            Refresh
          </button>
        </div>
      </div>

      <div style={{ marginTop: 16 }}>
        <div
          style={{
            fontSize: 12,
            fontWeight: 700,
            color: "#475569",
            textTransform: "uppercase",
            letterSpacing: "0.04em",
            marginBottom: 8,
          }}
        >
          Calendars
        </div>

        {availableCalendars.length === 0 ? (
          <div
            style={{
              fontSize: 13,
              color: theme.textMuted,
              background: theme.emptyBg,
              border: `1px dashed ${theme.borderSoft}`,
              borderRadius: 10,
              padding: "10px 12px",
              marginBottom: 14,
            }}
          >
            No accessible calendars found.
          </div>
        ) : (
          <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 14 }}>
            {availableCalendars.map((calendar) => {
              const active = filters.calendarIds.includes(calendar.id);

              return (
                <button
                  key={calendar.id}
                  type="button"
                  onClick={() => handleCalendarToggle(calendar.id)}
                  style={{
                    borderRadius: 999,
                    padding: "7px 12px",
                    cursor: "pointer",
                    fontSize: 13,
                    fontWeight: 700,
                    border: active ? "1px solid #93c5fd" : `1px solid ${theme.border}`,
                    background: active ? (isDark ? "rgba(37,99,235,0.18)" : "#dbeafe") : theme.chipInactiveBg,
                    color: active ? "#2563eb" : theme.chipInactiveText,
                    boxShadow: active ? (isDark ? "0 0 0 1px rgba(96,165,250,0.18)" : "0 0 0 2px rgba(15,23,42,0.04)") : "none",
                  }}
                >
                  {calendar.name}
                  {calendar.isDefaultCalendar ? " • Default" : ""}
                </button>
              );
            })}
          </div>
        )}

        <div
          style={{
            fontSize: 12,
            fontWeight: 700,
            color: "#475569",
            textTransform: "uppercase",
            letterSpacing: "0.04em",
            marginBottom: 8,
          }}
        >
          Technicians
        </div>

        {availableTechnicians.length === 0 ? (
          <div
            style={{
              fontSize: 13,
              color: theme.textMuted,
              background: theme.emptyBg,
              border: `1px dashed ${theme.borderSoft}`,
              borderRadius: 10,
              padding: "10px 12px",
              marginBottom: 14,
            }}
          >
            No required attendees found in the current view.
          </div>
        ) : (
          <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 14 }}>
            {availableTechnicians.map((technician) => {
              const active = filters.technicians.includes(technician);

              return (
                <button
                  key={technician}
                  type="button"
                  onClick={() => handleTechnicianToggle(technician)}
                  style={{
                    borderRadius: 999,
                    padding: "7px 12px",
                    cursor: "pointer",
                    fontSize: 13,
                    fontWeight: 700,
                    border: active ? "1px solid #93c5fd" : `1px solid ${theme.border}`,
                    background: active ? (isDark ? "rgba(37,99,235,0.18)" : "#dbeafe") : theme.chipInactiveBg,
                    color: active ? "#2563eb" : theme.chipInactiveText,
                    boxShadow: active ? (isDark ? "0 0 0 1px rgba(96,165,250,0.18)" : "0 0 0 2px rgba(15,23,42,0.04)") : "none",
                  }}
                >
                  {technician}
                </button>
              );
            })}
          </div>
        )}

        <div
          style={{
            fontSize: 12,
            fontWeight: 700,
            color: "#475569",
            textTransform: "uppercase",
            letterSpacing: "0.04em",
            marginBottom: 8,
          }}
        >
          Categories
        </div>

        {availableCategories.length === 0 ? (
          <div
            style={{
              fontSize: 13,
              color: theme.textMuted,
              background: theme.emptyBg,
              border: `1px dashed ${theme.borderSoft}`,
              borderRadius: 10,
              padding: "10px 12px",
            }}
          >
            No categories available in the current view.
          </div>
        ) : (
          <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
            {availableCategories.map((category) => {
              const active = filters.categories.includes(category);
              const color = getCategoryColor(category);

              return (
                <button
                  key={category}
                  type="button"
                  onClick={() => handleCategoryToggle(category)}
                  style={{
                    borderRadius: 999,
                    padding: "7px 12px",
                    cursor: "pointer",
                    fontSize: 13,
                    fontWeight: 700,
                    border: active
                      ? `1px solid ${color.border}`
                      : `1px solid ${theme.border}`,
                    background: active ? color.background : theme.chipInactiveBg,
                    color: active ? color.text : theme.chipInactiveText,
                    boxShadow: active ? (isDark ? "0 0 0 1px rgba(96,165,250,0.18)" : "0 0 0 2px rgba(15,23,42,0.04)") : "none",
                  }}
                >
                  {category}
                </button>
              );
            })}
          </div>
        )}
      </div>
    </div>
  );
}