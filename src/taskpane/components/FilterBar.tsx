import * as React from "react";
import { DatePreset, FilterState } from "../types/filters";
import { CalendarSource } from "../types/calendar";
import { getCategoryColor } from "../utils/categoryColors";

type ThemeMode = "light" | "dark";

type FilterBarProps = {
  filters: FilterState;
  availableCategories: string[];
  availableTechnicians: string[];
  availableCalendars?: CalendarSource[];
  onPresetChange: (preset: DatePreset) => void;
  onCustomDateChange: (field: "startDate" | "endDate", value: string) => void;
  onCategoriesChange: (categories: string[]) => void;
  onTechniciansChange: (technicians: string[]) => void;
  onCalendarIdsChange?: (calendarIds: string[]) => void;
  onRefresh: () => void;
  themeMode?: ThemeMode;
};

const presetOptions: Array<{ value: DatePreset; label: string }> = [
  { value: "today", label: "Today" },
  { value: "tomorrow", label: "Tomorrow" },
  { value: "thisWeek", label: "This Week" },
  { value: "custom", label: "Custom" },
];

function getTheme(mode: ThemeMode) {
  if (mode === "dark") {
    return {
      panelBg: "#111827",
      inputBg: "#0f172a",
      border: "#334155",
      borderSoft: "#475569",
      text: "#f8fafc",
      muted: "#94a3b8",
      sectionLabel: "#cbd5e1",
      emptyBg: "#0f172a",
      emptyText: "#94a3b8",
      chipBg: "#111827",
      chipText: "#cbd5e1",
      chipBorder: "#475569",
      techActiveBg: "rgba(59,130,246,0.18)",
      techActiveBorder: "#60a5fa",
      techActiveText: "#bfdbfe",
      buttonShadow: "0 1px 2px rgba(15, 23, 42, 0.45)",
    };
  }

  return {
    panelBg: "#ffffff",
    inputBg: "#ffffff",
    border: "#dbe2ea",
    borderSoft: "#cbd5e1",
    text: "#0f172a",
    muted: "#64748b",
    sectionLabel: "#475569",
    emptyBg: "#f8fafc",
    emptyText: "#64748b",
    chipBg: "#ffffff",
    chipText: "#334155",
    chipBorder: "#d1d5db",
    techActiveBg: "#dbeafe",
    techActiveBorder: "#93c5fd",
    techActiveText: "#1d4ed8",
    buttonShadow: "0 1px 2px rgba(37,99,235,0.18)",
  };
}

export default function FilterBar({
  filters,
  availableCategories,
  availableTechnicians,
  availableCalendars = [],
  onPresetChange,
  onCustomDateChange,
  onCategoriesChange,
  onTechniciansChange,
  onCalendarIdsChange,
  onRefresh,
  themeMode = "light",
}: FilterBarProps) {
  const theme = getTheme(themeMode);

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
    if (!onCalendarIdsChange) return;

    const existingIds = filters.calendarIds ?? [];
    const exists = existingIds.includes(calendarId);

    if (exists) {
      onCalendarIdsChange(existingIds.filter((id) => id !== calendarId));
      return;
    }

    onCalendarIdsChange([...existingIds, calendarId]);
  };

  return (
    <div
      style={{
        background: theme.panelBg,
        border: `1px solid ${theme.border}`,
        borderRadius: 16,
        padding: 16,
        boxShadow: themeMode === "dark" ? "0 10px 30px rgba(2, 6, 23, 0.28)" : "0 1px 2px rgba(0,0,0,0.04)",
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
              color: theme.sectionLabel,
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
              background: theme.inputBg,
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
                  color: theme.sectionLabel,
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
                  background: theme.inputBg,
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
                  color: theme.sectionLabel,
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
                  background: theme.inputBg,
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
              boxShadow: theme.buttonShadow,
            }}
          >
            Refresh
          </button>
        </div>
      </div>

      {availableCalendars.length > 0 ? (
        <div style={{ marginTop: 16 }}>
          <div
            style={{
              fontSize: 12,
              fontWeight: 700,
              color: theme.sectionLabel,
              textTransform: "uppercase",
              letterSpacing: "0.04em",
              marginBottom: 8,
            }}
          >
            Calendars
          </div>

          <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 14 }}>
            {availableCalendars.map((calendar) => {
              const active = (filters.calendarIds ?? []).includes(calendar.id);

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
                    border: active ? "1px solid #93c5fd" : `1px solid ${theme.chipBorder}`,
                    background: active ? (themeMode === "dark" ? "rgba(59,130,246,0.18)" : "#dbeafe") : theme.chipBg,
                    color: active ? (themeMode === "dark" ? "#bfdbfe" : "#1d4ed8") : theme.chipText,
                    boxShadow: active ? "0 0 0 2px rgba(15,23,42,0.04)" : "none",
                  }}
                >
                  {calendar.name}
                  {calendar.isDefaultCalendar ? " • Default" : ""}
                </button>
              );
            })}
          </div>
        </div>
      ) : null}

      <div style={{ marginTop: availableCalendars.length > 0 ? 0 : 16 }}>
        <div
          style={{
            fontSize: 12,
            fontWeight: 700,
            color: theme.sectionLabel,
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
              color: theme.emptyText,
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
                    border: active ? `1px solid ${theme.techActiveBorder}` : `1px solid ${theme.chipBorder}`,
                    background: active ? theme.techActiveBg : theme.chipBg,
                    color: active ? theme.techActiveText : theme.chipText,
                    boxShadow: active ? "0 0 0 2px rgba(15,23,42,0.04)" : "none",
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
            color: theme.sectionLabel,
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
              color: theme.emptyText,
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
                    border: active ? `1px solid ${color.border}` : `1px solid ${theme.chipBorder}`,
                    background: active ? color.background : theme.chipBg,
                    color: active ? color.text : theme.chipText,
                    boxShadow: active ? "0 0 0 2px rgba(15,23,42,0.04)" : "none",
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
