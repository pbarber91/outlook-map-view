import * as React from "react";
import { DatePreset, FilterState } from "../types/filters";
import { CalendarSource } from "../types/calendar";
import { getCategoryColor } from "../utils/categoryColors";

type FilterBarProps = {
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
}: FilterBarProps) {
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
        background: "#ffffff",
        border: "1px solid #dbe2ea",
        borderRadius: 16,
        padding: 16,
        boxShadow: "0 1px 2px rgba(0,0,0,0.04)",
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
              color: "#475569",
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
              border: "1px solid #cbd5e1",
              borderRadius: 10,
              padding: "0 12px",
              background: "#ffffff",
              color: "#0f172a",
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
                  color: "#475569",
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
                  border: "1px solid #cbd5e1",
                  borderRadius: 10,
                  padding: "0 12px",
                  background: "#ffffff",
                  color: "#0f172a",
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
                  color: "#475569",
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
                  border: "1px solid #cbd5e1",
                  borderRadius: 10,
                  padding: "0 12px",
                  background: "#ffffff",
                  color: "#0f172a",
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
              color: "#64748b",
              background: "#f8fafc",
              border: "1px dashed #cbd5e1",
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
                    border: active ? "1px solid #93c5fd" : "1px solid #d1d5db",
                    background: active ? "#dbeafe" : "#ffffff",
                    color: active ? "#1d4ed8" : "#334155",
                    boxShadow: active ? "0 0 0 2px rgba(15,23,42,0.04)" : "none",
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
              color: "#64748b",
              background: "#f8fafc",
              border: "1px dashed #cbd5e1",
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
                    border: active ? "1px solid #93c5fd" : "1px solid #d1d5db",
                    background: active ? "#dbeafe" : "#ffffff",
                    color: active ? "#1d4ed8" : "#334155",
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
              color: "#64748b",
              background: "#f8fafc",
              border: "1px dashed #cbd5e1",
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
                      : "1px solid #d1d5db",
                    background: active ? color.background : "#ffffff",
                    color: active ? color.text : "#334155",
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