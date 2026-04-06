export type DatePreset = "today" | "tomorrow" | "thisWeek" | "custom";

export type FilterState = {
  preset: DatePreset;
  startDate?: string;
  endDate?: string;
  categories: string[];
  technicians: string[];
  calendarIds: string[];
};