export type DateRangeResult = {
  startIso: string;
  endIso: string;
};

export type DatePreset = "today" | "tomorrow" | "thisWeek" | "custom";

function startOfDay(date: Date): Date {
  const next = new Date(date);
  next.setHours(0, 0, 0, 0);
  return next;
}

function endOfDay(date: Date): Date {
  const next = new Date(date);
  next.setHours(23, 59, 59, 999);
  return next;
}

function addDays(date: Date, days: number): Date {
  const next = new Date(date);
  next.setDate(next.getDate() + days);
  return next;
}

function startOfWeek(date: Date): Date {
  const next = startOfDay(date);
  const day = next.getDay();
  const diff = day === 0 ? -6 : 1 - day;
  next.setDate(next.getDate() + diff);
  return next;
}

function endOfWeek(date: Date): Date {
  const start = startOfWeek(date);
  const end = addDays(start, 6);
  return endOfDay(end);
}

export function getDateRange(
  preset: DatePreset,
  customStartDate?: string,
  customEndDate?: string
): DateRangeResult {
  const now = new Date();

  if (preset === "today") {
    return {
      startIso: startOfDay(now).toISOString(),
      endIso: endOfDay(now).toISOString(),
    };
  }

  if (preset === "tomorrow") {
    const tomorrow = addDays(now, 1);
    return {
      startIso: startOfDay(tomorrow).toISOString(),
      endIso: endOfDay(tomorrow).toISOString(),
    };
  }

  if (preset === "thisWeek") {
    return {
      startIso: startOfWeek(now).toISOString(),
      endIso: endOfWeek(now).toISOString(),
    };
  }

  const safeStart = customStartDate ? new Date(`${customStartDate}T00:00:00`) : now;
  const safeEnd = customEndDate ? new Date(`${customEndDate}T23:59:59`) : now;

  return {
    startIso: safeStart.toISOString(),
    endIso: safeEnd.toISOString(),
  };
}