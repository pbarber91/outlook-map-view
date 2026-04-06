export function getCategoryColor(category?: string): {
  background: string;
  text: string;
  border: string;
  marker: string;
} {
  const value = (category || "").trim().toLowerCase();

  const palette: Array<{
    keys: string[];
    background: string;
    text: string;
    border: string;
    marker: string;
  }> = [
    {
      keys: ["inspection", "site visit", "field"],
      background: "#dbeafe",
      text: "#1d4ed8",
      border: "#93c5fd",
      marker: "#2563eb",
    },
    {
      keys: ["follow-up", "follow up", "recheck"],
      background: "#dcfce7",
      text: "#166534",
      border: "#86efac",
      marker: "#16a34a",
    },
    {
      keys: ["admin", "office", "internal"],
      background: "#f3f4f6",
      text: "#374151",
      border: "#d1d5db",
      marker: "#6b7280",
    },
    {
      keys: ["meeting", "call", "teams", "zoom"],
      background: "#ede9fe",
      text: "#6d28d9",
      border: "#c4b5fd",
      marker: "#7c3aed",
    },
    {
      keys: ["travel", "drive"],
      background: "#ffedd5",
      text: "#c2410c",
      border: "#fdba74",
      marker: "#ea580c",
    },
    {
      keys: ["court", "legal", "deposition"],
      background: "#fee2e2",
      text: "#b91c1c",
      border: "#fca5a5",
      marker: "#dc2626",
    },
  ];

  for (const item of palette) {
    if (item.keys.some((key) => value.includes(key))) {
      return {
        background: item.background,
        text: item.text,
        border: item.border,
        marker: item.marker,
      };
    }
  }

  return {
    background: "#eef2ff",
    text: "#4338ca",
    border: "#c7d2fe",
    marker: "#4f46e5",
  };
}