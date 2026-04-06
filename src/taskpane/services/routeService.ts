import { GeocodedCalendarEvent } from "./geocodeService";

declare const __MAPBOX_ACCESS_TOKEN__: string;

export type RouteGeometry = {
  type: "LineString";
  coordinates: number[][];
};

export type RouteLeg = {
  fromEventId: string;
  toEventId: string;
  fromLabel: string;
  toLabel: string;
  durationSeconds: number;
  distanceMeters: number;
};

export type RoutePreview = {
  dayKey: string;
  dayLabel: string;
  geometry: RouteGeometry | null;
  totalDurationSeconds: number;
  totalDistanceMeters: number;
  totalEventDurationSeconds: number;
  totalFieldTimeSeconds: number;
  legs: RouteLeg[];
  orderedEventIds: string[];
};

export type RouteAnchor = {
  id: string;
  label: string;
  latitude: number;
  longitude: number;
};

type MapboxRouteLeg = {
  duration?: number;
  distance?: number;
};

type MapboxRoute = {
  geometry?: RouteGeometry;
  duration?: number;
  distance?: number;
  legs?: MapboxRouteLeg[];
};

type MapboxDirectionsResponse = {
  routes?: MapboxRoute[];
};

function getMapboxToken(): string {
  return typeof __MAPBOX_ACCESS_TOKEN__ === "string" ? __MAPBOX_ACCESS_TOKEN__ : "";
}

function getDayKey(date: Date): string {
  const year = date.getFullYear();
  const month = `${date.getMonth() + 1}`.padStart(2, "0");
  const day = `${date.getDate()}`.padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function getDayLabel(dayKey: string): string {
  const parsed = new Date(`${dayKey}T12:00:00`);
  return new Intl.DateTimeFormat(undefined, {
    weekday: "short",
    month: "short",
    day: "numeric",
  }).format(parsed);
}

function getEventDurationSeconds(event: GeocodedCalendarEvent): number {
  const start = new Date(event.startIso);
  const end = new Date(event.endIso);

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

export function groupEventsByDay(
  events: GeocodedCalendarEvent[]
): Array<{ dayKey: string; dayLabel: string; events: GeocodedCalendarEvent[] }> {
  const groups = new Map<string, GeocodedCalendarEvent[]>();

  events.forEach((event) => {
    const start = new Date(event.startIso);
    if (Number.isNaN(start.getTime())) {
      return;
    }

    const key = getDayKey(start);
    const existing = groups.get(key) ?? [];
    existing.push(event);
    groups.set(key, existing);
  });

  return Array.from(groups.entries())
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([dayKey, groupedEvents]) => ({
      dayKey,
      dayLabel: getDayLabel(dayKey),
      events: groupedEvents.sort((a, b) => {
        const aTime = new Date(a.startIso).getTime();
        const bTime = new Date(b.startIso).getTime();
        return aTime - bTime;
      }),
    }));
}

async function getSingleDayRoutePreview(
  events: GeocodedCalendarEvent[],
  dayKey: string,
  dayLabel: string,
  startAnchor?: RouteAnchor | null,
  endAnchor?: RouteAnchor | null
): Promise<RoutePreview> {
  const orderedEventIds = events.map((event) => event.id);
  const totalEventDurationSeconds = events.reduce(
    (sum, event) => sum + getEventDurationSeconds(event),
    0
  );

  if (events.length === 0) {
    return {
      dayKey,
      dayLabel,
      geometry: null,
      totalDurationSeconds: 0,
      totalDistanceMeters: 0,
      totalEventDurationSeconds,
      totalFieldTimeSeconds: totalEventDurationSeconds,
      legs: [],
      orderedEventIds,
    };
  }

  const token = getMapboxToken();
  if (!token) {
    throw new Error("Missing Mapbox access token.");
  }

  const routePoints = [
    ...(startAnchor
      ? [
          {
            id: startAnchor.id,
            label: startAnchor.label,
            latitude: startAnchor.latitude,
            longitude: startAnchor.longitude,
          },
        ]
      : []),
    ...events.map((event) => ({
      id: event.id,
      label: event.subject,
      latitude: event.latitude,
      longitude: event.longitude,
    })),
    ...(endAnchor
      ? [
          {
            id: endAnchor.id,
            label: endAnchor.label,
            latitude: endAnchor.latitude,
            longitude: endAnchor.longitude,
          },
        ]
      : []),
  ];

  if (routePoints.length <= 1) {
    return {
      dayKey,
      dayLabel,
      geometry: null,
      totalDurationSeconds: 0,
      totalDistanceMeters: 0,
      totalEventDurationSeconds,
      totalFieldTimeSeconds: totalEventDurationSeconds,
      legs: [],
      orderedEventIds,
    };
  }

  const coordinates = routePoints
    .map((point) => `${point.longitude},${point.latitude}`)
    .join(";");

  const url = new URL(
    `https://api.mapbox.com/directions/v5/mapbox/driving-traffic/${coordinates}`
  );

  url.searchParams.set("access_token", token);
  url.searchParams.set("geometries", "geojson");
  url.searchParams.set("overview", "full");
  url.searchParams.set("steps", "false");

  const response = await fetch(url.toString());
  if (!response.ok) {
    throw new Error(`Route preview failed with status ${response.status}.`);
  }

  const data = (await response.json()) as MapboxDirectionsResponse;
  const route = Array.isArray(data.routes) ? data.routes[0] : undefined;

  if (!route) {
    return {
      dayKey,
      dayLabel,
      geometry: null,
      totalDurationSeconds: 0,
      totalDistanceMeters: 0,
      totalEventDurationSeconds,
      totalFieldTimeSeconds: totalEventDurationSeconds,
      legs: [],
      orderedEventIds,
    };
  }

  const legs: RouteLeg[] = Array.isArray(route.legs)
    ? route.legs.map((leg, index) => ({
        fromEventId: routePoints[index]?.id ?? `from-${index}`,
        toEventId: routePoints[index + 1]?.id ?? `to-${index + 1}`,
        fromLabel: routePoints[index]?.label ?? `Stop ${index + 1}`,
        toLabel: routePoints[index + 1]?.label ?? `Stop ${index + 2}`,
        durationSeconds: typeof leg.duration === "number" ? leg.duration : 0,
        distanceMeters: typeof leg.distance === "number" ? leg.distance : 0,
      }))
    : [];

  const totalDurationSeconds =
    typeof route.duration === "number" ? route.duration : 0;

  return {
    dayKey,
    dayLabel,
    geometry: route.geometry ?? null,
    totalDurationSeconds,
    totalDistanceMeters: typeof route.distance === "number" ? route.distance : 0,
    totalEventDurationSeconds,
    totalFieldTimeSeconds: totalDurationSeconds + totalEventDurationSeconds,
    legs,
    orderedEventIds,
  };
}

export async function getRoutePreviewsByDay(
  events: GeocodedCalendarEvent[],
  startAnchor?: RouteAnchor | null,
  endAnchor?: RouteAnchor | null
): Promise<RoutePreview[]> {
  const grouped = groupEventsByDay(events);

  return Promise.all(
    grouped.map((group) =>
      getSingleDayRoutePreview(group.events, group.dayKey, group.dayLabel, startAnchor, endAnchor)
    )
  );
}