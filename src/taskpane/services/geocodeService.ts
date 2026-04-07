import { MappedCalendarEvent } from "../types/calendar";

declare const __MAPBOX_ACCESS_TOKEN__: string;

export type GeocodedCalendarEvent = MappedCalendarEvent & {
  latitude: number;
  longitude: number;
};

type MapboxFeature = {
  center?: [number, number];
};

type MapboxResponse = {
  features?: MapboxFeature[];
};

type CachedCoords = {
  latitude: number;
  longitude: number;
};

const STORAGE_KEY = "outlook-map-view-geocode-cache-v1";
const geocodeCache = new Map<string, CachedCoords | null>();

function getMapboxToken(): string {
  return typeof __MAPBOX_ACCESS_TOKEN__ === "string" ? __MAPBOX_ACCESS_TOKEN__ : "";
}

function canGeocode(address: string | null): address is string {
  return typeof address === "string" && address.trim().length > 0;
}

function normalizeWhitespace(value: string): string {
  return value.replace(/\s+/g, " ").trim();
}

function stripTrailingCoordinates(value: string): string {
  return value.replace(/;\s*-?\d{1,3}\.\d+\s*,\s*-?\d{1,3}\.\d+\s*$/i, "").trim();
}

function pickBestAddressSegment(value: string): string {
  const segments = value
    .split(";")
    .map((segment) => normalizeWhitespace(segment))
    .filter(Boolean);

  if (segments.length === 0) {
    return "";
  }

  const streetLike = segments.find((segment) => /\d+\s+[a-z0-9]/i.test(segment));

  return streetLike || segments[0];
}

function sanitizeAddressForGeocoding(address: string): string {
  const noCoords = stripTrailingCoordinates(address);
  const bestSegment = pickBestAddressSegment(noCoords);

  return normalizeWhitespace(bestSegment);
}

function isValidLatitude(value: number): boolean {
  return Number.isFinite(value) && value >= -90 && value <= 90;
}

function isValidLongitude(value: number): boolean {
  return Number.isFinite(value) && value >= -180 && value <= 180;
}

function parseEmbeddedCoordinates(value: string): CachedCoords | null {
  const normalized = normalizeWhitespace(value);

  const plainPairMatch = normalized.match(
    /(-?\d{1,2}(?:\.\d+)?)\s*,\s*(-?\d{1,3}(?:\.\d+)?)/
  );

  if (!plainPairMatch) {
    return null;
  }

  const latitude = Number(plainPairMatch[1]);
  const longitude = Number(plainPairMatch[2]);

  if (!isValidLatitude(latitude) || !isValidLongitude(longitude)) {
    return null;
  }

  return { latitude, longitude };
}

function loadPersistentCache(): void {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return;

    const parsed = JSON.parse(raw) as Record<string, CachedCoords | null>;
    Object.entries(parsed).forEach(([key, value]) => {
      if (
        value === null ||
        (typeof value === "object" &&
          value !== null &&
          typeof value.latitude === "number" &&
          typeof value.longitude === "number")
      ) {
        geocodeCache.set(key, value);
      }
    });
  } catch {
    // ignore bad cache data
  }
}

function persistCache(): void {
  try {
    const obj: Record<string, CachedCoords | null> = {};
    geocodeCache.forEach((value, key) => {
      obj[key] = value;
    });
    localStorage.setItem(STORAGE_KEY, JSON.stringify(obj));
  } catch {
    // ignore storage failures
  }
}

loadPersistentCache();

async function geocodeAddress(address: string): Promise<CachedCoords | null> {
  const directCoords = parseEmbeddedCoordinates(address);
  if (directCoords) {
    return directCoords;
  }

  const cleaned = sanitizeAddressForGeocoding(address);

  if (!cleaned) {
    return null;
  }

  if (geocodeCache.has(cleaned)) {
    return geocodeCache.get(cleaned) ?? null;
  }

  const token = getMapboxToken();

  if (!token) {
    throw new Error("Missing Mapbox access token.");
  }

  const url = new URL(
    `https://api.mapbox.com/geocoding/v5/mapbox.places/${encodeURIComponent(cleaned)}.json`
  );

  url.searchParams.set("access_token", token);
  url.searchParams.set("limit", "1");
  url.searchParams.set("country", "US");

  const response = await fetch(url.toString());

  if (!response.ok) {
    const body = await response.text();
    console.error("Original geocode address:", address);
    console.error("Sanitized geocode address:", cleaned);
    console.error("Geocode failed status:", response.status);
    console.error("Geocode failed body:", body);
    throw new Error(`Geocoding failed with status ${response.status}.`);
  }

  const data = (await response.json()) as MapboxResponse;
  const first = Array.isArray(data.features) ? data.features[0] : undefined;
  const center = first?.center;

  if (!center || center.length !== 2) {
    geocodeCache.set(cleaned, null);
    persistCache();
    return null;
  }

  const result: CachedCoords = {
    longitude: center[0],
    latitude: center[1],
  };

  geocodeCache.set(cleaned, result);
  persistCache();
  return result;
}

export async function geocodeEvents(
  events: MappedCalendarEvent[]
): Promise<GeocodedCalendarEvent[]> {
  const results = await Promise.all(
    events.map(async (event) => {
      if (!canGeocode(event.addressText)) {
        return null;
      }

      const coords = await geocodeAddress(event.addressText);
      if (!coords) {
        return null;
      }

      return {
        ...event,
        latitude: coords.latitude,
        longitude: coords.longitude,
      };
    })
  );

  return results.filter((event): event is GeocodedCalendarEvent => !!event);
}

export function getGeocodeCacheSize(): number {
  return geocodeCache.size;
}

export function clearGeocodeCache(): void {
  geocodeCache.clear();
  try {
    localStorage.removeItem(STORAGE_KEY);
  } catch {
    // ignore storage failures
  }
}