import * as React from "react";
import mapboxgl from "mapbox-gl";
import { GeocodedCalendarEvent } from "../services/geocodeService";
import { RouteGeometry } from "../services/routeService";
import { getCategoryColor } from "../utils/categoryColors";

declare const __MAPBOX_ACCESS_TOKEN__: string;

type EventMapProps = {
  events: GeocodedCalendarEvent[];
  selectedEventId: string | null;
  onSelectEvent: React.Dispatch<React.SetStateAction<string | null>>;
  routeGeometry?: RouteGeometry | null;
  showRouteOverlay?: boolean;
  standalone?: boolean;
};

type MarkerEntry = {
  marker: mapboxgl.Marker;
  popup: mapboxgl.Popup;
  eventId: string;
};

const ROUTE_SOURCE_ID = "route-preview-source";
const ROUTE_LAYER_ID = "route-preview-layer";

function getMapboxToken(): string {
  return typeof __MAPBOX_ACCESS_TOKEN__ === "string" ? __MAPBOX_ACCESS_TOKEN__ : "";
}

function buildMarkerElement(
  color: string,
  isSelected: boolean,
  title: string,
  orderNumber: number
): HTMLDivElement {
  const el = document.createElement("div");
  el.title = title;
  el.style.width = isSelected ? "32px" : "28px";
  el.style.height = isSelected ? "32px" : "28px";
  el.style.borderRadius = "999px";
  el.style.background = color;
  el.style.border = isSelected ? "4px solid #111827" : "4px solid #ffffff";
  el.style.boxShadow = "0 2px 10px rgba(0,0,0,0.45)";
  el.style.cursor = "pointer";
  el.style.display = "flex";
  el.style.alignItems = "center";
  el.style.justifyContent = "center";
  el.style.color = "#ffffff";
  el.style.fontSize = isSelected ? "13px" : "12px";
  el.style.fontWeight = "700";
  el.style.lineHeight = "1";
  el.style.zIndex = "999";
  el.textContent = String(orderNumber);
  return el;
}

function ensureRouteLayer(map: mapboxgl.Map): void {
  if (!map.getSource(ROUTE_SOURCE_ID)) {
    map.addSource(ROUTE_SOURCE_ID, {
      type: "geojson",
      data: {
        type: "Feature",
        geometry: {
          type: "LineString",
          coordinates: [],
        },
        properties: {},
      },
    });
  }

  if (!map.getLayer(ROUTE_LAYER_ID)) {
    map.addLayer({
      id: ROUTE_LAYER_ID,
      type: "line",
      source: ROUTE_SOURCE_ID,
      layout: {
        "line-join": "round",
        "line-cap": "round",
      },
      paint: {
        "line-color": "#0f172a",
        "line-width": 4,
        "line-opacity": 0.65,
      },
    });
  }
}

function updateRouteLayer(
  map: mapboxgl.Map,
  routeGeometry: RouteGeometry | null,
  showRouteOverlay: boolean
): void {
  const apply = () => {
    ensureRouteLayer(map);
    const source = map.getSource(ROUTE_SOURCE_ID) as mapboxgl.GeoJSONSource;

    source.setData({
      type: "Feature",
      geometry:
        showRouteOverlay && routeGeometry
          ? routeGeometry
          : {
              type: "LineString",
              coordinates: [],
            },
      properties: {},
    });
  };

  if (map.isStyleLoaded()) {
    apply();
  } else {
    map.once("load", apply);
  }
}

export default function EventMap({
  events,
  selectedEventId,
  onSelectEvent,
  routeGeometry = null,
  showRouteOverlay = true,
  standalone = false,
}: EventMapProps) {
  const containerRef = React.useRef<HTMLDivElement | null>(null);
  const mapRef = React.useRef<mapboxgl.Map | null>(null);
  const markersRef = React.useRef<MarkerEntry[]>([]);
  const token = getMapboxToken();

  React.useEffect((): void | (() => void) => {
    if (!containerRef.current || mapRef.current || !token) {
      return undefined;
    }

    mapboxgl.accessToken = token;

    const map = new mapboxgl.Map({
      container: containerRef.current,
      style: "mapbox://styles/mapbox/streets-v12",
      center: [-82.4572, 27.9506],
      zoom: 6,
    });

    mapRef.current = map;

    map.on("load", () => {
      ensureRouteLayer(map);
    });

    return () => {
      markersRef.current.forEach(({ marker }) => marker.remove());
      markersRef.current = [];
      map.remove();
      mapRef.current = null;
    };
  }, [token]);

  React.useEffect((): void => {
    const map = mapRef.current;
    if (!map) return;

    updateRouteLayer(map, routeGeometry, showRouteOverlay);
  }, [routeGeometry, showRouteOverlay]);

  React.useEffect((): void => {
    const map = mapRef.current;
    if (!map) return;

    markersRef.current.forEach(({ marker }) => marker.remove());
    markersRef.current = [];

    if (events.length === 0) return;

    const bounds = new mapboxgl.LngLatBounds();

    events.forEach((event, index) => {
      const isSelected = event.id === selectedEventId;
      const primaryCategory = event.categories[0];
      const color = getCategoryColor(primaryCategory);

      const popup = new mapboxgl.Popup({ offset: 18 }).setHTML(`
        <div style="font-family: Arial, Helvetica, sans-serif; min-width: 220px; color: #0f172a;">
          <div style="font-size: 12px; font-weight: 700; margin-bottom: 6px; color: #64748b;">Stop ${index + 1}</div>
          <div style="font-weight: 700; margin-bottom: 6px; color: #0f172a;">${escapeHtml(event.subject)}</div>
          <div style="font-size: 12px; margin-bottom: 4px; color: #334155;">${escapeHtml(
            formatDateRange(event.startIso, event.endIso)
          )}</div>
          <div style="font-size: 12px; margin-bottom: 8px; color: #334155;">${escapeHtml(event.addressText ?? "")}</div>
          <div style="font-size: 12px; color: ${color.text};">${escapeHtml(primaryCategory || "Uncategorized")}</div>
        </div>
      `);

      const element = buildMarkerElement(color.marker, isSelected, event.subject, index + 1);

      element.addEventListener("click", () => {
        onSelectEvent(event.id);
      });

      const marker = new mapboxgl.Marker({
        element,
        anchor: "center",
      })
        .setLngLat([event.longitude, event.latitude])
        .setPopup(popup)
        .addTo(map);

      markersRef.current.push({ marker, popup, eventId: event.id });
      bounds.extend([event.longitude, event.latitude]);
    });

    if (events.length === 1) {
      const only = events[0];
      map.flyTo({
        center: [only.longitude, only.latitude],
        zoom: 16,
        essential: true,
      });

      const onlyMarker = markersRef.current[0];
      if (onlyMarker && !onlyMarker.popup.isOpen()) {
        onlyMarker.marker.togglePopup();
      }
    } else if (!bounds.isEmpty()) {
      map.fitBounds(bounds, { padding: 60, maxZoom: 13 });
    }
  }, [events, selectedEventId, onSelectEvent]);

  React.useEffect((): void => {
    const map = mapRef.current;
    if (!map) return;

    let selectedCoords: [number, number] | null = null;

    markersRef.current.forEach(({ marker, popup, eventId }) => {
      const event = events.find((item) => item.id === eventId);
      const primaryCategory = event?.categories[0];
      const color = getCategoryColor(primaryCategory);
      const isSelected = eventId === selectedEventId;
      const element = marker.getElement() as HTMLDivElement;

      element.style.width = isSelected ? "32px" : "28px";
      element.style.height = isSelected ? "32px" : "28px";
      element.style.background = color.marker;
      element.style.border = isSelected ? "4px solid #111827" : "4px solid #ffffff";
      element.style.fontSize = isSelected ? "13px" : "12px";

      if (isSelected) {
        const lngLat = marker.getLngLat();
        selectedCoords = [lngLat.lng, lngLat.lat];

        if (!popup.isOpen()) {
          marker.togglePopup();
        }
      } else if (popup.isOpen()) {
        popup.remove();
      }
    });

    if (selectedCoords) {
      map.flyTo({
        center: selectedCoords,
        zoom: Math.max(map.getZoom(), 14),
        essential: true,
      });
    }
  }, [selectedEventId, events]);

  if (!token) {
    return (
      <div style={{ border: "1px solid #ccc", padding: 16, borderRadius: 8 }}>
        <h2 style={{ marginTop: 0 }}>Map</h2>
        <p>Missing Mapbox token.</p>
      </div>
    );
  }

 return (
  <div
    style={{
      border: "1px solid #ccc",
      padding: 16,
      borderRadius: 8,
      background: "#fff",
    }}
  >
    <h2 style={{ marginTop: 0, color: "#0f172a" }}>Map</h2>
    <p style={{ marginTop: 0, color: "#475569" }}>
      Mappable events in current view: {events.length}
    </p>
    <div
      ref={containerRef}
      style={{
        width: "100%",
        height: standalone ? "calc(100vh - 270px)" : 700,
        minHeight: standalone ? 820 : 700,
        borderRadius: 8,
        overflow: "hidden",
        background: "#e5e7eb",
        position: "relative",
      }}
    />
  </div>
  );
}

function formatDateRange(startIso: string, endIso: string): string {
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

  return `${dateFormatter.format(start)} ${timeFormatter.format(start)} - ${dateFormatter.format(
    end
  )} ${timeFormatter.format(end)}`;
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}