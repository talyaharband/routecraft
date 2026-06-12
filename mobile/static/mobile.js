const todayScreen = document.getElementById("todayScreen");
const activeScreen = document.getElementById("activeScreen");
const routeError = document.getElementById("routeError");
const restartDemoButton = document.getElementById("restartDemoButton");

const state = {
  payload: null,
  driver: null,
  stops: [],
  currentStopIndex: 0,
  todayMap: null,
  activeMap: null,
  todayLayer: null,
  activeLayer: null
};

document.getElementById("startRouteButton").addEventListener("click", showActiveRoute);
document.getElementById("viewStopsButton").addEventListener("click", showActiveRoute);
document.getElementById("activeBackButton").addEventListener("click", showTodayRoute);
document.getElementById("arrivedButton").addEventListener("click", markArrived);
document.getElementById("remainingStopsButton").addEventListener("click", toggleRemainingStops);
document.getElementById("recenterButton").addEventListener("click", () => renderActiveMap(true));
document.getElementById("navigateButton").addEventListener("click", openNavigation);
document.getElementById("detailsButton").addEventListener("click", toggleRemainingStops);
restartDemoButton?.addEventListener("click", restartDemo);

loadDriverRoute();

async function loadDriverRoute() {
  try {
    const params = new URLSearchParams(window.location.search);
    const driver = params.get("driver") || "דביר לוי";
    const source = params.get("source") || "comparison";
    const apiParams = new URLSearchParams({ driver, source });
    const response = await fetch(`/api/driver-route?${apiParams.toString()}`);
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || "Could not load route data.");
    state.payload = payload;
    state.driver = payload.driver;
    state.stops = routeStopsForDriver(payload.driver);
    state.currentStopIndex = firstDeliveryIndex(state.stops);
    renderTodayRoute();
    renderActiveRoute();
  } catch (error) {
    showError(error.message);
  }
}

function renderTodayRoute() {
  const driver = state.driver;
  const firstStop = currentStop();

  setText("todayDriverName", driver.name || "דביר לוי");
  setText("todayDate", new Intl.DateTimeFormat("en-US", {
    weekday: "long",
    month: "long",
    day: "numeric",
    year: "numeric"
  }).format(new Date()));
  setText("todayShiftType", `${driver.shift_type || "Route"} Shift`);
  setText("todayStops", driver.addresses || state.stops.length || 0);
  setText("todayDuration", driver.duration_label || formatDuration(driver.total_minutes));
  setText("todayDistance", Number(driver.distance_km || 0).toFixed(1));
  setText("todayFirstStop", stopDisplayLabel(firstStop));
  setText("todayEta", etaForStop(state.currentStopIndex));
  renderTodayMap();
}

function renderActiveRoute() {
  const driver = state.driver;
  const total = deliveryStops(state.stops).length || Number(driver.addresses || 0) || state.stops.length;
  const deliveryPosition = deliveryOrdinalForIndex(state.currentStopIndex);
  const remaining = Math.max(0, total - deliveryPosition);
  const percent = total ? Math.max(4, Math.min(100, (deliveryPosition / total) * 100)) : 0;
  const stop = currentStop();

  setText("activeStopCount", `Stop ${deliveryPosition} of ${total}`);
  setText("activeRemaining", `${remaining} stops remaining`);
  setText("activeStopAddress", stopDisplayLabel(stop));
  setText("activeEta", `ETA: ${etaForStop(state.currentStopIndex)}`);
  setText("activeOrderId", orderIdForStop(state.currentStopIndex));
  document.getElementById("activeProgressBar").style.width = `${percent}%`;
  renderRemainingStops();
  renderActiveMap();
}

function showActiveRoute() {
  todayScreen.classList.add("hidden");
  activeScreen.classList.remove("hidden");
  window.setTimeout(() => {
    renderActiveRoute();
    if (state.activeMap) state.activeMap.invalidateSize();
  }, 60);
}

function showTodayRoute() {
  activeScreen.classList.add("hidden");
  todayScreen.classList.remove("hidden");
  window.setTimeout(() => {
    if (state.todayMap) state.todayMap.invalidateSize();
  }, 60);
}

function restartDemo() {
  window.location.assign("/");
}

function markArrived() {
  if (!state.stops.length) return;
  const nextIndex = nextDeliveryIndex(state.currentStopIndex + 1);
  if (nextIndex === -1) {
    document.getElementById("arrivedButton").innerHTML = `<span class="material-symbols-outlined" style="font-variation-settings: 'FILL' 1;">task_alt</span> Route Complete`;
    document.getElementById("arrivedButton").disabled = true;
    return;
  }
  state.currentStopIndex = nextIndex;
  renderActiveRoute();
}

function toggleRemainingStops() {
  document.getElementById("remainingStopsList").classList.toggle("hidden");
}

function openNavigation() {
  const stop = currentStop();
  if (!stop || !Number.isFinite(Number(stop.lat)) || !Number.isFinite(Number(stop.lng))) return;
  const url = `https://www.google.com/maps/dir/?api=1&destination=${encodeURIComponent(`${stop.lat},${stop.lng}`)}`;
  window.open(url, "_blank", "noreferrer");
}

function renderTodayMap() {
  if (!window.L || !state.stops.length) return;
  if (!state.todayMap) {
    state.todayMap = L.map("todayMap", {
      zoomControl: false,
      attributionControl: false,
      dragging: true,
      scrollWheelZoom: false
    });
    addTiles(state.todayMap);
  }
  renderMapLayer(state.todayMap, "todayLayer", { compact: true });
}

function renderActiveMap(recenter = false) {
  if (!window.L || !state.stops.length) return;
  if (!state.activeMap) {
    state.activeMap = L.map("activeMap", {
      zoomControl: false,
      attributionControl: false,
      dragging: true,
      scrollWheelZoom: true
    });
    addTiles(state.activeMap);
  }
  renderMapLayer(state.activeMap, "activeLayer", { compact: false, recenter });
}

function renderMapLayer(map, layerName, options) {
  const points = state.stops.filter((stop) => Number.isFinite(Number(stop.lat)) && Number.isFinite(Number(stop.lng)));
  if (state[layerName]) map.removeLayer(state[layerName]);
  const layer = L.layerGroup().addTo(map);
  state[layerName] = layer;
  if (!points.length) {
    map.setView([31.778, 35.015], 8);
    return;
  }

  const latLngs = points.map((point) => [Number(point.lat), Number(point.lng)]);
  L.polyline(latLngs, {
    color: "#0051d5",
    weight: options.compact ? 4 : 5,
    opacity: 0.88,
    lineCap: "round",
    lineJoin: "round"
  }).addTo(layer);

  points.forEach((point) => {
    const realIndex = state.stops.indexOf(point);
    const isWarehouse = isWarehouseStop(point);
    const isDone = realIndex < state.currentStopIndex && !isWarehouse;
    const isCurrent = realIndex === state.currentStopIndex;
    const markerClass = [
      "route-stop-marker",
      isWarehouse ? "is-warehouse" : "",
      isDone ? "is-done" : "",
      isCurrent ? "is-current" : ""
    ].filter(Boolean).join(" ");
    const label = isWarehouse ? "W" : String(deliveryOrdinalForIndex(realIndex));
    L.marker([Number(point.lat), Number(point.lng)], {
      icon: L.divIcon({
        className: "",
        html: `<span class="${markerClass}">${escapeHtml(label)}</span>`,
        iconSize: [24, 24],
        iconAnchor: [12, 12]
      })
    }).bindTooltip(stopDisplayLabel(point), { direction: "top", offset: [0, -8] }).addTo(layer);
  });

  const current = currentStop();
  if (options.recenter && current && Number.isFinite(Number(current.lat)) && Number.isFinite(Number(current.lng))) {
    map.setView([Number(current.lat), Number(current.lng)], 16, { animate: true });
  } else {
    map.fitBounds(L.latLngBounds(latLngs), {
      padding: options.compact ? [18, 18] : [24, 24],
      maxZoom: options.compact ? 13 : 15
    });
  }
  window.setTimeout(() => map.invalidateSize(), 60);
}

function renderRemainingStops() {
  const list = document.getElementById("remainingStopsList");
  const remaining = state.stops
    .map((stop, index) => ({ stop, index }))
    .filter(({ stop, index }) => index >= state.currentStopIndex && !isWarehouseStop(stop));
  list.innerHTML = remaining.slice(0, 12).map(({ stop, index }) => `
    <li>
      <span class="stop-number">${deliveryOrdinalForIndex(index)}</span>
      <span class="stop-label" dir="${containsHebrew(stopDisplayLabel(stop)) ? "rtl" : "ltr"}">${escapeHtml(stopDisplayLabel(stop))}</span>
    </li>
  `).join("");
}

function addTiles(map) {
  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    maxZoom: 19
  }).addTo(map);
}

function routeStopsForDriver(driver) {
  const routePoints = (driver.route_points || []).filter((point) => point && point.label);
  if (routePoints.length) return routePoints;
  return (driver.stops || []).filter(Boolean).map((label) => ({ label }));
}

function deliveryStops(stops) {
  return stops.filter((stop) => !isWarehouseStop(stop));
}

function firstDeliveryIndex(stops) {
  const index = stops.findIndex((stop) => !isWarehouseStop(stop));
  return index >= 0 ? index : 0;
}

function nextDeliveryIndex(startIndex) {
  for (let index = startIndex; index < state.stops.length; index += 1) {
    if (!isWarehouseStop(state.stops[index])) return index;
  }
  return -1;
}

function currentStop() {
  return state.stops[state.currentStopIndex] || state.stops[0] || null;
}

function deliveryOrdinalForIndex(index) {
  let count = 0;
  for (let cursor = 0; cursor <= index && cursor < state.stops.length; cursor += 1) {
    if (!isWarehouseStop(state.stops[cursor])) count += 1;
  }
  return Math.max(1, count);
}

function isWarehouseStop(stop) {
  if (!stop) return false;
  const type = String(stop.stop_type || "");
  const label = stopDisplayLabel(stop).toUpperCase();
  return type.startsWith("warehouse") || label.includes("WAREHOUSE");
}

function etaForStop(index) {
  const start = new Date();
  start.setHours(8, 30, 0, 0);
  const minutes = index * 9 + deliveryOrdinalForIndex(index) * 4;
  start.setMinutes(start.getMinutes() + minutes);
  return start.toLocaleTimeString("en-US", { hour: "2-digit", minute: "2-digit", hour12: false });
}

function orderIdForStop(index) {
  return `#${String(10480 + deliveryOrdinalForIndex(index)).padStart(5, "0")}`;
}

function stopDisplayLabel(stop) {
  if (typeof stop === "string") return stop;
  const label = String(stop?.label || "").trim();
  const city = String(stop?.city || "").trim();
  if (label && city && !label.includes(city)) return `${label}, ${city}`;
  return label || city || "Unknown stop";
}

function formatDuration(minutes) {
  const total = Math.max(0, Math.round(Number(minutes || 0)));
  const hours = Math.floor(total / 60);
  const mins = total % 60;
  if (hours && mins) return `${hours}h ${mins}m`;
  if (hours) return `${hours}h`;
  return `${mins}m`;
}

function containsHebrew(value) {
  return /[\u0590-\u05ff]/.test(String(value || ""));
}

function setText(id, value) {
  document.getElementById(id).textContent = value;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function showError(message) {
  routeError.textContent = message;
  routeError.classList.remove("hidden");
}
