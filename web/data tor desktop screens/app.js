const screens = {
  setup: document.getElementById("dailySetupScreen"),
  processing: document.getElementById("processingScreen"),
  results: document.getElementById("routeResultsScreen")
};

const form = document.getElementById("plannerForm");
const driversInput = document.getElementById("drivers");
const capacityInput = document.getElementById("capacity");
const fileInput = document.getElementById("ordersFile");
const uploadZone = document.getElementById("uploadZone");
const fileName = document.getElementById("fileName");
const fileReadyBadge = document.getElementById("fileReadyBadge");
const formError = document.getElementById("formError");
const submitButton = document.getElementById("submitButton");
const demoResultsButton = document.getElementById("demoResultsButton");
const summarySource = document.getElementById("summarySource");
const summaryDrivers = document.getElementById("summaryDrivers");
const summaryCapacity = document.getElementById("summaryCapacity");
const summaryAddresses = document.getElementById("summaryAddresses");
const capacityForecast = document.getElementById("capacityForecast");
const capacityBar = document.getElementById("capacityBar");
const jobStatus = document.getElementById("jobStatus");
const jobLog = document.getElementById("jobLog");
const jobErrorPanel = document.getElementById("jobErrorPanel");
const jobErrorText = document.getElementById("jobErrorText");
const processingSpinner = document.getElementById("processingSpinner");
const backToSetupButton = document.getElementById("backToSetupButton");
const resultActions = document.getElementById("resultActions");
const driverAssignments = document.getElementById("driverAssignments");
const routeMapElement = document.getElementById("routeMap");
const routeMapEmpty = document.getElementById("routeMapEmpty");
const mapDriverLabel = document.getElementById("mapDriverLabel");
const dispatchAction = document.querySelector(".dispatch-action");

let pollTimer = null;
let routeMap = null;
let routeMapLayer = null;
let routeMapMarkersByStop = new Map();
let activeDrivers = [];
let selectedDriverIndex = 0;
let currentResultSource = "comparison";

const displayDriverNames = [
  "דביר לוי",
  "שמעיה סבן",
  "נהוראי מלצר",
  "יוסי אביטן",
  "איתיאל סופר",
  "יעל רבינוביץ"
];

document.getElementById("currentDate").textContent = new Date().toLocaleDateString("en-US", {
  weekday: "long",
  year: "numeric",
  month: "long",
  day: "numeric"
});

driversInput.addEventListener("input", updateSetupSummary);
capacityInput.addEventListener("input", updateSetupSummary);
fileInput.addEventListener("change", updateSelectedFile);
backToSetupButton.addEventListener("click", resetToSetup);
demoResultsButton.addEventListener("click", loadDemoResults);
if (dispatchAction) {
  dispatchAction.addEventListener("click", openMobileDispatch);
}

["dragenter", "dragover"].forEach((eventName) => {
  uploadZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    uploadZone.classList.add("is-dragging");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  uploadZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    uploadZone.classList.remove("is-dragging");
  });
});

uploadZone.addEventListener("drop", (event) => {
  const file = event.dataTransfer.files[0];
  if (!file) return;
  const transfer = new DataTransfer();
  transfer.items.add(file);
  fileInput.files = transfer.files;
  updateSelectedFile();
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  clearError();

  const validation = validateForm();
  if (!validation.ok) {
    formError.textContent = validation.message;
    return;
  }

  submitButton.disabled = true;
  submitButton.innerHTML = `<span class="material-symbols-outlined text-[20px]">hourglass_top</span> Generating...`;
  showProcessing();

  try {
    const payload = new FormData();
    payload.append("orders", fileInput.files[0]);
    payload.append("drivers", driversInput.value);
    payload.append("capacity", capacityInput.value);

    const response = await fetch("/api/jobs", { method: "POST", body: payload });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Could not start the route run.");
    pollJob(data.id);
  } catch (error) {
    showFailure(error.message);
  }
});

updateSetupSummary();

const demoMode = new URLSearchParams(window.location.search).get("demo");
if (demoMode === "results") {
  loadDemoResults();
} else {
  loadComparisonResults();
}

function updateSelectedFile() {
  const file = fileInput.files[0];
  if (!file) {
    fileName.textContent = "Choose the daily orders Excel file";
    summarySource.textContent = "No file selected";
    summaryAddresses.textContent = "Pending";
    fileReadyBadge.classList.add("hidden");
    fileReadyBadge.classList.remove("inline-flex");
    return;
  }
  fileName.textContent = file.name;
  summarySource.textContent = file.name;
  summaryAddresses.textContent = "Uploaded";
  fileReadyBadge.classList.remove("hidden");
  fileReadyBadge.classList.add("inline-flex");
  updateSetupSummary();
}

function updateSetupSummary() {
  const drivers = positiveNumber(driversInput.value);
  const capacity = positiveNumber(capacityInput.value);
  const totalCapacity = drivers * capacity;
  summaryDrivers.textContent = drivers || "-";
  summaryCapacity.textContent = capacity || "-";
  if (totalCapacity > 0) {
    capacityForecast.textContent = `Estimated fleet capacity: ${totalCapacity} packages`;
    capacityBar.style.width = `${Math.max(8, Math.min(100, totalCapacity / 8))}%`;
  } else {
    capacityForecast.textContent = "Estimated Load: waiting for fleet parameters";
    capacityBar.style.width = "8%";
  }
}

function validateForm() {
  const drivers = positiveNumber(driversInput.value);
  const capacity = positiveNumber(capacityInput.value);
  const file = fileInput.files[0];
  if (!drivers) return { ok: false, message: "Enter a positive number of available drivers." };
  if (!capacity) return { ok: false, message: "Enter a positive vehicle capacity." };
  if (!file) return { ok: false, message: "Choose the daily Excel orders file." };
  if (!/\.(xlsx|xls)$/i.test(file.name)) return { ok: false, message: "Upload must be an Excel file (.xlsx or .xls)." };
  return { ok: true };
}

async function loadDemoResults() {
  clearError();
  demoResultsButton.disabled = true;
  demoResultsButton.innerHTML = `<span class="material-symbols-outlined text-[20px]">hourglass_top</span> Loading saved results...`;
  try {
    const response = await fetch("/api/demo-result");
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || "Could not load saved real results.");
    renderResults(payload);
  } catch (error) {
    formError.textContent = error.message;
  } finally {
    demoResultsButton.disabled = false;
    demoResultsButton.innerHTML = `<span class="material-symbols-outlined text-[20px]">map</span> Open Saved Real Google Results`;
  }
}

async function loadComparisonResults() {
  clearError();
  demoResultsButton.disabled = true;
  demoResultsButton.innerHTML = `<span class="material-symbols-outlined text-[20px]">hourglass_top</span> Loading comparison results...`;
  try {
    const response = await fetch("/api/comparison-result");
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || "Could not load comparison results.");
    renderResults(payload);
  } catch (error) {
    formError.textContent = error.message;
  } finally {
    demoResultsButton.disabled = false;
    demoResultsButton.innerHTML = `<span class="material-symbols-outlined text-[20px]">map</span> Open Saved Real Google Results`;
  }
}

function pollJob(jobId) {
  clearInterval(pollTimer);
  pollTimer = setInterval(async () => {
    try {
      const response = await fetch(`/api/jobs/${jobId}`);
      const job = await response.json();
      if (!response.ok) throw new Error(job.error || "Could not read job status.");
      updateProgress(job);
      if (job.status === "completed") {
        clearInterval(pollTimer);
        restoreSubmitButton();
        renderResults(job);
      } else if (job.status === "failed") {
        clearInterval(pollTimer);
        showFailure(job.error || "The route run failed.");
      }
    } catch (error) {
      clearInterval(pollTimer);
      showFailure(error.message);
    }
  }, 1800);
}

function updateProgress(job) {
  jobStatus.textContent = titleCase(job.status || "running");
  jobLog.textContent = (job.log || []).join("\n");
  const lowerLog = jobLog.textContent.toLowerCase();
  const activeCount =
    job.status === "completed" ? 5 :
    lowerLog.includes("delivery") || lowerLog.includes("attempt") ? 5 :
    lowerLog.includes("distance") || lowerLog.includes("matrix") ? 4 :
    lowerLog.includes("geocode") ? 3 :
    lowerLog.includes("clean") || lowerLog.includes("addresses_for_geocoding") ? 2 :
    1;
  document.querySelectorAll("#progressSteps .progress-step").forEach((step, index) => {
    step.classList.toggle("is-active", index < activeCount);
  });
}

function renderResults(job) {
  const result = job.result || {};
  const summary = result.summary || {};
  const drivers = result.drivers || [];
  currentResultSource = summary.source === "demo" ? "demo" : "comparison";

  document.getElementById("kpiAssignedDeliveries").textContent = `${valueOrZero(summary.delivered_addresses)} / ${valueOrZero(summary.uploaded_orders)}`;
  document.getElementById("kpiAssignedDrivers").textContent = valueOrZero(summary.drivers);
  document.getElementById("kpiRegularShifts").textContent = valueOrZero(summary.regular_shifts);
  document.getElementById("kpiExtendedShifts").textContent = valueOrZero(summary.extended_shifts);
  renderKpiVisuals(summary);
  document.getElementById("failedAddressCount").textContent = Math.max(0, Number(summary.uploaded_orders || 0) - Number(summary.route_ready || 0));
  document.getElementById("unassignedDeliveryCount").textContent = valueOrZero(summary.leftover_orders || summary.deferred_orders);
  document.getElementById("driverCountLabel").textContent = `${drivers.length} Drivers Active`;

  renderActionRequiredButtons(result.files || {});
  renderActions(result.files || {});
  renderDrivers(drivers);
  showScreen("results");
  window.setTimeout(() => {
    selectDriver(0);
  }, 0);
}

function renderActionRequiredButtons(files) {
  setActionCardLink("failedAddressesAction", files.failed_addresses);
  setActionCardLink("unassignedOrdersAction", files.leftover_orders);
}

function setActionCardLink(id, href) {
  const link = document.getElementById(id);
  if (!link) return;
  link.classList.toggle("hidden", !href);
  if (href) {
    link.href = href;
  } else {
    link.removeAttribute("href");
  }
}

function setFillWidth(id, numerator, denominator) {
  const element = document.getElementById(id);
  if (!element) return;
  const total = Number(denominator || 0);
  const value = Number(numerator || 0);
  const percentage = total > 0 ? Math.max(0, Math.min(100, (value / total) * 100)) : 0;
  element.style.width = `${percentage}%`;
}

function renderKpiVisuals(summary) {
  const delivered = Number(summary.delivered_addresses || 0);
  const uploaded = Number(summary.uploaded_orders || 0);
  const regular = Number(summary.regular_shifts || 0);
  const extended = Number(summary.extended_shifts || 0);
  const shiftTotal = regular + extended;

  setFillWidth("kpiDeliveryFill", delivered, uploaded);
  setFillWidth("kpiRegularFill", regular, shiftTotal);
  setFillWidth("kpiExtendedFill", extended, shiftTotal);
}

function renderActions(files) {
  const links = [
    ["Open HTML Report", files.delivery_plan_html, "open_in_new"],
    ["Selected Orders", files.selected_orders, "fact_check"]
  ].filter(([, href]) => href);

  resultActions.innerHTML = links.map(([label, href, icon]) => `
    <a class="result-link" href="${escapeAttribute(href)}" target="_blank" rel="noreferrer">
      <span class="material-symbols-outlined text-[18px]">${icon}</span>
      ${escapeHtml(label)}
    </a>
  `).join("");
}

function openMobileDispatch() {
  const driverName = displayDriverNames[0] || "דביר לוי";
  const params = new URLSearchParams({
    source: currentResultSource,
    driver: driverName
  });
  const href = `/mobile?${params.toString()}`;
  if (dispatchAction) {
    dispatchAction.setAttribute("href", href);
  }
  window.location.assign(href);
}

function renderDrivers(drivers) {
  activeDrivers = drivers;
  selectedDriverIndex = 0;
  if (!drivers.length) {
    driverAssignments.innerHTML = `
      <div class="p-lg text-center text-on-surface-variant">
        <span class="material-symbols-outlined text-[32px]">route</span>
        <p class="font-body-md text-body-md mt-sm">No driver routes were returned.</p>
      </div>
    `;
    renderRouteMap(null);
    return;
  }

  driverAssignments.innerHTML = drivers.map((driver, index) => {
    const isLong = String(driver.shift_type || "").toLowerCase() !== "regular";
    const displayName = driverDisplayName(driver, index);
    const stops = routeStopsForDriver(driver);
    return `
      <article class="driver-panel ${index === 0 ? "is-selected" : ""}" data-driver-index="${index}" tabindex="0" role="button" aria-label="Show route for ${escapeAttribute(displayName)}">
        <div class="p-md border-b border-outline-variant/20">
          <div class="flex justify-between items-start gap-md mb-sm">
            <div class="driver-card-heading min-w-0">
              <div class="min-w-0">
                <h4 class="driver-card-name font-body-md text-on-surface truncate" dir="${containsHebrew(displayName) ? "rtl" : "ltr"}">${escapeHtml(displayName)}</h4>
                <div class="flex flex-wrap gap-2 mt-0.5">
                  <span class="driver-badge ${isLong ? "is-long" : ""}">${escapeHtml(driver.shift_type || "Unknown Shift")}</span>
                  <span class="driver-badge">${escapeHtml(driver.addresses || 0)} stops</span>
                  <span class="driver-badge">${escapeHtml(driver.clusters || 0)} clusters</span>
                </div>
              </div>
            </div>
            <span class="material-symbols-outlined text-secondary">${index === 0 ? "expand_less" : "expand_more"}</span>
          </div>
          <div class="flex gap-lg text-on-surface font-data-mono text-xs">
            <div class="flex flex-col"><span class="driver-stat-label">Stops</span>${escapeHtml(driver.addresses || 0)}</div>
            <div class="flex flex-col"><span class="driver-stat-label">Duration</span>${escapeHtml(driver.total_hours || "0")}h</div>
            <div class="flex flex-col"><span class="driver-stat-label">Minutes</span>${Math.round(Number(driver.total_minutes || 0))}</div>
          </div>
          <p class="font-body-sm text-body-sm text-on-surface-variant mt-sm">${escapeHtml((driver.cities || []).join(" / ") || "No city data")}</p>
        </div>
        <div class="p-md bg-surface-bright">
          <ol class="relative border-l-2 border-outline-variant/30 ml-2 space-y-3 driver-route-list">
            ${stops.map((stop, stopIndex) => renderStop(stop, stopIndex, index)).join("")}
          </ol>
        </div>
      </article>
    `;
  }).join("");

  driverAssignments.querySelectorAll(".driver-panel").forEach((panel) => {
    panel.addEventListener("click", () => selectDriver(Number(panel.dataset.driverIndex)));
    panel.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        selectDriver(Number(panel.dataset.driverIndex));
      }
    });
  });
  driverAssignments.querySelectorAll(".route-stop-item").forEach((item) => {
    item.addEventListener("click", (event) => {
      event.stopPropagation();
      focusRouteStop(Number(item.dataset.driverIndex), Number(item.dataset.stopIndex));
    });
    item.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        event.stopPropagation();
        focusRouteStop(Number(item.dataset.driverIndex), Number(item.dataset.stopIndex));
      }
    });
  });
}

function driverDisplayName(driver, index) {
  return displayDriverNames[index] || driver?.name || `Driver ${index + 1}`;
}

function renderStop(stop, index, driverIndex) {
  const label = stopDisplayLabel(stop);
  const kind = stopKindLabel(stop, index);
  const isWarehouse = stop && typeof stop === "object" && String(stop.stop_type || "").startsWith("warehouse");
  const rtl = containsHebrew(label);
  return `
    <li class="route-stop-item relative pl-6" data-driver-index="${driverIndex}" data-stop-index="${index}" tabindex="0" role="button" aria-label="Zoom map to ${escapeAttribute(label)}">
      <div class="absolute w-2.5 h-2.5 ${isWarehouse ? "bg-[#b45309]" : index === 0 ? "bg-secondary" : "bg-outline-variant"} rounded-full -left-[6px] top-1.5 border-2 border-white"></div>
      <div class="route-stop" dir="${rtl ? "rtl" : "ltr"}">
        <span class="block text-[10px] font-label-caps text-on-surface-variant">${escapeHtml(kind)}</span>
        <p class="text-sm font-medium text-on-surface">${escapeHtml(label)}</p>
      </div>
    </li>
  `;
}

function routeStopsForDriver(driver) {
  const routePoints = (driver.route_points || []).filter((point) => point && point.label);
  if (routePoints.length) return routePoints;
  return (driver.stops || []).filter(Boolean);
}

function stopKindLabel(stop, index) {
  if (!stop || typeof stop !== "object") return `Stop ${index + 1}`;
  if (stop.stop_type === "warehouse_start") return "Warehouse start";
  if (stop.stop_type === "warehouse_reload") return "Warehouse reload";
  return `Stop ${index + 1}`;
}

function mapStopKindLabel(stop, index) {
  if (!stop || typeof stop !== "object") return `Stop ${index + 1}`;
  if (stop.stop_type === "warehouse_start") return "Warehouse start";
  if (stop.stop_type === "warehouse_reload") return "Warehouse reload";
  const numbers = stop.stop_numbers || [index + 1];
  const label = formatStopNumbers(numbers);
  return numbers.length > 1 ? `Stops ${label}` : `Stop ${label}`;
}

function formatStopNumbers(numbers) {
  if (!numbers.length) return "";
  const ranges = [];
  let start = numbers[0];
  let previous = numbers[0];
  for (const number of numbers.slice(1)) {
    if (number === previous + 1) {
      previous = number;
      continue;
    }
    ranges.push(start === previous ? String(start) : `${start}-${previous}`);
    start = number;
    previous = number;
  }
  ranges.push(start === previous ? String(start) : `${start}-${previous}`);
  return ranges.join(", ");
}

function mapPointKey(point) {
  const lat = Number(point.lat).toFixed(6);
  const lng = Number(point.lng).toFixed(6);
  const stopType = String(point.stop_type || "delivery");
  return `${lat}|${lng}|${stopType}|${stopDisplayLabel(point)}`;
}

function groupedMapPoints(points) {
  const grouped = [];
  points.forEach((point, index) => {
    const key = mapPointKey(point);
    const previous = grouped[grouped.length - 1];
    if (previous && previous.map_key === key) {
      previous.stop_numbers.push(index + 1);
      previous.group_count += 1;
      return;
    }
    grouped.push({
      ...point,
      map_key: key,
      stop_numbers: [index + 1],
      group_count: 1
    });
  });
  return grouped;
}

function stopDisplayLabel(stop) {
  if (typeof stop === "string") return stop;
  const label = String(stop?.label || "").trim();
  const city = String(stop?.city || "").trim();
  if (label && city && !label.includes(city)) return `${label}, ${city}`;
  return label || city || "Unknown stop";
}

function showProcessing() {
  jobErrorPanel.classList.add("hidden");
  backToSetupButton.classList.add("hidden");
  processingSpinner.classList.remove("hidden");
  jobStatus.textContent = "Queued";
  jobLog.textContent = "";
  document.querySelectorAll("#progressSteps .progress-step").forEach((step, index) => {
    step.classList.toggle("is-active", index === 0);
  });
  showScreen("processing");
}

function selectDriver(index) {
  if (!activeDrivers.length) return;
  selectedDriverIndex = Math.max(0, Math.min(index, activeDrivers.length - 1));
  driverAssignments.querySelectorAll(".driver-panel").forEach((panel) => {
    panel.classList.toggle("is-selected", Number(panel.dataset.driverIndex) === selectedDriverIndex);
  });
  const driver = activeDrivers[selectedDriverIndex];
  mapDriverLabel.textContent = driverDisplayName(driver, selectedDriverIndex);
  renderRouteMap(driver);
}

function focusRouteStop(driverIndex, stopIndex) {
  if (!activeDrivers.length) return;
  if (driverIndex !== selectedDriverIndex) {
    selectDriver(driverIndex);
  }
  const stopNumber = stopIndex + 1;
  const marker = routeMapMarkersByStop.get(stopNumber);
  const driver = activeDrivers[selectedDriverIndex];
  const stop = routeStopsForDriver(driver)[stopIndex];
  driverAssignments.querySelectorAll(".route-stop-item").forEach((item) => {
    item.classList.toggle(
      "is-focused",
      Number(item.dataset.driverIndex) === selectedDriverIndex && Number(item.dataset.stopIndex) === stopIndex
    );
  });
  if (!marker || !routeMap) return;
  const latLng = marker.getLatLng();
  routeMap.setView(latLng, Math.max(routeMap.getZoom(), 17), { animate: true });
  marker.openTooltip();
  mapDriverLabel.textContent = `${driverDisplayName(driver, selectedDriverIndex)} - ${stopKindLabel(stop, stopIndex)}`;
}

function renderRouteMap(driver) {
  if (!routeMapElement || !window.L) return;
  const sourcePoints = (driver?.route_points?.length ? driver.route_points : driver?.delivery_points) || [];
  const points = sourcePoints.filter((point) => Number.isFinite(Number(point.lat)) && Number.isFinite(Number(point.lng)));
  const mapPoints = groupedMapPoints(points);
  if (!routeMap) {
    routeMap = L.map(routeMapElement, {
      zoomControl: true,
      scrollWheelZoom: true
    });
    L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
      maxZoom: 19,
      attribution: "&copy; OpenStreetMap contributors"
    }).addTo(routeMap);
  }

  if (routeMapLayer) {
    routeMap.removeLayer(routeMapLayer);
  }
  routeMapLayer = L.layerGroup().addTo(routeMap);
  routeMapMarkersByStop = new Map();
  routeMapEmpty.classList.toggle("hidden", points.length > 0);
  routeMapEmpty.classList.toggle("flex", points.length === 0);

  if (!points.length) {
    routeMap.setView([31.778, 35.015], 8);
    return;
  }

  const latLngs = mapPoints.map((point) => [Number(point.lat), Number(point.lng)]);
  L.polyline(latLngs, {
    color: "#0051d5",
    weight: 5,
    opacity: 0.88,
    lineCap: "round",
    lineJoin: "round",
    className: "animated-route-line"
  }).addTo(routeMapLayer);

  mapPoints.forEach((point, index) => {
    const isWarehouse = String(point.stop_type || "").startsWith("warehouse");
    const marker = L.circleMarker([Number(point.lat), Number(point.lng)], {
      radius: isWarehouse ? 8 : index === 0 ? 7 : 5,
      color: isWarehouse ? "#b45309" : index === 0 ? "#0b1c30" : "#0051d5",
      weight: 2,
      fillColor: isWarehouse ? "#fef3c7" : index === 0 ? "#0b1c30" : "#ffffff",
      fillOpacity: 1
    });
    const tooltipDetail = point.group_count > 1 ? ` (${point.group_count} deliveries)` : "";
    marker.bindTooltip(`${mapStopKindLabel(point, index)}: ${stopDisplayLabel(point)}${tooltipDetail}`, {
      direction: "top",
      offset: [0, -6]
    });
    marker.addTo(routeMapLayer);
    point.stop_numbers.forEach((stopNumber) => routeMapMarkersByStop.set(stopNumber, marker));
    L.marker([Number(point.lat), Number(point.lng)], {
      interactive: false,
      icon: L.divIcon({
        className: isWarehouse ? "route-stop-warehouse-label" : "route-stop-number-label",
        html: isWarehouse
          ? `<span class="material-symbols-outlined">warehouse</span>`
          : `<span>${formatStopNumbers(point.stop_numbers)}</span>`,
        iconSize: isWarehouse ? [30, 30] : [46, 24],
        iconAnchor: isWarehouse ? [15, 15] : [23, 12]
      })
    }).addTo(routeMapLayer);
  });

  routeMap.fitBounds(L.latLngBounds(latLngs), {
    padding: [18, 18],
    maxZoom: mapPoints.length <= 3 ? 16 : 15
  });
  window.setTimeout(() => routeMap.invalidateSize(), 50);
}

function showFailure(message) {
  restoreSubmitButton();
  showScreen("processing");
  jobStatus.textContent = "Failed";
  jobErrorText.textContent = message || "The route run failed.";
  jobErrorPanel.classList.remove("hidden");
  backToSetupButton.classList.remove("hidden");
  processingSpinner.classList.add("hidden");
}

function resetToSetup() {
  clearInterval(pollTimer);
  restoreSubmitButton();
  clearError();
  showScreen("setup");
}

function showScreen(name) {
  Object.entries(screens).forEach(([key, screen]) => {
    screen.classList.toggle("hidden", key !== name);
  });
}

function restoreSubmitButton() {
  submitButton.disabled = false;
  submitButton.innerHTML = `<span class="material-symbols-outlined text-[20px]">route</span> Generate Daily Routes`;
}

function clearError() {
  formError.textContent = "";
}

function positiveNumber(value) {
  const number = Number(value);
  return Number.isFinite(number) && number > 0 ? number : 0;
}

function valueOrZero(value) {
  return value ?? 0;
}

function titleCase(value) {
  const text = String(value || "");
  return text.slice(0, 1).toUpperCase() + text.slice(1);
}

function containsHebrew(value) {
  return /[\u0590-\u05ff]/.test(String(value || ""));
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function escapeAttribute(value) {
  return escapeHtml(value).replaceAll("`", "&#096;");
}
