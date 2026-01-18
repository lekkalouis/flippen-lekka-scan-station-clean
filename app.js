(() => {
  "use strict";

  const CONFIG = {
    COST_ALERT_THRESHOLD: 250.0,
    BOX_DIM: { dim1: 40, dim2: 40, dim3: 30, massKg: 5 },
    ORIGIN: {
      origpers: "Flippen Lekka Holdings (Pty) Ltd",
      origperadd1: "7 Papawer Street",
      origperadd2: "Blomtuin, Bellville",
      origperadd3: "Cape Town, Western Cape",
      origperadd4: "ZA",
      origperpcode: "7530",
      origtown: "Cape Town",
      origplace: 4663,
      origpercontact: "Louis",
      origperphone: "0730451885",
      origpercell: "0730451885",
      notifyorigpers: 1,
      origperemail: "admin@flippenlekkaspices.co.za",
      notes: "Louis 0730451885 / Michael 0783556277"
    },
    TRACKING_COMPANY: "SWE / ParcelPerfect",
    PP_ENDPOINT: "/pp",
    SHOPIFY: { PROXY_BASE: "/shopify" }
  };

  const $ = (id) => document.getElementById(id);
  const scanInput = $("scanInput");
  const uiOrderNo = $("uiOrderNo");
  const uiParcelCount = $("uiParcelCount");
  const uiCountdown = $("uiCountdown");
  const shipToCard = $("shipToCard");
  const parcelList = $("parcelList");
  const bookingSummary = $("bookingSummary");
  const bookingSummaryCard = $("bookingSummaryCard");
  const statusChip = $("statusChip");
  const stickerPreview = $("stickerPreview");
  const debugLog = $("debugLog");
  const quoteBox = $("quoteBox");
  const printMount = $("printMount");
  const addrSearch = $("addrSearch");
  const addrResults = $("addrResults");
  const placeCodeInput = $("placeCode");
  const serviceSelect = $("serviceOverride");

  const dispatchBoard = $("dispatchBoard");
  const dispatchStamp = $("dispatchStamp");

  const navScan = $("navScan");
  const navOps = $("navOps");
  const viewScan = $("viewScan");
  const viewOps = $("viewOps");
  const actionFlash = $("actionFlash");
  const emergencyStopBtn = $("emergencyStop");
  const modeToggleBtn = $("modeToggle");
  const focusScannerBtn = $("focusScanner");
  const toggleAddressBtn = $("toggleAddress");
  const toggleServiceBtn = $("toggleService");
  const copyShipToBtn = $("copyShipTo");
  const serviceDisplay = $("serviceDisplay");
  const autoModeStatus = $("autoModeStatus");

  const btnBookNow = $("btnBookNow");

  const MAX_ORDER_AGE_HOURS = 180;

  let activeOrderNo = null;
  let orderDetails = null;
  let parcelsForOrder = new Set();
  let armedForBooking = false;

  let placeCodeOverride = null;
  let serviceOverride = "RDF";
  let addressBook = [];
  let bookedOrders = new Set();
  let autoBookEnabled = true;
  let autoBookEndAtMs = null;
  let countdownInterval = null;
  let dispatchNotes = {};

  const dbgOn = new URLSearchParams(location.search).has("debug");
  if (dbgOn && debugLog) debugLog.style.display = "block";

  const statusExplain = (msg, tone = "info") => {
    if (statusChip) statusChip.textContent = msg;
    if (!actionFlash) return;
    actionFlash.textContent = msg;

    actionFlash.classList.remove(
      "actionFlash--info",
      "actionFlash--ok",
      "actionFlash--warn",
      "actionFlash--err"
    );

    const cls =
      tone === "ok"
        ? "actionFlash--ok"
        : tone === "warn"
        ? "actionFlash--warn"
        : tone === "err"
        ? "actionFlash--err"
        : "actionFlash--info";

    actionFlash.classList.add(cls);

    actionFlash.style.opacity = "1";
    clearTimeout(actionFlash._fadeTimer);
    actionFlash._fadeTimer = setTimeout(() => {
      actionFlash.style.opacity = "0.4";
    }, 2000);
  const STORAGE_KEYS = {
    plans: "fl_pack_plans_v2",
    recent: "fl_recent_fulfilled_v2"
  };

  const STEPS = [
    { key: "packed", label: "Packed" },
    { key: "booked", label: "Booked" },
    { key: "printed", label: "Printed" },
    { key: "fulfilled", label: "Fulfilled" },
    { key: "notified", label: "Notified" },
    { key: "archived", label: "Archived" }
  ];

  function updateCountdownDisplay() {
    if (!uiCountdown) return;
    if (!autoBookEndAtMs) {
      uiCountdown.textContent = "--";
      return;
    }
    const remainingMs = autoBookEndAtMs - Date.now();
    if (remainingMs <= 0) {
      uiCountdown.textContent = "0";
      autoBookEndAtMs = null;
      return;
    }
    uiCountdown.textContent = String(Math.ceil(remainingMs / 1000));
  }

  function startCountdownTicker() {
    if (countdownInterval) return;
    countdownInterval = setInterval(updateCountdownDisplay, 250);
  }

  function stopCountdownTicker() {
    if (!countdownInterval) return;
    clearInterval(countdownInterval);
    countdownInterval = null;
  }

  function renderCountdown() {
    updateCountdownDisplay();
  }

  function money(v) {
    return v == null || isNaN(v) ? "-" : `R${Number(v).toFixed(2)}`;
  }

  function setBookingSummary(text) {
    if (bookingSummary) bookingSummary.textContent = text;
    if (bookingSummaryCard) bookingSummaryCard.textContent = text;
  }

  function updateModeToggleButton() {
    if (!modeToggleBtn) return;
    modeToggleBtn.textContent = `MODE: ${autoBookEnabled ? "AUTO" : "MANUAL"}`;
    modeToggleBtn.classList.toggle("btnPrimary", autoBookEnabled);
  }

  function updateAutoModeUI() {
    if (autoModeStatus) autoModeStatus.textContent = autoBookEnabled ? "ON" : "OFF";
    updateModeToggleButton();
  }

  function updateServiceDisplay() {
    if (serviceDisplay) serviceDisplay.textContent = serviceOverride || "Auto";
  }
function isAutoBookOrder(details) {
  return hasParcelCountTag(details);
}
  const openOrdersEl = document.getElementById("openOrders");
  const refreshBtn = document.getElementById("refreshBtn");
  const clearStorageBtn = document.getElementById("clearStorageBtn");
  const recentTable = document.getElementById("recentFulfilledTable");
  const recentBody = document.getElementById("recentFulfilledBody");
  const recentEmpty = document.getElementById("recentFulfilledEmpty");

  let orders = [];
  let packPlans = {};
  let recentFulfilled = [];

  const loadStorage = () => {
    try {
      const rawPlans = localStorage.getItem(STORAGE_KEYS.plans);
      packPlans = rawPlans ? JSON.parse(rawPlans) : {};
    } catch {
      packPlans = {};
    }

    try {
      localStorage.setItem("fl_booked_orders_v1", JSON.stringify([...bookedOrders]));
    } catch {}
  }

  function loadDispatchNotes() {
    try {
      const raw = localStorage.getItem("fl_dispatch_notes_v1");
      dispatchNotes = raw ? JSON.parse(raw) : {};
    } catch {
      dispatchNotes = {};
    }
  }

  function saveDispatchNotes() {
    try {
      localStorage.setItem("fl_dispatch_notes_v1", JSON.stringify(dispatchNotes));
    } catch {}
  }

  function getDispatchKey(order) {
    return String(order?.id || order?.name || "");
  }

  function getDispatchEntry(order) {
    const key = getDispatchKey(order);
    if (!key) return { status: "not_started", boxes: "", notes: "" };
    if (!dispatchNotes[key]) {
      dispatchNotes[key] = { status: "not_started", boxes: "", notes: "" };
    }
    return dispatchNotes[key];
  }

  function updateDispatchBadge(card, status) {
    const badge = card.querySelector(".dispatchBadge");
    if (!badge) return;
    const label =
      status === "packing"
        ? "Packing"
        : status === "ready"
        ? "Ready"
        : status === "done"
        ? "Completed"
        : "Not started";
    badge.textContent = label;
    badge.classList.remove("status--packing", "status--ready", "status--done");
    if (status === "packing") badge.classList.add("status--packing");
    if (status === "ready") badge.classList.add("status--ready");
    if (status === "done") badge.classList.add("status--done");
  }

  function markBooked(orderNo) {
    bookedOrders.add(String(orderNo));
    saveBookedOrders();
  }

  function isBooked(orderNo) {
    return bookedOrders.has(String(orderNo));
  }

  function base64PdfToUrl(base64) {
    if (!base64) return null;
    const cleaned = base64.replace(/\s/g, "");
    const byteChars = atob(cleaned);
    const len = byteChars.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) bytes[i] = byteChars.charCodeAt(i);
    const blob = new Blob([bytes], { type: "application/pdf" });
    return URL.createObjectURL(blob);
  }

  const CODE128_PATTERNS = [
    "11011001100","11001101100","11001100110","10010011000","10010001100","10001001100","10011001000","10011000100","10001100100","11001001000","11001000100","11000100100","10110011100","10011011100","10011001110","10111001100","10011101100","10011100110","11001110010","11001011100","11001001110","11011100100","11001110100","11101101110","11101001100","11100101100","11100100110","11101100100","11100110100","11100110010","11011011000","11011000110","11000110110","10100011000","10001011000","10001000110","10110001000","10001101000","10001100010","11010001000","11000101000","11000100010","10110111000","10110001110","10001101110","10111011000","10111000110","10001110110","11101110110","11010001110","11000101110","11011101000","11011100010","11011101110","11101011000","11101000110","11100010110","11101101000","11101100010","11100011010","11101111010","11001000010","11110001010","10100110000","10100001100","10010110000","10010000110","10000101100","10000100110","10110010000","10110000100","10011010000","10011000010","10000110100","10000110010","11000010010","11001010000","11110111010","11000010100","10001111010","10100111100","10010111100","10010011110","10111100100","10011110100","10011110010","11110100100","11110010100","11110010010","11011011110","11011110110","11110110110","10101111000","10100011110","10001011110","10111101000","10111100010","11110101000","11110100010","10111011110","10111101110","11101011110","11110101110","11010000100","11010010000","11010011100","11000111010","11010111000","1100011101011"
  ];

  function code128BToValues(str) {
    const vals = [];
    for (let i = 0; i < str.length; i++) {
      const c = str.charCodeAt(i);
      if (c < 32 || c > 126) throw new Error("Unsupported char " + str[i]);
      vals.push(c - 32);
    }
    return vals;
  }

  function code128Encode(str) {
    const vals = code128BToValues(str);
    const full = [104, ...vals];
    let checksum = 104;
    for (let i = 0; i < vals.length; i++) checksum += vals[i] * (i + 1);
    checksum %= 103;
    full.push(checksum, 106);
    return full;
  }

  function code128Svg(str, h) {
    const vals = code128Encode(str);
    const height = h || 80;
    const moduleWidth = 2;
    const quietModules = 10;

    const barModules = vals.reduce((sum, code) => sum + CODE128_PATTERNS[code].length, 0);
    const totalModules = barModules + quietModules * 2;
    const totalWidth = totalModules * moduleWidth;

    let x = quietModules * moduleWidth;
    const rects = [];

    for (const c of vals) {
      const pattern = CODE128_PATTERNS[c];
      for (let i = 0; i < pattern.length; i++) {
        if (pattern[i] === "1") rects.push(`<rect x="${x}" y="0" width="${moduleWidth}" height="${height}" />`);
        x += moduleWidth;
      }
      const rawRecent = localStorage.getItem(STORAGE_KEYS.recent);
      recentFulfilled = rawRecent ? JSON.parse(rawRecent) : [];
    } catch {
      recentFulfilled = [];
    }
  };

    return `
<svg xmlns="http://www.w3.org/2000/svg"
     width="${totalWidth}"
     height="${height}"
     viewBox="0 0 ${totalWidth} ${height}"
     shape-rendering="crispEdges"
     style="background:#fff;display:block;max-width:100%;">
  <rect width="100%" height="100%" fill="#fff" />
  ${rects.join("")}
</svg>`;
  }
let autoBookTimer = null;

function cancelAutoBookTimer() {
  if (autoBookTimer) clearTimeout(autoBookTimer);
  autoBookTimer = null;
  autoBookEndAtMs = null;
  updateCountdownDisplay();
}

function scheduleIdleAutoBook() {
  cancelAutoBookTimer();

  if (!autoBookEnabled) return;
  // Only for untagged orders
  if (!activeOrderNo || !orderDetails) return;
  if (isBooked(activeOrderNo)) return;
  if (hasParcelCountTag(orderDetails)) return;

  // Need at least 1 scan
  if (parcelsForOrder.size <= 0) return;

  autoBookEndAtMs = Date.now() + CONFIG.BOOKING_IDLE_MS;
  updateCountdownDisplay();
  startCountdownTicker();

  autoBookTimer = setTimeout(async () => {
    autoBookTimer = null;
    autoBookEndAtMs = null;
    updateCountdownDisplay();

    // Still valid?
    if (!activeOrderNo || !orderDetails) return;
    if (isBooked(activeOrderNo)) return;
    if (armedForBooking) return;
    if (hasParcelCountTag(orderDetails)) return;
    if (parcelsForOrder.size <= 0) return;

    // Use scanned count as the parcel count (avoid prompt)
    orderDetails.manualParcelCount = parcelsForOrder.size;

    renderSessionUI();
    updateBookNowButton();

    statusExplain(`No tag. Auto-booking ${parcelsForOrder.size} parcels...`, "ok");
    await doBookingNow(); // will pass scanned==expected because expected becomes manualParcelCount
  }, CONFIG.BOOKING_IDLE_MS);
}

  const ADDR_FALLBACK = [
    {
      label: "Louis (Office)",
      name: "Louis Cabano",
      phone: "0730451885",
      email: "admin@flippenlekkaspices.co.za",
      address1: "37 Papawer Street",
      address2: "Oakdale",
      city: "Cape Town",
      province: "Western Cape",
      postal: "7530",
      placeCode: 4658
    },
    {
      label: "Michael (Warehouse)",
      name: "Michael Collison",
      phone: "0783556277",
      email: "admin@flippenlekkaspices.co.za",
      address1: "7 Papawer Street",
      address2: "Blomtuin, Bellville",
      city: "Cape Town",
      province: "Western Cape",
      postal: "7530",
      placeCode: 4001
  const savePlans = () => {
    localStorage.setItem(STORAGE_KEYS.plans, JSON.stringify(packPlans));
  };

  const saveRecent = () => {
    localStorage.setItem(STORAGE_KEYS.recent, JSON.stringify(recentFulfilled));
  };

  const formatDate = (value) => {
    if (!value) return "-";
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return "-";
    return date.toLocaleString();
  };

  const formatShipTo = (order) => {
    const parts = [order.shipping_address1, order.shipping_address2, order.shipping_city, order.shipping_postal]
      .map((p) => String(p || "").trim())
      .filter(Boolean);
    return parts.join(", ") || "-";
  };

  const getPlan = (order) => {
    const key = String(order.id);
    if (!packPlans[key]) {
      packPlans[key] = {
        orderId: order.id,
        orderName: order.name,
        orderGid: order.order_gid,
        boxes: [],
        milestones: {},
        log: [],
        expanded: false,
        bookingData: null,
        fulfillmentIds: []
      };
    }
    return packPlans[key];
  };

  const logPlan = (plan, status, message) => {
    plan.log = plan.log || [];
    plan.log.unshift({ status, message, at: new Date().toISOString() });
    plan.log = plan.log.slice(0, 50);
  };

  const setMilestone = (plan, key, status, message) => {
    const existing = plan.milestones[key];
    if (existing?.status === status && existing?.message === message) return;

    plan.milestones[key] = {
      status,
      message,
      at: new Date().toISOString()
    };
    logPlan(plan, status === "ok" ? "ok" : status === "err" ? "err" : "info", `${key}: ${message}`);
    savePlans();
  };

  const normalizeLineItems = (order) =>
    (order.line_items || []).map((li) => ({
      id: li.gid || li.id,
      title: li.title,
      quantity: Number(li.quantity || 0),
      fulfillable_quantity: Number(li.fulfillable_quantity ?? li.quantity ?? 0)
    }));

  const getAllocatedQty = (plan, lineItemId) =>
    plan.boxes.reduce((sum, box) => sum + Number(box.items?.[lineItemId] || 0), 0);

  const getRemainingQty = (order, plan, lineItem) => {
    const max = Number(lineItem.fulfillable_quantity || 0);
    const allocated = getAllocatedQty(plan, lineItem.id);
    return Math.max(0, max - allocated);
  };

  const getRemainingTotals = (order, plan) => {
    const items = normalizeLineItems(order);
    const remaining = items.reduce((sum, li) => sum + getRemainingQty(order, plan, li), 0);
    return { remaining, total: items.reduce((sum, li) => sum + li.fulfillable_quantity, 0) };
  };

  const updatePlanState = (order, plan) => {
    const items = normalizeLineItems(order);
    const remaining = items.map((li) => getRemainingQty(order, plan, li));
    const anyAllocated = plan.boxes.some((box) => Object.values(box.items || {}).some((qty) => qty > 0));
    const packed = remaining.every((qty) => qty === 0) && items.length > 0;

    if (packed) {
      if (plan.milestones.packed?.status !== "ok") {
        setMilestone(plan, "packed", "ok", "All items allocated to boxes.");
      }
    } else if (anyAllocated) {
      setMilestone(plan, "packed", "pending", "Packing in progress.");
    } else {
      setMilestone(plan, "packed", "pending", "No items allocated yet.");
    }

    savePlans();
  };

  const renderRecent = () => {
    if (!recentTable || !recentBody || !recentEmpty) return;
    if (!recentFulfilled.length) {
      recentTable.hidden = true;
      recentEmpty.hidden = false;
      return;
    }

    recentTable.hidden = false;
    recentEmpty.hidden = true;
    recentBody.innerHTML = recentFulfilled
      .map(
        (item) => `
        <tr>
          <td>${formatDate(item.date)}</td>
          <td>${item.orderName}</td>
          <td>${item.waybill}</td>
          <td>${item.customer}</td>
          <td>${item.shipTo}</td>
        </tr>
      `
      )
      .join("");
    addrResults.innerHTML =
      rows ||
      `<div class="addrItem" style="opacity:.7;cursor:default">No matches. Type city, postal, or name…</div>`;
  }

  async function initAddressSearch() {
    await loadAddressBook();
    renderAddrResults("");
    addrResults?.addEventListener("click", (e) => {
      const item = e.target.closest(".addrItem");
      if (!item) return;
      const idx = Number(item.dataset.idx);
      const entry = addressBook[idx];
      if (!entry) return;
      setDestinationFromEntry(entry);
    });
    addrSearch?.addEventListener("input", () => renderAddrResults(addrSearch.value));
    placeCodeInput?.addEventListener("input", () => {
      const v = (placeCodeInput.value || "").trim();
      placeCodeOverride = v ? Number(v) || null : null;
    });
    serviceSelect?.addEventListener("change", () => {
      serviceOverride = serviceSelect.value || "AUTO";
      updateServiceDisplay();
    });
  }

  function hasParcelCountTag(details) {
    return !!(details && typeof details.parcelCountFromTag === "number" && details.parcelCountFromTag > 0);
  }

  function shouldShowBookNow(details) {
    return !!activeOrderNo && !!details && !hasParcelCountTag(details) && !isBooked(activeOrderNo);
  }

  function updateBookNowButton() {
    if (!btnBookNow) return;
    const show = shouldShowBookNow(orderDetails);
    btnBookNow.hidden = !show;
    btnBookNow.disabled = !show;

    if (!show) return;
    const scanned = parcelsForOrder.size;
    btnBookNow.textContent = scanned > 0 ? `BOOK NOW (${scanned} scanned)` : "BOOK NOW";
  }

  function getExpectedParcelCount(details) {
    const fromTag =
      details && typeof details.parcelCountFromTag === "number" && details.parcelCountFromTag > 0
        ? details.parcelCountFromTag
        : null;
    const manual =
      details && typeof details.manualParcelCount === "number" && details.manualParcelCount > 0
        ? details.manualParcelCount
        : null;
    return fromTag || manual || null;
  }

  function getParcelIndexesForCurrentOrder(details) {
    const expected = getExpectedParcelCount(details);
    if (expected) return Array.from({ length: expected }, (_, i) => i + 1);
    if (parcelsForOrder.size > 0) return Array.from(parcelsForOrder).sort((a, b) => a - b);
    return [];
  }

  function renderSessionUI() {
    if (uiOrderNo) uiOrderNo.textContent = activeOrderNo || "--";

    const expected = getExpectedParcelCount(orderDetails || {});
    const idxs = getParcelIndexesForCurrentOrder(orderDetails || {});
    if (uiParcelCount) uiParcelCount.textContent = String(idxs.length);

    const tagInfo =
      orderDetails && typeof orderDetails.parcelCountFromTag === "number" && orderDetails.parcelCountFromTag > 0
        ? ` (tag: parcel_count_${orderDetails.parcelCountFromTag})`
        : "";

    const manualInfo =
      orderDetails && typeof orderDetails.manualParcelCount === "number" && orderDetails.manualParcelCount > 0
        ? ` (manual: ${orderDetails.manualParcelCount})`
        : "";

    if (parcelList) {
      parcelList.textContent = idxs.length
        ? `Parcels: ${idxs.join(", ")}${tagInfo}${manualInfo}`
        : "No parcels (scan parcel labels).";
    }
  };

  const render = () => {
    if (!openOrdersEl) return;
    const cards = orders
      .filter((order) => packPlans[String(order.id)]?.milestones?.archived?.status !== "ok")
      .map((order) => renderOrderCard(order))
      .join("");

    openOrdersEl.innerHTML = cards || `<div class="emptyState">No open orders available.</div>`;
    renderRecent();
  };

    updateAutoModeUI();
    updateServiceDisplay();
    updateBookNowButton();
  }

  function setDestinationFromEntry(entry) {
    orderDetails = {
      ...(orderDetails || {}),
      name: entry.name,
      phone: entry.phone,
      email: entry.email,
      address1: entry.address1,
      address2: entry.address2 || "",
      city: entry.city,
      province: entry.province,
      postal: entry.postal
    };
    placeCodeOverride = entry.placeCode || null;
    if (placeCodeInput) placeCodeInput.value = placeCodeOverride ? String(placeCodeOverride) : "";
    renderSessionUI();
  }

  function renderLabelHTML(waybillNo, service, cost, destDetails, parcelIdx, parcelCount) {
    const parcelStr = String(parcelIdx).padStart(3, "0");
    const codeParcel = `${waybillNo}0${parcelStr}`;
    const svgTopParcel = code128Svg(codeParcel, 70);
    const svgBig = code128Svg(codeParcel, 90);

    const fromHTML = `Flippen Lekka Holdings (Pty) Ltd
7 Papawer Street, Blomtuin, Bellville
Cape Town, Western Cape, 7530
Louis 0730451885 / Michael 0783556277
admin@flippenlekkaspices.co.za`.replace(/\n/g, "<br>");

    const toHTML = `${destDetails.name}<br>${destDetails.address1}${
      destDetails.address2 ? `<br>${destDetails.address2}` : ""
    }<br>${destDetails.city}, ${destDetails.province} ${destDetails.postal}<br>Tel: ${destDetails.phone || ""}`;
  const renderOrderCard = (order) => {
    const plan = getPlan(order);
    updatePlanState(order, plan);

    const items = normalizeLineItems(order);
    const { remaining, total } = getRemainingTotals(order, plan);
    const packed = remaining === 0 && total > 0;

    const statusChip = packed
      ? `<span class="chip ok">Packed/Ready to book</span>`
      : remaining < total
      ? `<span class="chip warn">In progress</span>`
      : `<span class="chip">Open</span>`;

    const fulfillChip = order.fulfillment_status
      ? `<span class="chip">Fulfillment: ${order.fulfillment_status}</span>`
      : "";

    const progress = STEPS.map((step) => {
      const status = plan.milestones[step.key]?.status;
      const cls = status === "ok" ? "done" : status === "pending" ? "active" : "";
      return `<span class="progressStep ${cls}">${step.label}</span>`;
    }).join("");

    const milestones = STEPS.map((step) => {
      const data = plan.milestones[step.key] || {};
      const status = data.status || "pending";
      const icon = status === "ok" ? "ok" : status === "err" ? "err" : "";
      return `
        <div class="milestone">
          <span>${step.label}</span>
          <span class="state ${icon}">${status === "ok" ? "✔" : status === "err" ? "✖" : "…"}</span>
        </div>
      `;
    }).join("");

    const logItems = (plan.log || [])
      .slice(0, 6)
      .map(
        (entry) => `
        <div class="logEntry ${entry.status}">
          <div>${entry.message}</div>
        </div>
      `
      )
      .join("") || `<div class="logEntry info">No milestones yet.</div>`;

    const boxesHtml = plan.expanded
      ? plan.boxes
          .map((box, idx) => renderBoxEditor(order, plan, box, idx))
          .join("")
      : "";

    const remainingRows = items
      .map(
        (li) => `
        <div class="lineItem">
          <span>${li.title}</span>
          <span>${getRemainingQty(order, plan, li)} / ${li.fulfillable_quantity}</span>
        </div>
      `
      )
      .join("");

    const bookingDisabled = !packed || plan.milestones.booked?.status === "ok" || plan.milestones.booked?.status === "pending";

    const retryPrint = plan.milestones.printed?.status === "err";
    const retryFulfill = plan.milestones.fulfilled?.status === "err";
    const retryNotify = plan.milestones.notified?.status === "err";

    return `
      <article class="orderCard" data-order-id="${order.id}">
        <div class="cardHeader">
          <div>
            <h3>${order.name}</h3>
            <div class="muted">${formatDate(order.created_at)} • ${order.customer_name || "Unknown"}</div>
            <div class="muted">${formatShipTo(order)}</div>
          </div>
          <div class="chipRow">
            ${statusChip}
            ${fulfillChip}
          </div>
        </div>

        <div class="lineItems">
          ${items
            .map((li) => `<div class="lineItem"><span>${li.title}</span><span>× ${li.quantity}</span></div>`)
            .join("")}
        </div>

        <div class="progress">${progress}</div>

        <div class="cardActions">
          <button class="ghost" data-action="toggle" data-order-id="${order.id}">
            ${plan.expanded ? "Hide pack plan" : "Fulfill / Pack"}
          </button>
          <button class="ghost" data-action="add-box" data-order-id="${order.id}">Add box</button>
          <button class="success" data-action="book" data-order-id="${order.id}" ${bookingDisabled ? "disabled" : ""}>
            BOOK NOW
          </button>
          ${retryPrint ? `<button class="warn" data-action="retry-print" data-order-id="${order.id}">Retry Print</button>` : ""}
          ${retryFulfill ? `<button class="warn" data-action="retry-fulfill" data-order-id="${order.id}">Retry Fulfill</button>` : ""}
          ${retryNotify ? `<button class="warn" data-action="retry-notify" data-order-id="${order.id}">Retry Notify</button>` : ""}
        </div>

        ${plan.expanded ? `
          <div class="editor">
            <h4>Packing editor</h4>
            <div class="remainingList">${remainingRows}</div>
            ${boxesHtml || `<div class="muted" style="margin-top:0.5rem;">No boxes created yet.</div>`}
            <div class="muted" style="margin-top:0.5rem;">Remaining to allocate: ${remaining} / ${total}</div>
          </div>
        ` : ""}

        <div class="milestones">${milestones}</div>
        <div class="logList">${logItems}</div>
      </article>
    `;
  };

  const renderBoxEditor = (order, plan, box, idx) => {
    const items = normalizeLineItems(order);
    const rows = items
      .map((li) => {
        const allocated = Number(box.items?.[li.id] || 0);
        const remaining = getRemainingQty(order, plan, li) + allocated;
        const disableDec = allocated <= 0;
        const disableInc = remaining <= 0;
        return `
          <div class="qtyRow">
            <span>${li.title}</span>
            <div class="qtyControls">
              <button data-action="dec" data-order-id="${order.id}" data-box-index="${idx}" data-line-id="${li.id}" ${disableDec ? "disabled" : ""}>-</button>
              <span>${allocated}</span>
              <button data-action="inc" data-order-id="${order.id}" data-box-index="${idx}" data-line-id="${li.id}" ${disableInc ? "disabled" : ""}>+</button>
            </div>
          </div>
        `;
      })
      .join("");

    return `
      <div class="boxCard">
        <div class="boxHeader">
          <strong>Box ${idx + 1}</strong>
          <button class="ghost" data-action="remove-box" data-order-id="${order.id}" data-box-index="${idx}">Remove</button>
        </div>
        ${rows}
      </div>
    `;
  };

  const fetchOpenOrders = async () => {
    const res = await fetch(`${CONFIG.SHOPIFY.PROXY_BASE}/orders/open`);
    if (!res.ok) throw new Error(`Open orders fetch failed: ${res.status}`);
    const data = await res.json();
    orders = data.orders || [];
  };

  const updateAllocation = (orderId, boxIndex, lineId, delta) => {
    const order = orders.find((o) => String(o.id) === String(orderId));
    if (!order) return;
    const plan = getPlan(order);
    const line = normalizeLineItems(order).find((li) => String(li.id) === String(lineId));
    if (!line) return;

    const current = Number(plan.boxes[boxIndex]?.items?.[line.id] || 0);
    const remaining = getRemainingQty(order, plan, line) + current;
    const next = Math.max(0, Math.min(current + delta, remaining));

    if (!plan.boxes[boxIndex].items) plan.boxes[boxIndex].items = {};
    plan.boxes[boxIndex].items[line.id] = next;

    savePlans();
    updatePlanState(order, plan);
    render();
  };

  const addBox = (orderId) => {
    const order = orders.find((o) => String(o.id) === String(orderId));
    if (!order) return;
    const plan = getPlan(order);
    plan.boxes.push({ boxIndex: plan.boxes.length + 1, items: {} });
    plan.expanded = true;
    savePlans();
    updatePlanState(order, plan);
    render();
  };

  const removeBox = (orderId, boxIndex) => {
    const order = orders.find((o) => String(o.id) === String(orderId));
    if (!order) return;
    const plan = getPlan(order);
    plan.boxes.splice(boxIndex, 1);
    savePlans();
    updatePlanState(order, plan);
    render();
  };

  const toggleExpand = (orderId) => {
    const order = orders.find((o) => String(o.id) === String(orderId));
    if (!order) return;
    const plan = getPlan(order);
    plan.expanded = !plan.expanded;
    savePlans();
    render();
  };

  const ppCall = async (payload) => {
    const res = await fetch(CONFIG.PP_ENDPOINT, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });

    const text = await res.text();
    let data;
    try {
      data = JSON.parse(text);
    } catch {
      data = { raw: text };
    }

    return { status: res.status, statusText: res.statusText, data };
  };

  const extractQuoteFromV28 = (shape) => {
    const obj = shape || {};
    if (obj.quoteno) return { quoteno: obj.quoteno, rates: obj.rates || [] };
    const res = Array.isArray(obj.results) && obj.results[0] ? obj.results[0] : null;
    const quoteno = (res && res.quoteno) || null;
    const rates = res && Array.isArray(res.rates) ? res.rates : [];
    return { quoteno, rates };
  };

  const resolvePlaceCode = async (order) => {
    const city = String(order.shipping_city || "").trim();
    const suburb = String(order.shipping_address2 || "").trim();
    const search = [suburb, city].filter(Boolean).join(" ");
    if (!search) return null;

    const res = await fetch(`/pp/place?q=${encodeURIComponent(search)}`);
    if (!res.ok) return null;
    const data = await res.json();
    const list = Array.isArray(data.results) ? data.results : [];
    if (!list.length) return null;
    const best = list.find((p) => String(p.ring) === "0") || list[0];
    if (!best || best.place == null) return null;
    return Number(best.place) || best.place;
  };

  const buildParcelPerfectPayload = async (order, parcelCount) => {
    const destplace = await resolvePlaceCode(order);
    const details = {
      ...CONFIG.ORIGIN,
      destpers: order.shipping_name || order.customer_name,
      destperadd1: order.shipping_address1,
      destperadd2: order.shipping_address2 || "",
      destperadd3: order.shipping_city,
      destperadd4: order.shipping_province,
      destperpcode: order.shipping_postal,
      desttown: order.shipping_city,
      destplace,
      destpercontact: order.shipping_name || order.customer_name,
      destperphone: order.shipping_phone || "",
      notifydestpers: 1,
      destpercell: order.shipping_phone || "",
      destperemail: order.email || "",
      reference: `Order ${order.name}`
    };

    const contents = Array.from({ length: parcelCount }, (_, i) => ({
      item: i + 1,
      pieces: 1,
      dim1: CONFIG.BOX_DIM.dim1,
      dim2: CONFIG.BOX_DIM.dim2,
      dim3: CONFIG.BOX_DIM.dim3,
      actmass: CONFIG.BOX_DIM.massKg
    }));

    return { details, contents };
  };

  const bookParcelPerfect = async (order, plan) => {
    const parcelCount = plan.boxes.length;
    if (!parcelCount) throw new Error("No boxes to book.");

    if (missing.length) {
      statusExplain("Quote failed", "err");
      setBookingSummary(
        `Cannot request quote — missing: ${missing.join(", ")}\n\nShip To:\n${JSON.stringify(orderDetails, null, 2)}`
      );
      armedForBooking = false;
      return;
    }

    const quoteRes = await ppCall({ method: "requestQuote", classVal: "quote", params: payload });
    if (!quoteRes || quoteRes.status !== 200) {
      statusExplain("Quote failed", "err");
      setBookingSummary(
        `Quote error (HTTP ${quoteRes?.status}): ${quoteRes?.statusText}\n\n${JSON.stringify(quoteRes?.data, null, 2)}`
      );
      if (quoteBox) quoteBox.textContent = "No quote — check place code / proxy / token.";
      armedForBooking = false;
      return;
    }

    const { quoteno, rates } = extractQuoteFromV28(quoteRes.data || {});
    if (!quoteno) {
      statusExplain("Quote failed", "err");
      setBookingSummary(`No quote number.\n${JSON.stringify(quoteRes.data, null, 2)}`);
      armedForBooking = false;
      return;
    }

    const pickedService = pickService(rates);
    const chosenRate = rates?.find((r) => r.service === pickedService) || rates?.[0] || null;
    const quoteCost = chosenRate ? Number(chosenRate.total ?? chosenRate.subtotal ?? chosenRate.charge ?? 0) : null;

    if (serviceDisplay) serviceDisplay.textContent = pickedService;

    if (quoteBox && rates?.length) {
      const fmt = (v) => (isNaN(Number(v)) ? "-" : `R${Number(v).toFixed(2)}`);
      const lines = rates
        .map((r) => `${r.service}: ${fmt(r.total ?? r.subtotal ?? r.charge)} ${r.name ? `(${r.name})` : ""}`)
        .join("\n");
      quoteBox.textContent = `Selected: ${pickedService} • Est: ${fmt(quoteCost)}${quoteCost > CONFIG.COST_ALERT_THRESHOLD ? "  ⚠ high" : ""}\nOptions:\n${lines}`;
    }
    const payload = await buildParcelPerfectPayload(order, parcelCount);
    if (!payload.details.destplace) {
      throw new Error("Missing destination place code. Update address or place code.");
    }

    const quoteRes = await ppCall({ method: "requestQuote", classVal: "quote", params: payload });
    if (quoteRes.status !== 200) {
      throw new Error(`Quote failed: ${quoteRes.status} ${quoteRes.statusText}`);
    }

    const { quoteno, rates } = extractQuoteFromV28(quoteRes.data || {});
    if (!quoteno) throw new Error("Quote missing quoteno.");

    const pickedService = (rates || []).find((rate) => rate.service === "RDF")?.service || rates?.[0]?.service || "RDF";

    await ppCall({
      method: "updateService",
      classVal: "quote",
      params: { quoteno, service: pickedService, reference: String(order.name) }
    });

    const collRes = await ppCall({
      method: "quoteToCollection",
      classVal: "collection",
      params: { quoteno, starttime: "12:00", endtime: "15:00", printLabels: 1, printWaybill: 0 }
    });

    if (!collRes || collRes.status !== 200) {
      statusExplain("Booking failed", "err");
      setBookingSummary(
        `Booking error: HTTP ${collRes?.status} ${collRes?.statusText}\n${JSON.stringify(collRes?.data, null, 2)}`
      );
      armedForBooking = false;
      return;
    if (collRes.status !== 200) {
      throw new Error(`Booking failed: ${collRes.status} ${collRes.statusText}`);
    }

    const cr = collRes.data || {};
    const maybe = cr.results?.[0] || cr;
    const waybill = String(maybe.waybill || maybe.waybillno || maybe.waybillNo || maybe.trackingNo || "");

    return {
      waybill,
      service: pickedService,
      labelsBase64: maybe.labelsBase64 || maybe.labelBase64 || maybe.labels_pdf || null,
      waybillBase64: maybe.waybillBase64 || maybe.waybillPdfBase64 || maybe.waybill_pdf || null,
      raw: cr
    };
  };

  const printLabels = async (booking, orderName) => {
    if (!booking?.labelsBase64) throw new Error("Labels not available to print.");
    const labelsResp = await fetch("/printnode/print", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ pdfBase64: booking.labelsBase64, title: `Labels ${orderName}` })
    });
    if (!labelsResp.ok) throw new Error(`PrintNode labels failed: ${labelsResp.status}`);

    if (booking.waybillBase64) {
      const wbResp = await fetch("/printnode/print", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ pdfBase64: booking.waybillBase64, title: `Waybill ${orderName}` })
      });
      if (!wbResp.ok) throw new Error(`PrintNode waybill failed: ${wbResp.status}`);
    }
  };

  const fetchFulfillmentMap = async (orderGid) => {
    const query = `query FulfillmentOrders($orderId: ID!) {
      order(id: $orderId) {
        fulfillmentOrders(first: 50) {
          edges {
            node {
              id
              status
              lineItems(first: 100) {
                edges {
                  node {
                    id
                    remainingQuantity
                    lineItem { id }
                  }
                }
              }
            }
          }
        }
      }
    }`;

    const res = await fetch(`${CONFIG.SHOPIFY.PROXY_BASE}/graphql`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ query, variables: { orderId: orderGid } })
    });

    if (!res.ok) throw new Error(`Fulfillment order fetch failed: ${res.status}`);
    const data = await res.json();
    const edges = data?.data?.order?.fulfillmentOrders?.edges || [];
    const map = {};
    for (const edge of edges) {
      const fo = edge.node;
      for (const liEdge of fo.lineItems.edges || []) {
        const node = liEdge.node;
        map[node.lineItem.id] = {
          fulfillmentOrderId: fo.id,
          fulfillmentOrderLineItemId: node.id,
          remainingQuantity: node.remainingQuantity
        };
      }
    }
    return map;
  };

    if (statusChip) statusChip.textContent = "Booked";
    setBookingSummary(`WAYBILL: ${waybillNo}
Service: ${pickedService}
Parcels: ${expected}
Estimated Cost: ${money(quoteCost)}
  const createFulfillment = async (order, plan) => {
    const map = await fetchFulfillmentMap(order.order_gid);
    const shipments = plan.boxes.map((box, idx) => ({
      boxIndex: idx + 1,
      lineItems: Object.entries(box.items || {})
        .filter(([, qty]) => qty > 0)
        .map(([lineItemId, quantity]) => ({
          fulfillmentOrderLineItemId: map[lineItemId]?.fulfillmentOrderLineItemId,
          quantity
        }))
        .filter((li) => li.fulfillmentOrderLineItemId),
      tracking: {
        number: plan.bookingData?.waybill || "",
        url: "",
        company: CONFIG.TRACKING_COMPANY
      }
    }));

    const res = await fetch(`${CONFIG.SHOPIFY.PROXY_BASE}/fulfillment/create`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        orderGid: order.order_gid,
        shipments,
        notifyCustomer: true
      })
    });

Raw:
${JSON.stringify(cr, null, 2)}`);

    await fulfillOnShopify(orderDetails, waybillNo);

    markBooked(activeOrderNo);
    resetSession();
  }

function resetSession() {
  cancelAutoBookTimer();

  activeOrderNo = null;
  orderDetails = null;
  parcelsForOrder = new Set();
  armedForBooking = false;

  placeCodeOverride = null;
  if (placeCodeInput) placeCodeInput.value = "";

  renderSessionUI();
  renderCountdown();
  updateBookNowButton();
  updateAutoModeUI();
  updateServiceDisplay();
}


  function parseScan(code) {
    if (!code || code.length < 9) return null;
    const orderNo = code.slice(0, code.length - 3);
    const seq = parseInt(code.slice(-3), 10);
    if (Number.isNaN(seq)) return null;
    return { orderNo, parcelSeq: seq };
  }

async function startOrder(orderNo) {
  cancelAutoBookTimer();

  activeOrderNo = orderNo;
  parcelsForOrder = new Set();
  armedForBooking = false;

  placeCodeOverride = null;
  if (placeCodeInput) placeCodeInput.value = "";

  orderDetails = await fetchShopifyOrder(activeOrderNo);

  if (orderDetails && orderDetails.placeCode != null) {
    placeCodeOverride = Number(orderDetails.placeCode) || orderDetails.placeCode;
    if (placeCodeInput) placeCodeInput.value = String(placeCodeOverride);
  }

  appendDebug("Started new order " + activeOrderNo);
  renderSessionUI();
  renderCountdown();
  updateBookNowButton();
}

async function handleScan(code) {
  const parsed = parseScan(code);
  if (!parsed) {
    appendDebug("Bad scan: " + code);
    statusExplain("Bad scan", "warn");
    return;
  }

  if (isBooked(parsed.orderNo)) {
    statusExplain(`Order ${parsed.orderNo} already booked — blocked.`, "warn");
    return;
  }

  if (!activeOrderNo) {
    await startOrder(parsed.orderNo);
} else if (parsed.orderNo !== activeOrderNo) {
  cancelAutoBookTimer(); // ADD THIS
  statusExplain(`Different order scanned (${parsed.orderNo}). Press CLEAR to reset.`, "warn");
  return;
}


  parcelsForOrder.add(parsed.parcelSeq);
  armedForBooking = false;

  const expected = getExpectedParcelCount(orderDetails);
    const data = await res.json();
    if (!res.ok) {
      const message = data?.message || "Shopify fulfillment failed";
      throw new Error(message);
    }

    return data;
  };

  const handleBookingFlow = async (orderId, mode) => {
    const order = orders.find((o) => String(o.id) === String(orderId));
    if (!order) return;
    const plan = getPlan(order);

    try {
      if (!plan.boxes.length) throw new Error("Add at least one box before booking.");

      if (plan.milestones.booked?.status === "ok" && mode !== "print" && mode !== "fulfill" && mode !== "notify") {
        return;
      }

      if (mode === "print") {
        if (!plan.bookingData) throw new Error("Missing booking data to reprint.");
        setMilestone(plan, "printed", "pending", "Retrying print.");
        await printLabels(plan.bookingData, order.name);
        setMilestone(plan, "printed", "ok", "Labels printed via PrintNode.");
        render();
        return;
      }

      if (mode === "fulfill") {
        setMilestone(plan, "fulfilled", "pending", "Retrying fulfillment.");
        const result = await createFulfillment(order, plan);
        plan.fulfillmentIds = (result.results || []).map((r) => r.fulfillmentId).filter(Boolean);
        const notifyErrors = (result.results || []).flatMap((r) => r.notifyErrors || []);
        const userErrors = (result.results || []).flatMap((r) => r.userErrors || []);

        if (userErrors.length) {
          throw new Error(userErrors.map((e) => e.message).join(" | "));
        }

        setMilestone(plan, "fulfilled", "ok", "Shopify fulfillment created.");
        if (notifyErrors.length) {
          setMilestone(plan, "notified", "err", notifyErrors.map((e) => e.message).join(" | "));
        } else {
          setMilestone(plan, "notified", "ok", "Customer notified.");
        }
        if (!notifyErrors.length) archiveOrder(order, plan);
        render();
        return;
      }

      if (mode === "notify") {
        if (!plan.fulfillmentIds.length) throw new Error("No fulfillment IDs available for notify.");
        setMilestone(plan, "notified", "pending", "Retrying notification.");
        await notifyFulfillments(plan.fulfillmentIds);
        setMilestone(plan, "notified", "ok", "Customer notified.");
        archiveOrder(order, plan);
        render();
        return;
      }

      setMilestone(plan, "booked", "pending", "Booking ParcelPerfect waybill.");
      const booking = await bookParcelPerfect(order, plan);
      if (!booking.waybill) throw new Error("Booking failed to return a waybill number.");
      plan.bookingData = booking;
      plan.boxes = plan.boxes.map((box) => ({
        ...box,
        trackingNumber: booking.waybill,
        trackingCompany: CONFIG.TRACKING_COMPANY
      }));
      setMilestone(plan, "booked", "ok", `Waybill ${booking.waybill} booked.`);

      setMilestone(plan, "printed", "pending", "Printing labels.");
      await printLabels(booking, order.name);
      setMilestone(plan, "printed", "ok", "Labels printed via PrintNode.");

      setMilestone(plan, "fulfilled", "pending", "Creating Shopify fulfillment.");
      const fulfillmentRes = await createFulfillment(order, plan);
      plan.fulfillmentIds = (fulfillmentRes.results || []).map((r) => r.fulfillmentId).filter(Boolean);
      const notifyErrors = (fulfillmentRes.results || []).flatMap((r) => r.notifyErrors || []);
      const userErrors = (fulfillmentRes.results || []).flatMap((r) => r.userErrors || []);

      if (userErrors.length) {
        setMilestone(plan, "fulfilled", "err", userErrors.map((e) => e.message).join(" | "));
        render();
        return;
      }

      return normalized;
    } catch (e) {
      appendDebug("Shopify fetch failed: " + String(e));
      return {
        raw: null,
        name: "Unknown",
        phone: "",
        email: "",
        address1: "",
        address2: "",
        city: "",
        province: "",
        postal: "",
        suburb: "",
        line_items: [],
        totalWeightKg: CONFIG.BOX_DIM.massKg,
        placeCode: null,
        placeLabel: null,
        parcelCountFromTag: null,
        manualParcelCount: null
      };
    }
  }

  function normalizeTags(tags) {
    if (!tags) return [];
    if (Array.isArray(tags)) return tags.map((t) => String(t).toLowerCase());
    return String(tags)
      .split(",")
      .map((t) => t.trim().toLowerCase())
      .filter(Boolean);
  }

  function getLaneForOrder(order) {
    const tags = normalizeTags(order?.tags);
    const shippingLines = (order?.shipping_lines || []).map((l) => String(l || "").toLowerCase());
    const title = String(order?.name || "").toLowerCase();

    const isB2B = tags.some((t) => ["b2b", "wholesale", "trade"].includes(t));
    if (isB2B) return "b2b";

    const isPickup = tags.some((t) => ["local_pickup", "pickup", "collection"].includes(t)) ||
      shippingLines.some((l) => l.includes("pickup") || l.includes("collect"));
    if (isPickup) return "pickup";

    const isDelivery = tags.some((t) => ["local_delivery", "delivery"].includes(t)) ||
      shippingLines.some((l) => l.includes("local delivery") || l.includes("delivery"));
    if (isDelivery) return "delivery";

    if (title.includes("pickup")) return "pickup";
    if (title.includes("delivery")) return "delivery";

    return "online";
  }

  function renderDispatchBoard(orders) {
    if (!dispatchBoard) return;

    const now = Date.now();
    const maxAgeMs = MAX_ORDER_AGE_HOURS * 60 * 60 * 1000;

    const filtered = (orders || []).filter((o) => {
      const fs = (o.fulfillment_status || "").toLowerCase();
      if (fs && fs !== "unfulfilled" && fs !== "in_progress") return false;
      if (!o.created_at) return true;
      const createdMs = new Date(o.created_at).getTime();
      if (!Number.isFinite(createdMs)) return true;
      return now - createdMs <= maxAgeMs;
    });

    filtered.sort((a, b) => new Date(b.created_at || 0).getTime() - new Date(a.created_at || 0).getTime());
    const list = filtered.slice(0, 40);
      setMilestone(plan, "fulfilled", "ok", "Shopify fulfillment created.");

      if (notifyErrors.length) {
        setMilestone(plan, "notified", "err", notifyErrors.map((e) => e.message).join(" | "));
      } else {
        setMilestone(plan, "notified", "ok", "Customer notified.");
        archiveOrder(order, plan);
      }
    } catch (err) {
      const message = String(err?.message || err);
      if (mode === "print") {
        setMilestone(plan, "printed", "err", message);
      } else if (mode === "fulfill") {
        setMilestone(plan, "fulfilled", "err", message);
      } else if (mode === "notify") {
        setMilestone(plan, "notified", "err", message);
      } else {
        setMilestone(plan, "booked", "err", message);
      }
    } finally {
      savePlans();
      render();
    }
  };

    const cols = [
      { id: "b2b", label: "B2B Shipping", className: "lane--b2b" },
      { id: "delivery", label: "Local Delivery", className: "lane--delivery" },
      { id: "pickup", label: "Local Pickup", className: "lane--pickup" },
      { id: "online", label: "Online Orders", className: "lane--online" }
    ];
    const lanes = { b2b: [], delivery: [], pickup: [], online: [] };
    list.forEach((o) => {
      const lane = getLaneForOrder(o);
      lanes[lane]?.push(o);
    });

    const cardHTML = (o) => {
      const entry = getDispatchEntry(o);
      const title = o.customer_name || o.name || `Order ${o.id}`;
      const city = o.shipping_city || "";
      const postal = o.shipping_postal || "";
      const created = o.created_at ? new Date(o.created_at).toLocaleTimeString() : "";
      const lines = (o.line_items || []).slice(0, 6).map((li) => `• ${li.quantity} × ${li.title}`).join("<br>");
      const addr1 = o.shipping_address1 || "";
      const addr2 = o.shipping_address2 || "";
      const addrHtml = `${addr1}${addr2 ? "<br>" + addr2 : ""}<br>${city} ${postal}`;
      const statusLabel =
        entry.status === "packing"
          ? "Packing"
          : entry.status === "ready"
          ? "Ready"
          : entry.status === "done"
          ? "Completed"
          : "Not started";
      const statusClass =
        entry.status === "packing"
          ? "status--packing"
          : entry.status === "ready"
          ? "status--ready"
          : entry.status === "done"
          ? "status--done"
          : "";

      return `
        <div class="dispatchCard" data-order-id="${getDispatchKey(o)}">
          <div class="dispatchCardTitle"><span>${title}</span><span class="dispatchBadge ${statusClass}">${statusLabel}</span></div>
          <div class="dispatchCardMeta">#${(o.name || "").replace("#", "")} · ${city} · ${created}</div>
          <div class="dispatchCardAddress">${addrHtml}</div>
          <div class="dispatchCardLines">${lines}</div>
          <div class="dispatchCardActions">
            <label class="panelSub">Status</label>
            <select class="dispatchSelect" data-field="status">
              <option value="not_started" ${entry.status === "not_started" ? "selected" : ""}>Not started</option>
              <option value="packing" ${entry.status === "packing" ? "selected" : ""}>Packing</option>
              <option value="ready" ${entry.status === "ready" ? "selected" : ""}>Ready</option>
              <option value="done" ${entry.status === "done" ? "selected" : ""}>Completed</option>
            </select>
            <label class="panelSub">Boxes / cartons</label>
            <input class="dispatchInput" data-field="boxes" placeholder="e.g. Box A, Box B" value="${entry.boxes || ""}" />
            <label class="panelSub">Packing notes</label>
            <textarea class="dispatchTextarea" data-field="notes" placeholder="Special instructions, missing items, etc.">${entry.notes || ""}</textarea>
          </div>
        </div>`;
    };

    dispatchBoard.innerHTML = cols
      .map((col, idx) => {
        const cards = lanes[col.id].map(cardHTML).join("") || `<div class="dispatchBoardEmptyCol">No jobs.</div>`;
        return `
          <div class="dispatchCol ${col.className}">
            <div class="dispatchColHeader">${col.label}<span class="laneCount">${lanes[col.id].length}</span></div>
            <div class="dispatchColBody">${cards}</div>
          </div>`;
      })
      .join("");
  }
  const notifyFulfillments = async (fulfillmentIds) => {
    const mutation = `mutation FulfillmentNotify($id: ID!) {
      fulfillmentNotify(fulfillmentId: $id, notifyCustomer: true) {
        fulfillment { id status }
        userErrors { field message }
      }
    }`;

    for (const id of fulfillmentIds) {
      const res = await fetch(`${CONFIG.SHOPIFY.PROXY_BASE}/graphql`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ query: mutation, variables: { id } })
      });
      const data = await res.json();
      const errors = data?.data?.fulfillmentNotify?.userErrors || [];
      if (errors.length) throw new Error(errors.map((e) => e.message).join(" | "));
    }
  };

  const archiveOrder = (order, plan) => {
    setMilestone(plan, "archived", "ok", "Moved to Recently Fulfilled.");
    recentFulfilled.unshift({
      date: new Date().toISOString(),
      orderName: order.name,
      waybill: plan.bookingData?.waybill || "-",
      customer: order.customer_name || "-",
      shipTo: formatShipTo(order)
    });
    recentFulfilled = recentFulfilled.slice(0, 50);
    saveRecent();
  };

  openOrdersEl?.addEventListener("click", (event) => {
    const button = event.target.closest("button");
    if (!button) return;

    const action = button.dataset.action;
    const orderId = button.dataset.orderId;
    const boxIndex = Number(button.dataset.boxIndex || 0);
    const lineId = button.dataset.lineId;

    if (action === "toggle") return toggleExpand(orderId);
    if (action === "add-box") return addBox(orderId);
    if (action === "remove-box") return removeBox(orderId, boxIndex);
    if (action === "inc") return updateAllocation(orderId, boxIndex, lineId, 1);
    if (action === "dec") return updateAllocation(orderId, boxIndex, lineId, -1);
    if (action === "book") return handleBookingFlow(orderId);
    if (action === "retry-print") return handleBookingFlow(orderId, "print");
    if (action === "retry-fulfill") return handleBookingFlow(orderId, "fulfill");
    if (action === "retry-notify") return handleBookingFlow(orderId, "notify");
  });

  refreshBtn?.addEventListener("click", async () => {
    try {
      await fetchOpenOrders();
    } catch (err) {
      console.error(err);
    } finally {
      render();
    }
  });

  modeToggleBtn?.addEventListener("click", () => {
    autoBookEnabled = !autoBookEnabled;
    if (!autoBookEnabled) cancelAutoBookTimer();
    updateAutoModeUI();
    statusExplain(autoBookEnabled ? "Auto-book enabled." : "Manual mode enabled.", "info");
  });

  focusScannerBtn?.addEventListener("click", () => {
    scanInput?.focus();
  });

  toggleAddressBtn?.addEventListener("click", () => {
    if (!addrResults) return;
    const show = addrResults.closest("#addrBox")?.hasAttribute("hidden");
    const box = addrResults.closest("#addrBox");
    if (!box) return;
    if (show) {
      box.removeAttribute("hidden");
      addrSearch?.focus();
    } else {
      box.setAttribute("hidden", "");
    }
  });

  toggleServiceBtn?.addEventListener("click", () => {
    const box = serviceSelect?.closest("#svcBox");
    if (!box) return;
    if (box.hasAttribute("hidden")) {
      box.removeAttribute("hidden");
      serviceSelect?.focus();
    } else {
      box.setAttribute("hidden", "");
    }
  });

  copyShipToBtn?.addEventListener("click", async () => {
    if (!shipToCard) return;
    const text = shipToCard.textContent?.trim();
    if (!text) {
      statusExplain("No ship-to data to copy.", "warn");
      return;
    }
    try {
      await navigator.clipboard.writeText(text);
      statusExplain("Ship-to copied to clipboard.", "ok");
    } catch {
      statusExplain("Clipboard blocked. Select text manually.", "warn");
    }
  });

  const persistDispatchField = (target) => {
    const card = target.closest(".dispatchCard");
    if (!card) return;
    const orderId = card.dataset.orderId;
    if (!orderId) return;
    const field = target.dataset.field;
    if (!field) return;
    if (!dispatchNotes[orderId]) dispatchNotes[orderId] = { status: "not_started", boxes: "", notes: "" };
    dispatchNotes[orderId][field] = target.value;
    if (field === "status") updateDispatchBadge(card, target.value);
    saveDispatchNotes();
  };

  dispatchBoard?.addEventListener("input", (e) => {
    const target = e.target;
    if (target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement) {
      if (target.dataset.field) persistDispatchField(target);
    }
  });

  dispatchBoard?.addEventListener("change", (e) => {
    const target = e.target;
    if (target instanceof HTMLSelectElement) {
      if (target.dataset.field) persistDispatchField(target);
    }
  });

btnBookNow?.addEventListener("click", async () => {
  cancelAutoBookTimer();

  if (!activeOrderNo || !orderDetails) {
    statusExplain("Scan an order first.", "warn");
    return;
  }
  if (isBooked(activeOrderNo)) {
    statusExplain(`Order ${activeOrderNo} already booked — blocked.`, "warn");
    return;
  }
  if (hasParcelCountTag(orderDetails)) {
    statusExplain("This order has a parcel_count tag — it auto-books on first scan.", "warn");
    return;
  }

  // Use scanned count as default to avoid prompt if you want:
  if (!getExpectedParcelCount(orderDetails) && parcelsForOrder.size > 0) {
    orderDetails.manualParcelCount = parcelsForOrder.size;
  }

  await doBookingNow();
});


  emergencyStopBtn?.addEventListener("click", () => {
    statusExplain("EMERGENCY STOP – session cleared", "err");
    resetSession();
    if (stickerPreview) stickerPreview.innerHTML = "";
    if (printMount) printMount.innerHTML = "";
    if (quoteBox) quoteBox.textContent = "No quote yet.";
    setBookingSummary("");
    if (scanInput) scanInput.value = "";
    if (dbgOn && debugLog) debugLog.textContent = "";
    switchMainView("scan");
  });

  navScan?.addEventListener("click", () => switchMainView("scan"));
  navOps?.addEventListener("click", () => switchMainView("ops"));

  loadBookedOrders();
  loadDispatchNotes();
  renderSessionUI();
  renderCountdown();
  updateAutoModeUI();
  updateServiceDisplay();
  startCountdownTicker();
  initAddressSearch();
  refreshDispatchData();
  setInterval(refreshDispatchData, 30000);
  switchMainView("scan");

  if (location.protocol === "file:") {
    alert("Open via http://localhost/... (not file://). Run a local server.");
  }

  window.__fl = window.__fl || {};
  window.__fl.bookNow = doBookingNow;
  clearStorageBtn?.addEventListener("click", () => {
    localStorage.removeItem(STORAGE_KEYS.plans);
    localStorage.removeItem(STORAGE_KEYS.recent);
    loadStorage();
    render();
  });

  const init = async () => {
    loadStorage();
    try {
      await fetchOpenOrders();
    } catch (err) {
      console.error(err);
    }
    render();
  };

  init();
})();
