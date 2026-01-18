(() => {
  "use strict";

  const CONFIG = {
    COST_ALERT_THRESHOLD: 250.0,
     BOOKING_IDLE_MS: 6000,
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

  const btnBookNow = $("btnBookNow");

  const MAX_ORDER_AGE_HOURS = 180;

  let activeOrderNo = null;
  let orderDetails = null;
  let parcelsForOrder = new Set();
  let armedForBooking = false;

  let placeCodeOverride = null;
  let serviceOverride = "RFX";
  let addressBook = [];
  let bookedOrders = new Set();

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
  };

  const appendDebug = (msg) => {
    if (!dbgOn || !debugLog) return;
    debugLog.textContent += `\n${new Date().toLocaleTimeString()} ${msg}`;
    debugLog.scrollTop = debugLog.scrollHeight;
  };

  function renderCountdown() {
    if (uiCountdown) uiCountdown.textContent = "--";
  }

  function money(v) {
    return v == null || isNaN(v) ? "-" : `R${Number(v).toFixed(2)}`;
  }
function isAutoBookOrder(details) {
  return hasParcelCountTag(details);
}

  function loadBookedOrders() {
    try {
      const raw = localStorage.getItem("fl_booked_orders_v1");
      if (raw) bookedOrders = new Set(JSON.parse(raw));
    } catch {}
  }

  function saveBookedOrders() {
    try {
      localStorage.setItem("fl_booked_orders_v1", JSON.stringify([...bookedOrders]));
    } catch {}
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
    }

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
}

function scheduleIdleAutoBook() {
  cancelAutoBookTimer();

  // Only for untagged orders
  if (!activeOrderNo || !orderDetails) return;
  if (isBooked(activeOrderNo)) return;
  if (hasParcelCountTag(orderDetails)) return;

  // Need at least 1 scan
  if (parcelsForOrder.size <= 0) return;

  autoBookTimer = setTimeout(async () => {
    autoBookTimer = null;

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
    }
  ];

  async function loadAddressBook() {
    try {
      const cached = localStorage.getItem("fl_addr_book_v2");
      if (cached) {
        const arr = JSON.parse(cached);
        if (Array.isArray(arr) && arr.length) {
          addressBook = arr;
          return;
        }
      }
    } catch {}

    try {
      const res = await fetch("/addresses.json", { cache: "no-cache" });
      if (!res.ok) throw new Error("HTTP " + res.status);
      const data = await res.json();
      if (!Array.isArray(data)) throw new Error("addresses.json must be an array");
      addressBook = data;
      localStorage.setItem("fl_addr_book_v2", JSON.stringify(addressBook));
    } catch {
      addressBook = ADDR_FALLBACK;
      try {
        localStorage.setItem("fl_addr_book_v2", JSON.stringify(addressBook));
      } catch {}
    }
  }

  function renderAddrResults(q) {
    if (!addrResults) return;
    const query = (q || "").trim().toLowerCase();
    const rows = addressBook
      .filter((e) => {
        const hay = `${e.label || ""} ${e.name || ""} ${e.address1 || ""} ${e.address2 || ""} ${e.city || ""} ${e.postal || ""}`.toLowerCase();
        return !query || hay.includes(query);
      })
      .slice(0, 100)
      .map((e, idx) =>
        `<div class="addrItem" data-idx="${idx}" role="option" tabindex="0"><strong>${e.label || e.name || "Address"}</strong> — ${e.city || ""} ${e.postal || ""}<br><span class="addrHint" style="color:#cbd5e1">${e.address1}${e.address2 ? ", " + e.address2 : ""}</span>${e.placeCode ? ` <span class='addrHint'>(code:${e.placeCode})</span>` : ""}</div>`
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

    if (shipToCard) {
      shipToCard.textContent = !orderDetails
        ? "None yet."
        : `${orderDetails.name}
${orderDetails.address1}
${orderDetails.address2 ? orderDetails.address2 + "\n" : ""}${orderDetails.city}
${orderDetails.province} ${orderDetails.postal}
Tel: ${orderDetails.phone || ""}
Email: ${orderDetails.email || ""}`.trim();
    }

    if (expected && parcelsForOrder.size) {
      statusExplain(
        `Scanning ${parcelsForOrder.size}/${expected} parcels`,
        parcelsForOrder.size === expected ? "ok" : "info"
      );
    } else if (activeOrderNo) {
      statusExplain(hasParcelCountTag(orderDetails) ? "Scan parcels until complete." : "Scan parcels, then BOOK NOW.", "info");
    }

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

    return `
    <div class="wb100x150" aria-label="Waybill ${waybillNo}, parcel ${parcelIdx} of ${parcelCount}">
      <div class="wbHead">
        <div class="wbMetaRow"><img src="img/download.jpg" style="width:80px; height:80px"></div>
        <div class="wbTopCode" aria-label="Parcel barcode top">
          <div class="wbTopHuman">${codeParcel}</div>
          ${svgBig}
        </div>
      </div>

      <div class="wbBody">
        <div class="wbFrom"><strong>FROM</strong><br>${fromHTML}</div>
        <div class="wbTo"><strong>SHIP&nbsp;TO</strong><br>${toHTML}</div>
      </div>

      <div class="wbFoot">
        <div class="wbTopCode" aria-label="Parcel barcode top">
          <div class="wbTopHuman">${codeParcel}</div>
          ${svgTopParcel}
        </div>
        <div class="podBox">
          <div class="podTitle">Proof of Delivery (POD) &nbsp; | &nbsp; Waybill: <strong>${waybillNo}</strong> </div>
          <div style="border-top:1px dotted #000"></div>
          <div>Receiver name:</div>
          <div style="border-top:1px dotted #000"></div>
          <div>Signature:</div>
          <div style="border-top:1px dotted #000"></div>
          <div>Date: ____/___/2025 &nbsp; | &nbsp; Time: ____:____</div>
        </div>
      </div>
    </div>`;
  }

  function mountLabelToPreviewAndPrint(firstHtml, allHtml) {
    if (stickerPreview) stickerPreview.innerHTML = `<div class="wbPreviewZoom">${firstHtml}</div>`;
    if (printMount) printMount.innerHTML = allHtml;
  }

  async function waitForImages(container) {
    const imgs = [...container.querySelectorAll("img")];
    await Promise.all(
      imgs.map((img) =>
        img.complete && img.naturalWidth
          ? Promise.resolve()
          : new Promise((res) => {
              img.onload = img.onerror = () => res();
            })
      )
    );
  }

  async function inlineImages(container) {
    const imgs = [...container.querySelectorAll("img")].filter((img) => img.src && !img.src.startsWith("data:"));
    for (const img of imgs) {
      try {
        const resp = await fetch(img.src, { cache: "force-cache" });
        const blob = await resp.blob();
        const dataURL = await new Promise((r) => {
          const fr = new FileReader();
          fr.onload = () => r(fr.result);
          fr.readAsDataURL(blob);
        });
        img.setAttribute("src", dataURL);
      } catch (e) {
        appendDebug("Inline image failed: " + img.src + " " + e);
      }
    }
  }

  async function ppCall(payload) {
    try {
      appendDebug(`PP CALL → method:${payload.method}, classVal:${payload.classVal}`);
      const res = await fetch(CONFIG.PP_ENDPOINT, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });
      const text = await res.text();
      let data;
      try { data = JSON.parse(text); } catch { data = { raw: text }; }
      appendDebug(`PP RESP ← HTTP ${res.status} ${res.statusText} :: ${text.slice(0, 500)}`);
      return { status: res.status, statusText: res.statusText, data };
    } catch (err) {
      appendDebug("PP NETWORK ERROR: " + String(err?.message || err));
      return { status: 0, statusText: String(err?.message || err), data: { error: "NETWORK", detail: String(err) } };
    }
  }

  function resolvePlaceCode(dest) {
    const key = `${(dest.city || "").trim().toLowerCase()}|${(dest.postal || "").trim()}`;
    const table = {
      "cape town|7530": 4001,
      "bellville|7530": 4001,
      "durbanville|7550": 4020,
      "cape town|8001": 3001
    };
    return table[key] || null;
  }

  async function lookupPlaceCodeFromPP(destDetails) {
    const suburb = (destDetails.suburb || destDetails.address2 || "").trim();
    const town = (destDetails.city || "").trim();

    const queries = [];
    if (suburb) queries.push(suburb);
    if (town && town.toLowerCase() !== suburb.toLowerCase()) {
      queries.push(town);
      if (suburb) queries.push(`${suburb} ${town}`);
    }

    if (!queries.length) return null;

    for (const q of queries) {
      try {
        appendDebug("PP getPlace query: " + q);
        const res = await fetch(`/pp/place?q=${encodeURIComponent(q)}`);
        if (!res.ok) throw new Error("HTTP " + res.status);
        const data = await res.json();

        if (data.errorcode && Number(data.errorcode) !== 0) continue;

        const list = Array.isArray(data.results) ? data.results : [];
        if (!list.length) continue;

        const targetTown = town.toLowerCase();
        const targetSuburb = suburb.toLowerCase();

        let best =
          (targetSuburb &&
            list.find((p) => {
              const name = (p.name || "").toLowerCase();
              const t = (p.town || "").toLowerCase();
              return ((name.includes(targetSuburb) || t.includes(targetSuburb)) && String(p.ring) === "0");
            })) ||
          (targetSuburb &&
            list.find((p) => {
              const name = (p.name || "").toLowerCase();
              const t = (p.town || "").toLowerCase();
              return name.includes(targetSuburb) || t.includes(targetSuburb);
            })) ||
          (targetTown &&
            list.find((p) => (p.town || "").trim().toLowerCase() === targetTown && String(p.ring) === "0")) ||
          (targetTown && list.find((p) => (p.town || "").trim().toLowerCase() === targetTown)) ||
          list.find((p) => String(p.ring) === "0") ||
          list[0];

        if (!best || best.place == null) continue;

        const code = Number(best.place);
        const label = (best.name || "").trim() + (best.town ? " – " + String(best.town).trim() : "");
        return { code, label, raw: best };
      } catch (e) {
        appendDebug("PP getPlace failed: " + String(e));
      }
    }
    return null;
  }

  function buildParcelPerfectPayload(destDetails, parcelCount) {
    const d = CONFIG.ORIGIN;

    let destplace =
      placeCodeOverride != null
        ? placeCodeOverride
        : destDetails && destDetails.placeCode != null
        ? Number(destDetails.placeCode) || destDetails.placeCode
        : resolvePlaceCode(destDetails) || null;

    let perParcelMass = CONFIG.BOX_DIM.massKg;
    if (destDetails && typeof destDetails.totalWeightKg === "number" && destDetails.totalWeightKg > 0 && parcelCount > 0) {
      perParcelMass = Number((destDetails.totalWeightKg / parcelCount).toFixed(2));
      if (perParcelMass <= 0) perParcelMass = CONFIG.BOX_DIM.massKg;
    }

    const details = {
      ...d,
      destpers: destDetails.name,
      destperadd1: destDetails.address1,
      destperadd2: destDetails.address2 || "",
      destperadd3: destDetails.city,
      destperadd4: destDetails.province,
      destperpcode: destDetails.postal,
      desttown: destDetails.city,
      destplace,
      destpercontact: destDetails.name,
      destperphone: destDetails.phone,
      notifydestpers: 1,
      destpercell: destDetails.phone || "0000000000",
      destperemail: destDetails.email,
      reference: `Order ${activeOrderNo}`
    };

    const contents = Array.from({ length: parcelCount }, (_, i) => ({
      item: i + 1,
      pieces: 1,
      dim1: CONFIG.BOX_DIM.dim1,
      dim2: CONFIG.BOX_DIM.dim2,
      dim3: CONFIG.BOX_DIM.dim3,
      actmass: perParcelMass
    }));

    return { details, contents };
  }

  async function fulfillOnShopify(details, waybillNo) {
    try {
      if (!details?.raw?.id) return;

      const orderId = details.raw.id;
      const lineItems = (details.raw.line_items || []).map((li) => ({ id: li.id, quantity: li.quantity }));

      const resp = await fetch(`${CONFIG.SHOPIFY.PROXY_BASE}/fulfill`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          orderId,
          lineItems,
          trackingNumber: waybillNo,
          trackingUrl: "",
          trackingCompany: "SWE / ParcelPerfect"
        })
      });

      const text = await resp.text();
      appendDebug("Shopify fulfill => " + resp.status + " " + text.slice(0, 300));
    } catch (e) {
      appendDebug("Shopify fulfill exception: " + String(e));
    }
  }

  function promptManualParcelCount(orderNo) {
    const raw = window.prompt(`Enter parcel count for order ${orderNo} (required):`, "");
    if (!raw) return null;
    const n = parseInt(String(raw).trim(), 10);
    if (!Number.isFinite(n) || n <= 0 || n > 999) return null;
    return n;
  }

  function extractQuoteFromV28(shape) {
    const obj = shape || {};
    if (obj.quoteno) return { quoteno: obj.quoteno, rates: obj.rates || [] };
    const res = Array.isArray(obj.results) && obj.results[0] ? obj.results[0] : null;
    const quoteno = (res && res.quoteno) || null;
    const rates = res && Array.isArray(res.rates) ? res.rates : [];
    return { quoteno, rates };
  }

  function pickService(rates) {
    const wanted = serviceOverride === "AUTO" ? ["RFX", "ECO", "RDF"] : [serviceOverride];
    const svcList = (rates || []).map((r) => r.service);
    for (const w of wanted) if (svcList.includes(w)) return w;
    return svcList[0] || "RDF";
  }

  async function doBookingNow(opts = {}) {
    if (!activeOrderNo || !orderDetails || armedForBooking) return;

    if (isBooked(activeOrderNo)) {
      statusExplain(`Order ${activeOrderNo} already booked — blocked.`, "warn");
      return;
    }

    const manual = !!opts.manual;
    const overrideCount = Number(opts.parcelCount || 0);

    let expected = getExpectedParcelCount(orderDetails);

    if (manual) {
      if (!overrideCount || overrideCount < 1) {
        statusExplain("Scan parcels first.", "warn");
        return;
      }
      expected = overrideCount;
      orderDetails.manualParcelCount = expected;
      renderSessionUI();
    } else {
      if (!expected) {
        const n = promptManualParcelCount(activeOrderNo);
        if (!n) {
          statusExplain("Parcel count required (cancelled).", "warn");
          return;
        }
        orderDetails.manualParcelCount = n;
        expected = n;
        renderSessionUI();
      }

      if (parcelsForOrder.size !== expected) {
        statusExplain(`Cannot book — scanned ${parcelsForOrder.size}/${expected}.`, "warn");
        return;
      }
    }

    const parcelIndexes = Array.from({ length: expected }, (_, i) => i + 1);

    armedForBooking = true;
    appendDebug("Booking order " + activeOrderNo + " parcels=" + parcelIndexes.join(", "));

    const missing = [];
    ["name", "address1", "city", "province", "postal"].forEach((k) => {
      if (!orderDetails[k]) missing.push(k);
    });

    const payload = buildParcelPerfectPayload(orderDetails, expected);
    if (!payload.details.destplace) missing.push("destplace (place code)");

    if (missing.length) {
      statusExplain("Quote failed", "err");
      if (bookingSummary) {
        bookingSummary.textContent = `Cannot request quote — missing: ${missing.join(", ")}\n\nShip To:\n${JSON.stringify(orderDetails, null, 2)}`;
      }
      armedForBooking = false;
      return;
    }

    const quoteRes = await ppCall({ method: "requestQuote", classVal: "quote", params: payload });
    if (!quoteRes || quoteRes.status !== 200) {
      statusExplain("Quote failed", "err");
      if (bookingSummary) {
        bookingSummary.textContent = `Quote error (HTTP ${quoteRes?.status}): ${quoteRes?.statusText}\n\n${JSON.stringify(quoteRes?.data, null, 2)}`;
      }
      if (quoteBox) quoteBox.textContent = "No quote — check place code / proxy / token.";
      armedForBooking = false;
      return;
    }

    const { quoteno, rates } = extractQuoteFromV28(quoteRes.data || {});
    if (!quoteno) {
      statusExplain("Quote failed", "err");
      if (bookingSummary) bookingSummary.textContent = `No quote number.\n${JSON.stringify(quoteRes.data, null, 2)}`;
      armedForBooking = false;
      return;
    }

    const pickedService = pickService(rates);
    const chosenRate = rates?.find((r) => r.service === pickedService) || rates?.[0] || null;
    const quoteCost = chosenRate ? Number(chosenRate.total ?? chosenRate.subtotal ?? chosenRate.charge ?? 0) : null;

    if (quoteBox && rates?.length) {
      const fmt = (v) => (isNaN(Number(v)) ? "-" : `R${Number(v).toFixed(2)}`);
      const lines = rates
        .map((r) => `${r.service}: ${fmt(r.total ?? r.subtotal ?? r.charge)} ${r.name ? `(${r.name})` : ""}`)
        .join("\n");
      quoteBox.textContent = `Selected: ${pickedService} • Est: ${fmt(quoteCost)}${quoteCost > CONFIG.COST_ALERT_THRESHOLD ? "  ⚠ high" : ""}\nOptions:\n${lines}`;
    }

    await ppCall({
      method: "updateService",
      classVal: "quote",
      params: { quoteno, service: pickedService, reference: String(activeOrderNo) }
    });

    const collRes = await ppCall({
      method: "quoteToCollection",
      classVal: "collection",
      params: { quoteno, starttime: "12:00", endtime: "15:00", printLabels: 1, printWaybill: 0 }
    });

    if (!collRes || collRes.status !== 200) {
      statusExplain("Booking failed", "err");
      if (bookingSummary) bookingSummary.textContent = `Booking error: HTTP ${collRes?.status} ${collRes?.statusText}\n${JSON.stringify(collRes?.data, null, 2)}`;
      armedForBooking = false;
      return;
    }

    const cr = collRes.data || {};
    const maybe = cr.results?.[0] || cr;
    const waybillNo = String(maybe.waybill || maybe.waybillno || maybe.waybillNo || maybe.trackingNo || "WB-TEST-12345");
    appendDebug("Waybill = " + waybillNo);

    const labelsBase64 = maybe.labelsBase64 || maybe.labelBase64 || maybe.labels_pdf || null;
    const waybillBase64 = maybe.waybillBase64 || maybe.waybillPdfBase64 || maybe.waybill_pdf || null;

    let usedPdf = false;

    if (labelsBase64) {
      usedPdf = true;

      try {
        await fetch("/printnode/print", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ pdfBase64: labelsBase64, title: `Labels ${waybillNo}` })
        });
      } catch (e) {
        appendDebug("PrintNode label error: " + String(e));
      }

      if (waybillBase64) {
        try {
          await fetch("/printnode/print", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ pdfBase64: waybillBase64, title: `Waybill ${waybillNo}` })
          });
        } catch (e) {
          appendDebug("PrintNode waybill error: " + String(e));
        }
      }

      if (stickerPreview) {
        stickerPreview.innerHTML = `
          <div class="wbPreviewPdf">
            <div style="font-weight:600;margin-bottom:0.25rem;">Labels sent to PrintNode</div>
            <div style="font-size:0.8rem;color:#64748b;">
              Waybill: <strong>${waybillNo}</strong><br>
              Service: ${pickedService} • Parcels: ${expected}
            </div>
          </div>`;
      }
      if (printMount) printMount.innerHTML = "";
    } else {
      const labels = parcelIndexes.map((idx) => renderLabelHTML(waybillNo, pickedService, quoteCost, orderDetails, idx, expected));
      mountLabelToPreviewAndPrint(labels[0], labels.join("\n"));
      statusExplain("Booked", "ok");

      await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
      await inlineImages(printMount);
      await waitForImages(printMount);

      if (labels.length) window.print();
    }

    if (statusChip) statusChip.textContent = "Booked";
    if (bookingSummary) {
      bookingSummary.textContent = `WAYBILL: ${waybillNo}
Service: ${pickedService}
Parcels: ${expected}
Estimated Cost: ${money(quoteCost)}

${usedPdf ? "Label + waybill generated by ParcelPerfect (PDF)." : "Using local HTML label layout."}

Raw:
${JSON.stringify(cr, null, 2)}`;
    }

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

  // TAGGED: auto-book immediately on first scan
  if (hasParcelCountTag(orderDetails) && expected) {
    cancelAutoBookTimer();

    parcelsForOrder = new Set(Array.from({ length: expected }, (_, i) => i + 1));
    renderSessionUI();
    updateBookNowButton();

    statusExplain(`Tag detected (parcel_count_${expected}). Auto-booking...`, "ok");
    await doBookingNow();
    return;
  }

  // UNTAGGED: schedule auto-book 6s after last scan
  renderSessionUI();
  updateBookNowButton();
  scheduleIdleAutoBook();
}



  async function fetchShopifyOrder(orderNo) {
    try {
      const url = `${CONFIG.SHOPIFY.PROXY_BASE}/orders/by-name/${encodeURIComponent(orderNo)}`;
      const res = await fetch(url);
      if (!res.ok) throw new Error("HTTP " + res.status);
      const data = await res.json();
      const o = data.order || data || {};
      const placeCodeFromMeta = data.customerPlaceCode || null;

      const shipping = o.shipping_address || {};
      const customer = o.customer || {};
      const lineItems = o.line_items || [];

      let parcelCountFromTag = null;
      if (typeof o.tags === "string" && o.tags.trim()) {
        const parts = o.tags.split(",").map((t) => t.trim().toLowerCase());
        for (const t of parts) {
          const m = t.match(/^parcel_count_(\d+)$/);
          if (m) { parcelCountFromTag = parseInt(m[1], 10); break; }
        }
      }

      let totalGrams = 0;
      for (const li of lineItems) {
        const gramsPerUnit = Number(li.grams || 0);
        const qty = Number(li.quantity || 1);
        totalGrams += gramsPerUnit * qty;
      }
      const totalWeightKg = totalGrams / 1000;

      const name =
        shipping.name ||
        `${customer.first_name || ""} ${customer.last_name || ""}`.trim() ||
        o.name ||
        String(orderNo);

      const normalized = {
        raw: o,
        name,
        phone: shipping.phone || customer.phone || "",
        email: o.email || "",
        address1: shipping.address1 || "",
        address2: shipping.address2 || "",
        city: shipping.city || "",
        province: shipping.province || "",
        postal: shipping.zip || "",
        suburb: shipping.address2 || "",
        line_items: lineItems,
        totalWeightKg,
        placeCode: placeCodeFromMeta,
        placeLabel: null,
        parcelCountFromTag,
        manualParcelCount: null
      };

      if (!placeCodeFromMeta) {
        const lookedUp = await lookupPlaceCodeFromPP(normalized);
        if (lookedUp?.code != null) {
          normalized.placeCode = lookedUp.code;
          normalized.placeLabel = lookedUp.label;
        }
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

    if (!list.length) {
      dispatchBoard.innerHTML = `<div class="dispatchBoardEmpty">No open shipping / delivery / collections right now.</div>`;
      return;
    }

    const cols = [{ id: "lane1", label: "" }, { id: "lane2", label: "" }, { id: "lane3", label: "" }, { id: "lane4", label: "" }];
    const lanes = cols.map(() => []);
    list.forEach((o, idx) => lanes[idx % lanes.length].push(o));

    const cardHTML = (o) => {
      const title = o.customer_name || o.name || `Order ${o.id}`;
      const city = o.shipping_city || "";
      const postal = o.shipping_postal || "";
      const created = o.created_at ? new Date(o.created_at).toLocaleTimeString() : "";
      const lines = (o.line_items || []).slice(0, 6).map((li) => `• ${li.quantity} × ${li.title}`).join("<br>");
      const addr1 = o.shipping_address1 || "";
      const addr2 = o.shipping_address2 || "";
      const addrHtml = `${addr1}${addr2 ? "<br>" + addr2 : ""}<br>${city} ${postal}`;

      return `
        <div class="dispatchCard">
          <div class="dispatchCardTitle"><span>${title}</span></div>
          <div class="dispatchCardMeta">#${(o.name || "").replace("#", "")} · ${city} · ${created}</div>
          <div class="dispatchCardAddress">${addrHtml}</div>
          <div class="dispatchCardLines">${lines}</div>
        </div>`;
    };

    dispatchBoard.innerHTML = cols
      .map((col, idx) => {
        const cards = lanes[idx].map(cardHTML).join("") || `<div class="dispatchBoardEmptyCol">No jobs.</div>`;
        return `
          <div class="dispatchCol">
            <div class="dispatchColHeader">${col.label}</div>
            <div class="dispatchColBody">${cards}</div>
          </div>`;
      })
      .join("");
  }

  async function refreshDispatchData() {
    try {
      const res = await fetch(`${CONFIG.SHOPIFY.PROXY_BASE}/orders/open`);
      const data = await res.json();
      renderDispatchBoard(data.orders || []);
      if (dispatchStamp) dispatchStamp.textContent = "Updated " + new Date().toLocaleTimeString();
    } catch (e) {
      appendDebug("Dispatch refresh failed: " + String(e));
      if (dispatchBoard) dispatchBoard.innerHTML = `<div class="dispatchBoardEmpty">Error loading orders.</div>`;
      if (dispatchStamp) dispatchStamp.textContent = "Dispatch: error";
    }
  }

  function switchMainView(view) {
    const showScan = view === "scan";

    if (viewScan) {
      viewScan.hidden = !showScan;
      viewScan.classList.toggle("flView--active", showScan);
    }
    if (viewOps) {
      viewOps.hidden = showScan;
      viewOps.classList.toggle("flView--active", !showScan);
    }

    navScan?.classList.toggle("flNavBtn--active", showScan);
    navOps?.classList.toggle("flNavBtn--active", !showScan);

    if (showScan) {
      statusExplain("Ready to scan orders…", "info");
      scanInput?.focus();
    } else {
      statusExplain("Viewing orders / ops dashboard", "info");
    }
  }

  scanInput?.addEventListener("keydown", async (e) => {
    if (e.key === "Enter") {
      const code = scanInput.value.trim();
      scanInput.value = "";
      if (!code) return;
      await handleScan(code);
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
    if (bookingSummary) bookingSummary.textContent = "";
    if (scanInput) scanInput.value = "";
    if (dbgOn && debugLog) debugLog.textContent = "";
    switchMainView("scan");
  });

  navScan?.addEventListener("click", () => switchMainView("scan"));
  navOps?.addEventListener("click", () => switchMainView("ops"));

  loadBookedOrders();
  renderSessionUI();
  renderCountdown();
  initAddressSearch();
  refreshDispatchData();
  setInterval(refreshDispatchData, 30000);
  switchMainView("scan");

  if (location.protocol === "file:") {
    alert("Open via http://localhost/... (not file://). Run a local server.");
  }

  window.__fl = window.__fl || {};
  window.__fl.bookNow = doBookingNow;
})();