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
    } catch (err) {
      packPlans = {};
    }

    try {
      const rawRecent = localStorage.getItem(STORAGE_KEYS.recent);
      recentFulfilled = rawRecent ? JSON.parse(rawRecent) : [];
    } catch (err) {
      recentFulfilled = [];
    }
  };

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
    } catch (err) {
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
