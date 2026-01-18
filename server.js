import path from "path";
import { fileURLToPath } from "url";

import express from "express";
import cors from "cors";
import helmet from "helmet";
import rateLimit from "express-rate-limit";
import morgan from "morgan";
import dotenv from "dotenv";

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();

const {
  PORT = "3000",
  NODE_ENV = "development",
  FRONTEND_ORIGIN = "http://localhost:3000",

  PP_BASE_URL = "",
  PP_TOKEN = "",
  PP_REQUIRE_TOKEN = "true",
  PP_ACCNUM = "",
  PP_PLACE_ID = "",

  SHOPIFY_STORE = "",
  SHOPIFY_CLIENT_ID = "",
  SHOPIFY_CLIENT_SECRET = "",
  SHOPIFY_API_VERSION = "2025-10",

  PRINTNODE_API_KEY = "",
  PRINTNODE_PRINTER_ID = ""
} = process.env;

app.disable("x-powered-by");
app.use(helmet({ crossOriginResourcePolicy: { policy: "cross-origin" } }));
app.use(express.json({ limit: "1mb" }));

const allowedOrigins = new Set(
  String(FRONTEND_ORIGIN || "")
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean)
);

app.use(
  cors({
    origin: (origin, cb) => {
      if (!origin || allowedOrigins.has(origin)) return cb(null, true);
      return cb(new Error("CORS: origin not allowed"));
    },
    methods: ["GET", "POST", "OPTIONS"],
    allowedHeaders: ["Content-Type", "Authorization"],
    credentials: false,
    maxAge: 86400
  })
);
app.options("*", (_req, res) => res.sendStatus(204));

app.use(
  rateLimit({
    windowMs: 60 * 1000,
    max: 120,
    standardHeaders: true,
    legacyHeaders: false
  })
);

app.use(morgan(NODE_ENV === "production" ? "combined" : "dev"));

function badRequest(res, message, detail) {
  return res.status(400).json({ error: "BAD_REQUEST", message, detail });
}

function requireShopifyConfigured(res) {
  if (!SHOPIFY_STORE || !SHOPIFY_CLIENT_ID || !SHOPIFY_CLIENT_SECRET) {
    res.status(501).json({
      error: "SHOPIFY_NOT_CONFIGURED",
      message:
        "Set SHOPIFY_STORE, SHOPIFY_CLIENT_ID, SHOPIFY_CLIENT_SECRET in .env (Dev Dashboard app credentials)"
    });
    return false;
  }
  return true;
}

let cachedToken = null;
let tokenExpiresAtMs = 0;

async function fetchNewShopifyAdminToken() {
  const url = `https://${SHOPIFY_STORE}.myshopify.com/admin/oauth/access_token`;

  const body = new URLSearchParams();
  body.set("grant_type", "client_credentials");
  body.set("client_id", SHOPIFY_CLIENT_ID);
  body.set("client_secret", SHOPIFY_CLIENT_SECRET);

  const resp = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body
  });

  const text = await resp.text();
  if (!resp.ok) throw new Error(`Shopify token request failed (${resp.status}): ${text}`);

  const data = JSON.parse(text);
  if (!data.access_token) throw new Error(`Shopify token response missing access_token: ${text}`);

  const expiresInSec = Number(data.expires_in || 0);
  const bufferMs = 60_000;
  cachedToken = data.access_token;
  tokenExpiresAtMs = Date.now() + Math.max(0, expiresInSec * 1000 - bufferMs);
  return cachedToken;
}

async function getShopifyAdminToken() {
  if (cachedToken && Date.now() < tokenExpiresAtMs) return cachedToken;
  return fetchNewShopifyAdminToken();
}

async function shopifyFetch(pathname, { method = "GET", headers = {}, body } = {}) {
  const token = await getShopifyAdminToken();
  const url = `https://${SHOPIFY_STORE}.myshopify.com${pathname}`;

  const doFetch = async (tokenToUse) =>
    fetch(url, {
      method,
      headers: {
        "X-Shopify-Access-Token": tokenToUse,
        "Content-Type": "application/json",
        Accept: "application/json",
        ...headers
      },
      body
    });

  const resp = await doFetch(token);

  if (resp.status === 401 || resp.status === 403) {
    cachedToken = null;
    tokenExpiresAtMs = 0;
    const token2 = await getShopifyAdminToken();
    return doFetch(token2);
  }

  return resp;
}

async function shopifyGraphql(query, variables = {}) {
  const resp = await shopifyFetch(`/admin/api/${SHOPIFY_API_VERSION}/graphql.json`, {
    method: "POST",
    body: JSON.stringify({ query, variables })
  });

  const text = await resp.text();
  let payload;
  try {
    payload = JSON.parse(text);
  } catch {
    payload = { raw: text };
  }

  if (!resp.ok) {
    const error = new Error(`Shopify GraphQL error ${resp.status}`);
    error.status = resp.status;
    error.payload = payload;
    throw error;
  }

  return payload;
}

app.post("/pp", async (req, res) => {
  try {
    const { method, classVal, params } = req.body || {};

    if (!method || !classVal || typeof params !== "object") {
      return badRequest(res, "Expected { method, classVal, params } in body");
    }

    if (!PP_BASE_URL || !PP_BASE_URL.startsWith("http")) {
      return res.status(500).json({ error: "CONFIG_ERROR", message: "PP_BASE_URL is not a valid URL" });
    }

    const form = new URLSearchParams();
    form.set("method", String(method));
    form.set("class", String(classVal));
    form.set("params", JSON.stringify(params));

    const mustUseToken = String(PP_REQUIRE_TOKEN).toLowerCase() === "true";
    if (mustUseToken && PP_TOKEN) form.set("token_id", PP_TOKEN);

    const upstream = await fetch(PP_BASE_URL, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: form.toString()
    });

    const text = await upstream.text();
    res.set("content-type", upstream.headers.get("content-type") || "application/json; charset=utf-8");

    try {
      return res.status(upstream.status).json(JSON.parse(text));
    } catch {
      return res.status(upstream.status).send(text);
    }
  } catch (err) {
    return res.status(502).json({ error: "UPSTREAM_ERROR", message: String(err?.message || err) });
  }
});

app.get("/pp/place", async (req, res) => {
  try {
    const query = String(req.query.q || req.query.query || "").trim();
    if (!query) return badRequest(res, "Missing ?q= query string for place search");

    if (!PP_BASE_URL || !PP_BASE_URL.startsWith("http")) {
      return res.status(500).json({ error: "CONFIG_ERROR", message: "PP_BASE_URL is not a valid URL" });
    }

    if (!PP_TOKEN) {
      return res.status(500).json({ error: "CONFIG_ERROR", message: "PP_TOKEN is required for getPlace" });
    }

    const paramsObj = {
      id: PP_PLACE_ID || "ShopifyScanStation",
      accnum: PP_ACCNUM || "",
      ppcust: ""
    };

    const qs = new URLSearchParams();
    qs.set("Class", "Waybill");
    qs.set("method", "getPlace");
    qs.set("token_id", PP_TOKEN);
    qs.set("params", JSON.stringify(paramsObj));
    qs.set("query", query);

    const base = PP_BASE_URL.endsWith("/") ? PP_BASE_URL : `${PP_BASE_URL}/`;
    const url = `${base}?${qs.toString()}`;

    const upstream = await fetch(url, { method: "GET" });
    const text = await upstream.text();

    try {
      return res.status(upstream.status).json(JSON.parse(text));
    } catch {
      return res.status(upstream.status).send(text);
    }
  } catch (err) {
    return res.status(502).json({ error: "UPSTREAM_ERROR", message: String(err?.message || err) });
  }
});

app.get("/shopify/orders/by-name/:name", async (req, res) => {
  try {
    if (!requireShopifyConfigured(res)) return;

    let name = String(req.params.name || "");
    if (!name.startsWith("#")) name = `#${name}`;

    const base = `/admin/api/${SHOPIFY_API_VERSION}`;
    const orderUrl = `${base}/orders.json?status=any&name=${encodeURIComponent(name)}`;

    const orderResp = await shopifyFetch(orderUrl, { method: "GET" });
    if (!orderResp.ok) {
      const body = await orderResp.text();
      return res.status(orderResp.status).json({
        error: "SHOPIFY_UPSTREAM",
        status: orderResp.status,
        statusText: orderResp.statusText,
        body
      });
    }

    const orderData = await orderResp.json();
    const order = Array.isArray(orderData.orders) && orderData.orders.length ? orderData.orders[0] : null;

    if (!order) return res.status(404).json({ error: "NOT_FOUND", message: "Order not found" });

    let customerPlaceCode = null;
    try {
      if (order.customer && order.customer.id) {
        const metaUrl = `${base}/customers/${order.customer.id}/metafields.json`;
        const metaResp = await shopifyFetch(metaUrl, { method: "GET" });
        if (metaResp.ok) {
          const metaData = await metaResp.json();
          const m = (metaData.metafields || []).find(
            (mf) => mf.namespace === "custom" && mf.key === "parcelperfect_place_code"
          );
          if (m && m.value) customerPlaceCode = m.value;
        }
      }
    } catch {
      customerPlaceCode = null;
    }

    return res.json({ order, customerPlaceCode });
  } catch (err) {
    return res.status(502).json({ error: "UPSTREAM_ERROR", message: String(err?.message || err) });
  }
});

app.get("/shopify/orders/open", async (req, res) => {
  try {
    if (!requireShopifyConfigured(res)) return;

    const base = `/admin/api/${SHOPIFY_API_VERSION}`;
    const url =
      `${base}/orders.json?status=any` +
      `&fulfillment_status=unfulfilled,in_progress` +
      `&limit=50&order=created_at+desc`;

    const resp = await shopifyFetch(url, { method: "GET" });
    if (!resp.ok) {
      const body = await resp.text();
      return res.status(resp.status).json({
        error: "SHOPIFY_UPSTREAM",
        status: resp.status,
        statusText: resp.statusText,
        body
      });
    }

    const data = await resp.json();
    const ordersRaw = Array.isArray(data.orders) ? data.orders : [];

    const orders = ordersRaw.map((o) => {
      const shipping = o.shipping_address || {};
      const customer = o.customer || {};

      let parcelCountFromTag = null;
      if (typeof o.tags === "string" && o.tags.trim()) {
        const parts = o.tags
          .split(",")
          .map((t) => t.trim().toLowerCase())
          .filter(Boolean);
        for (const t of parts) {
          const m = t.match(/^parcel_count_(\d+)$/);
          if (m) {
            parcelCountFromTag = parseInt(m[1], 10);
            break;
          }
        }
      }

      const companyName =
        (shipping.company && String(shipping.company).trim()) ||
        (customer?.default_address?.company && String(customer.default_address.company).trim());

      const customer_name =
        companyName ||
        shipping.name ||
        `${String(customer.first_name || "").trim()} ${String(customer.last_name || "").trim()}`.trim() ||
        (o.name ? String(o.name).replace(/^#/, "") : "");

      return {
        id: o.id,
        name: o.name,
        order_gid: o.admin_graphql_api_id,
        customer_name,
        created_at: o.processed_at || o.created_at,
        email: o.email,
        fulfillment_status: o.fulfillment_status,
        shipping_city: shipping.city || "",
        shipping_postal: shipping.zip || "",
        shipping_address1: shipping.address1 || "",
        shipping_address2: shipping.address2 || "",
        shipping_province: shipping.province || "",
        shipping_country: shipping.country || "",
        shipping_phone: shipping.phone || "",
        shipping_name: shipping.name || customer_name,
        parcel_count: parcelCountFromTag,
        line_items: (o.line_items || []).map((li) => ({
          id: li.id,
          gid: li.admin_graphql_api_id,
          title: li.title,
          quantity: li.quantity,
          fulfillable_quantity: li.fulfillable_quantity
        }))
      };
    });

    return res.json({ orders });
  } catch (err) {
    return res.status(502).json({ error: "UPSTREAM_ERROR", message: String(err?.message || err) });
  }
});

app.post("/shopify/fulfill", async (req, res) => {
  try {
    if (!requireShopifyConfigured(res)) return;

    const { orderId, trackingNumber, trackingUrl, trackingCompany } = req.body || {};
    if (!orderId) return res.status(400).json({ error: "MISSING_ORDER_ID", body: req.body });

    const base = `/admin/api/${SHOPIFY_API_VERSION}`;
    const trackingCompanyFinal = trackingCompany || process.env.TRACKING_COMPANY || "SWE / ParcelPerfect";

    const foUrl = `${base}/orders/${orderId}/fulfillment_orders.json`;
    const foResp = await shopifyFetch(foUrl, { method: "GET" });
    const foText = await foResp.text();

    let foData;
    try {
      foData = JSON.parse(foText);
    } catch {
      foData = { raw: foText };
    }

    if (!foResp.ok) {
      return res.status(foResp.status).json({
        error: "SHOPIFY_UPSTREAM",
        status: foResp.status,
        statusText: foResp.statusText,
        body: foData
      });
    }

    const fulfillmentOrders = Array.isArray(foData.fulfillment_orders) ? foData.fulfillment_orders : [];
    if (!fulfillmentOrders.length) {
      return res.status(409).json({
        error: "NO_FULFILLMENT_ORDERS",
        message: "No fulfillment_orders found for this order (cannot fulfill)",
        body: foData
      });
    }

    const fo = fulfillmentOrders.find((f) => f.status !== "closed" && f.status !== "cancelled") || fulfillmentOrders[0];
    const fulfillment_order_id = fo.id;

    const fulfillUrl = `${base}/fulfillments.json`;
    const fulfillmentPayload = {
      fulfillment: {
        message: "Shipped via Scan Station",
        notify_customer: true,
        tracking_info: {
          number: trackingNumber || "",
          url: trackingUrl || undefined,
          company: trackingCompanyFinal
        },
        line_items_by_fulfillment_order: [{ fulfillment_order_id }]
      }
    };

    const upstream = await shopifyFetch(fulfillUrl, {
      method: "POST",
      body: JSON.stringify(fulfillmentPayload)
    });

    const text = await upstream.text();
    let data;
    try {
      data = JSON.parse(text);
    } catch {
      data = { raw: text };
    }

    if (!upstream.ok) {
      return res.status(upstream.status).json({
        error: "SHOPIFY_UPSTREAM",
        status: upstream.status,
        statusText: upstream.statusText,
        body: data
      });
    }

    return res.json({ ok: true, fulfillment: data });
  } catch (err) {
    return res.status(502).json({ error: "UPSTREAM_ERROR", message: String(err?.message || err) });
  }
});

app.post("/shopify/graphql", async (req, res) => {
  try {
    if (!requireShopifyConfigured(res)) return;

    const { query, variables } = req.body || {};
    if (!query) return badRequest(res, "Missing GraphQL query in body");

    const payload = await shopifyGraphql(query, variables || {});
    return res.json(payload);
  } catch (err) {
    return res.status(err.status || 502).json({
      error: "SHOPIFY_GRAPHQL_ERROR",
      message: String(err?.message || err),
      payload: err.payload || null
    });
  }
});

app.post("/shopify/fulfillment/create", async (req, res) => {
  try {
    if (!requireShopifyConfigured(res)) return;

    const { orderGid, shipments, notifyCustomer } = req.body || {};
    if (!orderGid) return badRequest(res, "Missing orderGid");
    if (!Array.isArray(shipments) || !shipments.length) return badRequest(res, "Missing shipments array");

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

    const data = await shopifyGraphql(query, { orderId: orderGid });
    const edges = data?.data?.order?.fulfillmentOrders?.edges || [];

    const lineItemMap = new Map();
    for (const edge of edges) {
      const fo = edge.node;
      for (const liEdge of fo.lineItems.edges || []) {
        const node = liEdge.node;
        lineItemMap.set(node.id, {
          fulfillmentOrderId: fo.id,
          remainingQuantity: node.remainingQuantity
        });
      }
    }

    const results = [];
    for (let i = 0; i < shipments.length; i += 1) {
      const shipment = shipments[i];
      const tracking = shipment.tracking || {};
      const items = Array.isArray(shipment.lineItems) ? shipment.lineItems : [];

      const grouped = new Map();
      const userErrors = [];

      for (const item of items) {
        const mapEntry = lineItemMap.get(item.fulfillmentOrderLineItemId);
        if (!mapEntry) {
          userErrors.push({ message: `Unknown fulfillment order line item ${item.fulfillmentOrderLineItemId}` });
          continue;
        }

        const qty = Number(item.quantity || 0);
        if (!qty || qty <= 0) continue;

        if (mapEntry.remainingQuantity != null && qty > mapEntry.remainingQuantity) {
          userErrors.push({ message: `Quantity exceeds remaining for ${item.fulfillmentOrderLineItemId}` });
          continue;
        }

        if (!grouped.has(mapEntry.fulfillmentOrderId)) grouped.set(mapEntry.fulfillmentOrderId, []);
        grouped.get(mapEntry.fulfillmentOrderId).push({ id: item.fulfillmentOrderLineItemId, quantity: qty });
      }

      if (!grouped.size) {
        results.push({
          shipmentIndex: i,
          fulfillmentId: null,
          status: "error",
          userErrors: userErrors.length ? userErrors : [{ message: "No line items to fulfill" }],
          notifyErrors: []
        });
        continue;
      }

      const fulfillmentInput = {
        lineItemsByFulfillmentOrder: Array.from(grouped.entries()).map(([fulfillmentOrderId, fulfillmentOrderLineItems]) => ({
          fulfillmentOrderId,
          fulfillmentOrderLineItems
        })),
        trackingInfo: {
          number: tracking.number || "",
          url: tracking.url || null,
          company: tracking.company || "SWE / ParcelPerfect"
        },
        notifyCustomer: false
      };

      const mutation = `mutation FulfillmentCreate($fulfillment: FulfillmentV2Input!) {
        fulfillmentCreate(fulfillment: $fulfillment) {
          fulfillment { id status }
          userErrors { field message }
        }
      }`;

      let fulfillmentId = null;
      let notifyErrors = [];
      let status = "created";
      let createErrors = userErrors;

      if (!userErrors.length) {
        const createResp = await shopifyGraphql(mutation, { fulfillment: fulfillmentInput });
        const createData = createResp?.data?.fulfillmentCreate;
        const createUserErrors = createData?.userErrors || [];
        if (createUserErrors.length) {
          createErrors = createUserErrors;
          status = "error";
        } else {
          fulfillmentId = createData?.fulfillment?.id || null;
          status = createData?.fulfillment?.status || "success";
        }

        if (fulfillmentId && notifyCustomer) {
          const notifyMutation = `mutation FulfillmentNotify($id: ID!) {
            fulfillmentNotify(fulfillmentId: $id, notifyCustomer: true) {
              fulfillment { id status }
              userErrors { field message }
            }
          }`;
          const notifyResp = await shopifyGraphql(notifyMutation, { id: fulfillmentId });
          notifyErrors = notifyResp?.data?.fulfillmentNotify?.userErrors || [];
        }
      }

      results.push({
        shipmentIndex: i,
        fulfillmentId,
        status,
        userErrors: createErrors,
        notifyErrors
      });
    }

    return res.json({ ok: true, results });
  } catch (err) {
    return res.status(err.status || 502).json({
      error: "UPSTREAM_ERROR",
      message: String(err?.message || err),
      payload: err.payload || null
    });
  }
});

app.post("/printnode/print", async (req, res) => {
  try {
    const { pdfBase64, title } = req.body || {};
    if (!pdfBase64) return res.status(400).json({ error: "BAD_REQUEST", message: "Missing pdfBase64" });

    if (!PRINTNODE_API_KEY || !PRINTNODE_PRINTER_ID) {
      return res.status(500).json({
        error: "PRINTNODE_NOT_CONFIGURED",
        message: "Set PRINTNODE_API_KEY and PRINTNODE_PRINTER_ID in your .env file"
      });
    }

    const auth = Buffer.from(`${PRINTNODE_API_KEY}:`).toString("base64");

    const payload = {
      printerId: Number(PRINTNODE_PRINTER_ID),
      title: title || "Parcel Label",
      contentType: "pdf_base64",
      content: String(pdfBase64).replace(/\s/g, ""),
      source: "Flippen Lekka Scan Station"
    };

    const upstream = await fetch("https://api.printnode.com/printjobs", {
      method: "POST",
      headers: {
        Authorization: `Basic ${auth}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify(payload)
    });

    const text = await upstream.text();
    let data;
    try {
      data = JSON.parse(text);
    } catch {
      data = { raw: text };
    }

    if (!upstream.ok) {
      return res.status(upstream.status).json({
        error: "PRINTNODE_UPSTREAM",
        status: upstream.status,
        statusText: upstream.statusText,
        body: data
      });
    }

    return res.json({ ok: true, printJob: data });
  } catch (err) {
    return res.status(502).json({ error: "UPSTREAM_ERROR", message: String(err?.message || err) });
  }
});

app.get("/healthz", (_req, res) => res.json({ ok: true }));

app.use("/img", express.static(path.join(__dirname, "img")));

app.get("/", (_req, res) => res.sendFile(path.join(__dirname, "index.html")));
app.get("/app.js", (_req, res) => res.sendFile(path.join(__dirname, "app.js")));

app.get("*", (_req, res) => res.sendFile(path.join(__dirname, "index.html")));

app.listen(Number(PORT), () => {
  console.log(`Scan Station server listening on http://localhost:${PORT}`);
  console.log(`Allowed origins: ${[...allowedOrigins].join(", ") || "(none)"}`);
});
