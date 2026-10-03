import { createServer } from "node:http";
import { readFile, stat } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.dirname(fileURLToPath(import.meta.url));
const dist = path.join(root, "dist");
const maxFileBytes = 4 * 1024 * 1024;
const maxBodyBytes = 5_800_000;
const requestsByIp = new Map();

function sendJson(res, status, value) {
  res.writeHead(status, { "Content-Type": "application/json; charset=utf-8", "Cache-Control": "no-store", "X-Content-Type-Options": "nosniff" });
  res.end(JSON.stringify(value));
}

async function readJson(req) {
  const chunks = [];
  let size = 0;
  for await (const chunk of req) {
    size += chunk.length;
    if (size > maxBodyBytes) throw Object.assign(new Error("Upload is too large. Maximum file size is 4 MB."), { status: 413 });
    chunks.push(chunk);
  }
  try { return JSON.parse(Buffer.concat(chunks).toString("utf8")); }
  catch { throw Object.assign(new Error("Invalid request."), { status: 400 }); }
}

function getFieldValue(field) {
  if (!field) return undefined;
  if (field.valueCurrency && Number.isFinite(field.valueCurrency.amount)) return field.valueCurrency.amount;
  if (Number.isFinite(field.valueNumber)) return field.valueNumber;
  if (typeof field.valueString === "string") return field.valueString;
  if (typeof field.content === "string") return field.content;
  return undefined;
}

function extractLineItems(payload, model) {
  const fields = payload?.analyzeResult?.documents?.[0]?.fields || {};
  const itemField = fields.Items || fields.LineItems;
  const itemRows = itemField?.valueArray || [];
  return itemRows.map(entry => {
    const item = entry?.valueObject || entry?.valueArray || {};
    const get = (...names) => {
      for (const name of names) {
        const value = getFieldValue(item[name]);
        if (value !== undefined && value !== "") return value;
      }
      return undefined;
    };
    const name = get("Name", "Description", "ProductName");
    const quantity = Number(get("Quantity")) || 1;
    let price = Number(get(model === "prebuilt-receipt" ? "Price" : "UnitPrice", "Price"));
    if (!Number.isFinite(price)) {
      const total = Number(get("TotalPrice", "Amount"));
      if (Number.isFinite(total)) price = quantity > 1 ? total / quantity : total;
    }
    return { name: typeof name === "string" ? name.trim() : "", price: Math.round(price * 100) / 100 };
  }).filter(item => item.name && Number.isFinite(item.price) && item.price >= 0);
}

async function scanInvoice(req, res) {
  const endpoint = process.env.AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT?.trim().replace(/\/+$/, "");
  const key = process.env.AZURE_DOCUMENT_INTELLIGENCE_KEY?.trim();
  if (!endpoint || !key) {
    return sendJson(res, 503, { error: "Azure scanning isn’t configured yet. Add AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT and AZURE_DOCUMENT_INTELLIGENCE_KEY to Railway variables." });
  }
  let azureEndpoint;
  try { azureEndpoint = new URL(endpoint); } catch { return sendJson(res, 500, { error: "The Azure Document Intelligence endpoint is invalid." }); }
  if (azureEndpoint.protocol !== "https:") return sendJson(res, 500, { error: "The Azure endpoint must use HTTPS." });

  const ip = req.headers["x-forwarded-for"]?.split(",")[0]?.trim() || req.socket.remoteAddress || "unknown";
  const now = Date.now();
  const bucket = requestsByIp.get(ip) || { startedAt: now, count: 0 };
  if (now - bucket.startedAt > 60_000) { bucket.startedAt = now; bucket.count = 0; }
  if (bucket.count >= 10) return sendJson(res, 429, { error: "Too many scans in a short time. Please wait a minute and try again." });
  bucket.count += 1; requestsByIp.set(ip, bucket);

  const body = await readJson(req);
  const model = body.model;
  if (!["prebuilt-invoice", "prebuilt-receipt"].includes(model)) return sendJson(res, 400, { error: "Choose the Invoice or Receipt model." });
  if (typeof body.base64 !== "string" || !/^[A-Za-z0-9+/]+={0,2}$/.test(body.base64)) return sendJson(res, 400, { error: "The uploaded file could not be read." });
  const file = Buffer.from(body.base64, "base64");
  if (!file.length || file.length > maxFileBytes) return sendJson(res, 413, { error: "Upload a non-empty PDF or image no larger than 4 MB." });
  const allowedTypes = new Set(["application/pdf", "image/jpeg", "image/png"]);
  const mimeType = allowedTypes.has(body.mimeType) ? body.mimeType : "application/octet-stream";
  if (mimeType === "application/octet-stream") return sendJson(res, 400, { error: "Use a PDF, JPG, or PNG file." });

  const analyzeUrl = new URL(`${endpoint}/documentintelligence/documentModels/${model}:analyze?_overload=analyzeDocument&api-version=2024-11-30`);
  const analyze = await fetch(analyzeUrl, {
    method: "POST",
    headers: { "Ocp-Apim-Subscription-Key": key, "Content-Type": mimeType },
    body: file,
    signal: AbortSignal.timeout(30_000),
  });
  if (!analyze.ok) {
    const details = await analyze.json().catch(() => ({}));
    const message = details.error?.message || `Azure analysis request failed (${analyze.status}).`;
    return sendJson(res, analyze.status === 429 ? 429 : 502, { error: message });
  }
  const operationLocation = analyze.headers.get("operation-location");
  if (!operationLocation) return sendJson(res, 502, { error: "Azure accepted the scan but returned no result location." });
  let operationUrl;
  try { operationUrl = new URL(operationLocation); } catch { return sendJson(res, 502, { error: "Azure returned an invalid analysis result location." }); }
  if (operationUrl.origin !== azureEndpoint.origin) return sendJson(res, 502, { error: "Azure returned an unexpected analysis result location." });

  let result;
  for (let attempt = 0; attempt < 30; attempt++) {
    await new Promise(resolve => setTimeout(resolve, 1000));
    const poll = await fetch(operationUrl, {
      headers: { "Ocp-Apim-Subscription-Key": key },
      signal: AbortSignal.timeout(10_000),
    });
    if (!poll.ok) return sendJson(res, 502, { error: `Could not retrieve Azure analysis results (${poll.status}).` });
    result = await poll.json();
    if (result.status === "succeeded") return sendJson(res, 200, { items: extractLineItems(result, model) });
    if (result.status === "failed") return sendJson(res, 422, { error: result.error?.message || "Azure could not read this document." });
  }
  return sendJson(res, 504, { error: "Azure analysis is taking too long. Please try again." });
}

const mimeTypes = {
  ".css": "text/css; charset=utf-8", ".html": "text/html; charset=utf-8", ".ico": "image/x-icon",
  ".js": "text/javascript; charset=utf-8", ".json": "application/json; charset=utf-8", ".png": "image/png",
  ".svg": "image/svg+xml", ".webp": "image/webp", ".woff": "font/woff", ".woff2": "font/woff2",
};

async function serveStatic(req, res, pathname) {
  const requested = decodeURIComponent(pathname === "/" ? "/index.html" : pathname);
  const filePath = path.resolve(dist, `.${requested}`);
  if (filePath !== dist && !filePath.startsWith(`${dist}${path.sep}`)) return sendJson(res, 400, { error: "Invalid path." });
  let target = filePath;
  try { if (!(await stat(target)).isFile()) target = path.join(dist, "index.html"); }
  catch { target = path.join(dist, "index.html"); }
  try {
    const data = await readFile(target);
    res.writeHead(200, { "Content-Type": mimeTypes[path.extname(target)] || "application/octet-stream", "X-Content-Type-Options": "nosniff" });
    res.end(data);
  } catch { res.writeHead(404); res.end("Not found"); }
}

const port = Number(process.env.PORT || 3000);
createServer(async (req, res) => {
  try {
    const requestUrl = new URL(req.url || "/", "http://localhost");
    if (requestUrl.pathname === "/api/invoice-scan" && req.method === "POST") return await scanInvoice(req, res);
    if (requestUrl.pathname.startsWith("/api/")) return sendJson(res, 404, { error: "API route not found." });
    if (req.method === "GET" || req.method === "HEAD") return await serveStatic(req, res, requestUrl.pathname);
    return sendJson(res, 405, { error: "Method not allowed." });
  } catch (error) {
    if (!res.headersSent) sendJson(res, error.status || 500, { error: error.status ? error.message : "The scan could not be completed. Please try again." });
    else res.end();
  }
}).listen(port, "0.0.0.0", () => console.log(`App server listening on ${port}`));
