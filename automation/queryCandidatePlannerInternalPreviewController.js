const crypto = require("crypto");
const {
  getQueryCandidatePlannerInternalPreviewConfig,
} = require("./queryCandidatePlannerInternalPreviewConfig");
const {
  getQueryCandidatePlannerInternalPreviewStore,
} = require("./queryCandidatePlannerInternalPreviewStore");
const {
  buildQueryCandidatePlannerInternalPreviewHtml,
} = require("./queryCandidatePlannerInternalPreviewPage");

const CONTROLLER_VERSION =
  "query_candidate_planner_internal_preview_controller_v1";

function noStoreHeaders(res) {
  res.set("Cache-Control", "no-store, max-age=0");
  res.set("Pragma", "no-cache");
  res.set("Expires", "0");
  res.set("X-Content-Type-Options", "nosniff");
  res.set("X-Frame-Options", "DENY");
  res.set("Referrer-Policy", "no-referrer");
}

function notFound(res) {
  return res.status(404).json({
    ok: false,
    code: "NOT_FOUND",
    error: "요청한 리소스를 찾을 수 없습니다.",
  });
}

function internalPreviewPage(req, res) {
  const config = getQueryCandidatePlannerInternalPreviewConfig();
  if (!config.enabled) return notFound(res);

  const nonce = crypto.randomBytes(18).toString("base64url");
  noStoreHeaders(res);
  res.set(
    "Content-Security-Policy",
    [
      "default-src 'none'",
      `style-src 'nonce-${nonce}'`,
      `script-src 'nonce-${nonce}'`,
      "connect-src 'self'",
      "img-src 'none'",
      "font-src 'none'",
      "object-src 'none'",
      "base-uri 'none'",
      "form-action 'none'",
      "frame-ancestors 'none'",
    ].join("; "),
  );
  return res
    .status(200)
    .type("html")
    .send(buildQueryCandidatePlannerInternalPreviewHtml({ nonce }));
}

function internalPreviewStatus(req, res) {
  const config = getQueryCandidatePlannerInternalPreviewConfig();
  if (!config.enabled) return notFound(res);
  noStoreHeaders(res);
  return res.json({
    ok: true,
    version: CONTROLLER_VERSION,
    preview: config.publicSnapshot(),
    store: getQueryCandidatePlannerInternalPreviewStore().summary(),
    guardrails: {
      readOnly: true,
      observationOnly: true,
      productionCandidateMerge: false,
      productionReadyAssignment: false,
      productionRouteChanged: false,
      candidateExecutionAvailable: false,
      candidateSelectionAvailable: false,
    },
  });
}

function internalPreviewObservations(req, res) {
  const config = getQueryCandidatePlannerInternalPreviewConfig();
  if (!config.enabled) return notFound(res);
  noStoreHeaders(res);
  const store = getQueryCandidatePlannerInternalPreviewStore();
  const entries = store.list({
    limit: req.query?.limit,
    status: req.query?.status,
    verdict: req.query?.verdict,
  });
  return res.json({
    ok: true,
    version: CONTROLLER_VERSION,
    readOnly: true,
    entries,
    summary: store.summary(),
    privacy: {
      rawRowsIncluded: false,
      sampleValuesIncluded: false,
      fileNameIncluded: false,
      originalFileNameIncluded: false,
      queryTablesKeyIncluded: false,
      tenantIdIncluded: false,
      emailIncluded: false,
      rawCandidatePayloadIncluded: false,
      rawIdentifiersIncluded: false,
    },
    productionCandidateMerge: false,
    productionReadyAssignment: false,
    productionRouteChanged: false,
  });
}

module.exports = Object.freeze({
  CONTROLLER_VERSION,
  noStoreHeaders,
  internalPreviewPage,
  internalPreviewStatus,
  internalPreviewObservations,
});
