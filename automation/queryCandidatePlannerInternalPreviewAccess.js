const {
  getQueryCandidatePlannerInternalPreviewConfig,
} = require("./queryCandidatePlannerInternalPreviewConfig");

const ACCESS_VERSION = "query_candidate_planner_internal_preview_access_v1";
const TOKEN_HEADER = "x-beebee-internal-preview-token";

function headerValue(req, name) {
  if (typeof req?.get === "function") return req.get(name) || "";
  const headers = req?.headers || {};
  return headers[name] || headers[name.toLowerCase()] || "";
}

function notFound(res) {
  return res.status(404).json({
    ok: false,
    code: "NOT_FOUND",
    error: "요청한 리소스를 찾을 수 없습니다.",
  });
}

function forbidden(res) {
  return res.status(403).json({
    ok: false,
    code: "INTERNAL_PREVIEW_ACCESS_DENIED",
    error: "내부 미리보기 접근 권한이 없습니다.",
  });
}

function requireQueryCandidatePlannerInternalPreviewAccess(req, res, next) {
  const config = getQueryCandidatePlannerInternalPreviewConfig();
  if (!config.enabled) return notFound(res);

  const token = headerValue(req, TOKEN_HEADER);
  if (!config.verifyToken(token)) return forbidden(res);

  res.locals = res.locals || {};
  res.locals.queryCandidatePlannerInternalPreviewAccess = Object.freeze({
    version: ACCESS_VERSION,
    allowed: true,
    tokenIncluded: false,
    tokenHashIncluded: false,
  });
  return next();
}

module.exports = Object.freeze({
  ACCESS_VERSION,
  TOKEN_HEADER,
  headerValue,
  requireQueryCandidatePlannerInternalPreviewAccess,
});
