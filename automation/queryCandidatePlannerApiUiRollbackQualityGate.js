const crypto = require("crypto");
const {
  createQueryCandidatePlannerFeatureControl,
  SCOPES,
} = require("./queryCandidatePlannerFeatureControl");
const {
  createQueryCandidatePlannerApiShadowBoundary,
} = require("./queryCandidatePlannerApiShadowBoundary");
const {
  primaryResponseContractSha256,
} = require("./queryCandidatePlannerApiShadowRunner");
const {
  createQueryCandidatePlannerInternalPreviewStore,
} = require("./queryCandidatePlannerInternalPreviewStore");
const {
  buildQueryCandidatePlannerInternalPreviewHtml,
} = require("./queryCandidatePlannerInternalPreviewPage");
const {
  createQueryCandidatePlannerMutationBoundary,
  createQueryCandidatePlannerDownloadRetentionBoundary,
} = require("./queryCandidatePlannerFileLifecycleBoundary");
const {
  evaluateControlledProductionPromotionGate,
} = require("./queryCandidatePlannerControlledProductionPromotionGate");
const {
  controlledProductionMergeAdapter,
  ADAPTER_VERSION,
} = require("./queryCandidatePlannerControlledProductionMergeAdapter");

const QUALITY_GATE_VERSION =
  "query_candidate_planner_api_ui_rollback_quality_gate_v1";
const REPORT_VERSION =
  "query_candidate_planner_api_ui_rollback_quality_gate_report_v1";
const POLICY_VERSION =
  "api_ui_e2e_failure_isolation_immediate_rollback_policy_v1";

const DEFAULT_TIMEOUT_MS = 12;
const FIXED_NOW = "2026-08-04T12:00:00.000Z";
const FIXED_ROLLOUT_SALT = "beebeeai-patch14-5-quality-gate-salt-v1";

function canonicalize(value) {
  if (Array.isArray(value)) return value.map(canonicalize);
  if (!value || typeof value !== "object") return value;
  return Object.fromEntries(
    Object.keys(value)
      .sort()
      .map((key) => [key, canonicalize(value[key])]),
  );
}

function sha256(value) {
  const serialized =
    typeof value === "string" ? value : JSON.stringify(canonicalize(value));
  return crypto.createHash("sha256").update(serialized).digest("hex");
}

function canonicalClone(value) {
  if (value === undefined) return undefined;
  return JSON.parse(JSON.stringify(value));
}

function freeze(value) {
  if (Array.isArray(value)) {
    value.forEach(freeze);
    return Object.freeze(value);
  }
  if (value && typeof value === "object" && !Object.isFrozen(value)) {
    Object.values(value).forEach(freeze);
    Object.freeze(value);
  }
  return value;
}

function createResponseHarness() {
  const state = {
    statusCode: 200,
    headers: {},
    contentType: "",
    payload: undefined,
    kind: "",
  };
  const res = {
    locals: {},
    status(code) {
      state.statusCode = Number(code) || 200;
      return this;
    },
    set(name, value) {
      if (name && typeof name === "object") {
        Object.assign(state.headers, name);
      } else {
        state.headers[String(name)] = String(value);
      }
      return this;
    },
    type(value) {
      state.contentType = String(value || "");
      return this;
    },
    json(payload) {
      state.kind = "json";
      state.payload = payload;
      return this;
    },
    send(payload) {
      state.kind = "send";
      state.payload = payload;
      return this;
    },
  };
  return { res, state };
}

function sampleSubjectSha256() {
  return sha256("beebeeai:patch14.5:internal-canary-subject");
}

function createQualityGateFeatureControl(overrides = {}) {
  return createQueryCandidatePlannerFeatureControl({
    env: {
      QUERY_CANDIDATE_PLANNER_FEATURE_ENABLED: "1",
      QUERY_CANDIDATE_PLANNER_SHADOW_ENABLED: "1",
      QUERY_CANDIDATE_PLANNER_PROVIDER_ENABLED: "0",
      QUERY_CANDIDATE_PLANNER_CACHE_READ_ENABLED: "1",
      QUERY_CANDIDATE_PLANNER_CACHE_WRITE_ENABLED: "1",
      QUERY_CANDIDATE_PLANNER_PRODUCTION_ENABLED: "1",
      QUERY_CANDIDATE_PLANNER_PRODUCTION_CANDIDATE_MERGE_ENABLED: "1",
      QUERY_CANDIDATE_PLANNER_PRODUCTION_READY_ASSIGNMENT_ENABLED: "0",
      QUERY_CANDIDATE_PLANNER_PRODUCTION_ROUTE_ENABLED: "0",
      QUERY_CANDIDATE_PLANNER_KILL_SWITCH: "0",
      QUERY_CANDIDATE_PLANNER_PROVIDER_KILL_SWITCH: "0",
      QUERY_CANDIDATE_PLANNER_CACHE_KILL_SWITCH: "0",
      QUERY_CANDIDATE_PLANNER_PRODUCTION_KILL_SWITCH: "0",
      ...overrides,
    },
    now: () => new Date(FIXED_NOW),
  });
}

function readinessEvidence(overrides = {}) {
  return freeze({
    eligible: true,
    decision: "ELIGIBLE_FOR_CONTROLLED_PRODUCTION_PROMOTION_REVIEW",
    guardrails: {
      manualPromotionReviewRequired: true,
      failClosed: true,
      productionRouteAutoWired: false,
      productionCandidateMergeAllowed: false,
      productionReadyAssignmentAllowed: false,
    },
    ...overrides,
  });
}

function allowlistPromotionEnvironment(subjectSha256 = sampleSubjectSha256()) {
  return freeze({
    QUERY_CANDIDATE_PLANNER_PROMOTION_GATE_ENABLED: "1",
    QUERY_CANDIDATE_PLANNER_PROMOTION_AUDIENCE_MODE: "ALLOWLIST",
    QUERY_CANDIDATE_PLANNER_PROMOTION_ALLOWLIST_SHA256: subjectSha256,
    QUERY_CANDIDATE_PLANNER_PROMOTION_ROLLOUT_PERCENT: "0",
    QUERY_CANDIDATE_PLANNER_PROMOTION_ROLLOUT_SALT: FIXED_ROLLOUT_SALT,
  });
}

function samplePrimaryPayload() {
  return freeze({
    ok: true,
    source: "DETERMINISTIC_PRIMARY",
    fileHash: sha256("patch14.5-sample-file"),
    sheetStateSig: sha256("patch14.5-sheet-state"),
    queryTablesKey: "encrypted/query-json/patch14-5.json",
    normalizedQueryTables: [
      {
        tableId: "table_primary",
        rowCount: 12,
        columnCount: 3,
        isPrimary: true,
        columns: [
          { columnId: "date", header: "Date", type: "DATE" },
          { columnId: "amount", header: "Amount", type: "NUMBER" },
          { columnId: "category", header: "Category", type: "TEXT" },
        ],
      },
    ],
    topCandidates: [
      {
        candidateId: "candidate_summary",
        candidateType: "ANALYSIS_RECIPE",
        recipeId: "summary_recipe",
        rank: 1,
        title: "Summary",
      },
      {
        candidateId: "candidate_category",
        candidateType: "BUSINESS_TEMPLATE",
        recipeId: "category_recipe",
        rank: 2,
        title: "Category",
      },
    ],
    candidateUiPayload: {
      recommendedCandidates: [
        { candidateId: "candidate_summary", rank: 1, title: "Summary" },
        { candidateId: "candidate_category", rank: 2, title: "Category" },
      ],
    },
    analysisRecipeCandidates: [],
    businessTemplateCandidates: [],
    multiSourceCandidates: [],
    categoryCandidates: [],
    dashboardCandidates: [],
  });
}

function sampleShadowResolution() {
  return freeze({
    status: "SHADOW_COMPLETED",
    items: [
      {
        candidateId: "candidate_summary",
        candidateType: "ANALYSIS_RECIPE",
        recipeId: "summary_recipe",
        rank: 1,
        title: "Summary",
        status: "ACCEPTED",
      },
      {
        candidateId: "candidate_category",
        candidateType: "BUSINESS_TEMPLATE",
        recipeId: "category_recipe",
        rank: 2,
        title: "Category",
        status: "ACCEPTED",
      },
    ],
    counts: { accepted: 2 },
    providerCallCount: 0,
    policy: {
      productionCandidateMerge: false,
      productionReadyAssignment: false,
      productionRouteChanged: false,
    },
  });
}

function sampleRequest() {
  const queryJsonKey = "encrypted/query-json/patch14-5.json";
  return {
    method: "POST",
    path: "/analysis-candidates",
    originalUrl: "/api/automation/analysis-candidates",
    body: { queryTablesKey: queryJsonKey },
    params: {},
    query: {},
    user: {
      id: "quality-gate-user",
      uploadedFiles: [
        {
          originalName: "quality-gate.xlsx",
          queryJsonKey,
          storageKey: "encrypted/source/patch14-5.xlsx",
          fileHash: sha256("patch14.5-sample-file"),
          sheetStateSig: sha256("patch14.5-sheet-state"),
        },
      ],
    },
  };
}

async function runObservedApiCall({
  primaryPayload = samplePrimaryPayload(),
  request = sampleRequest(),
  featureControl = createQualityGateFeatureControl(),
  shadowRunner = async () => sampleShadowResolution(),
  timeoutMs = DEFAULT_TIMEOUT_MS,
  onObservation = () => {},
} = {}) {
  const beforeSha256 = primaryResponseContractSha256(primaryPayload);
  const { res, state } = createResponseHarness();
  const handler = async (_req, response) =>
    response.status(200).json(primaryPayload);
  const boundary = createQueryCandidatePlannerApiShadowBoundary({
    handler,
    featureControl,
    shadowRunner,
    timeoutMs,
    onObservation,
  });

  await boundary(request, res, () => {});
  const observation = res.locals.queryCandidatePlannerShadowTask
    ? await res.locals.queryCandidatePlannerShadowTask
    : null;
  const responsePayload = state.payload;
  const afterSha256 = primaryResponseContractSha256(primaryPayload);

  return freeze({
    statusCode: state.statusCode,
    responseKind: state.kind,
    responsePayloadSha256: sha256(responsePayload),
    primaryBeforeSha256: beforeSha256,
    primaryAfterSha256: afterSha256,
    primaryUnchanged: beforeSha256 === afterSha256,
    responseMatchesPrimary: sha256(responsePayload) === sha256(primaryPayload),
    observation,
  });
}

function evaluateInternalPreview(observation) {
  const store = createQueryCandidatePlannerInternalPreviewStore({
    maxEntries: 10,
    ttlMs: 60 * 60 * 1000,
    now: () => Date.parse(FIXED_NOW),
  });
  const entry = store.record(observation || {});
  const entries = store.list({ limit: 10 });
  const summary = store.summary();
  const html = buildQueryCandidatePlannerInternalPreviewHtml({
    nonce: "patch14-5-fixed-nonce",
  });

  const forbiddenActionFragments = [
    "/execute-analysis-candidate",
    "/query-execute",
    "/execute-business-template",
    "PRODUCTION_CANDIDATE_MERGE_ENABLED",
    'method="post"',
    "method='post'",
  ];
  const forbiddenMatches = forbiddenActionFragments.filter((fragment) =>
    html.includes(fragment),
  );

  return freeze({
    entryCount: entries.length,
    summaryTotal: summary.total,
    readOnly:
      entry.guardrails?.candidateExecutionAvailable === false &&
      entry.guardrails?.candidateSelectionAvailable === false &&
      entry.guardrails?.productionCandidateMerge === false,
    productionBadgePresent: html.includes("Production 반영 없음"),
    forbiddenActionMatches: forbiddenMatches,
    privacySafe:
      entry.privacy?.rawRowsIncluded === false &&
      entry.privacy?.fileNameIncluded === false &&
      entry.privacy?.tenantIdIncluded === false &&
      entry.privacy?.rawCandidatePayloadIncluded === false,
    persistence: summary.persistence,
  });
}

async function runCacheInvalidationFailureScenario() {
  const request = sampleRequest();
  request.params = { originalName: "quality-gate.xlsx" };
  const { res, state } = createResponseHarness();
  const boundary = createQueryCandidatePlannerMutationBoundary({
    action: "DELETE",
    handler: async (_req, response) =>
      response.status(200).json({ ok: true, deleted: true }),
    runtimeProvider: () => ({
      enabled: true,
      hierarchicalCache: {},
      cacheSecret: "test-only-secret-not-observed",
    }),
    invalidate: async () => {
      const error = new Error("injected cache invalidation failure");
      error.code = "CACHE_INVALIDATION_FAULT_INJECTED";
      throw error;
    },
    onObservation: () => {},
  });
  await boundary(request, res, () => {});
  return freeze({
    statusCode: state.statusCode,
    responseSucceeded: state.payload?.ok === true,
    cacheDisposition:
      res.locals.queryCandidatePlannerCacheLifecycleObservation
        ?.cacheDisposition || "",
    reason:
      res.locals.queryCandidatePlannerCacheLifecycleObservation?.reason || "",
  });
}

async function runDownloadRetentionScenario() {
  const request = sampleRequest();
  const { res, state } = createResponseHarness();
  const boundary = createQueryCandidatePlannerDownloadRetentionBoundary({
    action: "SOURCE_DOWNLOAD",
    handler: async (_req, response) =>
      response.status(200).send(Buffer.from("download-ok")),
    onObservation: () => {},
  });
  await boundary(request, res, () => {});
  return freeze({
    statusCode: state.statusCode,
    responseKind: state.kind,
    cacheDisposition:
      res.locals.queryCandidatePlannerCacheLifecycleObservation
        ?.cacheDisposition || "",
    reason:
      res.locals.queryCandidatePlannerCacheLifecycleObservation?.reason || "",
  });
}

function check(id, passed, details = {}) {
  return freeze({
    id,
    passed: passed === true,
    details,
  });
}

function safeObservationDetails(observation = {}) {
  return freeze({
    status: String(observation?.status || ""),
    reason: String(observation?.reason || ""),
    primaryResponseUnchanged: observation?.primaryResponseUnchanged !== false,
    providerCallCount: Number(observation?.shadow?.providerCallCount || 0),
    productionCandidateMerge:
      observation?.guardrails?.productionCandidateMerge === true,
    productionReadyAssignment:
      observation?.guardrails?.productionReadyAssignment === true,
    productionRouteChanged:
      observation?.guardrails?.productionRouteChanged === true,
  });
}

async function runQueryCandidatePlannerApiUiRollbackQualityGate({
  timeoutMs = DEFAULT_TIMEOUT_MS,
} = {}) {
  const primaryPayload = samplePrimaryPayload();
  const subjectSha256 = sampleSubjectSha256();
  const readinessGate = readinessEvidence();
  const featureControl = createQualityGateFeatureControl();

  const baseline = await runObservedApiCall({
    primaryPayload,
    request: sampleRequest(),
    featureControl,
    timeoutMs,
  });
  const preview = evaluateInternalPreview(baseline.observation);

  const defaultPromotionDecision = evaluateControlledProductionPromotionGate({
    env: {},
    featureControl,
    readinessGate,
    subjectSha256,
    adapterVersion: ADAPTER_VERSION,
  });
  const preRollbackPromotionDecision =
    evaluateControlledProductionPromotionGate({
      env: allowlistPromotionEnvironment(subjectSha256),
      featureControl,
      readinessGate,
      subjectSha256,
      adapterVersion: ADAPTER_VERSION,
    });
  const dryRun = controlledProductionMergeAdapter({
    primaryPayload,
    shadowResolution: sampleShadowResolution(),
    featureControl,
    readinessGate,
    promotionGateDecision: preRollbackPromotionDecision,
    apply: false,
  });

  const failureControl = createQualityGateFeatureControl();
  const shadowFailure = await runObservedApiCall({
    primaryPayload,
    request: sampleRequest(),
    featureControl: failureControl,
    shadowRunner: async () => {
      const error = new Error("injected shadow failure");
      error.code = "SHADOW_FAULT_INJECTED";
      throw error;
    },
    timeoutMs,
  });

  const timeoutControl = createQualityGateFeatureControl();
  const shadowTimeout = await runObservedApiCall({
    primaryPayload,
    request: sampleRequest(),
    featureControl: timeoutControl,
    shadowRunner: () => new Promise(() => {}),
    timeoutMs,
  });

  const recorderControl = createQualityGateFeatureControl();
  const previewRecorderFailure = await runObservedApiCall({
    primaryPayload,
    request: sampleRequest(),
    featureControl: recorderControl,
    timeoutMs,
    onObservation: () => {
      const error = new Error("injected preview recorder failure");
      error.code = "PREVIEW_RECORDER_FAULT_INJECTED";
      throw error;
    },
  });

  const cacheFailure = await runCacheInvalidationFailureScenario();
  const downloadRetention = await runDownloadRetentionScenario();

  const revisionBeforeRollback = featureControl.snapshot().runtimeRevision;
  featureControl.activateKillSwitch({
    scope: SCOPES.GLOBAL,
    reason: "PATCH14_5_IMMEDIATE_ROLLBACK_DRILL",
    actor: "QUALITY_GATE",
  });
  const revisionAfterRollback = featureControl.snapshot().runtimeRevision;

  const postRollbackApi = await runObservedApiCall({
    primaryPayload,
    request: sampleRequest(),
    featureControl,
    timeoutMs,
  });
  const postRollbackPromotionDecision =
    evaluateControlledProductionPromotionGate({
      env: allowlistPromotionEnvironment(subjectSha256),
      featureControl,
      readinessGate,
      subjectSha256,
      adapterVersion: ADAPTER_VERSION,
    });
  const postRollbackMerge = controlledProductionMergeAdapter({
    primaryPayload,
    shadowResolution: sampleShadowResolution(),
    featureControl,
    readinessGate,
    promotionGateDecision: postRollbackPromotionDecision,
    apply: true,
  });
  const secondRollbackSnapshot = featureControl.activateKillSwitch({
    scope: SCOPES.GLOBAL,
    reason: "PATCH14_5_IDEMPOTENT_ROLLBACK_DRILL",
    actor: "QUALITY_GATE",
  });

  const checks = [
    check(
      "API_PRIMARY_HTTP_CONTRACT",
      baseline.statusCode === 200 && baseline.responseKind === "json",
      {
        statusCode: baseline.statusCode,
        responseKind: baseline.responseKind,
      },
    ),
    check(
      "API_PRIMARY_RESPONSE_UNCHANGED",
      baseline.primaryUnchanged && baseline.responseMatchesPrimary,
      {
        primaryBeforeSha256: baseline.primaryBeforeSha256,
        primaryAfterSha256: baseline.primaryAfterSha256,
      },
    ),
    check(
      "SHADOW_OBSERVATION_COMPLETED",
      ["COMPLETED", "COMPLETED_SAFE"].includes(baseline.observation?.status),
      safeObservationDetails(baseline.observation),
    ),
    check(
      "SHADOW_COMPARATOR_AVAILABLE",
      Boolean(baseline.observation?.comparison?.verdict),
      {
        verdict: baseline.observation?.comparison?.verdict || "",
      },
    ),
    check(
      "INTERNAL_PREVIEW_MEMORY_ONLY",
      preview.entryCount === 1 &&
        preview.summaryTotal === 1 &&
        preview.persistence === "MEMORY_ONLY",
      preview,
    ),
    check(
      "INTERNAL_PREVIEW_READ_ONLY",
      preview.readOnly &&
        preview.productionBadgePresent &&
        preview.forbiddenActionMatches.length === 0,
      preview,
    ),
    check("INTERNAL_PREVIEW_PRIVACY", preview.privacySafe, preview),
    check(
      "PROMOTION_GATE_DEFAULT_BLOCKED",
      defaultPromotionDecision.allowed === false &&
        defaultPromotionDecision.decision === "BLOCK",
      {
        reason: defaultPromotionDecision.reason,
      },
    ),
    check(
      "CONTROLLED_ALLOWLIST_PATH_READY",
      preRollbackPromotionDecision.allowed === true &&
        preRollbackPromotionDecision.decision === "ALLOW",
      {
        reason: preRollbackPromotionDecision.reason,
        audiencePath: preRollbackPromotionDecision.audience?.path || "",
      },
    ),
    check(
      "MERGE_ADAPTER_DRY_RUN_ONLY",
      dryRun.status === "DRY_RUN_READY" &&
        dryRun.applied === false &&
        dryRun.mergedPayload === null,
      {
        status: dryRun.status,
        reason: dryRun.reason,
      },
    ),
    check(
      "SHADOW_FAILURE_ISOLATED",
      shadowFailure.statusCode === 200 &&
        shadowFailure.responseMatchesPrimary &&
        shadowFailure.observation?.status === "FAILED_SAFE",
      safeObservationDetails(shadowFailure.observation),
    ),
    check(
      "SHADOW_TIMEOUT_ISOLATED",
      shadowTimeout.statusCode === 200 &&
        shadowTimeout.responseMatchesPrimary &&
        shadowTimeout.observation?.status === "TIMEOUT_SAFE",
      safeObservationDetails(shadowTimeout.observation),
    ),
    check(
      "PREVIEW_RECORDER_FAILURE_ISOLATED",
      previewRecorderFailure.statusCode === 200 &&
        previewRecorderFailure.responseMatchesPrimary &&
        previewRecorderFailure.observation?.status === "BOUNDARY_FAILED_SAFE",
      safeObservationDetails(previewRecorderFailure.observation),
    ),
    check(
      "CACHE_INVALIDATION_FAILURE_ISOLATED",
      cacheFailure.statusCode === 200 &&
        cacheFailure.responseSucceeded &&
        cacheFailure.cacheDisposition === "INVALIDATION_FAILED_SAFE",
      cacheFailure,
    ),
    check(
      "DOWNLOAD_RETAINS_CACHE",
      downloadRetention.statusCode === 200 &&
        downloadRetention.cacheDisposition === "RETAINED",
      downloadRetention,
    ),
    check(
      "KILL_SWITCH_REVISION_IMMEDIATE",
      revisionAfterRollback === revisionBeforeRollback + 1,
      {
        revisionBeforeRollback,
        revisionAfterRollback,
      },
    ),
    check(
      "POST_ROLLBACK_SHADOW_BLOCKED",
      postRollbackApi.statusCode === 200 &&
        postRollbackApi.responseMatchesPrimary &&
        postRollbackApi.observation?.status === "BLOCKED" &&
        postRollbackApi.observation?.reason === "GLOBAL_KILL_SWITCH_ACTIVE",
      safeObservationDetails(postRollbackApi.observation),
    ),
    check(
      "POST_ROLLBACK_PROMOTION_BLOCKED",
      postRollbackPromotionDecision.allowed === false &&
        postRollbackPromotionDecision.reason === "GLOBAL_KILL_SWITCH_ACTIVE",
      {
        reason: postRollbackPromotionDecision.reason,
      },
    ),
    check(
      "POST_ROLLBACK_MERGE_BLOCKED",
      postRollbackMerge.status === "BLOCKED" &&
        postRollbackMerge.applied === false &&
        postRollbackMerge.mergedPayload === null,
      {
        status: postRollbackMerge.status,
        reason: postRollbackMerge.reason,
      },
    ),
    check(
      "ROLLBACK_IDEMPOTENT_FAIL_CLOSED",
      secondRollbackSnapshot.killSwitches.global === true &&
        secondRollbackSnapshot.runtimeRevision === revisionAfterRollback + 1,
      {
        runtimeRevision: secondRollbackSnapshot.runtimeRevision,
        globalKillSwitch: secondRollbackSnapshot.killSwitches.global,
      },
    ),
    check(
      "PROVIDER_CALL_COUNT_ZERO",
      Number(baseline.observation?.shadow?.providerCallCount || 0) === 0 &&
        Number(shadowFailure.observation?.shadow?.providerCallCount || 0) === 0,
      {
        baselineProviderCallCount: Number(
          baseline.observation?.shadow?.providerCallCount || 0,
        ),
        failureProviderCallCount: Number(
          shadowFailure.observation?.shadow?.providerCallCount || 0,
        ),
      },
    ),
    check(
      "PRODUCTION_GUARDRAILS_UNCHANGED",
      [
        baseline.observation,
        shadowFailure.observation,
        shadowTimeout.observation,
        postRollbackApi.observation,
      ].every(
        (observation) =>
          observation?.guardrails?.productionCandidateMerge !== true &&
          observation?.guardrails?.productionReadyAssignment !== true &&
          observation?.guardrails?.productionRouteChanged !== true,
      ),
      {
        productionCandidateMerge: false,
        productionReadyAssignment: false,
        productionRouteChanged: false,
      },
    ),
  ];

  const failedChecks = checks.filter((item) => !item.passed);
  const report = {
    version: REPORT_VERSION,
    qualityGateVersion: QUALITY_GATE_VERSION,
    policyVersion: POLICY_VERSION,
    decision: failedChecks.length === 0 ? "PASS" : "FAIL",
    passed: failedChecks.length === 0,
    failClosed: true,
    counts: {
      total: checks.length,
      passed: checks.length - failedChecks.length,
      failed: failedChecks.length,
    },
    checks,
    failedCheckIds: failedChecks.map((item) => item.id),
    rollback: {
      mechanism: "RUNTIME_GLOBAL_KILL_SWITCH",
      synchronousDecisionBoundary: true,
      primaryResponseAuthorityRetained: true,
      postRollbackShadowStatus: postRollbackApi.observation?.status || "",
      postRollbackPromotionReason: postRollbackPromotionDecision.reason,
      postRollbackMergeStatus: postRollbackMerge.status,
    },
    guardrails: {
      routeWired: false,
      controllerWired: false,
      generalUserUiChanged: false,
      primaryResponseMutation: false,
      responseHeaderMutation: false,
      responseStatusMutation: false,
      productionCandidateMergeApplied: false,
      productionReadyAssignment: false,
      productionRouteChanged: false,
      providerCalls: 0,
      railwayEnvironmentChanged: false,
      failClosed: true,
    },
    privacy: {
      rawRowsIncluded: false,
      fileNameIncluded: false,
      queryTablesKeyIncluded: false,
      tenantIdIncluded: false,
      rawCandidatePayloadIncluded: false,
      rawSubjectIncluded: false,
    },
  };
  report.reportSha256 = sha256(report);
  return freeze(report);
}

module.exports = Object.freeze({
  QUALITY_GATE_VERSION,
  REPORT_VERSION,
  POLICY_VERSION,
  DEFAULT_TIMEOUT_MS,
  FIXED_NOW,
  FIXED_ROLLOUT_SALT,
  sha256,
  canonicalClone,
  createResponseHarness,
  sampleSubjectSha256,
  createQualityGateFeatureControl,
  readinessEvidence,
  allowlistPromotionEnvironment,
  samplePrimaryPayload,
  sampleShadowResolution,
  sampleRequest,
  runObservedApiCall,
  evaluateInternalPreview,
  runCacheInvalidationFailureScenario,
  runDownloadRetentionScenario,
  runQueryCandidatePlannerApiUiRollbackQualityGate,
});
