const {
  CANDIDATE_GROUPS,
  normalizeText,
  sha256,
  classifyObservedCandidate,
} = require("./queryCandidateObservation");

const QUERY_CANDIDATE_CONTRACT_VERSION = "query_candidate_contract_v1";
const QUERY_CANDIDATE_ITEM_VERSION = "query_candidate_item_v1";

const CANDIDATE_STATUS = Object.freeze([
  "UNASSESSED",
  "READY",
  "CONDITIONAL",
  "UNSUPPORTED",
  "REJECTED",
]);

const CANDIDATE_TYPE = Object.freeze([
  "ANALYSIS_RECIPE",
  "BUSINESS_TEMPLATE",
  "DASHBOARD",
  "CATEGORY",
  "MULTI_SOURCE",
  "UNKNOWN",
]);

const CANDIDATE_VISIBILITY = Object.freeze([
  "RECOMMENDED",
  "VISIBLE",
  "HIDDEN",
]);

const OBSERVED_CLASS = Object.freeze([
  "ELIGIBLE",
  "CONDITIONAL",
  "HIDDEN",
  "UNKNOWN",
]);

const EVIDENCE_TYPE = Object.freeze([
  "SOURCE_TABLE",
  "SOURCE_COLUMN",
  "SEMANTIC_ROLE",
  "CAPABILITY",
  "RECIPE_BINDING",
  "QUERY_PATH",
  "OTHER",
]);

const RISK_LEVEL = Object.freeze(["INFO", "WARNING", "BLOCKING"]);

const ALL_GROUPS = Object.freeze([
  ...CANDIDATE_GROUPS,
  "uiRecommendedCandidates",
]);

const GROUP_PRIORITY = Object.freeze([
  "uiRecommendedCandidates",
  "businessTemplateCandidates",
  "analysisRecipeCandidates",
  "dashboardCandidates",
  "categoryCandidates",
  "multiSourceCandidates",
  "topCandidates",
  "secondaryCandidates",
]);

function asArray(value) {
  if (Array.isArray(value)) return value.filter((item) => item != null);
  if (value == null || value === "") return [];
  return [value];
}

function unique(values = []) {
  const seen = new Set();
  const result = [];
  for (const value of asArray(values)) {
    const normalized = normalizeText(value);
    if (!normalized || seen.has(normalized)) continue;
    seen.add(normalized);
    result.push(normalized);
  }
  return result;
}

function firstText(values = []) {
  for (const value of values) {
    const normalized = normalizeText(value);
    if (normalized) return normalized;
  }
  return "";
}

function finiteOrNull(value) {
  if (value == null || value === "") return null;
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

function normalizeScore(value) {
  const number = finiteOrNull(value);
  if (number == null) return null;
  return Number(Math.min(100, Math.max(0, number)).toFixed(6));
}

function normalizeConfidence(value) {
  const number = finiteOrNull(value);
  if (number == null) return null;
  const normalized = number > 1 && number <= 100 ? number / 100 : number;
  return Number(Math.min(1, Math.max(0, normalized)).toFixed(6));
}

function positiveIntegerOrNull(value) {
  const number = Number(value);
  if (!Number.isInteger(number) || number < 1) return null;
  return number;
}

function candidateIdOf(candidate = {}, index = 0) {
  return normalizeText(
    candidate.candidateId ||
      candidate.id ||
      candidate.templateId ||
      candidate.recipeId ||
      `${candidate.candidateType || candidate.type || "candidate"}_${index + 1}`,
  );
}

function normalizeCandidateType(value = "") {
  const key = normalizeText(value)
    .replace(/[\s-]+/g, "_")
    .toUpperCase();
  const aliases = {
    ANALYSISRECIPE: "ANALYSIS_RECIPE",
    ANALYSIS_RECIPE: "ANALYSIS_RECIPE",
    BUSINESSTEMPLATE: "BUSINESS_TEMPLATE",
    BUSINESS_TEMPLATE: "BUSINESS_TEMPLATE",
    DASHBOARD: "DASHBOARD",
    CATEGORY: "CATEGORY",
    MULTISOURCE: "MULTI_SOURCE",
    MULTI_SOURCE: "MULTI_SOURCE",
  };
  return aliases[key] || "UNKNOWN";
}

function observedClassOf(value = "") {
  const key = normalizeText(value).toUpperCase();
  return OBSERVED_CLASS.includes(key) ? key : "UNKNOWN";
}

function groupPayload(payload = {}, groupName = "") {
  if (groupName === "uiRecommendedCandidates") {
    return asArray(payload.candidateUiPayload?.recommendedCandidates);
  }
  return asArray(payload[groupName]);
}

function collectCandidateRecords(payload = {}) {
  const byId = new Map();
  for (const groupName of GROUP_PRIORITY) {
    groupPayload(payload, groupName).forEach((candidate, index) => {
      if (!candidate || typeof candidate !== "object") return;
      const candidateId = candidateIdOf(candidate, index);
      if (!candidateId) return;
      if (!byId.has(candidateId)) byId.set(candidateId, []);
      byId.get(candidateId).push({ candidate, groupName, index });
    });
  }
  return byId;
}

function observationCandidateMap(observation = {}) {
  const result = new Map();
  for (const candidate of asArray(
    observation.candidateObservation?.candidates,
  )) {
    const candidateId = normalizeText(candidate.candidateId);
    if (candidateId) result.set(candidateId, candidate);
  }
  return result;
}

function numericCandidates(records = [], selectors = []) {
  const result = [];
  for (const { candidate } of records) {
    for (const selector of selectors) {
      const value = selector(candidate);
      const number = finiteOrNull(value);
      if (number != null) result.push(number);
    }
  }
  return result;
}

function minPositive(values = []) {
  const filtered = values.filter(
    (value) => Number.isFinite(value) && value >= 1,
  );
  return filtered.length ? Math.min(...filtered) : null;
}

function maxFinite(values = []) {
  const filtered = values.filter(Number.isFinite);
  return filtered.length ? Math.max(...filtered) : null;
}

function extractEvidence(
  records = [],
  recipeIds = [],
  sourceTableIds = [],
  sourceSheetNames = [],
) {
  const evidence = [];
  const keys = new Set();

  function add(item) {
    const normalized = {
      evidenceId: normalizeText(item.evidenceId || ""),
      type: EVIDENCE_TYPE.includes(item.type) ? item.type : "OTHER",
      path: normalizeText(item.path || ""),
      tableId: normalizeText(item.tableId || ""),
      sheetName: normalizeText(item.sheetName || ""),
      columnId: normalizeText(item.columnId || ""),
      columnName: normalizeText(item.columnName || ""),
      role: normalizeText(item.role || ""),
      value: item.value == null ? "" : normalizeText(item.value),
      source: normalizeText(item.source || ""),
    };
    const key = JSON.stringify(normalized);
    if (keys.has(key)) return;
    keys.add(key);
    normalized.evidenceId =
      normalized.evidenceId ||
      `evidence_${String(evidence.length + 1).padStart(3, "0")}`;
    evidence.push(normalized);
  }

  for (const { candidate, groupName } of records) {
    for (const raw of asArray(candidate.evidence)) {
      if (typeof raw === "string") {
        add({ type: "OTHER", value: raw, source: groupName });
        continue;
      }
      if (!raw || typeof raw !== "object") continue;
      add({
        evidenceId: raw.evidenceId || raw.id,
        type: normalizeText(raw.type || raw.kind).toUpperCase(),
        path: raw.path || raw.queryPath,
        tableId: raw.tableId || raw.sourceTableId,
        sheetName: raw.sheetName || raw.sourceSheetName,
        columnId: raw.columnId,
        columnName: raw.columnName || raw.header,
        role: raw.role || raw.semanticRole,
        value: raw.value || raw.message,
        source: raw.source || groupName,
      });
    }
  }

  sourceTableIds.forEach((tableId, index) =>
    add({
      type: "SOURCE_TABLE",
      tableId,
      sheetName: sourceSheetNames[index] || "",
      source: "legacy_candidate_reference",
    }),
  );
  sourceSheetNames
    .filter((_, index) => !sourceTableIds[index])
    .forEach((sheetName) =>
      add({
        type: "SOURCE_TABLE",
        sheetName,
        source: "legacy_candidate_reference",
      }),
    );
  recipeIds.forEach((recipeId) =>
    add({
      type: "RECIPE_BINDING",
      value: recipeId,
      source: "legacy_candidate_reference",
    }),
  );

  return evidence;
}

function extractRisks(records = []) {
  const rows = [];
  const keys = new Set();
  for (const { candidate, groupName } of records) {
    for (const raw of asArray(candidate.risks || candidate.warnings)) {
      const risk =
        typeof raw === "string"
          ? { code: "legacy_warning", level: "WARNING", message: raw }
          : {
              code: raw?.code || raw?.reasonCode || "legacy_warning",
              level: normalizeText(
                raw?.level || raw?.severity || "WARNING",
              ).toUpperCase(),
              message: raw?.message || raw?.reason || "",
            };
      const normalized = {
        code: normalizeText(risk.code),
        level: RISK_LEVEL.includes(risk.level) ? risk.level : "WARNING",
        message: normalizeText(risk.message),
        source: groupName,
      };
      const key = JSON.stringify(normalized);
      if (!normalized.code && !normalized.message) continue;
      if (keys.has(key)) continue;
      keys.add(key);
      rows.push(normalized);
    }
  }
  return rows;
}

function buildCandidateItem({
  candidateId,
  records,
  observedCandidate = {},
  payload = {},
} = {}) {
  const candidates = records.map((record) => record.candidate);
  const sourceGroups = unique(records.map((record) => record.groupName));
  const recipeIds = unique(
    candidates.flatMap((candidate) => [
      ...asArray(candidate.recipeIds),
      candidate.recipeId,
    ]),
  );
  const outputTypes = unique(
    candidates.flatMap((candidate) => [
      ...asArray(candidate.outputTypes),
      candidate.outputType,
    ]),
  );
  const sourceTableIds = unique(
    candidates.flatMap((candidate) => [
      ...asArray(candidate.sourceTableIds),
      candidate.sourceTableId,
    ]),
  );
  const sourceSheetNames = unique(
    candidates.flatMap((candidate) => [
      ...asArray(candidate.sourceSheetNames),
      candidate.sourceSheetName,
    ]),
  );
  const reasonCodes = unique(
    candidates.flatMap((candidate) => asArray(candidate.reasonCodes)),
  );
  const requiredCapabilities = unique(
    candidates.flatMap((candidate) =>
      asArray(
        candidate.requiredCapabilities ||
          candidate.requirements?.requiredCapabilities,
      ),
    ),
  );
  const missingRequirements = unique(
    candidates.flatMap((candidate) =>
      asArray(
        candidate.missingRequirements ||
          candidate.requirements?.missingRequirements,
      ).map((item) =>
        typeof item === "string" ? item : item?.message || item?.code || "",
      ),
    ),
  );

  const observedClass = observedClassOf(
    observedCandidate.observedClass ||
      classifyObservedCandidate(candidates[0] || {}),
  );
  const recommendedIds = new Set(
    groupPayload(payload, "uiRecommendedCandidates").map(candidateIdOf),
  );
  const visibility = recommendedIds.has(candidateId)
    ? "RECOMMENDED"
    : observedClass === "HIDDEN"
      ? "HIDDEN"
      : "VISIBLE";

  const type = firstText([
    ...candidates.map((candidate) => candidate.candidateType),
    ...candidates.map((candidate) => candidate.recipeType),
    ...candidates.map((candidate) => candidate.type),
    observedCandidate.candidateType,
  ]);
  const title = firstText([
    ...candidates.map((candidate) => candidate.title),
    ...candidates.map((candidate) => candidate.label),
    ...candidates.map((candidate) => candidate.name),
    observedCandidate.title,
    candidateId,
  ]);
  const templateId = firstText([
    ...candidates.map((candidate) => candidate.templateId),
    observedCandidate.templateId,
  ]);
  const score = normalizeScore(
    maxFinite(
      numericCandidates(records, [
        (candidate) => candidate.rankScore,
        (candidate) => candidate.score,
        (candidate) => candidate.score?.total,
        (candidate) => candidate.score?.finalScore,
        (candidate) => candidate.score?.value,
      ]),
    ),
  );
  const confidence = normalizeConfidence(
    maxFinite(
      numericCandidates(records, [
        (candidate) => candidate.confidence,
        (candidate) => candidate.score?.confidence,
      ]),
    ),
  );
  const rank = positiveIntegerOrNull(
    minPositive(
      numericCandidates(records, [
        (candidate) => candidate.rank,
        (candidate) => candidate.uiRank,
      ]),
    ),
  );
  const reasonSummary = firstText([
    ...candidates.map((candidate) => candidate.reason),
    ...candidates.map((candidate) => candidate.explanation),
    ...candidates.map((candidate) => candidate.rationale),
    ...candidates.map((candidate) => candidate.matchingReason),
  ]);

  return {
    version: QUERY_CANDIDATE_ITEM_VERSION,
    candidateId,
    recipeId: recipeIds[0] || "",
    recipeIds,
    templateId,
    candidateType: normalizeCandidateType(type),
    title,
    outputType: outputTypes[0] || "",
    outputTypes,
    status: "UNASSESSED",
    observedClass,
    visibility,
    confidence,
    score,
    rank,
    reason: {
      summary: reasonSummary,
      codes: reasonCodes,
      source: reasonSummary || reasonCodes.length ? "legacy_candidate" : "",
    },
    evidence: extractEvidence(
      records,
      recipeIds,
      sourceTableIds,
      sourceSheetNames,
    ),
    requiredCapabilities,
    missingRequirements,
    risks: extractRisks(records),
    sourceTableIds,
    sourceSheetNames,
    provenance: {
      sourceGroups,
      sourceCandidateContractVersion: normalizeText(
        payload.candidateContract?.version ||
          payload.candidateGeneration?.candidateContract?.version ||
          "",
      ),
      sourceCandidateScoringVersion: normalizeText(
        payload.candidateScoring?.version ||
          payload.candidateGeneration?.candidateScoring?.version ||
          "",
      ),
      sourceCandidateUiPayloadVersion: normalizeText(
        payload.candidateUiPayload?.version || "",
      ),
    },
  };
}

function sortCandidates(candidates = []) {
  return [...candidates].sort((left, right) => {
    const leftRank = left.rank == null ? Number.MAX_SAFE_INTEGER : left.rank;
    const rightRank = right.rank == null ? Number.MAX_SAFE_INTEGER : right.rank;
    if (leftRank !== rightRank) return leftRank - rightRank;
    const leftScore = left.score == null ? -1 : left.score;
    const rightScore = right.score == null ? -1 : right.score;
    if (leftScore !== rightScore) return rightScore - leftScore;
    return left.candidateId.localeCompare(right.candidateId, "ko");
  });
}

function buildQueryCandidateContract({
  observation = {},
  candidatePayload = {},
} = {}) {
  const recordsById = collectCandidateRecords(candidatePayload);
  const observedById = observationCandidateMap(observation);

  for (const [candidateId, observedCandidate] of observedById.entries()) {
    if (!recordsById.has(candidateId)) {
      recordsById.set(candidateId, [
        {
          candidate: observedCandidate,
          groupName: observedCandidate.groupName || "observation",
          index: 0,
        },
      ]);
    }
  }

  const candidates = sortCandidates(
    [...recordsById.entries()].map(([candidateId, records]) =>
      buildCandidateItem({
        candidateId,
        records,
        observedCandidate: observedById.get(candidateId) || {},
        payload: candidatePayload,
      }),
    ),
  );

  const contract = {
    version: QUERY_CANDIDATE_CONTRACT_VERSION,
    itemVersion: QUERY_CANDIDATE_ITEM_VERSION,
    source: {
      caseId: normalizeText(observation.caseId || ""),
      fileName: normalizeText(observation.fileName || ""),
      observationVersion: normalizeText(observation.version || ""),
      observationSha256: normalizeText(observation.observationSha256 || ""),
      queryJsonSha256: normalizeText(observation.source?.queryJsonSha256 || ""),
      candidatePayloadSha256: normalizeText(
        observation.source?.candidatePayloadSha256 || "",
      ),
    },
    counts: {
      total: candidates.length,
      unassessed: candidates.filter(
        (candidate) => candidate.status === "UNASSESSED",
      ).length,
      ready: 0,
      conditional: 0,
      unsupported: 0,
      rejected: 0,
    },
    candidates,
  };
  contract.contractSha256 = sha256({ ...contract, contractSha256: undefined });
  return contract;
}

function issue(path, code, message) {
  return { path, code, message };
}

function validateCandidate(candidate = {}, index = 0) {
  const path = `candidates[${index}]`;
  const errors = [];
  const warnings = [];

  if (candidate.version !== QUERY_CANDIDATE_ITEM_VERSION) {
    errors.push(
      issue(
        `${path}.version`,
        "invalid_item_version",
        `version은 ${QUERY_CANDIDATE_ITEM_VERSION}이어야 합니다.`,
      ),
    );
  }
  if (!normalizeText(candidate.candidateId)) {
    errors.push(
      issue(`${path}.candidateId`, "required", "candidateId가 필요합니다."),
    );
  } else if (/\s|[\\/]/u.test(candidate.candidateId)) {
    errors.push(
      issue(
        `${path}.candidateId`,
        "invalid_candidate_id",
        "candidateId에는 공백 또는 경로 구분자를 사용할 수 없습니다.",
      ),
    );
  }
  if (!normalizeText(candidate.title)) {
    errors.push(issue(`${path}.title`, "required", "title이 필요합니다."));
  }
  if (!CANDIDATE_TYPE.includes(candidate.candidateType)) {
    errors.push(
      issue(
        `${path}.candidateType`,
        "invalid_enum",
        "candidateType이 유효하지 않습니다.",
      ),
    );
  }
  if (!CANDIDATE_STATUS.includes(candidate.status)) {
    errors.push(
      issue(`${path}.status`, "invalid_enum", "status가 유효하지 않습니다."),
    );
  }
  if (!OBSERVED_CLASS.includes(candidate.observedClass)) {
    errors.push(
      issue(
        `${path}.observedClass`,
        "invalid_enum",
        "observedClass가 유효하지 않습니다.",
      ),
    );
  }
  if (!CANDIDATE_VISIBILITY.includes(candidate.visibility)) {
    errors.push(
      issue(
        `${path}.visibility`,
        "invalid_enum",
        "visibility가 유효하지 않습니다.",
      ),
    );
  }
  if (!Array.isArray(candidate.recipeIds)) {
    errors.push(
      issue(
        `${path}.recipeIds`,
        "invalid_type",
        "recipeIds는 배열이어야 합니다.",
      ),
    );
  }
  if (!Array.isArray(candidate.outputTypes)) {
    errors.push(
      issue(
        `${path}.outputTypes`,
        "invalid_type",
        "outputTypes는 배열이어야 합니다.",
      ),
    );
  }
  if (!Array.isArray(candidate.evidence)) {
    errors.push(
      issue(
        `${path}.evidence`,
        "invalid_type",
        "evidence는 배열이어야 합니다.",
      ),
    );
  }
  if (!Array.isArray(candidate.requiredCapabilities)) {
    errors.push(
      issue(
        `${path}.requiredCapabilities`,
        "invalid_type",
        "requiredCapabilities는 배열이어야 합니다.",
      ),
    );
  }
  if (!Array.isArray(candidate.missingRequirements)) {
    errors.push(
      issue(
        `${path}.missingRequirements`,
        "invalid_type",
        "missingRequirements는 배열이어야 합니다.",
      ),
    );
  }
  if (!Array.isArray(candidate.risks)) {
    errors.push(
      issue(`${path}.risks`, "invalid_type", "risks는 배열이어야 합니다."),
    );
  }
  if (
    !candidate.reason ||
    typeof candidate.reason !== "object" ||
    Array.isArray(candidate.reason)
  ) {
    errors.push(
      issue(`${path}.reason`, "invalid_type", "reason은 객체여야 합니다."),
    );
  } else {
    if (!Array.isArray(candidate.reason.codes)) {
      errors.push(
        issue(
          `${path}.reason.codes`,
          "invalid_type",
          "reason.codes는 배열이어야 합니다.",
        ),
      );
    }
    if (typeof candidate.reason.summary !== "string") {
      errors.push(
        issue(
          `${path}.reason.summary`,
          "invalid_type",
          "reason.summary는 문자열이어야 합니다.",
        ),
      );
    }
    if (typeof candidate.reason.source !== "string") {
      errors.push(
        issue(
          `${path}.reason.source`,
          "invalid_type",
          "reason.source는 문자열이어야 합니다.",
        ),
      );
    }
  }
  if (
    normalizeText(candidate.recipeId) &&
    !asArray(candidate.recipeIds).includes(candidate.recipeId)
  ) {
    errors.push(
      issue(
        `${path}.recipeId`,
        "recipe_id_not_in_recipe_ids",
        "recipeId는 recipeIds에 포함되어야 합니다.",
      ),
    );
  }
  if (
    normalizeText(candidate.outputType) &&
    !asArray(candidate.outputTypes).includes(candidate.outputType)
  ) {
    errors.push(
      issue(
        `${path}.outputType`,
        "output_type_not_in_output_types",
        "outputType은 outputTypes에 포함되어야 합니다.",
      ),
    );
  }

  if (candidate.confidence != null) {
    if (
      !Number.isFinite(candidate.confidence) ||
      candidate.confidence < 0 ||
      candidate.confidence > 1
    ) {
      errors.push(
        issue(
          `${path}.confidence`,
          "out_of_range",
          "confidence는 null 또는 0~1 범위여야 합니다.",
        ),
      );
    }
  }
  if (candidate.score != null) {
    if (
      !Number.isFinite(candidate.score) ||
      candidate.score < 0 ||
      candidate.score > 100
    ) {
      errors.push(
        issue(
          `${path}.score`,
          "out_of_range",
          "score는 null 또는 0~100 범위여야 합니다.",
        ),
      );
    }
  }
  if (
    candidate.rank != null &&
    (!Number.isInteger(candidate.rank) || candidate.rank < 1)
  ) {
    errors.push(
      issue(
        `${path}.rank`,
        "invalid_rank",
        "rank는 null 또는 1 이상의 정수여야 합니다.",
      ),
    );
  }

  const evidence = asArray(candidate.evidence);
  evidence.forEach((item, evidenceIndex) => {
    if (!EVIDENCE_TYPE.includes(item?.type)) {
      errors.push(
        issue(
          `${path}.evidence[${evidenceIndex}].type`,
          "invalid_enum",
          "evidence type이 유효하지 않습니다.",
        ),
      );
    }
  });
  asArray(candidate.risks).forEach((item, riskIndex) => {
    if (!RISK_LEVEL.includes(item?.level)) {
      errors.push(
        issue(
          `${path}.risks[${riskIndex}].level`,
          "invalid_enum",
          "risk level이 유효하지 않습니다.",
        ),
      );
    }
  });

  if (candidate.status === "READY") {
    if (!normalizeText(candidate.recipeId)) {
      errors.push(
        issue(
          `${path}.recipeId`,
          "ready_requires_recipe",
          "READY 후보는 recipeId가 필요합니다.",
        ),
      );
    }
    if (candidate.confidence == null) {
      errors.push(
        issue(
          `${path}.confidence`,
          "ready_requires_confidence",
          "READY 후보는 confidence가 필요합니다.",
        ),
      );
    }
    if (!evidence.length) {
      errors.push(
        issue(
          `${path}.evidence`,
          "ready_requires_evidence",
          "READY 후보는 evidence가 필요합니다.",
        ),
      );
    }
    if (asArray(candidate.missingRequirements).length) {
      errors.push(
        issue(
          `${path}.missingRequirements`,
          "ready_cannot_have_missing_requirements",
          "READY 후보에는 missingRequirements가 없어야 합니다.",
        ),
      );
    }
  }

  if (
    candidate.status === "CONDITIONAL" &&
    !asArray(candidate.missingRequirements).length
  ) {
    errors.push(
      issue(
        `${path}.missingRequirements`,
        "conditional_requires_missing_requirement",
        "CONDITIONAL 후보는 missingRequirements가 필요합니다.",
      ),
    );
  }

  if (["UNSUPPORTED", "REJECTED"].includes(candidate.status)) {
    const hasReason =
      normalizeText(candidate.reason?.summary) ||
      asArray(candidate.reason?.codes).length;
    if (!hasReason) {
      errors.push(
        issue(
          `${path}.reason`,
          "terminal_status_requires_reason",
          `${candidate.status} 후보는 reason이 필요합니다.`,
        ),
      );
    }
  }

  if (
    candidate.status === "UNASSESSED" &&
    candidate.observedClass === "ELIGIBLE"
  ) {
    warnings.push(
      issue(
        `${path}.status`,
        "observed_eligible_not_ready",
        "관측상 eligible이지만 생성 가능성 검증 전이므로 READY로 간주하지 않습니다.",
      ),
    );
  }
  if (!normalizeText(candidate.recipeId)) {
    warnings.push(
      issue(
        `${path}.recipeId`,
        "recipe_not_bound",
        "대표 recipeId가 아직 연결되지 않았습니다.",
      ),
    );
  }

  return {
    candidateId: candidate.candidateId || "",
    valid: errors.length === 0,
    errors,
    warnings,
  };
}

function validateQueryCandidateContract(contract = {}) {
  const errors = [];
  const warnings = [];

  if (contract.version !== QUERY_CANDIDATE_CONTRACT_VERSION) {
    errors.push(
      issue(
        "version",
        "invalid_contract_version",
        `version은 ${QUERY_CANDIDATE_CONTRACT_VERSION}이어야 합니다.`,
      ),
    );
  }
  if (contract.itemVersion !== QUERY_CANDIDATE_ITEM_VERSION) {
    errors.push(
      issue(
        "itemVersion",
        "invalid_item_version",
        `itemVersion은 ${QUERY_CANDIDATE_ITEM_VERSION}이어야 합니다.`,
      ),
    );
  }
  if (
    !contract.source ||
    typeof contract.source !== "object" ||
    Array.isArray(contract.source)
  ) {
    errors.push(issue("source", "invalid_type", "source는 객체여야 합니다."));
  }
  if (
    !contract.counts ||
    typeof contract.counts !== "object" ||
    Array.isArray(contract.counts)
  ) {
    errors.push(issue("counts", "invalid_type", "counts는 객체여야 합니다."));
  }
  if (!Array.isArray(contract.candidates)) {
    errors.push(
      issue("candidates", "invalid_type", "candidates는 배열이어야 합니다."),
    );
  }
  if (!/^[a-f0-9]{64}$/.test(normalizeText(contract.contractSha256))) {
    errors.push(
      issue(
        "contractSha256",
        "invalid_hash",
        "contractSha256는 64자리 SHA-256이어야 합니다.",
      ),
    );
  }

  const candidateResults = asArray(contract.candidates).map(validateCandidate);
  candidateResults.forEach((result) => {
    errors.push(...result.errors);
    warnings.push(...result.warnings);
  });

  const ids = asArray(contract.candidates).map(
    (candidate) => candidate.candidateId,
  );
  const duplicateIds = ids.filter(
    (candidateId, index) => ids.indexOf(candidateId) !== index,
  );
  unique(duplicateIds).forEach((candidateId) =>
    errors.push(
      issue(
        "candidates",
        "duplicate_candidate_id",
        `candidateId가 중복되었습니다: ${candidateId}`,
      ),
    ),
  );

  const counts = contract.counts || {};
  const expectedCounts = {
    total: asArray(contract.candidates).length,
    unassessed: asArray(contract.candidates).filter(
      (candidate) => candidate.status === "UNASSESSED",
    ).length,
    ready: asArray(contract.candidates).filter(
      (candidate) => candidate.status === "READY",
    ).length,
    conditional: asArray(contract.candidates).filter(
      (candidate) => candidate.status === "CONDITIONAL",
    ).length,
    unsupported: asArray(contract.candidates).filter(
      (candidate) => candidate.status === "UNSUPPORTED",
    ).length,
    rejected: asArray(contract.candidates).filter(
      (candidate) => candidate.status === "REJECTED",
    ).length,
  };
  Object.entries(expectedCounts).forEach(([key, expected]) => {
    if (counts[key] !== expected) {
      errors.push(
        issue(
          `counts.${key}`,
          "count_mismatch",
          `counts.${key}는 ${expected}이어야 합니다.`,
        ),
      );
    }
  });

  const expectedHash = sha256({ ...contract, contractSha256: undefined });
  if (contract.contractSha256 && contract.contractSha256 !== expectedHash) {
    errors.push(
      issue(
        "contractSha256",
        "hash_mismatch",
        "contractSha256가 내용과 일치하지 않습니다.",
      ),
    );
  }

  return {
    version: "query_candidate_contract_validation_v1",
    valid: errors.length === 0,
    errorCount: errors.length,
    warningCount: warnings.length,
    errors,
    warnings,
    candidateResults,
    expectedContractSha256: expectedHash,
  };
}

function assertValidQueryCandidateContract(contract = {}) {
  const validation = validateQueryCandidateContract(contract);
  if (!validation.valid) {
    const error = new Error(
      `query candidate contract validation failed: ${validation.errors
        .map((item) => `${item.path}:${item.code}`)
        .join(", ")}`,
    );
    error.validation = validation;
    throw error;
  }
  return validation;
}

function recountContract(contract = {}) {
  const candidates = asArray(contract.candidates);
  contract.counts = {
    total: candidates.length,
    unassessed: candidates.filter(
      (candidate) => candidate.status === "UNASSESSED",
    ).length,
    ready: candidates.filter((candidate) => candidate.status === "READY")
      .length,
    conditional: candidates.filter(
      (candidate) => candidate.status === "CONDITIONAL",
    ).length,
    unsupported: candidates.filter(
      (candidate) => candidate.status === "UNSUPPORTED",
    ).length,
    rejected: candidates.filter((candidate) => candidate.status === "REJECTED")
      .length,
  };
  contract.contractSha256 = sha256({ ...contract, contractSha256: undefined });
  return contract;
}

function applyCandidateAssessment(candidate = {}, assessment = {}) {
  const next = JSON.parse(JSON.stringify(candidate));
  if (assessment.status != null)
    next.status = normalizeText(assessment.status).toUpperCase();
  if (assessment.recipeId != null)
    next.recipeId = normalizeText(assessment.recipeId);
  if (assessment.recipeIds != null)
    next.recipeIds = unique(assessment.recipeIds);
  if (assessment.confidence !== undefined)
    next.confidence = normalizeConfidence(assessment.confidence);
  if (assessment.score !== undefined)
    next.score = normalizeScore(assessment.score);
  if (assessment.reason != null) {
    next.reason = {
      summary: normalizeText(assessment.reason.summary || assessment.reason),
      codes: unique(assessment.reason.codes),
      source: normalizeText(assessment.reason.source || "assessment"),
    };
  }
  if (assessment.evidence != null) next.evidence = asArray(assessment.evidence);
  if (assessment.requiredCapabilities != null) {
    next.requiredCapabilities = unique(assessment.requiredCapabilities);
  }
  if (assessment.missingRequirements != null) {
    next.missingRequirements = unique(assessment.missingRequirements);
  }
  if (assessment.risks != null) next.risks = asArray(assessment.risks);

  const validation = validateCandidate(next, 0);
  if (!validation.valid) {
    const error = new Error(
      `candidate assessment validation failed: ${validation.errors
        .map((item) => `${item.path}:${item.code}`)
        .join(", ")}`,
    );
    error.validation = validation;
    throw error;
  }
  return next;
}

module.exports = {
  QUERY_CANDIDATE_CONTRACT_VERSION,
  QUERY_CANDIDATE_ITEM_VERSION,
  CANDIDATE_STATUS,
  CANDIDATE_TYPE,
  CANDIDATE_VISIBILITY,
  OBSERVED_CLASS,
  EVIDENCE_TYPE,
  RISK_LEVEL,
  ALL_GROUPS,
  buildQueryCandidateContract,
  validateQueryCandidateContract,
  assertValidQueryCandidateContract,
  applyCandidateAssessment,
  recountContract,
};
