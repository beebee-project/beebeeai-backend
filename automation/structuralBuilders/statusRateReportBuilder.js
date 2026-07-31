const {
  findColumnHeader,
  makeTemplateCandidate,
  makeTemplateSection,
  executeTemplateSections,
  getRows,
  getRowValue,
} = require("../businessTemplates/commonTemplateHelpers");

const STATUS_RATE_REPORT_VERSION =
  "status_rate_report_builder_v4_legacy_overview_compatibility";
const STATUS_COLUMN_SELECTION_VERSION =
  "status_column_selection_v2_semantic_evidence";
const STATUS_CLASSIFICATION_VERSION =
  "canonical_business_status_v2_1_exact_progress";
const STATUS_RATIO_CONTRACT_VERSION = "status_ratio_denominator_v2_all_rows";
const STATUS_OVERVIEW_COMPATIBILITY_VERSION =
  "status_overview_compatibility_v1_legacy_pending_rollup";
const LEGACY_PENDING_ROLLUP_CLASSES = Object.freeze([
  "pending",
  "incomplete",
  "delayed",
]);

const DEFAULT_STATUS_HINTS = [
  "조치상태",
  "계약상태",
  "이수상태",
  "합격상태",
  "참여상태",
  "배송상태",
  "발주상태",
  "검수상태",
  "참석상태",
  "참석여부",
  "이수여부",
  "완료여부",
  "처리상태",
  "진행상태",
  "승인상태",
  "정산상태",
  "신청상태",
  "문의상태",
  "재직상태",
  "장비상태",
  "자산상태",
  "보유상태",
  "취소여부",
  "처리결과",
  "결과",
  "상태",
  "status",
  "state",
  "result",
];

const DEFAULT_DATE_HINTS = [
  "일자",
  "날짜",
  "월",
  "연월",
  "기준월",
  "등록일",
  "신청일",
  "접수일",
  "요청일",
  "처리일",
  "완료일",
  "승인일",
  "발주일",
  "검수일",
  "계약일",
  "시작일",
  "종료일",
  "date",
  "month",
  "period",
];

const DEFAULT_METRIC_HINTS = [
  "금액",
  "비용",
  "사용금액",
  "승인금액",
  "지출금액",
  "구매금액",
  "취득금액",
  "발주금액",
  "계약금액",
  "집행금액",
  "정산금액",
  "수량",
  "건수",
  "승인건수",
  "점수",
  "평점",
  "amount",
  "cost",
  "price",
  "count",
  "score",
];

const DEFAULT_OWNER_HINTS = [
  "담당자",
  "사용자",
  "신청자",
  "요청자",
  "처리자",
  "검수자",
  "출장자",
  "성명",
  "이름",
  "직원명",
  "owner",
  "manager",
  "user",
  "name",
];

const DEFAULT_DEPARTMENT_HINTS = [
  "부서",
  "소속",
  "조직",
  "팀",
  "부서명",
  "소속부서",
  "담당부서",
  "기관",
  "department",
  "team",
  "organization",
];

const DEFAULT_CATEGORY_HINTS = [
  "유형",
  "구분",
  "분류",
  "카테고리",
  "항목",
  "비목",
  "세목",
  "품목",
  "물품",
  "장비",
  "자산",
  "문의유형",
  "지원분야",
  "전형",
  "회차",
  "업체",
  "거래처",
  "공급사",
  "가맹점",
  "행사명",
  "교육명",
  "과제명",
  "운송사",
  "category",
  "type",
  "item",
  "vendor",
];

const CANONICAL_STATUS_ORDER = Object.freeze([
  "completed",
  "pending",
  "incomplete",
  "cancelled",
  "delayed",
  "terminated",
  "other",
  "unknown",
]);

const CANONICAL_STATUS_LABELS = Object.freeze({
  completed: "완료·승인",
  pending: "진행·대기",
  incomplete: "미완료",
  cancelled: "취소·반려",
  delayed: "지연",
  terminated: "종료·중단",
  other: "기타",
  unknown: "미입력",
});

function normalizeText(value = "") {
  return String(value == null ? "" : value)
    .normalize("NFKC")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeStatusText(value = "") {
  return normalizeText(value)
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[()[\]{}"'`~!@#$%^&*+=|\\/:;,.?<>·ㆍ]/g, "")
    .trim();
}

function normalizeKey(value = "") {
  return normalizeText(value)
    .toLowerCase()
    .replace(/[\s_\-./\\|:;,'"‘’“”()[\]{}<>]+/g, "");
}

function includesAny(text = "", keywords = []) {
  const normalized = normalizeStatusText(text);
  return keywords.some((keyword) => {
    const target = normalizeStatusText(keyword);
    return target && normalized.includes(target);
  });
}

function classifyStatus(value = "", context = {}) {
  const text = normalizeStatusText(value);
  if (!text) return "unknown";

  /*
   * 부정·미완료는 긍정어보다 먼저 판정한다.
   * 예: 미이수에는 "이수", 불참에는 "참"이 포함된다.
   */
  if (
    includesAny(text, [
      "미이수",
      "불참",
      "결석",
      "미참석",
      "미출석",
      "미참여",
      "미완료",
      "미처리",
      "미검수",
      "미달",
      "누락",
      "미제출",
      "미응답",
      "incomplete",
      "absent",
      "notcompleted",
    ])
  ) {
    return "incomplete";
  }

  if (
    includesAny(text, [
      "취소",
      "반려",
      "거절",
      "불합격",
      "미승인",
      "불가",
      "탈락",
      "폐기",
      "cancel",
      "rejected",
      "declined",
      "failed",
      "inactive",
    ])
  ) {
    return "cancelled";
  }

  if (
    includesAny(text, [
      "중도종료",
      "중도중단",
      "조기종료",
      "참여종료",
      "중단",
      "퇴사",
      "탈퇴",
      "terminated",
      "stopped",
      "withdrawn",
    ]) ||
    /^(?:종료|해지)$/.test(text)
  ) {
    return "terminated";
  }

  if (
    includesAny(text, [
      "지연",
      "연체",
      "기한초과",
      "미납",
      "delayed",
      "late",
      "overdue",
    ])
  ) {
    return "delayed";
  }

  /*
   * 진행 상태는 완료 키워드보다 먼저 판정한다.
   * 예: 참여중에는 "참여", 검수중에는 "검수"가 포함된다.
   */
  if (
    text === "진행" ||
    includesAny(text, [
      "배송중",
      "검수중",
      "발주중",
      "조치중",
      "진행중",
      "처리중",
      "확인중",
      "참여중",
      "심사중",
      "검토중",
      "접수중",
      "대기",
      "예비",
      "신청",
      "접수",
      "요청",
      "예정",
      "검토",
      "보류",
      "부분검수",
      "pending",
      "progress",
      "open",
      "requested",
      "review",
      "waiting",
      "active",
    ]) ||
    /(?:업무|처리|배송|검수|발주|조치|진행|참여|심사|검토|접수)중$/.test(text)
  ) {
    return "pending";
  }

  if (
    includesAny(text, [
      "처리완료",
      "검수완료",
      "입고완료",
      "수령완료",
      "지급완료",
      "정산완료",
      "계약완료",
      "참석완료",
      "완료됨",
      "완료",
      "승인",
      "도착",
      "정상",
      "참석",
      "이수",
      "합격",
      "선정",
      "입고",
      "수령",
      "completed",
      "complete",
      "approved",
      "done",
      "closed",
      "success",
    ])
  ) {
    return "completed";
  }

  /*
   * context는 관측용으로만 보존하며 값 자체의 의미를 덮어쓰지 않는다.
   */
  void context;
  return "other";
}

function statusClassLabel(statusClass = "") {
  return CANONICAL_STATUS_LABELS[statusClass] || CANONICAL_STATUS_LABELS.other;
}

function columnHeader(column = {}, index = 0) {
  if (typeof column === "string") return normalizeText(column);
  return normalizeText(
    column.header ||
      column.originalHeader ||
      column.name ||
      column.key ||
      column.label ||
      `열${index + 1}`,
  );
}

function tableHeaders(table = {}) {
  const headers = [];
  const seen = new Set();

  const columns = Array.isArray(table.columns) ? table.columns : [];
  columns.forEach((column, index) => {
    const header = columnHeader(column, index);
    const key = normalizeKey(header);
    if (header && !seen.has(key)) {
      seen.add(key);
      headers.push(header);
    }
  });

  getRows(table)
    .slice(0, 20)
    .forEach((row) => {
      if (!row || typeof row !== "object" || Array.isArray(row)) return;
      Object.keys(row).forEach((header) => {
        const key = normalizeKey(header);
        if (header && !seen.has(key)) {
          seen.add(key);
          headers.push(header);
        }
      });
    });

  return headers;
}

function statusContextText(config = {}) {
  return [config.templateId, config.title, config.description]
    .map(normalizeText)
    .filter(Boolean)
    .join(" ");
}

function headerContextTokens(header = "") {
  return normalizeText(header)
    .replace(/상태|여부|결과|구분|분류|등급/g, " ")
    .split(/\s+/)
    .map(normalizeKey)
    .filter((token) => token.length >= 2);
}

function statusHeaderShapeScore(header = "") {
  const text = normalizeText(header);
  let score = 0;

  if (/여부$/i.test(text)) score += 180;
  if (/상태$/i.test(text)) score += 165;
  if (/처리상태|진행상태|승인상태/i.test(text)) score += 35;
  if (/결과$/i.test(text)) score += 35;
  if (/^상태$/i.test(text)) score += 20;

  if (
    /조치|계약|이수|합격|참여|배송|발주|검수|참석|출석|처리|진행|승인|정산|신청/i.test(
      text,
    )
  ) {
    score += 70;
  }

  if (
    /분류|유형|등급|점수|금액|수량|인원|시간|날짜|일자|위험|과정|단계/i.test(
      text,
    ) &&
    !/(?:상태|여부)$/i.test(text)
  ) {
    score -= 110;
  }

  return score;
}

function statusValueEvidence(table = {}, header = "") {
  const values = getRows(table)
    .map((row) => normalizeText(getRowValue(row, header)))
    .filter(Boolean)
    .slice(0, 500);

  if (!values.length) {
    return {
      nonBlankCount: 0,
      distinctCount: 0,
      recognizedCount: 0,
      recognizedRatio: 0,
      classCount: 0,
      score: -120,
      samples: [],
    };
  }

  const classes = values.map((value) => classifyStatus(value));
  const recognized = classes.filter(
    (value) => value !== "other" && value !== "unknown",
  );
  const classSet = new Set(recognized);
  const distinctSet = new Set(values.map(normalizeStatusText));

  let score = (recognized.length / values.length) * 150;
  if (distinctSet.size >= 2 && distinctSet.size <= 16) score += 35;
  if (classSet.size >= 2) score += 35;
  if (distinctSet.size > Math.max(30, values.length * 0.7)) score -= 80;

  return {
    nonBlankCount: values.length,
    distinctCount: distinctSet.size,
    recognizedCount: recognized.length,
    recognizedRatio: recognized.length / values.length,
    classCount: classSet.size,
    score,
    samples: values.slice(0, 8),
  };
}

function scoreStatusHeader({
  table = {},
  header = "",
  config = {},
  hintIndex = -1,
} = {}) {
  const contextText = normalizeKey(statusContextText(config));
  const evidence = statusValueEvidence(table, header);
  let score = statusHeaderShapeScore(header) + evidence.score;

  if (hintIndex >= 0) {
    score += Math.max(5, 40 - hintIndex);
  }

  headerContextTokens(header).forEach((token) => {
    if (contextText.includes(token)) score += 65;
  });

  return {
    header,
    score: Number(score.toFixed(6)),
    hintIndex,
    shapeScore: statusHeaderShapeScore(header),
    evidence,
  };
}

function selectStatusHeader(table = {}, config = {}) {
  const hints = [...(config.hints?.status || []), ...DEFAULT_STATUS_HINTS].map(
    normalizeText,
  );

  const hintIndexByKey = new Map();
  hints.forEach((hint, index) => {
    const key = normalizeKey(hint);
    if (key && !hintIndexByKey.has(key)) {
      hintIndexByKey.set(key, index);
    }
  });

  const candidates = tableHeaders(table)
    .map((header) =>
      scoreStatusHeader({
        table,
        header,
        config,
        hintIndex: hintIndexByKey.has(normalizeKey(header))
          ? hintIndexByKey.get(normalizeKey(header))
          : -1,
      }),
    )
    .filter((candidate) => candidate.score > 20)
    .sort(
      (left, right) =>
        right.score - left.score ||
        right.evidence.recognizedRatio - left.evidence.recognizedRatio ||
        left.header.localeCompare(right.header, "ko"),
    );

  return {
    statusHeader: candidates[0]?.header || "",
    candidates,
    version: STATUS_COLUMN_SELECTION_VERSION,
  };
}

function safeRate(numerator = 0, denominator = 0) {
  const n = Number(numerator || 0);
  const d = Number(denominator || 0);
  if (!Number.isFinite(n) || !Number.isFinite(d) || d === 0) return null;
  return n / d;
}

function makePercent(value) {
  return value == null ? null : value * 100;
}

function findStatusRateHeaders(table = {}, config = {}) {
  const statusSelection = selectStatusHeader(table, config);
  const hints = config.hints || {};

  const dateHeader = findColumnHeader(table, [
    ...(hints.date || []),
    ...DEFAULT_DATE_HINTS,
  ]);

  const metricHeader = findColumnHeader(
    table,
    [...(hints.metric || []), ...DEFAULT_METRIC_HINTS],
    { type: "number" },
  );

  const departmentHeader = findColumnHeader(table, [
    ...(hints.department || []),
    ...DEFAULT_DEPARTMENT_HINTS,
  ]);

  const ownerHeader = findColumnHeader(table, [
    ...(hints.owner || []),
    ...DEFAULT_OWNER_HINTS,
  ]);

  const categoryHeader = findColumnHeader(table, [
    ...(hints.category || []),
    ...DEFAULT_CATEGORY_HINTS,
  ]);

  return {
    statusHeader: statusSelection.statusHeader,
    statusSelection,
    dateHeader,
    metricHeader,
    departmentHeader,
    ownerHeader,
    categoryHeader,
  };
}

function commonStatusMeta(headers = {}) {
  return {
    statusRateReportVersion: STATUS_RATE_REPORT_VERSION,
    statusColumnSelectionVersion: STATUS_COLUMN_SELECTION_VERSION,
    statusClassificationVersion: STATUS_CLASSIFICATION_VERSION,
    statusRatioContractVersion: STATUS_RATIO_CONTRACT_VERSION,
    selectedStatusHeader: headers.statusHeader || "",
    statusHeaderCandidates:
      headers.statusSelection?.candidates?.slice(0, 5) || [],
  };
}

function makeCustomMetricSection({
  sectionId,
  sectionType,
  title,
  table,
  rows,
  columns = {},
  chartHint = {},
  narrativeHint = {},
  meta = {},
}) {
  return makeTemplateSection({
    sectionId,
    sectionType,
    title,
    candidate: {
      recipeType: "custom_metric",
      title,
      tableId: table.tableId,
      columns,
      meta: {
        ...meta,
        statusRateReportVersion: STATUS_RATE_REPORT_VERSION,
        statusClassificationVersion: STATUS_CLASSIFICATION_VERSION,
        statusRatioContractVersion: STATUS_RATIO_CONTRACT_VERSION,
      },
    },
    result: {
      ok: true,
      recipeType: "custom_metric",
      resultType: sectionType,
      title,
      tableId: table.tableId,
      sheetName: table.sheetName,
      columns,
      rows,
      rowCount: rows.length,
      meta: {
        ...meta,
        statusRateReportVersion: STATUS_RATE_REPORT_VERSION,
        statusClassificationVersion: STATUS_CLASSIFICATION_VERSION,
        statusRatioContractVersion: STATUS_RATIO_CONTRACT_VERSION,
      },
    },
    chartHint,
    narrativeHint,
  });
}

function emptyClassCounts() {
  return CANONICAL_STATUS_ORDER.reduce((result, key) => {
    result[key] = 0;
    return result;
  }, {});
}

function buildStatusCounts(rows = [], statusHeader = "") {
  const total = rows.length;
  const statusMap = new Map();
  const classCounts = emptyClassCounts();

  rows.forEach((row) => {
    const rawStatus = normalizeText(getRowValue(row, statusHeader)) || "미입력";
    const normalizedClass = classifyStatus(rawStatus, {
      statusHeader,
    });

    classCounts[normalizedClass] = (classCounts[normalizedClass] || 0) + 1;

    if (!statusMap.has(rawStatus)) {
      statusMap.set(rawStatus, {
        상태: rawStatus,
        건수: 0,
        상태그룹: normalizedClass,
        상태그룹명: statusClassLabel(normalizedClass),
      });
    }

    statusMap.get(rawStatus).건수 += 1;
  });

  const statusRows = Array.from(statusMap.values())
    .map((item) => {
      const ratio = safeRate(item.건수, total);
      return {
        ...item,
        전체건수: total,
        비율: ratio,
        비율Percent: makePercent(ratio),
      };
    })
    .sort(
      (left, right) =>
        Number(right.건수 || 0) - Number(left.건수 || 0) ||
        left.상태.localeCompare(right.상태, "ko"),
    );

  const classifiedTotal = Object.values(classCounts).reduce(
    (sum, value) => sum + Number(value || 0),
    0,
  );

  return {
    total,
    classifiedTotal,
    classCounts,
    statusRows,
    ratioSum: statusRows.reduce((sum, row) => sum + Number(row.비율 || 0), 0),
  };
}

function legacyPendingRollupCount(classCounts = {}) {
  return LEGACY_PENDING_ROLLUP_CLASSES.reduce(
    (sum, statusClass) => sum + Number(classCounts[statusClass] || 0),
    0,
  );
}

function makeStatusOverviewRow({
  label = "",
  count = 0,
  total = 0,
  statusGroup = "",
  statusGroupName = "",
} = {}) {
  const ratio = safeRate(count, total);

  /*
   * 객체 key 순서는 Summary Sheet의 출력 열 순서 계약이다.
   * 지표 바로 오른쪽에 숫자 값이 위치해야 scalarByLabel이
   * 이전 성공 기준선과 동일하게 읽을 수 있다.
   */
  return {
    지표: label,
    값: count,
    비율: ratio,
    비율Percent: makePercent(ratio),
    상태그룹: statusGroup,
    상태그룹명: statusGroupName,
  };
}

function buildStatusOverviewRows({
  total = 0,
  classCounts = {},
  config = {},
} = {}) {
  const labelOverrides = {
    completed: config.labels?.completed || "완료·승인 건수",
    pending: config.labels?.pending || "진행·대기 건수",
    incomplete: config.labels?.incomplete || "미완료 건수",
    cancelled: config.labels?.cancelled || "취소·반려 건수",
    delayed: config.labels?.delayed || "지연 건수",
    terminated: config.labels?.terminated || "종료·중단 건수",
    other: config.labels?.other || "기타 건수",
    unknown: config.labels?.unknown || "미입력 건수",
  };

  const rows = [
    makeStatusOverviewRow({
      label: "전체 건수",
      count: total,
      total,
      statusGroup: "total",
      statusGroupName: "전체",
    }),
  ];

  CANONICAL_STATUS_ORDER.forEach((statusClass) => {
    const canonicalCount = Number(classCounts[statusClass] || 0);

    /*
     * 성공 기준선의 '진행·대기 건수'는 과거의 넓은
     * 미완료 처리 묶음이다. 세분화된 canonical 상태를
     * 유지하면서 overview KPI에서만 다음을 합산한다.
     *
     * pending + incomplete + delayed
     */
    const count =
      statusClass === "pending"
        ? legacyPendingRollupCount(classCounts)
        : canonicalCount;

    if (count <= 0) return;

    rows.push(
      makeStatusOverviewRow({
        label: labelOverrides[statusClass],
        count,
        total,
        statusGroup: statusClass === "pending" ? "pending_rollup" : statusClass,
        statusGroupName:
          statusClass === "pending"
            ? "진행·대기(호환 집계)"
            : statusClassLabel(statusClass),
      }),
    );
  });

  return rows;
}

function buildStatusOverviewSection({ table, headers, config = {} }) {
  const { statusHeader } = headers || {};
  if (!table?.tableId || !statusHeader) return null;

  const summary = buildStatusCounts(getRows(table), statusHeader);
  if (!summary.total) return null;

  return makeCustomMetricSection({
    sectionId: config.sectionIds?.overview || "status_rate_overview",
    sectionType: "status_rate_overview",
    title: config.titles?.overview || "상태 처리율 요약",
    table,
    rows: buildStatusOverviewRows({
      total: summary.total,
      classCounts: summary.classCounts,
      config,
    }),
    columns: {
      status: statusHeader,
      total: "전체 건수",
      completed: "완료·승인 건수",
      pending: "진행·대기 건수",
      incomplete: "미완료 건수",
      cancelled: "취소·반려 건수",
      delayed: "지연 건수",
      terminated: "종료·중단 건수",
      other: "기타 건수",
      unknown: "미입력 건수",
    },
    chartHint: {
      preferredType: "metric_card",
      valueField: "값",
      ratioField: "비율Percent",
    },
    narrativeHint: {
      focus: "status_rate_overview",
      status: statusHeader,
    },
    meta: {
      ...commonStatusMeta(headers),
      statusOverviewCompatibilityVersion: STATUS_OVERVIEW_COMPATIBILITY_VERSION,
      legacyPendingRollupClasses: LEGACY_PENDING_ROLLUP_CLASSES,
      legacyPendingRollupCount: legacyPendingRollupCount(summary.classCounts),
      canonicalPendingCount: Number(summary.classCounts.pending || 0),
      canonicalIncompleteCount: Number(summary.classCounts.incomplete || 0),
      canonicalDelayedCount: Number(summary.classCounts.delayed || 0),
      sourceRowCount: summary.total,
      classifiedRowCount: summary.classifiedTotal,
      ratioSum: summary.ratioSum,
    },
  });
}

function buildStatusRatioSection({ table, headers, config = {} }) {
  const { statusHeader } = headers || {};
  if (!table?.tableId || !statusHeader) return null;

  const summary = buildStatusCounts(getRows(table), statusHeader);
  if (!summary.statusRows.length) return null;

  return makeCustomMetricSection({
    sectionId: config.sectionIds?.statusRatio || "status_ratio_breakdown",
    sectionType: "status_ratio_breakdown",
    title: config.titles?.statusRatio || `${statusHeader}별 구성비`,
    table,
    rows: summary.statusRows,
    columns: {
      status: statusHeader,
      count: "건수",
      canonicalStatus: "상태그룹",
      ratio: "비율Percent",
    },
    chartHint: {
      preferredType: "donut",
      categoryField: "상태",
      valueField: "건수",
      ratioField: "비율Percent",
    },
    narrativeHint: {
      focus: "status_ratio",
      status: statusHeader,
    },
    meta: {
      ...commonStatusMeta(headers),
      sourceRowCount: summary.total,
      classifiedRowCount: summary.classifiedTotal,
      ratioSum: summary.ratioSum,
    },
  });
}

function initialDimensionItem(dimensionHeader = "", dimension = "") {
  return {
    [dimensionHeader]: dimension,
    전체건수: 0,
    완료승인건수: 0,
    진행대기건수: 0,
    미완료건수: 0,
    취소반려건수: 0,
    지연건수: 0,
    종료중단건수: 0,
    기타미분류건수: 0,
  };
}

function incrementDimensionClass(item = {}, statusClass = "") {
  if (statusClass === "completed") {
    item.완료승인건수 += 1;
  } else if (statusClass === "pending") {
    item.진행대기건수 += 1;
  } else if (statusClass === "incomplete") {
    item.미완료건수 += 1;
  } else if (statusClass === "cancelled") {
    item.취소반려건수 += 1;
  } else if (statusClass === "delayed") {
    item.지연건수 += 1;
  } else if (statusClass === "terminated") {
    item.종료중단건수 += 1;
  } else {
    item.기타미분류건수 += 1;
  }
}

function withDimensionRates(item = {}) {
  const total = Number(item.전체건수 || 0);
  const rate = (value) => safeRate(value, total);
  const percent = (value) => makePercent(rate(value));

  return {
    ...item,
    완료율: rate(item.완료승인건수),
    완료율Percent: percent(item.완료승인건수),
    진행대기율: rate(item.진행대기건수),
    진행대기율Percent: percent(item.진행대기건수),
    미완료율: rate(item.미완료건수),
    미완료율Percent: percent(item.미완료건수),
    취소반려율: rate(item.취소반려건수),
    취소반려율Percent: percent(item.취소반려건수),
    지연율: rate(item.지연건수),
    지연율Percent: percent(item.지연건수),
    종료중단율: rate(item.종료중단건수),
    종료중단율Percent: percent(item.종료중단건수),
    기타미분류율: rate(item.기타미분류건수),
    기타미분류율Percent: percent(item.기타미분류건수),
  };
}

function buildDimensionStatusRows({
  table,
  statusHeader = "",
  dimensionHeader = "",
} = {}) {
  const map = new Map();

  getRows(table).forEach((row) => {
    const dimension =
      normalizeText(getRowValue(row, dimensionHeader)) || "미입력";
    const statusClass = classifyStatus(getRowValue(row, statusHeader), {
      statusHeader,
    });

    if (!map.has(dimension)) {
      map.set(dimension, initialDimensionItem(dimensionHeader, dimension));
    }

    const item = map.get(dimension);
    item.전체건수 += 1;
    incrementDimensionClass(item, statusClass);
  });

  return Array.from(map.values())
    .map(withDimensionRates)
    .sort(
      (left, right) =>
        Number(right.전체건수 || 0) - Number(left.전체건수 || 0) ||
        String(left[dimensionHeader]).localeCompare(
          String(right[dimensionHeader]),
          "ko",
        ),
    );
}

function buildDimensionStatusRateSection({
  table,
  headers,
  dimensionHeader = "",
  title = "",
  sectionId = "",
}) {
  const { statusHeader } = headers || {};
  if (!table?.tableId || !statusHeader || !dimensionHeader) {
    return null;
  }

  const resultRows = buildDimensionStatusRows({
    table,
    statusHeader,
    dimensionHeader,
  });
  if (!resultRows.length) return null;

  return makeCustomMetricSection({
    sectionId: sectionId || `status_rate_by_${dimensionHeader}`,
    sectionType: "status_rate_by_dimension",
    title: title || `${dimensionHeader}별 상태 처리율`,
    table,
    rows: resultRows,
    columns: {
      dimension: dimensionHeader,
      status: statusHeader,
      total: "전체건수",
      completed: "완료승인건수",
      pending: "진행대기건수",
      incomplete: "미완료건수",
      cancelled: "취소반려건수",
      delayed: "지연건수",
      terminated: "종료중단건수",
      other: "기타미분류건수",
      completionRate: "완료율Percent",
      pendingRate: "진행대기율Percent",
      incompleteRate: "미완료율Percent",
      cancelledRate: "취소반려율Percent",
    },
    chartHint: {
      preferredType: "bar",
      categoryField: dimensionHeader,
      valueField: "완료율Percent",
    },
    narrativeHint: {
      focus: "status_rate_by_dimension",
      dimension: dimensionHeader,
      status: statusHeader,
    },
    meta: {
      ...commonStatusMeta(headers),
      dimensionHeader,
      dimensionRowCount: resultRows.length,
    },
  });
}

function buildStatusRateCandidates({ table, headers, config = {} }) {
  if (!table?.tableId) return [];

  const {
    statusHeader,
    dateHeader,
    metricHeader,
    departmentHeader,
    ownerHeader,
    categoryHeader,
  } = headers || {};

  const candidates = [];
  const tableId = table.tableId;

  if (statusHeader) {
    candidates.push(
      makeTemplateCandidate({
        sectionId: config.sectionIds?.statusCount || "status_count",
        sectionType: "status_count",
        recipeType: "category_count",
        title: config.titles?.statusCount || `${statusHeader}별 건수`,
        tableId,
        columns: {
          dimension: statusHeader,
        },
        chartHint: {
          preferredType: "bar",
          categoryField: statusHeader,
          valueField: "count",
        },
        narrativeHint: {
          focus: "status_count",
          status: statusHeader,
        },
      }),
    );

    candidates.push(
      makeTemplateCandidate({
        sectionId: config.sectionIds?.statusComposition || "status_composition",
        sectionType: "status_composition",
        recipeType: "composition_ratio",
        title: config.titles?.statusComposition || `${statusHeader} 구성비`,
        tableId,
        columns: {
          dimension: statusHeader,
        },
        chartHint: {
          preferredType: "donut",
          categoryField: statusHeader,
          valueField: "value",
        },
        narrativeHint: {
          focus: "status_composition",
          status: statusHeader,
        },
      }),
    );
  }

  if (dateHeader) {
    candidates.push(
      makeTemplateCandidate({
        sectionId: config.sectionIds?.timeCount || "status_time_count",
        sectionType: "status_time_count",
        recipeType: "time_count",
        title: config.titles?.timeCount || `${dateHeader}별 건수 추이`,
        tableId,
        columns: {
          date: dateHeader,
        },
        chartHint: {
          preferredType: "line",
          categoryField: dateHeader,
          valueField: "count",
        },
        narrativeHint: {
          focus: "time_count",
          date: dateHeader,
        },
      }),
    );
  }

  const summaryDimensions = [
    departmentHeader,
    categoryHeader,
    ownerHeader,
  ].filter((value, index, values) => value && values.indexOf(value) === index);

  summaryDimensions.forEach((dimensionHeader) => {
    candidates.push(
      makeTemplateCandidate({
        sectionId: `count_by_${dimensionHeader}`,
        sectionType: "dimension_count",
        recipeType: "category_count",
        title: `${dimensionHeader}별 건수`,
        tableId,
        columns: {
          dimension: dimensionHeader,
        },
        chartHint: {
          preferredType: "bar",
          categoryField: dimensionHeader,
          valueField: "count",
        },
        narrativeHint: {
          focus: "dimension_count",
          dimension: dimensionHeader,
        },
      }),
    );

    if (metricHeader) {
      candidates.push(
        makeTemplateCandidate({
          sectionId: `metric_by_${dimensionHeader}`,
          sectionType: "dimension_metric_summary",
          recipeType: "group_summary",
          title: `${dimensionHeader}별 ${metricHeader} 요약`,
          tableId,
          columns: {
            dimension: dimensionHeader,
            metric: metricHeader,
          },
          chartHint: {
            preferredType: "bar",
            categoryField: dimensionHeader,
            valueField: metricHeader,
          },
          narrativeHint: {
            focus: "dimension_metric_summary",
            dimension: dimensionHeader,
            metric: metricHeader,
          },
        }),
      );
    }
  });

  if (metricHeader) {
    const topDimension =
      categoryHeader || ownerHeader || departmentHeader || statusHeader;

    candidates.push(
      makeTemplateCandidate({
        sectionId: "status_metric_top_bottom",
        sectionType: "status_metric_top_bottom",
        recipeType: "top_bottom",
        title: `${metricHeader} 상위·하위 항목`,
        tableId,
        columns: {
          dimension: topDimension,
          metric: metricHeader,
        },
        chartHint: {
          preferredType: "bar",
          categoryField: topDimension,
          valueField: metricHeader,
        },
        narrativeHint: {
          focus: "top_bottom",
          metric: metricHeader,
        },
      }),
    );
  }

  return candidates.filter((candidate) => {
    if (!candidate.columns) return true;
    return Object.values(candidate.columns).every(Boolean);
  });
}

function buildStatusRateReportSections({
  normalizedQueryTables = [],
  table,
  templateCandidate = {},
  config = {},
}) {
  if (!table?.tableId) return [];

  const headers = findStatusRateHeaders(table, config);

  if (!headers.statusHeader) {
    const fallbackCandidates = Array.isArray(templateCandidate.candidates)
      ? templateCandidate.candidates
      : [];

    if (!fallbackCandidates.length) return [];

    return executeTemplateSections({
      normalizedQueryTables,
      templateCandidate,
    });
  }

  const customSections = [
    buildStatusOverviewSection({
      table,
      headers,
      config,
    }),
    buildStatusRatioSection({
      table,
      headers,
      config,
    }),
    buildDimensionStatusRateSection({
      table,
      headers,
      dimensionHeader: headers.departmentHeader,
      sectionId: "status_rate_by_department",
      title: headers.departmentHeader
        ? `${headers.departmentHeader}별 상태 처리율`
        : "부서별 상태 처리율",
    }),
    buildDimensionStatusRateSection({
      table,
      headers,
      dimensionHeader: headers.categoryHeader,
      sectionId: "status_rate_by_category",
      title: headers.categoryHeader
        ? `${headers.categoryHeader}별 상태 처리율`
        : "유형별 상태 처리율",
    }),
  ].filter(Boolean);

  const candidates = buildStatusRateCandidates({
    table,
    headers,
    config,
  });
  const recipeSections = executeTemplateSections({
    normalizedQueryTables,
    templateCandidate: {
      ...templateCandidate,
      candidates,
    },
  });

  return [...customSections, ...recipeSections];
}

module.exports = {
  STATUS_RATE_REPORT_VERSION,
  STATUS_COLUMN_SELECTION_VERSION,
  STATUS_CLASSIFICATION_VERSION,
  STATUS_RATIO_CONTRACT_VERSION,
  STATUS_OVERVIEW_COMPATIBILITY_VERSION,
  LEGACY_PENDING_ROLLUP_CLASSES,
  CANONICAL_STATUS_ORDER,
  CANONICAL_STATUS_LABELS,
  normalizeStatusText,
  classifyStatus,
  statusClassLabel,
  scoreStatusHeader,
  selectStatusHeader,
  findStatusRateHeaders,
  buildStatusCounts,
  buildStatusOverviewRows,
  legacyPendingRollupCount,
  makeStatusOverviewRow,
  buildDimensionStatusRows,
  buildStatusRateCandidates,
  buildStatusRateReportSections,
};
