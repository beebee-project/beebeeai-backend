const {
  findTableForTemplate,
  executeTemplateSections,
  findColumnHeader,
  getRows,
  getRowValue,
  makeTemplateSection,
  toNumber,
} = require("./commonTemplateHelpers");
const {
  buildRosterStatusReportSections,
} = require("../structuralBuilders/rosterStatusReportBuilder");
const {
  buildCategorySummaryReportSections,
} = require("../structuralBuilders/categorySummaryReportBuilder");

const HR_MONTHLY_REPORT_VERSION = "hr_monthly_report_v2";

const HR_HINTS = {
  name: [
    "성명",
    "이름",
    "직원명",
    "사원명",
    "임직원명",
    "구성원명",
    "참여자",
    "연구원명",
    "name",
    "employee name",
  ],
  employeeId: [
    "사번",
    "직번",
    "직원번호",
    "사원번호",
    "인사번호",
    "employee id",
    "emp id",
  ],
  department: [
    "부서",
    "부서명",
    "소속",
    "소속부서",
    "조직",
    "조직명",
    "본부",
    "실",
    "팀",
    "department",
    "organization",
    "team",
  ],
  position: [
    "직급",
    "직위",
    "직책",
    "직무등급",
    "직군",
    "직무",
    "등급",
    "position",
    "rank",
    "grade",
    "job",
  ],
  status: [
    "재직상태",
    "근무상태",
    "상태",
    "참여상태",
    "진행상태",
    "재직여부",
    "퇴사여부",
    "휴직여부",
    "status",
  ],
  employmentType: [
    "고용형태",
    "근무형태",
    "계약형태",
    "채용구분",
    "정규직",
    "비정규직",
    "계약직",
    "employment type",
    "contract type",
  ],
  hireDate: [
    "입사일",
    "입사일자",
    "입사년월",
    "채용일",
    "채용일자",
    "시작일",
    "참여시작일",
    "등록일",
    "hire date",
    "join date",
    "start date",
  ],
  leaveDate: [
    "퇴사일",
    "퇴사일자",
    "퇴직일",
    "종료일",
    "참여종료일",
    "이탈일",
    "leave date",
    "termination date",
    "end date",
  ],
  salary: [
    "연봉",
    "급여",
    "월급",
    "기본급",
    "보수",
    "임금",
    "인건비",
    "지급액",
    "금액",
    "salary",
    "pay",
    "wage",
    "amount",
  ],
  performance: [
    "평가등급",
    "평가",
    "성과등급",
    "성과평가",
    "고과",
    "인사평가",
    "performance",
    "rating",
    "evaluation",
  ],
  training: [
    "교육이수",
    "이수여부",
    "수료여부",
    "교육상태",
    "법정교육",
    "training",
    "education",
    "completion",
  ],
  age: ["나이", "연령", "만나이", "age"],
  tenure: ["근속", "근속년수", "재직기간", "근무기간", "tenure"],
};

function normalizeText(value = "") {
  return String(value || "")
    .toLowerCase()
    .replace(/[\s_]+/g, "")
    .replace(/[^\p{Letter}\p{Number}]/gu, "")
    .trim();
}

function findFirstColumnHeader(table = {}, hintGroups = []) {
  for (const hints of hintGroups) {
    const matched = findColumnHeader(table, hints || []);
    if (matched) return matched;
  }
  return "";
}

function resolveHrColumns(table = {}, config = {}) {
  const hints = config.hints || HR_HINTS;

  const salaryHeader = findColumnHeader(table, hints.salary || [], {
    type: "number",
  });

  return {
    nameHeader: findColumnHeader(table, hints.name || []),
    employeeIdHeader: findColumnHeader(table, hints.employeeId || []),
    departmentHeader: findColumnHeader(table, hints.department || []),
    positionHeader: findColumnHeader(table, hints.position || []),
    statusHeader: findColumnHeader(table, hints.status || []),
    employmentTypeHeader: findColumnHeader(table, hints.employmentType || []),
    hireDateHeader: findFirstColumnHeader(table, [
      hints.hireDate,
      ["입사일", "입사년월", "채용일", "시작일"],
    ]),
    leaveDateHeader: findFirstColumnHeader(table, [
      hints.leaveDate,
      ["퇴사일", "퇴직일", "종료일"],
    ]),
    salaryHeader,
    performanceHeader: findColumnHeader(table, hints.performance || []),
    trainingHeader: findColumnHeader(table, hints.training || []),
    ageHeader: findColumnHeader(table, hints.age || []),
    tenureHeader: findColumnHeader(table, hints.tenure || [], {
      type: "number",
    }),
  };
}

function normalizePeriodValue(value = "") {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const year = String(value.getFullYear());
    const month = String(value.getMonth() + 1).padStart(2, "0");
    return `${year}-${month}`;
  }

  const raw = String(value ?? "").trim();
  if (!raw) return "";

  const compact = raw.replace(/\s+/g, "");
  const ym = compact.match(/((?:19|20)\d{2})[.\-/년]?(0?[1-9]|1[0-2])?/);
  if (ym) {
    const year = ym[1];
    const month = ym[2] ? String(Number(ym[2])).padStart(2, "0") : "";
    return month ? `${year}-${month}` : year;
  }

  return raw;
}

function periodSortValue(value = "") {
  const normalized = normalizePeriodValue(value);
  const ym = normalized.match(/^((?:19|20)\d{2})(?:-(0[1-9]|1[0-2]))?$/);
  if (ym) return Number(ym[1]) * 100 + Number(ym[2] || 0);
  const n = Number(String(normalized).replace(/,/g, ""));
  return Number.isFinite(n) ? n : normalized;
}

function comparePeriodValues(a, b) {
  const av = periodSortValue(a);
  const bv = periodSortValue(b);
  if (typeof av === "number" && typeof bv === "number") return av - bv;
  return String(av).localeCompare(String(bv), "ko");
}

function getDimensionValue(row = {}, header = "") {
  const value = getRowValue(row, header);
  const normalized = String(value ?? "").trim();
  return normalized || "미분류";
}

function isActiveStatus(value = "") {
  const text = normalizeText(value);
  if (!text) return null;
  if (
    [
      "퇴사",
      "퇴직",
      "종료",
      "이탈",
      "해지",
      "inactive",
      "terminated",
      "resigned",
    ].some((token) => text.includes(normalizeText(token)))
  ) {
    return false;
  }
  if (
    ["재직", "근무", "재임", "참여", "활성", "active", "employed"].some(
      (token) => text.includes(normalizeText(token)),
    )
  ) {
    return true;
  }
  return null;
}

function countByDimension({ table, dimensionHeader }) {
  if (!dimensionHeader) return [];
  const map = new Map();
  for (const row of getRows(table)) {
    const key = getDimensionValue(row, dimensionHeader);
    const current = map.get(key) || { label: key, count: 0 };
    current.count += 1;
    map.set(key, current);
  }
  return Array.from(map.values()).sort((a, b) => b.count - a.count);
}

function sumByDimension({ table, dimensionHeader, metricHeader }) {
  if (!dimensionHeader || !metricHeader) return [];
  const map = new Map();
  for (const row of getRows(table)) {
    const key = getDimensionValue(row, dimensionHeader);
    const value = toNumber(getRowValue(row, metricHeader));
    if (value == null) continue;
    const current = map.get(key) || {
      label: key,
      count: 0,
      sum: 0,
      min: null,
      max: null,
    };
    current.count += 1;
    current.sum += value;
    current.min = current.min == null ? value : Math.min(current.min, value);
    current.max = current.max == null ? value : Math.max(current.max, value);
    map.set(key, current);
  }
  return Array.from(map.values())
    .map((row) => ({
      ...row,
      average: row.count ? row.sum / row.count : null,
    }))
    .sort((a, b) => (b.average || 0) - (a.average || 0));
}

function makeHrSection({
  sectionId,
  sectionType,
  title,
  table,
  rows,
  columns = {},
  chartHint = {},
  narrativeHint = {},
}) {
  if (!Array.isArray(rows) || !rows.length) return null;

  return makeTemplateSection({
    sectionId,
    sectionType,
    title,
    candidate: {
      recipeType: "hr_monthly_report_v2_custom",
      sectionType,
      title,
      tableId: table.tableId,
      columns,
      meta: {
        hrMonthlyReportVersion: HR_MONTHLY_REPORT_VERSION,
        sectionType,
      },
    },
    result: {
      ok: true,
      recipeType: "hr_monthly_report_v2_custom",
      title,
      tableId: table.tableId,
      sheetName: table.sheetName,
      columns,
      rows,
      rowCount: rows.length,
      meta: {
        hrMonthlyReportVersion: HR_MONTHLY_REPORT_VERSION,
      },
    },
    chartHint,
    narrativeHint: {
      ...narrativeHint,
      hrMonthlyReportVersion: HR_MONTHLY_REPORT_VERSION,
    },
  });
}

function buildHeadcountOverviewSection({ table, columns }) {
  const {
    statusHeader,
    departmentHeader,
    positionHeader,
    employmentTypeHeader,
  } = columns;
  const rows = getRows(table);
  const statusCounts = statusHeader
    ? countByDimension({ table, dimensionHeader: statusHeader })
    : [];

  let activeCount = 0;
  let inactiveCount = 0;
  let unknownStatusCount = 0;

  if (statusHeader) {
    for (const row of rows) {
      const active = isActiveStatus(getRowValue(row, statusHeader));
      if (active === true) activeCount += 1;
      else if (active === false) inactiveCount += 1;
      else unknownStatusCount += 1;
    }
  }

  const metrics = [
    { 지표: "전체 인원", 값: rows.length },
    ...(statusHeader
      ? [
          { 지표: "재직/활성 추정 인원", 값: activeCount },
          { 지표: "퇴사/종료 추정 인원", 값: inactiveCount },
          { 지표: "상태 미분류 인원", 값: unknownStatusCount },
        ]
      : []),
    ...(departmentHeader
      ? [
          {
            지표: "부서 수",
            값: countByDimension({ table, dimensionHeader: departmentHeader })
              .length,
          },
        ]
      : []),
    ...(positionHeader
      ? [
          {
            지표: "직급/직위 수",
            값: countByDimension({ table, dimensionHeader: positionHeader })
              .length,
          },
        ]
      : []),
    ...(employmentTypeHeader
      ? [
          {
            지표: "고용형태 수",
            값: countByDimension({
              table,
              dimensionHeader: employmentTypeHeader,
            }).length,
          },
        ]
      : []),
  ];

  if (!statusHeader && !rows.length) return null;

  const statusRows = statusCounts.slice(0, 10).map((row) => ({
    지표: `${statusHeader}=${row.label}`,
    값: row.count,
  }));

  return makeHrSection({
    sectionId: "hr_headcount_overview_v2",
    sectionType: "hr_headcount_overview",
    title: "월간 인원 현황 요약",
    table,
    rows: [...metrics, ...statusRows],
    columns: {
      status: statusHeader || null,
      department: departmentHeader || null,
      position: positionHeader || null,
      employmentType: employmentTypeHeader || null,
    },
    chartHint: {
      preferredType: "metric_card",
      categoryField: "지표",
      valueField: "값",
    },
    narrativeHint: {
      focus: "hr_headcount_overview",
    },
  });
}

function buildCountCompositionSection({
  table,
  dimensionHeader,
  sectionId,
  sectionType,
  title,
}) {
  if (!dimensionHeader) return null;

  const counted = countByDimension({ table, dimensionHeader }).slice(0, 30);
  const total = counted.reduce((sum, row) => sum + row.count, 0);
  if (!total) return null;

  const rows = counted.map((row, index) => ({
    순위: index + 1,
    [dimensionHeader]: row.label,
    인원수: row.count,
    구성비: row.count / total,
  }));

  return makeHrSection({
    sectionId,
    sectionType,
    title,
    table,
    rows,
    columns: {
      dimension: dimensionHeader,
      count: "인원수",
    },
    chartHint: {
      preferredType: "bar",
      categoryField: dimensionHeader,
      valueField: "인원수",
      secondaryValueFields: ["구성비"],
    },
    narrativeHint: {
      focus: sectionType,
      dimension: dimensionHeader,
    },
  });
}

function buildHireLeaveTrendSection({ table, columns, mode = "hire" }) {
  const dateHeader =
    mode === "leave" ? columns.leaveDateHeader : columns.hireDateHeader;
  if (!dateHeader) return null;

  const map = new Map();
  for (const row of getRows(table)) {
    const period = normalizePeriodValue(getRowValue(row, dateHeader));
    if (!period) continue;
    const current = map.get(period) || { 기간: period, 인원수: 0 };
    current.인원수 += 1;
    map.set(period, current);
  }

  const rows = Array.from(map.values()).sort((a, b) =>
    comparePeriodValues(a.기간, b.기간),
  );

  return makeHrSection({
    sectionId: mode === "leave" ? "leave_trend_v2" : "hire_trend_v2",
    sectionType: mode === "leave" ? "leave_trend" : "hire_trend",
    title: `${dateHeader} 기준 ${mode === "leave" ? "퇴사/종료" : "입사/채용"} 추이`,
    table,
    rows,
    columns: {
      date: dateHeader,
      count: "인원수",
    },
    chartHint: {
      preferredType: "line",
      categoryField: "기간",
      valueField: "인원수",
    },
    narrativeHint: {
      focus: mode === "leave" ? "leave_trend" : "hire_trend",
      date: dateHeader,
    },
  });
}

function buildMovementSection({ table, columns }) {
  const { hireDateHeader, leaveDateHeader } = columns;
  if (!hireDateHeader && !leaveDateHeader) return null;

  const map = new Map();
  function ensure(period) {
    const current = map.get(period) || { 기간: period, 입사수: 0, 퇴사수: 0 };
    map.set(period, current);
    return current;
  }

  for (const row of getRows(table)) {
    if (hireDateHeader) {
      const period = normalizePeriodValue(getRowValue(row, hireDateHeader));
      if (period) ensure(period).입사수 += 1;
    }
    if (leaveDateHeader) {
      const period = normalizePeriodValue(getRowValue(row, leaveDateHeader));
      if (period) ensure(period).퇴사수 += 1;
    }
  }

  const rows = Array.from(map.values())
    .sort((a, b) => comparePeriodValues(a.기간, b.기간))
    .map((row) => ({
      ...row,
      순증감: row.입사수 - row.퇴사수,
    }));

  return makeHrSection({
    sectionId: "hr_movement_v2",
    sectionType: "hr_movement",
    title: "입사·퇴사 인원 증감 추이",
    table,
    rows,
    columns: {
      hireDate: hireDateHeader || null,
      leaveDate: leaveDateHeader || null,
    },
    chartHint: {
      preferredType: "line",
      categoryField: "기간",
      valueField: "순증감",
      secondaryValueFields: ["입사수", "퇴사수"],
    },
    narrativeHint: {
      focus: "hr_movement",
    },
  });
}

function buildSalarySummarySection({
  table,
  dimensionHeader,
  salaryHeader,
  sectionId,
  sectionType,
  title,
}) {
  if (!dimensionHeader || !salaryHeader) return null;

  const rows = sumByDimension({
    table,
    dimensionHeader,
    metricHeader: salaryHeader,
  })
    .slice(0, 30)
    .map((row, index) => ({
      순위: index + 1,
      [dimensionHeader]: row.label,
      인원수: row.count,
      [`${salaryHeader} 합계`]: row.sum,
      [`${salaryHeader} 평균`]: row.average,
      [`${salaryHeader} 최소`]: row.min,
      [`${salaryHeader} 최대`]: row.max,
    }));

  return makeHrSection({
    sectionId,
    sectionType,
    title,
    table,
    rows,
    columns: {
      dimension: dimensionHeader,
      metric: salaryHeader,
    },
    chartHint: {
      preferredType: "bar",
      categoryField: dimensionHeader,
      valueField: `${salaryHeader} 평균`,
      secondaryValueFields: [`${salaryHeader} 합계`, "인원수"],
    },
    narrativeHint: {
      focus: sectionType,
      dimension: dimensionHeader,
      metric: salaryHeader,
    },
  });
}

function buildSalaryTopBottomSection({ table, columns }) {
  const {
    salaryHeader,
    nameHeader,
    employeeIdHeader,
    departmentHeader,
    positionHeader,
  } = columns;
  if (!salaryHeader) return null;

  const labelHeader =
    nameHeader ||
    employeeIdHeader ||
    departmentHeader ||
    positionHeader ||
    "순번";
  const salaryRows = getRows(table)
    .map((row, index) => {
      const salary = toNumber(getRowValue(row, salaryHeader));
      if (salary == null) return null;
      return {
        순번: index + 1,
        구분: getRowValue(row, labelHeader) || `${index + 1}`,
        부서: departmentHeader ? getRowValue(row, departmentHeader) : "",
        직급: positionHeader ? getRowValue(row, positionHeader) : "",
        [salaryHeader]: salary,
      };
    })
    .filter(Boolean)
    .sort((a, b) => b[salaryHeader] - a[salaryHeader]);

  if (!salaryRows.length) return null;

  const top = salaryRows.slice(0, 10).map((row, index) => ({
    순위구분: "상위",
    순위: index + 1,
    ...row,
  }));
  const bottom = salaryRows
    .slice(-10)
    .reverse()
    .map((row, index) => ({
      순위구분: "하위",
      순위: index + 1,
      ...row,
    }));

  return makeHrSection({
    sectionId: "salary_top_bottom_v2",
    sectionType: "salary_top_bottom",
    title: `${salaryHeader} 상위·하위 항목`,
    table,
    rows: [...top, ...bottom],
    columns: {
      dimension: labelHeader,
      metric: salaryHeader,
    },
    chartHint: {
      preferredType: "bar",
      categoryField: "구분",
      valueField: salaryHeader,
    },
    narrativeHint: {
      focus: "salary_top_bottom",
      metric: salaryHeader,
    },
  });
}

function buildSalaryBandSection({ table, columns }) {
  const { salaryHeader } = columns;
  if (!salaryHeader) return null;

  const values = getRows(table)
    .map((row) => toNumber(getRowValue(row, salaryHeader)))
    .filter((value) => value != null)
    .sort((a, b) => a - b);
  if (!values.length) return null;

  const max = values[values.length - 1];
  const unit = max > 100000 ? 10000000 : 1000;
  const unitLabel = max > 100000 ? "천만원" : "천";
  const map = new Map();

  values.forEach((value) => {
    const bucketStart = Math.floor(value / unit) * unit;
    const bucketEnd = bucketStart + unit;
    const label = `${bucketStart / unit}${unitLabel}~${bucketEnd / unit}${unitLabel}`;
    const current = map.get(label) || { 급여구간: label, 인원수: 0 };
    current.인원수 += 1;
    map.set(label, current);
  });

  const rows = Array.from(map.values());
  return makeHrSection({
    sectionId: "salary_band_distribution_v2",
    sectionType: "salary_band_distribution",
    title: `${salaryHeader} 구간별 인원 분포`,
    table,
    rows,
    columns: {
      metric: salaryHeader,
      dimension: "급여구간",
    },
    chartHint: {
      preferredType: "bar",
      categoryField: "급여구간",
      valueField: "인원수",
    },
    narrativeHint: {
      focus: "salary_band_distribution",
      metric: salaryHeader,
    },
  });
}

function buildTrainingStatusSection({ table, columns }) {
  const { trainingHeader } = columns;
  if (!trainingHeader) return null;
  return buildCountCompositionSection({
    table,
    dimensionHeader: trainingHeader,
    sectionId: "training_status_distribution_v2",
    sectionType: "training_status_distribution",
    title: `${trainingHeader}별 인원 현황`,
  });
}

function hasSameSection(sections = [], sectionId = "") {
  if (!sectionId) return false;
  return sections.some((section) => section.sectionId === sectionId);
}

function uniqueSections(sections = []) {
  const seen = new Set();
  return sections.filter((section) => {
    if (!section) return false;
    const key = [
      section.sectionId,
      section.sectionType,
      section.title,
      section.candidate?.columns?.dimension,
      section.candidate?.columns?.metric,
      section.candidate?.columns?.date,
    ].join("|");
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function buildHrMonthlyV2Sections({ table, templateCandidate, config = {} }) {
  const columns = resolveHrColumns(table, config);
  const sections = [
    buildHeadcountOverviewSection({ table, columns }),
    buildCountCompositionSection({
      table,
      dimensionHeader: columns.departmentHeader,
      sectionId: "department_headcount_composition_v2",
      sectionType: "department_headcount_composition",
      title: columns.departmentHeader
        ? `${columns.departmentHeader}별 인원 구성`
        : "부서별 인원 구성",
    }),
    buildCountCompositionSection({
      table,
      dimensionHeader: columns.positionHeader,
      sectionId: "position_headcount_composition_v2",
      sectionType: "position_headcount_composition",
      title: columns.positionHeader
        ? `${columns.positionHeader}별 인원 구성`
        : "직급별 인원 구성",
    }),
    buildCountCompositionSection({
      table,
      dimensionHeader: columns.statusHeader,
      sectionId: "status_headcount_composition_v2",
      sectionType: "status_headcount_composition",
      title: columns.statusHeader
        ? `${columns.statusHeader}별 인원 구성`
        : "재직상태별 인원 구성",
    }),
    buildCountCompositionSection({
      table,
      dimensionHeader: columns.employmentTypeHeader,
      sectionId: "employment_type_composition_v2",
      sectionType: "employment_type_composition",
      title: columns.employmentTypeHeader
        ? `${columns.employmentTypeHeader}별 인원 구성`
        : "고용형태별 인원 구성",
    }),
    buildHireLeaveTrendSection({ table, columns, mode: "hire" }),
    buildHireLeaveTrendSection({ table, columns, mode: "leave" }),
    buildMovementSection({ table, columns }),
    buildSalarySummarySection({
      table,
      dimensionHeader: columns.departmentHeader,
      salaryHeader: columns.salaryHeader,
      sectionId: "department_salary_summary_v2",
      sectionType: "department_salary_summary",
      title:
        columns.departmentHeader && columns.salaryHeader
          ? `${columns.departmentHeader}별 ${columns.salaryHeader} 평균·합계`
          : "부서별 급여 요약",
    }),
    buildSalarySummarySection({
      table,
      dimensionHeader: columns.positionHeader,
      salaryHeader: columns.salaryHeader,
      sectionId: "position_salary_summary_v2",
      sectionType: "position_salary_summary",
      title:
        columns.positionHeader && columns.salaryHeader
          ? `${columns.positionHeader}별 ${columns.salaryHeader} 평균·합계`
          : "직급별 급여 요약",
    }),
    buildCountCompositionSection({
      table,
      dimensionHeader: columns.performanceHeader,
      sectionId: "performance_rating_distribution_v2",
      sectionType: "performance_rating_distribution",
      title: columns.performanceHeader
        ? `${columns.performanceHeader}별 인원 현황`
        : "평가등급별 인원 현황",
    }),
    buildSalarySummarySection({
      table,
      dimensionHeader: columns.performanceHeader,
      salaryHeader: columns.salaryHeader,
      sectionId: "performance_salary_summary_v2",
      sectionType: "performance_salary_summary",
      title:
        columns.performanceHeader && columns.salaryHeader
          ? `${columns.performanceHeader}별 ${columns.salaryHeader} 평균`
          : "평가등급별 급여 요약",
    }),
    buildSalaryBandSection({ table, columns }),
    buildSalaryTopBottomSection({ table, columns }),
    buildTrainingStatusSection({ table, columns }),
  ].filter(Boolean);

  return uniqueSections(sections).map((section) => ({
    ...section,
    candidate: {
      ...section.candidate,
      templateId: templateCandidate.templateId || "hr_monthly_report",
    },
  }));
}

function buildHrMonthlyReportConfig() {
  return {
    hints: HR_HINTS,
    sectionIds: {
      totalCount: "total_roster_count",
      department: "department_count",
      position: "position_count",
      status: "status_count",
      hireTrend: "hire_year_month_count",
      departmentSalary: "department_salary_summary",
      salaryTopBottom: "salary_top_bottom",
    },
    sectionTypes: {
      totalCount: "total_roster_count",
      department: "department_count",
      position: "position_count",
      status: "status_count",
      hireTrend: "hire_trend",
      departmentSalary: "department_salary",
      salaryTopBottom: "salary_top_bottom",
    },
    titles: {
      totalCount: "전체 대상 수",
      department: "부서별 인원 현황",
      position: "직급별 인원 현황",
      status: "재직상태별 인원 현황",
      hireTrend: "입사연월별 입사 현황",
      departmentSalary: "부서별 연봉 요약",
      salaryTopBottom: "연봉 상위 하위 항목",
    },
  };
}

function executeHrMonthlyReport({
  normalizedQueryTables = [],
  templateCandidate = {},
}) {
  const table = findTableForTemplate(normalizedQueryTables, templateCandidate);

  if (!table?.tableId) {
    return executeTemplateSections({
      normalizedQueryTables,
      templateCandidate,
    });
  }

  const config = buildHrMonthlyReportConfig();

  const v2Sections = buildHrMonthlyV2Sections({
    table,
    templateCandidate,
    config,
  });

  const rosterSections = buildRosterStatusReportSections({
    normalizedQueryTables,
    table,
    templateCandidate,
    config,
  });

  const categorySections = buildCategorySummaryReportSections({
    normalizedQueryTables,
    table,
    templateCandidate,
    config: {
      metricHints: HR_HINTS.salary,
      dimensions: [
        {
          sectionId: "department_salary_summary",
          sectionType: "department_salary",
          hints: HR_HINTS.department,
        },
        {
          sectionId: "position_salary_summary",
          sectionType: "position_salary",
          hints: HR_HINTS.position,
        },
        {
          sectionId: "performance_salary_summary",
          sectionType: "performance_salary",
          hints: HR_HINTS.performance,
        },
      ],
      topBottom: {
        sectionId: "salary_top_bottom",
        sectionType: "salary_top_bottom",
        dimensionHints: [
          ...HR_HINTS.name,
          ...HR_HINTS.employeeId,
          ...HR_HINTS.department,
        ],
      },
    },
  });

  const sections = uniqueSections([
    ...v2Sections,
    ...rosterSections.filter(
      (section) => !hasSameSection(v2Sections, section.sectionId),
    ),
    ...categorySections.filter(
      (section) => !hasSameSection(v2Sections, section.sectionId),
    ),
  ]).map((section) => ({
    ...section,
    meta: {
      ...(section.meta || {}),
      hrMonthlyReportVersion: HR_MONTHLY_REPORT_VERSION,
    },
  }));

  if (!sections.length) {
    return executeTemplateSections({
      normalizedQueryTables,
      templateCandidate,
    });
  }

  return sections;
}

module.exports = {
  HR_MONTHLY_REPORT_VERSION,
  HR_HINTS,
  executeHrMonthlyReport,
  buildHrMonthlyReportConfig,
  buildHrMonthlyV2Sections,
};
