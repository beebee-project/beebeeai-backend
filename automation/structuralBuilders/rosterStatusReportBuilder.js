const {
  findColumnHeader,
  makeTemplateCandidate,
  makeTemplateSection,
  executeTemplateSections,
  getRows,
  getRowValue,
  createVirtualTable,
  toNumber,
} = require("../businessTemplates/commonTemplateHelpers");

function normalizeDateParts(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const year = String(value.getFullYear());
    const month = String(value.getMonth() + 1).padStart(2, "0");

    return {
      year,
      month,
      yearMonth: `${year}-${month}`,
    };
  }

  const raw = String(value ?? "").trim();

  if (!raw) {
    return {
      year: "",
      month: "",
      yearMonth: "",
    };
  }

  const matched = raw.match(/((?:19|20)\d{2})[.\-/년\s]*(0?[1-9]|1[0-2])?/);

  if (!matched) {
    return {
      year: "",
      month: "",
      yearMonth: "",
    };
  }

  const year = matched[1];
  const month = matched[2] ? String(Number(matched[2])).padStart(2, "0") : "";

  return {
    year,
    month,
    yearMonth: month ? `${year}-${month}` : year,
  };
}

function findRosterHeaders(table = {}, config = {}) {
  const hints = config.hints || {};

  const nameHeader = findColumnHeader(table, [
    ...(hints.name || []),
    "성명",
    "이름",
    "직원명",
    "사원명",
    "참여자",
    "연구원명",
    "name",
  ]);

  const departmentHeader = findColumnHeader(table, [
    ...(hints.department || []),
    "부서",
    "소속",
    "조직",
    "팀",
    "부서명",
    "소속부서",
    "department",
  ]);

  const positionHeader = findColumnHeader(table, [
    ...(hints.position || []),
    "직급",
    "직위",
    "직책",
    "등급",
    "position",
    "rank",
  ]);

  const statusHeader = findColumnHeader(table, [
    ...(hints.status || []),
    "재직상태",
    "상태",
    "근무상태",
    "이수여부",
    "참여상태",
    "진행상태",
    "status",
  ]);

  const hireDateHeader = findColumnHeader(table, [
    ...(hints.hireDate || []),
    "입사일",
    "입사일자",
    "채용일",
    "시작일",
    "참여시작일",
    "등록일",
    "date",
  ]);

  const salaryHeader = findColumnHeader(
    table,
    [
      ...(hints.salary || []),
      "연봉",
      "급여",
      "월급",
      "인건비",
      "지급액",
      "금액",
      "salary",
      "pay",
      "amount",
    ],
    { type: "number" },
  );

  return {
    nameHeader,
    departmentHeader,
    positionHeader,
    statusHeader,
    hireDateHeader,
    salaryHeader,
  };
}

function buildRosterVirtualTable({ table, headers }) {
  if (!table?.tableId) return null;

  const {
    nameHeader,
    departmentHeader,
    positionHeader,
    statusHeader,
    hireDateHeader,
    salaryHeader,
  } = headers || {};

  const rows = [];

  getRows(table).forEach((row, index) => {
    const dateParts = hireDateHeader
      ? normalizeDateParts(getRowValue(row, hireDateHeader))
      : { year: "", month: "", yearMonth: "" };

    const item = {
      순번: index + 1,
    };

    if (nameHeader) item.성명 = getRowValue(row, nameHeader);
    if (departmentHeader) item.부서 = getRowValue(row, departmentHeader);
    if (positionHeader) item.직급 = getRowValue(row, positionHeader);
    if (statusHeader) item.상태 = getRowValue(row, statusHeader);
    if (hireDateHeader) item.입사연도 = dateParts.year;
    if (hireDateHeader) item.입사연월 = dateParts.yearMonth;

    if (salaryHeader) {
      const salary = toNumber(getRowValue(row, salaryHeader));
      item.연봉 = salary;
    }

    rows.push(item);
  });

  if (!rows.length) return null;

  const columns = [
    { header: "순번", type: "number", role: "id" },
    ...(nameHeader
      ? [{ header: "성명", type: "string", role: "dimension" }]
      : []),
    ...(departmentHeader
      ? [{ header: "부서", type: "string", role: "dimension" }]
      : []),
    ...(positionHeader
      ? [{ header: "직급", type: "string", role: "dimension" }]
      : []),
    ...(statusHeader
      ? [{ header: "상태", type: "string", role: "status" }]
      : []),
    ...(hireDateHeader
      ? [
          { header: "입사연도", type: "string", role: "date" },
          { header: "입사연월", type: "string", role: "date" },
        ]
      : []),
    ...(salaryHeader
      ? [{ header: "연봉", type: "number", role: "metric" }]
      : []),
  ];

  return createVirtualTable({
    sourceTable: table,
    tableId: `${table.tableId}_roster_status`,
    tableName: `${table.tableName || table.sheetName || "명단"}_상태관리`,
    columns,
    rows,
  });
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
    },
    result: {
      ok: true,
      recipeType: "custom_metric",
      title,
      tableId: table.tableId,
      sheetName: table.sheetName,
      columns,
      rows,
      rowCount: rows.length,
    },
    chartHint,
    narrativeHint,
  });
}

function buildMissingCheckSection({ table, headers }) {
  const {
    nameHeader,
    departmentHeader,
    positionHeader,
    statusHeader,
    hireDateHeader,
    salaryHeader,
  } = headers || {};

  const checkHeaders = [
    nameHeader,
    departmentHeader,
    positionHeader,
    statusHeader,
    hireDateHeader,
    salaryHeader,
  ].filter(Boolean);

  if (!checkHeaders.length) return null;

  const rows = [];

  getRows(table).forEach((row, index) => {
    const missing = checkHeaders.filter((header) => {
      const value = getRowValue(row, header);
      return value == null || String(value).trim() === "";
    });

    if (!missing.length) return;

    rows.push({
      순번: index + 1,
      확인필요항목: missing.join(", "),
      성명: nameHeader ? getRowValue(row, nameHeader) : "",
      부서: departmentHeader ? getRowValue(row, departmentHeader) : "",
      상태: statusHeader ? getRowValue(row, statusHeader) : "",
    });
  });

  if (!rows.length) return null;

  return makeCustomMetricSection({
    sectionId: "missing_required_fields",
    sectionType: "missing_required_fields",
    title: "확인 필요 항목",
    table,
    rows,
    columns: {
      checkedHeaders: checkHeaders,
    },
    chartHint: {
      preferredType: "table",
    },
    narrativeHint: {
      focus: "data_quality",
      checkedHeaders: checkHeaders,
    },
  });
}

function buildRosterStatusCandidates({ table, config = {} }) {
  if (!table?.tableId) return [];

  const candidates = [];
  const tableId = table.tableId;

  const departmentHeader = findColumnHeader(table, [
    ...(config.hints?.department || []),
    "부서",
    "소속",
    "조직",
    "팀",
  ]);

  const positionHeader = findColumnHeader(table, [
    ...(config.hints?.position || []),
    "직급",
    "직위",
    "직책",
    "등급",
  ]);

  const statusHeader = findColumnHeader(table, [
    ...(config.hints?.status || []),
    "상태",
    "재직상태",
    "근무상태",
  ]);

  const hireYearHeader = findColumnHeader(table, ["입사연도"]);
  const hireYearMonthHeader = findColumnHeader(table, ["입사연월"]);

  const salaryHeader = findColumnHeader(
    table,
    [...(config.hints?.salary || []), "연봉", "급여", "인건비", "금액"],
    { type: "number" },
  );

  const nameHeader = findColumnHeader(table, [
    ...(config.hints?.name || []),
    "성명",
    "이름",
    "직원명",
    "사원명",
  ]);

  if (departmentHeader) {
    candidates.push(
      makeTemplateCandidate({
        sectionId: config.sectionIds?.department || "department_count",
        sectionType: config.sectionTypes?.department || "department_count",
        recipeType: "category_count",
        title: config.titles?.department || `${departmentHeader}별 인원 현황`,
        tableId,
        columns: {
          dimension: departmentHeader,
        },
        chartHint: {
          preferredType: "bar",
          categoryField: departmentHeader,
          valueField: "count",
        },
        narrativeHint: {
          focus: "department_count",
          dimension: departmentHeader,
        },
      }),
    );
  }

  if (positionHeader) {
    candidates.push(
      makeTemplateCandidate({
        sectionId: config.sectionIds?.position || "position_count",
        sectionType: config.sectionTypes?.position || "position_count",
        recipeType: "category_count",
        title: config.titles?.position || `${positionHeader}별 인원 현황`,
        tableId,
        columns: {
          dimension: positionHeader,
        },
        chartHint: {
          preferredType: "bar",
          categoryField: positionHeader,
          valueField: "count",
        },
        narrativeHint: {
          focus: "position_count",
          dimension: positionHeader,
        },
      }),
    );
  }

  if (statusHeader) {
    candidates.push(
      makeTemplateCandidate({
        sectionId: config.sectionIds?.status || "status_count",
        sectionType: config.sectionTypes?.status || "status_count",
        recipeType: "category_count",
        title: config.titles?.status || `${statusHeader}별 현황`,
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
          dimension: statusHeader,
        },
      }),
    );
  }

  if (hireYearMonthHeader) {
    candidates.push(
      makeTemplateCandidate({
        sectionId: config.sectionIds?.hireTrend || "hire_year_month_count",
        sectionType: config.sectionTypes?.hireTrend || "hire_trend",
        recipeType: "category_count",
        title: config.titles?.hireTrend || `${hireYearMonthHeader}별 입사 현황`,
        tableId,
        columns: {
          dimension: hireYearMonthHeader,
        },
        chartHint: {
          preferredType: "line",
          categoryField: hireYearMonthHeader,
          valueField: "count",
        },
        narrativeHint: {
          focus: "hire_trend",
          dimension: hireYearMonthHeader,
        },
      }),
    );
  } else if (hireYearHeader) {
    candidates.push(
      makeTemplateCandidate({
        sectionId: config.sectionIds?.hireTrend || "hire_year_count",
        sectionType: config.sectionTypes?.hireTrend || "hire_trend",
        recipeType: "category_count",
        title: config.titles?.hireTrend || `${hireYearHeader}별 입사 현황`,
        tableId,
        columns: {
          dimension: hireYearHeader,
        },
        chartHint: {
          preferredType: "line",
          categoryField: hireYearHeader,
          valueField: "count",
        },
        narrativeHint: {
          focus: "hire_trend",
          dimension: hireYearHeader,
        },
      }),
    );
  }

  if (departmentHeader && salaryHeader) {
    candidates.push(
      makeTemplateCandidate({
        sectionId:
          config.sectionIds?.departmentSalary || "department_salary_summary",
        sectionType:
          config.sectionTypes?.departmentSalary || "department_salary",
        recipeType: "group_summary",
        title:
          config.titles?.departmentSalary ||
          `${departmentHeader}별 ${salaryHeader} 요약`,
        tableId,
        columns: {
          dimension: departmentHeader,
          metric: salaryHeader,
        },
        chartHint: {
          preferredType: "bar",
          categoryField: departmentHeader,
          valueField: salaryHeader,
        },
        narrativeHint: {
          focus: "department_salary",
          dimension: departmentHeader,
          metric: salaryHeader,
        },
      }),
    );
  }

  if (salaryHeader) {
    candidates.push(
      makeTemplateCandidate({
        sectionId: config.sectionIds?.salaryTopBottom || "salary_top_bottom",
        sectionType:
          config.sectionTypes?.salaryTopBottom || "salary_top_bottom",
        recipeType: "top_bottom",
        title:
          config.titles?.salaryTopBottom || `${salaryHeader} 상위/하위 항목`,
        tableId,
        columns: {
          dimension: nameHeader || departmentHeader || positionHeader || "순번",
          metric: salaryHeader,
        },
        chartHint: {
          preferredType: "bar",
          categoryField:
            nameHeader || departmentHeader || positionHeader || "순번",
          valueField: salaryHeader,
        },
        narrativeHint: {
          focus: "salary_top_bottom",
          metric: salaryHeader,
        },
      }),
    );
  }

  return candidates;
}

function buildRosterStatusReportSections({
  normalizedQueryTables = [],
  table,
  templateCandidate = {},
  config = {},
}) {
  if (!table?.tableId) return [];

  const headers = findRosterHeaders(table, config);
  const virtualTable = buildRosterVirtualTable({ table, headers });

  const selectedTable = virtualTable || table;
  const executionTables = virtualTable
    ? [
        ...normalizedQueryTables.filter(
          (item) => item.tableId !== virtualTable.tableId,
        ),
        virtualTable,
      ]
    : normalizedQueryTables;

  const candidates = buildRosterStatusCandidates({
    table: selectedTable,
    config,
  });

  const customSections = [
    makeCustomMetricSection({
      sectionId: config.sectionIds?.totalCount || "total_roster_count",
      sectionType: config.sectionTypes?.totalCount || "total_roster_count",
      title: config.titles?.totalCount || "전체 대상 수",
      table: selectedTable,
      rows: [
        {
          지표: "전체 대상 수",
          값: getRows(table).length,
        },
      ],
      columns: {
        count: "rows",
      },
      chartHint: {
        preferredType: "metric_card",
      },
      narrativeHint: {
        focus: "total_count",
      },
    }),
    buildMissingCheckSection({
      table,
      headers,
    }),
  ].filter(Boolean);

  if (!candidates.length && !customSections.length) return [];

  const recipeSections = executeTemplateSections({
    normalizedQueryTables: executionTables,
    templateCandidate: {
      ...templateCandidate,
      candidates,
    },
  });

  return [...customSections, ...recipeSections];
}

module.exports = {
  buildRosterStatusReportSections,
  buildRosterStatusCandidates,
  buildRosterVirtualTable,
  findRosterHeaders,
};
