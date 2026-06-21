const {
  findTableForTemplate,
  findColumnHeader,
  executeTemplateSections,
  getRows,
  getRowValue,
  createVirtualTable,
  toNumber,
} = require("./commonTemplateHelpers");

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

function makeCandidate({
  sectionId,
  recipeType,
  title,
  tableId,
  columns,
  meta,
}) {
  return {
    sectionId,
    recipeType,
    title,
    tableId,
    columns,
    meta: meta || {},
  };
}

function makeCustomSection({ sectionId, title, table, rows, columns = {} }) {
  return {
    sectionId,
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
  };
}

function findHrHeaders(table = {}) {
  const nameHeader = findColumnHeader(table, [
    "성명",
    "이름",
    "직원명",
    "사원명",
    "참여자",
    "연구원명",
    "name",
  ]);

  const departmentHeader = findColumnHeader(table, [
    "부서",
    "소속",
    "조직",
    "팀",
    "부서명",
    "소속부서",
    "department",
  ]);

  const positionHeader = findColumnHeader(table, [
    "직급",
    "직위",
    "직책",
    "등급",
    "position",
    "rank",
  ]);

  const statusHeader = findColumnHeader(table, [
    "재직상태",
    "상태",
    "근무상태",
    "이수여부",
    "참여상태",
    "진행상태",
    "status",
  ]);

  const hireDateHeader = findColumnHeader(table, [
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
  const {
    nameHeader,
    departmentHeader,
    positionHeader,
    statusHeader,
    hireDateHeader,
    salaryHeader,
  } = headers;

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
      if (salary != null) item.연봉 = salary;
    }

    rows.push(item);
  });

  if (!rows.length) return null;

  const columns = [{ header: "순번", type: "number", role: "dimension" }];

  if (nameHeader)
    columns.push({ header: "성명", type: "category", role: "entity" });
  if (departmentHeader)
    columns.push({ header: "부서", type: "category", role: "dimension" });
  if (positionHeader)
    columns.push({ header: "직급", type: "category", role: "dimension" });
  if (statusHeader)
    columns.push({ header: "상태", type: "category", role: "status" });
  if (hireDateHeader)
    columns.push({ header: "입사연도", type: "category", role: "date" });
  if (hireDateHeader)
    columns.push({ header: "입사연월", type: "category", role: "date" });
  if (salaryHeader)
    columns.push({ header: "연봉", type: "number", role: "metric" });

  return createVirtualTable({
    sourceTable: table,
    tableId: `${table.tableId}_hr_roster`,
    tableName: `${table.tableName || table.sheetName || "인사"}_명단상태변환`,
    columns,
    rows,
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
  } = headers;

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

  return makeCustomSection({
    sectionId: "missing_required_fields",
    title: "확인 필요 항목",
    table,
    rows,
    columns: {
      checkedHeaders: checkHeaders,
    },
  });
}

function buildHrCandidates({ table }) {
  const candidates = [];
  const tableId = table.tableId;

  const departmentHeader = findColumnHeader(table, [
    "부서",
    "소속",
    "조직",
    "팀",
  ]);
  const positionHeader = findColumnHeader(table, [
    "직급",
    "직위",
    "직책",
    "등급",
  ]);
  const statusHeader = findColumnHeader(table, [
    "상태",
    "재직상태",
    "근무상태",
  ]);
  const hireYearHeader = findColumnHeader(table, ["입사연도"]);
  const hireYearMonthHeader = findColumnHeader(table, ["입사연월"]);
  const salaryHeader = findColumnHeader(
    table,
    ["연봉", "급여", "인건비", "금액"],
    {
      type: "number",
    },
  );
  const nameHeader = findColumnHeader(table, [
    "성명",
    "이름",
    "직원명",
    "사원명",
  ]);

  if (departmentHeader) {
    candidates.push(
      makeCandidate({
        sectionId: "department_count",
        recipeType: "category_count",
        title: `${departmentHeader}별 인원 현황`,
        tableId,
        columns: {
          dimension: departmentHeader,
        },
        meta: {
          sectionType: "department_count",
        },
      }),
    );
  }

  if (positionHeader) {
    candidates.push(
      makeCandidate({
        sectionId: "position_count",
        recipeType: "category_count",
        title: `${positionHeader}별 인원 현황`,
        tableId,
        columns: {
          dimension: positionHeader,
        },
        meta: {
          sectionType: "position_count",
        },
      }),
    );
  }

  if (statusHeader) {
    candidates.push(
      makeCandidate({
        sectionId: "status_count",
        recipeType: "category_count",
        title: `${statusHeader}별 현황`,
        tableId,
        columns: {
          dimension: statusHeader,
        },
        meta: {
          sectionType: "status_count",
        },
      }),
    );
  }

  if (hireYearMonthHeader) {
    candidates.push(
      makeCandidate({
        sectionId: "hire_year_month_count",
        recipeType: "category_count",
        title: `${hireYearMonthHeader}별 입사 현황`,
        tableId,
        columns: {
          dimension: hireYearMonthHeader,
        },
        meta: {
          sectionType: "hire_trend",
        },
      }),
    );
  } else if (hireYearHeader) {
    candidates.push(
      makeCandidate({
        sectionId: "hire_year_count",
        recipeType: "category_count",
        title: `${hireYearHeader}별 입사 현황`,
        tableId,
        columns: {
          dimension: hireYearHeader,
        },
        meta: {
          sectionType: "hire_trend",
        },
      }),
    );
  }

  if (departmentHeader && salaryHeader) {
    candidates.push(
      makeCandidate({
        sectionId: "department_salary_summary",
        recipeType: "group_summary",
        title: `${departmentHeader}별 ${salaryHeader} 요약`,
        tableId,
        columns: {
          dimension: departmentHeader,
          metric: salaryHeader,
        },
        meta: {
          sectionType: "department_salary",
        },
      }),
    );
  }

  if (salaryHeader) {
    candidates.push(
      makeCandidate({
        sectionId: "salary_top_bottom",
        recipeType: "top_bottom",
        title: `${salaryHeader} 상위/하위 항목`,
        tableId,
        columns: {
          dimension: nameHeader || departmentHeader || positionHeader || "순번",
          metric: salaryHeader,
        },
        meta: {
          sectionType: "salary_top_bottom",
        },
      }),
    );
  }

  return candidates;
}

function executeHrMonthlyReport({
  normalizedQueryTables = [],
  templateCandidate = {},
}) {
  const sourceTable = findTableForTemplate(
    normalizedQueryTables,
    templateCandidate,
  );

  if (!sourceTable?.tableId) {
    return executeTemplateSections({
      normalizedQueryTables,
      templateCandidate,
    });
  }

  const sourceRows = getRows(sourceTable);
  const headers = findHrHeaders(sourceTable);
  const virtualTable = buildRosterVirtualTable({
    table: sourceTable,
    headers,
  });

  const selectedTable = virtualTable || sourceTable;
  const executionTables = virtualTable
    ? [
        ...normalizedQueryTables.filter(
          (table) => table.tableId !== virtualTable.tableId,
        ),
        virtualTable,
      ]
    : normalizedQueryTables;

  const candidates = buildHrCandidates({
    table: selectedTable,
  });

  const customSections = [
    makeCustomSection({
      sectionId: "total_roster_count",
      title: "전체 대상 수",
      table: selectedTable,
      rows: [
        {
          지표: "전체 대상 수",
          값: sourceRows.length,
        },
      ],
      columns: {
        count: "rows",
      },
    }),
    buildMissingCheckSection({
      table: sourceTable,
      headers,
    }),
  ].filter(Boolean);

  if (!candidates.length && !customSections.length) {
    return executeTemplateSections({
      normalizedQueryTables,
      templateCandidate,
    });
  }

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
  executeHrMonthlyReport,
};
