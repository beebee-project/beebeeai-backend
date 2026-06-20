const { executeAnalysisRecipeCandidate } = require("./analysisRecipeExecutor");
const {
  executeResearchBudgetReport,
} = require("./businessTemplates/researchBudgetReportBuilder");
const {
  executeSalesReport,
} = require("./businessTemplates/salesReportBuilder");

function executeTemplateSections({
  normalizedQueryTables = [],
  templateCandidate = {},
}) {
  const candidates = Array.isArray(templateCandidate.candidates)
    ? templateCandidate.candidates
    : [];

  return candidates
    .map((candidate, index) => {
      const result = executeAnalysisRecipeCandidate({
        normalizedQueryTables,
        candidate,
      });

      if (!result?.ok) return null;

      return {
        sectionId:
          candidate.recipeType ||
          candidate.type ||
          candidate.recipeId ||
          `section_${index + 1}`,
        title:
          candidate.title ||
          candidate.name ||
          candidate.label ||
          `섹션 ${index + 1}`,
        candidate,
        result,
      };
    })
    .filter(Boolean);
}

function normalizeHeader(value = "") {
  return String(value || "")
    .toLowerCase()
    .replace(/[\s_]+/g, "")
    .replace(/[^\p{Letter}\p{Number}]/gu, "")
    .trim();
}

function headerMatches(header = "", hints = []) {
  const h = normalizeHeader(header);

  return hints.some((hint) => {
    const normalizedHint = normalizeHeader(hint);
    return h.includes(normalizedHint) || normalizedHint.includes(h);
  });
}

function findTableForTemplate(
  normalizedQueryTables = [],
  templateCandidate = {},
) {
  const firstCandidate = Array.isArray(templateCandidate.candidates)
    ? templateCandidate.candidates[0]
    : null;

  if (firstCandidate?.tableId) {
    const matched = normalizedQueryTables.find(
      (table) => table.tableId === firstCandidate.tableId,
    );

    if (matched) return matched;
  }

  return (
    normalizedQueryTables.find((table) => table.isPrimary) ||
    normalizedQueryTables[0] ||
    null
  );
}

function findColumnHeader(table = {}, hints = [], options = {}) {
  const columns = Array.isArray(table.columns) ? table.columns : [];

  const matchedByHint = columns.find((col) =>
    headerMatches(col.header || col.key || "", hints),
  );

  if (matchedByHint?.header) return matchedByHint.header;

  if (options.type) {
    const matchedByType = columns.find(
      (col) => col.type === options.type || col.dominantType === options.type,
    );

    if (matchedByType?.header) return matchedByType.header;
  }

  if (options.role) {
    const matchedByRole = columns.find(
      (col) => col.role === options.role || col.inferredRole === options.role,
    );

    if (matchedByRole?.header) return matchedByRole.header;
  }

  return "";
}

function makeHrCandidate({ recipeType, title, tableId, columns }) {
  return {
    recipeType,
    title,
    tableId,
    columns,
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

  const tableId = table.tableId;

  const departmentHeader = findColumnHeader(table, [
    "부서",
    "소속",
    "소속팀",
    "팀",
    "조직",
    "부문",
  ]);

  const positionHeader = findColumnHeader(table, [
    "직급",
    "직위",
    "직책",
    "레벨",
    "등급",
  ]);

  const ratingHeader = findColumnHeader(table, [
    "평가등급",
    "평가 등급",
    "평가",
    "성과등급",
    "등급",
  ]);

  const salaryHeader = findColumnHeader(
    table,
    ["연봉", "급여", "보수", "salary", "pay"],
    { type: "number" },
  );

  const hireDateHeader = findColumnHeader(
    table,
    ["입사일", "입직일", "채용일", "입사", "hiredate", "hire"],
    { type: "date" },
  );

  const personHeader =
    findColumnHeader(table, [
      "이름",
      "성명",
      "직원명",
      "사원명",
      "성명",
      "name",
    ]) ||
    findColumnHeader(table, [
      "직원 ID",
      "사번",
      "직원번호",
      "employeeid",
      "id",
    ]);

  const candidates = [];

  if (departmentHeader) {
    candidates.push(
      makeHrCandidate({
        recipeType: "category_count",
        title: `${departmentHeader}별 인원`,
        tableId,
        columns: {
          dimension: departmentHeader,
        },
      }),
    );
  }

  if (positionHeader) {
    candidates.push(
      makeHrCandidate({
        recipeType: "category_count",
        title: `${positionHeader}별 인원`,
        tableId,
        columns: {
          dimension: positionHeader,
        },
      }),
    );
  }

  if (ratingHeader && ratingHeader !== positionHeader) {
    candidates.push(
      makeHrCandidate({
        recipeType: "category_count",
        title: `${ratingHeader}별 인원`,
        tableId,
        columns: {
          dimension: ratingHeader,
        },
      }),
    );
  }

  if (departmentHeader && salaryHeader) {
    candidates.push(
      makeHrCandidate({
        recipeType: "group_summary",
        title: `${departmentHeader}별 ${salaryHeader} 요약`,
        tableId,
        columns: {
          dimension: departmentHeader,
          metric: salaryHeader,
        },
      }),
    );
  }

  if (hireDateHeader && salaryHeader) {
    candidates.push(
      makeHrCandidate({
        recipeType: "time_trend",
        title: `${hireDateHeader} 기준 ${salaryHeader} 추이`,
        tableId,
        columns: {
          date: hireDateHeader,
          metric: salaryHeader,
        },
      }),
    );
  }

  if (salaryHeader) {
    candidates.push(
      makeHrCandidate({
        recipeType: "top_bottom",
        title: `${salaryHeader} 상위/하위 항목`,
        tableId,
        columns: {
          dimension: personHeader || departmentHeader || salaryHeader,
          metric: salaryHeader,
        },
      }),
    );
  }

  if (!candidates.length) {
    return executeTemplateSections({
      normalizedQueryTables,
      templateCandidate,
    });
  }

  return executeTemplateSections({
    normalizedQueryTables,
    templateCandidate: {
      ...templateCandidate,
      candidates,
    },
  });
}

function executeBusinessTemplate({
  normalizedQueryTables = [],
  templateCandidate = {},
}) {
  const templateId = templateCandidate.templateId;

  if (!templateId) {
    return {
      ok: false,
      code: "BUSINESS_TEMPLATE_ID_REQUIRED",
      message: "templateId가 필요합니다.",
    };
  }

  let sections = [];

  switch (templateId) {
    case "hr_monthly_report":
      sections = executeHrMonthlyReport({
        normalizedQueryTables,
        templateCandidate,
      });
      break;

    case "research_budget_report":
      sections = executeResearchBudgetReport({
        normalizedQueryTables,
        templateCandidate,
      });
      break;

    case "sales_report":
      sections = executeSalesReport({
        normalizedQueryTables,
        templateCandidate,
      });
      break;

    default:
      sections = executeTemplateSections({
        normalizedQueryTables,
        templateCandidate,
      });
      break;
  }

  if (!sections.length) {
    return {
      ok: false,
      code: "BUSINESS_TEMPLATE_EXECUTION_EMPTY",
      message: "실행 가능한 템플릿 섹션이 없습니다.",
    };
  }

  return {
    ok: true,
    resultType: "businessTemplate",
    templateId,
    title: templateCandidate.title || templateId,
    description: templateCandidate.description || "",
    sections,
  };
}

module.exports = {
  executeBusinessTemplate,
};
