const {
  findTableForTemplate,
  executeTemplateSections,
} = require("./commonTemplateHelpers");
const {
  buildRosterStatusReportSections,
} = require("../structuralBuilders/rosterStatusReportBuilder");
const {
  buildCategorySummaryReportSections,
} = require("../structuralBuilders/categorySummaryReportBuilder");

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

  const rosterSections = buildRosterStatusReportSections({
    normalizedQueryTables,
    table,
    templateCandidate,
    config: {
      hints: {
        name: [
          "성명",
          "이름",
          "직원명",
          "사원명",
          "참여자",
          "연구원명",
          "name",
        ],
        department: ["부서", "소속", "조직", "팀", "department"],
        position: ["직급", "직위", "직책", "등급", "position", "rank"],
        status: [
          "재직상태",
          "상태",
          "근무상태",
          "이수여부",
          "참여상태",
          "진행상태",
          "status",
        ],
        date: [
          "입사일",
          "입사일자",
          "채용일",
          "시작일",
          "참여시작일",
          "등록일",
          "date",
        ],
        salary: [
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
      },
      totalTitle: "전체 대상 수",
    },
  });

  const categorySections = buildCategorySummaryReportSections({
    normalizedQueryTables,
    table,
    templateCandidate,
    config: {
      metricHints: [
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
      dimensions: [
        {
          sectionId: "department_salary_summary",
          sectionType: "department_salary",
          hints: ["부서", "소속", "조직", "팀", "department"],
        },
      ],
      topBottom: {
        sectionId: "salary_top_bottom",
        sectionType: "salary_top_bottom",
        dimensionHints: ["성명", "이름", "직원명", "사원명", "부서", "소속"],
      },
    },
  });

  const existingIds = new Set(rosterSections.map((s) => s.sectionId));
  const dedupedCategorySections = categorySections.filter(
    (section) => !existingIds.has(section.sectionId),
  );

  const sections = [...rosterSections, ...dedupedCategorySections];

  if (!sections.length) {
    return executeTemplateSections({
      normalizedQueryTables,
      templateCandidate,
    });
  }

  return sections;
}

module.exports = {
  executeHrMonthlyReport,
};
