const {
  findTableForTemplate,
  executeTemplateSections,
} = require("./commonTemplateHelpers");
const {
  buildPeriodMetricReportSections,
} = require("../structuralBuilders/periodMetricReportBuilder");

function executeSalesReport({
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

  const sections = buildPeriodMetricReportSections({
    normalizedQueryTables,
    table,
    templateCandidate,
    config: {
      hints: {
        metric: [
          "순매출액",
          "매출액",
          "판매금액",
          "매출금액",
          "카드매출액",
          "sales",
          "revenue",
        ],
        quantity: ["매출수량", "판매수량", "수량", "quantity", "qty"],
        year: ["연도 구분", "기준년도", "매출연도", "연도", "년도", "year"],
        month: ["월 구분", "기준월", "매출월", "월", "month"],
        period: ["연월", "기간", "기준년월", "매출년월", "period"],
        item: [
          "제품명",
          "상품명",
          "품목",
          "세부사업명",
          "제품류",
          "product",
          "item",
        ],
        category: [
          "제품분류",
          "상품분류",
          "업종",
          "업태",
          "구분",
          "분류",
          "category",
        ],
        group: ["구분", "제품류", "제품분류", "대분류"],
      },
      sectionIds: {
        year: "yearly_sales",
        period: "period_sales",
        month: "monthly_sales",
        quantityMonth: "monthly_quantity",
        topBottom: "top_bottom_sales",
      },
      sectionTypes: {
        year: "yearly_sales",
        period: "period_sales",
        month: "monthly_sales",
        quantityMonth: "monthly_quantity",
        topBottom: "top_bottom_sales",
      },
      dimensions: [
        {
          sectionId: "product_sales",
          sectionType: "product_sales",
          hints: [
            "제품명",
            "상품명",
            "품목",
            "세부사업명",
            "제품류",
            "product",
            "item",
          ],
        },
        {
          sectionId: "category_sales",
          sectionType: "category_sales",
          hints: [
            "제품분류",
            "상품분류",
            "업종",
            "업태",
            "구분",
            "분류",
            "category",
          ],
        },
        {
          sectionId: "region_sales",
          sectionType: "region_sales",
          hints: ["지역", "구", "자치구", "시군구", "region"],
        },
      ],
      rankingDimensionHints: [
        "연월",
        "제품명",
        "상품명",
        "품목",
        "제품분류",
        "업종",
        "지역",
      ],
      averagePerUnit: {
        sectionId: "average_sales_amount",
        sectionType: "average_sales_amount",
        title: "평균 판매금액",
      },
    },
  });

  if (!sections.length) {
    return executeTemplateSections({
      normalizedQueryTables,
      templateCandidate,
    });
  }

  return sections;
}

module.exports = {
  executeSalesReport,
};
