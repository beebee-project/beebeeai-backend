function text(value) {
  return String(value ?? "");
}

function isPlainObject(value) {
  return value != null && typeof value === "object" && !Array.isArray(value);
}

function firstRow(rows = []) {
  return Array.isArray(rows) && rows.length && isPlainObject(rows[0])
    ? rows[0]
    : {};
}

function hasTopBottomRows(rows = []) {
  return (rows || []).some((row) => /^(top|bottom)$/i.test(text(row?.type)));
}

function cleanInternalFieldName(field = "") {
  const key = text(field).trim();
  const normalized = key.toLowerCase();

  if (!key) return "값";
  if (normalized === "count" || normalized === "cnt") return "건수";
  if (normalized === "sum" || normalized === "total") return "합계";
  if (normalized === "average" || normalized === "avg" || normalized === "mean") {
    return "평균";
  }
  if (normalized === "max") return "최댓값";
  if (normalized === "min") return "최솟값";
  if (normalized === "median") return "중앙값";
  if (normalized === "value") return "값";
  if (normalized === "label") return "항목";
  if (normalized === "type") return "구분";
  if (normalized === "rowcount") return "행 수";

  return key.replace(/_/g, " ").replace(/\s+/g, " ").trim();
}

function inferMetricLabelFromContext(field = "", context = {}) {
  const key = text(field).toLowerCase();
  const source = text([
    context.title,
    context.sectionTitle,
    context.tableTitle,
    context.metricHeader,
    context.metricLabel,
    context.categoryField,
  ].filter(Boolean).join(" "));

  if (key === "average" || key === "avg" || key === "mean") {
    if (/연봉|급여|salary/i.test(source)) return "평균 연봉 (만원)";
    if (/금액|매출|예산|집행|비용|amount|revenue|budget/i.test(source)) return "평균 금액";
    return "평균";
  }

  if (key === "sum" || key === "total") {
    if (/연봉|급여|salary/i.test(source)) return "연봉 합계 (만원)";
    if (/금액|매출|예산|집행|비용|amount|revenue|budget/i.test(source)) return "합계 금액";
    return "합계";
  }

  if (key === "count" || key === "cnt") {
    if (/인원|직원|부서|직급|입사|퇴사|명단|상태|사원|employee|headcount/i.test(source)) {
      return "인원 수";
    }
    return "건수";
  }

  if (key === "value") {
    if (/연봉|급여|salary/i.test(source)) return "연봉 (만원)";
    if (/금액|매출|예산|집행|비용|amount|revenue|budget/i.test(source)) return "금액";
    return "값";
  }

  return cleanInternalFieldName(field);
}

function getDisplayFieldLabel(field = "", context = {}) {
  const key = text(field);
  if (context?.fieldLabels && context.fieldLabels[key]) return context.fieldLabels[key];
  if (context?.chartSpec?.fieldLabels && context.chartSpec.fieldLabels[key]) {
    return context.chartSpec.fieldLabels[key];
  }
  return inferMetricLabelFromContext(key, context);
}

function inferCategoryLabel(field = "", context = {}) {
  const key = text(field).trim();
  const lower = key.toLowerCase();
  const rows = context.rows || context.sampleRows || [];

  if (!key) return "항목";
  if (lower === "label") {
    if (hasTopBottomRows(rows)) return "이름";
    return "항목";
  }
  if (lower === "type") return "구분";
  return cleanInternalFieldName(key);
}

function buildFieldLabels(fields = [], context = {}) {
  const labels = {};
  for (const field of fields.filter(Boolean)) {
    const key = text(field);
    labels[key] =
      key === context.categoryField
        ? inferCategoryLabel(key, context)
        : getDisplayFieldLabel(key, context);
  }
  return labels;
}

function buildChartTitle({ categoryField, valueField, seriesFields, rows, context } = {}) {
  const ctx = { ...(context || {}), rows };
  const categoryLabel = inferCategoryLabel(categoryField, ctx);
  const metricLabel = valueField
    ? getDisplayFieldLabel(valueField, ctx)
    : Array.isArray(seriesFields) && seriesFields.length === 1
      ? getDisplayFieldLabel(seriesFields[0], ctx)
      : "복수 지표";

  if (hasTopBottomRows(rows) && (text(valueField).toLowerCase() === "value" || text(categoryField).toLowerCase() === "label")) {
    if (/연봉|급여|salary/i.test(text(ctx.metricHeader || ctx.title || ctx.sectionTitle))) {
      return "연봉 상위/하위 항목";
    }
    return `${metricLabel} 상위/하위 항목`;
  }

  if (/입사연월|입사월|입사일/i.test(text(categoryField)) && /건수|인원 수/.test(metricLabel)) {
    return `${categoryLabel}별 입사 건수`;
  }

  if (metricLabel === "복수 지표") return `${categoryLabel} 기준 복수 지표 분석`;
  return `${categoryLabel}별 ${metricLabel}`;
}

function getTopRow(rows = [], valueField = "") {
  if (!Array.isArray(rows) || !rows.length || !valueField) return null;
  return rows
    .filter((row) => Number.isFinite(Number(row?.[valueField])))
    .sort((a, b) => Number(b[valueField]) - Number(a[valueField]))[0] || null;
}

function buildInsightText({ categoryField, valueField, seriesFields, rows, context } = {}) {
  const ctx = { ...(context || {}), rows };
  const valueKey = valueField || (Array.isArray(seriesFields) ? seriesFields[0] : "");
  const top = getTopRow(rows, valueKey);
  if (!top) return "차트 후보 데이터를 보고서에 정리했습니다.";

  const categoryLabel = inferCategoryLabel(categoryField, ctx);
  const metricLabel = getDisplayFieldLabel(valueKey, ctx);
  const categoryValue = text(top?.[categoryField] ?? top?.label ?? top?.항목 ?? "상위 항목");
  const metricValue = formatDisplayValue(top?.[valueKey], valueKey, ctx);

  return `${categoryLabel} 기준 ${metricLabel} 최상위 항목은 ${categoryValue}(${metricValue})입니다.`;
}

function decorateChartSpec(spec = null, result = {}, context = {}) {
  if (!spec) return null;
  const rows = Array.isArray(result?.rows) ? result.rows : Array.isArray(context.rows) ? context.rows : [];
  const categoryField = spec.categoryField;
  const seriesFields = Array.isArray(spec.seriesFields)
    ? spec.seriesFields
    : spec.valueField
      ? [spec.valueField]
      : [];
  const ctx = {
    ...context,
    rows,
    title: context.title || spec.title || result.title,
    sectionTitle: context.sectionTitle || result.title,
    metricHeader: context.metricHeader || result.metric?.header || result.metricHeader,
    categoryField,
  };
  const fieldLabels = buildFieldLabels([categoryField, ...seriesFields], ctx);
  const title = buildChartTitle({
    categoryField,
    valueField: spec.valueField,
    seriesFields,
    rows,
    context: ctx,
  });
  const insight = buildInsightText({
    categoryField,
    valueField: spec.valueField,
    seriesFields,
    rows,
    context: { ...ctx, fieldLabels },
  });

  return {
    ...spec,
    title,
    fieldLabels,
    insight,
  };
}

function decimalPlacesForField(field = "", context = {}) {
  const key = text(field).toLowerCase();
  const source = text([field, context.title, context.sectionTitle, context.metricHeader].filter(Boolean).join(" "));

  if (/율|비율|rate|ratio|증감/.test(source)) return 1;
  if (key === "average" || key === "avg" || key === "mean") return 1;
  return 0;
}

function formatDisplayValue(value, field = "", context = {}) {
  if (value == null || value === "") return "";
  const n = Number(value);
  if (!Number.isFinite(n)) return text(value);

  const fractionDigits = decimalPlacesForField(field, context);
  return new Intl.NumberFormat("ko-KR", {
    maximumFractionDigits: fractionDigits,
    minimumFractionDigits: 0,
  }).format(n);
}

function toDisplayNumber(value, field = "", context = {}) {
  const n = Number(value);
  if (!Number.isFinite(n)) return 0;
  const fractionDigits = decimalPlacesForField(field, context);
  const factor = 10 ** fractionDigits;
  return Math.round(n * factor) / factor;
}

function tableColumnLabels(rows = [], context = {}) {
  const row = firstRow(rows);
  const keys = Object.keys(row || {});
  const labels = {};
  for (const key of keys) {
    labels[key] = key === context.categoryField
      ? inferCategoryLabel(key, { ...context, rows })
      : getDisplayFieldLabel(key, { ...context, rows });
  }
  return labels;
}

function decorateReportSections(sections = []) {
  const list = Array.isArray(sections) ? sections : [];
  const decorated = list.map((section) => {
    if (!isPlainObject(section)) return section;

    if (section.type === "chart") {
      const chartSpec = decorateChartSpec(section.chartSpec, { rows: section.rows || [] }, {
        title: section.title,
        sectionTitle: section.title,
        rows: section.rows || [],
      });
      return {
        ...section,
        title: chartSpec?.title || cleanDisplayTitle(section.title),
        chartSpec,
        insight: chartSpec?.insight || cleanDisplayTitle(section.insight),
      };
    }

    if (section.type === "table") {
      return {
        ...section,
        columnLabels: section.columnLabels || tableColumnLabels(section.rows || [], {
          title: section.title,
          sectionTitle: section.title,
        }),
      };
    }

    if (section.type === "insight") {
      return {
        ...section,
        bullets: (section.bullets || []).map(cleanDisplayTitle),
      };
    }

    return section;
  });

  const chartInsights = decorated
    .filter((section) => section?.type === "chart" && section.insight)
    .map((section) => section.insight);
  const insight = decorated.find((section) => section?.type === "insight");
  if (insight && chartInsights.length) {
    insight.bullets = [...new Set(chartInsights.concat(insight.bullets || []))].slice(0, 8);
  }

  return decorated;
}

function cleanDisplayTitle(value = "") {
  let s = text(value);
  s = s.replace(/\blabel\b/gi, "항목");
  s = s.replace(/\bvalue\b/gi, "값");
  s = s.replace(/\bcount\b/gi, "건수");
  s = s.replace(/\baverage\b|\bavg\b|\bmean\b/gi, "평균");
  s = s.replace(/\bsum\b|\btotal\b/gi, "합계");
  s = s.replace(/\s+/g, " ").trim();
  return s;
}

module.exports = {
  buildChartTitle,
  buildFieldLabels,
  buildInsightText,
  cleanDisplayTitle,
  cleanInternalFieldName,
  decorateChartSpec,
  decorateReportSections,
  formatDisplayValue,
  getDisplayFieldLabel,
  inferCategoryLabel,
  tableColumnLabels,
  toDisplayNumber,
};
