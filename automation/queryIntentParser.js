const {
  findMetricRecipe,
  findCompareRecipe,
  normalizeText,
} = require("./metricRecipeRegistry");

function findDateColumn(columns = [], text = "") {
  const direct = findColumnByText(columns, text);
  if (direct && String(direct.type) === "date") return direct;

  return columns.find((c) => String(c.type) === "date") || null;
}

function findColumnByHints(columns = [], hints = [], excludeKey = null) {
  return (
    columns.find((c) => {
      if (excludeKey && c.key === excludeKey) return false;

      const haystack = normalizeText(
        `${c.header || ""} ${c.originalHeader || ""} ${c.key || ""}`,
      );

      return hints.some((h) => haystack.includes(normalizeText(h)));
    }) || null
  );
}

function detectDerivedGroup(message = "", columns = []) {
  const s = String(message || "");

  if (!s.includes("연도별")) return null;

  const dateCol = findDateColumn(columns, s);
  if (!dateCol) return null;

  return {
    type: "derive",
    fn: "year",
    sourceColumnKey: dateCol.key,
    sourceHeader: dateCol.header,
    outputKey: `${dateCol.key}_year`,
    outputHeader: `${dateCol.header} 연도`,
  };
}

function detectRateSpec(message = "", columns = [], deriveStep = null) {
  const recipe = findMetricRecipe(message);
  if (!recipe) return null;

  const numeratorCol = findColumnByHints(
    columns,
    recipe.numerator?.columnHints || [],
    deriveStep?.sourceColumnKey,
  );

  if (!numeratorCol) {
    return {
      type: "rate",
      unresolved: true,
      reason: "NUMERATOR_COLUMN_NOT_FOUND",
      requiredHints: recipe.numerator?.columnHints || [],
      outputHeader: recipe.outputHeader || "비율",
      multiplier: recipe.multiplier || 100,
    };
  }

  return {
    type: "rate",
    recipeId: recipe.id,
    numerator: {
      type: recipe.numerator?.type || "exists",
      columnKey: numeratorCol.key,
      header: numeratorCol.header,
    },
    denominator: recipe.denominator || { type: "count" },
    outputHeader: recipe.outputHeader || "비율",
    multiplier: recipe.multiplier || 100,
  };
}

function includesAny(s, words = []) {
  return words.some((w) => s.includes(w));
}

function findColumnByText(columns = [], text = "") {
  const s = normalizeText(text);

  return columns.find((c) => {
    const h = normalizeText(c.header || c.originalHeader || c.name || "");

    const k = normalizeText(c.key || c.accessor || c.name || "");

    return s.includes(h) || h.includes(s) || s.includes(k) || k.includes(s);
  });
}

function detectOperation(message = "") {
  const s = String(message);

  if (includesAny(s, ["평균", "average", "avg"])) return "average";
  if (includesAny(s, ["합계", "총합", "sum", "total"])) return "sum";
  if (includesAny(s, ["개수", "몇 명", "몇개", "몇 개", "수"])) return "count";
  if (includesAny(s, ["최대", "최고", "max"])) return "max";
  if (includesAny(s, ["최소", "최저", "min"])) return "min";
  if (includesAny(s, ["목록", "리스트", "보여", "출력"])) return "list";

  return "list";
}

function detectMetricColumn(message = "", columns = [], operation = "") {
  if (operation === "count") return null;

  const direct = findColumnByText(columns, message);
  if (direct && ["number", "date"].includes(String(direct.type))) return direct;

  const numberCols = columns.filter((c) => String(c.type) === "number");
  if (numberCols.length === 1) return numberCols[0];

  return numberCols[0] || null;
}

function detectGroupBy(message = "", columns = []) {
  const m = String(message).match(/(.+?)별/);
  if (!m) return null;

  const hint = m[1].trim();
  return findColumnByText(columns, hint) || null;
}

function detectFilters(message = "", columns = []) {
  const filters = [];
  const s = String(message || "");

  for (const col of columns) {
    const header = String(col.header || "");
    if (!header || !s.includes(header)) continue;

    const escaped = header.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

    const numRe = new RegExp(
      `${escaped}\\s*(?:이|가)?\\s*(\\d+(?:\\.\\d+)?)\\s*(이상|초과|이하|미만)`,
    );
    const numMatch = s.match(numRe);

    if (numMatch) {
      const word = numMatch[2];
      let operator = ">=";
      if (word === "초과") operator = ">";
      if (word === "이하") operator = "<=";
      if (word === "미만") operator = "<";

      filters.push({
        columnKey: col.key,
        header: col.header,
        operator,
        value: Number(numMatch[1]),
        valueType: "number",
      });
      continue;
    }

    const textRe = new RegExp(`${escaped}\\s*(?:이|가)?\\s*([^\\s]+)`);
    const textMatch = s.match(textRe);

    if (textMatch && String(col.type) !== "number") {
      const value = textMatch[1]
        .replace(/인|인$|인\\s*$/g, "")
        .replace(/직원|학생|환자|데이터|목록|수|평균|합계/g, "")
        .trim();

      if (!value || value === "별" || value.endsWith("별")) continue;

      if (value) {
        filters.push({
          columnKey: col.key,
          header: col.header,
          operator: "=",
          value,
          valueType: "text",
        });
      }
    }
  }

  return filters;
}

function detectLimit(message = "") {
  const s = String(message || "");

  const m =
    s.match(/상위\s*(\d+)\s*(개|명|건)?/) ||
    s.match(/top\s*(\d+)/i) ||
    s.match(/(\d+)\s*(개|명|건)\s*(?:만|까지)?/);

  if (!m) return null;

  return {
    type: "limit",
    count: Number(m[1]),
  };
}

function detectDeriveStep(message = "", table = null) {
  const s = String(message || "");

  const columns = table?.columns || [];

  const dateColumn =
    columns.find((c) => c.type === "date") ||
    columns.find((c) =>
      /(일|날짜|date|입사|퇴사|등록|생성)/i.test(c.header || ""),
    );

  if (!dateColumn) return null;

  if (/연월별|월도별|년월별|연월\s*별|년월\s*별/.test(s)) {
    return {
      type: "derive",
      fn: "yearMonth",
      sourceColumnKey: dateColumn.key,
      sourceHeader: dateColumn.header,
      outputKey: `${dateColumn.key}_year_month`,
      outputHeader: `${dateColumn.header} 연월`,
    };
  }

  if (/연도별|년도별|연도\s*별/.test(s)) {
    return {
      type: "derive",
      fn: "year",
      sourceColumnKey: dateColumn.key,
      sourceHeader: dateColumn.header,
      outputKey: `${dateColumn.key}_year`,
      outputHeader: `${dateColumn.header} 연도`,
    };
  }

  if (/분기별|분기\s*별|quarter/i.test(s)) {
    return {
      type: "derive",
      fn: "quarter",
      sourceColumnKey: dateColumn.key,
      sourceHeader: dateColumn.header,
      outputKey: `${dateColumn.key}_quarter`,
      outputHeader: `${dateColumn.header} 분기`,
    };
  }

  if (/월별|개월별|월\s*별/.test(s)) {
    return {
      type: "derive",
      fn: "month",
      sourceColumnKey: dateColumn.key,
      sourceHeader: dateColumn.header,
      outputKey: `${dateColumn.key}_month`,
      outputHeader: `${dateColumn.header} 월`,
    };
  }

  return null;
}

function detectSort(message = "", metric = null) {
  const s = String(message || "");

  if (!metric?.columnKey) return null;

  if (/높은\s*순|내림차순|큰\s*순|상위|top/i.test(s)) {
    return {
      type: "sort",
      by: metric.columnKey,
      header: metric.header,
      direction: "desc",
    };
  }

  if (/낮은\s*순|오름차순|작은\s*순|하위/i.test(s)) {
    return {
      type: "sort",
      by: metric.columnKey,
      header: metric.header,
      direction: "asc",
    };
  }

  return null;
}

function detectCompareOffset(message = "", recipe = {}) {
  const s = normalizeText(message);

  for (const hint of recipe.offsetHints || []) {
    if ((hint.match || []).some((kw) => s.includes(normalizeText(kw)))) {
      return {
        offset: hint.offset,
        unit: hint.unit,
      };
    }
  }

  return {
    offset: 1,
    unit: "previous",
  };
}

function detectCompareStep(message = "", metric = null, groupBySpec = null) {
  const recipe = findCompareRecipe(message);

  if (!recipe) return null;
  if (!metric?.columnKey || !groupBySpec?.columnKey) {
    return {
      type: "compare",
      unresolved: true,
      reason: "COMPARE_REQUIRES_METRIC_AND_GROUP",
      outputHeader: recipe.outputHeader || "증감률",
    };
  }

  const offsetSpec = detectCompareOffset(message, recipe);

  return {
    type: "compare",
    recipeId: recipe.id,
    mode: recipe.mode || "previous",
    method: recipe.method || "growthRate",
    metric,
    groupBy: {
      columnKey: groupBySpec.columnKey,
      header: groupBySpec.header,
    },
    outputHeader: recipe.outputHeader || "증감률",
    multiplier: recipe.multiplier || 100,
    defaultAggregate: recipe.defaultAggregate || "sum",
    offset: offsetSpec.offset,
    offsetUnit: offsetSpec.unit,
  };
}

function detectWindowStep(message = "", metric = null, groupBySpec = null) {
  const s = normalizeText(message);

  if (!metric?.columnKey) return null;

  if (/누적|누계|runningtotal|cumulative/i.test(message)) {
    return {
      type: "window",
      method: "cumulativeSum",
      metric,
      groupBy: groupBySpec
        ? {
            columnKey: groupBySpec.columnKey,
            header: groupBySpec.header,
          }
        : null,
      outputHeader: `누적 ${metric.header || "값"}`,
    };
  }

  const rollingMatch =
    String(message).match(
      /최근\s*(\d+)\s*(개월|월|일|년)\s*(?:평균|이동평균)/,
    ) ||
    String(message).match(
      /(\d+)\s*(개월|월|일|년)\s*(?:이동평균|rolling\s*average)/i,
    );

  if (rollingMatch) {
    const size = Number(rollingMatch[1]);

    return {
      type: "window",
      method: "rollingAverage",
      size,
      unit: rollingMatch[2],
      metric,
      groupBy: groupBySpec
        ? {
            columnKey: groupBySpec.columnKey,
            header: groupBySpec.header,
          }
        : null,
      outputHeader: `${size}${rollingMatch[2]} 이동평균`,
    };
  }

  return null;
}

function scoreTableForMessage(table = {}, message = "") {
  const s = normalizeText(message);

  const columns = table.columns || [];
  let score = Number(table.isPrimary ? 5 : 0);
  score += Number(table.confidence || 0) / 20;

  const tableText = normalizeText(
    `${table.tableName || ""} ${table.sheetName || ""}`,
  );

  if (tableText && s.includes(tableText)) score += 20;

  for (const col of columns) {
    const header = normalizeText(
      `${col.header || ""} ${col.originalHeader || ""} ${col.key || ""}`,
    );

    if (!header) continue;

    if (s.includes(header)) score += 10;
    else if (header.includes(s)) score += 3;
  }

  const groupHint = String(message)
    .match(/(.+?)별/)?.[1]
    ?.trim();
  if (groupHint) {
    const matched = findColumnByText(columns, groupHint);
    if (matched) score += 15;
  }

  return score;
}

function selectBestTable(queryTables = [], message = "") {
  if (!Array.isArray(queryTables) || !queryTables.length) return null;

  const ranked = [...queryTables]
    .map((table) => ({
      table,
      score: scoreTableForMessage(table, message),
    }))
    .sort((a, b) => b.score - a.score);

  const best = ranked[0];
  const second = ranked[1];

  if (second && best.score - second.score <= 2) {
    return {
      ambiguous: true,
      candidates: ranked.slice(0, 3).map((r) => ({
        tableId: r.table.tableId,
        tableName: r.table.tableName,
        sheetName: r.table.sheetName,
        score: r.score,
      })),
    };
  }

  return best.table;
}

function parseQueryIntent(message = "", queryTables = []) {
  const selected = selectBestTable(queryTables, message);

  if (!selected) {
    return {
      ok: false,
      error: "분석 가능한 테이블이 없습니다.",
    };
  }

  if (selected?.ambiguous) {
    return {
      ok: false,
      code: "AMBIGUOUS_TABLE",
      error: "요청에 맞는 테이블을 명확히 선택할 수 없습니다.",
      candidates: selected.candidates,
      message,
    };
  }

  const table = selected;

  const columns = table.columns || [];
  let operation = detectOperation(message);
  const metricColumn = detectMetricColumn(message, columns, operation);
  const groupBy = detectGroupBy(message, columns);
  const filters = detectFilters(message, columns);
  const deriveStep = detectDeriveStep(message, table);
  const rateStep = detectRateSpec(message, columns, deriveStep);

  const isYearlyRateRequest =
    /연도별|년도별|연도\s*별/.test(message) && /율|비율/.test(message);

  if (isYearlyRateRequest && rateStep && !rateStep.unresolved) {
    const baseDeriveStep = deriveStep;

    if (!baseDeriveStep) {
      return {
        ok: false,
        code: "DERIVE_COLUMN_NOT_FOUND",
        error: "연도별 계산에 사용할 날짜 컬럼을 찾을 수 없습니다.",
        message,
      };
    }

    return {
      ok: true,
      version: "query_intent_v1",
      message,
      table: {
        tableId: table.tableId,
        tableName: table.tableName,
        sheetName: table.sheetName,
        confidence: table.confidence,
        isPrimary: !!table.isPrimary,
      },
      operation: "multi",
      metric: null,
      groupBy: {
        columnKey: baseDeriveStep.outputKey,
        header: baseDeriveStep.outputHeader,
        type: "category",
        derived: true,
      },
      filters,
      plan: {
        version: "query_plan_v1",
        tableId: table.tableId,
        pipelines: [
          {
            id: "base_count",
            label: "기준 건수",
            steps: [
              baseDeriveStep,
              {
                type: "groupBy",
                columnKey: baseDeriveStep.outputKey,
                header: baseDeriveStep.outputHeader,
              },
              {
                type: "aggregate",
                operation: "count",
                metric: null,
              },
            ],
          },
          {
            id: "event_count",
            label: rateStep?.outputHeader || "이벤트 건수",
            steps: [
              baseDeriveStep,
              {
                type: "filter",
                filters: [
                  {
                    columnKey: rateStep?.numerator?.columnKey,
                    header: rateStep?.numerator?.header,
                    operator: "exists",
                    valueType: "exists",
                  },
                ],
              },
              {
                type: "groupBy",
                columnKey: baseDeriveStep.outputKey,
                header: baseDeriveStep.outputHeader,
              },
              {
                type: "aggregate",
                operation: "count",
                metric: null,
              },
            ],
          },
        ],
        combine: {
          type: "combineRatio",
          numeratorPipeline: "event_count",
          denominatorPipeline: "base_count",
          operation: "rate",
          outputHeader: rateStep?.outputHeader || "비율",
          multiplier: rateStep?.multiplier || 100,
        },
      },
      derive: baseDeriveStep,
      rate: null,
    };
  }

  if (rateStep?.unresolved) {
    return {
      ok: false,
      code: rateStep.reason,
      error: "비율 계산에 필요한 기준 컬럼을 찾을 수 없습니다.",
      requiredHints: rateStep.requiredHints,
      message,
    };
  }

  if (rateStep) operation = "rate";

  const metric = metricColumn
    ? {
        columnKey: metricColumn.key,
        header: metricColumn.header,
        type: metricColumn.type,
      }
    : null;

  const sortStep = detectSort(message, metric);
  const limitStep = detectLimit(message);

  const groupBySpec = deriveStep
    ? {
        columnKey: deriveStep.outputKey,
        header: deriveStep.outputHeader,
        type: "category",
        derived: true,
      }
    : groupBy
      ? {
          columnKey: groupBy.key,
          header: groupBy.header,
          type: groupBy.type,
        }
      : null;

  const compareStep = detectCompareStep(message, metric, groupBySpec);

  const windowStep = detectWindowStep(message, metric, groupBySpec);

  if (windowStep) operation = windowStep.method;

  if (compareStep?.unresolved) {
    return {
      ok: false,
      code: compareStep.reason,
      error: "비교 계산에는 기준 지표와 그룹 기준이 필요합니다.",
      message,
    };
  }

  if (compareStep) operation = "growthRate";

  const plan = {
    version: "query_plan_v1",
    tableId: table.tableId,
    steps: [
      ...(deriveStep ? [deriveStep] : []),
      ...(filters.length ? [{ type: "filter", filters }] : []),
      ...(groupBySpec
        ? [
            {
              type: "groupBy",
              columnKey: groupBySpec.columnKey,
              header: groupBySpec.header,
            },
          ]
        : []),
      ...(rateStep
        ? [rateStep]
        : [
            {
              type: "aggregate",
              operation: windowStep
                ? includesAny(String(message), ["평균", "average", "avg"])
                  ? "average"
                  : "sum"
                : compareStep
                  ? compareStep.defaultAggregate || "sum"
                  : operation,
              metric,
            },
          ]),
      ...(compareStep ? [compareStep] : []),
      ...(windowStep ? [windowStep] : []),
      ...(sortStep ? [sortStep] : []),
      ...(limitStep ? [limitStep] : []),
    ],
  };

  return {
    ok: true,
    version: "query_intent_v1",
    message,
    table: {
      tableId: table.tableId,
      tableName: table.tableName,
      sheetName: table.sheetName,
      confidence: table.confidence,
      isPrimary: !!table.isPrimary,
    },
    operation,
    metric,
    groupBy: groupBySpec,
    filters,
    plan,
    derive: deriveStep,
    rate: rateStep,
    sort: sortStep,
    limit: limitStep,
    compare: compareStep,
    window: windowStep,
  };
}

module.exports = { parseQueryIntent };
