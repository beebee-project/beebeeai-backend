const formulaUtils = require("./formulaUtils");

function toRef(sheetName, header, meta, sheetInfo) {
  if (!meta || !sheetInfo) return null;
  const columnLetter = meta.columnLetter;
  const startRow = sheetInfo.startRow || meta.startRow;
  const lastDataRow = sheetInfo.lastDataRow || meta.lastRow;

  return {
    sheetName,
    header,
    columnLetter,
    startRow,
    lastDataRow,
    cell: `'${sheetName}'!${columnLetter}${startRow}`,
    range: `'${sheetName}'!${columnLetter}${startRow}:${columnLetter}${lastDataRow}`,
  };
}

function pickBestColumnAnySheet(ctx, headerLike, role = "lookup") {
  if (!ctx?.allSheetsData || !headerLike) return null;
  const terms = formulaUtils.expandTermsFromText(String(headerLike));
  const hit = formulaUtils.findBestColumnAcrossSheets(
    ctx.allSheetsData,
    terms,
    role,
  );
  if (!hit) return null;

  return {
    sheetName: hit.sheetName,
    header: hit.header,
    columnLetter: hit.columnLetter,
    startRow: hit.startRow,
    lastDataRow: hit.lastDataRow,
    cell: `'${hit.sheetName}'!${hit.columnLetter}${hit.startRow}`,
    range: `'${hit.sheetName}'!${hit.columnLetter}${hit.startRow}:${hit.columnLetter}${hit.lastDataRow}`,
    score: hit.score,
    isAmbiguous: !!hit.isAmbiguous,
    ambiguityGap: hit.ambiguityGap ?? null,
    runnerUpHeader: hit.runnerUpHeader ?? null,
  };
}

function pickBestColumnInSheet(ctx, headerLike, sheetName, role = "lookup") {
  if (!ctx?.allSheetsData || !headerLike || !sheetName) return null;
  const info = ctx.allSheetsData[sheetName];
  if (!info?.metaData) return null;

  const terms = formulaUtils.expandTermsFromText(String(headerLike));
  let winner = null;

  for (const [header, meta] of Object.entries(info.metaData)) {
    const score =
      formulaUtils.norm(headerLike) === formulaUtils.norm(header) ? 999 : 0;
    const partial = [...terms].some((t) => {
      const h = formulaUtils.norm(header);
      return h.includes(t) || t.includes(h);
    });

    const finalScore = score || (partial ? 50 : 0);
    if (!finalScore) continue;

    const ref = toRef(sheetName, header, meta, info);
    if (!winner || finalScore > winner.score) {
      winner = { ...ref, score: finalScore };
    }
  }

  return winner;
}

function listColumnsInSheet(ctx, sheetName) {
  const info = ctx?.allSheetsData?.[sheetName];
  if (!info?.metaData) return [];
  return Object.entries(info.metaData).map(([header, meta]) =>
    toRef(sheetName, header, meta, info),
  );
}

function listColumnsAnySheet(ctx) {
  const out = [];
  for (const sheetName of Object.keys(ctx?.allSheetsData || {})) {
    out.push(...listColumnsInSheet(ctx, sheetName));
  }
  return out;
}

function isDerivedSheetName(sheetName = "") {
  const s = String(sheetName || "");
  return /(요약|summary|result|결과|pivot|집계)/i.test(s);
}

function getPreferredSheetNames(ctx, preferredBaseSheet = null) {
  const names = Object.keys(ctx?.allSheetsData || {});
  const ordered = [];

  if (preferredBaseSheet && names.includes(preferredBaseSheet)) {
    ordered.push(preferredBaseSheet);
  }

  for (const n of names) {
    if (ordered.includes(n)) continue;
    if (!isDerivedSheetName(n)) ordered.push(n);
  }

  for (const n of names) {
    if (ordered.includes(n)) continue;
    ordered.push(n);
  }

  return ordered;
}

function pickNumericColumnInSheet(ctx, sheetName, preferredHint = null) {
  const cols = listColumnsInSheet(ctx, sheetName);
  if (!cols.length) return null;

  let best = null;
  for (const c of cols) {
    const info = ctx.allSheetsData?.[c.sheetName];
    const meta = info?.metaData?.[c.header];
    const numericRatio = Number(meta?.numericRatio || 0);
    let score = numericRatio;

    if (preferredHint) {
      const normH = formulaUtils.norm(c.header);
      const normP = formulaUtils.norm(preferredHint);
      if (normH === normP) score += 1000;
      else if (normH.includes(normP) || normP.includes(normH)) score += 100;
    }

    if (!best || score > best.score) {
      best = { ...c, score };
    }
  }
  return best;
}

function pickNumericColumnAnySheet(
  ctx,
  preferredHint = null,
  preferredBaseSheet = null,
) {
  const sheets = getPreferredSheetNames(ctx, preferredBaseSheet);
  let best = null;
  for (const s of sheets) {
    const hit = pickNumericColumnInSheet(ctx, s, preferredHint);
    if (!hit) continue;
    if (isDerivedSheetName(hit.sheetName)) {
      hit.score -= 500;
    }
    if (!best || hit.score > best.score) best = hit;
  }
  return best;
}

function looksLikeDateHeader(header = "") {
  const h = String(header || "");
  return /(일자|날짜|date|time|월|년|입사|시작|종료)/i.test(h);
}

function pickDateColumnInSheet(ctx, sheetName, preferredHint = null) {
  const cols = listColumnsInSheet(ctx, sheetName);
  let best = null;
  for (const c of cols) {
    let score = 0;
    if (looksLikeDateHeader(c.header)) score += 100;

    if (preferredHint) {
      const normH = formulaUtils.norm(c.header);
      const normP = formulaUtils.norm(preferredHint);
      if (normH === normP) score += 1000;
      else if (normH.includes(normP) || normP.includes(normH)) score += 100;
    }

    if (!score) continue;
    if (!best || score > best.score) best = { ...c, score };
  }
  return best;
}

function pickDateColumnAnySheet(
  ctx,
  preferredHint = null,
  preferredBaseSheet = null,
) {
  const sheets = getPreferredSheetNames(ctx, preferredBaseSheet);
  let best = null;
  for (const s of sheets) {
    const hit = pickDateColumnInSheet(ctx, s, preferredHint);
    if (!hit) continue;
    if (isDerivedSheetName(hit.sheetName)) {
      hit.score -= 500;
    }
    if (!best || hit.score > best.score) best = hit;
  }
  return best;
}

function resolveBaseSheet(ctx, schema) {
  const candidates = [];

  if (schema.lookup?.key_header) {
    const c = pickBestColumnAnySheet(ctx, schema.lookup.key_header, "lookup");
    if (c) candidates.push(c);
  }

  for (const rf of schema.return_fields || []) {
    const c = pickBestColumnAnySheet(ctx, rf, "return");
    if (c) candidates.push(c);
  }

  if (schema.header_hint) {
    const c =
      pickBestColumnAnySheet(ctx, schema.header_hint, "return") ||
      pickNumericColumnAnySheet(ctx, schema.header_hint) ||
      pickDateColumnAnySheet(ctx, schema.header_hint);
    if (c) candidates.push(c);
  }

  if (schema.return_hint) {
    const c = pickBestColumnAnySheet(ctx, schema.return_hint, "return");
    if (c) candidates.push(c);
  }

  if (schema.group_by) {
    const c = pickBestColumnAnySheet(ctx, schema.group_by, "group");
    if (c) candidates.push(c);
  }

  for (const f of schema.filters || []) {
    if (f?.header) {
      const c = pickBestColumnAnySheet(ctx, f.header, "filter");
      if (c) candidates.push(c);
    } else if (f?.role === "date_filter") {
      const c = pickDateColumnAnySheet(
        ctx,
        schema.header_hint || null,
        ctx.bestReturn?.sheetName || null,
      );
      if (c) candidates.push(c);
    } else if (f?.role === "metric_filter") {
      const c = pickNumericColumnAnySheet(
        ctx,
        schema.header_hint || null,
        ctx.bestReturn?.sheetName || null,
      );
      if (c) candidates.push(c);
    }
  }

  if (!candidates.length) {
    return ctx.bestReturn?.sheetName || ctx.bestLookup?.sheetName || null;
  }

  const scoreBySheet = new Map();
  for (const c of candidates) {
    scoreBySheet.set(
      c.sheetName,
      (scoreBySheet.get(c.sheetName) || 0) + (c.score || 1),
    );
  }

  return [...scoreBySheet.entries()].sort((a, b) => b[1] - a[1])[0][0];
}

function resolveReturnColumns(ctx, schema, baseSheet) {
  const out = [];
  const seen = new Set();
  const op = String(schema?.operation || "").toLowerCase();
  const hasExplicitReturn =
    Array.isArray(schema.return_fields) && schema.return_fields.length > 0;
  const returnRole = ["average", "sum", "stdev", "min", "max"].includes(op)
    ? op
    : "return";

  // 이름 목록 요청은 "이름" 열을 최우선으로 resolved
  if (String(schema?.return_role || "") === "entity_name") {
    const nameCol =
      pickBestColumnInSheet(ctx, "이름", baseSheet, "return") ||
      pickBestColumnAnySheet(ctx, "이름", "return");

    if (nameCol) {
      const key = `${nameCol.sheetName}::${nameCol.header}::${nameCol.columnLetter}`;
      if (!seen.has(key)) {
        seen.add(key);
        out.push(nameCol);
      }
      return out;
    }
  }
  if (op === "count" && !hasExplicitReturn) {
    return out;
  }

  for (const rf of schema.return_fields || []) {
    const inBase = pickBestColumnInSheet(ctx, rf, baseSheet, returnRole);
    const any = inBase || pickBestColumnAnySheet(ctx, rf, returnRole);
    if (!any) continue;

    const key = `${any.sheetName}::${any.header}::${any.columnLetter}`;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(any);
  }

  if (!out.length && schema.header_hint && op !== "count") {
    const hinted =
      pickBestColumnInSheet(ctx, schema.header_hint, baseSheet, returnRole) ||
      pickBestColumnAnySheet(ctx, schema.header_hint, returnRole) ||
      (["average", "sum", "min", "max", "median", "stdev", "var_s"].includes(op)
        ? pickNumericColumnInSheet(ctx, baseSheet, schema.header_hint) ||
          pickNumericColumnAnySheet(ctx, schema.header_hint)
        : null);

    if (hinted) {
      const key = `${hinted.sheetName}::${hinted.header}::${hinted.columnLetter}`;
      if (!seen.has(key)) {
        seen.add(key);
        out.push(hinted);
      }
    }
  }

  if (
    !out.length &&
    !hasExplicitReturn &&
    ["average", "sum", "min", "max", "median", "stdev", "var_s"].includes(op)
  ) {
    const numeric =
      pickNumericColumnInSheet(ctx, baseSheet, schema.header_hint || null) ||
      pickNumericColumnAnySheet(
        ctx,
        schema.header_hint || null,
        baseSheet || ctx.bestReturn?.sheetName || null,
      );
    if (numeric) {
      const key = `${numeric.sheetName}::${numeric.header}::${numeric.columnLetter}`;
      if (!seen.has(key)) {
        seen.add(key);
        out.push(numeric);
      }
    }
  }

  if (
    !out.length &&
    ctx.bestReturn &&
    op !== "count" &&
    op !== "filter" &&
    String(schema?.return_role || "") !== "entity_name"
  ) {
    const fallback = {
      ...ctx.bestReturn,
      cell: `'${ctx.bestReturn.sheetName}'!${ctx.bestReturn.columnLetter}${ctx.bestReturn.startRow}`,
      range: `'${ctx.bestReturn.sheetName}'!${ctx.bestReturn.columnLetter}${ctx.bestReturn.startRow}:${ctx.bestReturn.columnLetter}${ctx.bestReturn.lastDataRow}`,
    };

    const key = `${fallback.sheetName}::${fallback.header}::${fallback.columnLetter}`;
    if (!seen.has(key)) {
      seen.add(key);
      out.push(fallback);
    }
  }

  return out;
}

function resolveLookupColumn(ctx, schema, baseSheet) {
  if (!schema.lookup?.key_header) return null;
  return (
    pickBestColumnInSheet(ctx, schema.lookup.key_header, baseSheet, "lookup") ||
    pickBestColumnAnySheet(ctx, schema.lookup.key_header, "lookup")
  );
}

function resolveGroupColumn(ctx, schema, baseSheet) {
  if (!schema.group_by) return null;
  return (
    pickBestColumnInSheet(ctx, schema.group_by, baseSheet, "group") ||
    pickBestColumnAnySheet(ctx, schema.group_by, "group")
  );
}

function resolveSortColumn(ctx, schema, baseSheet) {
  const explicitSortHeader = schema.sort?.header || null;
  const op = String(schema?.operation || "").toLowerCase();
  const sortHint =
    explicitSortHeader ||
    (op === "maxrow" || op === "minrow" || op === "topnrows" || op === "sortby"
      ? schema.header_hint || null
      : null);

  if (!sortHint) return null;

  return (
    pickBestColumnInSheet(ctx, sortHint, baseSheet, "sort") ||
    pickBestColumnAnySheet(ctx, sortHint, "sort") ||
    pickNumericColumnInSheet(ctx, baseSheet, sortHint) ||
    pickNumericColumnAnySheet(ctx, sortHint, baseSheet) ||
    pickDateColumnInSheet(ctx, baseSheet, sortHint) ||
    pickDateColumnAnySheet(ctx, sortHint, baseSheet)
  );
}

function resolveFilterColumns(ctx, schema, baseSheet) {
  const out = [];
  const seen = new Set();

  const pushUnique = (item) => {
    if (!item) return;
    const key = JSON.stringify([
      item.logical_operator || "",
      item.header || "",
      item.operator || "",
      item.value ?? "",
      item.min ?? "",
      item.max ?? "",
      item.value_type || "",
      item.ref?.sheetName || "",
      item.ref?.columnLetter || "",
    ]);
    if (seen.has(key)) return;
    seen.add(key);
    out.push(item);
  };

  for (const f of schema.filters || []) {
    if (f?.logical_operator && Array.isArray(f.conditions)) {
      const innerSeen = new Set();
      const inner = f.conditions
        .map((x) => {
          const ref =
            pickBestColumnInSheet(ctx, x.header, baseSheet, "filter") ||
            pickBestColumnAnySheet(ctx, x.header, "filter");
          const item = { ...x, ref };
          const innerKey = JSON.stringify([
            item.header || "",
            item.operator || "",
            item.value ?? "",
            item.min ?? "",
            item.max ?? "",
            item.value_type || "",
            item.ref?.sheetName || "",
            item.ref?.columnLetter || "",
          ]);
          if (innerSeen.has(innerKey)) return null;
          innerSeen.add(innerKey);
          return item;
        })
        .filter(Boolean);
      pushUnique({ logical_operator: f.logical_operator, conditions: inner });
      continue;
    }

    let ref = null;

    if (f?.header) {
      ref =
        pickBestColumnInSheet(ctx, f.header, baseSheet, "filter") ||
        pickBestColumnAnySheet(ctx, f.header, "filter");
    } else if (f?.role === "date_filter") {
      ref =
        pickDateColumnInSheet(ctx, baseSheet, schema.header_hint || null) ||
        pickDateColumnAnySheet(ctx, schema.header_hint || null);
    } else if (f?.role === "metric_filter") {
      ref =
        pickNumericColumnInSheet(
          ctx,
          baseSheet,
          f.header || schema.header_hint || null,
        ) ||
        pickNumericColumnAnySheet(
          ctx,
          f.header || schema.header_hint || null,
          baseSheet || ctx.bestReturn?.sheetName || null,
        );
    }

    pushUnique({ ...f, ref });
  }

  return out;
}

function resolveIntent(ctx) {
  const schema = ctx.intent || {};
  const baseSheet = resolveBaseSheet(ctx, schema);

  const resolved = {
    platform: ctx.engine || "excel",
    baseSheet,
    returnColumns: resolveReturnColumns(ctx, schema, baseSheet),
    lookupColumn: resolveLookupColumn(ctx, schema, baseSheet),
    groupColumn: resolveGroupColumn(ctx, schema, baseSheet),
    sortColumn: resolveSortColumn(ctx, schema, baseSheet),
    filterColumns: resolveFilterColumns(ctx, schema, baseSheet),
    ambiguities: [],
  };

  const candidates = [
    ...resolved.returnColumns,
    resolved.lookupColumn,
    resolved.groupColumn,
    resolved.sortColumn,
  ].filter(Boolean);

  for (const c of candidates) {
    if (c.isAmbiguous) {
      resolved.ambiguities.push({
        header: c.header,
        sheetName: c.sheetName,
        gap: c.ambiguityGap ?? null,
        runnerUpHeader: c.runnerUpHeader ?? null,
      });
    }
  }

  return resolved;
}

function buildResolvedContext(ctx, resolved) {
  return {
    ...ctx,
    resolved,
    bestReturn: resolved?.returnColumns?.[0] || ctx.bestReturn || null,
    bestLookup: resolved?.lookupColumn || ctx.bestLookup || null,
  };
}

module.exports = {
  resolveIntent,
  buildResolvedContext,
};
