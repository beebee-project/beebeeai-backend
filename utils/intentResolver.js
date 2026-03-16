const formulaUtils = require("./formulaUtils");

function toRef(sheetName, header, meta, sheetInfo) {
  if (!meta || !sheetInfo) return null;
  const columnLetter = meta.columnLetter;
  const startRow = meta.startRow || sheetInfo.startRow;
  const lastDataRow = meta.lastRow || sheetInfo.lastDataRow;

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

  if (schema.group_by) {
    const c = pickBestColumnAnySheet(ctx, schema.group_by, "lookup");
    if (c) candidates.push(c);
  }

  for (const f of schema.filters || []) {
    if (f?.header) {
      const c = pickBestColumnAnySheet(ctx, f.header, "lookup");
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
  for (const rf of schema.return_fields || []) {
    const inBase = pickBestColumnInSheet(ctx, rf, baseSheet, "return");
    const any = inBase || pickBestColumnAnySheet(ctx, rf, "return");
    if (any) out.push(any);
  }

  if (!out.length && ctx.bestReturn) {
    out.push({
      ...ctx.bestReturn,
      cell: `'${ctx.bestReturn.sheetName}'!${ctx.bestReturn.columnLetter}${ctx.bestReturn.startRow}`,
      range: `'${ctx.bestReturn.sheetName}'!${ctx.bestReturn.columnLetter}${ctx.bestReturn.startRow}:${ctx.bestReturn.columnLetter}${ctx.bestReturn.lastDataRow}`,
    });
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
    pickBestColumnInSheet(ctx, schema.group_by, baseSheet, "lookup") ||
    pickBestColumnAnySheet(ctx, schema.group_by, "lookup")
  );
}

function resolveSortColumn(ctx, schema, baseSheet) {
  if (!schema.sort?.header) return null;
  return (
    pickBestColumnInSheet(ctx, schema.sort.header, baseSheet, "lookup") ||
    pickBestColumnAnySheet(ctx, schema.sort.header, "lookup")
  );
}

function resolveFilterColumns(ctx, schema, baseSheet) {
  const out = [];

  for (const f of schema.filters || []) {
    if (f?.logical_operator && Array.isArray(f.conditions)) {
      const inner = f.conditions.map((x) => {
        const ref =
          pickBestColumnInSheet(ctx, x.header, baseSheet, "lookup") ||
          pickBestColumnAnySheet(ctx, x.header, "lookup");
        return { ...x, ref };
      });
      out.push({ logical_operator: f.logical_operator, conditions: inner });
      continue;
    }

    const ref =
      pickBestColumnInSheet(ctx, f.header, baseSheet, "lookup") ||
      pickBestColumnAnySheet(ctx, f.header, "lookup");

    out.push({ ...f, ref });
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
