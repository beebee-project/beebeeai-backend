const formulaUtils = require("./formulaUtils");

function _isRedundantQuotedEqualityFilter(f, raw = "") {
  const quoted = _extractQuotedLiteral(raw);
  if (!quoted || !_inferTextOperator(raw)) return false;

  const op = String(f?.operator || "=").toLowerCase();
  if (op !== "=" && op !== "==") return false;

  const v = String(f?.value ?? "").trim();
  return _normToken(v) === _normToken(quoted);
}

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
    ambiguityReason: hit.ambiguityReason || "",
  };
}

function _ambiguityIssue(c, role = "unknown") {
  if (!c?.isAmbiguous) return null;
  return {
    role,
    header: c.header || null,
    sheetName: c.sheetName || null,
    gap: c.ambiguityGap ?? null,
    runnerUpHeader: c.runnerUpHeader ?? null,
    message:
      c.ambiguityReason ||
      `후보 열이 모호합니다: ${c.header || "알 수 없음"} / ${
        c.runnerUpHeader || "다른 후보"
      }`,
  };
}

function _collectAmbiguities(resolved = {}) {
  const out = [];

  for (const c of resolved.returnColumns || []) {
    const issue = _ambiguityIssue(c, "return");
    if (issue) out.push(issue);
  }

  for (const [role, c] of [
    ["lookup", resolved.lookupColumn],
    ["group", resolved.groupColumn],
    ["sort", resolved.sortColumn],
  ]) {
    const issue = _ambiguityIssue(c, role);
    if (issue) out.push(issue);
  }

  for (const f of resolved.filterColumns || []) {
    const ref = f?.ref || f;
    const issue = _ambiguityIssue(ref, "filter");
    if (issue) out.push(issue);

    if (f?.logical_operator && Array.isArray(f.conditions)) {
      for (const sub of f.conditions) {
        const subIssue = _ambiguityIssue(sub?.ref || sub, "filter");
        if (subIssue) out.push(subIssue);
      }
    }
  }

  const seen = new Set();
  return out.filter((x) => {
    const key = `${x.role}|${x.sheetName}|${x.header}|${x.runnerUpHeader}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
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

function pickBestColumnInTableBlock(
  ctx,
  headerLike,
  tableBlock,
  role = "lookup",
) {
  if (!ctx?.allSheetsData || !headerLike || !tableBlock?.sheetName) return null;

  const terms = formulaUtils.expandTermsFromText(String(headerLike));
  const sheetName = tableBlock.sheetName;
  const info = ctx.allSheetsData[sheetName];
  if (!info?.metaData) return null;

  let winner = null;

  for (const col of tableBlock.columns || []) {
    const header = String(col.header || "").trim();
    if (!header) continue;

    const score =
      formulaUtils.norm(headerLike) === formulaUtils.norm(header) ? 999 : 0;

    const partial = [...terms].some((t) => {
      const h = formulaUtils.norm(header);
      return h.includes(t) || t.includes(h);
    });

    const finalScore = score || (partial ? 60 : 0);
    if (!finalScore) continue;

    const meta = info.metaData[header] || {
      columnLetter: col.columnLetter,
      startRow: tableBlock.dataStartRow,
      lastRow: tableBlock.dataEndRow,
    };

    const ref = {
      sheetName,
      header,
      columnLetter: col.columnLetter,
      startRow: tableBlock.dataStartRow,
      lastDataRow: tableBlock.dataEndRow,
      cell: `'${sheetName}'!${col.columnLetter}${tableBlock.dataStartRow}`,
      range: `'${sheetName}'!${col.columnLetter}${tableBlock.dataStartRow}:${col.columnLetter}${tableBlock.dataEndRow}`,
      score: finalScore,
      tableId: tableBlock.tableId,
    };

    if (!winner || ref.score > winner.score) {
      winner = ref;
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

function _looksLikeEntityNameReturn(schema = {}) {
  const raw = _rawText(schema);
  const op = String(schema?.operation || "").toLowerCase();

  if (!["maxrow", "minrow", "topnrows"].includes(op)) return false;
  return /(이름|성명|\bname\b)/i.test(raw);
}

function _pickEntityNameColumn(ctx, schema, baseSheet) {
  const raw = _rawText(schema);

  // "선수 이름", "직원 이름", "학생 이름"처럼 이름 앞의 대상어를 추출
  const entityMatch = raw.match(
    /([가-힣A-Za-z0-9_]+)\s*(?:의\s*)?(?:이름|성명|\bname\b)/i,
  );
  const entityToken = entityMatch?.[1] ? _normToken(entityMatch[1]) : "";

  const sheets = getPreferredSheetNames(ctx, baseSheet);
  let best = null;

  for (const sheetName of sheets) {
    const cols = listColumnsInSheet(ctx, sheetName);

    for (const c of cols) {
      const h = _normToken(c.header);
      if (!h) continue;

      let score = 0;

      // 직접적인 이름 열
      if (/(이름|성명|name)/i.test(String(c.header || ""))) score += 100;

      const info = ctx?.allSheetsData?.[sheetName];
      const meta = info?.metaData?.[c.header] || {};
      const semanticRole = _metaSemanticRole(meta, c.header);

      if (semanticRole === "entity_name") score += 80;
      if (semanticRole === "id") score -= 40;
      if (semanticRole === "metric") score -= 50;
      if (semanticRole === "date") score -= 50;

      // 대상어 + 명 패턴: 선수명, 직원명, 고객명 등
      if (entityToken && h.includes(entityToken) && /명$/.test(h)) {
        score += 90;
      }

      // 너무 일반적인 집계/분류 열은 감점
      if (/(팀|부서|학과|분류|구분|등급|상태|지역|카테고리)/.test(h)) {
        score -= 30;
      }

      if (score <= 0) continue;
      if (!best || score > best.score) best = { ...c, score };
    }
  }

  return best;
}

function _scoreTableBlockForIntent(ctx, schema = {}, block = {}) {
  const raw = _rawText(schema);
  const rawNorm = _normToken(raw);
  const op = String(schema?.operation || "").toLowerCase();

  let score = 0;

  const headers = (block.columns || []).map((c) => String(c.header || ""));
  const headerNorms = headers.map(_normToken).filter(Boolean);

  const hints = [
    schema.header_hint,
    schema.return_hint,
    schema.group_by,
    schema.lookup?.key_header,
    ...(schema.return_fields || []),
    ...(schema.filters || []).map((f) => f?.header).filter(Boolean),
  ].filter(Boolean);

  for (const h of hints) {
    const nh = _normToken(h);
    if (!nh) continue;

    for (const bh of headerNorms) {
      if (bh === nh) score += 80;
      else if (bh.includes(nh) || nh.includes(bh)) score += 35;
    }
  }

  for (const bh of headerNorms) {
    if (bh && rawNorm.includes(bh)) score += 20;
  }

  if (["sum", "average", "min", "max", "median"].includes(op)) {
    const info = ctx?.allSheetsData?.[block.sheetName];
    const hasMetricColumn = (block.columns || []).some((c) => {
      const meta = info?.metaData?.[c.header] || {};
      return _metaSemanticRole(meta, c.header) === "metric";
    });

    if (hasMetricColumn) score += 15;
  }

  if (op === "count") {
    score += Math.min(block.columns?.length || 0, 10);
  }

  score += Math.min(Number(block.score || 0), 30);

  return score;
}

function selectBestTableBlock(ctx, schema = {}, baseSheet = null) {
  const blocks = [];

  for (const sheetName of getPreferredSheetNames(ctx, baseSheet)) {
    const sheetBlocks = ctx?.allSheetsData?.[sheetName]?.tableBlocks || [];
    for (const block of sheetBlocks) {
      const score = _scoreTableBlockForIntent(ctx, schema, block);
      if (score > 0) {
        blocks.push({ ...block, scoreForIntent: score });
      }
    }
  }

  if (!blocks.length) return null;

  blocks.sort((a, b) => b.scoreForIntent - a.scoreForIntent);
  const best = blocks[0];
  const runnerUp = blocks[1] || null;

  return {
    ...best,
    runnerUpTableId: runnerUp?.tableId || null,
    runnerUpScoreForIntent: runnerUp?.scoreForIntent ?? null,
    tableAmbiguityGap: runnerUp
      ? best.scoreForIntent - runnerUp.scoreForIntent
      : null,
  };
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

  // 이름/성명 요청은 "이름 계열" 열을 최우선으로 resolved
  if (
    String(schema?.return_role || "") === "entity_name" ||
    _looksLikeEntityNameReturn(schema)
  ) {
    const nameCol =
      _pickEntityNameColumn(ctx, schema, baseSheet) ||
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
    const inTable = pickBestColumnInTableBlock(
      ctx,
      rf,
      schema.selectedTableBlock || ctx?.resolved?.selectedTableBlock,
      returnRole,
    );
    const inBase =
      inTable || pickBestColumnInSheet(ctx, rf, baseSheet, returnRole);
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
    pickBestColumnInTableBlock(
      ctx,
      schema.group_by,
      schema.selectedTableBlock || ctx?.resolved?.selectedTableBlock,
      "group",
    ) ||
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
    pickBestColumnInTableBlock(
      ctx,
      sortHint,
      schema.selectedTableBlock || ctx?.resolved?.selectedTableBlock,
      "sort",
    ) ||
    pickBestColumnInSheet(ctx, sortHint, baseSheet, "sort") ||
    pickBestColumnAnySheet(ctx, sortHint, "sort") ||
    pickNumericColumnInSheet(ctx, baseSheet, sortHint) ||
    pickNumericColumnAnySheet(ctx, sortHint, baseSheet) ||
    pickDateColumnInSheet(ctx, baseSheet, sortHint) ||
    pickDateColumnAnySheet(ctx, sortHint, baseSheet)
  );
}

function _normToken(s = "") {
  return formulaUtils.norm
    ? formulaUtils.norm(s)
    : String(s || "")
        .trim()
        .toLowerCase();
}

function _rawText(schema = {}) {
  return String(schema.raw_message || schema.message || schema.prompt || "");
}

function _extractRawTokens(raw = "") {
  return String(raw || "")
    .split(/[^가-힣A-Za-z0-9_.-]+/)
    .map((x) => x.trim())
    .filter((x) => x.length >= 1);
}

function _stripTokenSuffix(token = "") {
  return String(token || "")
    .trim()
    .replace(/(이면서|이지만|이고|이며|인데)$/u, "")
    .replace(/(인|인것|인값|인항목)$/u, "")
    .replace(/(이|가|은|는|을|를|의|로|으로)$/u, "")
    .trim();
}

function _expandRawTokens(raw = "") {
  const out = [];
  const seen = new Set();

  for (const token of _extractRawTokens(raw)) {
    for (const cand of [token, _stripTokenSuffix(token)]) {
      const t = String(cand || "").trim();
      if (!t || seen.has(t)) continue;
      seen.add(t);
      out.push(t);
    }
  }

  return out;
}

function _isNegativeExistenceRaw(raw = "") {
  return /(존재하지\s*않는|존재하지않는|없는|없음)/.test(String(raw || ""));
}

function _isFragmentFromNegativeExistence(value = "") {
  const t = _normToken(value);
  return ["않", "않는", "없는", "없음", "존재하지", "존재하지않는"].includes(t);
}

function _isStopToken(token = "") {
  const t = _normToken(token);
  if (!t) return true;
  if (/^top\d+$/i.test(String(token).trim())) return true;
  if (/^(상위|하위)\d+$/u.test(String(token).trim())) return true;
  if (/^\d+$/.test(t)) return false;
  return [
    "계산",
    "계산해줘",
    "구해줘",
    "보여줘",
    "가져와줘",
    "뽑아줘",
    "목록",
    "리스트",
    "표",
    "전체",
    "해당",
    "있는",
    "에서",
    "이고",
    "이며",
    "이면서",
    "그리고",
    "또는",
    "없는",
    "없음",
    "존재하지",
    "존재하지않는",
    "않",
    "않는",
    "평균",
    "합계",
    "최고",
    "최저",
    "중앙값",
    "직원",
    "사람",
    "인원",
    "수",
    "명",
    "개수",
  ].includes(t);
}

function _sampleValues(meta = {}) {
  return Array.isArray(meta.sampleValues) ? meta.sampleValues : [];
}

function _columnType(meta = {}) {
  return String(meta.dominantType || meta.clusterType || "").toLowerCase();
}

function _profileType(meta = {}) {
  return String(meta.profileType || meta.dominantType || "").toLowerCase();
}

function _inferSemanticRoleFromMeta(header = "", meta = {}) {
  const h = _normToken(header);
  const profileType = _profileType(meta);
  const numericRatio = Number(meta.numericRatio || 0);
  const dateRatio =
    Number(meta.dateRatio || 0) + Number(meta.datetimeRatio || 0);
  const uniqueRatio = Number(meta.uniqueRatio || 0);
  const uniqueCount = Number(meta.uniqueCount || 0);
  const textRatio = Number(meta.textRatio || 0);

  if (dateRatio >= 0.5 || profileType === "date") return "date";

  if (
    numericRatio >= 0.7 &&
    !/(id|코드|번호|사번|학번|환자번호|관리번호)$/i.test(h)
  ) {
    return "metric";
  }

  if (
    /(id|코드|번호|사번|학번|환자번호|관리번호)$/i.test(h) ||
    (textRatio >= 0.5 && uniqueRatio >= 0.8 && uniqueCount >= 5)
  ) {
    return "id";
  }

  if (
    profileType === "category" ||
    (textRatio >= 0.5 &&
      uniqueCount >= 2 &&
      uniqueCount <= 30 &&
      uniqueRatio <= 0.6)
  ) {
    return "category";
  }

  if (/(이름|성명|name|명)$/i.test(String(header || ""))) {
    return "entity_name";
  }

  return "unknown";
}

function _metaSemanticRole(meta = {}, header = "") {
  return (
    meta.inferredRole ||
    meta.semanticRole ||
    meta.clusterRole ||
    _inferSemanticRoleFromMeta(header, meta)
  );
}

function _isNumericColumn(meta = {}) {
  return (
    _columnType(meta) === "number" || Number(meta.numericRatio || 0) >= 0.5
  );
}

function _isDateColumn(meta = {}, header = "") {
  const t = _columnType(meta);
  if (t === "date" || t === "datetime") return true;
  return looksLikeDateHeader(header);
}

function _valueAppearsInSamples(token, meta = {}) {
  const nt = _normToken(token);
  if (!nt || _isStopToken(nt)) return false;
  return _sampleValues(meta).some((v) => _normToken(v) === nt);
}

function _collectAllMetaColumns(ctx, preferredBaseSheet = null) {
  const out = [];
  for (const sheetName of getPreferredSheetNames(ctx, preferredBaseSheet)) {
    const info = ctx?.allSheetsData?.[sheetName];
    if (!info?.metaData) continue;
    for (const [header, meta] of Object.entries(info.metaData)) {
      out.push({
        sheetName,
        info,
        header,
        meta,
        ref: toRef(sheetName, header, meta, info),
      });
    }
  }
  return out.filter((x) => x.ref);
}

function _normalizeRefRange(ref, ctx) {
  if (!ref?.sheetName || !ref?.columnLetter) return ref;
  const info = ctx?.allSheetsData?.[ref.sheetName];
  if (!info) return ref;

  const startRow = info.startRow || ref.startRow;
  const lastDataRow = info.lastDataRow || ref.lastDataRow;

  return {
    ...ref,
    startRow,
    lastDataRow,
    cell: `'${ref.sheetName}'!${ref.columnLetter}${startRow}`,
    range: `'${ref.sheetName}'!${ref.columnLetter}${startRow}:${ref.columnLetter}${lastDataRow}`,
  };
}

function _scoreSampleValueCandidate(raw, token, col) {
  let score = 0;
  const nt = _normToken(token);
  const nh = _normToken(col.header);

  if (!nt || !nh) return 0;

  // 토큰이 실제 sampleValues에 있는 것은 기본 조건
  if (!_valueAppearsInSamples(token, col.meta)) return 0;

  score += 10;

  // 문장에 헤더명이 함께 등장하면 강한 보너스
  if (_normToken(raw).includes(nh)) score += 30;

  // 헤더 동의어/확장어가 문장에 있으면 보너스
  const terms =
    typeof formulaUtils.expandTermsFromText === "function"
      ? [...formulaUtils.expandTermsFromText(col.header)]
      : [nh];
  if (terms.some((t) => t && _normToken(raw).includes(_normToken(t)))) {
    score += 15;
  }

  // 1글자 토큰은 오탐 위험이 높지만,
  // 문장에 해당 컬럼 헤더/동의어 문맥이 있으면 정상 조건으로 인정
  if (nt.length === 1 && score < 25) score -= 20;
  if (nt.length === 1 && score >= 25) score += 10;

  // 숫자/날짜 컬럼은 sample text filter 후보에서 감점
  if (_isNumericColumn(col.meta) || _isDateColumn(col.meta, col.header)) {
    score -= 50;
  }

  const semanticRole = _metaSemanticRole(col.meta, col.header);

  if (semanticRole === "category") score += 10;
  if (semanticRole === "id") score -= 10;

  return score;
}

function _dedupeFilters(filters = []) {
  const out = [];
  const seen = new Set();

  for (const f of filters) {
    if (!f) continue;
    const key = [
      f.sheetName || "",
      f.columnLetter || "",
      f.header || "",
      f.operator || "=",
      String(f.value ?? ""),
      f.value_type || "",
    ].join("::");

    if (seen.has(key)) continue;
    seen.add(key);
    out.push(f);
  }

  return out;
}

function _inferFiltersBySampleValues(ctx, schema, baseSheet) {
  const raw = _rawText(schema);
  if (!raw || !ctx?.allSheetsData) return [];
  if (/(존재하지\s*않는|존재하지않는|없는|없음)/.test(raw)) {
    return [];
  }

  const quotedLiteral = _extractQuotedLiteral(raw);
  const textOperator = _inferTextOperator(raw);

  // quoted literal + text operator 요청에서는 quoted literal 자체만
  // sample equality 후보에서 제외한다.
  // 예: 이름에 "민" 포함 + 영업 부서
  //   - "민"은 contains 조건으로만 처리
  //   - "영업"은 sample equality 조건으로 유지
  const skipSampleToken = quotedLiteral && textOperator ? quotedLiteral : null;

  const tokens = _expandRawTokens(raw).filter((t) => {
    if (_isStopToken(t)) return false;
    if (skipSampleToken && _normToken(t) === _normToken(skipSampleToken)) {
      return false;
    }
    return true;
  });

  const out = [];
  const columns = _collectAllMetaColumns(ctx, baseSheet);

  for (const token of tokens) {
    let best = null;

    for (const col of columns) {
      const score = _scoreSampleValueCandidate(raw, token, col);
      if (score <= 0) continue;

      const ref = _normalizeRefRange(col.ref, ctx);
      const candidate = {
        ...ref,
        header: col.header,
        operator: "=",
        value: token,
        value_type: "text",
        source: "sample_value_match",
        score,
      };

      if (!best || candidate.score > best.score) {
        best = candidate;
      }
    }

    // 1글자 토큰은 명확한 후보만 채택
    if (best && (String(token).length > 1 || best.score >= 25)) {
      out.push(best);
    }
  }

  return _dedupeFilters(out);
}

function _extractQuotedLiteral(raw = "") {
  const m = String(raw || "").match(/["']([^"']+)["']/);
  return m?.[1] ? String(m[1]).trim() : null;
}

function _inferTextOperator(raw = "") {
  const s = String(raw || "");
  if (/(포함|contains|include)/i.test(s)) return "contains";
  if (/(시작|starts?\s*with)/i.test(s)) return "starts_with";
  if (/(끝나는|끝\s*나는|끝|ends?\s*with)/i.test(s)) return "ends_with";
  return null;
}

function _inferTextOperatorFilters(ctx, schema, baseSheet) {
  const raw = _rawText(schema);
  if (!raw || !ctx?.allSheetsData) return [];

  const operator = _inferTextOperator(raw);
  const value = _extractQuotedLiteral(raw);
  if (!operator || !value) return [];

  const columns = _collectAllMetaColumns(ctx, baseSheet).filter(
    (c) => !_isNumericColumn(c.meta) && !_isDateColumn(c.meta, c.header),
  );

  let best = null;
  for (const col of columns) {
    let score = 0;
    const headerNorm = _normToken(col.header);
    const rawNorm = _normToken(raw);

    if (rawNorm.includes(headerNorm)) score += 40;

    const terms =
      typeof formulaUtils.expandTermsFromText === "function"
        ? [...formulaUtils.expandTermsFromText(col.header)]
        : [headerNorm];

    if (terms.some((t) => t && rawNorm.includes(_normToken(t)))) {
      score += 20;
    }

    // 명시 헤더/동의어 문맥이 없는 경우에는 오탐 방지를 위해 채택하지 않음
    if (score <= 0) continue;

    const ref = _normalizeRefRange(col.ref, ctx);
    const candidate = {
      ...ref,
      header: col.header,
      operator,
      value,
      value_type: "text",
      source: "text_operator_match",
      score,
    };

    if (!best || candidate.score > best.score) best = candidate;
  }

  return best ? [best] : [];
}

function _extractCellRefLiteral(raw = "") {
  const m = String(raw || "").match(/\b(\$?[A-Z]{1,3}\$?\d{1,7})\b/i);
  return m?.[1] ? m[1].toUpperCase() : null;
}

function _inferCellRefEqualityFilters(ctx, schema, baseSheet) {
  const raw = _rawText(schema);
  if (!raw || !ctx?.allSheetsData) return [];

  const cellRef = _extractCellRefLiteral(raw);
  if (!cellRef) return [];

  const returnHeaders = new Set(
    (schema.return_fields || []).map((x) => _normToken(x)),
  );

  const columns = _collectAllMetaColumns(ctx, baseSheet).filter(
    (c) =>
      !_isNumericColumn(c.meta) &&
      !_isDateColumn(c.meta, c.header) &&
      !returnHeaders.has(_normToken(c.header)),
  );

  let best = null;
  const rawNorm = _normToken(raw);
  const cellNorm = _normToken(cellRef);

  for (const col of columns) {
    let score = 0;
    const headerNorm = _normToken(col.header);

    if (headerNorm && rawNorm.includes(headerNorm)) score += 40;

    // 셀 참조 바로 주변의 헤더를 조건열로 우선 판단
    // 예: "J3 부서에 해당하는 ..." => 부서 컬럼에 강한 보너스
    const cellIdx = rawNorm.indexOf(cellNorm);
    const headerIdx = headerNorm ? rawNorm.indexOf(headerNorm) : -1;
    if (cellIdx >= 0 && headerIdx >= 0 && Math.abs(headerIdx - cellIdx) <= 12) {
      score += 80;
    }

    const terms =
      typeof formulaUtils.expandTermsFromText === "function"
        ? [...formulaUtils.expandTermsFromText(col.header)]
        : [headerNorm];

    if (terms.some((t) => t && rawNorm.includes(_normToken(t)))) {
      score += 20;
    }

    if (score <= 0) continue;

    const ref = _normalizeRefRange(col.ref, ctx);
    const candidate = {
      ...ref,
      header: col.header,
      operator: "=",
      value: cellRef,
      value_type: "cell",
      source: "cell_ref_equality_match",
      score,
    };

    if (!best || candidate.score > best.score) best = candidate;
  }

  return best ? [best] : [];
}

function _inferNumericFiltersByTypedColumns(ctx, schema, baseSheet) {
  const raw = _rawText(schema);
  if (!raw || !ctx?.allSheetsData) return [];

  const out = [];
  const rules = [
    {
      re: /(\d+(?:,\d{3})*(?:\.\d+)?)\s*(?:[가-힣A-Za-z%]+)?\s*(?:이상|부터|>=|≥)/g,
      op: ">=",
    },
    {
      re: /(\d+(?:,\d{3})*(?:\.\d+)?)\s*(?:[가-힣A-Za-z%]+)?\s*(?:초과|>|보다\s*큰)/g,
      op: ">",
    },
    {
      re: /(\d+(?:,\d{3})*(?:\.\d+)?)\s*(?:[가-힣A-Za-z%]+)?\s*(?:이하|까지|<=|≤)/g,
      op: "<=",
    },
    {
      re: /(\d+(?:,\d{3})*(?:\.\d+)?)\s*(?:[가-힣A-Za-z%]+)?\s*(?:미만|<|보다\s*작은)/g,
      op: "<",
    },
  ];

  const numericCols = _collectAllMetaColumns(ctx, baseSheet).filter((c) =>
    _isNumericColumn(c.meta),
  );

  for (const { re, op } of rules) {
    re.lastIndex = 0;
    let m;
    while ((m = re.exec(raw))) {
      const value = String(m[1]).replace(/,/g, "");

      // 날짜/연도 범위에 사용된 숫자는 metric filter로 오탐하지 않는다.
      // 예: "2022년부터 2023년 사이", "2022-01-01부터 ..."
      if (_isNumericTokenConsumedByDate(raw, value)) continue;

      const hinted =
        numericCols.find((c) => raw.includes(c.header)) ||
        numericCols.find((c) => {
          const key = String(
            c.meta.canonicalKey || c.meta.clusterCandidate || "",
          );
          return key && raw.toLowerCase().includes(key.toLowerCase());
        }) ||
        null;

      const ref =
        hinted ||
        pickNumericColumnInSheet(ctx, baseSheet, schema.header_hint || null) ||
        pickNumericColumnAnySheet(ctx, schema.header_hint || null, baseSheet);

      const normalized = _normalizeRefRange(ref, ctx);
      if (normalized?.header) {
        out.push({
          ...normalized,
          header: normalized.header,
          operator: op,
          value,
          value_type: "number",
          source: "typed_numeric_match",
        });
      }
    }
  }

  return out;
}

function _escapeRegExp(s = "") {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function _normalizeIsoDateToken(s = "") {
  return String(s || "")
    .trim()
    .replace(/[./]/g, "-");
}

function _dateRangeConsumedTokens(raw = "") {
  const s = String(raw || "");
  const out = [];

  const push = (...xs) => {
    for (const x of xs) {
      const v = String(x || "").trim();
      if (v && !out.includes(v)) out.push(v);
    }
  };

  // 2022-01-01부터 2023-12-31 사이
  for (const m of s.matchAll(
    /((?:19|20)\d{2}[-/.]\d{1,2}[-/.]\d{1,2})\s*(?:부터|~|-)\s*((?:19|20)\d{2}[-/.]\d{1,2}[-/.]\d{1,2})\s*(?:까지|사이)?/g,
  )) {
    push(m[1], m[2]);
  }

  // 2022년부터 2023년 사이 / 2022~2023
  for (const m of s.matchAll(
    /((?:19|20)\d{2})\s*년?\s*(?:부터|~|-)\s*((?:19|20)\d{2})\s*년?\s*(?:까지|사이)?/g,
  )) {
    push(m[1], m[2]);
  }

  // 2023년에 / 2023년 입사
  for (const m of s.matchAll(/((?:19|20)\d{2})\s*년(?:에)?/g)) {
    push(m[1]);
  }

  return out;
}

function _isNumericTokenConsumedByDate(raw = "", value = "") {
  const v = String(value || "").trim();
  if (!v) return false;
  return _dateRangeConsumedTokens(raw).some((x) => String(x).includes(v));
}

function _pickDateFilterRef(ctx, schema, baseSheet) {
  return (
    pickDateColumnInSheet(ctx, baseSheet, schema.header_hint || null) ||
    pickDateColumnAnySheet(
      ctx,
      schema.header_hint || null,
      baseSheet || ctx.bestReturn?.sheetName || null,
    )
  );
}

function _inferDateFiltersByTypedColumns(ctx, schema, baseSheet) {
  const raw = _rawText(schema);
  if (!raw || !ctx?.allSheetsData) return [];

  const out = [];
  const dateCols = _collectAllMetaColumns(ctx, baseSheet).filter((c) =>
    _isDateColumn(c.meta, c.header),
  );

  const pushDateRange = (min, max, source = "typed_date_range_match") => {
    const ref =
      dateCols.find((c) => raw.includes(c.header)) ||
      _pickDateFilterRef(ctx, schema, baseSheet);
    const normalized = _normalizeRefRange(ref, ctx);
    if (!normalized?.header) return;

    out.push({
      ...normalized,
      header: normalized.header,
      operator: "between",
      min,
      max,
      value_type: "date",
      source,
    });
  };

  // YYYY-MM-DD부터 YYYY-MM-DD 사이/까지
  const explicitRange = raw.match(
    /((?:19|20)\d{2}[-/.]\d{1,2}[-/.]\d{1,2})\s*(?:부터|~|-)\s*((?:19|20)\d{2}[-/.]\d{1,2}[-/.]\d{1,2})\s*(?:까지|사이)?/,
  );
  if (explicitRange) {
    pushDateRange(
      _normalizeIsoDateToken(explicitRange[1]),
      _normalizeIsoDateToken(explicitRange[2]),
      "typed_date_explicit_range_match",
    );
  }

  // YYYY년부터 YYYY년 사이 / YYYY~YYYY
  const yearRange = raw.match(
    /((?:19|20)\d{2})\s*년?\s*(?:부터|~|-)\s*((?:19|20)\d{2})\s*년?\s*(?:까지|사이)?/,
  );
  if (yearRange) {
    pushDateRange(
      `${yearRange[1]}-01-01`,
      `${yearRange[2]}-12-31`,
      "typed_date_year_range_match",
    );
  }

  // 2023년에 입사한 / 2022년에 입사한
  // 단, 위 range 문장과 중복되면 추가하지 않음
  if (!explicitRange && !yearRange) {
    const yearOnly = raw.match(/((?:19|20)\d{2})\s*년(?:에)?/);
    if (
      yearOnly &&
      /(입사|날짜|일자|date)/i.test(raw) &&
      !/(이후|부터|이전|전|후|까지|>=|<=|>|<)/i.test(raw)
    ) {
      pushDateRange(
        `${yearOnly[1]}-01-01`,
        `${yearOnly[1]}-12-31`,
        "typed_date_year_only_match",
      );
    }
  }

  const rules = [
    {
      re: /((?:19|20)\d{2}[-/.]\d{1,2}[-/.]\d{1,2})\s*(?:이후|부터|>=|≥)/g,
      op: ">=",
    },
    {
      re: /((?:19|20)\d{2}[-/.]\d{1,2}[-/.]\d{1,2})\s*(?:이전|전|<)/g,
      op: "<",
    },
    { re: /((?:19|20)\d{2})\s*년\s*(?:이후|부터|>=|≥)/g, op: ">=", year: true },
    { re: /((?:19|20)\d{2})\s*년\s*(?:이전|전|<)/g, op: "<", year: true },
  ];

  for (const { re, op, year } of rules) {
    re.lastIndex = 0;
    let m;
    while ((m = re.exec(raw))) {
      const value = year ? `${m[1]}-01-01` : String(m[1]).replace(/[./]/g, "-");
      const hinted = dateCols.find((c) => raw.includes(c.header)) || null;
      const ref =
        hinted ||
        pickDateColumnInSheet(ctx, baseSheet, schema.header_hint || null) ||
        pickDateColumnAnySheet(ctx, schema.header_hint || null, baseSheet);

      const normalized = _normalizeRefRange(ref, ctx);
      if (normalized?.header) {
        out.push({
          ...normalized,
          header: normalized.header,
          operator: op,
          value,
          value_type: "date",
          source: "typed_date_match",
        });
      }
    }
  }

  return out;
}

function augmentFiltersFromRaw(ctx, schema, baseSheet) {
  const raw = _rawText(schema);
  const hasQuotedTextOperator = !!(
    _extractQuotedLiteral(raw) && _inferTextOperator(raw)
  );
  const existing = [
    ...(Array.isArray(schema.filters) ? schema.filters : []),
    ...(Array.isArray(schema.filter_specs) ? schema.filter_specs : []),
  ].filter((f) => {
    if (
      _isNegativeExistenceRaw(raw) &&
      String(f?.value_type || "").toLowerCase() === "text" &&
      _isFragmentFromNegativeExistence(f?.value)
    ) {
      return false;
    }

    if (!hasQuotedTextOperator) return true;
    const op = String(f?.operator || "=").toLowerCase();
    const vt = String(f?.value_type || "").toLowerCase();
    // quoted literal + starts/ends/contains 요청에서는
    // 기존 equality text filter가 중복 조건으로 들어오는 것을 차단
    return !(vt === "text" && (op === "=" || op === "=="));
  });
  return _dedupeFilters([
    ...existing,
    ..._inferCellRefEqualityFilters(ctx, schema, baseSheet),
    ..._inferTextOperatorFilters(ctx, schema, baseSheet),
    ..._inferFiltersBySampleValues(ctx, schema, baseSheet),
    ..._inferNumericFiltersByTypedColumns(ctx, schema, baseSheet),
    ..._inferDateFiltersByTypedColumns(ctx, schema, baseSheet),
  ]);
}

function resolveFilterColumns(ctx, schema, baseSheet) {
  const out = [];
  const seen = new Set();
  const raw = _rawText(schema);
  const hasQuotedTextOperator = !!(
    _extractQuotedLiteral(raw) && _inferTextOperator(raw)
  );
  const filters = augmentFiltersFromRaw(ctx, schema, baseSheet);

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

  for (const f of filters) {
    if (
      _isNegativeExistenceRaw(raw) &&
      String(f?.value_type || "").toLowerCase() === "text" &&
      _isFragmentFromNegativeExistence(f?.value)
    ) {
      continue;
    }
    if (f?.logical_operator && Array.isArray(f.conditions)) {
      const innerSeen = new Set();
      const inner = f.conditions
        .map((x) => {
          const ref =
            pickBestColumnInSheet(ctx, x.header, baseSheet, "filter") ||
            pickBestColumnAnySheet(ctx, x.header, "filter");
          const item = { ...x, ref: _normalizeRefRange(ref, ctx) };
          if (_isRedundantQuotedEqualityFilter(item, raw)) return null;
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

    // raw augmentation에서 이미 sheetName/columnLetter가 정해진 경우
    // 다시 pickBestColumn*을 타지 말고 해당 ref를 우선 사용한다.
    // 그래야 H91:H177이 H101:H177처럼 밀리는 문제를 막을 수 있다.
    if (f?.ref?.range) {
      ref = _normalizeRefRange(f.ref, ctx);
    } else if (f?.sheetName && f?.columnLetter) {
      ref = _normalizeRefRange(f, ctx);
    }

    if (!ref && f?.header) {
      ref =
        pickBestColumnInSheet(ctx, f.header, baseSheet, "filter") ||
        pickBestColumnAnySheet(ctx, f.header, "filter");
    } else if (f?.role === "date_filter") {
      ref =
        pickDateColumnInSheet(ctx, baseSheet, schema.header_hint || null) ||
        pickDateColumnAnySheet(ctx, schema.header_hint || null);
    } else if (
      f?.role === "metric_filter" ||
      f?.role === "aggregate_filter" ||
      f?.value_type === "aggregate"
    ) {
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

    const op = String(f?.operator || "=").toLowerCase();
    const vt = String(f?.value_type || "").toLowerCase();
    const source = String(f?.source || "");

    if (_isRedundantQuotedEqualityFilter(f, raw)) {
      continue;
    }

    pushUnique({ ...f, ref: _normalizeRefRange(ref, ctx) });
  }

  return out;
}

function _dedupeDateRangeFilters(filters = []) {
  const ranges = filters.filter(
    (f) =>
      f &&
      String(f.operator || "").toLowerCase() === "between" &&
      f.value_type === "date" &&
      f.min != null &&
      f.max != null,
  );

  if (!ranges.length) return filters;

  const sameHeader = (a, b) =>
    String(a?.header || a?.ref?.header || "").trim() ===
    String(b?.header || b?.ref?.header || "").trim();

  return filters.filter((f) => {
    if (!f || String(f.operator || "").toLowerCase() === "between") {
      return true;
    }

    const op = String(f.operator || "").toLowerCase();
    const val = String(f.value || "")
      .trim()
      .replace(/[./]/g, "-");

    return !ranges.some((r) => {
      if (!sameHeader(f, r)) return false;
      const min = String(r.min || "")
        .trim()
        .replace(/[./]/g, "-");
      const max = String(r.max || "")
        .trim()
        .replace(/[./]/g, "-");

      return (
        (f.value_type === "date" &&
          (op === ">=" || op === "gte") &&
          val === min) ||
        ((op === "<=" || op === "lte") && val === max)
      );
    });
  });
}

function _dedupeOrdinalEqualityFilters(filters = []) {
  const ordinalKeys = new Set();

  const collect = (items = []) => {
    for (const f of items) {
      if (!f) continue;
      if (f.logical_operator && Array.isArray(f.conditions)) {
        collect(f.conditions);
        continue;
      }

      if (f.role === "ordinal_filter" || f.value_type === "ordinal_text") {
        ordinalKeys.add(
          [
            String(f.header || f.header_hint || f.ref?.header || "").trim(),
            String(f.value ?? "").trim(),
          ].join("::"),
        );
      }
    }
  };

  collect(filters);
  if (!ordinalKeys.size) return filters;

  const shouldDrop = (f) => {
    if (!f) return false;

    const key = [
      String(f.header || f.header_hint || f.ref?.header || "").trim(),
      String(f.value ?? "").trim(),
    ].join("::");

    const isPlainEquality =
      String(f.operator || "").trim() === "=" &&
      f.role !== "ordinal_filter" &&
      f.value_type !== "ordinal_text";

    return isPlainEquality && ordinalKeys.has(key);
  };

  return filters
    .map((f) => {
      if (!f) return null;

      if (f.logical_operator && Array.isArray(f.conditions)) {
        const inner = f.conditions.filter((x) => !shouldDrop(x));
        if (!inner.length) return null;
        return { ...f, conditions: inner };
      }

      return shouldDrop(f) ? null : f;
    })
    .filter(Boolean);
}

function _dedupeAggregateTextFilters(filters = []) {
  const aggregateHeaders = new Set();

  const collect = (items = []) => {
    for (const f of items) {
      if (!f) continue;

      if (f.logical_operator && Array.isArray(f.conditions)) {
        collect(f.conditions);
        continue;
      }

      if (f.role === "aggregate_filter" || f.value_type === "aggregate") {
        const header = String(
          f.header || f.header_hint || f.ref?.header || "",
        ).trim();
        if (header) aggregateHeaders.add(header);
      }
    }
  };

  collect(filters);
  if (!aggregateHeaders.size) return filters;

  const shouldDrop = (f) => {
    if (!f) return false;

    // aggregate 자체는 유지
    if (f.role === "aggregate_filter" || f.value_type === "aggregate") {
      return false;
    }

    const header = String(
      f.header || f.header_hint || f.ref?.header || "",
    ).trim();

    // 같은 header에 대한 일반 비교는 제거
    return aggregateHeaders.has(header);
  };

  return filters
    .map((f) => {
      if (!f) return null;

      if (f.logical_operator && Array.isArray(f.conditions)) {
        const inner = f.conditions.filter((x) => !shouldDrop(x));
        if (!inner.length) return null;
        return { ...f, conditions: inner };
      }

      return shouldDrop(f) ? null : f;
    })
    .filter(Boolean);
}

function resolveIntent(ctx) {
  const schema = ctx.intent || {};
  const baseSheet = resolveBaseSheet(ctx, schema);
  const selectedTableBlock = selectBestTableBlock(ctx, schema, baseSheet);
  const schemaForResolve = {
    ...schema,
    selectedTableBlock,
  };
  const filterColumns = _dedupeOrdinalEqualityFilters(
    _dedupeAggregateTextFilters(
      _dedupeDateRangeFilters(resolveFilterColumns(ctx, schema, baseSheet)),
    ),
  );

  const resolved = {
    platform: ctx.engine || "excel",
    baseSheet,
    selectedTableBlock,
    returnColumns: resolveReturnColumns(ctx, schemaForResolve, baseSheet),
    lookupColumn: resolveLookupColumn(ctx, schemaForResolve, baseSheet),
    groupColumn: resolveGroupColumn(ctx, schemaForResolve, baseSheet),
    sortColumn: resolveSortColumn(ctx, schemaForResolve, baseSheet),
    filterColumns: resolveFilterColumns(ctx, schemaForResolve, baseSheet),
    ambiguities: [],
    hasBlockingAmbiguity: false,
  };

  resolved.ambiguities = _collectAmbiguities(resolved);
  resolved.hasBlockingAmbiguity = resolved.ambiguities.some((a) => {
    const op = String(
      schema?.operation || ctx?.intent?.operation || "",
    ).toLowerCase();

    // 목록/필터 요청은 조건열이 명확하면 return 후보 ambiguity만으로 막지 않는다.
    if (op === "filter" && a.role === "return") return false;

    // 월별/연도별 날짜 파생 group은 group_by 후보가 임시로 잡힌 경우가 있어 막지 않는다.
    if (
      a.role === "group" &&
      /(월별|연도별|일별|주별|monthly|yearly|daily|weekly)/i.test(
        String(ctx?.intent?.raw_message || ctx?.message || ""),
      )
    ) {
      return false;
    }

    return true;
  });

  return resolved;
}

function buildResolvedContext(ctx, resolved) {
  return {
    ...ctx,
    resolved,
    selectedTableBlock: resolved?.selectedTableBlock || null,
    bestReturn: resolved?.returnColumns?.[0] || ctx.bestReturn || null,
    bestLookup: resolved?.lookupColumn || ctx.bestLookup || null,
  };
}

module.exports = {
  resolveIntent,
  buildResolvedContext,
};
