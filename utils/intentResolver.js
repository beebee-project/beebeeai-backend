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

function _isStopToken(token = "") {
  const t = _normToken(token);
  if (!t) return true;
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

  const tokens = _expandRawTokens(raw).filter((t) => !_isStopToken(t));
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

function _inferDateFiltersByTypedColumns(ctx, schema, baseSheet) {
  const raw = _rawText(schema);
  if (!raw || !ctx?.allSheetsData) return [];

  const out = [];
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

  const dateCols = _collectAllMetaColumns(ctx, baseSheet).filter((c) =>
    _isDateColumn(c.meta, c.header),
  );

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
  const existing = Array.isArray(schema.filters) ? schema.filters : [];
  return _dedupeFilters([
    ...existing,
    ..._inferTextOperatorFilters(ctx, schema, baseSheet),
    ..._inferFiltersBySampleValues(ctx, schema, baseSheet),
    ..._inferNumericFiltersByTypedColumns(ctx, schema, baseSheet),
    ..._inferDateFiltersByTypedColumns(ctx, schema, baseSheet),
  ]);
}

function resolveFilterColumns(ctx, schema, baseSheet) {
  const out = [];
  const seen = new Set();
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
    if (f?.logical_operator && Array.isArray(f.conditions)) {
      const innerSeen = new Set();
      const inner = f.conditions
        .map((x) => {
          const ref =
            pickBestColumnInSheet(ctx, x.header, baseSheet, "filter") ||
            pickBestColumnAnySheet(ctx, x.header, "filter");
          const item = { ...x, ref: _normalizeRefRange(ref, ctx) };
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

    pushUnique({ ...f, ref: _normalizeRefRange(ref, ctx) });
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
