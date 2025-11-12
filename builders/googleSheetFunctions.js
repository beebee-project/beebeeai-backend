const googleSheetFunctionBuilder = {
  // =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
  // 외부 데이터 가져오기 함수
  //-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
  importrange: (ctx, formatValue) => {
    // AI intent: { operation: "importrange", spreadsheet_url: "...", range_string: "Sheet1!A1:B10" }
    const { intent } = ctx;
    return `=IMPORTRANGE(${formatValue(intent.spreadsheet_url)}, ${formatValue(
      intent.range_string
    )})`;
  },

  importhtml: (ctx, formatValue) => {
    // AI intent: { operation: "importhtml", url: "...", query: "table", index: 1 }
    const { intent } = ctx;
    return `=IMPORTHTML(${formatValue(intent.url)}, ${formatValue(
      intent.query
    )}, ${intent.index})`;
  },

  importdata: (ctx, formatValue) =>
    `=IMPORTDATA(${formatValue(ctx.intent.url)})`,

  importxml: (ctx, formatValue) =>
    `=IMPORTXML(${formatValue(ctx.intent.url)}, ${formatValue(
      ctx.intent.xpath_query
    )})`,

  googlefinance: (ctx, formatValue) =>
    `=GOOGLEFINANCE(${(ctx.intent.params || ["GOOG", "price"])
      .map((p) => formatValue(p))
      .join(", ")})`,

  //-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
  // 셀 상태 확인 함수
  //-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
  isformula: (ctx) => `=ISFORMULA(${ctx.intent.target_cell || "A1"})`,

  //-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
  // 정규표현식 함수
  //-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
  regexmatch: (ctx, formatValue) => {
    // AI intent: { operation: "regexmatch", text: "A2", regex: "^prefix" }
    const { intent } = ctx;
    return `=REGEXMATCH(${intent.text || "A2"}, ${formatValue(intent.regex)})`;
  },

  regexextract: (ctx, formatValue) => {
    const { intent } = ctx;
    return `=REGEXEXTRACT(${intent.text || "A2"}, ${formatValue(
      intent.regex
    )})`;
  },

  regexreplace: (ctx, formatValue) => {
    const { intent } = ctx;
    return `=REGEXREPLACE(${intent.text || "A2"}, ${formatValue(
      intent.regex
    )}, ${formatValue(intent.replacement)})`;
  },

  //-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
  // 하이퍼링크 함수
  //-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
  hyperlink: (ctx, formatValue) => {
    const { intent } = ctx;
    return `=HYPERLINK(${formatValue(intent.url)}, ${formatValue(
      intent.link_label || intent.url
    )})`;
  },
};

module.exports = googleSheetFunctionBuilder;
