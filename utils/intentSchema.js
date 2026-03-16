const DEFAULT_POLICY = {
  mode: "loose",
  value_if_not_found: "",
  value_if_error: "",
};

const DEFAULT_FORMAT_OPTIONS = {
  case_sensitive: false,
  trim_text: true,
  coerce_number: true,
};

function asArray(v) {
  if (v == null) return [];
  return Array.isArray(v) ? v.filter(Boolean) : [v].filter(Boolean);
}

function normalizeOperation(op = "") {
  const s = String(op || "")
    .trim()
    .toLowerCase();

  if (["lookup", "xlookup", "find", "search", "reference"].includes(s))
    return "xlookup";
  if (["avg", "mean"].includes(s)) return "average";
  if (["total"].includes(s)) return "sum";
  if (["cnt", "countifs"].includes(s)) return "count";
  if (["minimum"].includes(s)) return "min";
  if (["maximum"].includes(s)) return "max";
  if (["med", "middle"].includes(s)) return "median";
  if (["sort", "sortby"].includes(s)) return "sortby";
  if (["filter"].includes(s)) return "filter";

  return s || "formula";
}

function normalizeEngine(intent = {}) {
  const raw =
    intent.engine ||
    intent.platform ||
    intent.target_engine ||
    intent.target ||
    "";
  const s = String(raw || "")
    .trim()
    .toLowerCase();
  if (
    ["sheets", "googlesheets", "google sheets", "google_sheets"].includes(s)
  ) {
    return "sheets";
  }
  return "excel";
}

function normalizePolicy(intent = {}) {
  const engine = normalizeEngine(intent);

  const policy = {
    mode: intent.mode === "strict" ? "strict" : DEFAULT_POLICY.mode,
    value_if_not_found:
      intent.value_if_not_found != null
        ? String(intent.value_if_not_found)
        : DEFAULT_POLICY.value_if_not_found,
    value_if_error:
      intent.value_if_error != null
        ? String(intent.value_if_error)
        : DEFAULT_POLICY.value_if_error,
  };

  const formatOptions = { ...DEFAULT_FORMAT_OPTIONS };
  if (policy.mode === "strict") {
    formatOptions.case_sensitive = true;
    formatOptions.trim_text = false;
    formatOptions.coerce_number = false;
  }

  return { engine, policy, formatOptions };
}

function normalizeReturnFields(intent = {}, rawMessage = "") {
  const out = [];

  if (Array.isArray(intent.return_fields)) {
    out.push(...intent.return_fields.filter(Boolean).map(String));
  }

  if (intent.return_array?.header) out.push(String(intent.return_array.header));
  if (intent.return?.header) out.push(String(intent.return.header));
  if (intent.return_hint) out.push(String(intent.return_hint));
  if (intent.header_hint && !out.length) out.push(String(intent.header_hint));

  // 다중 반환 자연어 흔적 보정
  const msg = String(rawMessage || "");
  const known = [
    "이름",
    "부서",
    "직급",
    "연봉",
    "입사일",
    "평가 등급",
    "직원 ID",
  ];
  for (const k of known) {
    if (msg.includes(k) && /가져와|보여줘|출력/.test(msg) && !out.includes(k)) {
      out.push(k);
    }
  }

  return [...new Set(out)];
}

function normalizeLookup(intent = {}) {
  const lookup = {
    key_header: null,
    value: null,
    value_ref: null,
    match_mode: "exact",
  };

  if (intent.lookup_array?.header)
    lookup.key_header = String(intent.lookup_array.header);
  if (intent.lookup?.header && !lookup.key_header) {
    lookup.key_header = String(intent.lookup.header);
  }
  if (intent.lookup_hint && !lookup.key_header) {
    lookup.key_header = String(intent.lookup_hint);
  }

  if (intent.lookup_value != null) {
    const s = String(intent.lookup_value).trim();
    if (
      /^'?[^'!]+!'?[A-Z]{1,3}\d{1,7}$/.test(s) ||
      /^[A-Z]{1,3}\d{1,7}$/i.test(s)
    ) {
      lookup.value_ref = s;
    } else {
      lookup.value = intent.lookup_value;
    }
  }

  if (
    intent.lookup?.value != null &&
    lookup.value == null &&
    lookup.value_ref == null
  ) {
    const s = String(intent.lookup.value).trim();
    if (/^[A-Z]{1,3}\d{1,7}$/i.test(s)) lookup.value_ref = s;
    else lookup.value = intent.lookup.value;
  }

  if (intent.match_mode) lookup.match_mode = String(intent.match_mode);
  if (intent.lookup_mode) lookup.match_mode = String(intent.lookup_mode);

  return lookup;
}

function normalizeFilters(intent = {}) {
  const src = Array.isArray(intent.filters)
    ? intent.filters
    : Array.isArray(intent.conditions)
      ? intent.conditions
      : [];

  const out = [];

  for (const c of src) {
    if (!c) continue;

    if (c.logical_operator && Array.isArray(c.conditions)) {
      out.push({
        logical_operator: String(c.logical_operator).toUpperCase(),
        conditions: c.conditions.map((x) => ({
          header:
            typeof x?.target === "string"
              ? x.target
              : x?.target?.header || x?.header || x?.hint || null,
          operator: x?.operator || "=",
          value: x?.value,
          value_type: x?.value_type || null,
        })),
      });
      continue;
    }

    out.push({
      header:
        typeof c.target === "string"
          ? c.target
          : c?.target?.header || c?.header || c?.hint || null,
      operator: c.operator || "=",
      value: c.value,
      value_type: c.value_type || null,
    });
  }

  return out.filter(Boolean);
}

function normalizeSort(intent = {}) {
  if (!intent.sort && !intent.sort_by && !intent.sort_order && !intent.sorted)
    return null;

  return {
    header:
      intent.sort?.header ||
      intent.sort_by ||
      intent.header_hint ||
      intent.return_hint ||
      null,
    order: String(
      intent.sort?.order || intent.sort_order || "desc",
    ).toLowerCase(),
  };
}

function normalizeLimit(intent = {}, rawMessage = "") {
  if (intent.limit != null) return Number(intent.limit) || null;
  if (intent.top_n != null) return Number(intent.top_n) || null;

  const m = String(rawMessage || "").match(/상위\s*(\d+)|top\s*(\d+)/i);
  if (m) return Number(m[1] || m[2]) || null;
  return null;
}

function normalizeDuplicateRule(intent = {}, rawMessage = "") {
  if (intent.duplicate_rule) return String(intent.duplicate_rule);
  const msg = String(rawMessage || "");
  if (/가장\s*최근|최신|마지막/.test(msg)) return "latest";
  return null;
}

function normalizeIntentSchema(rawIntent = {}, rawMessage = "") {
  const intent = { ...(rawIntent || {}) };
  intent.operation = normalizeOperation(intent.operation);

  const schema = {
    ...intent,
    operation: intent.operation,
    engine: normalizeEngine(intent),
    return_fields: normalizeReturnFields(intent, rawMessage),
    lookup: normalizeLookup(intent),
    filters: normalizeFilters(intent),
    sort: normalizeSort(intent),
    limit: normalizeLimit(intent, rawMessage),
    duplicate_rule: normalizeDuplicateRule(intent, rawMessage),
    group_by: intent.group_by || null,
    raw_message: rawMessage || intent.raw_message || "",
  };

  // lookup 계열 최소 보정
  if (schema.operation === "xlookup") {
    if (!schema.lookup.key_header && intent.lookup_hint) {
      schema.lookup.key_header = String(intent.lookup_hint);
    }
  }

  return schema;
}

module.exports = {
  normalizeIntentSchema,
  normalizePolicy,
  normalizeOperation,
};
