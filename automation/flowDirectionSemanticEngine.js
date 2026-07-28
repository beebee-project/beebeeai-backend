"use strict";

const FLOW_DIRECTION_SEMANTIC_ENGINE_VERSION =
  "flow_direction_semantic_engine_v1";
const FLOW_DIRECTION_SECTION_REPAIR_VERSION =
  "flow_direction_section_repair_v1_system_net_dual_entry";

const DIRECTION_HEADER_PATTERN =
  /^(?:(?:입출고|이동|수불|재고이동|창고이동|거래|물류)(?:구분|유형|방향)|(?:flow|movement|transaction)(?:type|direction)|direction)$/i;
const QUANTITY_HEADER_PATTERN = /(?:수량|quantity|qty)$/i;
const SOURCE_LOCATION_HEADER_PATTERN =
  /^(?:출발|출고|발송|source|from)(?:창고|위치|지점|장소|location|warehouse)?$/i;
const DESTINATION_LOCATION_HEADER_PATTERN =
  /^(?:도착|입고|수신|destination|dest|to)(?:창고|위치|지점|장소|location|warehouse)?$/i;
const PERIOD_HEADER_PATTERN =
  /^(?:기간|기준기간|연월|년월|기준월|월|일자|날짜|이동일|거래일|처리일|date|period|month)$/i;
const ENTITY_HEADER_PATTERN =
  /(?:품목|소모품|제품|상품|자산|장비|항목)(?:명|이름)$|^(?:item|product|asset|equipment|entity)$/i;
const MIXED_FLOW_TITLE_PATTERN =
  /(?:재고|입출고|수불).*(?:흐름|요약)|(?:inventory|stock).*(?:flow|summary)/i;

function normalizeText(value = "") {
  return String(value == null ? "" : value)
    .normalize("NFKC")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeKey(value = "") {
  return normalizeText(value)
    .toLowerCase()
    .replace(/[\s_\-./\\|:;,'"‘’“”()[\]{}<>]+/g, "");
}

function cloneValue(value) {
  return JSON.parse(JSON.stringify(value));
}

function numericValue(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const source = normalizeText(value);
  if (!source || source === "-") return null;
  const normalized = source.replace(/,/g, "").replace(/%$/g, "").trim();
  if (!/^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(normalized)) return null;
  const result = Number(normalized);
  return Number.isFinite(result) ? result : null;
}

function tableRows(table = {}) {
  return Array.isArray(table.rows) ? table.rows : [];
}

function tableColumns(table = {}) {
  return Array.isArray(table.columns) ? table.columns : [];
}

function columnHeader(column = {}, index = 0) {
  return normalizeText(
    column.header ||
      column.originalHeader ||
      column.name ||
      column.label ||
      `열${index + 1}`,
  );
}

function rowValue(row, column = {}, index = 0) {
  if (Array.isArray(row)) return row[index];
  if (!row || typeof row !== "object") return undefined;

  const keys = [
    column.key,
    column.canonicalKey,
    column.accessor,
    column.name,
    column.header,
    column.originalHeader,
    column.label,
    column.id,
  ]
    .map((value) => String(value || "").trim())
    .filter(Boolean);

  for (const key of keys) {
    if (Object.prototype.hasOwnProperty.call(row, key)) return row[key];
  }

  const normalizedTargets = new Set(keys.map(normalizeKey));
  for (const [key, value] of Object.entries(row)) {
    if (normalizedTargets.has(normalizeKey(key))) return value;
  }

  return Object.values(row)[index];
}

function tableLabel(table = {}, index = 0) {
  return normalizeText(
    table.tableName ||
      table.sheetName ||
      table.title ||
      table.tableId ||
      `표 ${index + 1}`,
  );
}

function canonicalFlowDirection(value = "") {
  const key = normalizeKey(value);
  if (!key) return "";
  if (/^(?:입고|입하|반입|수취|inbound|receipt|received)$/.test(key)) {
    return "inbound";
  }
  if (/^(?:출고|출하|반출|outbound|shipment|shipped)$/.test(key)) {
    return "outbound";
  }
  if (/^(?:이동|내부이동|재고이동|창고이동|이관|transfer|movement)$/.test(key)) {
    return "transfer";
  }
  return "";
}

function normalizedPeriod(value = "") {
  const text = normalizeText(value);
  if (!text) return "";
  const match = text.match(/^(\d{4})[-./년\s]*(\d{1,2})/);
  if (match) return `${match[1]}-${String(Number(match[2])).padStart(2, "0")}`;
  return text;
}

function distinctValuesForColumn(table = {}, column = {}, index = 0) {
  return new Set(
    tableRows(table)
      .map((row) => normalizeText(rowValue(row, column, index)))
      .filter(Boolean)
      .map(normalizeKey),
  );
}

function resolveDirectionalTable(table = {}, tableIndex = 0) {
  const columns = tableColumns(table);
  const rows = tableRows(table);
  if (!columns.length || !rows.length) {
    return {
      applied: false,
      tableIndex,
      tableLabel: tableLabel(table, tableIndex),
      reason: "empty_table",
    };
  }

  const facts = columns.map((column, index) => {
    const header = columnHeader(column, index);
    return { column, index, header, key: normalizeKey(header) };
  });

  const directionCandidates = facts.filter((fact) =>
    DIRECTION_HEADER_PATTERN.test(fact.key),
  );
  const quantityCandidates = facts.filter((fact) =>
    QUANTITY_HEADER_PATTERN.test(fact.key),
  );

  let best = null;
  for (const direction of directionCandidates) {
    for (const quantity of quantityCandidates) {
      const records = [];
      let unknownDirectionRowCount = 0;
      for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
        const row = rows[rowIndex];
        const canonical = canonicalFlowDirection(
          rowValue(row, direction.column, direction.index),
        );
        const numeric = numericValue(
          rowValue(row, quantity.column, quantity.index),
        );
        if (numeric == null) continue;
        if (!canonical) {
          unknownDirectionRowCount += 1;
          continue;
        }
        records.push({ row, rowIndex, direction: canonical, quantity: numeric });
      }

      const classes = new Set(records.map((record) => record.direction));
      const hasExternal = classes.has("inbound") || classes.has("outbound");
      if (!records.length || !hasExternal) continue;

      const score =
        records.length * 10 +
        classes.size * 25 +
        (classes.has("transfer") ? 60 : 0) -
        unknownDirectionRowCount * 2;
      if (!best || score > best.score) {
        best = {
          score,
          direction,
          quantity,
          records,
          classes,
          unknownDirectionRowCount,
        };
      }
    }
  }

  if (!best) {
    return {
      applied: false,
      tableIndex,
      tableLabel: tableLabel(table, tableIndex),
      reason: "no_direction_quantity_evidence",
    };
  }

  const source = facts.find((fact) =>
    SOURCE_LOCATION_HEADER_PATTERN.test(fact.key),
  );
  const destination = facts.find((fact) =>
    DESTINATION_LOCATION_HEADER_PATTERN.test(fact.key),
  );
  const period = facts.find((fact) => PERIOD_HEADER_PATTERN.test(fact.key));

  const entityCandidates = facts
    .filter((fact) => ENTITY_HEADER_PATTERN.test(fact.header))
    .map((fact) => ({
      ...fact,
      distinctCount: distinctValuesForColumn(table, fact.column, fact.index).size,
    }))
    .sort((left, right) => right.distinctCount - left.distinctCount);
  const entity = entityCandidates[0] || null;

  const records = best.records.map((record) => {
    const dimensions = {};
    for (const fact of facts) {
      dimensions[fact.header] = rowValue(record.row, fact.column, fact.index);
    }
    return {
      rowIndex: record.rowIndex,
      direction: record.direction,
      quantity: record.quantity,
      sourceLocation: source
        ? normalizeText(rowValue(record.row, source.column, source.index))
        : "",
      destinationLocation: destination
        ? normalizeText(
            rowValue(record.row, destination.column, destination.index),
          )
        : "",
      period: period
        ? normalizedPeriod(rowValue(record.row, period.column, period.index))
        : "",
      entity: entity
        ? normalizeText(rowValue(record.row, entity.column, entity.index))
        : "",
      dimensions,
    };
  });

  return {
    applied: true,
    version: FLOW_DIRECTION_SEMANTIC_ENGINE_VERSION,
    tableIndex,
    tableLabel: tableLabel(table, tableIndex),
    directionHeader: best.direction.header,
    quantityHeader: best.quantity.header,
    sourceLocationHeader: source?.header || "",
    destinationLocationHeader: destination?.header || "",
    periodHeader: period?.header || "",
    entityHeader: entity?.header || "",
    directionClasses: Array.from(best.classes).sort(),
    recognizedDirectionRowCount: records.length,
    unknownDirectionRowCount: best.unknownDirectionRowCount,
    dualEntryAvailable: Boolean(source && destination),
    records,
    reason: "direction_quantity_resolved",
  };
}

function resolveFlowDirectionEvidence(tables = []) {
  const candidates = (Array.isArray(tables) ? tables : [])
    .map((table, tableIndex) => resolveDirectionalTable(table, tableIndex));
  const applied = candidates
    .filter((candidate) => candidate.applied)
    .sort((left, right) =>
      Number(right.dualEntryAvailable) - Number(left.dualEntryAvailable) ||
      right.recognizedDirectionRowCount - left.recognizedDirectionRowCount ||
      right.directionClasses.length - left.directionClasses.length,
    )[0];

  if (!applied) {
    return {
      applied: false,
      version: FLOW_DIRECTION_SEMANTIC_ENGINE_VERSION,
      reason: "no_directional_flow_table",
      tableEvidence: candidates.map((candidate) => ({
        tableIndex: candidate.tableIndex,
        tableLabel: candidate.tableLabel,
        applied: candidate.applied,
        reason: candidate.reason,
      })),
      records: [],
    };
  }

  return {
    ...applied,
    tableEvidence: candidates.map((candidate) => ({
      tableIndex: candidate.tableIndex,
      tableLabel: candidate.tableLabel,
      applied: candidate.applied,
      reason: candidate.reason,
      directionHeader: candidate.directionHeader || "",
      quantityHeader: candidate.quantityHeader || "",
      directionClasses: candidate.directionClasses || [],
      recognizedDirectionRowCount:
        candidate.recognizedDirectionRowCount || 0,
    })),
  };
}

function sumDirection(records = [], direction = "") {
  return records.reduce(
    (sum, record) =>
      record.direction === direction ? sum + record.quantity : sum,
    0,
  );
}

function countDirection(records = [], direction = "") {
  return records.filter((record) => record.direction === direction).length;
}

function buildSystemFlowSummary(records = []) {
  const inboundQuantity = sumDirection(records, "inbound");
  const outboundQuantity = sumDirection(records, "outbound");
  const internalTransferQuantity = sumDirection(records, "transfer");
  return {
    rowCount: records.length,
    inboundQuantity,
    outboundQuantity,
    internalTransferQuantity,
    totalHandledQuantity:
      inboundQuantity + outboundQuantity + internalTransferQuantity,
    netInventoryChange: inboundQuantity - outboundQuantity,
    inboundRowCount: countDirection(records, "inbound"),
    outboundRowCount: countDirection(records, "outbound"),
    transferRowCount: countDirection(records, "transfer"),
  };
}

function buildDirectionRows(records = []) {
  const definitions = [
    ["inbound", "입고"],
    ["outbound", "출고"],
    ["transfer", "내부이동"],
  ];
  return definitions
    .map(([direction, label]) => {
      const quantity = sumDirection(records, direction);
      const count = countDirection(records, direction);
      const inbound = direction === "inbound" ? quantity : 0;
      const outbound = direction === "outbound" ? quantity : 0;
      const transfer = direction === "transfer" ? quantity : 0;
      return {
        구분: label,
        건수: count,
        입고수량: inbound,
        출고수량: outbound,
        내부이동수량: transfer,
        순증감수량: inbound - outbound,
        총처리수량: inbound + outbound + transfer,
      };
    })
    .filter((row) => row.건수 > 0);
}

function groupedSummary(records = [], keySelector = () => "") {
  const grouped = new Map();
  for (const record of records) {
    const key = normalizeText(keySelector(record)) || "미분류";
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key).push(record);
  }
  return grouped;
}

function buildGroupedFlowRows(records = [], groupHeader = "구분", keySelector) {
  return [...groupedSummary(records, keySelector).entries()]
    .map(([group, groupedRecords]) => {
      const summary = buildSystemFlowSummary(groupedRecords);
      return {
        [groupHeader]: group,
        건수: summary.rowCount,
        입고수량: summary.inboundQuantity,
        출고수량: summary.outboundQuantity,
        내부이동수량: summary.internalTransferQuantity,
        순증감수량: summary.netInventoryChange,
        총처리수량: summary.totalHandledQuantity,
      };
    })
    .sort((left, right) =>
      String(left[groupHeader]).localeCompare(String(right[groupHeader]), "ko", {
        numeric: true,
      }),
    );
}

function buildPeriodFlowRows(records = []) {
  return buildGroupedFlowRows(records, "기간", (record) => record.period || "미분류");
}

function buildEntityFlowRows(records = [], entityHeader = "항목") {
  return buildGroupedFlowRows(
    records,
    entityHeader,
    (record) => record.entity || record.dimensions?.[entityHeader] || "미분류",
  );
}

function buildLocationLedgerRows(records = []) {
  const ledgers = new Map();
  const ensure = (location) => {
    const normalized = normalizeText(location);
    if (!normalized) return null;
    if (!ledgers.has(normalized)) {
      ledgers.set(normalized, {
        "창고·위치": normalized,
        건수: 0,
        외부입고수량: 0,
        외부출고수량: 0,
        내부이동입고수량: 0,
        내부이동출고수량: 0,
        순증감수량: 0,
        총처리수량: 0,
      });
    }
    return ledgers.get(normalized);
  };

  for (const record of records) {
    if (record.direction === "inbound") {
      const target = ensure(record.destinationLocation);
      if (!target) continue;
      target.건수 += 1;
      target.외부입고수량 += record.quantity;
      target.순증감수량 += record.quantity;
      target.총처리수량 += record.quantity;
      continue;
    }
    if (record.direction === "outbound") {
      const source = ensure(record.sourceLocation);
      if (!source) continue;
      source.건수 += 1;
      source.외부출고수량 += record.quantity;
      source.순증감수량 -= record.quantity;
      source.총처리수량 += record.quantity;
      continue;
    }
    if (record.direction === "transfer") {
      const source = ensure(record.sourceLocation);
      const target = ensure(record.destinationLocation);
      if (source) {
        source.건수 += 1;
        source.내부이동출고수량 += record.quantity;
        source.순증감수량 -= record.quantity;
        source.총처리수량 += record.quantity;
      }
      if (target) {
        target.건수 += 1;
        target.내부이동입고수량 += record.quantity;
        target.순증감수량 += record.quantity;
        target.총처리수량 += record.quantity;
      }
    }
  }

  return [...ledgers.values()].sort((left, right) =>
    left["창고·위치"].localeCompare(right["창고·위치"], "ko", {
      numeric: true,
    }),
  );
}

function sectionRows(section = {}) {
  return Array.isArray(section.result?.rows) ? section.result.rows : [];
}

function rowKeySet(row = {}) {
  return new Set(Object.keys(row || {}).map(normalizeKey));
}

function firstRow(section = {}) {
  return sectionRows(section).find(
    (row) => row && typeof row === "object" && !Array.isArray(row),
  ) || null;
}

function sectionText(section = {}) {
  return normalizeText(
    [
      section.sectionId,
      section.title,
      section.sectionType,
      section.result?.operation,
      section.result?.groupBy?.header,
      section.groupHeader,
    ]
      .filter(Boolean)
      .join(" "),
  );
}

function mixedFlowRowShape(section = {}) {
  const row = firstRow(section);
  if (!row) return false;
  const keys = rowKeySet(row);
  return (
    keys.has(normalizeKey("입고수량")) &&
    keys.has(normalizeKey("출고수량")) &&
    (keys.has(normalizeKey("조정수량")) ||
      keys.has(normalizeKey("내부이동수량"))) &&
    keys.has(normalizeKey("순증감수량"))
  );
}

function isSystemFlowOverview(section = {}) {
  const rows = sectionRows(section);
  const labels = new Set(
    rows.map((row) => normalizeKey(row?.지표 || row?.metric || "")),
  );
  return (
    MIXED_FLOW_TITLE_PATTERN.test(sectionText(section)) &&
    labels.has(normalizeKey("입고 수량")) &&
    labels.has(normalizeKey("출고 수량")) &&
    (labels.has(normalizeKey("순증감 수량")) ||
      labels.has(normalizeKey("추정 재고 증감")))
  );
}

function groupHeaderFromSection(section = {}) {
  const explicit = normalizeText(
    section.result?.groupBy?.header ||
      section.groupBy?.header ||
      section.groupHeader ||
      "",
  );
  if (explicit) return explicit;
  const row = firstRow(section);
  if (!row) return "";
  const excluded = new Set(
    [
      "건수",
      "입고수량",
      "출고수량",
      "조정수량",
      "내부이동수량",
      "순증감수량",
      "재고수량",
      "금액",
      "출고율",
      "총처리수량",
    ].map(normalizeKey),
  );
  return Object.keys(row).find((key) => !excluded.has(normalizeKey(key))) || "";
}

function isDirectionGroupHeader(header = "") {
  return /구분|유형|방향|direction|type/i.test(normalizeText(header));
}

function isPeriodGroupHeader(header = "") {
  return PERIOD_HEADER_PATTERN.test(normalizeText(header));
}

function isLocationGroupHeader(header = "") {
  return /창고|위치|지점|장소|warehouse|location/i.test(normalizeText(header));
}

function setSectionRows(section = {}, rows = [], patch = {}) {
  const next = cloneValue(section);
  next.title = patch.title || next.title;
  next.sectionType = patch.sectionType || next.sectionType;
  next.result = {
    ...(next.result || {}),
    rows: cloneValue(rows),
    operation: patch.operation || next.result?.operation || "flowDirection",
    meta: {
      ...(next.result?.meta || {}),
      flowDirectionSemanticEngineVersion:
        FLOW_DIRECTION_SEMANTIC_ENGINE_VERSION,
      flowDirectionSectionRepairVersion:
        FLOW_DIRECTION_SECTION_REPAIR_VERSION,
      flowDirectionScope: patch.scope || "",
      flowDirectionRepairApplied: true,
      internalTransferNetEffect: 0,
    },
  };
  if (patch.groupHeader) {
    next.groupHeader = patch.groupHeader;
    next.result.groupBy = {
      ...(next.result?.groupBy || {}),
      header: patch.groupHeader,
    };
  }
  return next;
}

function buildSystemOverviewRows(summary = {}) {
  const total = summary.totalHandledQuantity || 0;
  const percent = (value) => (total ? (value / total) * 100 : null);
  return [
    { 지표: "전체 행 수", 값: summary.rowCount, 보조값: null, 비율Percent: null },
    {
      지표: "입고 수량",
      값: summary.inboundQuantity,
      보조값: null,
      비율Percent: percent(summary.inboundQuantity),
    },
    {
      지표: "출고 수량",
      값: summary.outboundQuantity,
      보조값: null,
      비율Percent: percent(summary.outboundQuantity),
    },
    {
      지표: "내부 이동 수량",
      값: summary.internalTransferQuantity,
      보조값: null,
      비율Percent: percent(summary.internalTransferQuantity),
    },
    {
      지표: "총 이동 처리량",
      값: summary.totalHandledQuantity,
      보조값: null,
      비율Percent: 100,
    },
    {
      지표: "순증감 수량",
      값: summary.netInventoryChange,
      보조값: null,
      비율Percent: null,
    },
  ];
}

function applyFlowDirectionSemantics({ sections = [], tables = [] } = {}) {
  const inputSections = Array.isArray(sections) ? sections : [];
  const evidence = resolveFlowDirectionEvidence(tables);
  if (!evidence.applied) {
    return {
      sections: cloneValue(inputSections),
      applied: false,
      version: FLOW_DIRECTION_SEMANTIC_ENGINE_VERSION,
      repairVersion: FLOW_DIRECTION_SECTION_REPAIR_VERSION,
      reason: evidence.reason,
      evidence,
      repairedSectionCount: 0,
      repairedSectionIds: [],
      scopes: [],
      systemSummary: {},
      dualEntryApplied: false,
      locationEntryCount: 0,
    };
  }

  const systemSummary = buildSystemFlowSummary(evidence.records);
  const directionRows = buildDirectionRows(evidence.records);
  const periodRows = buildPeriodFlowRows(evidence.records);
  const entityHeader = evidence.entityHeader || "항목";
  const entityRows = buildEntityFlowRows(evidence.records, entityHeader);
  const locationRows = evidence.dualEntryAvailable
    ? buildLocationLedgerRows(evidence.records)
    : [];

  const repairedSectionIds = [];
  const scopes = [];
  let locationSectionRepaired = false;

  const repaired = inputSections.map((section) => {
    let next = null;
    let scope = "";

    if (isSystemFlowOverview(section)) {
      scope = "system_overview";
      next = setSectionRows(section, buildSystemOverviewRows(systemSummary), {
        scope,
        operation: "flowDirectionSystemSummary",
      });
    } else if (mixedFlowRowShape(section)) {
      const groupHeader = groupHeaderFromSection(section);
      if (isDirectionGroupHeader(groupHeader)) {
        scope = "direction_breakdown";
        next = setSectionRows(section, directionRows, {
          scope,
          groupHeader,
          operation: "flowDirectionBreakdown",
        });
      } else if (isPeriodGroupHeader(groupHeader)) {
        scope = "period_breakdown";
        next = setSectionRows(section, periodRows, {
          scope,
          groupHeader,
          operation: "flowDirectionPeriod",
        });
      } else if (
        isLocationGroupHeader(groupHeader) &&
        locationRows.length &&
        !locationSectionRepaired
      ) {
        locationSectionRepaired = true;
        scope = "location_dual_entry";
        next = setSectionRows(section, locationRows, {
          scope,
          title: "창고·위치별 재고 증감",
          sectionType: "flow_direction_location_ledger",
          groupHeader: "창고·위치",
          operation: "flowDirectionDualEntry",
        });
      } else if (
        groupHeader &&
        normalizeKey(groupHeader) === normalizeKey(entityHeader)
      ) {
        scope = "entity_system_net";
        next = setSectionRows(section, entityRows, {
          scope,
          groupHeader,
          operation: "flowDirectionEntity",
        });
      }
    }

    if (!next) return cloneValue(section);
    repairedSectionIds.push(
      normalizeText(section.sectionId || section.title || `section_${repairedSectionIds.length + 1}`),
    );
    scopes.push(scope);
    return next;
  });

  return {
    sections: repaired,
    applied: repairedSectionIds.length > 0,
    version: FLOW_DIRECTION_SEMANTIC_ENGINE_VERSION,
    repairVersion: FLOW_DIRECTION_SECTION_REPAIR_VERSION,
    reason: repairedSectionIds.length
      ? "directional_flow_sections_repaired"
      : "no_matching_flow_sections",
    evidence: {
      ...cloneValue(evidence),
      records: undefined,
    },
    repairedSectionCount: repairedSectionIds.length,
    repairedSectionIds,
    scopes: Array.from(new Set(scopes)),
    systemSummary,
    dualEntryApplied: locationSectionRepaired,
    locationEntryCount: locationRows.length,
  };
}

module.exports = {
  FLOW_DIRECTION_SEMANTIC_ENGINE_VERSION,
  FLOW_DIRECTION_SECTION_REPAIR_VERSION,
  canonicalFlowDirection,
  resolveDirectionalTable,
  resolveFlowDirectionEvidence,
  buildSystemFlowSummary,
  buildDirectionRows,
  buildPeriodFlowRows,
  buildEntityFlowRows,
  buildLocationLedgerRows,
  applyFlowDirectionSemantics,
};
