"use strict";

const path = require("path");
const { readJsonObject } = require("../utils/storage");
const {
  buildNormalizedQueryTables,
} = require("./normalizedQueryTableBuilder");
const {
  BUSINESS_TEMPLATE_EXECUTOR_VERSION,
  SEMANTIC_OUTPUT_PLANNER_VERSION,
  executeBusinessTemplate,
} = require("./businessTemplateExecutor");

const AUTOMATION_EXECUTION_ROUTE_BRIDGE_VERSION =
  "automation_execution_route_bridge_v1";
const AUTOMATION_EXECUTE_BUSINESS_ROUTE_VERSION =
  "automation_execute_business_route_v2_direct_executor";

function modulePathForObservation(filePath = __filename) {
  const relative = path.relative(process.cwd(), filePath);
  return String(relative || filePath).replace(/\\/g, "/");
}

function runtimeFingerprint() {
  let executorModulePath = "";
  let plannerModulePath = "";

  try {
    executorModulePath = modulePathForObservation(
      require.resolve("./businessTemplateExecutor"),
    );
  } catch (_error) {
    executorModulePath = "UNRESOLVED";
  }

  try {
    plannerModulePath = modulePathForObservation(
      require.resolve("./semanticOutputPlanner"),
    );
  } catch (_error) {
    plannerModulePath = "UNRESOLVED";
  }

  return {
    automationExecutionRouteBridgeVersion:
      AUTOMATION_EXECUTION_ROUTE_BRIDGE_VERSION,
    automationExecuteBusinessRouteVersion:
      AUTOMATION_EXECUTE_BUSINESS_ROUTE_VERSION,
    businessTemplateExecutorVersion:
      BUSINESS_TEMPLATE_EXECUTOR_VERSION,
    semanticOutputPlannerVersion:
      SEMANTIC_OUTPUT_PLANNER_VERSION,
    routeBridgeModulePath: modulePathForObservation(__filename),
    executorModulePath,
    plannerModulePath,
    routeExecutorConnected: true,
  };
}

function withRouteExecutionMeta(result = {}, extra = {}) {
  return {
    ...result,
    executionMeta: {
      ...(result.executionMeta || {}),
      ...runtimeFingerprint(),
      ...extra,
    },
  };
}

async function resolveExecutionTables({
  queryTablesKey = "",
  normalizedQueryTables,
} = {}) {
  if (Array.isArray(normalizedQueryTables)) {
    return {
      tables: normalizedQueryTables,
      source: "request.normalizedQueryTables",
      saved: null,
    };
  }

  if (!queryTablesKey) {
    return {
      tables: null,
      source: "",
      saved: null,
    };
  }

  const saved = await readJsonObject(queryTablesKey);
  if (Array.isArray(saved?.normalizedQueryTables)) {
    return {
      tables: saved.normalizedQueryTables,
      source: "saved.normalizedQueryTables",
      saved,
    };
  }

  return {
    tables: buildNormalizedQueryTables(saved?.tables || []),
    source: "rebuilt.saved.tables",
    saved,
  };
}

async function executeBusinessTemplateObserved(req, res) {
  const startedAt = Date.now();

  try {
    const {
      queryTablesKey = "",
      normalizedQueryTables,
      templateCandidate,
    } = req.body || {};

    const resolved = await resolveExecutionTables({
      queryTablesKey,
      normalizedQueryTables,
    });

    if (!Array.isArray(resolved.tables)) {
      return res.status(400).json(
        withRouteExecutionMeta(
          {
            ok: false,
            code: "NORMALIZED_QUERY_TABLES_REQUIRED",
            message:
              "normalizedQueryTables 또는 queryTablesKey가 필요합니다.",
          },
          {
            routeInputSource: resolved.source,
            routeExecutionElapsedMs: Date.now() - startedAt,
          },
        ),
      );
    }

    if (!templateCandidate || !templateCandidate.templateId) {
      return res.status(400).json(
        withRouteExecutionMeta(
          {
            ok: false,
            code: "BUSINESS_TEMPLATE_REQUIRED",
            message: "실행할 업무 템플릿 후보가 필요합니다.",
          },
          {
            routeInputSource: resolved.source,
            routeExecutionTableCount: resolved.tables.length,
            routeExecutionElapsedMs: Date.now() - startedAt,
          },
        ),
      );
    }

    const result = executeBusinessTemplate({
      normalizedQueryTables: resolved.tables,
      templateCandidate,
    });

    const observed = withRouteExecutionMeta(result, {
      routeInputSource: resolved.source,
      routeQueryTablesKeyPresent: Boolean(queryTablesKey),
      routeExecutionTableCount: resolved.tables.length,
      routeSavedPhysicalTableCount: Array.isArray(resolved.saved?.tables)
        ? resolved.saved.tables.length
        : 0,
      routeSavedNormalizedTableCount: Array.isArray(
        resolved.saved?.normalizedQueryTables,
      )
        ? resolved.saved.normalizedQueryTables.length
        : 0,
      routeExecutionElapsedMs: Date.now() - startedAt,
    });

    return res.status(observed.ok ? 200 : 400).json(observed);
  } catch (error) {
    console.error("executeBusinessTemplateObserved error:", error);

    return res.status(500).json(
      withRouteExecutionMeta(
        {
          ok: false,
          code: "BUSINESS_TEMPLATE_EXECUTE_FAILED",
          message: "업무 템플릿 실행 중 오류가 발생했습니다.",
        },
        {
          routeExecutionElapsedMs: Date.now() - startedAt,
          routeExecutionError:
            error?.message || String(error),
        },
      ),
    );
  }
}

module.exports = {
  AUTOMATION_EXECUTION_ROUTE_BRIDGE_VERSION,
  AUTOMATION_EXECUTE_BUSINESS_ROUTE_VERSION,
  executeBusinessTemplateObserved,
  resolveExecutionTables,
  runtimeFingerprint,
  withRouteExecutionMeta,
};
