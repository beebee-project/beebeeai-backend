const fs = require("fs");
const path = require("path");

const root = process.cwd();
const target = path.join(root, "controllers", "automationController.js");

function read(file) {
  return fs.readFileSync(file, "utf8");
}

function write(file, content) {
  fs.writeFileSync(file, content, "utf8");
  console.log(`[download-cleanup] wrote ${path.relative(root, file)}`);
}

function addDeleteObjectImport(content) {
  if (
    content.includes("deleteObject") &&
    content.includes("../utils/storage")
  ) {
    return content;
  }

  const storageRequireRe =
    /const\s*\{([\s\S]*?)\}\s*=\s*require\(["']\.\.\/utils\/storage["']\);/m;
  const match = content.match(storageRequireRe);
  if (!match) {
    throw new Error("Could not find destructured ../utils/storage import");
  }

  const names = match[1]
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);

  if (!names.includes("deleteObject")) names.push("deleteObject");

  const replacement = [
    "const {",
    ...names.map((name) => `  ${name},`),
    '} = require("../utils/storage");',
  ].join("\n");

  return content.replace(storageRequireRe, replacement);
}

function patchBuildGeneratedDownloadUrl(content) {
  if (
    /function\s+buildGeneratedDownloadUrl[\s\S]*?queryTablesKey\s*=/.test(
      content,
    )
  ) {
    return content;
  }

  const signatureRe =
    /function\s+buildGeneratedDownloadUrl\(\{\s*storageKey\s*=\s*"",\s*filePath\s*=\s*"",\s*displayName\s*=\s*"",\s*outputType\s*=\s*"",\s*\}\s*=\s*\{\}\)\s*\{/m;

  if (!signatureRe.test(content)) {
    throw new Error("Could not find buildGeneratedDownloadUrl signature");
  }

  content = content.replace(
    signatureRe,
    [
      "function buildGeneratedDownloadUrl({",
      '  storageKey = "",',
      '  filePath = "",',
      '  displayName = "",',
      '  outputType = "",',
      '  queryTablesKey = "",',
      "  deleteAfterDownload = true,",
      "} = {}) {",
    ].join("\n"),
  );

  content = content.replace(
    /if\s*\(outputType\)\s*params\.set\("outputType",\s*outputType\);/,
    [
      'if (outputType) params.set("outputType", outputType);',
      '  if (queryTablesKey) params.set("queryTablesKey", queryTablesKey);',
      '  if (deleteAfterDownload) params.set("deleteAfterDownload", "1");',
    ].join("\n"),
  );

  return content;
}

function insertCleanupHelpers(content) {
  if (content.includes("function assertQueryTablesKeyAccess(")) return content;

  const helperLines = [
    "",
    'function assertQueryTablesKeyAccess(req, queryTablesKey = "") {',
    '  const key = String(queryTablesKey || "");',
    "  if (!key) return false;",
    "",
    '  const userId = req.user?.id ? String(req.user.id) : "local-dev";',
    "",
    "  return (",
    '    key.startsWith("query-tables/" + userId + "/") ||',
    '    key.startsWith("query-tables/local-dev/")',
    "  );",
    "}",
    "",
    'function shouldDeleteAfterDownload(value = "1") {',
    '  const normalized = String(value ?? "1").trim().toLowerCase();',
    '  return !["0", "false", "no", "off"].includes(normalized);',
    "}",
    "",
    "async function cleanupGeneratedDownloadArtifacts({",
    "  req,",
    '  storageKey = "",',
    '  filePath = "",',
    '  queryTablesKey = "",',
    "} = {}) {",
    "  const deleted = [];",
    "  const errors = [];",
    "",
    "  async function tryDelete(label, fn) {",
    "    try {",
    "      await fn();",
    "      deleted.push(label);",
    "    } catch (error) {",
    "      errors.push({ label, message: error?.message || String(error) });",
    "    }",
    "  }",
    "",
    "  if (storageKey && assertGeneratedStorageKeyAccess(req, storageKey)) {",
    '    await tryDelete("storage:" + storageKey, () => deleteObject(storageKey));',
    "  }",
    "",
    "  if (filePath) {",
    "    const safePath = resolveSafeGeneratedLocalPath(filePath);",
    "    if (safePath && fs.existsSync(safePath)) {",
    '      await tryDelete("file:" + safePath, async () => {',
    "        fs.unlinkSync(safePath);",
    "      });",
    "    }",
    "  }",
    "",
    "  if (queryTablesKey && assertQueryTablesKeyAccess(req, queryTablesKey)) {",
    '    await tryDelete("queryTables:" + queryTablesKey, () => deleteObject(queryTablesKey));',
    "  }",
    "",
    "  if (deleted.length || errors.length) {",
    '    console.log("[automation.download.cleanup]", { deleted, errors });',
    "  }",
    "}",
    "",
    "function scheduleGeneratedDownloadCleanup(res, cleanupFn) {",
    "  let done = false;",
    "",
    "  async function run() {",
    "    if (done) return;",
    "    done = true;",
    "",
    "    try {",
    "      await cleanupFn();",
    "    } catch (error) {",
    '      console.error("[automation.download.cleanup] failed", error);',
    "    }",
    "  }",
    "",
    '  res.once("finish", run);',
    '  res.once("close", run);',
    "}",
    "",
  ].join("\n");

  if (!content.includes("function resolveSafeGeneratedLocalPath(")) {
    throw new Error(
      "Could not find resolveSafeGeneratedLocalPath insertion point",
    );
  }

  return content.replace(
    /function\s+resolveSafeGeneratedLocalPath\(filePath\s*=\s*""\)\s*\{/,
    helperLines + 'function resolveSafeGeneratedLocalPath(filePath = "") {',
  );
}

function patchDownloadGeneratedFileParams(content) {
  if (
    content.includes("deleteAfterDownload") &&
    /const\s*\{[\s\S]*?queryTablesKey[\s\S]*?\}\s*=\s*req\.query/.test(content)
  ) {
    return content;
  }

  const destructuringRe =
    /const\s*\{\s*storageKey\s*=\s*"",\s*filePath\s*=\s*"",\s*displayName\s*=\s*"",\s*outputType\s*=\s*"",\s*\}\s*=\s*req\.query\s*\|\|\s*\{\};/m;

  if (!destructuringRe.test(content)) {
    throw new Error(
      "Could not find downloadGeneratedFile req.query destructuring",
    );
  }

  return content.replace(
    destructuringRe,
    [
      "const {",
      '      storageKey = "",',
      '      filePath = "",',
      '      displayName = "",',
      '      outputType = "",',
      '      queryTablesKey = "",',
      '      deleteAfterDownload = "1",',
      "    } = req.query || {};",
    ].join("\n"),
  );
}

function patchDownloadCleanupSchedule(content) {
  if (content.includes("cleanupGeneratedDownloadArtifacts({")) return content;

  const responseRe =
    /res\.setHeader\(\s*"Content-Type",\s*contentTypeForGeneratedFile\(safeDisplayName,\s*outputType\),\s*\);\s*\n\s*return\s+res\.end\(buffer\);/m;

  if (!responseRe.test(content)) {
    throw new Error("Could not find generated file response block");
  }

  return content.replace(
    responseRe,
    [
      "res.setHeader(",
      '      "Content-Type",',
      "      contentTypeForGeneratedFile(safeDisplayName, outputType),",
      "    );",
      "",
      "    if (shouldDeleteAfterDownload(deleteAfterDownload)) {",
      "      scheduleGeneratedDownloadCleanup(res, () =>",
      "        cleanupGeneratedDownloadArtifacts({",
      "          req,",
      "          storageKey,",
      "          filePath,",
      "          queryTablesKey,",
      "        }),",
      "      );",
      "    }",
      "",
      "    return res.end(buffer);",
    ].join("\n"),
  );
}

function addQueryTablesKeyToDownloadUrls(content) {
  return content.replace(
    /buildGeneratedDownloadUrl\(\{([\s\S]*?)\n(\s*)\}\)/g,
    (match, body, indent) => {
      if (body.includes("queryTablesKey")) return match;
      if (!body.includes("outputType")) return match;
      return `buildGeneratedDownloadUrl({${body}\n${indent}  queryTablesKey,\n${indent}})`;
    },
  );
}

function patchAutomationController() {
  if (!fs.existsSync(target)) {
    throw new Error(`Missing file: ${target}`);
  }

  let content = read(target);
  content = addDeleteObjectImport(content);
  content = patchBuildGeneratedDownloadUrl(content);
  content = insertCleanupHelpers(content);
  content = patchDownloadGeneratedFileParams(content);
  content = patchDownloadCleanupSchedule(content);
  content = addQueryTablesKeyToDownloadUrls(content);

  write(target, content);
}

patchAutomationController();
console.log("[download-cleanup] done");
console.log(
  "[download-cleanup] next: node --check controllers/automationController.js",
);
