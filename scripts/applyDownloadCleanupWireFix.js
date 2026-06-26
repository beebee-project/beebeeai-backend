/**
 * Wire download cleanup into automationController.downloadGeneratedFile.
 * - Keeps encrypted originals and query-json cache.
 * - Deletes generated artifact and query-tables only after download response finishes/closes.
 */
const fs = require("fs");
const path = require("path");

const ROOT = process.cwd();
const target = path.join(ROOT, "controllers", "automationController.js");

function fail(message) {
  console.error("[download-cleanup-wire-fix] " + message);
  process.exit(1);
}

if (!fs.existsSync(target)) {
  fail("controllers/automationController.js not found");
}

let src = fs.readFileSync(target, "utf8");

if (!src.includes("function scheduleGeneratedDownloadCleanup")) {
  fail(
    "scheduleGeneratedDownloadCleanup helper not found. Apply base cleanup patch first.",
  );
}

const startMarker =
  "exports.downloadGeneratedFile = async (req, res, next) => {";
const start = src.indexOf(startMarker);
if (start < 0) fail("downloadGeneratedFile export not found");

const nextMarker = "\nexports.createSummarySheet = async";
const end = src.indexOf(nextMarker, start);
if (end < 0)
  fail("createSummarySheet marker not found after downloadGeneratedFile");

const replacement = [
  "exports.downloadGeneratedFile = async (req, res, next) => {",
  "  try {",
  "    const {",
  '      storageKey = "",',
  '      filePath = "",',
  '      displayName = "",',
  '      outputType = "",',
  '      queryTablesKey = "",',
  '      deleteAfterDownload = "1",',
  "    } = req.query || {};",
  "",
  "    const safeDisplayName =",
  '      String(displayName || "").trim() ||',
  '      path.basename(String(storageKey || filePath || "download"));',
  "",
  "    let buffer = null;",
  "",
  "    if (storageKey) {",
  "      if (!assertGeneratedStorageKeyAccess(req, storageKey)) {",
  "        return res.status(403).json({",
  "          ok: false,",
  '          code: "GENERATED_FILE_FORBIDDEN",',
  '          error: "다운로드 권한이 없습니다.",',
  "        });",
  "      }",
  "",
  "      buffer = await downloadToBuffer(storageKey);",
  "    } else if (filePath) {",
  "      const safePath = resolveSafeGeneratedLocalPath(filePath);",
  "",
  "      if (!safePath || !fs.existsSync(safePath)) {",
  "        return res.status(404).json({",
  "          ok: false,",
  '          code: "GENERATED_FILE_NOT_FOUND",',
  '          error: "생성 파일을 찾을 수 없습니다.",',
  "        });",
  "      }",
  "",
  "      buffer = fs.readFileSync(safePath);",
  "    } else {",
  "      return res.status(400).json({",
  "        ok: false,",
  '        code: "DOWNLOAD_TARGET_REQUIRED",',
  '        error: "storageKey 또는 filePath가 필요합니다.",',
  "      });",
  "    }",
  "",
  "    res.setHeader(",
  '      "Content-Disposition",',
  "      `attachment; filename*=UTF-8''${encodeDownloadName(safeDisplayName)}`,",
  "    );",
  "    res.setHeader(",
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
  "  } catch (error) {",
  '    console.error("[automation.downloadGeneratedFile]", error);',
  "    next(error);",
  "  }",
  "};",
  "",
].join("\n");

const before = src.slice(0, start);
const after = src.slice(end);
src = before + replacement + after;

fs.writeFileSync(target, src, "utf8");

console.log(
  "[download-cleanup-wire-fix] wrote controllers/automationController.js",
);
console.log(
  "[download-cleanup-wire-fix] next: node --check controllers/automationController.js",
);
