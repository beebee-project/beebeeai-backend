const { readJsonObject } = require("../utils/storage");
const { readEncryptedQueryJson } = require("./encryptedJsonStorageService");

const ENCRYPTED_QUERY_JSON_PREFIX = "query-json/encrypted/";

function isEncryptedQueryTablesKey(queryTablesKey = "") {
  return String(queryTablesKey || "").startsWith(ENCRYPTED_QUERY_JSON_PREFIX);
}

function storageStatusCode(error) {
  return Number(
    error?.code ||
      error?.status ||
      error?.statusCode ||
      error?.response?.statusCode ||
      error?.response?.status ||
      0,
  );
}

function isMissingQueryTablesObjectError(error) {
  const message = String(error?.message || "");

  return (
    storageStatusCode(error) === 404 ||
    /No such object/i.test(message) ||
    /not found/i.test(message)
  );
}

function createQueryTablesReadError({
  code,
  message,
  status,
  queryTablesKey = "",
  cause = null,
}) {
  const error = new Error(message);

  error.code = code;
  error.status = status;
  error.queryTablesKey = queryTablesKey;
  error.cause = cause;

  return error;
}

async function readQueryTablesPayload(queryTablesKey) {
  const key = String(queryTablesKey || "").trim();

  if (!key) {
    throw createQueryTablesReadError({
      code: "QUERY_TABLES_KEY_REQUIRED",
      message: "queryTablesKey가 필요합니다.",
      status: 400,
    });
  }

  try {
    if (isEncryptedQueryTablesKey(key)) {
      const decrypted = await readEncryptedQueryJson(key);

      if (!decrypted || typeof decrypted !== "object") {
        throw createQueryTablesReadError({
          code: "QUERY_TABLE_INVALID_ENCRYPTED_PAYLOAD",
          message: "암호화된 작업 데이터를 정상적으로 복호화하지 못했습니다.",
          status: 422,
          queryTablesKey: key,
        });
      }

      /*
       * 구버전에서 { payload: {...} } 형태로 저장된 객체와
       * 현재 payload 직접 저장 형식을 모두 지원한다.
       */
      return decrypted.payload || decrypted;
    }

    /*
     * 로컬 회귀 및 기존 평문 query-tables 호환.
     * 운영 신규 객체는 이 분기를 사용하지 않는다.
     */
    return await readJsonObject(key);
  } catch (error) {
    if (
      error?.code === "QUERY_TABLES_KEY_REQUIRED" ||
      error?.code === "QUERY_TABLE_INVALID_ENCRYPTED_PAYLOAD"
    ) {
      throw error;
    }

    if (isMissingQueryTablesObjectError(error)) {
      console.warn("[query-tables] storage object missing", {
        queryTablesKey: key,
        originalMessage: error?.message || String(error),
      });

      throw createQueryTablesReadError({
        code: "QUERY_TABLE_NOT_FOUND",
        message:
          "작업 데이터가 만료되었거나 존재하지 않습니다. 다시 준비해주세요.",
        status: 410,
        queryTablesKey: key,
        cause: error,
      });
    }

    throw error;
  }
}

module.exports = {
  ENCRYPTED_QUERY_JSON_PREFIX,
  isEncryptedQueryTablesKey,
  isMissingQueryTablesObjectError,
  readQueryTablesPayload,
};
