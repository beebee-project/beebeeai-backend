const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

function assertCodecFunction(name, value) {
  if (typeof value !== "function") {
    throw new TypeError(`${name} 함수가 필요합니다.`);
  }
}

function cacheFileName(cacheKey) {
  return `${crypto.createHash("sha256").update(String(cacheKey)).digest("hex")}.enc`;
}

function createEncryptedSemanticProfilerCache({
  rootDir,
  encryptBuffer,
  decryptBuffer,
} = {}) {
  if (!rootDir) throw new Error("암호화 캐시 rootDir이 필요합니다.");
  assertCodecFunction("encryptBuffer", encryptBuffer);
  assertCodecFunction("decryptBuffer", decryptBuffer);
  const absoluteRoot = path.resolve(rootDir);

  function filePath(cacheKey) {
    return path.join(absoluteRoot, cacheFileName(cacheKey));
  }

  return {
    async get(cacheKey) {
      const target = filePath(cacheKey);
      if (!fs.existsSync(target)) return null;
      const encrypted = fs.readFileSync(target);
      const plaintext = await decryptBuffer(encrypted, {
        purpose: "query-semantic-profiler",
        cacheKey,
      });
      const buffer = Buffer.isBuffer(plaintext)
        ? plaintext
        : Buffer.from(plaintext);
      return JSON.parse(buffer.toString("utf8"));
    },

    async set(cacheKey, profile) {
      fs.mkdirSync(absoluteRoot, { recursive: true });
      const plaintext = Buffer.from(`${JSON.stringify(profile)}\n`, "utf8");
      const encrypted = await encryptBuffer(plaintext, {
        purpose: "query-semantic-profiler",
        cacheKey,
      });
      if (!Buffer.isBuffer(encrypted)) {
        throw new TypeError("encryptBuffer는 Buffer를 반환해야 합니다.");
      }
      const target = filePath(cacheKey);
      const temporary = `${target}.${process.pid}.${Date.now()}.tmp`;
      fs.writeFileSync(temporary, encrypted);
      fs.renameSync(temporary, target);
      return target;
    },

    async delete(cacheKey) {
      const target = filePath(cacheKey);
      if (!fs.existsSync(target)) return false;
      fs.unlinkSync(target);
      return true;
    },

    pathFor(cacheKey) {
      return filePath(cacheKey);
    },
  };
}

module.exports = {
  cacheFileName,
  createEncryptedSemanticProfilerCache,
};
