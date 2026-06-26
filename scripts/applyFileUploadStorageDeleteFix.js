#!/usr/bin/env node
/**
 * Patch: make uploaded encrypted file deletion robust.
 *
 * Goal:
 * - Keep download-after-cleanup behavior unchanged.
 * - When an uploaded file is deleted or replaced, delete all possible storage keys
 *   (gcsName + localName + storageName/storageKey legacy fields) instead of only
 *   `localName || gcsName`.
 */

const fs = require('fs');
const path = require('path');

const root = process.cwd();
const filePath = path.join(root, 'controllers', 'fileController.js');

function read(file) {
  return fs.readFileSync(file, 'utf8');
}

function write(file, content) {
  fs.writeFileSync(file, content, 'utf8');
  console.log(`[file-delete-fix] wrote ${path.relative(root, file)}`);
}

function ensureHelper(content) {
  if (content.includes('function collectUploadedFileStorageKeys(')) return content;

  const marker = 'exports.upload = async';
  const idx = content.indexOf(marker);
  if (idx === -1) {
    throw new Error('Could not find insertion point before exports.upload');
  }

  const helper = `
function collectUploadedFileStorageKeys(fileInfo = {}) {
  return Array.from(
    new Set(
      [
        fileInfo.gcsName,
        fileInfo.localName,
        fileInfo.storageName,
        fileInfo.storageKey,
      ]
        .map((value) => String(value || '').trim())
        .filter(Boolean),
    ),
  );
}

async function deleteUploadedFileStorageObjects(fileInfo = {}, context = {}) {
  const keys = collectUploadedFileStorageKeys(fileInfo);
  const deleted = [];
  const errors = [];

  for (const key of keys) {
    try {
      await deleteObject(key);
      deleted.push(key);
    } catch (error) {
      errors.push({ key, message: error?.message || String(error) });
    }
  }

  console.log('[file.storage.delete]', {
    reason: context.reason || 'delete',
    userId: context.userId || '',
    originalName: fileInfo.originalName || '',
    keys,
    deleted,
    errors,
  });

  if (errors.length) {
    const error = new Error('업로드 파일 저장소 삭제 중 오류가 발생했습니다.');
    error.code = 'FILE_STORAGE_DELETE_FAILED';
    error.storageDeleteErrors = errors;
    throw error;
  }

  return { keys, deleted };
}

`;

  return content.slice(0, idx) + helper + content.slice(idx);
}

function replaceDirectDeletes(content) {
  let out = content;

  // Existing-file replacement during upload.
  out = out.replace(
    /await\s+deleteObject\(existingFile\.localName\s*\|\|\s*existingFile\.gcsName\);/g,
    `await deleteUploadedFileStorageObjects(existingFile, {
        reason: 'replace-existing-upload',
        userId: String(user._id),
      });`,
  );

  // Normal file deletion endpoint.
  out = out.replace(
    /await\s+deleteObject\(fileInfo\.localName\s*\|\|\s*fileInfo\.gcsName\);/g,
    `await deleteUploadedFileStorageObjects(fileInfo, {
      reason: 'delete-uploaded-file',
      userId: String(user._id),
    });`,
  );

  return out;
}

function main() {
  if (!fs.existsSync(filePath)) {
    throw new Error(`controllers/fileController.js not found at ${filePath}`);
  }

  let content = read(filePath);
  content = ensureHelper(content);
  content = replaceDirectDeletes(content);
  write(filePath, content);
  console.log('[file-delete-fix] done');
  console.log('[file-delete-fix] next: node --check controllers/fileController.js');
}

main();
