"use strict";
const assert=require("assert");const {encryptEvidencePayload,decryptEvidencePayload}=require("../automation/queryCandidatePlannerRealShadowEvidenceCrypto");const encrypted=encryptEvidencePayload({ok:true},"a".repeat(64));assert.throws(()=>decryptEvidencePayload(encrypted,"b".repeat(64)));console.log("PASS query candidate patch15.3.2 wrong decryption secret fail-closed smoke");
