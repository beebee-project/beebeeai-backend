"use strict";
const assert=require("assert");const {parseRegistry}=require("../automation/queryCandidatePlannerRealShadowEvidenceConfig");const r=parseRegistry("{bad");assert.strictEqual(r.valid,false);assert.strictEqual(r.reason,"REAL_SHADOW_CASE_REGISTRY_JSON_INVALID");console.log("PASS query candidate patch15.3.2 invalid registry fail-closed smoke");
