"use strict";
const assert=require("assert");const s=require("./queryCandidatePatch15_3_2TestSupport");const r=s.build({records:s.realRecords().filter(x=>x.kind==="EXECUTION").slice(0,29)});assert.strictEqual(r.decision,"EVALUATION_BLOCKED");assert.strictEqual(r.reason,"REAL_SHADOW_MINIMUM_SAMPLE_SIZE_NOT_MET");console.log("PASS query candidate patch15.3.2 minimum sample fail-closed smoke");
