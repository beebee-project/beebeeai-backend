"use strict";
const assert=require("assert");const s=require("./queryCandidatePatch15_3_2TestSupport");const r=s.build();assert.strictEqual(r.evidenceBundle.llmPolicy.mode,"SEMANTIC_PROFILER_ONLY");assert.strictEqual(r.evidenceBundle.llmPolicy.plannerEscalationAllowed,false);console.log("PASS query candidate patch15.3.2 profiler-only policy smoke");
