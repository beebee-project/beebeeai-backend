"use strict";
const assert=require("assert");const s=require("./queryCandidatePatch15_3_2TestSupport");const r=s.build();assert.strictEqual(r.reports.accuracy.decision,"EVALUATION_PASS");assert.strictEqual(r.reports.accuracy.caseCount,s.accuracyDataset().cases.length);console.log("PASS query candidate patch15.3.2 accuracy integration smoke");
