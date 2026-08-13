# Patch 15.3.3-B-4-F-A.2 — Readiness-aware Runtime Compatibility Repair

## 목적

B-4-F-A.1 실제 Feature Control 진단에서 `PRODUCTION_CANDIDATE_MERGE`가
`READINESS_EVIDENCE_INVALID`로 차단되어 F.1.6 및 A Gate가 연쇄 BLOCK되는
호환성 결함을 최소 범위로 복구한다.

## 핵심 설계

- F.1.6 Approval Binding Gate는 byte-identical 유지한다.
- B-2 Live Bootstrap Runtime은 byte-identical 유지한다.
- A Gate에서 F.1.6 호출에만 readiness-aware Feature Control bridge를 적용한다.
- bridge는 `PRODUCTION_CANDIDATE_MERGE`에만 정확한 sanitized Patch 13.3
  readiness를 주입한다.
- Production Route / Ready Assignment에는 readiness를 주입하지 않는다.
- B-2 runtime의 `bootstrapObserveOnlyMergeAdapter`와 `readinessGate: null`은 유지한다.
- Railway 변수, route, provider runner를 변경하지 않는다.
- 이 패치 자체는 Provider를 호출하거나 실제 사용자 요청을 실행하지 않는다.

## 보호 predecessor

- F.1.6 Gate: `ED43CFAF798FE904EDB0308EE82EFDB5A17D599EC44416072DE152F625E436E7`
- B-2 Runtime: `F52737193BCCA38C8534BB12698D96E2B162381E62AA7170B82AB0E01C19519A`
- Feature Control: `E80A47537ECDB4454C6120693A9F3E725F74AC986C42ABF52E6AD163B30EAB07`
- A Gate predecessor: `4585B4549B0F756274F47FBB9089E56A07D21C6EFE3C1929214E856B068B5498`

## 신규 identities

- Readiness bridge SHA-256: `77DB527F808BBB61BD63BD61913E01A489AB25E154C5D4C0E67DAC730AB81259`
- Patched A Gate SHA-256: `DFE04C089F0F514FA60026BE9FD3EF4EDA0DD584B4B55ECC6C2AF54FDECECD7D`
- Sanitized readiness SHA-256: `46D1211AF4F318DAB91D137F0728C3AE6F246CD8B85A2582802CCB6DB1475AC4`
- Historical readiness source SHA-256: `33B70E7B4278CBC7E6F66D10CC6AA0F8FA7219A46E553EAD70612494E654F7D5`

## 검증

```powershell
node --check .\automation\queryCandidatePlannerInternalCanaryBootstrapReadinessBridge.js
node --check .\automation\queryCandidatePlannerInternalCanaryLiveBootstrapGate.js
node .\tests\queryCandidatePatch15_3_3_B_4_F_A_2ReadinessAwareFeatureControlSmokeTest.js
node .\tests\queryCandidatePatch15_3_3_B_4_F_A_2FailClosedReadinessSmokeTest.js
node .\tests\queryCandidatePatch15_3_3_B_4_F_A_2RealF16FeatureControlCompatibilitySmokeTest.js
node .\tests\queryCandidatePatch15_3_3_B_4_F_A_2SourceContractSmokeTest.js
node .\tests\queryCandidatePatch15_3_3_B_4_F_A_2ProtectedSourcesSmokeTest.js
node .\scripts\queryCandidatePlannerVerifyPatch15_3_3_B_4_F_A_2.js
```

## 주의

과거 B-2 exact source integrity smoke는 A Gate의 과거 SHA를 고정하고 있으므로
A.2 적용 뒤 A Gate supersession 때문에 실패할 수 있다. A.2 protected-source smoke가
새 누적 기준이며, F.1.6/B-2 Runtime/Feature Control의 byte identity는 계속 고정한다.
