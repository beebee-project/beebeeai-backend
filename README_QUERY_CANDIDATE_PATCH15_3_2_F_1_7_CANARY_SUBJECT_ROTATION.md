# Patch 15.3.2-F.1.7 — Canary Subject Rotation Preparation & Approval Rebinding Readiness

## 목적

현재 `QUERY_CANDIDATE_PLANNER_PROMOTION_ALLOWLIST_SHA256`가 현재 MongoDB 사용자와 매칭되지 않는 상태에서, 현재 사용할 내부 테스트 계정의 canonical Canary subject SHA를 새로 산출하고 F.1.5 이후 승인 바인딩을 안전하게 재발급할 준비를 만든다.

## 범위

- F.1.0~F.1.4 평가 기준선은 변경하지 않는다.
- F.1.4 Candidate Payload SHA는 `928F6A6E0AA8683D63A5A2CB62199FA460EB84494B119EB7E171000843D484EA`로 고정한다.
- 새 subject는 기존 `queryCandidatePlannerInternalCanarySubject` canonical derivation을 사용한다.
- raw account/tenant ID를 tracked output에 기록하지 않는다.
- private rotation plan에는 old/new SHA와 guardrail만 기록한다.
- Railway, route, feature flag, kill switch를 변경하지 않는다.
- provider call, shadow runner, merge adapter를 실행하지 않는다.
- 새 subject가 기존 allowlist와 같으면 회전을 BLOCK한다.
- 새 allowlist SHA가 확정되면 기존 F.1.5 Receipt는 반드시 재발급해야 한다.

## 다음 단계

1. 현재 내부 테스트 계정의 raw account/tenant를 process env에만 로드한다.
2. rotation plan을 생성한다.
3. rotation plan을 검증한다.
4. 새 allowlist SHA를 현재 PowerShell Process에만 적용한다.
5. 기존 F.1.5 builder의 정확한 explicit approval 계약을 read-only로 확인한다.
6. F.1.5 Receipt를 새 allowlist에 맞춰 재발급한다.
7. F.1.6 offline binding을 재검증한다.
8. F.1.6.1 composition을 재검증한다.
9. Provider-free Runtime Preflight E2E를 실행한다.

광범위 rollout 및 production promotion은 승인하지 않는다.
