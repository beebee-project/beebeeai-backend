const PAGE_VERSION = "query_candidate_planner_internal_preview_page_v1";

function buildQueryCandidatePlannerInternalPreviewHtml({ nonce = "" } = {}) {
  const safeNonce = String(nonce || "").replace(/[^a-zA-Z0-9_-]/g, "");
  return `<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>BeeBee AI · 후보군 Shadow 내부 미리보기</title>
  <style nonce="${safeNonce}">
    :root { color-scheme: light; font-family: Inter, Pretendard, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; }
    * { box-sizing: border-box; }
    body { margin: 0; background: #f5f7fb; color: #172033; }
    main { width: min(1500px, calc(100% - 32px)); margin: 24px auto 48px; }
    header { display: flex; gap: 16px; align-items: flex-start; justify-content: space-between; margin-bottom: 18px; }
    h1 { margin: 0 0 6px; font-size: 24px; }
    p { margin: 0; color: #5f6b7a; }
    .badges { display: flex; gap: 8px; flex-wrap: wrap; justify-content: flex-end; }
    .badge { border: 1px solid #cfd7e6; background: #fff; border-radius: 999px; padding: 6px 10px; font-size: 12px; font-weight: 700; }
    .badge.safe { color: #0f6b46; border-color: #a8dbc8; background: #ecfaf4; }
    .panel { background: #fff; border: 1px solid #dfe5ef; border-radius: 14px; box-shadow: 0 8px 24px rgba(35, 48, 74, .06); padding: 16px; margin-bottom: 16px; }
    .controls { display: grid; grid-template-columns: minmax(240px, 1fr) auto auto auto; gap: 10px; align-items: center; }
    input, select, button { height: 40px; border-radius: 9px; border: 1px solid #cbd4e3; background: #fff; padding: 0 12px; font: inherit; }
    button { cursor: pointer; font-weight: 700; }
    button.primary { background: #172033; color: #fff; border-color: #172033; }
    button:disabled { opacity: .5; cursor: not-allowed; }
    .message { min-height: 22px; margin-top: 10px; font-size: 13px; color: #5f6b7a; }
    .summary { display: grid; grid-template-columns: repeat(5, minmax(130px, 1fr)); gap: 12px; }
    .metric { background: #f8faff; border: 1px solid #e3e8f2; border-radius: 12px; padding: 14px; }
    .metric span { display: block; color: #667085; font-size: 12px; margin-bottom: 7px; }
    .metric strong { font-size: 22px; }
    .table-wrap { overflow: auto; max-height: 62vh; border: 1px solid #e3e8f2; border-radius: 12px; }
    table { width: 100%; border-collapse: collapse; min-width: 1180px; }
    th, td { padding: 11px 10px; border-bottom: 1px solid #e7ebf2; text-align: left; font-size: 12px; white-space: nowrap; }
    th { position: sticky; top: 0; background: #f8faff; z-index: 1; color: #475467; }
    tbody tr { cursor: pointer; }
    tbody tr:hover { background: #f8faff; }
    .status { font-weight: 800; }
    .MATCH { color: #0f6b46; }
    .PARTIAL_MATCH { color: #9a6700; }
    .MISMATCH, .FAILED_SAFE, .TIMEOUT_SAFE { color: #b42318; }
    .empty { padding: 36px; text-align: center; color: #667085; }
    details { margin-top: 14px; }
    pre { overflow: auto; max-height: 420px; background: #0f172a; color: #dbeafe; border-radius: 12px; padding: 14px; font-size: 12px; }
    .filters { display: flex; gap: 8px; margin-bottom: 10px; flex-wrap: wrap; }
    .filters select { min-width: 160px; }
    @media (max-width: 900px) {
      header { display: block; }
      .badges { justify-content: flex-start; margin-top: 12px; }
      .controls { grid-template-columns: 1fr 1fr; }
      .controls input { grid-column: 1 / -1; }
      .summary { grid-template-columns: repeat(2, 1fr); }
    }
  </style>
</head>
<body data-page-version="${PAGE_VERSION}">
<main>
  <header>
    <div>
      <h1>후보군 Shadow 내부 미리보기</h1>
      <p>Primary 응답과 Shadow Planner의 비교 관찰만 표시합니다.</p>
    </div>
    <div class="badges">
      <span class="badge safe">읽기 전용</span>
      <span class="badge">Production 반영 없음</span>
      <span class="badge">메모리 저장</span>
    </div>
  </header>

  <section class="panel">
    <div class="controls">
      <input id="token" type="password" autocomplete="off" placeholder="내부 Preview Token">
      <button id="connect" class="primary" type="button">연결</button>
      <button id="refresh" type="button" disabled>새로고침</button>
      <button id="disconnect" type="button">세션 해제</button>
    </div>
    <div id="message" class="message">토큰은 URL이나 로컬 저장소에 기록되지 않고 현재 탭의 sessionStorage에만 유지됩니다.</div>
  </section>

  <section class="panel summary">
    <div class="metric"><span>관찰 건수</span><strong id="total">0</strong></div>
    <div class="metric"><span>완료</span><strong id="completed">0</strong></div>
    <div class="metric"><span>Mismatch</span><strong id="mismatch">0</strong></div>
    <div class="metric"><span>평균 지연</span><strong id="latency">0 ms</strong></div>
    <div class="metric"><span>Provider 호출</span><strong id="providerCalls">0</strong></div>
  </section>

  <section class="panel">
    <div class="filters">
      <select id="statusFilter"><option value="">모든 상태</option><option>COMPLETED</option><option>COMPLETED_SAFE</option><option>BLOCKED</option><option>FAILED_SAFE</option><option>TIMEOUT_SAFE</option><option>BOUNDARY_FAILED_SAFE</option></select>
      <select id="verdictFilter"><option value="">모든 비교 결과</option><option>MATCH</option><option>PARTIAL_MATCH</option><option>MISMATCH</option><option>NO_SHADOW_CANDIDATES</option></select>
      <label><input id="autoRefresh" type="checkbox"> 5초 자동 새로고침</label>
    </div>
    <div class="table-wrap">
      <table>
        <thead><tr><th>시각</th><th>상태</th><th>비교</th><th>Primary</th><th>Shadow</th><th>공통</th><th>Top-1</th><th>Jaccard</th><th>순위 일치</th><th>Provider</th><th>Cache</th><th>지연</th></tr></thead>
        <tbody id="rows"></tbody>
      </table>
      <div id="empty" class="empty">내부 토큰으로 연결하면 관찰 결과가 표시됩니다.</div>
    </div>
    <details id="details"><summary>선택한 관찰의 정제 JSON</summary><pre id="json">항목을 선택하세요.</pre></details>
  </section>
</main>
<script nonce="${safeNonce}">
(() => {
  "use strict";
  const TOKEN_KEY = "beebee.queryCandidate.internalPreviewToken";
  const tokenInput = document.getElementById("token");
  const message = document.getElementById("message");
  const rows = document.getElementById("rows");
  const empty = document.getElementById("empty");
  const refreshButton = document.getElementById("refresh");
  const autoRefresh = document.getElementById("autoRefresh");
  const statusFilter = document.getElementById("statusFilter");
  const verdictFilter = document.getElementById("verdictFilter");
  const base = window.location.pathname.replace(/\/$/, "");
  let timer = null;
  let currentEntries = [];

  function token() { return sessionStorage.getItem(TOKEN_KEY) || ""; }
  function headers() { return { "x-beebee-internal-preview-token": token() }; }
  function setMessage(value, error = false) {
    message.textContent = value;
    message.style.color = error ? "#b42318" : "#5f6b7a";
  }
  function count(summary, group, key) {
    return Number(summary && summary[group] && summary[group][key] || 0);
  }
  function updateSummary(summary = {}) {
    document.getElementById("total").textContent = Number(summary.total || 0);
    document.getElementById("completed").textContent = count(summary, "statusCounts", "COMPLETED") + count(summary, "statusCounts", "COMPLETED_SAFE");
    document.getElementById("mismatch").textContent = count(summary, "verdictCounts", "MISMATCH");
    document.getElementById("latency").textContent = Number(summary.averageLatencyMs || 0) + " ms";
    document.getElementById("providerCalls").textContent = Number(summary.providerCallTotal || 0);
  }
  function cell(value, className = "") {
    const td = document.createElement("td");
    td.textContent = value == null || value === "" ? "-" : String(value);
    if (className) td.className = className;
    return td;
  }
  function render(entries = []) {
    currentEntries = entries;
    rows.replaceChildren();
    empty.style.display = entries.length ? "none" : "block";
    entries.forEach((entry, index) => {
      const tr = document.createElement("tr");
      const comparison = entry.comparison || {};
      const counts = comparison.counts || {};
      const metrics = comparison.metrics || {};
      const cache = entry.cacheLifecycle || {};
      tr.append(
        cell(new Date(entry.observedAt).toLocaleString("ko-KR")),
        cell(entry.status, "status " + entry.status),
        cell(comparison.verdict || "NOT_AVAILABLE", comparison.verdict || ""),
        cell(counts.primary), cell(counts.shadow), cell(counts.shared),
        cell(metrics.top1Same === true ? "일치" : metrics.top1Same === false ? "불일치" : "-"),
        cell(metrics.jaccard), cell(metrics.rankAgreement),
        cell(entry.shadow && entry.shadow.providerCallCount),
        cell(cache.cacheReadAllowed ? "READ" : cache.cacheWriteAllowed ? "WRITE" : "OFF"),
        cell(Number(entry.latencyMs || 0) + " ms")
      );
      tr.addEventListener("click", () => {
        document.getElementById("json").textContent = JSON.stringify(currentEntries[index], null, 2);
        document.getElementById("details").open = true;
      });
      rows.appendChild(tr);
    });
  }
  async function requestJson(path) {
    const response = await fetch(path, { headers: headers(), credentials: "same-origin", cache: "no-store" });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(payload.error || payload.code || "요청 실패");
    return payload;
  }
  async function refresh() {
    if (!token()) { setMessage("내부 Preview Token을 입력하세요.", true); return; }
    refreshButton.disabled = true;
    try {
      const params = new URLSearchParams({ limit: "100" });
      if (statusFilter.value) params.set("status", statusFilter.value);
      if (verdictFilter.value) params.set("verdict", verdictFilter.value);
      const [status, observations] = await Promise.all([
        requestJson(base + "/status"),
        requestJson(base + "/observations?" + params.toString()),
      ]);
      updateSummary(observations.summary || status.store || {});
      render(observations.entries || []);
      setMessage("연결됨 · Production 후보 선택과 병합 기능은 제공되지 않습니다.");
    } catch (error) {
      render([]);
      updateSummary({});
      setMessage(error.message || "내부 미리보기 연결 실패", true);
    } finally {
      refreshButton.disabled = false;
    }
  }
  function updateTimer() {
    if (timer) clearInterval(timer);
    timer = autoRefresh.checked ? setInterval(refresh, 5000) : null;
  }
  document.getElementById("connect").addEventListener("click", () => {
    const value = tokenInput.value.trim();
    if (!value) { setMessage("토큰을 입력하세요.", true); return; }
    sessionStorage.setItem(TOKEN_KEY, value);
    tokenInput.value = "";
    refresh();
  });
  document.getElementById("disconnect").addEventListener("click", () => {
    sessionStorage.removeItem(TOKEN_KEY);
    render([]); updateSummary({});
    setMessage("현재 탭의 내부 토큰을 삭제했습니다.");
  });
  refreshButton.addEventListener("click", refresh);
  autoRefresh.addEventListener("change", updateTimer);
  statusFilter.addEventListener("change", refresh);
  verdictFilter.addEventListener("change", refresh);
  if (token()) refresh();
})();
</script>
</body>
</html>`;
}

module.exports = Object.freeze({
  PAGE_VERSION,
  buildQueryCandidatePlannerInternalPreviewHtml,
});
