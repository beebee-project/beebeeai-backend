const CLUSTER_DEFS = {
  id_key: {
    aliases: [
      "직원id",
      "직원 id",
      "사번",
      "직원번호",
      "employeeid",
      "empid",
      "id",
    ],
    role: "lookup",
    type: "text",
  },
  person_name: {
    aliases: ["이름", "성명", "직원명", "name", "employee name"],
    role: "lookup",
    type: "text",
  },
  group_label: {
    aliases: ["부서", "팀", "조직", "그룹", "category", "group", "department"],
    role: "group",
    type: "text",
  },
  amount_metric: {
    aliases: [
      "연봉",
      "급여",
      "금액",
      "매출",
      "비용",
      "salary",
      "amount",
      "revenue",
    ],
    role: "metric",
    type: "number",
  },
  date_field: {
    aliases: ["입사일", "날짜", "일자", "date", "joined date"],
    role: "date",
    type: "date",
  },
  rating_label: {
    aliases: ["평가등급", "평가 등급", "등급", "rating", "grade"],
    role: "group",
    type: "ordered_text",
  },
  rank_label: {
    aliases: ["직급", "직위", "rank", "position", "title"],
    role: "group",
    type: "ordered_text",
  },
};

function norm(s = "") {
  return String(s)
    .toLowerCase()
    .replace(/\(.*?\)/g, "")
    .replace(/[^\p{Letter}\p{Number}]+/gu, "")
    .trim();
}

function inferClusterFromText(text = "") {
  const base = norm(text);
  if (!base) return null;

  for (const [key, def] of Object.entries(CLUSTER_DEFS)) {
    const aliases = (def.aliases || []).map(norm).filter(Boolean);
    if (
      aliases.some((a) => a === base || base.includes(a) || a.includes(base))
    ) {
      return key;
    }
  }

  return null;
}

function inferClusterCandidate(
  header = "",
  sampleValues = [],
  dominantType = "",
) {
  const byHeader = inferClusterFromText(header);
  if (byHeader) return byHeader;

  const joinedSamples = Array.isArray(sampleValues)
    ? sampleValues.slice(0, 5).join(" ")
    : "";
  const bySamples = inferClusterFromText(joinedSamples);
  if (bySamples) return bySamples;

  const t = String(dominantType || "").toLowerCase();
  if (t === "date") return "date_field";
  if (t === "number") return "amount_metric";

  return null;
}

function getClusterRole(clusterKey = "") {
  return CLUSTER_DEFS[clusterKey]?.role || null;
}

function getClusterType(clusterKey = "") {
  return CLUSTER_DEFS[clusterKey]?.type || null;
}

module.exports = {
  CLUSTER_DEFS,
  inferClusterFromText,
  inferClusterCandidate,
  getClusterRole,
  getClusterType,
};
