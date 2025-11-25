module.exports.parseMacroIntent = function (text) {
  const t = text.toLowerCase().replace(/\s+/g, " ").trim();

  // 1) 글씨 굵게
  if (t.includes("굵게") || t.includes("bold")) {
    const range = extractRange(t) || "A:A";
    return {
      type: "formatRange",
      target: { range },
      style: { bold: true },
    };
  }

  // 2) 배경색 칠하기
  if (t.includes("배경") || t.includes("색칠")) {
    const range = extractRange(t) || "A:A";
    const color = extractColor(t) || "#FFFF00";
    return {
      type: "formatRange",
      target: { range },
      style: { fillColor: color },
    };
  }

  // 3) 행 삽입
  if (t.includes("행") && (t.includes("추가") || t.includes("삽입"))) {
    const row = extractNumber(t) || 1;
    return {
      type: "insertRow",
      target: { rowIndex: row, position: "above" },
    };
  }

  // 4) 시트 생성
  if (t.includes("시트") && (t.includes("생성") || t.includes("만들어"))) {
    const sheetName = extractSheetName(t) || "NewSheet";
    return {
      type: "createSheet",
      name: sheetName,
    };
  }

  // 5) 시트 복사
  if (t.includes("시트") && t.includes("복사")) {
    const name = extractSheetName(t) || "Backup";
    return {
      type: "duplicateSheet",
      newSheetName: name,
    };
  }

  // fallback
  return {
    type: "unknown",
    text,
  };
};

/* ---------- Helper functions ---------- */

function extractRange(text) {
  // 'A열', 'B 열', 'C column', 'A:A' 등 처리
  const colMatch = text.match(/([a-zA-Z])\s*열/);
  if (colMatch)
    return `${colMatch[1].toUpperCase()}:${colMatch[1].toUpperCase()}`;

  const rangeMatch = text.match(/[A-Z]:[A-Z]/i);
  if (rangeMatch) return rangeMatch[0].toUpperCase();

  return null;
}

function extractColor(text) {
  if (text.includes("노란")) return "#FFFF00";
  if (text.includes("빨간")) return "#FF0000";
  if (text.includes("파란")) return "#0000FF";
  return null;
}

function extractNumber(text) {
  const m = text.match(/\d+/);
  return m ? parseInt(m[0], 10) : null;
}

function extractSheetName(text) {
  const m = text.match(/["'](.+?)["']/);
  return m ? m[1] : null;
}
