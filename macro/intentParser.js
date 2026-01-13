function normalize(text) {
  return text.replace(/\s+/g, " ").trim();
}

// ✅ 범위 인식: 블록 / 단일셀 / 열 / 행
function detectRange(text) {
  const upper = text.toUpperCase();

  // A1~C10, A1부터 C10까지
  const blockMatch = upper.match(
    /([A-Z][0-9]+)\s*(?:부터|~|-)\s*([A-Z][0-9]+)/
  );
  if (blockMatch) {
    return `${blockMatch[1]}:${blockMatch[2]}`;
  }

  // 단일 셀 (A1, B3 등) – 첫 번째 매치만 사용
  const cellMatch = upper.match(/([A-Z][0-9]+)/);
  if (cellMatch) {
    return cellMatch[1];
  }

  // B열, C 열
  const colMatch = upper.match(/([A-Z])\s*열/);
  if (colMatch) {
    const col = colMatch[1];
    return `${col}:${col}`;
  }

  // 3행, 10 행
  const rowMatch = upper.match(/([0-9]+)\s*행/);
  if (rowMatch) {
    const row = rowMatch[1];
    return `${row}:${row}`;
  }

  return null;
}

// ✅ 복사/이동용 범위 2개 추출(A열, B열 / A1:A10, B1:B10 등)
function detectTwoRanges(text) {
  const upper = text.toUpperCase();
  const tokens =
    upper.match(/[A-Z]+[0-9]*:[A-Z]+[0-9]*|[A-Z][0-9]+|[A-Z]:[A-Z]|[A-Z]/g) ||
    [];

  const normalizeToken = (t) => {
    if (/^[A-Z]$/.test(t)) return `${t}:${t}`; // B -> B:B
    return t;
  };

  const from = tokens[0] ? normalizeToken(tokens[0]) : null;
  const to = tokens[1] ? normalizeToken(tokens[1]) : null;

  return { from, to };
}

// ✅ 열 정보 추출 (insertColumn / deleteColumn 용)
function detectColumnInfo(text) {
  const upper = text.toUpperCase();

  // B열, C 열
  const colLetterMatch = upper.match(/([A-Z])\s*열/);
  if (colLetterMatch) {
    const letter = colLetterMatch[1];
    return { letter, index: null };
  }

  // 2열, 3 열 (숫자 열)
  const colIndexMatch = upper.match(/([0-9]+)\s*열/);
  if (colIndexMatch) {
    const idx = parseInt(colIndexMatch[1], 10);
    return { letter: null, index: idx };
  }

  return { letter: null, index: null };
}

function detectColors(text) {
  const t = text.toLowerCase();
  const colorMap = [
    { keys: ["노란", "노랑", "yellow"], value: "#FFFF00" },
    { keys: ["빨간", "빨강", "red"], value: "#FF0000" },
    { keys: ["파란", "파랑", "blue"], value: "#0000FF" },
    { keys: ["초록", "green"], value: "#00AA00" },
    { keys: ["회색", "grey", "gray"], value: "#CCCCCC" },
    { keys: ["검정", "까만", "black"], value: "#000000" },
    { keys: ["흰색", "흰", "white"], value: "#FFFFFF" },
  ];

  let fillColor = null;
  let fontColor = null;

  for (const { keys, value } of colorMap) {
    if (keys.some((k) => t.includes(k))) {
      if (t.includes("배경") || t.includes("바탕") || t.includes("색칠")) {
        fillColor = value;
      } else if (
        t.includes("글씨") ||
        t.includes("폰트") ||
        t.includes("글자")
      ) {
        fontColor = value;
      } else {
        fillColor = value; // 애매하면 배경색
      }
      break;
    }
  }

  return { fillColor, fontColor };
}

function detectStyleFlags(text) {
  const t = text.toLowerCase();
  const style = {};

  // 굵게 / bold
  if (t.includes("굵게") || t.includes("bold") || t.includes("강조")) {
    style.bold = true;
  }

  // 이탤릭
  if (t.includes("이탤릭") || t.includes("기울임") || t.includes("italic")) {
    style.italic = true;
  }

  // 밑줄
  if (t.includes("밑줄") || t.includes("underline")) {
    style.underline = true;
  }

  // 정렬
  if (
    t.includes("가운데 정렬") ||
    t.includes("중앙 정렬") ||
    t.includes("센터")
  ) {
    style.horizontalAlign = "Center";
  } else if (t.includes("오른쪽 정렬") || t.includes("우측 정렬")) {
    style.horizontalAlign = "Right";
  } else if (t.includes("왼쪽 정렬") || t.includes("좌측 정렬")) {
    style.horizontalAlign = "Left";
  }

  // 테두리
  if (t.includes("테두리") || t.includes("border")) {
    style.border = "thin";
  }

  return style;
}

// 값 추출: " '합계' 라고 입력 " 패턴 우선
function detectValue(text) {
  const quoteMatch = text.match(/["“”‘’']([^"“”‘’']+)["“”‘’']/);
  if (quoteMatch) {
    return quoteMatch[1].trim();
  }

  // 따옴표가 없으면 '입력/적어/써줘/기록' 뒤를 값으로 추정 (MVP)
  const parts = text.split(/입력|적어|써줘|써 줘|기록/);
  if (parts.length > 1) {
    return parts[1]
      .replace(/해\s*줘.*/, "")
      .replace(/해주세요.*/, "")
      .trim();
  }

  return null;
}

// 필터 기준 값 추출: "완료만", "'완료'만" 등에서 값 추정
function detectFilterCriteria(text) {
  // 1) 따옴표가 있으면 그 안의 값을 우선 사용
  const quoted = text.match(/["“”‘’']([^"“”‘’']+)["“”‘’']/);
  if (quoted) {
    return quoted[1].trim();
  }

  // 2) "<단어>만" 패턴에서 단어 뽑기 (예: 완료만, 지연만)
  const m = text.match(/([^\s"“”‘’']+)\s*만/);
  if (m) {
    return m[1].trim();
  }

  return null;
}

// 시트 이름 추출: 따옴표 기반 우선
function detectSheetNameInQuotes(text) {
  const m = text.match(/["“”‘’']([^"“”‘’']+)["“”‘’']/);
  return m ? m[1].trim() : null;
}

// 시트 이름 추정(따옴표 없을 때): "<단어> 시트" 패턴
function detectSheetNameLoose(text) {
  const m = text.match(/([^\s"“”‘’']+)\s*시트/);
  return m ? m[1].trim() : null;
}

function parseMacroIntent(text) {
  if (!text || typeof text !== "string") {
    return { type: "unknown", text: "" };
  }

  const originalText = text;
  const tNorm = normalize(text.toLowerCase());

  // ─────────────────────────────
  // 1) 형식(서식) 관련 (1단계)
  // ─────────────────────────────
  const hasFormatKeyword =
    tNorm.includes("배경") ||
    tNorm.includes("색칠") ||
    tNorm.includes("색상") ||
    tNorm.includes("색 ") ||
    tNorm.includes("굵게") ||
    tNorm.includes("bold") ||
    tNorm.includes("글씨") ||
    tNorm.includes("폰트") ||
    tNorm.includes("글자") ||
    tNorm.includes("정렬") ||
    tNorm.includes("테두리");

  if (hasFormatKeyword) {
    const range = detectRange(originalText);
    const { fillColor, fontColor } = detectColors(originalText);
    const flags = detectStyleFlags(originalText);

    const style = { ...flags };
    if (fillColor) style.fillColor = fillColor;
    if (fontColor) style.fontColor = fontColor;

    return {
      type: "formatRange",
      target: { range },
      style,
      text: originalText,
    };
  }

  // ─────────────────────────────
  // 2) 값 입력 (setValue)
  // ─────────────────────────────
  const hasSetValueKeyword =
    tNorm.includes("입력") ||
    tNorm.includes("적어") ||
    tNorm.includes("써줘") ||
    tNorm.includes("써 줘") ||
    tNorm.includes("기록");

  if (hasSetValueKeyword) {
    const range = detectRange(originalText) || "A1";
    const value = detectValue(originalText) || "";

    return {
      type: "setValue",
      target: { range },
      value,
      text: originalText,
    };
  }

  // ─────────────────────────────
  // 3) 복사 (copyRange)
  // ─────────────────────────────
  if (tNorm.includes("복사")) {
    const { from, to } = detectTwoRanges(originalText);
    return {
      type: "copyRange",
      from: from || "A1:A1",
      to: to || from || "B1:B1",
      text: originalText,
    };
  }

  // ─────────────────────────────
  // 4) 이동 (moveRange)
  // ─────────────────────────────
  if (tNorm.includes("이동") || tNorm.includes("옮겨")) {
    const { from, to } = detectTwoRanges(originalText);
    return {
      type: "moveRange",
      from: from || "A1:A1",
      to: to || "B1:B1",
      text: originalText,
    };
  }

  // ─────────────────────────────
  // X) 정렬 (sortRange)
  // ─────────────────────────────
  const hasSortKeyword =
    tNorm.includes("정렬") ||
    tNorm.includes("오름차순") ||
    tNorm.includes("내림차순") ||
    tNorm.includes("역순") ||
    tNorm.includes("순으로");

  if (hasSortKeyword) {
    const colInfo = detectColumnInfo(originalText); // { letter, index }

    let direction = "ascending";
    if (
      tNorm.includes("내림") ||
      tNorm.includes("큰 순") ||
      tNorm.includes("z-") ||
      tNorm.includes("z~") ||
      tNorm.includes("역순")
    ) {
      direction = "descending";
    }

    return {
      type: "sortRange",
      column: colInfo, // { letter, index }
      direction,
      text: originalText,
    };
  }

  // ─────────────────────────────
  // X) 필터 (filterRange)
  // ─────────────────────────────
  const hasFilterKeyword =
    tNorm.includes("필터") ||
    tNorm.includes("걸러") ||
    tNorm.includes("만 보") || // "완료만 보이게"
    tNorm.includes("만 남기"); // "완료만 남기고"

  if (hasFilterKeyword) {
    const colInfo = detectColumnInfo(originalText); // { letter, index }
    const criteria = detectFilterCriteria(originalText) || "";

    return {
      type: "filterRange",
      column: colInfo,
      criteria,
      text: originalText,
    };
  }

  // ─────────────────────────────
  // 5) 행/열 삭제 (deleteRow / deleteColumn)
  // ─────────────────────────────
  const hasDeleteKeyword = tNorm.includes("삭제") || tNorm.includes("지워");

  if (hasDeleteKeyword && tNorm.includes("행")) {
    const rowMatch = tNorm.match(/([0-9]+)\s*행/);
    const rowIndex = rowMatch ? parseInt(rowMatch[1], 10) : 1;
    return {
      type: "deleteRow",
      rowIndex,
      text: originalText,
    };
  }

  if (hasDeleteKeyword && tNorm.includes("열")) {
    const colInfo = detectColumnInfo(originalText);
    return {
      type: "deleteColumn",
      column: colInfo,
      text: originalText,
    };
  }

  // ─────────────────────────────
  // 6) 행 삽입 (insertRow)
  // ─────────────────────────────
  if (
    tNorm.includes("행") &&
    (tNorm.includes("추가") || tNorm.includes("삽입"))
  ) {
    const rowMatch = tNorm.match(/([0-9]+)\s*행/);
    const rowIndex = rowMatch ? parseInt(rowMatch[1], 10) : 1;
    return {
      type: "insertRow",
      rowIndex,
      position: "above",
      text: originalText,
    };
  }

  // ─────────────────────────────
  // 7) 열 삽입 (insertColumn)
  // ─────────────────────────────
  if (
    tNorm.includes("열") &&
    (tNorm.includes("추가") || tNorm.includes("삽입"))
  ) {
    const colInfo = detectColumnInfo(originalText);
    let position = "right";
    if (tNorm.includes("왼쪽") || tNorm.includes("좌측")) position = "left";
    if (tNorm.includes("오른쪽") || tNorm.includes("우측")) position = "right";

    return {
      type: "insertColumn",
      column: colInfo,
      position,
      text: originalText,
    };
  }

  // ─────────────────────────────
  // 8) 범위 전체 삭제(내용 지우기) – 위 케이스 아닌 삭제/지우기
  // ─────────────────────────────
  if (hasDeleteKeyword) {
    const range = detectRange(originalText) || "A1:A10";
    return {
      type: "clearRange",
      target: { range },
      text: originalText,
    };
  }

  // ─────────────────────────────
  // 9) 시트 관련 (생성/복사/이름 변경/삭제/이동)
  // ─────────────────────────────
  const hasSheetKeyword = tNorm.includes("시트");

  if (hasSheetKeyword) {
    const quotedName = detectSheetNameInQuotes(originalText);
    const looseName = detectSheetNameLoose(originalText);
    const anyName = quotedName || looseName;

    // 9-1) 시트 생성
    if (tNorm.includes("생성") || tNorm.includes("만들")) {
      const name = quotedName || "NewSheet";
      return {
        type: "createSheet",
        name,
        text: originalText,
      };
    }

    // 9-2) 시트 복사
    if (tNorm.includes("복사")) {
      const name = quotedName || "Backup";
      return {
        type: "duplicateSheet",
        name,
        text: originalText,
      };
    }

    // 9-3) 시트 이름 변경
    // 패턴: 현재 시트 이름을 "요약"으로 바꿔줘
    // 또는: "데이터" 시트 이름을 "원본"으로 변경해줘
    if (
      tNorm.includes("이름") &&
      (tNorm.includes("변경") || tNorm.includes("바꿔"))
    ) {
      const names = originalText.match(/["“”‘’']([^"“”‘’']+)["“”‘’']/g);
      let fromName = null;
      let toName = null;

      if (names && names.length >= 2) {
        fromName = names[0].replace(/["“”‘’']/g, "").trim();
        toName = names[1].replace(/["“”‘’']/g, "").trim();
      } else {
        // 따옴표 없는 경우: anyName 을 새 이름으로 보고, 대상은 현재 시트
        toName = anyName || "RenamedSheet";
      }

      return {
        type: "renameSheet",
        fromName,
        toName,
        text: originalText,
      };
    }

    // 9-4) 시트 삭제
    if (hasDeleteKeyword) {
      const name = anyName || quotedName || null; // 없으면 나중에 active 삭제도 가능하지만, 안전 위해 null 허용
      return {
        type: "deleteSheet",
        name,
        text: originalText,
      };
    }

    // 9-5) 시트 이동/활성화
    if (
      tNorm.includes("이동") ||
      tNorm.includes("전환") ||
      tNorm.includes("가줘") ||
      tNorm.includes("열어") ||
      tNorm.includes("선택")
    ) {
      const name = anyName || quotedName || "Sheet1";
      return {
        type: "activateSheet",
        name,
        text: originalText,
      };
    }
  }

  // ─────────────────────────────
  // 10) 그 외
  // ─────────────────────────────
  return { type: "unknown", text: originalText };
}

module.exports = {
  parseMacroIntent,
};
