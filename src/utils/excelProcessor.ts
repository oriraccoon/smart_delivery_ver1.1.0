import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { PlatformConfig, ProductMapping, ProcessedOrder, MatchingError, ColumnMapping, ValidationIssue, CourierConfig } from '../types';
import { DEFAULT_COURIERS } from './defaultData';

// ==========================================
// 1. 텍스트 정제 및 정규화 유틸리티
// ==========================================

export function colNumToLetter(colNum: number): string {
  let temp = colNum;
  let letter = '';
  while (temp > 0) {
    const mod = (temp - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    temp = Math.floor((temp - mod) / 26);
  }
  return letter || 'A';
}

export function normalizeTextForMapping(text: string): string {
  if (!text) return "";
  let raw = String(text);
  // 대괄호 및 소괄호 내용 제거
  raw = raw.replace(/\[.*?\]|\(.*?\)/g, '');
  // 숫자 + 용량/수량 단위 제거 (대소문자 구분 없음)
  raw = raw.replace(/\d+\.?\d*\s*(kg|g|k|개|세트|팩|입|통|봉)/gi, '');
  // 문자와 숫자 이외의 특수문자 및 공백 제거
  raw = raw.replace(/[^\wㄱ-ㅎㅏ-ㅣ가-힣]/g, '');
  return raw.trim();
}

export function normalizeValueForMatching(val: any): string {
  if (val === null || val === undefined) return "";
  let s = String(val).trim();
  if (s.endsWith('.0')) {
    s = s.slice(0, -2);
  }
  return s;
}

export function cleanStr(val: any): string {
  if (val === null || val === undefined) return "";
  return String(val).trim();
}

export function cleanZip(val: any): string {
  if (val === null || val === undefined) return "";
  return String(val).replace(/[^\d]/g, '').trim();
}

// 상품 정보에서 중량(Unit)과 배수(Multiplier)를 완벽히 정밀 추출하는 스마트 가공 함수
export function extractWeightAndMultiplier(
  rawProdName: string,
  originalUnitText: string,
  mappedName: string
): { unitWeight: string; multiplier: number } {
  const combinedText = `${rawProdName} ${originalUnitText}`.trim();
  
  // 1. 용량/중량(Weight) 파싱
  // 주요 타겟 용량 패턴: kg, g, ml, k, l 등
  const weightRegex = /(\d+(?:\.\d+)?)\s*(kg|g|ml|k|l)/i;
  const weightMatches = combinedText.match(weightRegex);
  
  let unitWeight = "1kg"; // 기본 디폴트
  if (weightMatches) {
    const val = weightMatches[1];
    const unit = weightMatches[2].toLowerCase();
    
    let standardUnit = unit;
    if (unit === 'k') standardUnit = 'kg';
    if (unit === 'l') standardUnit = 'L';
    
    unitWeight = `${val}${standardUnit}`;
  } else {
    // 텍스트에 명시적인 중량이 없다면 품목명(mappedName)별 기본 중량을 스마트하게 매핑해 줍니다.
    const defaultWeights: { [key: string]: string } = {
      "간장연탄": "200g",
      "고추장연탄": "200g",
      "육전": "150g",
      "진미채": "150g",
      "우엉": "200g",
      "도라지": "300g",
      "아삭이고추": "500g",
      "순태": "500g",
      "양념조개": "500g",
      "연근": "700g",
      "통마늘": "2kg",
      "반달단무지": "2.68kg",
      "실채": "150g",
      "더덕": "200g",
      "고추채": "10kg",
      "락교": "1kg",
      "알밥단무지": "1kg",
      "대두조림": "1kg",
      "서리태조림": "1kg",
      "곡물강낭콩": "1kg",
      "곡물서리태": "1kg",
      "단풍콩잎": "1kg",
      "된장고추": "1kg",
      "땅콩": "1kg",
      "땅콩진미": "1kg",
      "양념고추": "1kg",
      "양념깻잎": "1kg",
      "양념창란": "1kg",
      "양대조림": "1kg",
      "여러콩": "1kg",
      "오이피클": "1kg",
      "오징어": "10kg",
      "게장": "1kg",
      "무말랭이": "1kg"
    };
    
    if (defaultWeights[mappedName]) {
      unitWeight = defaultWeights[mappedName];
    }
  }

  // 2. 배수(Multiplier) 파싱
  // 중량 단위(예: "1kg")는 배수로 혼동되지 않도록 먼저 공백으로 소거합니다.
  let textForQty = combinedText;
  if (weightMatches) {
    textForQty = textForQty.replace(weightMatches[0], " ");
  }

  let multiplier = 1;

  // 수량 배수 정규식: "숫자 + (개|팩|봉|입|통|세트|묶음|박스|p)"
  const qtyRegex = /(\d+)\s*(개|팩|봉|입|통|세트|묶음|박스|p)/gi;
  let qtyMatch;
  const qtyMatches: number[] = [];
  
  while ((qtyMatch = qtyRegex.exec(textForQty)) !== null) {
    const num = parseInt(qtyMatch[1], 10);
    if (!isNaN(num) && num > 0) {
      qtyMatches.push(num);
    }
  }

  if (qtyMatches.length > 0) {
    multiplier = qtyMatches[0];
  } else {
    // 텍스트 매칭 보완
    if (combinedText.includes("5팩")) {
      multiplier = 5;
    } else if (combinedText.includes("10팩")) {
      multiplier = 10;
    } else if (combinedText.includes("2개")) {
      multiplier = 2;
    }
  }

  return {
    unitWeight,
    multiplier
  };
}

// 상품명에서 단위(kg, g, 개, 팩 등), 수량, 괄호 수식어를 제거하여 순수 키값으로 단순화하는 함수
export function simplifyRawProductName(rawProdName: string): string {
  if (!rawProdName) return "";
  let name = String(rawProdName).trim();
  
  // 1. 대괄호 및 소괄호 지우기
  name = name.replace(/\[.*?\]|\(.*?\)/g, ' ');
  
  // 2. 용량/수량 패턴 지우기 (예: 1kg, 500g, 2개, 3팩, 10봉, 1박스, 2세트 등)
  const optionRegex = /\d+(?:\.\d+)?\s*(?:kg|g|k|L|ml|봉|팩|입|통|세트|개|p|짜리|묶음|박스)/gi;
  name = name.replace(optionRegex, ' ');
  
  // 3. 남은 기호 및 공백 정제
  name = name.replace(/^[\s\-,_/]+|[\s\-,_/]+$/g, '').trim();
  name = name.replace(/\s+/g, ' ');
  
  return name || rawProdName.trim();
}

// 상품명에서 옵션(중량, 개수 등)을 스마트하게 발라내고 매핑 키로 간소화하는 함수
export function extractOptionAndCleanName(
  rawProdName: string,
  productMap: ProductMapping[],
  originalUnit?: string
): { cleanName: string; mappedName: string; optionText: string; matchedRawName: string; isMatched: boolean } {
  const cleanName = rawProdName.trim();
  let optionText = String(originalUnit || "").trim();
  let mappedName = "";

  // 1단계: 매핑 사전의 rawName 중 가장 긴 것부터 체크하여 포함(substring) 관계 확인
  const sortedMap = [...productMap].sort((a, b) => b.rawName.length - a.rawName.length);
  let foundMapping: ProductMapping | null = null;

  for (const mapping of sortedMap) {
    if (cleanName.includes(mapping.rawName)) {
      foundMapping = mapping;
      break;
    }
  }

  if (foundMapping) {
    mappedName = foundMapping.mappedName;
    // 매핑 키를 지운 나머지 텍스트를 옵션(수량/용량 등)으로 추출
    let remaining = cleanName.replace(foundMapping.rawName, "").trim();
    remaining = remaining.replace(/^[\s\-,_/]+|[\s\-,_/]+$/g, "").trim();
    
    if (remaining) {
      if (optionText) {
        if (!optionText.includes(remaining)) {
          optionText = `${optionText} ${remaining}`;
        }
      } else {
        optionText = remaining;
      }
    }

    return {
      cleanName: foundMapping.rawName,
      mappedName: mappedName,
      optionText: optionText || "기본",
      matchedRawName: foundMapping.rawName,
      isMatched: true
    };
  }

  // 2단계: 매핑 테이블에 없는 경우, 정규식을 사용하여 용량/수량/세트 옵션 패턴을 먼저 추출
  const optionRegex = /(\d+(?:\.\d+)?\s*(?:kg|g|k|L|ml|봉|팩|입|통|세트|개|p|짜리|묶음|박스))\s*(\(.*\))?|\((?:\d+(?:\.\d+)?\s*(?:kg|g|k|L|ml|봉|팩|입|통|세트|개|p))\)|\[(?:\d+(?:\.\d+)?\s*(?:kg|g|k|L|ml|봉|팩|입|통|세트|개|p))\]/gi;
  
  const matches: string[] = [];
  let match;
  const regex = new RegExp(optionRegex);
  while ((match = regex.exec(cleanName)) !== null) {
    matches.push(match[0]);
  }

  if (matches.length > 0) {
    const foundOption = matches.join(" ").trim();
    if (optionText) {
      if (!optionText.includes(foundOption)) {
        optionText = `${optionText} ${foundOption}`;
      }
    } else {
      optionText = foundOption;
    }

    // 옵션 패턴들을 제거한 이름 만들기
    let cleaned = cleanName;
    matches.forEach(m => {
      cleaned = cleaned.replace(m, "");
    });
    cleaned = cleaned.replace(/^[\s\-,_/]+|[\s\-,_/]+$/g, "").trim();
    
    // 제거하고 남은 이름으로 다시 1단계 매칭(포함 관계)을 시도합니다.
    for (const mapping of sortedMap) {
      if (cleaned.includes(mapping.rawName) || mapping.rawName.includes(cleaned)) {
        return {
          cleanName: mapping.rawName,
          mappedName: mapping.mappedName,
          optionText: optionText || "기본",
          matchedRawName: mapping.rawName,
          isMatched: true
        };
      }
    }

    // 텍스트 매칭 정규화로 한번 더 시도
    const normClean = normalizeTextForMapping(cleaned);
    for (const mapping of sortedMap) {
      const normKey = normalizeTextForMapping(mapping.rawName);
      if (normKey && (normClean === normKey || normClean.includes(normKey) || normKey.includes(normClean))) {
        return {
          cleanName: mapping.rawName,
          mappedName: mapping.mappedName,
          optionText: optionText || "기본",
          matchedRawName: mapping.rawName,
          isMatched: true
        };
      }
    }

    const simplifiedCleaned = simplifyRawProductName(cleaned || cleanName);

    return {
      cleanName: simplifiedCleaned,
      mappedName: simplifiedCleaned,
      optionText: optionText || "기본",
      matchedRawName: "",
      isMatched: false
    };
  }

  // 3단계: 여전히 매칭되지 않는다면 기존 매핑 전처리를 활용해 매칭 시도
  const normClean = normalizeTextForMapping(cleanName);
  for (const mapping of sortedMap) {
    const normKey = normalizeTextForMapping(mapping.rawName);
    if (normKey && (normClean === normKey || normClean.includes(normKey) || normKey.includes(normClean))) {
      return {
        cleanName: mapping.rawName,
        mappedName: mapping.mappedName,
        optionText: optionText || "기본",
        matchedRawName: mapping.rawName,
        isMatched: true
      };
    }
  }

  const simplifiedCleanName = simplifyRawProductName(cleanName);

  return {
    cleanName: simplifiedCleanName,
    mappedName: simplifiedCleanName,
    optionText: optionText || "기본",
    matchedRawName: "",
    isMatched: false
  };
}

// 품목 정제 (사전에 정의된 대치 단어 매핑 적용)
export function cleanProductName(text: string, productMap: ProductMapping[]): string {
  const result = extractOptionAndCleanName(text, productMap);
  return result.mappedName;
}

// 가이드 행(수정가능/수정불가 설명글이 포함된 행)인지 확인하는 유틸리티
export function isGuideRow(row: any[]): boolean {
  if (!row || row.length === 0) return false;
  
  const guideKeywords = [
    "수정불가", "수정가능", "수정 불가", "수정 가능", 
    "수정금지", "수정 금지", "입력불가", "입력 불가"
  ];
  
  // 배송메시지 등 고객이 입력하는 뒷쪽 열에서 우연히 매칭되어 주문이 유실되는 것을 방지하기 위해,
  // 앞쪽 10개 열까지만 검사하여 가이드 행 여부를 판별합니다.
  const limit = Math.min(10, row.length);
  for (let c = 0; c < limit; c++) {
    const cellStr = String(row[c] || "").trim().toLowerCase().replace(/\s+/g, "");
    if (!cellStr) continue;
    
    for (const kw of guideKeywords) {
      const cleanKw = kw.toLowerCase().replace(/\s+/g, "");
      if (cellStr === cleanKw || cellStr.includes(cleanKw)) {
        return true;
      }
    }
  }
  return false;
}

// 신규 품목 추출 유틸리티 (대치 단어 사전에 없는 신규 품목 발굴)
export function extractRawUniqueNames(
  rows: any[][], 
  colMapIndex: number, 
  startRow: number
): string[] {
  const uniqueNames = new Set<string>();
  for (let i = startRow; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    if (isGuideRow(row)) continue; // 가이드 행 건너뛰기

    const val = row[colMapIndex];
    if (val !== undefined && val !== null && String(val).trim() !== "") {
      uniqueNames.add(String(val).trim());
    }
  }
  return Array.from(uniqueNames);
}

// 헤더 셀 명칭을 기반으로 유연하게 열 인덱스를 탐색 및 보정하는 스마트 함수
export function resolveColumnMapping(headerRow: any[], defaultColMap: ColumnMapping): ColumnMapping {
  const colMap = { ...defaultColMap };
  if (!headerRow || headerRow.length === 0) return colMap;

  const headers = headerRow.map(h => String(h || "").trim().toLowerCase().replace(/\s+/g, ""));

  const findColumnIndex = (keywords: string[]): number => {
    for (const kw of keywords) {
      const cleanKw = kw.toLowerCase().replace(/\s+/g, "");
      const exactIdx = headers.findIndex(h => h === cleanKw);
      if (exactIdx !== -1) return exactIdx;
    }
    for (const kw of keywords) {
      const cleanKw = kw.toLowerCase().replace(/\s+/g, "");
      const partialIdx = headers.findIndex(h => h.includes(cleanKw));
      if (partialIdx !== -1) return partialIdx;
    }
    return -1;
  };

  const mappings: { [key in keyof ColumnMapping]: string[] } = {
    상품명: ["노출상품명", "상품명", "상품명(노출상품명)", "등록상품명"],
    용량: ["등록옵션명", "옵션명", "옵션정보", "상세옵션", "옵션", "상품상세"],
    수량: ["구매수", "수량", "구매수량", "수량(고정)"],
    수령인: ["수취인명", "수령인명", "수령인", "받는사람", "수취인"],
    연락처: ["수취인이동전화번호", "수취인연락처", "연락처", "전화번호", "휴대폰번호", "수취인전화번호", "수취인 연락처1"],
    우편번호: ["우편번호", "수취인우편번호", "배송지우편번호"],
    주소: ["수취인주소", "배송지", "주소", "배송지주소", "수취인 주소"],
    배송메시지: ["배송메세지", "배송메시지", "배송요청사항", "배송메세지1"]
  };

  const keys: Array<keyof ColumnMapping> = [
    "상품명",
    "용량",
    "수량",
    "수령인",
    "연락처",
    "우편번호",
    "주소",
    "배송메시지"
  ];

  keys.forEach(key => {
    const matchedIdx = findColumnIndex(mappings[key]);
    if (matchedIdx !== -1) {
      colMap[key] = matchedIdx;
    }
  });

  return colMap;
}

// 여러 시트 중 실데이터가 존재하고 주문 관련 정보가 가장 많은 최적의 시트 이름을 찾는 스마트 유틸리티
export function findBestOrderSheet(workbook: XLSX.WorkBook): string {
  if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) {
    return "";
  }
  if (workbook.SheetNames.length === 1) {
    return workbook.SheetNames[0];
  }

  let bestSheetName = workbook.SheetNames[0];
  let maxScore = -1;

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json<any[]>(sheet, { header: 1, range: 0, defval: "" });
    let score = 0;
    
    const scanRows = Math.min(15, rows.length);
    for (let i = 0; i < scanRows; i++) {
      const rowStr = (rows[i] || []).map(v => String(v)).join(" ");
      if (rowStr.includes("주문번호")) score += 10;
      if (rowStr.includes("수취인명") || rowStr.includes("수령인") || rowStr.includes("수취인")) score += 8;
      if (rowStr.includes("상품명")) score += 5;
      if (rowStr.includes("옵션") || rowStr.includes("용량")) score += 3;
      if (rowStr.includes("묶음배송번호")) score += 10;
    }

    if (score > maxScore) {
      maxScore = score;
      bestSheetName = sheetName;
    }
  }

  return bestSheetName;
}

// 여러 행 중에서 실제 컬럼 헤더들이 나열되어 있는 최적의 헤더 행 인덱스를 찾는 스마트 함수
export function findHeaderRowIndex(rawRows: any[][], defaultStartRow: number): number {
  if (rawRows.length === 0) return 0;

  let bestRowIdx = Math.max(0, defaultStartRow - 1);
  let maxScore = -1;

  // 헤더 매칭용 키워드 정의
  const keywords = [
    "주문번호", "상품명", "노출상품명", "등록상품명", "옵션", "옵션정보", "상세옵션",
    "수취인명", "수령인", "수령인명", "받는사람", "수취인",
    "연락처", "전화번호", "휴대폰번호", "우편번호", "주소", "배송지", "배송지주소",
    "구매수", "수량", "구매수량"
  ];

  const scanLimit = Math.min(15, rawRows.length);
  for (let r = 0; r < scanLimit; r++) {
    const row = rawRows[r] || [];
    let score = 0;

    row.forEach(cell => {
      const cellStr = String(cell || "").trim().toLowerCase().replace(/\s+/g, "");
      if (!cellStr) return;

      for (const kw of keywords) {
        const cleanKw = kw.toLowerCase().replace(/\s+/g, "");
        if (cellStr === cleanKw) {
          // 완벽히 일치하는 단어가 헤더 행에 존재할 때 높은 점수 부여
          score += 10;
        } else if (cellStr.includes(cleanKw) && cellStr.length <= cleanKw.length + 6) {
          // 일치하진 않지만 키워드가 포함되어 있고 셀 길이가 키워드와 비슷할 때만 약간의 점수 부여
          score += 1;
        }
      }
    });

    if (score > maxScore) {
      maxScore = score;
      bestRowIdx = r;
    }
  }

  // 최적의 헤더 행의 매칭 점수가 너무 낮으면 기본 설정값을 사용
  if (maxScore < 8) {
    return Math.max(0, defaultStartRow - 1);
  }

  return bestRowIdx;
}

// ==========================================
// 2. 엑셀 파일 파싱 및 가공 (1단계)
// ==========================================

export interface FileParseResult {
  detectedPlatform: PlatformConfig | null;
  orders: ProcessedOrder[];
  rawRows: any[][];
  headers: any[];
  newRawItems: string[]; // 새로 발견되어 매핑이 필요한 원본 상품명들
  validationIssues: ValidationIssue[]; // 필수 항목 누락/오류 검사 결과 목록
}

// 2-1. 복호화되거나 이미 읽어온 2차원 배열 데이터를 기반으로 플랫폼 자동 감지 및 가공 분석을 실행하는 핵심 로직
export function processRawRows(
  rawRows: any[][],
  fileName: string,
  platforms: PlatformConfig[],
  productMap: ProductMapping[],
  forcedPlatformId?: string
): FileParseResult {
  if (rawRows.length === 0) {
    throw new Error("엑셀 파일에 데이터가 없습니다.");
  }

  // 플랫폼 자동 인식 또는 강제 지정
  let detectedPlatform: PlatformConfig | null = null;

  if (forcedPlatformId) {
    detectedPlatform = platforms.find(p => p.id === forcedPlatformId) || null;
  }

  if (!detectedPlatform) {
    // 1단계: 파일 이름 기반 매칭 시도 (Mac/Windows 한글 NFD/NFC 정규화)
    const normFileName = fileName.normalize('NFC').toLowerCase();
    
    for (const plat of platforms) {
      const platNameLower = plat.name.normalize('NFC').toLowerCase();
      
      const rawPatterns = plat.filepath_pattern
        ? plat.filepath_pattern.split(',').map(k => k.trim()).filter(k => k !== '')
        : [];
        
      let isMatched = false;
      
      if (rawPatterns.length > 0) {
        for (const pattern of rawPatterns) {
          // 패턴 정제: input/ 경로, 확장자, [오늘날짜] 등 제거
          const cleanedPattern = pattern
            .normalize('NFC')
            .replace(/^input\//i, '')
            .replace(/\.xlsx$/i, '')
            .replace(/\.xls$/i, '')
            .replace(/\[.*?\]/g, '')
            .replace(/\(.*?\)/g, '')
            .toLowerCase();

          if (!cleanedPattern) continue;

          // 단일 통째로 포함 검사 (예: "주문배송관리")
          const pureString = cleanedPattern.replace(/[*_\-\s]+/g, '');
          const pureFileName = normFileName.replace(/[*_\-\s]+/g, '');
          if (pureString.length >= 2 && pureFileName.includes(pureString)) {
            isMatched = true;
            break;
          }

          // 와일드카드(*) 또는 구분자(-, _)로 나눈 핵심 토큰들 추출 (예: ["주문배송관리", "상품준비중"])
          const tokens = cleanedPattern
            .split(/[*_\-\s]+/)
            .map(t => t.trim())
            .filter(t => t.length >= 2);

          if (tokens.length > 0) {
            // 모든 핵심 토큰이 파일명에 포함되어 있으면 매칭 성공
            const allTokensPresent = tokens.every(token => normFileName.includes(token));
            if (allTokensPresent) {
              isMatched = true;
              break;
            }
          }
        }
      }
      
      if (!isMatched) {
        if (
          normFileName.includes(platNameLower) ||
          (platNameLower.includes("네이버") && (normFileName.includes("naver") || normFileName.includes("스마트스토어") || normFileName.includes("smartstore") || normFileName.includes("발주발송"))) ||
          (platNameLower.includes("쿠팡") && (normFileName.includes("coupang") || normFileName.includes("deliverylist") || normFileName.includes("wing") || normFileName.includes("쿠팡"))) ||
          (platNameLower.includes("토스") && (normFileName.includes("toss") || normFileName.includes("주문배송관리") || normFileName.includes("상품준비중") || normFileName.includes("주문배송") || normFileName.includes("토스"))) ||
          (platNameLower.includes("g마켓") && (normFileName.includes("gmarket") || normFileName.includes("esm"))) ||
          (platNameLower.includes("옥션") && (normFileName.includes("auction") || normFileName.includes("esm"))) ||
          (platNameLower.includes("11번가") && (normFileName.includes("11st") || normFileName.includes("11번가")))
        ) {
          isMatched = true;
        }
      }
      
      if (isMatched) {
        detectedPlatform = plat;
        break;
      }
    }
  }

  // 2단계: 파일 이름 매칭이 안 된 경우, 시트 셀 안의 식별자 단어로 매칭 시도 (식별자가 지정된 경우)
  if (!detectedPlatform) {
    const searchLimit = Math.min(5, rawRows.length);
    for (let r = 0; r < searchLimit; r++) {
      const rowStr = rawRows[r].map(v => String(v)).join(" ").normalize('NFC').toLowerCase();
      for (const plat of platforms) {
        if (plat.identifier && plat.identifier.trim() !== '') {
          const normIdent = plat.identifier.normalize('NFC').toLowerCase();
          if (rowStr.includes(normIdent)) {
            detectedPlatform = plat;
            break;
          }
        }
      }
      if (detectedPlatform) break;
    }
  }

  // 감지되지 않았으면 기본적으로 첫 번째 플랫폼 또는 null
  const platform = detectedPlatform || platforms[0];
  
  // 헤더 행의 인덱스를 본문 검색을 통해 동적으로 완벽 보정
  const headerRowIndex = findHeaderRowIndex(rawRows, platform.start_row);

  const startRow = headerRowIndex + 1;
  const headerRow = rawRows[headerRowIndex] || [];
  const colMap = resolveColumnMapping(headerRow, platform.col_map);

  // 원본 상품명 추출 및 신규 품목 정제 목록 구축 (단위, 수량, 괄호 수식어가 제거된 단순 키값 추출)
  const rawNames = extractRawUniqueNames(rawRows, colMap.상품명, startRow);
  const newRawItemsSet = new Set<string>();

  rawNames.forEach(rawName => {
    const { isMatched, cleanName } = extractOptionAndCleanName(rawName, productMap);
    if (!isMatched) {
      const simplifiedKey = simplifyRawProductName(cleanName || rawName);
      if (simplifiedKey) {
        newRawItemsSet.add(simplifiedKey);
      }
    }
  });
  const newRawItems = Array.from(newRawItemsSet);

  // 주문 데이터 추출 및 필수 유효성 검사
  const orders: ProcessedOrder[] = [];
  const validationIssues: ValidationIssue[] = [];

  for (let i = startRow; i < rawRows.length; i++) {
    const row = rawRows[i];
    if (!row) continue;
    if (isGuideRow(row)) continue; // 가이드 행 건너뛰기

    const rawRecipient = String(row[colMap.수령인] || "").trim();
    const rawProdName = String(row[colMap.상품명] || "").trim();
    const rawAddr = String(row[colMap.주소] || "").trim();
    const rawPhone = String(row[colMap.연락처] || "").trim();
    const rawZip = normalizeValueForMatching(row[colMap.우편번호]);
    const rawQtyVal = row[colMap.수량];
    const qtyNum = Number(rawQtyVal) || 0;

    // 완전히 빈 행(모든 주요 데이터가 비어있는 경우) 건너뛰기
    const isCompletelyEmpty = (!rawRecipient || rawRecipient === "nan" || rawRecipient === "undefined") &&
                              !rawProdName && !rawAddr && !rawPhone;
    if (isCompletelyEmpty) continue;

    // 필수 항목 검사
    const missingFields: string[] = [];

    if (!rawRecipient || rawRecipient === "nan" || rawRecipient === "undefined" || rawRecipient === "-") {
      missingFields.push("수령인");
    }

    if (!rawAddr || rawAddr === "nan" || rawAddr === "undefined" || rawAddr === "-") {
      missingFields.push("주소");
    }

    if (!rawPhone || rawPhone === "nan" || rawPhone === "undefined" || rawPhone === "-") {
      missingFields.push("연락처");
    }

    if (!rawProdName || rawProdName === "nan" || rawProdName === "undefined") {
      missingFields.push("상품명");
    }

    if (qtyNum <= 0) {
      missingFields.push("수량");
    }

    if (!rawZip || rawZip === "nan" || rawZip === "undefined" || rawZip === "-") {
      missingFields.push("우편번호");
    }

    const recipientName = rawRecipient && rawRecipient !== "nan" && rawRecipient !== "undefined" ? rawRecipient : "(수령인 미입력)";
    const originalUnit = String(row[colMap.용량] || "");
    
    const { mappedName } = extractOptionAndCleanName(rawProdName, productMap, originalUnit);
    const { unitWeight, multiplier } = extractWeightAndMultiplier(rawProdName, originalUnit, mappedName);
    
    const zipCode = rawZip;
    const addr = rawAddr;
    const ordererName = recipientName;

    const uniqueIdName = recipientName;
    const combinedId = `${uniqueIdName}_${zipCode}_${addr}`;

    const excelRowIndex = i + 1; // 엑셀 실제 1-based 행 번호

    if (missingFields.length > 0) {
      validationIssues.push({
        id: `val_${i}_${Date.now()}`,
        rowIndex: excelRowIndex,
        fileName: fileName,
        recipient: recipientName,
        productName: rawProdName || mappedName || "(상품명 미입력)",
        address: addr || "(주소 미입력)",
        phone: rawPhone || "(연락처 미입력)",
        zipCode: zipCode || "(우편번호 미입력)",
        missingFields: missingFields
      });
    }

    orders.push({
      id: String(i),
      용량: unitWeight,
      상품명: mappedName,
      수량: qtyNum,
      주문자: ordererName,
      수령인: recipientName,
      연락처: rawPhone,
      우편번호: zipCode,
      주소: addr,
      배송메시지: String(row[colMap.배송메시지] || ""),
      sortKey: mappedName,
      combinedId: combinedId,
      isMulti: false,
      multiplier: multiplier,
      originalOptionName: originalUnit,
      originalProductName: rawProdName,
      validationIssues: missingFields.length > 0 ? missingFields : undefined
    });
  }

  const idCounts: { [key: string]: number } = {};
  orders.forEach(o => {
    idCounts[o.combinedId] = (idCounts[o.combinedId] || 0) + 1;
  });

  orders.forEach(o => {
    o.isMulti = idCounts[o.combinedId] > 1;
  });

  const multiOrders = orders.filter(o => o.isMulti).sort((a, b) => {
    if (a.combinedId !== b.combinedId) return a.combinedId.localeCompare(b.combinedId);
    if (a.sortKey !== b.sortKey) return a.sortKey.localeCompare(b.sortKey);
    return a.상품명.localeCompare(b.상품명);
  });

  const singleOrders = orders.filter(o => !o.isMulti).sort((a, b) => {
    if (a.sortKey !== b.sortKey) return a.sortKey.localeCompare(b.sortKey);
    if (a.상품명 !== b.상품명) return a.상품명.localeCompare(b.상품명);
    return a.수령인.localeCompare(b.수령인);
  });

  const sortedOrders = [...multiOrders, ...singleOrders];

  return {
    detectedPlatform: detectedPlatform,
    orders: sortedOrders,
    rawRows,
    headers: headerRow,
    newRawItems,
    validationIssues
  };
}

export async function parseAndProcessOrderExcel(
  file: File,
  platforms: PlatformConfig[],
  productMap: ProductMapping[],
  forcedPlatformId?: string
): Promise<FileParseResult> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const bestSheetName = findBestOrderSheet(workbook);
        const worksheet = workbook.Sheets[bestSheetName];
        
        const rawRows = XLSX.utils.sheet_to_json<any[]>(worksheet, { header: 1, defval: "" });
        const result = processRawRows(rawRows, file.name, platforms, productMap, forcedPlatformId);
        resolve(result);
      } catch (err: any) {
        const errMsg = err?.message || "";
        const lowerMsg = errMsg.toLowerCase();
        if (
          lowerMsg.includes("password") ||
          lowerMsg.includes("protected") ||
          lowerMsg.includes("decrypt") ||
          lowerMsg.includes("encrypted")
        ) {
          reject(new Error("PASSWORD_PROTECTED"));
        } else {
          reject(err);
        }
      }
    };
    reader.onerror = () => reject(new Error("파일을 읽는 도중 오류가 발생했습니다."));
    reader.readAsArrayBuffer(file);
  });
}

// ==========================================
// 3. 정렬 파일 및 요약 보고서 저장 (ExcelJS 활용)
// ==========================================

export function sortUnitsSmart(units: string[]): string[] {
  const getGroupAndValue = (unitStr: string) => {
    const s = unitStr.trim().toLowerCase();
    
    // 1. kg 단위 체크 (숫자 + kg)
    const kgMatch = s.match(/^(\d+(?:\.\d+)?)\s*(?:kg|k)$/i);
    if (kgMatch) {
      return { group: 0, val: parseFloat(kgMatch[1]), raw: unitStr };
    }
    
    // 2. g/ml 단위 체크 (숫자 + g 또는 ml)
    const gMatch = s.match(/^(\d+(?:\.\d+)?)\s*(?:g|ml)$/i);
    if (gMatch) {
      return { group: 1, val: parseFloat(gMatch[1]), raw: unitStr };
    }
    
    // 3. 기타 수량 단위 체크 (숫자 + 개/팩/세트 등)
    const numMatch = s.match(/^(\d+(?:\.\d+)?)\s*(.*)$/);
    if (numMatch) {
      return { group: 2, val: parseFloat(numMatch[1]), suffix: numMatch[2], raw: unitStr };
    }
    
    // 4. 비수량 텍스트 (기본 등)
    return { group: 3, val: 0, suffix: s, raw: unitStr };
  };

  const parsed = units.map(getGroupAndValue);

  parsed.sort((a, b) => {
    if (a.group !== b.group) {
      return a.group - b.group;
    }
    if (a.group === 0 || a.group === 1) {
      return a.val - b.val;
    }
    if (a.group === 2) {
      if (a.val !== b.val) {
        return a.val - b.val;
      }
      return (a.suffix || '').localeCompare(b.suffix || '');
    }
    return (a.suffix || '').localeCompare(b.suffix || '');
  });

  return parsed.map(item => item.raw);
}

export async function exportSummaryReportExcel(orders: ProcessedOrder[]): Promise<void> {
  const wb = new ExcelJS.Workbook();
  
  // ----------------------------------------------------
  // 시트 1. 상세내역 (용량, 상품명, 수량, 주문자)
  // ----------------------------------------------------
  const wsDetail = wb.addWorksheet('상세내역');
  wsDetail.addRow(["용량", "상품명", "수량", "주문자"]);

  orders.forEach(o => {
    wsDetail.addRow([
      o.originalOptionName || "",
      o.originalProductName || "",
      o.수량,
      o.주문자 || ""
    ]);
  });

  styleSheetCommon(wsDetail, 4, true);

  // ----------------------------------------------------
  // 시트 2. 상품준비요약 (기존 피벗 구조 및 종횡 총합계 포함)
  // ----------------------------------------------------
  const wsSummary = wb.addWorksheet('상품준비요약');

  const unitSet = new Set<string>();
  const pivotMap: { [baseName: string]: { [unit: string]: number } } = {};

  orders.forEach(o => {
    const unit = o.용량 || "기본";
    const baseName = o.상품명 || "미분류";
    unitSet.add(unit);

    if (!pivotMap[baseName]) {
      pivotMap[baseName] = {};
    }
    pivotMap[baseName][unit] = (pivotMap[baseName][unit] || 0) + (o.수량 * (o.multiplier || 1));
  });

  const sortedUnits = sortUnitsSmart(Array.from(unitSet));
  const sortedBaseNames = Object.keys(pivotMap).sort();

  // 요약 헤더 작성: [상품 분류 (행 레이블), 옵션1, 옵션2, ..., 총합계]
  const summaryHeader = ["상품 분류 (행 레이블)", ...sortedUnits, "총합계"];
  wsSummary.addRow(summaryHeader);

  // 데이터 행 작성
  sortedBaseNames.forEach(baseName => {
    const rowData: any[] = [baseName];
    let totalRowQty = 0;
    sortedUnits.forEach(unit => {
      const qty = pivotMap[baseName][unit] || 0;
      rowData.push(qty > 0 ? qty : "");
      totalRowQty += qty;
    });
    rowData.push(totalRowQty);
    wsSummary.addRow(rowData);
  });

  // 하단 총합계 행 작성
  const totalRow: any[] = ["총합계"];
  let grandTotal = 0;
  sortedUnits.forEach(unit => {
    let colTotal = 0;
    sortedBaseNames.forEach(baseName => {
      colTotal += pivotMap[baseName][unit] || 0;
    });
    totalRow.push(colTotal > 0 ? colTotal : "");
    grandTotal += colTotal;
  });
  totalRow.push(grandTotal);
  wsSummary.addRow(totalRow);

  styleSheetCommon(wsSummary, sortedUnits.length + 2, false);

  // 다운로드 처리
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const today = new Date().toISOString().slice(0, 10);
  triggerDownload(blob, `요약보고서_${today}.xlsx`);
}

export async function exportProductSortedExcel(orders: ProcessedOrder[]): Promise<void> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('가공데이터');

  // 노출상품명과 구매수 사이에 수량(고정)이 들어간 총 9개 컬럼 구성
  const headerRow = [
    "용량",
    "상품명",
    "수량(고정)",
    "수량",
    "수령인",
    "연락처",
    "우편번호",
    "주소",
    "배송메시지"
  ];
  ws.addRow(headerRow);

  orders.forEach(o => {
    ws.addRow([
      o.originalOptionName || "",
      o.originalProductName || "",
      1, // 수량(고정) 고정값 1
      o.수량,
      o.수령인 || "",
      o.연락처 || "",
      o.우편번호 || "",
      o.주소 || "",
      o.배송메시지 || ""
    ]);
  });

  styleSheetCommon(ws, 9, true);

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  
  // 날짜 형식: YYYY_MM_DD
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  const todayStr = `${year}_${month}_${day}`;

  triggerDownload(blob, `품목별정렬_${todayStr}.xlsx`);
}

// 엑셀 서식 공통 적용 헬퍼 (파이썬 openpyxl 스타일 완벽 반영)
function styleSheetCommon(ws: ExcelJS.Worksheet, maxCol: number, isDetail: boolean = false) {
  // 폰트 및 스타일 상수 정의
  const font9 = { name: 'Inter', size: 9 };
  const font9Bold = { name: 'Inter', size: 9, bold: true };
  const headerFont = { name: 'Inter', size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
  
  const headerFill: ExcelJS.Fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF2F5597' }
  };
  
  const stripeFill: ExcelJS.Fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFF8F9FA' }
  };

  const thinBorder: any = {
    left: { style: 'thin', color: { argb: 'FFD9D9D9' } },
    right: { style: 'thin', color: { argb: 'FFD9D9D9' } },
    top: { style: 'thin', color: { argb: 'FFD9D9D9' } },
    bottom: { style: 'thin', color: { argb: 'FFD9D9D9' } }
  };

  const alignCenter: any = { horizontal: 'center', vertical: 'middle' };
  const alignLeft: any = { horizontal: 'left', vertical: 'middle', indent: 1 };

  // 1. 헤더 행(1번 행) 스타일
  const headerRow = ws.getRow(1);
  headerRow.height = 24;
  for (let c = 1; c <= maxCol; c++) {
    const cell = headerRow.getCell(c);
    cell.font = headerFont;
    cell.fill = headerFill;
    cell.alignment = alignCenter;
    cell.border = thinBorder;
  }

  // 2. 데이터 행 스타일
  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return; // 헤더 제외
    
    row.height = 20;
    const isEven = rowNumber % 2 === 0;

    for (let c = 1; c <= maxCol; c++) {
      const cell = row.getCell(c);
      cell.border = thinBorder;
      if (isEven) {
        cell.fill = stripeFill;
      }

      // 기본 폰트
      cell.font = font9;

      // 정렬 설정
      if (isDetail) {
        // 상세내역 시트: 상품명(2번 열)과 주문자(4번 열)는 왼쪽 정렬, 나머지는 가운데
        if (c === 2 || c === 4) {
          cell.alignment = alignLeft;
        } else {
          cell.alignment = alignCenter;
        }
      } else {
        // 일반 정렬 (텍스트는 왼쪽, 우편번호나 수량 등은 가운데)
        if (c === 1) {
          cell.alignment = alignLeft;
        } else {
          cell.alignment = alignCenter;
          // 요약 시트에서 수량이 0보다 큰 경우 볼드 처리
          if (cell.value && typeof cell.value === 'number') {
            cell.font = font9Bold;
          }
        }
      }
    }
  });

  // 3. 틀 고정 및 오토 필터
  ws.views = [{ state: 'frozen', ySplit: 1, xSplit: 0, activeCell: 'A2' }];
  ws.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: ws.actualRowCount, column: maxCol }
  };

  // 4. 셀 크기 자동 조정 (초밀착 정밀 너비 계산)
  ws.columns.forEach(col => {
    let maxLen = 0;
    col.eachCell!({ includeEmpty: false }, (cell) => {
      if (cell.value !== undefined && cell.value !== null) {
        const strVal = String(cell.value);
        // 한글은 대략 1.7배, 영숫자는 0.85배로 길이 산정
        let length = 0;
        for (let i = 0; i < strVal.length; i++) {
          length += strVal.charCodeAt(i) > 128 ? 1.7 : 0.85;
        }
        if (length > maxLen) maxLen = length;
      }
    });

    const padding = isDetail ? 2.5 : 1.8;
    col.width = Math.max(8, maxLen + padding);
  });
}

// ==========================================
// 4. 운송장 매칭 및 기입 프로세스 (2단계)
// ==========================================

export interface TrackingUpdateResult {
  matchCount: number;
  softMatchCount: number;
  bundledCount: number;      // 합포장 품목 건수
  bundledGroupCount: number; // 합포장 묶음(고객/송장) 그룹 수
  bundleMismatchCount?: number; // C열 합포장 수량 불일치 건수
  bundleMismatchDetails?: string[]; // C열 합포장 수량 불일치 내역
  unfilledErrors: MatchingError[];
  unusedErrors: MatchingError[];
  outputOrders: ProcessedOrder[];
}

// 헬퍼: 해당 주문/플랫폼이 쿠팡(DeliveryList)인지 판별하는 함수 (쿠팡만 합포장 연동 허용)
function isCoupangOrder(o: { platformName?: string; platformId?: string; sourceFileName?: string }): boolean {
  const plat = (o.platformName || o.platformId || "").toLowerCase();
  const file = (o.sourceFileName || "").toLowerCase();
  return (
    plat.includes("쿠팡") ||
    plat.includes("coupang") ||
    plat.includes("deliverylist") ||
    plat.includes("wing") ||
    file.includes("쿠팡") ||
    file.includes("coupang") ||
    file.includes("deliverylist")
  );
}

export function isSoftMatch(inKey: string[], srcKey: string[]): boolean {
  const [inN, inU, inP, inQ, inZ] = inKey;
  const [srcN, srcU, srcP, srcQ, srcZ] = srcKey;

  // 우편번호와 수량은 무조건 같아야 함
  if (inZ !== srcZ || inQ !== srcQ) return false;

  // 이름(2글자 이상) 및 품목명(3글자 이상) 포함 관계 확인
  const nameMatch = (inN.includes(srcN) || srcN.includes(inN)) && Math.min(inN.length, srcN.length) >= 2;
  const prodMatch = (inP.includes(srcP) || srcP.includes(inP)) && Math.min(inP.length, srcP.length) >= 3;

  return nameMatch && prodMatch;
}

// 헬퍼: 공통 최장 부분자열(LCS) 길이 계산 (수령인 이름 유사도 계산용)
function getLCSLength(s1: string, s2: string): number {
  const m = s1.length;
  const n = s2.length;
  if (m === 0 || n === 0) return 0;
  const dp = Array.from({ length: m + 1 }, () => new Array(n + 1).fill(0));
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      if (s1[i - 1] === s2[j - 1]) {
        dp[i][j] = dp[i - 1][j - 1] + 1;
      } else {
        dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
      }
    }
  }
  return dp[m][n];
}

// 헬퍼: 수령인(성명) 교차 일치 검사 (공백 제거 완전일치, 부분 포함, 긴 성명 80% 이상 유사 매칭)
function checkRecipientNameMatch(name1: string, name2: string): { isMatch: boolean; isExact: boolean; reason?: string } {
  const c1 = cleanStr(name1);
  const c2 = cleanStr(name2);
  if (!c1 || !c2) return { isMatch: false, isExact: false };

  // 1) 공백 및 특수문자 제거 후 100% 동일
  if (c1 === c2) {
    return { isMatch: true, isExact: true };
  }

  // 2) 한쪽이 다른 쪽에 완벽히 포함되는 경우 (예: "홍길동(네이버)" vs "홍길동", 10~14자 길어서 잘린 수령인명)
  if (c1.includes(c2) || c2.includes(c1)) {
    const minLen = Math.min(c1.length, c2.length);
    if (minLen >= 2) {
      return { isMatch: true, isExact: false, reason: `수령인 부분 포함 일치 ('${c1}' ↔ '${c2}')` };
    }
  }

  // 3) 긴 성명(8글자 이상) 또는 80% 이상 유사도 측정 (글자 일부 잘림/오탈자 대응)
  const lcsLen = getLCSLength(c1, c2);
  const maxLen = Math.max(c1.length, c2.length);
  const similarity = lcsLen / maxLen;

  if (similarity >= 0.78 || (lcsLen >= 8 && maxLen >= 10)) {
    return { isMatch: true, isExact: false, reason: `수령인명 지능형 유사 일치 (유사도 ${Math.round(similarity * 100)}%)` };
  }

  return { isMatch: false, isExact: false };
}

// 헬퍼: 정규화 텍스트 (특수문자 및 공백 완전 제거)
function normalizeText(val: any): string {
  if (!val) return "";
  return String(val)
    .replace(/[^\w\d가-힣]/g, "")
    .toLowerCase()
    .trim();
}

// 헬퍼: 텍스트에 대해 사전 등록된 대치 규칙(ProductMapping)을 양방향으로 적용하여 가능한 모든 대치 명칭 리스트 반환
function getPossibleMappedNames(text: string, mappings?: ProductMapping[]): string[] {
  if (!text) return [];
  const cleanText = String(text).replace(/\s+/g, "").toLowerCase();
  const results = new Set<string>();
  if (cleanText) results.add(cleanText);

  if (mappings && mappings.length > 0) {
    for (const m of mappings) {
      const raw = String(m.rawName || "").replace(/\s+/g, "").toLowerCase();
      const mapped = String(m.mappedName || "").replace(/\s+/g, "").toLowerCase();

      if (raw && cleanText.includes(raw)) {
        if (mapped) {
          results.add(cleanText.replace(raw, mapped));
          results.add(mapped);
        }
      }
      if (mapped && cleanText.includes(mapped)) {
        if (raw) {
          results.add(cleanText.replace(mapped, raw));
          results.add(raw);
        }
      }
    }
  }
  return Array.from(results);
}

// 헬퍼: 상품명/옵션명 및 품목 대치(ProductMapping) 양쪽 교차 일치 검사
function checkProductMatch(
  sourceProd: string,     // 2단계 송장파일의 상품명
  sourceUnit: string,     // 2단계 송장파일의 옵션/용량
  oOriginalProd: string,  // 1단계 원본 상품명
  oMappedProd: string,    // 1단계 매핑/표준 상품명
  oOriginalOpt: string,   // 1단계 원본 옵션명
  oMappedOpt: string,     // 1단계 매핑/표준 옵션명
  mappings?: ProductMapping[]
): { isMatch: boolean; reason: string } {
  // 송장 측 대치어 후보군 생성 (송장 상품명/옵션명도 대치 규칙 적용)
  const srcProdCandidates = getPossibleMappedNames(sourceProd, mappings);
  const srcUnitCandidates = getPossibleMappedNames(sourceUnit, mappings);

  // 주문서 측 대치어 후보군 생성 (원본/대치 상품명 및 옵션명 모두 대치 규칙 적용)
  const orderProdCandidates = new Set<string>([
    ...getPossibleMappedNames(oOriginalProd, mappings),
    ...getPossibleMappedNames(oMappedProd, mappings)
  ]);
  const orderUnitCandidates = new Set<string>([
    ...getPossibleMappedNames(oOriginalOpt, mappings),
    ...getPossibleMappedNames(oMappedOpt, mappings)
  ]);

  // 1) 상품명 후보군 간 양쪽 대치 교차 일치 검사
  for (const sP of srcProdCandidates) {
    if (!sP) continue;
    for (const oP of orderProdCandidates) {
      if (!oP) continue;
      if (sP === oP) {
        return { isMatch: true, reason: `상품명 양쪽 대치 교차 완전 일치 ('${sP}')` };
      }
      if (sP.length >= 2 && oP.length >= 2 && (sP.includes(oP) || oP.includes(sP))) {
        return { isMatch: true, reason: `상품명 양쪽 대치 교차 포함 일치 ('${sP}' ↔ '${oP}')` };
      }
    }
  }

  // 2) 옵션/용량 후보군 간 양쪽 대치 교차 일치 검사
  for (const sU of srcUnitCandidates) {
    if (!sU) continue;
    for (const oU of orderUnitCandidates) {
      if (!oU) continue;
      if (sU === oU) {
        return { isMatch: true, reason: `옵션/용량 양쪽 대치 교차 완전 일치 ('${sU}')` };
      }
      if (sU.length >= 2 && oU.length >= 2 && (sU.includes(oU) || oU.includes(sU))) {
        return { isMatch: true, reason: `옵션/용량 양쪽 대치 교차 포함 일치 ('${sU}' ↔ '${oU}')` };
      }
    }
  }

  return { isMatch: false, reason: "상품명/옵션명 미일치" };
}

// 2-2. 복호화되거나 이미 읽어온 2차원 배열 데이터를 기반으로 운송장 소스 매칭 분석을 수행하는 로직
export function processTrackingUpdateWithRows(
  inputOrders: ProcessedOrder[],
  srcRows: any[][],
  mappings?: ProductMapping[],
  couriers?: CourierConfig[],
  selectedCourierName?: string,
  sourceFileName?: string
): TrackingUpdateResult {
  if (!srcRows || srcRows.length === 0) {
    return {
      matchCount: 0,
      softMatchCount: 0,
      bundledCount: 0,
      bundledGroupCount: 0,
      unfilledErrors: inputOrders.map(o => ({
        type: 'unfilled',
        name: o.수령인,
        prod: o.상품명,
        qty: String(o.수량),
        zipCode: o.우편번호,
        reason: "소스 파일에 데이터가 없습니다.",
        sourceFileName: o.sourceFileName || "주문서 엑셀",
        platformName: o.platformName
      })),
      unusedErrors: [],
      outputOrders: inputOrders.map(o => ({ ...o }))
    };
  }

  // 헤더 행 및 컬럼 인덱스 동적 감지
  let SRC_TRACK_IDX = -1;
  let SRC_NAME_IDX = -1;
  let SRC_ZIP_IDX = -1;
  let SRC_PHONE_IDX = -1;
  let SRC_PROD_IDX = -1;
  let SRC_QTY_IDX = -1;
  let SRC_UNIT_IDX = -1;
  let SRC_COURIER_IDX = -1;
  let SRC_ADDR_IDX = -1;
  let startDataRow = 1;

  let courierConfigFound = false;

  // 0-1) [최우선 1순위] UI에서 선택한 택배사(또는 기본 한진택배)에 대한 컬럼 설정을 무조건 우선 적용
  const allAvailableCouriers = [
    ...(couriers || []),
    ...DEFAULT_COURIERS
  ];
  const targetCourierName = (selectedCourierName || "한진택배").trim();
  
  let targetCourier = allAvailableCouriers.find(c => 
    c.id === targetCourierName ||
    c.name === targetCourierName ||
    targetCourierName.includes(c.name) ||
    c.name.includes(targetCourierName)
  );

  // 만약 선택된 택배사가 한진택배 관련 명칭이면 'hanjin_express' 규격으로 포괄 지정
  if (!targetCourier && targetCourierName.includes("한진")) {
    targetCourier = DEFAULT_COURIERS.find(c => c.id === 'hanjin_express' || c.name.includes('한진'));
  }

  if (targetCourier) {
    startDataRow = Math.max(1, targetCourier.start_row);
    SRC_TRACK_IDX = targetCourier.tracking_col;
    if (targetCourier.courier_col !== undefined && targetCourier.courier_col !== -1) {
      SRC_COURIER_IDX = targetCourier.courier_col;
    }
    SRC_NAME_IDX = targetCourier.name_col;
    SRC_ZIP_IDX = targetCourier.zip_col;
    SRC_PHONE_IDX = targetCourier.phone_col;
    SRC_ADDR_IDX = targetCourier.addr_col;
    SRC_PROD_IDX = targetCourier.prod_col;
    SRC_UNIT_IDX = targetCourier.unit_col;
    SRC_QTY_IDX = targetCourier.qty_col;
    courierConfigFound = true;
  }

  // 0-2) 선택된 택배사 매칭 실패 시만 엑셀 상단 셀 텍스트로 택배사 패턴 자동 탐색
  if (!courierConfigFound && allAvailableCouriers.length > 0) {
    const sampleText = srcRows.slice(0, 15).map(r => (r ? r.join(" ") : "")).join(" ").toLowerCase();
    
    for (const c of allAvailableCouriers) {
      const patterns = [c.name, ...(c.filepath_pattern ? c.filepath_pattern.split(',') : [])]
        .map(p => p.trim().toLowerCase())
        .filter(p => p.length > 0);
        
      const matched = patterns.some(p => sampleText.includes(p));
      if (matched) {
        startDataRow = Math.max(1, c.start_row);
        SRC_TRACK_IDX = c.tracking_col;
        if (c.courier_col !== undefined && c.courier_col !== -1) SRC_COURIER_IDX = c.courier_col;
        SRC_NAME_IDX = c.name_col;
        SRC_ZIP_IDX = c.zip_col;
        SRC_PHONE_IDX = c.phone_col;
        SRC_ADDR_IDX = c.addr_col;
        SRC_PROD_IDX = c.prod_col;
        SRC_UNIT_IDX = c.unit_col;
        SRC_QTY_IDX = c.qty_col;
        courierConfigFound = true;
        break;
      }
    }
  }

  // 1) 택배사 특정 설정이 발견되지 않은 경우 헤더 자동 감지 스캔
  if (!courierConfigFound) {
    const scanLimit = Math.min(20, srcRows.length);
    let bestScore = -1;

    for (let r = 0; r < scanLimit; r++) {
      const row = srcRows[r];
      if (!row || row.length === 0) continue;

      let localTrack = -1, localName = -1, localZip = -1, localPhone = -1;
      let localProd = -1, localQty = -1, localUnit = -1, localCourier = -1, localAddr = -1;
      let score = 0;

      row.forEach((cell: any, colIdx: number) => {
        const s = String(cell || "").trim().toLowerCase().replace(/\s+/g, "");
        if (!s) return;

        // 보낸분 / 송하인 / 발송인 / 위탁자 헤더는 수취인 정보 식별에서 제외
        const isSender = s.includes("보낸분") || s.includes("송하인") || s.includes("발송인") || s.includes("보내는분") || s.includes("위탁자");

        if (localTrack === -1 && (s.includes("송장") || s.includes("운송장") || s.includes("등기") || s.includes("waybill") || s.includes("tracking"))) {
          localTrack = colIdx;
          score += 10;
        } else if (!isSender && (s.includes("받는분") || s.includes("수령인") || s.includes("수취인") || s.includes("받는사람") || s.includes("수화인") || s.includes("고객명") || s === "이름" || s === "성명")) {
          if (localName === -1 || s.includes("받는") || s.includes("수령") || s.includes("수취")) {
            localName = colIdx;
            score += (s.includes("받는") || s.includes("수령") || s.includes("수취")) ? 10 : 7;
          }
        } else if (!isSender && (s.includes("우편번호") || s === "우편" || s.includes("zip"))) {
          // 보낸분 우편번호 배제 및 받는분/수령인/배송지 우편번호 우선 감지
          if (localZip === -1 || s.includes("받는") || s.includes("수령") || s.includes("배송지") || s.includes("수취")) {
            localZip = colIdx;
            score += (s.includes("받는") || s.includes("수령") || s.includes("배송지") || s.includes("수취")) ? 8 : 5;
          }
        } else if (!isSender && (s.includes("전화번호") || s.includes("연락처") || s.includes("휴대폰") || s.includes("핸드폰"))) {
          if (localPhone === -1 || s.includes("받는") || s.includes("수령") || s.includes("수취")) {
            localPhone = colIdx;
            score += 5;
          }
        } else if (localProd === -1 && (s.includes("상품명") || s.includes("품목명") || s.includes("품명") || s.includes("내용물") || s === "상품" || s === "품목")) {
          localProd = colIdx;
          score += 4;
        } else if (localQty === -1 && (s.includes("수량") || s.includes("구매수") || s === "qty")) {
          localQty = colIdx;
          score += 4;
        } else if (localUnit === -1 && (s.includes("옵션명") || s.includes("등록옵션") || s === "용량" || s === "옵션" || s === "규격")) {
          localUnit = colIdx;
          score += 3;
        } else if (localCourier === -1 && (s.includes("택배사") || s.includes("배송업체") || s.includes("택배사명") || s === "택배" || s === "배송사")) {
          localCourier = colIdx;
          score += 3;
        } else if (!isSender && (s.includes("주소") || s.includes("배송지"))) {
          if (localAddr === -1 || s.includes("받는") || s.includes("수령") || s.includes("배송지") || s.includes("수취")) {
            localAddr = colIdx;
            score += (s.includes("받는") || s.includes("수령") || s.includes("배송지") || s.includes("수취")) ? 5 : 2;
          }
        }
      });

      if (score > bestScore && (localTrack !== -1 || localName !== -1)) {
        bestScore = score;
        SRC_TRACK_IDX = localTrack;
        SRC_NAME_IDX = localName;
        SRC_ZIP_IDX = localZip;
        SRC_PHONE_IDX = localPhone;
        SRC_PROD_IDX = localProd;
        SRC_QTY_IDX = localQty;
        SRC_UNIT_IDX = localUnit;
        SRC_COURIER_IDX = localCourier;
        SRC_ADDR_IDX = localAddr;
        startDataRow = r + 1;
      }
    }
  }

  // 만약 헤더 스캔으로 찾지 못한 경우 데이터 패턴 스캔 시도 (상단 30행 스캔)
  if (SRC_TRACK_IDX === -1 || SRC_NAME_IDX === -1) {
    const dataScanRows = Math.min(30, srcRows.length);
    for (let r = 0; r < dataScanRows; r++) {
      const row = srcRows[r];
      if (!row) continue;
      row.forEach((cell: any, cIdx: number) => {
        const val = String(cell || "").trim();
        // 송장번호 패턴 (숫자/하이픈 9자리 이상)
        if (SRC_TRACK_IDX === -1 && /^[\d-]{9,16}$/.test(val)) {
          SRC_TRACK_IDX = cIdx;
        }
        // 한국어 수령인 이름 패턴 (2~4글자 한글)
        if (SRC_NAME_IDX === -1 && /^[가-힣]{2,5}$/.test(val) && val !== "수령인" && val !== "수취인" && val !== "상품명" && val !== "배송지") {
          SRC_NAME_IDX = cIdx;
        }
        // 우편번호 패턴 (5자리 숫자)
        if (SRC_ZIP_IDX === -1 && /^\d{5}$/.test(val)) {
          SRC_ZIP_IDX = cIdx;
        }
      });
    }
  }

  // 헬퍼: 정규화 및 크리닝
  const cleanStr = (val: any) => String(val || "").replace(/\s+/g, "").trim().toLowerCase();
  const cleanZip = (val: any) => {
    let z = String(val || "").replace(/[^\d]/g, "");
    if (z.length === 4) z = "0" + z;
    return z;
  };
  const cleanPhone = (val: any) => String(val || "").replace(/[^\d]/g, "");

  // SourceItems 배열 구조 생성
  interface SourceItem {
    rawIndex: number;
    trackingNumber: string;
    courierName?: string;
    name: string;
    zip: string;
    phone: string;
    prod: string;
    unit: string;
    qty: number;
    addr: string;
    used: boolean;
    extraProducts?: { prod: string; unit: string; qty: number }[]; // C열 수량 > 1일 때 AW, BD, BK... 열의 추가 합포장 상품들
  }

  const sourceItems: SourceItem[] = [];

  for (let i = startDataRow; i < srcRows.length; i++) {
    const row = srcRows[i];
    if (!row || row.length === 0) continue;

    const rawTrack = SRC_TRACK_IDX !== -1 ? String(row[SRC_TRACK_IDX] || "").trim() : "";
    if (!rawTrack || rawTrack === "nan" || rawTrack === "undefined" || rawTrack === "송장번호") continue;

    const name = SRC_NAME_IDX !== -1 ? cleanStr(row[SRC_NAME_IDX]) : "";
    const zip = SRC_ZIP_IDX !== -1 ? cleanZip(row[SRC_ZIP_IDX]) : "";
    const phone = SRC_PHONE_IDX !== -1 ? cleanPhone(row[SRC_PHONE_IDX]) : "";
    const prod = SRC_PROD_IDX !== -1 ? cleanStr(row[SRC_PROD_IDX]) : "";
    const unit = SRC_UNIT_IDX !== -1 ? cleanStr(row[SRC_UNIT_IDX]) : "";
    const qty = SRC_QTY_IDX !== -1 ? (Number(row[SRC_QTY_IDX]) || 1) : 1;
    const addr = SRC_ADDR_IDX !== -1 ? cleanStr(row[SRC_ADDR_IDX]) : "";
    const courierName = SRC_COURIER_IDX !== -1 ? String(row[SRC_COURIER_IDX] || "").trim() : undefined;

    // C열(인덱스 2)의 숫자가 2 이상이거나 AW(48), BD(55), BK(62)... 열에 추가 상품 정보가 기재된 경우 파싱
    const extraProducts: { prod: string; unit: string; qty: number }[] = [];
    const cColVal = Number(row[2]);
    const maxCheckCount = !isNaN(cColVal) && cColVal > 1 ? Math.min(10, cColVal - 1) : 10;

    for (let n = 1; n <= maxCheckCount; n++) {
      const pIdx = 48 + (n - 1) * 7; // AW = 48, BD = 55, BK = 62, BR = 69, BY = 76...
      const qIdx = 49 + (n - 1) * 7; // AX = 49, BE = 56, BL = 63...
      const uIdx = 50 + (n - 1) * 7; // AY = 50, BF = 57, BM = 64...

      if (pIdx >= row.length) break;

      const exProd = cleanStr(row[pIdx]);
      const exUnit = cleanStr(row[uIdx]);
      const exQty = Number(row[qIdx]) || 1;

      if (exProd || exUnit) {
        extraProducts.push({
          prod: exProd,
          unit: exUnit,
          qty: exQty
        });
      }
    }

    sourceItems.push({
      rawIndex: i,
      trackingNumber: rawTrack,
      courierName: courierName || undefined,
      name,
      zip,
      phone,
      prod,
      unit,
      qty,
      addr,
      used: false,
      extraProducts: extraProducts.length > 0 ? extraProducts : undefined
    });
  }

  // 복제본 생성
  const outputOrders = inputOrders.map(o => ({ ...o }));
  let matchCount = 0;
  let softMatchCount = 0;
  const unfilledErrors: MatchingError[] = [];

  // 헬퍼: 메인 상품 및 추가 합포장 상품(AW, BD, BK열 등) 대조
  const checkProductMatchWithExtra = (
    item: SourceItem,
    oOrigProd: string,
    oMappedProd: string,
    oOrigOpt: string,
    oMappedOpt: string
  ): { isMatch: boolean; reason: string } => {
    // 1) 메인 상품 대조
    const mainCheck = checkProductMatch(item.prod, item.unit, oOrigProd, oMappedProd, oOrigOpt, oMappedOpt, mappings);
    if (mainCheck.isMatch) return mainCheck;

    // 2) 합포장 추가 상품들 대조 (AW, BD, BK열 등)
    if (item.extraProducts && item.extraProducts.length > 0) {
      for (let idx = 0; idx < item.extraProducts.length; idx++) {
        const ex = item.extraProducts[idx];
        const exCheck = checkProductMatch(ex.prod, ex.unit, oOrigProd, oMappedProd, oOrigOpt, oMappedOpt, mappings);
        if (exCheck.isMatch) {
          return {
            isMatch: true,
            reason: `합포장 추가 품목(${idx + 1}번째 AW/BD/BK열) ${exCheck.reason}`
          };
        }
      }
    }

    return { isMatch: false, reason: "상품명 미일치" };
  };

  // 헬퍼: 상품 및 주소 검증 조건 (OR 조건: 상품명 100% OR 주소 100% 정규화 OR 사전 등록 매핑 100%)
  const checkProductOrAddressMatch = (
    item: SourceItem,
    oOrigProd: string,
    oMappedProd: string,
    oOrigOpt: string,
    oMappedOpt: string,
    oAddr: string
  ): { isMatched: boolean; isExact: boolean; reason: string } => {
    const normSrcProd = normalizeText(item.prod);
    const normSrcUnit = normalizeText(item.unit);
    const normOrigProd = normalizeText(oOrigProd);
    const normMappedProd = normalizeText(oMappedProd);
    const normOrigOpt = normalizeText(oOrigOpt);
    const normMappedOpt = normalizeText(oMappedOpt);

    const normSrcAddr = normalizeText(item.addr);
    const normOrderAddr = normalizeText(oAddr);

    // 1) [Condition 1] 상품명 및 옵션 완전 일치 (100% 동일)
    const isProdExact =
      (normSrcProd && (normSrcProd === normOrigProd || normSrcProd === normMappedProd)) ||
      (normSrcUnit && (normSrcUnit === normOrigOpt || normSrcUnit === normMappedOpt));

    if (isProdExact) {
      return { isMatched: true, isExact: true, reason: "상품명/옵션 100% 완전 일치" };
    }

    // 2) [Condition 2] 주소 정규화 일치 (100% 동일)
    if (normSrcAddr && normOrderAddr && normSrcAddr === normOrderAddr) {
      return { isMatched: true, isExact: true, reason: "주소 정규화(100%) 완전 일치" };
    }

    // 3) [Condition 3] 사전 등록된 상품/고객 매핑 규칙 일치
    if (mappings && mappings.length > 0 && (normSrcProd || normSrcUnit)) {
      for (const m of mappings) {
        const mOrig = normalizeText(m.rawName);
        const mMapped = normalizeText(m.mappedName);

        if (!mOrig && !mMapped) continue;

        const isRuleExact =
          (mOrig && normSrcProd === mOrig) ||
          (mMapped && normSrcProd === mMapped) ||
          (mOrig && normSrcUnit === mOrig) ||
          (mMapped && normSrcUnit === mMapped);

        if (isRuleExact) {
          if (mMapped === normMappedProd || mOrig === normOrigProd || normSrcProd === mMapped) {
            return {
              isMatched: true,
              isExact: true,
              reason: `사전 등록 상품 매핑 규칙 완전 일치 (${m.rawName} ↔ ${m.mappedName})`
            };
          }
        }
      }
    }

    // C열 > 1 합포장 추가 상품들 (AW, BD, BK...) 대조
    if (item.extraProducts && item.extraProducts.length > 0) {
      for (let idx = 0; idx < item.extraProducts.length; idx++) {
        const ex = item.extraProducts[idx];
        const exProd = normalizeText(ex.prod);
        const exUnit = normalizeText(ex.unit);

        if (
          (exProd && (exProd === normOrigProd || exProd === normMappedProd)) ||
          (exUnit && (exUnit === normOrigOpt || exUnit === normMappedOpt))
        ) {
          return { isMatched: true, isExact: true, reason: `합포장 추가품목(${idx + 1}열) 완전 일치` };
        }
      }
    }

    // 1~3 완전 일치에 미달할 경우: 텍스트 포함(부분) 매칭 수행 (지능형 부분 매칭 판단용)
    const partialCheck = checkProductMatchWithExtra(item, oOrigProd, oMappedProd, oOrigOpt, oMappedOpt);
    if (partialCheck.isMatch) {
      return { isMatched: true, isExact: false, reason: partialCheck.reason };
    }

    return { isMatched: false, isExact: false, reason: "상품/주소 매칭 미일치" };
  };

  // 헬퍼: 매칭 실패 시 원인을 정밀 분석하는 디버그 헬퍼
  const analyzeUnmatchedReason = (
    o: ProcessedOrder,
    oName: string,
    oZip: string,
    oPhone: string,
    oOrigProd: string,
    oMappedProd: string,
    oOrigOpt: string,
    oMappedOpt: string,
    oQty: number,
    oAddr: string
  ): { reason: string; debugDetail: string } => {
    // 혹시 동일 수취인에게 이미 매칭된 타 주문건 운송장이 존재하는지 확인 (합포장건 확인 - 단, 쿠팡 주문에 한함!)
    if (isCoupangOrder(o)) {
      const existingOrderWithSameName = outputOrders.find(prev => 
        prev.trackingNumber && isCoupangOrder(prev) && cleanStr(prev.수령인) === oName
      );

      if (existingOrderWithSameName) {
        return {
          reason: `📦 [합포장 묶음 출고 안내] 동일 수취인('${o.수령인}') 명의의 타 주문건으로 운송장(${existingOrderWithSameName.trackingNumber})이 통합 발행됨`,
          debugDetail: `[합포장 출고 안내] 수취인 '${o.수령인}' 고객에게 대표 운송장(${existingOrderWithSameName.trackingNumber})이 발행되어 본 주문건은 별도 송장 없이 합포장 묶음배송 처리된 상태입니다.`
        };
      }
    }

    const sameNameItems = sourceItems.filter(item => item.name && item.name === oName);

    if (sameNameItems.length === 0) {
      const similarNameItems = sourceItems.filter(item => item.name && (item.name.includes(oName) || oName.includes(item.name)));
      if (similarNameItems.length > 0) {
        return {
          reason: `수취인명 유사항목 발견 ('${similarNameItems[0].name}') - 정확히 일치하지 않음`,
          debugDetail: `주문서 수취인명: '${o.수령인}' | 송장파일 유사 수취인명: '${similarNameItems[0].name}' (스펠링/공백 차이)`
        };
      }
      return {
        reason: `송장 엑셀에서 수취인명 '${o.수령인}'을 가진 데이터 행을 찾을 수 없음`,
        debugDetail: `[수취인명 미존재] 송장파일 읽기열(${SRC_NAME_IDX >= 0 ? colNumToLetter(SRC_NAME_IDX + 1) : '미지정'}열) 내에 '${o.수령인}' 텍스트가 존재하지 않습니다.`
      };
    }

    const availableSameName = sameNameItems.filter(item => !item.used);
    if (availableSameName.length === 0) {
      const usedItem = sameNameItems[0];
      const sameRecipientCountInCoupangOrders = inputOrders.filter(ord => isCoupangOrder(ord) && checkRecipientNameMatch(ord.수령인, o.수령인).isMatch).length;
      
      if (isCoupangOrder(o) && sameRecipientCountInCoupangOrders >= 2) {
        return {
          reason: `📦 [합포장 묶음 출고 안내] 동일 수취인('${o.수령인}') 대표 송장(${usedItem.trackingNumber})에 합포장 포함됨`,
          debugDetail: `[합포장 소진] 송장 엑셀의 '${o.수령인}' 수량(${sameNameItems.length}건)이 앞선 주문건들에 매칭되어 소진되었습니다. 본 품목은 대표 송장(${usedItem.trackingNumber})으로 합포장 출고 처리되었습니다.`
        };
      } else {
        return {
          reason: `송장 엑셀의 '${o.수령인}' 데이터가 이미 앞선 항목에 매칭 소진됨`,
          debugDetail: `[송장 소진] 송장 엑셀의 수량이 이미 앞선 항목에 대조 매칭되었습니다. (타 플랫폼 주문건은 무조건 독립적인 운송장 매칭이 필요합니다.)`
        };
      }
    }

    const diffs: string[] = [];
    const item = availableSameName[0];

    const zipMatch = oZip && item.zip ? cleanZip(oZip) === cleanZip(item.zip) : true;
    if (!zipMatch) {
      diffs.push(`우편번호 불일치 (주문서: '${oZip}' vs 송장: '${item.zip}')`);
    }

    const qtyMatch = !item.qty || item.qty === oQty;
    if (!qtyMatch) {
      diffs.push(`수량 불일치 (주문서: ${oQty}개 vs 송장: ${item.qty}개)`);
    }

    const prodOrAddrCheck = checkProductOrAddressMatch(item, oOrigProd, oMappedProd, oOrigOpt, oMappedOpt, oAddr);
    if (!prodOrAddrCheck.isMatched) {
      const srcProdDesc = [item.prod, item.unit].filter(Boolean).join(' / ') || '없음';
      const orderProdDesc = [oOrigProd, oOrigOpt].filter(Boolean).join(' / ') || '없음';
      diffs.push(`상품명/옵션명 불일치 (주문서: '${orderProdDesc}' [대치후: '${oMappedProd}'] vs 송장: '${srcProdDesc}')`);
      
      if (item.extraProducts && item.extraProducts.length > 0) {
        const extraDesc = item.extraProducts.map(e => `${e.prod}(${e.qty})`).join(', ');
        diffs.push(`참고: 송장 엑셀 합포장 추가열(AW, BD 등) 품목 [${extraDesc}]과도 일치하지 않음`);
      }
    }

    const mainReason = diffs.length > 0 ? diffs[0] : "매칭 조건 미충족";
    const debugDetail = `[수취인 '${o.수령인}' 송장 데이터 대조 구문]\n- ` + diffs.join('\n- ');

    return {
      reason: mainReason,
      debugDetail
    };
  };

  // 다단계(Multi-Pass) 스마트 매칭 실행
  outputOrders.forEach(o => {
    if (!o.수령인) return;

    const oName = cleanStr(o.수령인);
    const oZip = cleanZip(o.우편번호);
    const oPhone = cleanPhone(o.연락처);
    const oOrigProd = o.originalProductName || o.상품명 || "";
    const oMappedProd = o.상품명 || "";
    const oOrigOpt = o.originalOptionName || o.용량 || "";
    const oMappedOpt = o.용량 || "";
    const oQty = Number(o.수량) || 1;
    const oAddr = cleanStr(o.주소);

    let matchedItem: SourceItem | null = null;
    let matchType: 'exact' | 'soft' | 'bundled' = 'exact';
    let softMatchReason = '';

    // Tier 1: [100% 완전 매칭 (Exact Match)]
    // 수령인 성명 100% 일치 AND (우편번호/연락처/주소 포함 일치) AND 수량 일치 AND [상품 조건 100% 완전 일치]
    for (const item of sourceItems) {
      if (item.used) continue;
      if (item.name) {
        const nameCheck = checkRecipientNameMatch(item.name, oName);
        if (!nameCheck.isMatch || !nameCheck.isExact) continue;

        const zipMatch = oZip && item.zip && oZip === item.zip;
        const phoneMatch = oPhone && item.phone && oPhone === item.phone;
        const addrMatch = oAddr && item.addr && (item.addr.includes(oAddr.slice(0, 8)) || oAddr.includes(item.addr.slice(0, 8)));
        const infoMatch = zipMatch || phoneMatch || addrMatch;
        const qtyMatch = !item.qty || item.qty === oQty;

        const prodOrAddrCheck = checkProductOrAddressMatch(item, oOrigProd, oMappedProd, oOrigOpt, oMappedOpt, oAddr);

        if (infoMatch && qtyMatch && prodOrAddrCheck.isMatched && prodOrAddrCheck.isExact) {
          matchedItem = item;
          matchType = 'exact';
          softMatchReason = `수령인/연락처/수량 일치 및 [${prodOrAddrCheck.reason}] 완전 매칭`;
          break;
        }
      }
    }

    // Tier 2: [식별정보 완화 완전 매칭]
    if (!matchedItem) {
      for (const item of sourceItems) {
        if (item.used) continue;
        if (item.name) {
          const nameCheck = checkRecipientNameMatch(item.name, oName);
          if (!nameCheck.isMatch || !nameCheck.isExact) continue;

          const zipMatch = oZip && item.zip && oZip === item.zip;
          const phoneMatch = oPhone && item.phone && oPhone === item.phone;
          const infoMatch = zipMatch || phoneMatch;
          const qtyMatch = !item.qty || item.qty === oQty;

          const prodOrAddrCheck = checkProductOrAddressMatch(item, oOrigProd, oMappedProd, oOrigOpt, oMappedOpt, oAddr);

          if (infoMatch && qtyMatch && prodOrAddrCheck.isMatched) {
            matchedItem = item;
            matchType = 'exact';
            softMatchReason = `수령인/연락처/수량 일치 및 [${prodOrAddrCheck.reason}] 교차 완전 매칭`;
            break;
          }
        }
      }
    }

    // Tier 3: [지능형 부분 매칭 (Soft Match) - 수령인 성명 유사/잘림/부분일치 포함]
    if (!matchedItem) {
      for (const item of sourceItems) {
        if (item.used) continue;
        if (item.name) {
          const nameCheck = checkRecipientNameMatch(item.name, oName);
          if (!nameCheck.isMatch) continue;

          // 우편번호가 양쪽에 모두 있는 경우 일치 여부 검사
          const zipMatch = oZip && item.zip ? cleanZip(oZip) === cleanZip(item.zip) : true;
          if (!zipMatch) continue;

          const prodOrAddrCheck = checkProductOrAddressMatch(item, oOrigProd, oMappedProd, oOrigOpt, oMappedOpt, oAddr);

          if (prodOrAddrCheck.isMatched) {
            matchedItem = item;
            matchType = 'soft';
            softMatchReason = `${nameCheck.reason || '수령인 일치'} + [${prodOrAddrCheck.reason}] 기반 지능형 부분 매칭`;
            break;
          }
        }
      }
    }

    // Tier 4: [지능형 부분 매칭] 고유 수령인 명의 및 우편번호 1:1 유체 매칭
    if (!matchedItem) {
      const candidates = sourceItems.filter(item => {
        if (item.used) return false;
        if (!item.name) return false;
        const nameCheck = checkRecipientNameMatch(item.name, oName);
        if (!nameCheck.isMatch) return false;
        if (oZip && item.zip && cleanZip(oZip) !== cleanZip(item.zip)) return false;
        return true;
      });
      const orderCountWithSameName = inputOrders.filter(ord => checkRecipientNameMatch(ord.수령인, oName).isMatch).length;

      if (candidates.length === 1 && orderCountWithSameName === 1) {
        matchedItem = candidates[0];
        matchType = 'soft';
        softMatchReason = '수령인명 유사 및 우편번호 일치 기반 고유 1:1 지능형 매칭';
      }
    }

    // Tier 5: [독립 카테고리: 동일 수취인/배송지 합포장 묶음 자동 운송장 연동]
    // 단, 1) 해당 주문이 '쿠팡(DeliveryList)' 플랫폼일 때만 합포장 연동 수행! (다른 플랫폼은 합포장 미적용)
    // 2) 전체 주문서 내에 동일 수취인/배송지의 쿠팡 주문이 2건 이상일 때만 합포장 연동 수행!
    if (!matchedItem && isCoupangOrder(o)) {
      const sameRecipientOrdersInInput = inputOrders.filter(ord => 
        isCoupangOrder(ord) &&
        checkRecipientNameMatch(ord.수령인, oName).isMatch &&
        (!oZip || !ord.우편번호 || cleanZip(ord.우편번호) === oZip)
      );

      if (sameRecipientOrdersInInput.length >= 2) {
        const alreadyMatchedOrder = outputOrders.find(prev => 
          prev.trackingNumber && 
          isCoupangOrder(prev) &&
          checkRecipientNameMatch(prev.수령인, oName).isMatch &&
          (!oZip || !prev.우편번호 || cleanZip(prev.우편번호) === oZip)
        );

        if (alreadyMatchedOrder) {
          o.trackingNumber = alreadyMatchedOrder.trackingNumber;
          o.matchType = 'bundled';
          o.softMatchReason = `📦 [합포장 묶음 출고] 동일 수취인('${oName}') 타 주문건과 대표 송장(${alreadyMatchedOrder.trackingNumber})으로 통합 발송`;
          if (alreadyMatchedOrder.courierName) {
            o.courierName = alreadyMatchedOrder.courierName;
          }
          return;
        }
      }
    }

    if (matchedItem) {
      matchedItem.used = true;
      o.trackingNumber = matchedItem.trackingNumber;
      o.matchType = matchType;
      o.softMatchReason = softMatchReason;
      if (matchedItem.courierName) {
        o.courierName = matchedItem.courierName;
      }
    } else {
      const analysis = analyzeUnmatchedReason(o, oName, oZip, oPhone, oOrigProd, oMappedProd, oOrigOpt, oMappedOpt, oQty, oAddr);
      o.matchType = 'unmatched';
      o.unmatchedDetail = analysis.debugDetail;

      unfilledErrors.push({
        type: 'unfilled',
        name: o.수령인,
        prod: o.상품명,
        qty: String(o.수량),
        zipCode: o.우편번호,
        reason: analysis.reason,
        debugDetail: analysis.debugDetail,
        sourceFileName: o.sourceFileName || "주문서 엑셀",
        platformName: o.platformName
      });
    }
  });

  // 미사용 운송장 탐색
  const unusedErrors: MatchingError[] = [];
  sourceItems.forEach(item => {
    if (!item.used) {
      unusedErrors.push({
        type: 'unused',
        name: item.name || "알수없음",
        prod: item.prod || item.unit || "알수없음",
        qty: String(item.qty),
        zipCode: item.zip || "알수없음",
        reason: "취소 주문이거나 인풋 파일에서 이미 제외된 항목",
        sourceFileName: sourceFileName || "송장 엑셀 파일"
      });
    }
  });

  // ==========================================
  // 쿠팡 플랫폼 합포장(묶음배송) 그룹 후처리 및 C열 수량 대조 검사
  // ==========================================
  const trackingNumberCounts: { [tracking: string]: number } = {};
  outputOrders.forEach(o => {
    if (o.trackingNumber && isCoupangOrder(o)) {
      trackingNumberCounts[o.trackingNumber] = (trackingNumberCounts[o.trackingNumber] || 0) + 1;
    }
  });

  outputOrders.forEach(o => {
    if (o.trackingNumber && isCoupangOrder(o)) {
      const count = trackingNumberCounts[o.trackingNumber] || 0;
      if (count >= 2) {
        // 동일 대표 송장번호로 2개 이상의 쿠팡 품목이 매칭된 경우, 합포장 묶음 출고건으로 처리
        if (o.matchType !== 'soft') {
          o.matchType = 'bundled';
        }
        o.softMatchReason = `📦 [합포장 묶음 출고] 동일 수취인('${o.수령인}') 총 ${count}개 품목 통합 배송 (대표송장: ${o.trackingNumber})`;
      }
    }
  });

  // C열(합포) 수량 및 실제 매칭 수량 일치 여부 검사
  const bundleMismatchDetails: string[] = [];
  const checkedTrackings = new Set<string>();

  sourceItems.forEach(item => {
    if (!item.trackingNumber || checkedTrackings.has(item.trackingNumber)) return;
    checkedTrackings.add(item.trackingNumber);

    const cColVal = Number(srcRows[item.rawIndex] ? srcRows[item.rawIndex][2] : 0);
    const expectedQty = !isNaN(cColVal) && cColVal >= 2 ? cColVal : 1;

    const matchedForTrack = outputOrders.filter(o => o.trackingNumber === item.trackingNumber && isCoupangOrder(o));
    const actualQty = matchedForTrack.length;

    if (expectedQty >= 2 || actualQty >= 2) {
      if (expectedQty !== actualQty) {
        const recipName = item.name || (matchedForTrack[0] ? matchedForTrack[0].수령인 : "미상");
        bundleMismatchDetails.push(
          `[합포장 수량 불일치] 수령인 '${recipName}' (송장번호: ${item.trackingNumber}): 2단계 엑셀 C열(합포) 수량은 ${expectedQty}개이나, 실제 매칭된 주문 품목수는 ${actualQty}개입니다.`
        );
      }
    }
  });

  const exactMatchOrders = outputOrders.filter(o => o.matchType === 'exact');
  const softMatchOrders = outputOrders.filter(o => o.matchType === 'soft');
  const bundledOrders = outputOrders.filter(o => o.matchType === 'bundled');

  matchCount = exactMatchOrders.length;
  softMatchCount = softMatchOrders.length;
  const bundledCount = bundledOrders.length;

  const bundledGroupCount = new Set(
    bundledOrders.map(o => `${cleanStr(o.수령인)}_${cleanZip(o.우편번호)}_${o.trackingNumber || ''}`)
  ).size;

  return {
    matchCount,
    softMatchCount,
    bundledCount,
    bundledGroupCount,
    bundleMismatchCount: bundleMismatchDetails.length,
    bundleMismatchDetails,
    unfilledErrors,
    unusedErrors,
    outputOrders
  };
}

export async function processTrackingUpdate(
  inputOrders: ProcessedOrder[],
  sourceFile: File,
  mappings?: ProductMapping[],
  couriers?: CourierConfig[],
  selectedCourierName?: string
): Promise<TrackingUpdateResult> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const bestSheetName = findBestOrderSheet(workbook);
        const worksheet = workbook.Sheets[bestSheetName];
        
        // 2차원 배열 데이터 로드
        const srcRows = XLSX.utils.sheet_to_json<any[]>(worksheet, { header: 1, defval: "" });
        const result = processTrackingUpdateWithRows(inputOrders, srcRows, mappings, couriers, selectedCourierName, sourceFile.name);
        resolve(result);
      } catch (err: any) {
        const errMsg = err?.message || "";
        const lowerMsg = errMsg.toLowerCase();
        if (
          lowerMsg.includes("password") ||
          lowerMsg.includes("protected") ||
          lowerMsg.includes("decrypt") ||
          lowerMsg.includes("encrypted")
        ) {
          reject(new Error("PASSWORD_PROTECTED"));
        } else {
          reject(err);
        }
      }
    };
    reader.onerror = () => reject(new Error("소스 파일을 읽는 도중 오류가 발생했습니다."));
    reader.readAsArrayBuffer(sourceFile);
  });
}

// 최종 운송장이 채워진 엑셀 파일 내보내기 (원본 양식, 노란 배경, 파란 헤더 100% 보존)
export async function exportFinalTrackingExcel(
  orders: ProcessedOrder[],
  originalRows: any[][],
  platform: PlatformConfig,
  outputFileName?: string,
  originalFileOrBuffer?: File | ArrayBuffer | Uint8Array,
  defaultCourierName: string = "한진택배"
): Promise<void> {
  const wb = new ExcelJS.Workbook();
  let ws: ExcelJS.Worksheet | undefined;
  let loadedFromOriginal = false;

  if (originalFileOrBuffer) {
    try {
      let buffer: ArrayBuffer;
      if (originalFileOrBuffer instanceof File) {
        buffer = await originalFileOrBuffer.arrayBuffer();
      } else if (originalFileOrBuffer instanceof Uint8Array) {
        buffer = originalFileOrBuffer.buffer.slice(
          originalFileOrBuffer.byteOffset,
          originalFileOrBuffer.byteOffset + originalFileOrBuffer.byteLength
        ) as ArrayBuffer;
      } else {
        buffer = originalFileOrBuffer;
      }

      await wb.xlsx.load(buffer);
      
      let bestWs: ExcelJS.Worksheet | undefined = wb.worksheets[0];
      let maxScore = -1;

      wb.worksheets.forEach(sheet => {
        if (!sheet) return;
        let score = 0;
        const sheetName = (sheet.name || "").toLowerCase();
        if (sheetName.includes("발송") || sheetName.includes("주문") || sheetName.includes("delivery") || sheetName.includes("sheet1")) {
          score += 5;
        }
        
        const scanRows = Math.min(15, sheet.rowCount || 15);
        for (let r = 1; r <= scanRows; r++) {
          const rowVal = sheet.getRow(r).values;
          if (Array.isArray(rowVal)) {
            const rowStr = rowVal.map(v => String(v || "")).join(" ").toLowerCase();
            if (rowStr.includes("주문번호")) score += 10;
            if (rowStr.includes("수취인명") || rowStr.includes("수령인") || rowStr.includes("수취인")) score += 8;
            if (rowStr.includes("상품명")) score += 5;
            if (rowStr.includes("송장번호") || rowStr.includes("운송장")) score += 10;
            if (rowStr.includes("택배사")) score += 5;
          }
        }

        if (score > maxScore) {
          maxScore = score;
          bestWs = sheet;
        }
      });

      ws = bestWs;
      if (ws) {
        loadedFromOriginal = true;
      }
    } catch (e) {
      console.warn("원본 엑셀 템플릿 로드 실패, 기본 방식 사용:", e);
    }
  }

  // 송장번호 열 및 택배사 열 결정 (플랫폼 설정값 1순위 절대 기준)
  const trackingColIdx = platform.tracking_col !== undefined ? (platform.tracking_col + 1) : 9;
  const courierColIdx = platform.courier_col !== undefined ? (platform.courier_col + 1) : undefined;

  if (loadedFromOriginal && ws) {
    // orders에 존재하는 rawIdx와 { trackingNumber, courierName } 맵 생성
    const trackingByRawIdx = new Map<number, { trackingNumber: string; courierName?: string }>();
    orders.forEach(o => {
      const rawIdx = Number(o.id);
      if (!isNaN(rawIdx) && o.trackingNumber) {
        trackingByRawIdx.set(rawIdx, {
          trackingNumber: o.trackingNumber,
          courierName: o.courierName
        });
      }
    });

    // trackingByRawIdx에 존재하는 모든 0-based rawIdx 항목에 대해 직접 1-based 행(rawIdx + 1)을 얻어 값 기입
    trackingByRawIdx.forEach((item, rawIdx) => {
      const rowNum = rawIdx + 1; // 1-based index
      const row = ws!.getRow(rowNum);

      if (item.trackingNumber && trackingColIdx) {
        const cell = row.getCell(trackingColIdx);
        cell.value = item.trackingNumber;
      }
      const courierToSet = item.courierName || defaultCourierName;
      if (courierToSet && courierColIdx) {
        const courierCell = row.getCell(courierColIdx);
        courierCell.value = courierToSet;
      }
      row.commit();
    });
  } else {
    // 원본 파일을 로드하지 못한 경우의 폴백 생성기
    ws = wb.addWorksheet('최종송장반영');
    const seenIds = new Set<string>();
    const rowsToSave: any[][] = [];

    const startRow = platform.start_row;
    for (let r = 0; r < startRow; r++) {
      rowsToSave.push(originalRows[r] || []);
    }

    orders.forEach(o => {
      const rawIdx = Number(o.id);
      const originalRow = [...(originalRows[rawIdx] || [])];
      
      const orderId = String(originalRow[0] || "").trim();
      
      if (orderId) {
        if (seenIds.has(orderId)) {
          return;
        }
        seenIds.add(orderId);
      }

      if (o.trackingNumber && trackingColIdx) {
        originalRow[trackingColIdx - 1] = o.trackingNumber;
      }
      const courierToSet = o.courierName || defaultCourierName;
      if (courierToSet && courierColIdx) {
        originalRow[courierColIdx - 1] = courierToSet;
      }
      
      rowsToSave.push(originalRow);
    });

    rowsToSave.forEach(row => {
      ws!.addRow(row);
    });

    ws.columns.forEach(col => {
      let maxLen = 0;
      col.eachCell!({ includeEmpty: false }, (cell) => {
        if (cell.value) {
          const length = String(cell.value).length;
          if (length > maxLen) maxLen = length;
        }
      });
      col.width = Math.max(10, maxLen + 1.5);
    });
  }

  // ----------------------------------------------------
  // F열(6번 열) 분리배송 'N' 중복 행 삭제 처리 (쿠팡/DeliveryList 합포장건 1개만 남기기)
  // 예: [분리배송불가, N, N, N, 분리배송불가] -> [분리배송불가, N, 분리배송불가]
  // ----------------------------------------------------
  if (ws) {
    const totalRows = ws.rowCount;
    const rowsToDelete: number[] = [];
    let seenNInGroup = false;
    let currentGroupKey = "";

    for (let r = 1; r <= totalRows; r++) {
      const row = ws.getRow(r);
      const valF = String(row.getCell(6).value || "").trim().toUpperCase();

      const colAKey = String(row.getCell(1).value || "").trim();

      const recipientKey = String(
        row.getCell(platform.col_map.수령인 !== undefined ? platform.col_map.수령인 + 1 : 2).value ||
        row.getCell(2).value || ""
      ).trim();

      const addrKey = String(
        row.getCell(platform.col_map.주소 !== undefined ? platform.col_map.주소 + 1 : 4).value || ""
      ).trim();

      // A열 번호가 제공된 경우 A열 번호를 그룹 키로 사용하고, 없으면 수령인+주소 조합 사용
      const groupKey = colAKey ? `ORDER_${colAKey}` : `${recipientKey}_${addrKey}`;

      if (groupKey !== currentGroupKey && recipientKey !== "") {
        currentGroupKey = groupKey;
        seenNInGroup = false;
      }

      if (valF === "N") {
        if (seenNInGroup) {
          // 이미 그룹 내에서 첫 번째 'N' 행을 하나 유지하였으므로 중복 'N' 행 삭제
          rowsToDelete.push(r);
        } else {
          // 그룹 내 첫 번째 'N' 행은 그대로 유지
          seenNInGroup = true;
        }
      }
    }

    if (rowsToDelete.length > 0) {
      for (let i = rowsToDelete.length - 1; i >= 0; i--) {
        ws.spliceRows(rowsToDelete[i], 1);
      }
    }
  }

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const today = new Date().toISOString().slice(0, 10);
  const finalName = outputFileName || `DeliveryList_최종송장_${platform.id}_${today}.xlsx`;
  triggerDownload(blob, finalName);
}

// 아이프레임(Iframe) 및 다양한 브라우저 환경에서 안정적으로 다운로드를 수행하는 헬퍼 함수
export function triggerDownload(blob: Blob, fileName: string) {
  try {
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
  } catch (err) {
    console.error("Standard download failed, falling back to file-saver saveAs:", err);
    saveAs(blob, fileName);
  }
}
