import * as XLSX from 'xlsx';

export type SheetRole =
  | 'HOUSEHOLD'
  | 'GIRO'
  | 'POST_OFFICE'
  | 'PAYMENT_STANDARD'
  | 'HISTORICAL_MAPPING'
  | 'UNKNOWN';

export interface SheetAnalysis {
  sheetName: string;
  role: SheetRole;
  confidence: number;
  reason: string;
  rowCount: number;
  headers: string[];
}

function normalizeSheetName(value: string) {
  return String(value ?? '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(/[^가-힣a-z0-9]/g, '');
}

function normalizeHeaderName(value: string) {
  return String(value ?? '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(/[^가-힣a-z0-9]/g, '');
}

function headerMatches(headers: string[], aliases: string[]) {
  const normalizedAliases = aliases.map(normalizeHeaderName).filter(Boolean);
  return headers.filter((header) => normalizeHeaderName(header)).some((header) => {
    const normalizedHeader = normalizeHeaderName(header);
    return normalizedAliases.some((alias) => normalizedHeader === alias || normalizedHeader.includes(alias) || alias.includes(normalizedHeader));
  });
}

function hasAllHeaders(headers: string[], aliases: string[]) {
  const normalizedHeaders = headers.map(normalizeHeaderName).filter(Boolean);
  return aliases.every((alias) => {
    const normalizedAlias = normalizeHeaderName(alias);
    return normalizedHeaders.some((header) => header === normalizedAlias || header.includes(normalizedAlias) || normalizedAlias.includes(header));
  });
}

function getSheetMatrix(worksheet: XLSX.WorkSheet) {
  return XLSX.utils.sheet_to_json<unknown[]>(worksheet, { header: 1, defval: '', raw: false })
    .map((row) => row.map((cell) => String(cell ?? '').trim()));
}

function findHeaderRow(matrix: string[][]) {
  let bestIndex = 0;
  let bestScore = 0;

  matrix.slice(0, 30).forEach((row, index) => {
    const headers = row.filter(Boolean);
    const score = [
      ['세대주', '주소', '자부담산정'],
      ['고객상호명', '납부금액', '설치주소'],
      ['거래일시', '입금액', '내역'],
      ['입금액원', '내역'],
      ['일반주택', '공동주택'],
    ].reduce((total, aliases) => total + (aliases.every((alias) => headerMatches(headers, [alias])) ? aliases.length : 0), 0);

    if (score > bestScore) {
      bestScore = score;
      bestIndex = index;
    }
  });

  return bestIndex;
}

function getRowsAndHeaders(worksheet: XLSX.WorkSheet) {
  const matrix = getSheetMatrix(worksheet);
  const headerRow = findHeaderRow(matrix);
  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(worksheet, {
    defval: '',
    raw: false,
    range: headerRow,
  });

  const rowsWithSourceRow = rows.map((row, index) => ({
    ...row,
    __sourceRow: headerRow + index + 2,
  }));

  return { rows: rowsWithSourceRow, headers: rows.length ? Object.keys(rows[0]) : (matrix[headerRow] ?? []) };
}

function inferSheetRole(sheetName: string, headers: string[], rows: Record<string, unknown>[]) {
  const normalizedSheetName = normalizeSheetName(sheetName);
  const headerList = headers.map(normalizeHeaderName);
  let score = 0;
  let reason = '기본값';

  if (normalizedSheetName.includes('신청') && normalizedSheetName.includes('세대')) score += 30;
  if (normalizedSheetName.includes('세대') && normalizedSheetName.includes('명단')) score += 28;
  if (normalizedSheetName.includes('지로')) score += 24;
  if (normalizedSheetName.includes('우체국')) score += 24;
  if (normalizedSheetName.includes('자부담')) score += 22;
  if (normalizedSheetName.includes('확인완료') || normalizedSheetName.includes('명단확인')) score += 20;

  if (headerMatches(headers, ['세대주', '주소', '자부담산정', '입금확인'])) score += 35;
  if (headerMatches(headers, ['세대주', '주소'])) score += 12;
  if (headerMatches(headers, ['이름', '주소', '자부담산정'])) score += 10;
  if (headerMatches(headers, ['고객상호명', '납부금액', '설치주소'])) score += 35;
  if (headerMatches(headers, ['거래일시', '입금액', '내역'])) score += 35;
  if (headerMatches(headers, ['거래일시', '입금액원', '내역'])) score += 35;
  if (headerMatches(headers, ['일반주택', '공동주택', '취소'])) score += 30;
  if (headerMatches(headers, ['입금날짜', '입금금액', '적용여부'])) score += 26;

  const amountColumns = headerList.filter((h) => h.includes('금액') || h.includes('자부담') || h.includes('납부'));
  if (amountColumns.length > 0) score += 6;

  const hasNameAndAddress = hasAllHeaders(headers, ['세대주', '주소']) || hasAllHeaders(headers, ['이름', '주소']);
  const hasHouseholdAmount = hasAllHeaders(headers, ['자부담산정']) || hasAllHeaders(headers, ['자부담금']);
  const hasGiroPattern = hasAllHeaders(headers, ['고객상호명', '납부금액', '설치주소']) || hasAllHeaders(headers, ['입금자명', '금액', '주소']);
  const hasPostPattern = hasAllHeaders(headers, ['거래일시', '입금액', '내역']) || hasAllHeaders(headers, ['거래일시', '입금액원', '내역']);
  const hasPaymentStandardPattern = hasAllHeaders(headers, ['일반주택', '공동주택', '취소']) || hasAllHeaders(headers, ['기준', '일반주택']);
  const hasHistoricalPattern = hasAllHeaders(headers, ['입금날짜', '입금금액', '적용여부']) || hasAllHeaders(headers, ['세대주', '입금금액', '비고']);

  let role: SheetRole = 'UNKNOWN';
  if (hasHouseholdAmount && hasNameAndAddress) {
    role = 'HOUSEHOLD';
    reason = '신청세대 필수 헤더 조합 감지';
  } else if (hasGiroPattern) {
    role = 'GIRO';
    reason = '지로 입금내역 구조 감지';
  } else if (hasPostPattern) {
    role = 'POST_OFFICE';
    reason = '우체국 입금내역 구조 감지';
  } else if (hasPaymentStandardPattern) {
    role = 'PAYMENT_STANDARD';
    reason = '자부담 기준표 구조 감지';
  } else if (hasHistoricalPattern) {
    role = 'HISTORICAL_MAPPING';
    reason = '과거 매칭자료 구조 감지';
  } else if ((normalizedSheetName.includes('신청') && normalizedSheetName.includes('세대')) || normalizedSheetName.includes('신청명단')) {
    role = 'HOUSEHOLD';
    reason = '시트명 기반 신청세대 판별';
  } else if (normalizedSheetName.includes('지로')) {
    role = 'GIRO';
    reason = '시트명 기반 지로 판별';
  } else if (normalizedSheetName.includes('우체국') || normalizedSheetName.includes('post')) {
    role = 'POST_OFFICE';
    reason = '시트명 기반 우체국 판별';
  } else if (normalizedSheetName.includes('자부담') || normalizedSheetName.includes('기준')) {
    role = 'PAYMENT_STANDARD';
    reason = '시트명 기반 자부담 기준 판별';
  } else if (normalizedSheetName.includes('완료') || normalizedSheetName.includes('매칭')) {
    role = 'HISTORICAL_MAPPING';
    reason = '시트명 기반 과거 매칭자료 판별';
  }

  if (role === 'UNKNOWN' && score <= 0 && rows.length > 0) {
    reason = '명확한 패턴을 찾지 못해 사용하지 않는 시트로 분류';
  }

  const confidence = Math.min(Math.max((score + (role !== 'UNKNOWN' ? 15 : 0)) / 100, 0.05), 0.99);

  return { role, confidence, reason };
}

export function analyzeWorkbook(workbook: XLSX.WorkBook): SheetAnalysis[] {
  return workbook.SheetNames.map((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    const { rows, headers } = getRowsAndHeaders(worksheet);
    const detected = inferSheetRole(sheetName, headers, rows);

    return {
      sheetName,
      role: detected.role,
      confidence: detected.confidence,
      reason: detected.reason,
      rowCount: rows.length,
      headers,
    };
  });
}

export function getRowsForRole(workbook: XLSX.WorkBook, role: SheetRole): Record<string, unknown>[] {
  const results: Record<string, unknown>[] = [];
  for (const sheetName of workbook.SheetNames) {
    const worksheet = workbook.Sheets[sheetName];
    const { rows, headers } = getRowsAndHeaders(worksheet);
    if (!rows.length) continue;
    const detected = inferSheetRole(sheetName, headers, rows);
    if (detected.role === role) {
      results.push(...rows);
    }
  }
  return results;
}

export function findSheetByRole(workbook: XLSX.WorkBook, role: SheetRole) {
  return workbook.SheetNames.map((sheetName) => ({ sheetName, worksheet: workbook.Sheets[sheetName] }))
    .find(({ sheetName, worksheet }) => {
      const { rows, headers } = getRowsAndHeaders(worksheet);
      return inferSheetRole(sheetName, headers, rows).role === role;
    });
}

export function getSheetRoleLabel(role: SheetRole) {
  switch (role) {
    case 'HOUSEHOLD':
      return '신청세대';
    case 'GIRO':
      return '지로';
    case 'POST_OFFICE':
      return '우체국';
    case 'PAYMENT_STANDARD':
      return '자부담 기준';
    case 'HISTORICAL_MAPPING':
      return '과거 매칭자료';
    default:
      return '기타';
  }
}
