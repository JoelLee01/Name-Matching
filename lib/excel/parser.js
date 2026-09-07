const XLSX = require('xlsx');

function normalizeHeader(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(/[()\[\]{}]/g, '')
    .replace(/[^가-힣a-z0-9]/g, '');
}

function detectColumns(headers) {
  const normalized = headers.map((header) => normalizeHeader(header));
  const map = {
    name: ['세대주명', '세대주', '성명', '신청자명', '입금자명', '고객상호명', '고객명', '입금자'],
    amount: ['자부담', '자부담금', '납부금액', '입금금액', '금액', '입금액', '거래금액', '납부금', '납부금액'],
    address: ['주소', '세대주소', '설치주소', '납부처주소'],
    confirmed: ['입금확인', '확인여부', '입금확인여부'],
  };

  const result = {};
  Object.entries(map).forEach(([key, aliases]) => {
    const found = headers.find((header, index) => {
      const candidate = normalizeHeader(header);
      return aliases.some((alias) => candidate.includes(normalizeHeader(alias)) || normalizeHeader(alias).includes(candidate));
    });
    result[key] = found || null;
  });

  return result;
}

function parseWorkbookFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const buffer = event.target?.result;
        const workbook = XLSX.read(buffer, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[firstSheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });
        resolve({ workbook, rows, sheetName: firstSheetName });
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = () => reject(new Error('엑셀 파일을 읽을 수 없습니다.'));
    reader.readAsArrayBuffer(file);
  });
}

function parseHouseholds(rows) {
  if (!rows.length) return [];
  const headers = Object.keys(rows[0]);
  const columns = detectColumns(headers);

  return rows.map((row, index) => ({
    id: `household-${index + 1}`,
    rowIndex: index + 1,
    name: row[columns.name] || '',
    address: row[columns.address] || '',
    requiredAmount: row[columns.amount] || 0,
    confirmed: row[columns.confirmed] || '',
    originalData: row,
  }));
}

function parseDeposits(rows, source) {
  if (!rows.length) return [];
  const headers = Object.keys(rows[0]);
  const columns = detectColumns(headers);

  return rows.map((row, index) => ({
    id: `${source.toLowerCase()}-${index + 1}`,
    rowIndex: index + 1,
    source,
    depositorName: row[columns.name] || '',
    amount: row[columns.amount] || 0,
    address: row[columns.address] || '',
    originalData: row,
  }));
}

module.exports = {
  normalizeHeader,
  detectColumns,
  parseWorkbookFile,
  parseHouseholds,
  parseDeposits,
};
