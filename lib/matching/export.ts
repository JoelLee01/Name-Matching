import * as XLSX from 'xlsx';
import type { MatchResult } from './types';
import type { ValidationReportRow } from './validation';

export function buildResultWorkbook(
  sourceWorkbook: XLSX.WorkBook,
  results: MatchResult[],
  validationRows: ValidationReportRow[],
) {
  const workbook = XLSX.utils.book_new();

  sourceWorkbook.SheetNames.forEach((sheetName) => {
    XLSX.utils.book_append_sheet(workbook, sourceWorkbook.Sheets[sheetName], sheetName);
  });

  const resultRows = results.map((row) => ({
    세대주: row.householdName,
    주소: row.householdAddress ?? '',
    자부담금: row.householdAmount ?? '',
    상태: row.status,
    입금자명: row.matchedDeposit?.depositorName ?? '',
    입금금액: row.matchedDeposit?.amount ?? '',
    입금일: row.matchedDeposit?.originalData?.['거래일시'] ?? row.matchedDeposit?.originalData?.['입금일시'] ?? '',
    입금처: row.matchedDeposit?.source ?? '',
    원본행: row.matchedDeposit?.rowIndex ?? '',
    신뢰도: row.score,
    사유: row.reasons.join(', '),
  }));
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(resultRows), '입금매칭결과');

  const allocationRows = results.flatMap((row) => (row.allocations ?? []).map((allocation) => {
    const deposit = row.candidates.find((candidate) => candidate.id === allocation.depositId) ?? row.matchedDeposit;
    return {
      입금ID: allocation.depositId,
      출처: deposit?.source ?? '',
      입금자: deposit?.depositorName ?? '',
      입금일시: deposit?.originalData?.['거래일시'] ?? deposit?.originalData?.['입금일시'] ?? '',
      원입금액: deposit?.amount ?? '',
      배분세대: row.householdName,
      배분금액: allocation.allocatedAmount,
      배분상태: row.status,
    };
  }));
  if (allocationRows.length) XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(allocationRows), '입금배분결과');

  const verificationRows = validationRows.map((row) => ({
    '신청세대 행번호': row.householdRow,
    세대주: row.householdName,
    주소: row.householdAddress,
    자부담금: row.householdAmount,
    '기존 W 수식': row.formula,
    '기존 입금출처': row.existingSource,
    '기존 입금행': row.existingDepositRow,
    '기존 배분금액': row.existingAllocatedAmount,
    '프로그램 상태': row.programStatus,
    '프로그램 입금출처': row.programSource,
    '프로그램 입금행': row.programDepositRow,
    '프로그램 입금자': row.programDepositorName,
    '프로그램 입금금액': row.programDepositAmount,
    '프로그램 점수': row.programScore,
    비교결과: row.comparison,
    차이사유: row.differenceReason,
  }));
  if (verificationRows.length) XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(verificationRows), '기존매칭_검증');

  return workbook;
}
