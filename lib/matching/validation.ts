import * as XLSX from 'xlsx';
import type { DepositRow, HouseholdRow, MatchResult } from './types';

export type ExistingSource = 'GIRO' | 'POST';
export type ValidationComparison =
  | 'EXACT_MATCH'
  | 'EXACT_SPLIT_MATCH'
  | 'PROGRAM_REVIEW'
  | 'PROGRAM_MISSED'
  | 'PROGRAM_DIFFERENT'
  | 'NO_EXISTING_MATCH';

export interface ExistingFormulaReference {
  householdRow: number;
  householdName: string;
  householdAddress: string;
  householdAmount: number;
  formula: string;
  source: ExistingSource;
  depositRow: number;
  depositRows: number[];
  divisor: number;
  allocatedAmount: number;
  deposit: DepositRow | null;
}

export interface ValidationReportRow {
  householdRow: number;
  householdName: string;
  householdAddress: string;
  householdAmount: number;
  formula: string;
  existingSource: string;
  existingDepositRow: string;
  existingAllocatedAmount: string | number;
  programStatus: string;
  programSource: string;
  programDepositRow: string | number;
  programDepositorName: string;
  programDepositAmount: string | number;
  programScore: string | number;
  comparison: ValidationComparison;
  differenceReason: string;
  programCandidates: Array<{
    source: string;
    row: number;
    depositorName: string;
    amount: number;
    address: string;
    date: string;
    score: number;
  }>;
}

export interface ValidationSummary {
  existingTotal: number;
  existingGiro: number;
  existingPost: number;
  existingSplit: number;
  exactMatch: number;
  exactSplitMatch: number;
  programReview: number;
  programMissed: number;
  programDifferent: number;
  noExistingMatch: number;
  reproductionRate: number;
}

export interface SafetyAudit {
  duplicateNameAuto: number;
  amountMismatchAuto: number;
  duplicateDepositAuto: number;
  fuzzyOnlyAuto: number;
  totalViolations: number;
}

function parseFormula(formula: string) {
  const normalized = String(formula ?? '').trim().replace(/^=/, '').replace(/\s+/g, '');
  const terms = normalized.split('+').map((term) => term.match(/^(우체국|지로)!([A-Z]+)(\d+)(?:\/(\d+(?:\.\d+)?))?$/i));
  if (terms.some((term) => !term)) return null;
  const parsedTerms = terms as RegExpMatchArray[];

  return {
    source: parsedTerms[0][1] === '우체국' ? 'POST' as const : 'GIRO' as const,
    depositRow: Number(parsedTerms[0][3]),
    depositRows: parsedTerms.map((term) => Number(term[3])),
    divisors: parsedTerms.map((term) => Number(term[4] ?? 1)),
  };
}

function getWorksheetRowCount(worksheet: XLSX.WorkSheet) {
  const range = worksheet['!ref'] ? XLSX.utils.decode_range(worksheet['!ref']) : { s: { r: 0 }, e: { r: 0 } };
  return range.e.r + 1;
}

export function readExistingFormulaReferences(
  workbook: XLSX.WorkBook,
  households: HouseholdRow[],
  deposits: DepositRow[],
): ExistingFormulaReference[] {
  const worksheet = workbook.Sheets['신청세대명단'];
  if (!worksheet) return [];

  const householdByRow = new Map(households.map((household) => [household.rowIndex, household]));
  const depositByKey = new Map(deposits.map((deposit) => [`${deposit.source}:${deposit.rowIndex}`, deposit]));
  const references: ExistingFormulaReference[] = [];

  for (let row = 1; row <= getWorksheetRowCount(worksheet); row += 1) {
    const formulaCell = worksheet[`W${row}`];
    if (!formulaCell?.f) continue;

    const parsed = parseFormula(formulaCell.f);
    if (!parsed) continue;

    const household = householdByRow.get(row);
    const depositsForFormula = parsed.depositRows.map((depositRow) => depositByKey.get(`${parsed.source}:${depositRow}`) ?? null);
    const deposit = depositsForFormula[0] ?? null;
    const rawAmount = depositsForFormula.reduce((total, item, index) => total + ((item?.amount ?? 0) / parsed.divisors[index]), 0) || Number(formulaCell.v ?? 0);

    references.push({
      householdRow: row,
      householdName: household?.name ?? String(worksheet[`D${row}`]?.v ?? ''),
      householdAddress: household?.address ?? String(worksheet[`H${row}`]?.v ?? ''),
      householdAmount: household?.requiredAmount ?? Number(worksheet[`V${row}`]?.v ?? 0),
      formula: `=${String(formulaCell.f).replace(/^=/, '')}`,
      source: parsed.source,
      depositRow: parsed.depositRow,
      depositRows: parsed.depositRows,
      divisor: parsed.divisors[0],
      allocatedAmount: Math.round(rawAmount),
      deposit: depositsForFormula.length === 1 ? deposit : null,
    });
  }

  return references;
}

function isSameDeposit(reference: ExistingFormulaReference, result: MatchResult | undefined) {
  if (!result) return false;
  const resultKeys = new Set(result.candidates.map((candidate) => `${candidate.source}:${candidate.rowIndex}`));
  if (result.matchedDeposit) resultKeys.add(`${result.matchedDeposit.source}:${result.matchedDeposit.rowIndex}`);
  return reference.depositRows.every((row) => resultKeys.has(`${reference.source}:${row}`));
}

export function compareExistingMatches(
  references: ExistingFormulaReference[],
  households: HouseholdRow[],
  results: MatchResult[],
): { rows: ValidationReportRow[]; summary: ValidationSummary } {
  const resultByHouseholdRow = new Map(results.map((result, index) => [households[index]?.rowIndex, result]));
  const rows: ValidationReportRow[] = [];

  for (const reference of references) {
    const result = resultByHouseholdRow.get(reference.householdRow);
    const sameDeposit = isSameDeposit(reference, result);
    let comparison: ValidationComparison;
    let differenceReason = '';

    if (!result || result.status === 'UNMATCHED') {
      comparison = 'PROGRAM_MISSED';
      differenceReason = '기존 W열에는 입금 참조가 있으나 프로그램은 후보를 찾지 못했습니다.';
    } else if (!sameDeposit) {
      comparison = 'PROGRAM_DIFFERENT';
      differenceReason = '프로그램이 기존 W열과 다른 입금행을 추천 또는 선택했습니다.';
    } else if (reference.divisor > 1) {
      if (String(result.status) === 'SPLIT_MATCHED') {
        comparison = 'EXACT_SPLIT_MATCH';
        differenceReason = '기존 분할 입금과 동일한 입금행을 분할 후보로 처리했습니다.';
      } else {
        comparison = 'PROGRAM_REVIEW';
        differenceReason = '동일 입금행은 찾았지만 기존 분할 배분을 자동 재현하지 못했습니다.';
      }
    } else if (result.status === 'AUTO_MATCHED') {
      comparison = 'EXACT_MATCH';
    } else {
      comparison = 'PROGRAM_REVIEW';
      differenceReason = '동일 입금행을 찾았지만 자동확정하지 않고 확인필요로 분류했습니다.';
    }

    rows.push({
      householdRow: reference.householdRow,
      householdName: reference.householdName,
      householdAddress: reference.householdAddress,
      householdAmount: reference.householdAmount,
      formula: reference.formula,
      existingSource: reference.source === 'GIRO' ? '지로' : '우체국',
      existingDepositRow: reference.depositRows.join('+'),
      existingAllocatedAmount: reference.allocatedAmount,
      programStatus: result?.status ?? 'UNMATCHED',
      programSource: result?.matchedDeposit?.source ?? '',
      programDepositRow: result?.matchedDeposit?.rowIndex ?? '',
      programDepositorName: result?.matchedDeposit?.depositorName ?? '',
      programDepositAmount: result?.matchedDeposit?.amount ?? '',
      programScore: result?.score ?? '',
      comparison,
      differenceReason,
      programCandidates: (result?.candidates ?? []).map((candidate) => ({
        source: candidate.source,
        row: candidate.rowIndex,
        depositorName: candidate.depositorName,
        amount: candidate.amount,
        address: candidate.address ?? '',
        date: String(candidate.originalData['거래일시'] ?? candidate.originalData['납부일자'] ?? candidate.originalData['입금일시'] ?? '-'),
        score: result?.score ?? 0,
      })),
    });
  }

  const summary = {
    existingTotal: rows.length,
    existingGiro: rows.filter((row) => row.existingSource === '지로').length,
    existingPost: rows.filter((row) => row.existingSource === '우체국').length,
    existingSplit: references.filter((reference) => reference.divisor > 1).length,
    exactMatch: rows.filter((row) => row.comparison === 'EXACT_MATCH').length,
    exactSplitMatch: rows.filter((row) => row.comparison === 'EXACT_SPLIT_MATCH').length,
    programReview: rows.filter((row) => row.comparison === 'PROGRAM_REVIEW').length,
    programMissed: rows.filter((row) => row.comparison === 'PROGRAM_MISSED').length,
    programDifferent: rows.filter((row) => row.comparison === 'PROGRAM_DIFFERENT').length,
    noExistingMatch: households.length - rows.length,
    reproductionRate: rows.length ? (rows.filter((row) => row.comparison === 'EXACT_MATCH' || row.comparison === 'EXACT_SPLIT_MATCH').length / rows.length) * 100 : 0,
  };

  return { rows, summary };
}

export function auditAutoMatches(households: HouseholdRow[], results: MatchResult[]): SafetyAudit {
  const householdById = new Map(households.map((household) => [household.id, household]));
  const autoResults = results.filter((result) => result.status === 'AUTO_MATCHED');
  const nameCounts = new Map<string, number>();
  households.forEach((household) => nameCounts.set(household.normalizedName, (nameCounts.get(household.normalizedName) ?? 0) + 1));
  const depositCounts = new Map<string, number>();
  autoResults.forEach((result) => result.depositIds.forEach((id) => depositCounts.set(id, (depositCounts.get(id) ?? 0) + 1)));

  const duplicateNameAuto = autoResults.filter((result) => (nameCounts.get(householdById.get(result.householdId)?.normalizedName ?? '') ?? 0) > 1).length;
  const amountMismatchAuto = autoResults.filter((result) => result.matchedDeposit && result.householdAmount !== result.matchedDeposit.amount).length;
  const duplicateDepositAuto = autoResults.filter((result) => result.depositIds.some((id) => (depositCounts.get(id) ?? 0) > 1)).length;
  const fuzzyOnlyAuto = autoResults.filter((result) => result.matchedDeposit && householdById.get(result.householdId)?.normalizedName !== result.matchedDeposit.normalizedDepositorName).length;

  return {
    duplicateNameAuto,
    amountMismatchAuto,
    duplicateDepositAuto,
    fuzzyOnlyAuto,
    totalViolations: duplicateNameAuto + amountMismatchAuto + duplicateDepositAuto + fuzzyOnlyAuto,
  };
}
