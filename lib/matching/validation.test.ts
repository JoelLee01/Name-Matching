import { describe, expect, it } from 'vitest';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as XLSX from 'xlsx';
import { matchHouseholds, normalizeDepositRow, normalizeHouseholdRow } from './engine.ts';
import { getRowsForRole } from './workbook';
import { compareExistingMatches, readExistingFormulaReferences } from './validation';
import { buildSplitCandidates, validateAllocations } from './allocation';
import { buildResultWorkbook } from './export';

describe('real workbook validation', () => {
  it('parses W formulas and compares the actual workbook data', () => {
    const filePath = path.join(process.cwd(), '260902_창녕 영산면_신청세대명단(자부담확인)_R1.xlsx');
    expect(fs.existsSync(filePath)).toBe(true);

    const workbook = XLSX.read(fs.readFileSync(filePath), { type: 'buffer', cellFormula: true, cellNF: true });
    const householdRows = getRowsForRole(workbook, 'HOUSEHOLD');
    const giroRows = getRowsForRole(workbook, 'GIRO');
    const postRows = getRowsForRole(workbook, 'POST_OFFICE');
    const households = householdRows.map((row, index) => normalizeHouseholdRow({ ...row, id: `household-${index + 1}`, rowIndex: Number(row.__sourceRow ?? index + 1) }));
    const deposits = [
      ...giroRows.map((row, index) => normalizeDepositRow({ ...row, id: `giro-${index + 1}`, rowIndex: Number(row.__sourceRow ?? index + 1) }, 'GIRO')),
      ...postRows.map((row, index) => normalizeDepositRow({ ...row, id: `post-${index + 1}`, rowIndex: Number(row.__sourceRow ?? index + 1) }, 'POST')),
    ];
    const results = matchHouseholds(households, deposits);
    const references = readExistingFormulaReferences(workbook, households, deposits);
    const report = compareExistingMatches(references, households, results);
    const splitCandidates = buildSplitCandidates(references, households, deposits);

    expect(workbook.SheetNames).toContain('신청세대명단');
    expect(workbook.SheetNames).toContain('지로');
    expect(workbook.SheetNames).toContain('우체국');
    expect(households.length).toBe(798);
    expect(giroRows.length).toBe(169);
    expect(postRows.length).toBe(203);
    expect(references.length).toBe(355);
    expect(references.filter((reference) => reference.divisor > 1)).toHaveLength(14);
    expect(references.some((reference) => reference.formula === '=우체국!D201/2')).toBe(true);
    expect(report.summary.existingTotal).toBe(355);
    expect(report.summary.exactMatch).toBe(170);
    expect(report.summary.exactSplitMatch).toBe(0);
    expect(report.summary.programReview).toBe(170);
    expect(report.summary.programMissed).toBe(12);
    expect(report.summary.programDifferent).toBe(3);
    expect(splitCandidates).toHaveLength(5);
    const twoHouseholdSplit = splitCandidates.find((candidate) => candidate.formula === '=우체국!D201/2');
    const sixHouseholdSplit = splitCandidates.find((candidate) => candidate.formula === '=우체국!D121/6');
    expect(twoHouseholdSplit?.households).toHaveLength(2);
    expect(twoHouseholdSplit ? validateAllocations(twoHouseholdSplit.deposit.amount, twoHouseholdSplit.allocations).valid : false).toBe(true);
    expect(sixHouseholdSplit?.households).toHaveLength(6);
    expect(sixHouseholdSplit ? validateAllocations(sixHouseholdSplit.deposit.amount, sixHouseholdSplit.allocations).valid : false).toBe(true);

    const approvedResults = results.map((result) => {
      const allocation = splitCandidates.flatMap((candidate) => candidate.allocations).find((item) => item.householdId === result.householdId);
      if (!allocation) return result;
      const deposit = deposits.find((item) => item.id === allocation.depositId) ?? null;
      return { ...result, status: 'SPLIT_MATCHED' as const, allocations: [allocation], matchedDeposit: deposit, candidates: deposit ? [deposit] : [] };
    });
    const outputPath = path.join(process.cwd(), 'tmp', 'split-export-reopen.xlsx');
    const exported = buildResultWorkbook(workbook, approvedResults, report.rows);
    XLSX.writeFile(exported, outputPath);
    const reopened = XLSX.read(fs.readFileSync(outputPath), { type: 'buffer', cellFormula: true });
    expect(reopened.SheetNames).toEqual(expect.arrayContaining([...workbook.SheetNames, '입금매칭결과', '입금배분결과', '기존매칭_검증']));
    expect(reopened.Sheets['입금배분결과']['!ref']).toBe('A1:H15');
    expect(reopened.Sheets['기존매칭_검증']['!ref']).toBe('A1:P356');
    expect(reopened.Sheets['신청세대명단']['W6'].f).toBe('우체국!D201/2');
    expect(reopened.Sheets['신청세대명단']['W534'].f).toBe('우체국!D113/2');
  });
});
