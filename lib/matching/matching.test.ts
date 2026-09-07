import { describe, expect, it } from 'vitest';
import { matchHouseholds, normalizeHouseholdRow, normalizeDepositRow } from './engine.ts';
import { normalizeName, parseAmount } from './normalize';


describe('matching engine', () => {
  it('exact name and amount become auto match', () => {
    const households = [{
      id: 'h1',
      rowIndex: 1,
      name: '김철수',
      requiredAmount: 1500000,
      address: 'A주소',
      normalizedName: '김철수',
      normalizedAddress: 'a주소',
      originalData: {},
    }];

    const deposits = [{
      id: 'd1',
      rowIndex: 1,
      source: 'GIRO' as const,
      depositorName: '김철수',
      amount: 1500000,
      address: 'A주소',
      normalizedDepositorName: '김철수',
      normalizedAddress: 'a주소',
      originalData: {},
    }];

    const result = matchHouseholds(households, deposits);
    expect(result[0].status).toBe('AUTO_MATCHED');
  });

  it('duplicate household name requires review', () => {
    const households = [
      { id: 'h1', rowIndex: 1, name: '김철수', requiredAmount: 1500000, address: 'A주소', normalizedName: '김철수', normalizedAddress: 'a주소', originalData: {} },
      { id: 'h2', rowIndex: 2, name: '김철수', requiredAmount: 1500000, address: 'B주소', normalizedName: '김철수', normalizedAddress: 'b주소', originalData: {} },
    ];

    const deposits = [{
      id: 'd1', rowIndex: 1, source: 'POST' as const, depositorName: '김철수', amount: 1500000, address: 'A주소', normalizedDepositorName: '김철수', normalizedAddress: 'a주소', originalData: {} }];

    const result = matchHouseholds(households, deposits);
    expect(result[0].status).toBe('REVIEW_REQUIRED');
    expect(result[1].status).toBe('REVIEW_REQUIRED');
  });

  it('amount mismatch stays review required', () => {
    const households = [{ id: 'h1', rowIndex: 1, name: '이영희', requiredAmount: 1500000, address: 'A주소', normalizedName: '이영희', normalizedAddress: 'a주소', originalData: {} }];
    const deposits = [{ id: 'd1', rowIndex: 1, source: 'POST' as const, depositorName: '이영희', amount: 1000000, address: 'A주소', normalizedDepositorName: '이영희', normalizedAddress: 'a주소', originalData: {} }];

    const result = matchHouseholds(households, deposits);
    expect(result[0].status).toBe('REVIEW_REQUIRED');
  });

  it('normalizeName strips spaces and suffix text', () => {
    expect(normalizeName(' 김철수 ')).toBe('김철수');
    expect(normalizeName('김 철수')).toBe('김철수');
    expect(normalizeName('김철수(입금)')).toBe('김철수');
    expect(normalizeName('김철수_자부담')).toBe('김철수');
  });

  it('parseAmount handles various formatted amounts', () => {
    expect(parseAmount('1,500,000원')).toBe(1500000);
    expect(parseAmount('₩1,500,000')).toBe(1500000);
    expect(parseAmount('1500000.0')).toBe(1500000);
  });

  it('matches rows when Excel headers include whitespace, BOM, or formatting noise', () => {
    const households = [
      normalizeHouseholdRow({
        '﻿세대주명 ': '김철수',
        ' 자부담금 ': 1500000,
        ' 주소 ': '수원시 영통구',
      }),
    ];

    const deposits = [
      normalizeDepositRow({
        '입금자명 ': '김철수',
        '입금금액 ': 1500000,
        ' 주소 ': '수원시 영통구',
      }, 'GIRO'),
    ];

    const result = matchHouseholds(households, deposits);
    expect(result[0].status).toBe('AUTO_MATCHED');
  });
});
