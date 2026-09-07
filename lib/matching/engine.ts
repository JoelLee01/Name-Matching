import { findBestNameMatch, normalizeAddress, normalizeName, parseAmount } from './normalize';
import type { DepositRow, HouseholdRow, MatchResult } from './types';

function resolveRowValue(row: Record<string, unknown>, aliases: string[]) {
  const keys = Object.keys(row);
  const normalizeKey = (value: string) => String(value).replace(/^\uFEFF/, '').trim().toLowerCase().replace(/\s+/g, '').replace(/[^가-힣a-z0-9]/g, '');

  for (const alias of aliases) {
    const normalizedAlias = normalizeKey(alias);
    const match = keys.find((key) => {
      const normalizedKey = normalizeKey(key);
      return normalizedKey === normalizedAlias || normalizedKey.includes(normalizedAlias) || normalizedAlias.includes(normalizedKey);
    });

    if (match && row[match] !== undefined && row[match] !== null && String(row[match]).trim() !== '') {
      return row[match];
    }
  }

  return undefined;
}
export function normalizeHouseholdRow(row: Record<string, unknown>): HouseholdRow {
  const nameValue = row.name || resolveRowValue(row, ['세대주명', '세대주', '성명', '신청자명', '이름']) || '';
  const amountValue = row.requiredAmount || resolveRowValue(row, ['자부담', '자부담금', '납부금액', '금액']) || 0;
  const addressValue = row.address || resolveRowValue(row, ['주소', '세대주소']) || '';

  const name = String(nameValue ?? '').trim();
  const address = String(addressValue ?? '').trim();

  return {
    id: String(row.id ?? `household-${Math.random().toString(16).slice(2)}`),
    rowIndex: Number(row.rowIndex ?? 0),
    name,
    requiredAmount: parseAmount(amountValue),
    address,
    normalizedName: normalizeName(name),
    normalizedAddress: normalizeAddress(address),
    originalData: row,
  };
}

export function normalizeDepositRow(row: Record<string, unknown>, source: 'GIRO' | 'POST'): DepositRow {
  const nameValue = row.depositorName || resolveRowValue(row, ['입금자명', '고객상호명', '고객명', '입금자', '성명', '내역']) || '';
  const amountValue = row.amount || resolveRowValue(row, ['입금금액', '입금액원', '납부금액', '거래금액', '금액']) || 0;
  const addressValue = row.address || resolveRowValue(row, ['주소', '설치주소', '납부처주소']) || '';

  const name = String(nameValue ?? '').trim();
  const address = String(addressValue ?? '').trim();

  return {
    id: String(row.id ?? `${source.toLowerCase()}-${Math.random().toString(16).slice(2)}`),
    rowIndex: Number(row.rowIndex ?? 0),
    source,
    depositorName: name,
    amount: parseAmount(amountValue),
    address,
    normalizedDepositorName: normalizeName(name),
    normalizedAddress: normalizeAddress(address),
    originalData: row,
  };
}

export function buildDepositIndex(deposits: DepositRow[]) {
  return deposits.reduce<Map<string, DepositRow[]>>((acc, deposit) => {
    const key = deposit.normalizedDepositorName;
    if (!key) return acc;
    const list = acc.get(key) ?? [];
    list.push(deposit);
    acc.set(key, list);
    return acc;
  }, new Map());
}

export function matchHouseholds(households: HouseholdRow[], deposits: DepositRow[]): MatchResult[] {
  const depositIndex = buildDepositIndex(deposits);
  const usedDepositIds = new Set<string>();
  const sameNameHouseholds = new Map<string, number>();

  households.forEach((household) => {
    const key = household.normalizedName;
    if (!key) return;
    sameNameHouseholds.set(key, (sameNameHouseholds.get(key) ?? 0) + 1);
  });

  return households.map((household) => {
    const duplicateHouseholdCount = sameNameHouseholds.get(household.normalizedName) ?? 0;
    const exactCandidates = (depositIndex.get(household.normalizedName) ?? []).filter((d) => !usedDepositIds.has(d.id));
    const similarCandidates = deposits
      .filter((d) => !usedDepositIds.has(d.id))
      .filter((d) => {
        const similarity = findBestNameMatch(household.normalizedName, [d.normalizedDepositorName]);
        return similarity.score >= 70 && similarity.score < 100;
      });

    const allCandidates = [...exactCandidates, ...similarCandidates].filter(
      (candidate, index, array) => array.findIndex((item) => item.id === candidate.id) === index,
    );

    if (!allCandidates.length) {
      return {
        householdId: household.id,
        householdName: household.name,
        depositIds: [],
        status: 'UNMATCHED',
        score: 0,
        reasons: ['입금내역 후보를 찾지 못했습니다.'],
        matchedDeposit: null,
        candidates: [],
        householdAddress: household.address,
        householdAmount: household.requiredAmount,
      };
    }

    const exactNameMatches = allCandidates.filter((d) => d.normalizedDepositorName === household.normalizedName);
    const exactAmountMatches = allCandidates.filter((d) => d.amount === household.requiredAmount);
    const sameNameDepositCount = exactNameMatches.length;
    const addressMatches = allCandidates.filter((d) => {
      if (!household.normalizedAddress || !d.normalizedAddress) return false;
      return d.normalizedAddress === household.normalizedAddress;
    });

    if (duplicateHouseholdCount > 1) {
      return {
        householdId: household.id,
        householdName: household.name,
        depositIds: allCandidates.map((candidate) => candidate.id),
        status: 'REVIEW_REQUIRED',
        score: 75,
        reasons: ['동일 이름 신청세대가 여러 건 존재합니다.', '자동확정은 보류됩니다.'],
        matchedDeposit: allCandidates[0] ?? null,
        candidates: allCandidates,
        householdAddress: household.address,
        householdAmount: household.requiredAmount,
      };
    }

    if (sameNameDepositCount > 1) {
      return {
        householdId: household.id,
        householdName: household.name,
        depositIds: exactNameMatches.map((candidate) => candidate.id),
        status: 'REVIEW_REQUIRED',
        score: 78,
        reasons: ['동일 이름 입금내역이 여러 건 존재합니다.', '중복 후보로 인해 확인이 필요합니다.'],
        matchedDeposit: exactNameMatches[0] ?? null,
        candidates: exactNameMatches,
        householdAddress: household.address,
        householdAmount: household.requiredAmount,
      };
    }

    const exactMatch = exactNameMatches.find((candidate) => candidate.amount === household.requiredAmount && !usedDepositIds.has(candidate.id));
    if (exactMatch) {
      usedDepositIds.add(exactMatch.id);
      return {
        householdId: household.id,
        householdName: household.name,
        depositIds: [exactMatch.id],
        status: 'AUTO_MATCHED',
        score: 100,
        reasons: ['이름 정확히 일치', '입금금액 일치', '중복 후보 없음'],
        matchedDeposit: exactMatch,
        candidates: [exactMatch],
        householdAddress: household.address,
        householdAmount: household.requiredAmount,
      };
    }

    if (exactNameMatches.length === 1 && household.requiredAmount !== exactNameMatches[0].amount) {
      return {
        householdId: household.id,
        householdName: household.name,
        depositIds: [exactNameMatches[0].id],
        status: 'REVIEW_REQUIRED',
        score: 65,
        reasons: ['이름은 일치하지만 입금금액이 다릅니다.'],
        matchedDeposit: exactNameMatches[0],
        candidates: exactNameMatches,
        householdAddress: household.address,
        householdAmount: household.requiredAmount,
      };
    }

    if (exactAmountMatches.length === 1 && exactAmountMatches[0].normalizedDepositorName !== household.normalizedName) {
      return {
        householdId: household.id,
        householdName: household.name,
        depositIds: [exactAmountMatches[0].id],
        status: 'REVIEW_REQUIRED',
        score: 72,
        reasons: ['금액은 일치하지만 이름이 확실하지 않습니다.'],
        matchedDeposit: exactAmountMatches[0],
        candidates: exactAmountMatches,
        householdAddress: household.address,
        householdAmount: household.requiredAmount,
      };
    }

    if (addressMatches.length && household.normalizedAddress && allCandidates[0].normalizedAddress) {
      const deposit = allCandidates[0];
      usedDepositIds.add(deposit.id);
      return {
        householdId: household.id,
        householdName: household.name,
        depositIds: [deposit.id],
        status: 'AUTO_MATCHED',
        score: 96,
        reasons: ['이름 일치', '금액 일치', '주소 일치'],
        matchedDeposit: deposit,
        candidates: [deposit],
        householdAddress: household.address,
        householdAmount: household.requiredAmount,
      };
    }

    if (household.normalizedName !== allCandidates[0].normalizedDepositorName) {
      return {
        householdId: household.id,
        householdName: household.name,
        depositIds: allCandidates.map((candidate) => candidate.id),
        status: 'REVIEW_REQUIRED',
        score: 70,
        reasons: ['이름이 완전히 일치하지 않아 유사 후보만 발견되었습니다.'],
        matchedDeposit: allCandidates[0],
        candidates: allCandidates,
        householdAddress: household.address,
        householdAmount: household.requiredAmount,
      };
    }

    return {
      householdId: household.id,
      householdName: household.name,
      depositIds: allCandidates.map((candidate) => candidate.id),
      status: 'REVIEW_REQUIRED',
      score: 68,
      reasons: ['여러 후보가 존재하거나 금액 조건이 불명확합니다.'],
      matchedDeposit: allCandidates[0] ?? null,
      candidates: allCandidates,
      householdAddress: household.address,
      householdAmount: household.requiredAmount,
    };
  });
}
