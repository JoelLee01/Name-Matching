const { normalizeName, normalizeAddress, normalizeAmount, findBestNameMatch } = require('./normalize');

const STATUS_LABEL = {
  UNMATCHED: '미매칭',
  AUTO_MATCHED: '자동확정',
  REVIEW_REQUIRED: '확인필요',
  CONFIRMED: '사용자확정',
  IGNORED: '제외',
};

function normalizeHouseholdRow(row) {
  const name = row.name || row['세대주명'] || row['세대주'] || row['성명'] || row['신청자명'] || '';
  const amount = row.requiredAmount ?? row['자부담'] ?? row['자부담금'] ?? row['납부금액'] ?? row['금액'] ?? 0;
  const address = row.address || row['주소'] || row['세대주소'] || '';

  return {
    id: row.id || `household-${row.rowIndex || Math.random().toString(16).slice(2)}`,
    rowIndex: row.rowIndex ?? 0,
    name: String(name || '').trim(),
    requiredAmount: normalizeAmount(amount),
    address: String(address || '').trim(),
    normalizedName: normalizeName(name),
    normalizedAddress: normalizeAddress(address),
    originalData: row.originalData || row,
  };
}

function normalizeDepositRow(row, source) {
  const name = row.depositorName || row['입금자명'] || row['고객상호명'] || row['입금자'] || row['성명'] || row['고객명'] || '';
  const amount = row.amount ?? row['입금금액'] ?? row['납부금액'] ?? row['거래금액'] ?? row['금액'] ?? 0;
  const address = row.address || row['주소'] || row['설치주소'] || row['납부처주소'] || '';

  return {
    id: row.id || `${source.toLowerCase()}-${row.rowIndex || Math.random().toString(16).slice(2)}`,
    rowIndex: row.rowIndex ?? 0,
    source,
    depositorName: String(name || '').trim(),
    amount: normalizeAmount(amount),
    address: String(address || '').trim(),
    normalizedDepositorName: normalizeName(name),
    normalizedAddress: normalizeAddress(address),
    originalData: row.originalData || row,
  };
}

function buildDepositIndex(deposits) {
  const index = new Map();
  deposits.forEach((deposit) => {
    const key = deposit.normalizedDepositorName || '';
    if (!key) return;
    if (!index.has(key)) index.set(key, []);
    index.get(key).push(deposit);
  });
  return index;
}

function matchHouseholds(households, deposits) {
  const depositIndex = buildDepositIndex(deposits);
  const usedDepositIds = new Set();
  const sameNameHouseholds = new Map();

  households.forEach((household) => {
    const key = household.normalizedName;
    if (!key) return;
    sameNameHouseholds.set(key, (sameNameHouseholds.get(key) ?? 0) + 1);
  });

  const results = households.map((household) => {
    const duplicateHouseholdCount = sameNameHouseholds.get(household.normalizedName) ?? 0;
    const exactCandidates = (depositIndex.get(household.normalizedName) || []).filter(
      (deposit) => !usedDepositIds.has(deposit.id),
    );

    const allMatches = [...exactCandidates, ...deposits.filter((deposit) => {
      if (usedDepositIds.has(deposit.id)) return false;
      const similarity = findBestNameMatch(household.normalizedName, [deposit.normalizedDepositorName]);
      return similarity.score >= 70 && similarity.score < 100;
    })].filter((candidate, index, array) => array.findIndex((item) => item.id === candidate.id) === index);

    if (!allMatches.length) {
      return {
        householdId: household.id,
        householdName: household.name,
        depositIds: [],
        status: 'UNMATCHED',
        score: 0,
        reasons: ['입금내역 후보를 찾지 못했습니다.'],
        matchedDeposit: null,
      };
    }

    const exactNameMatches = allMatches.filter((deposit) => deposit.normalizedDepositorName === household.normalizedName);
    const amountMatches = allMatches.filter((deposit) => deposit.amount === household.requiredAmount);
    const addressMatches = allMatches.filter((deposit) => {
      if (!household.normalizedAddress || !deposit.normalizedAddress) return false;
      return deposit.normalizedAddress === household.normalizedAddress;
    });

    if (duplicateHouseholdCount > 1) {
      return {
        householdId: household.id,
        householdName: household.name,
        depositIds: allMatches.map((candidate) => candidate.id),
        status: 'REVIEW_REQUIRED',
        score: 75,
        reasons: ['동일 이름 신청세대가 여러 건 존재합니다.', '자동확정은 보류됩니다.'],
        matchedDeposit: allMatches[0],
      };
    }

    if (exactNameMatches.length > 1) {
      return {
        householdId: household.id,
        householdName: household.name,
        depositIds: exactNameMatches.map((candidate) => candidate.id),
        status: 'REVIEW_REQUIRED',
        score: 78,
        reasons: ['동일 이름 입금내역이 여러 건 존재합니다.', '중복 후보로 인해 확인이 필요합니다.'],
        matchedDeposit: exactNameMatches[0],
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
      };
    }

    if (amountMatches.length === 1 && amountMatches[0].normalizedDepositorName !== household.normalizedName) {
      return {
        householdId: household.id,
        householdName: household.name,
        depositIds: [amountMatches[0].id],
        status: 'REVIEW_REQUIRED',
        score: 72,
        reasons: ['금액은 일치하지만 이름이 확실하지 않습니다.'],
        matchedDeposit: amountMatches[0],
      };
    }

    if (addressMatches.length && household.normalizedAddress && allMatches[0].normalizedAddress) {
      const deposit = allMatches[0];
      usedDepositIds.add(deposit.id);
      return {
        householdId: household.id,
        householdName: household.name,
        depositIds: [deposit.id],
        status: 'AUTO_MATCHED',
        score: 96,
        reasons: ['이름 일치', '금액 일치', '주소 일치'],
        matchedDeposit: deposit,
      };
    }

    if (household.normalizedName !== allMatches[0].normalizedDepositorName) {
      return {
        householdId: household.id,
        householdName: household.name,
        depositIds: allMatches.map((candidate) => candidate.id),
        status: 'REVIEW_REQUIRED',
        score: 70,
        reasons: ['이름이 완전히 일치하지 않아 유사 후보만 발견되었습니다.'],
        matchedDeposit: allMatches[0],
      };
    }

    return {
      householdId: household.id,
      householdName: household.name,
      depositIds: allMatches.map((candidate) => candidate.id),
      status: 'REVIEW_REQUIRED',
      score: 68,
      reasons: ['여러 후보가 존재하거나 금액 조건이 불명확합니다.'],
      matchedDeposit: allMatches[0],
    };
  });

  return results;
}

module.exports = {
  STATUS_LABEL,
  normalizeHouseholdRow,
  normalizeDepositRow,
  buildDepositIndex,
  matchHouseholds,
};
