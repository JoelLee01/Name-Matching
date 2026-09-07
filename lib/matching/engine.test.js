const test = require('node:test');
const assert = require('node:assert/strict');
const { matchHouseholds } = require('./engine.js');

test('exact match auto confirms', () => {
  const households = [{ id: 'h1', name: '김철수', requiredAmount: 1500000, address: 'A주소', normalizedName: '김철수', normalizedAddress: 'a주소' }];
  const deposits = [{ id: 'd1', source: 'GIRO', depositorName: '김철수', amount: 1500000, address: 'A주소', normalizedDepositorName: '김철수', normalizedAddress: 'a주소' }];
  const result = matchHouseholds(households, deposits);
  assert.equal(result[0].status, 'AUTO_MATCHED');
  assert.ok(result[0].score >= 90);
});

test('duplicate exact candidate triggers review', () => {
  const households = [{ id: 'h1', name: '김철수', requiredAmount: 1500000, address: 'A주소', normalizedName: '김철수', normalizedAddress: 'a주소' }];
  const deposits = [
    { id: 'd1', source: 'GIRO', depositorName: '김철수', amount: 1500000, address: 'A주소', normalizedDepositorName: '김철수', normalizedAddress: 'a주소' },
    { id: 'd2', source: 'POST', depositorName: '김철수', amount: 1500000, address: 'B주소', normalizedDepositorName: '김철수', normalizedAddress: 'b주소' }
  ];
  const result = matchHouseholds(households, deposits);
  assert.equal(result[0].status, 'REVIEW_REQUIRED');
});
