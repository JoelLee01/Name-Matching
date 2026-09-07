const test = require('node:test');
const assert = require('node:assert/strict');
const { normalizeName, normalizeAmount, parseAddress, findBestNameMatch } = require('./normalize.js');

test('normalizeName collapses spacing and suffix text', () => {
  assert.equal(normalizeName(' 김철수 '), '김철수');
  assert.equal(normalizeName('김 철수'), '김철수');
  assert.equal(normalizeName('김철수(입금)'), '김철수');
  assert.equal(normalizeName('김철수_자부담'), '김철수');
});

test('parseAmount handles Korean currency and number formats', () => {
  assert.equal(normalizeAmount('1,500,000원'), 1500000);
  assert.equal(normalizeAmount('₩1,500,000'), 1500000);
  assert.equal(normalizeAmount('1500000.0'), 1500000);
});

test('findBestNameMatch only returns strong candidates', () => {
  assert.equal(findBestNameMatch('김철수', ['김철수', '박철수']).name, '김철수');
  assert.equal(findBestNameMatch('김용우', ['JINLONGYU', '김용우']).name, '김용우');
});
