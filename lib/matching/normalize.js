function normalizeName(value) {
  if (value === null || value === undefined) return '';

  const cleaned = String(value)
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/[()\[\]{}]/g, ' ')
    .replace(/[_.\-/]/g, ' ')
    .replace(/(입금|자부담|납부|수납|세대주|주민|고객|상호|씨|님)$/gi, '')
    .replace(/[^가-힣a-zA-Z0-9\s]/g, ' ')
    .replace(/\s+/g, '')
    .trim();

  return cleaned.replace(/(입금|자부담|납부|수납|세대주|주민|고객|상호|씨|님)$/gi, '');
}

function normalizeAddress(value) {
  if (value === null || value === undefined) return '';
  return String(value)
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(/[^가-힣a-z0-9]/g, '')
    .trim();
}

function normalizeAmount(value) {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return Math.round(value);

  const text = String(value)
    .replace(/원|₩|,/g, '')
    .replace(/\s+/g, '')
    .trim();

  const num = Number.parseFloat(text);
  if (Number.isFinite(num)) return Math.round(num);

  return 0;
}

function findBestNameMatch(value, candidates = []) {
  const target = normalizeName(value);
  if (!target) return { name: '', score: 0, exact: false };

  const results = candidates
    .map((candidate) => {
      const normalized = normalizeName(candidate);
      if (!normalized) return null;

      let score = 0;
      if (normalized === target) score = 100;
      else if (normalized.includes(target) || target.includes(normalized)) score = 80;
      else {
        const sameLength = Math.max(target.length, normalized.length);
        const shared = [...target].filter((char) => normalized.includes(char)).length;
        const ratio = sameLength ? shared / sameLength : 0;
        score = ratio > 0.7 ? 70 : 0;
      }

      return { name: candidate, normalized, score, exact: normalized === target };
    })
    .filter(Boolean)
    .sort((a, b) => b.score - a.score);

  return results[0] || { name: '', score: 0, exact: false };
}

module.exports = {
  normalizeName,
  normalizeAddress,
  normalizeAmount,
  parseAmount: normalizeAmount,
  findBestNameMatch,
};
