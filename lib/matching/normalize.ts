export function normalizeName(value: string | number | null | undefined): string {
  if (value === null || value === undefined) return '';

  const withoutPunctuation = String(value)
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/[()\[\]{}]/g, ' ')
    .replace(/[_.\-/]/g, ' ')
    .replace(/(입금|자부담|납부|수납|세대주|주민|고객|상호|씨|님)$/gi, '')
    .replace(/[^가-힣a-zA-Z0-9\s]/g, ' ')
    .replace(/\s+/g, '')
    .trim();

  return withoutPunctuation.replace(/(입금|자부담|납부|수납|세대주|주민|고객|상호|씨|님)$/gi, '');
}

export function normalizeAddress(value: string | number | null | undefined): string {
  if (value === null || value === undefined) return '';
  return String(value)
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(/[^가-힣a-z0-9]/g, '')
    .trim();
}

export function parseAmount(value: unknown): number {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return Math.round(value);

  const text = String(value)
    .replace(/원|₩|,/g, '')
    .replace(/\s+/g, '')
    .trim();

  const num = Number.parseFloat(text);
  return Number.isFinite(num) ? Math.round(num) : 0;
}

export const normalizeAmount = parseAmount;

export function findBestNameMatch(target: string, candidates: string[]): { name: string; score: number; exact: boolean } {
  const normalizedTarget = normalizeName(target);
  if (!normalizedTarget) return { name: '', score: 0, exact: false };

  const ranked = candidates
    .map((candidate) => {
      const normalized = normalizeName(candidate);
      if (!normalized) return null;

      let score = 0;
      if (normalized === normalizedTarget) score = 100;
      else if (normalized.includes(normalizedTarget) || normalizedTarget.includes(normalized)) score = 80;
      else {
        const sameLength = Math.max(normalizedTarget.length, normalized.length);
        const shared = [...normalizedTarget].filter((char) => normalized.includes(char)).length;
        const ratio = sameLength ? shared / sameLength : 0;
        score = ratio > 0.7 ? 70 : 0;
      }

      return { name: candidate, score, exact: normalized === normalizedTarget };
    })
    .filter(Boolean) as { name: string; score: number; exact: boolean }[];

  return ranked.sort((a, b) => b.score - a.score)[0] || { name: '', score: 0, exact: false };
}
