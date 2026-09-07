'use client';

import { Fragment, useEffect, useMemo, useState } from 'react';
import * as XLSX from 'xlsx';
import {
  AlertTriangle,
  CheckCircle2,
  Download,
  FileSpreadsheet,
  Search,
  ShieldAlert,
  Upload,
  X,
} from 'lucide-react';
import { matchHouseholds, normalizeDepositRow, normalizeHouseholdRow } from '@/lib/matching/engine.ts';
import { analyzeWorkbook, getRowsForRole, getSheetRoleLabel, type SheetAnalysis } from '@/lib/matching/workbook';
import type { DepositRow, HouseholdRow, MatchStatus } from '@/lib/matching/types';
import { auditAutoMatches, compareExistingMatches, readExistingFormulaReferences, type SafetyAudit, type ValidationReportRow, type ValidationSummary } from '@/lib/matching/validation';
import { buildSplitCandidates, validateAllocations, type SplitCandidate } from '@/lib/matching/allocation';
import { buildResultWorkbook } from '@/lib/matching/export';
import { normalizeAddress, normalizeName } from '@/lib/matching/normalize';

type FileKind = 'household' | 'giro' | 'post';

type UploadedFile = {
  kind: FileKind;
  file: File;
  name: string;
  rows: number;
};

type ReviewQueueFilter = 'ALL' | 'DUPLICATE_NAME' | 'AMOUNT' | 'FUZZY' | 'OTHER';
type ValidationFilter = 'ALL' | 'PROGRAM_MISSED' | 'PROGRAM_DIFFERENT';

function getReviewCategory(row: { reasons: string[] }): Exclude<ReviewQueueFilter, 'ALL'> {
  const reason = row.reasons.join(' ');
  if (reason.includes('동일 이름')) return 'DUPLICATE_NAME';
  if (reason.includes('금액')) return 'AMOUNT';
  if (reason.includes('유사') || reason.includes('완전히 일치')) return 'FUZZY';
  return 'OTHER';
}

function getCandidateDate(candidate: DepositRow) {
  return String(candidate.originalData['거래일시'] ?? candidate.originalData['납부일자'] ?? candidate.originalData['입금일시'] ?? '-');
}

function getCandidateScore(household: { name: string; address?: string; amount: number }, candidate: DepositRow) {
  const nameScore = normalizeName(household.name) === candidate.normalizedDepositorName ? 100 : 70;
  const amountScore = household.amount === candidate.amount ? 100 : 0;
  const addressScore = normalizeAddress(household.address) && candidate.normalizedAddress && normalizeAddress(household.address) === candidate.normalizedAddress ? 100 : 0;
  return Math.round(nameScore * 0.5 + amountScore * 0.3 + addressScore * 0.2);
}

const aliasMap: Record<string, string[]> = {
  householdName: ['세대주명', '세대주', '성명', '신청자명', '이름'],
  householdAddress: ['주소', '세대주소', '연락처주소'],
  householdAmount: ['자부담', '자부담금', '금액', '납부금액'],
  depositorName: ['입금자명', '고객상호명', '고객명', '입금자', '성명'],
  depositAmount: ['입금금액', '거래금액', '납부금액', '금액'],
  depositAddress: ['주소', '설치주소', '납부처주소'],
};

function normalizeHeader(header: string) {
  return String(header)
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(/[^가-힣a-z0-9]/g, '');
}

function findColumn(headers: string[], aliases: string[]) {
  const target = aliases.map(normalizeHeader);
  return headers.find((header) => target.includes(normalizeHeader(header))) ?? null;
}

async function readExcelRows(file: File) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array' });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: '', raw: false });
  const headers = rows.length ? Object.keys(rows[0]) : [];
  return { rows, headers, workbook };
}

function detectMatchingData(rows: Record<string, unknown>[], kind: 'household'): HouseholdRow[];
function detectMatchingData(rows: Record<string, unknown>[], kind: 'giro' | 'post'): DepositRow[];
function detectMatchingData(rows: Record<string, unknown>[], kind: FileKind): HouseholdRow[] | DepositRow[] {
  const headers = rows.length ? Object.keys(rows[0]) : [];
  if (kind === 'household') {
    return rows.map((row, index) => {
      const nameKey = findColumn(headers, aliasMap.householdName);
      const amountKey = findColumn(headers, aliasMap.householdAmount);
      const addressKey = findColumn(headers, aliasMap.householdAddress);
      return normalizeHouseholdRow({
        ...row,
        id: `household-${index + 1}`,
        rowIndex: index + 1,
        name: nameKey ? row[nameKey] : '',
        requiredAmount: amountKey ? row[amountKey] : 0,
        address: addressKey ? row[addressKey] : '',
        originalData: row,
      });
    });
  }

  const depositSource = kind === 'giro' ? 'GIRO' : 'POST';
  return rows.map((row, index) => {
    const nameKey = findColumn(headers, aliasMap.depositorName);
    const amountKey = findColumn(headers, aliasMap.depositAmount);
    const addressKey = findColumn(headers, aliasMap.depositAddress);
    return normalizeDepositRow(
      {
        ...row,
        id: `${depositSource.toLowerCase()}-${index + 1}`,
        rowIndex: index + 1,
        depositorName: nameKey ? row[nameKey] : '',
        amount: amountKey ? row[amountKey] : 0,
        address: addressKey ? row[addressKey] : '',
        originalData: row,
      },
      depositSource,
    );
  });
}

export default function Home() {
  const [files, setFiles] = useState<Record<FileKind, UploadedFile | null>>({
    household: null,
    giro: null,
    post: null,
  });
  const [processing, setProcessing] = useState(false);
  const [error, setError] = useState('');
  const [results, setResults] = useState<any[]>([]);
  const [filter, setFilter] = useState<'ALL' | MatchStatus>('ALL');
  const [search, setSearch] = useState('');
  const [sortKey, setSortKey] = useState<'householdName' | 'amount' | 'status' | 'score'>('householdName');
  const [sheetAnalysis, setSheetAnalysis] = useState<SheetAnalysis[]>([]);
  const [validationRows, setValidationRows] = useState<ValidationReportRow[]>([]);
  const [validationSummary, setValidationSummary] = useState<ValidationSummary | null>(null);
  const [safetyAudit, setSafetyAudit] = useState<SafetyAudit | null>(null);
  const [splitCandidates, setSplitCandidates] = useState<SplitCandidate[]>([]);
  const [sourceWorkbook, setSourceWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [reviewQueueFilter, setReviewQueueFilter] = useState<ReviewQueueFilter>('ALL');
  const [reviewPage, setReviewPage] = useState(0);
  const [activeReviewId, setActiveReviewId] = useState<string | null>(null);
  const [validationFilter, setValidationFilter] = useState<ValidationFilter>('ALL');
  const [datasetStats, setDatasetStats] = useState({
    household: 0,
    giro: 0,
    post: 0,
    totalDeposits: 0,
    auto: 0,
    review: 0,
    unmatched: 0,
    usedDeposits: 0,
    unusedDeposits: 0,
  });

  const handleFileUpload = async (kind: FileKind, file: File | null) => {
    if (!file) {
      setFiles((prev) => ({ ...prev, [kind]: null }));
      return;
    }

    const allowed = ['xlsx', 'xls', 'csv'];
    const ext = file.name.split('.').pop()?.toLowerCase() ?? '';
    if (!allowed.includes(ext)) {
      setError('엑셀 또는 CSV 파일만 업로드할 수 있습니다.');
      return;
    }

    const rows = await readExcelRows(file);
    const workbookAnalysis = analyzeWorkbook(rows.workbook);
    setFiles((prev) => ({
      ...prev,
      [kind]: {
        kind,
        file,
        name: file.name,
        rows: rows.rows.length,
      },
    }));
    setSheetAnalysis(workbookAnalysis);
    setError('');
  };

  const startMatching = async () => {
    if (!files.household) {
      setError('신청세대명단을 먼저 업로드해 주세요.');
      return;
    }

    setProcessing(true);
    setError('');

    try {
      const householdWorkbook = (await readExcelRows(files.household.file)).workbook;
      setSourceWorkbook(householdWorkbook);
      const householdSourceRows: Record<string, unknown>[] = getRowsForRole(householdWorkbook, 'HOUSEHOLD');
      const householdData: HouseholdRow[] = householdSourceRows.length
        ? householdSourceRows.map((row, index) => normalizeHouseholdRow({ ...row, id: `household-${index + 1}`, rowIndex: Number(row.__sourceRow ?? index + 1) }))
        : detectMatchingData((await readExcelRows(files.household.file)).rows, 'household');

      const depositRows: DepositRow[] = [];
      if (files.giro) {
        const giroWorkbook = (await readExcelRows(files.giro.file)).workbook;
        const giroRows: Record<string, unknown>[] = getRowsForRole(giroWorkbook, 'GIRO');
        if (giroRows.length) {
          depositRows.push(...giroRows.map((row, index) => normalizeDepositRow({ ...row, id: `giro-${index + 1}`, rowIndex: Number(row.__sourceRow ?? index + 1) }, 'GIRO')));
        } else {
          depositRows.push(...detectMatchingData((await readExcelRows(files.giro.file)).rows, 'giro'));
        }
      } else {
        const giroRows: Record<string, unknown>[] = getRowsForRole(householdWorkbook, 'GIRO');
        depositRows.push(...giroRows.map((row, index) => normalizeDepositRow({ ...row, id: `giro-${index + 1}`, rowIndex: Number(row.__sourceRow ?? index + 1) }, 'GIRO')));
      }
      if (files.post) {
        const postWorkbook = (await readExcelRows(files.post.file)).workbook;
        const postRows: Record<string, unknown>[] = getRowsForRole(postWorkbook, 'POST_OFFICE');
        if (postRows.length) {
          depositRows.push(...postRows.map((row, index) => normalizeDepositRow({ ...row, id: `post-${index + 1}`, rowIndex: Number(row.__sourceRow ?? index + 1) }, 'POST')));
        } else {
          depositRows.push(...detectMatchingData((await readExcelRows(files.post.file)).rows, 'post'));
        }
      } else {
        const postRows: Record<string, unknown>[] = getRowsForRole(householdWorkbook, 'POST_OFFICE');
        depositRows.push(...postRows.map((row, index) => normalizeDepositRow({ ...row, id: `post-${index + 1}`, rowIndex: Number(row.__sourceRow ?? index + 1) }, 'POST')));
      }

      const matches = matchHouseholds(householdData, depositRows);
      const existingReferences = readExistingFormulaReferences(householdWorkbook, householdData, depositRows);
      const validation = compareExistingMatches(existingReferences, householdData, matches);
      const allUsedDepositIds = new Set<string>(matches.flatMap((row) => row.depositIds ?? []));
      const usedDeposits = allUsedDepositIds.size;
      setResults(matches);
      setValidationRows(validation.rows);
      setValidationSummary(validation.summary);
      setSafetyAudit(auditAutoMatches(householdData, matches));
      setSplitCandidates(buildSplitCandidates(existingReferences, householdData, depositRows));
      setDatasetStats({
        household: householdData.length,
        giro: depositRows.filter((deposit) => deposit.source === 'GIRO').length,
        post: depositRows.filter((deposit) => deposit.source === 'POST').length,
        totalDeposits: depositRows.length,
        auto: matches.filter((row) => row.status === 'AUTO_MATCHED').length,
        review: matches.filter((row) => row.status === 'REVIEW_REQUIRED').length,
        unmatched: matches.filter((row) => row.status === 'UNMATCHED').length,
        usedDeposits,
        unusedDeposits: Math.max(depositRows.length - usedDeposits, 0),
      });
    } catch (err) {
      setError(err instanceof Error ? err.message : '처리 중 오류가 발생했습니다.');
    } finally {
      setProcessing(false);
    }
  };

  const summary = useMemo(() => {
    const total = results.length;
    const auto = results.filter((row) => row.status === 'AUTO_MATCHED').length;
    const review = results.filter((row) => row.status === 'REVIEW_REQUIRED').length;
    const unmatched = results.filter((row) => row.status === 'UNMATCHED').length;
    const split = results.filter((row) => row.status === 'SPLIT_MATCHED').length;
    return { total, auto, review, unmatched, split };
  }, [results]);

  const filteredResults = useMemo(() => {
    const rows = [...results].filter((row) => {
      const passStatus = filter === 'ALL' || row.status === filter;
      const haystack = [row.householdName, row.matchedDeposit?.depositorName ?? '', row.matchedDeposit?.address ?? '']
        .join(' ')
        .toLowerCase();
      const passSearch = !search || haystack.includes(search.toLowerCase());
      return passStatus && passSearch;
    });

    rows.sort((a, b) => {
      if (sortKey === 'householdName') return (a.householdName ?? '').localeCompare(b.householdName ?? '');
      if (sortKey === 'amount') return (b.householdAmount ?? 0) - (a.householdAmount ?? 0);
      if (sortKey === 'status') return (a.status ?? '').localeCompare(b.status ?? '');
      return (b.score ?? 0) - (a.score ?? 0);
    });

    return rows;
  }, [filter, results, search, sortKey]);

  const reviewQueueRows = useMemo(() => results.filter((row: any) => row.status === 'REVIEW_REQUIRED' && (reviewQueueFilter === 'ALL' || getReviewCategory(row) === reviewQueueFilter)), [results, reviewQueueFilter]);
  const reviewItems = useMemo(() => {
    const pageSize = 20;
    return reviewQueueRows.slice(reviewPage * pageSize, (reviewPage + 1) * pageSize);
  }, [reviewPage, reviewQueueRows]);
  const reviewTotal = results.filter((row: any) => row.status === 'REVIEW_REQUIRED').length;
  const reviewFilteredTotal = reviewQueueRows.length;
  const reviewPageCount = Math.max(Math.ceil(reviewFilteredTotal / 20), 1);
  const filteredValidationRows = validationRows.filter((row) => validationFilter === 'ALL' || row.comparison === validationFilter);

  const handleDecision = (householdId: string, nextStatus: 'CONFIRMED' | 'UNMATCHED') => {
    setResults((prev) => prev.map((row) => (row.householdId === householdId ? { ...row, status: nextStatus } : row)));
  };

  const moveToNextReview = (householdId: string) => {
    const currentIndex = reviewQueueRows.findIndex((row) => row.householdId === householdId);
    const nextIndex = currentIndex + 1;
    const next = reviewQueueRows[nextIndex];
    setActiveReviewId(next?.householdId ?? null);
    if (next) setReviewPage(Math.floor(nextIndex / 20));
  };

  const decideAndAdvance = (householdId: string, nextStatus: 'CONFIRMED' | 'UNMATCHED') => {
    handleDecision(householdId, nextStatus);
    moveToNextReview(householdId);
  };

  useEffect(() => {
    const handleReviewKey = (event: KeyboardEvent) => {
      const target = event.target as HTMLElement | null;
      if (target && ['INPUT', 'TEXTAREA', 'SELECT', 'BUTTON'].includes(target.tagName)) return;
      const currentId = activeReviewId ?? reviewQueueRows[0]?.householdId;
      if (!currentId) return;
      if (event.key === 'Enter') {
        event.preventDefault();
        decideAndAdvance(currentId, 'CONFIRMED');
      } else if (event.key.toLowerCase() === 'x') {
        event.preventDefault();
        decideAndAdvance(currentId, 'UNMATCHED');
      } else if (event.key === 'ArrowRight') {
        event.preventDefault();
        moveToNextReview(currentId);
      }
    };

    window.addEventListener('keydown', handleReviewKey);
    return () => window.removeEventListener('keydown', handleReviewKey);
  }, [activeReviewId, reviewQueueRows]);

  const handleSplitDecision = (candidate: SplitCandidate) => {
    const validation = validateAllocations(candidate.deposit.amount, candidate.allocations);
    if (!validation.valid) {
      setError(`배분 합계가 입금액과 일치하지 않습니다. 차액: ${validation.remaining.toLocaleString()}원`);
      return;
    }

    const householdIds = new Set(candidate.households.map((household) => household.id));
    setResults((prev) => prev.map((row) => {
      if (!householdIds.has(row.householdId)) return row;
      return {
        ...row,
        status: 'SPLIT_MATCHED',
        depositIds: [candidate.deposit.id],
        allocations: candidate.allocations,
        matchedDeposit: candidate.deposit,
        reasons: [`${candidate.formula} 분할 배분`, '사용자 확인 완료'],
      };
    }));
    const householdRows = new Set(candidate.households.map((household) => household.rowIndex));
    setValidationRows((prev) => prev.map((row) => householdRows.has(row.householdRow) ? { ...row, programStatus: 'SPLIT_MATCHED', comparison: 'EXACT_SPLIT_MATCH', differenceReason: '사용자가 분할 배분을 확정했습니다.' } : row));
    setValidationSummary((prev) => {
      if (!prev) return prev;
      const confirmedCount = candidate.households.length;
      const nextReview = Math.max(prev.programReview - confirmedCount, 0);
      const nextSplit = prev.exactSplitMatch + confirmedCount;
      return { ...prev, programReview: nextReview, exactSplitMatch: nextSplit, reproductionRate: prev.existingTotal ? ((prev.exactMatch + nextSplit) / prev.existingTotal) * 100 : 0 };
    });
    setSplitCandidates((prev) => prev.filter((item) => item.id !== candidate.id));
    setError('');
  };

  const downloadWorkbook = () => {
    if (!results.length) return;

    const reviewCount = results.filter((row) => row.status === 'REVIEW_REQUIRED').length;
    const unmatchedCount = results.filter((row) => row.status === 'UNMATCHED').length;
    if (reviewCount > 0 || unmatchedCount > 0) {
      const accept = window.confirm(`확인필요 ${reviewCount}건, 미매칭 ${unmatchedCount}건이 남아 있습니다. 그래도 결과 파일을 생성하시겠습니까?`);
      if (!accept) return;
    }

    if (!sourceWorkbook) return;
    const wb = buildResultWorkbook(sourceWorkbook, results, validationRows);
    XLSX.writeFile(wb, '[입금매칭완료]_입금확인결과.xlsx');
  };

  return (
    <main className="min-h-screen bg-slate-50 text-slate-900">
      <div className="mx-auto max-w-7xl px-5 py-8">
        <header className="mb-8 rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
          <h1 className="text-3xl font-bold text-slate-900">입금확인 자동화</h1>
          <p className="mt-2 text-sm text-slate-600">
            신청세대명단과 입금내역을 업로드하면 세대주 이름과 입금금액을 기준으로 자동 매칭합니다.
          </p>
        </header>

        {sheetAnalysis.length > 0 ? (
          <section className="mt-8 rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
            <h2 className="text-lg font-semibold text-slate-900">파일 분석 완료</h2>
            <div className="mt-4 overflow-x-auto">
              <table className="min-w-full text-left text-sm">
                <thead className="bg-slate-50 text-slate-600">
                  <tr>
                    <th className="px-3 py-2 font-medium">시트명</th>
                    <th className="px-3 py-2 font-medium">역할</th>
                    <th className="px-3 py-2 font-medium">confidence</th>
                    <th className="px-3 py-2 font-medium">감지 이유</th>
                  </tr>
                </thead>
                <tbody>
                  {sheetAnalysis.map((sheet) => (
                    <tr key={sheet.sheetName} className="border-t border-slate-200">
                      <td className="px-3 py-2 font-medium text-slate-800">{sheet.sheetName}</td>
                      <td className="px-3 py-2">{getSheetRoleLabel(sheet.role)}</td>
                      <td className="px-3 py-2">{(sheet.confidence * 100).toFixed(0)}%</td>
                      <td className="px-3 py-2 text-slate-600">{sheet.reason}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>
        ) : null}

        <section className="mt-8 grid gap-5 md:grid-cols-3">
          {[
            { kind: 'household', title: '신청세대명단', desc: '입금확인 대상 주민 목록이 들어있는 엑셀 파일을 업로드하세요.', accept: '.xlsx,.xls,.csv' },
            { kind: 'giro', title: '지로 입금내역', desc: '지로 프로그램에서 내려받은 입금내역을 업로드하세요.', accept: '.xlsx,.xls,.csv' },
            { kind: 'post', title: '우체국 입금내역', desc: '우체국 간편결제/입금조회에서 내려받은 입금내역을 업로드하세요.', accept: '.xlsx,.xls,.csv' },
          ].map((item) => (
            <div key={item.kind} className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm">
              <div className="mb-3 flex items-center gap-3">
                <div className="rounded-xl bg-sky-50 p-2 text-sky-600"><Upload size={18} /></div>
                <div>
                  <h2 className="text-lg font-semibold">{item.title}</h2>
                  <p className="text-xs text-slate-500">{item.desc}</p>
                </div>
              </div>

              <label className="flex cursor-pointer flex-col items-center justify-center rounded-2xl border border-dashed border-slate-300 bg-slate-50 px-4 py-8 text-center transition hover:border-sky-400 hover:bg-sky-50">
                <input
                  type="file"
                  accept={item.accept}
                  className="hidden"
                  onChange={(e) => handleFileUpload(item.kind as FileKind, e.target.files?.[0] ?? null)}
                />
                <FileSpreadsheet className="mb-2 text-slate-400" />
                <span className="text-sm font-medium text-slate-700">파일 선택</span>
                <span className="mt-1 text-xs text-slate-500">Drag & Drop 지원</span>
              </label>

              {files[item.kind as FileKind] ? (
                <div className="mt-4 rounded-xl border border-emerald-200 bg-emerald-50 p-3">
                  <div className="flex items-start justify-between gap-3">
                    <div>
                      <div className="flex items-center gap-2 text-sm font-semibold text-emerald-700">
                        <CheckCircle2 size={16} /> 업로드 완료
                      </div>
                      <div className="mt-2 text-xs text-slate-600">{files[item.kind as FileKind]?.name}</div>
                      <div className="mt-1 text-xs text-slate-500">{files[item.kind as FileKind]?.rows}건</div>
                    </div>
                    <button
                      type="button"
                      className="rounded-md border border-slate-200 bg-white p-1 text-slate-500"
                      onClick={() => setFiles((prev) => ({ ...prev, [item.kind as FileKind]: null }))}
                    >
                      <X size={14} />
                    </button>
                  </div>
                </div>
              ) : null}
            </div>
          ))}
        </section>

        {error ? (
          <div className="mt-6 flex items-center gap-2 rounded-xl border border-red-200 bg-red-50 p-3 text-sm text-red-700">
            <ShieldAlert size={16} /> {error}
          </div>
        ) : null}

        <div className="mt-8 flex flex-wrap items-center gap-3">
          <button
            type="button"
            onClick={startMatching}
            className="rounded-xl bg-slate-900 px-5 py-3 text-sm font-semibold text-white shadow-sm hover:bg-slate-800 disabled:cursor-not-allowed disabled:bg-slate-300"
            disabled={processing || !files.household}
          >
            {processing ? '분석 중...' : '입금 매칭 시작'}
          </button>
          <button type="button" className="rounded-xl border border-slate-300 bg-white px-4 py-3 text-sm font-medium text-slate-700">
            샘플 데이터로 체험하기
          </button>
        </div>

        {results.length > 0 ? (
          <>
            <section className="mt-8 grid gap-4 sm:grid-cols-5">
              <div className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm">
                <div className="text-sm text-slate-500">전체 신청세대</div>
                <div className="mt-2 text-3xl font-bold">{datasetStats.household || summary.total}</div>
              </div>
              <div className="rounded-2xl border border-sky-200 bg-sky-50 p-4 shadow-sm">
                <div className="text-sm text-sky-700">전체 지로</div>
                <div className="mt-2 text-3xl font-bold text-sky-700">{datasetStats.giro}</div>
              </div>
              <div className="rounded-2xl border border-violet-200 bg-violet-50 p-4 shadow-sm">
                <div className="text-sm text-violet-700">전체 우체국</div>
                <div className="mt-2 text-3xl font-bold text-violet-700">{datasetStats.post}</div>
              </div>
              <div className="rounded-2xl border border-emerald-200 bg-emerald-50 p-4 shadow-sm">
                <div className="text-sm text-emerald-700">자동확정</div>
                <div className="mt-2 text-3xl font-bold text-emerald-700">{summary.auto}</div>
              </div>
              <div className="rounded-2xl border border-amber-200 bg-amber-50 p-4 shadow-sm">
                <div className="text-sm text-amber-700">확인필요</div>
                <div className="mt-2 text-3xl font-bold text-amber-700">{summary.review}</div>
              </div>
            </section>

            <section className="mt-4 grid gap-4 sm:grid-cols-6">
              <div className="rounded-2xl border border-red-200 bg-red-50 p-4 shadow-sm">
                <div className="text-sm text-red-700">미매칭</div>
                <div className="mt-2 text-3xl font-bold text-red-700">{summary.unmatched}</div>
              </div>
              <div className="rounded-2xl border border-slate-200 bg-slate-100 p-4 shadow-sm">
                <div className="text-sm text-slate-700">전체 입금건수</div>
                <div className="mt-2 text-3xl font-bold text-slate-700">{datasetStats.totalDeposits}</div>
              </div>
              <div className="rounded-2xl border border-emerald-100 bg-emerald-50 p-4 shadow-sm">
                <div className="text-sm text-emerald-700">사용된 입금</div>
                <div className="mt-2 text-3xl font-bold text-emerald-700">{datasetStats.usedDeposits}</div>
              </div>
              <div className="rounded-2xl border border-amber-100 bg-amber-50 p-4 shadow-sm">
                <div className="text-sm text-amber-700">미사용 입금</div>
                <div className="mt-2 text-3xl font-bold text-amber-700">{datasetStats.unusedDeposits}</div>
              </div>
              <div className="rounded-2xl border border-sky-200 bg-sky-50 p-4 shadow-sm">
                <div className="text-sm text-sky-700">분할배분 확정</div>
                <div className="mt-2 text-3xl font-bold text-sky-700">{summary.split}</div>
              </div>
              <div className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm">
                <div className="text-sm text-slate-500">사용자확정</div>
                <div className="mt-2 text-3xl font-bold text-slate-700">{results.filter((row) => row.status === 'CONFIRMED').length}</div>
              </div>
            </section>

            {validationSummary ? (
              <section className="mt-8 rounded-2xl border border-sky-200 bg-sky-50 p-4 shadow-sm">
                <div className="flex flex-wrap items-start justify-between gap-3">
                  <div>
                    <h2 className="text-lg font-semibold text-slate-900">기존 W열 매칭 검증</h2>
                    <p className="mt-1 text-sm text-slate-600">기존 사람이 입력한 수식을 정답으로 확정하지 않고, 프로그램 재현 여부만 비교합니다.</p>
                  </div>
                  <div className="rounded-xl bg-white px-4 py-3 text-right shadow-sm">
                    <div className="text-xs text-slate-500">재현율</div>
                    <div className="text-2xl font-bold text-sky-700">{validationSummary.reproductionRate.toFixed(1)}%</div>
                  </div>
                </div>
                <div className="mt-4 grid gap-3 sm:grid-cols-3 lg:grid-cols-6">
                  {[
                    ['기존 매칭', validationSummary.existingTotal],
                    ['정확 재현', validationSummary.exactMatch],
                    ['분할 재현', validationSummary.exactSplitMatch],
                    ['프로그램 확인필요', validationSummary.programReview],
                    ['프로그램 놓침', validationSummary.programMissed],
                    ['다른 결과', validationSummary.programDifferent],
                  ].map(([label, value]) => (
                    <div key={String(label)} className="rounded-xl border border-sky-100 bg-white p-3">
                      <div className="text-xs text-slate-500">{label}</div>
                      <div className="mt-1 text-xl font-bold text-slate-900">{value}</div>
                    </div>
                  ))}
                </div>
                {safetyAudit ? (
                  <div className={`mt-4 rounded-xl border p-3 text-sm ${safetyAudit.totalViolations ? 'border-red-200 bg-red-50 text-red-800' : 'border-emerald-200 bg-emerald-50 text-emerald-800'}`}>
                    자동확정 안전성 위반: <strong>{safetyAudit.totalViolations}건</strong>
                    {' '} (동명이인 {safetyAudit.duplicateNameAuto}, 금액불일치 {safetyAudit.amountMismatchAuto}, 중복입금 {safetyAudit.duplicateDepositAuto}, fuzzy-only {safetyAudit.fuzzyOnlyAuto})
                  </div>
                ) : null}
                <div className="mt-4 flex flex-wrap items-center gap-2 text-sm">
                  <span className="mr-1 font-semibold text-slate-700">예외 집중분석:</span>
                  {(['ALL', 'PROGRAM_MISSED', 'PROGRAM_DIFFERENT'] as const).map((item) => {
                    const count = item === 'ALL' ? validationRows.length : validationRows.filter((row) => row.comparison === item).length;
                    return (
                      <button key={item} type="button" onClick={() => setValidationFilter(item)} className={`rounded-full px-3 py-1.5 ${validationFilter === item ? 'bg-sky-700 text-white' : 'bg-white text-sky-800'}`}>
                        {item === 'ALL' ? '전체' : item === 'PROGRAM_MISSED' ? '놓침' : '다른 결과'} {count}
                      </button>
                    );
                  })}
                </div>
                <div className="mt-4 max-h-96 overflow-auto rounded-xl border border-sky-100 bg-white">
                  <table className="min-w-full text-left text-xs">
                    <thead className="sticky top-0 bg-slate-100 text-slate-600">
                      <tr>
                        <th className="px-3 py-2">행</th>
                        <th className="px-3 py-2">세대주</th>
                        <th className="px-3 py-2">기존 W 수식</th>
                        <th className="px-3 py-2">프로그램 상태</th>
                        <th className="px-3 py-2">비교결과</th>
                        <th className="px-3 py-2">차이사유</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredValidationRows.map((row) => (
                        <Fragment key={`${row.householdRow}-${row.formula}`}>
                          <tr key={`${row.householdRow}-${row.formula}`} className="border-t border-slate-200">
                            <td className="px-3 py-2">{row.householdRow}</td>
                            <td className="px-3 py-2">{row.householdName}</td>
                            <td className="px-3 py-2 font-mono">{row.formula} ({row.existingSource} {row.existingDepositRow})</td>
                            <td className="px-3 py-2">{row.programStatus}</td>
                            <td className="px-3 py-2 font-semibold">{row.comparison}</td>
                            <td className="px-3 py-2 text-slate-600">{row.differenceReason || '-'}</td>
                          </tr>
                          {validationFilter !== 'ALL' ? (
                            <tr key={`${row.householdRow}-${row.formula}-detail`} className="bg-sky-50">
                              <td colSpan={6} className="px-3 py-3">
                                <div className="grid gap-2 text-xs text-slate-700 md:grid-cols-4">
                                  <div><strong>세대주</strong><br />{row.householdName}</div>
                                  <div><strong>기준금액</strong><br />{row.householdAmount.toLocaleString()}원</div>
                                  <div><strong>주소</strong><br />{row.householdAddress || '-'}</div>
                                  <div><strong>기존 W</strong><br />{row.formula} · 배분 {row.existingAllocatedAmount}</div>
                                </div>
                                <div className="mt-3 overflow-x-auto rounded-lg border border-sky-100 bg-white">
                                  <table className="min-w-full text-left text-xs">
                                    <thead className="bg-slate-100 text-slate-600">
                                      <tr><th className="px-2 py-1.5">출처/행</th><th className="px-2 py-1.5">입금자명</th><th className="px-2 py-1.5">입금금액</th><th className="px-2 py-1.5">주소</th><th className="px-2 py-1.5">거래일</th><th className="px-2 py-1.5">점수</th></tr>
                                    </thead>
                                    <tbody>
                                      {row.programCandidates.length ? row.programCandidates.map((candidate) => (
                                        <tr key={`${row.householdRow}-${candidate.source}-${candidate.row}`} className="border-t border-slate-200">
                                          <td className="px-2 py-1.5">{candidate.source} / {candidate.row}</td>
                                          <td className="px-2 py-1.5">{candidate.depositorName || '-'}</td>
                                          <td className="px-2 py-1.5">{candidate.amount.toLocaleString()}원</td>
                                          <td className="px-2 py-1.5">{candidate.address || '-'}</td>
                                          <td className="px-2 py-1.5 whitespace-nowrap">{candidate.date}</td>
                                          <td className="px-2 py-1.5 font-semibold text-sky-700">{candidate.score}점</td>
                                        </tr>
                                      )) : (
                                        <tr><td colSpan={6} className="px-2 py-2 text-slate-500">프로그램 후보 없음</td></tr>
                                      )}
                                    </tbody>
                                  </table>
                                </div>
                              </td>
                            </tr>
                          ) : null}
                        </Fragment>
                      ))}
                    </tbody>
                  </table>
                </div>
              </section>
            ) : null}

            <section className="mt-8 rounded-2xl border border-slate-200 bg-white p-4 shadow-sm">
              <div className="mb-4 flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
                <div className="flex flex-wrap gap-2">
                  {(['ALL', 'AUTO_MATCHED', 'REVIEW_REQUIRED', 'SPLIT_MATCHED', 'UNMATCHED', 'CONFIRMED'] as const).map((item) => (
                    <button
                      key={item}
                      type="button"
                      onClick={() => setFilter(item)}
                      className={`rounded-full px-3 py-1.5 text-sm ${filter === item ? 'bg-slate-900 text-white' : 'bg-slate-100 text-slate-700'}`}
                    >
                      {item === 'ALL' ? '전체' : item === 'AUTO_MATCHED' ? '자동확정' : item === 'REVIEW_REQUIRED' ? '확인필요' : item === 'SPLIT_MATCHED' ? '분할배분 확정' : item === 'UNMATCHED' ? '미매칭' : '사용자확정'}
                    </button>
                  ))}
                </div>
                <div className="flex w-full max-w-md items-center gap-2">
                  <select
                    value={sortKey}
                    onChange={(e) => setSortKey(e.target.value as 'householdName' | 'amount' | 'status' | 'score')}
                    className="rounded-xl border border-slate-200 bg-slate-50 px-3 py-2 text-sm outline-none"
                  >
                    <option value="householdName">세대주</option>
                    <option value="amount">입금액</option>
                    <option value="status">상태</option>
                    <option value="score">score</option>
                  </select>
                  <div className="relative w-full">
                  <Search className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
                    <input
                      value={search}
                      onChange={(e) => setSearch(e.target.value)}
                      className="w-full rounded-xl border border-slate-200 bg-slate-50 py-2.5 pl-9 pr-3 text-sm outline-none ring-0"
                      placeholder="세대주 / 입금자 / 주소"
                    />
                  </div>
                </div>
              </div>

              <div className="overflow-x-auto">
                <table className="min-w-full text-left text-sm">
                  <thead className="bg-slate-50 text-slate-600">
                    <tr>
                      <th className="px-3 py-3 font-medium">상태</th>
                      <th className="px-3 py-3 font-medium">세대주</th>
                      <th className="px-3 py-3 font-medium">주소</th>
                      <th className="px-3 py-3 font-medium">자부담</th>
                      <th className="px-3 py-3 font-medium">입금자명</th>
                      <th className="px-3 py-3 font-medium">입금금액</th>
                      <th className="px-3 py-3 font-medium">입금처</th>
                      <th className="px-3 py-3 font-medium">신뢰도</th>
                      <th className="px-3 py-3 font-medium">매칭사유</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredResults.map((row) => (
                      <tr key={row.householdId} className="border-t border-slate-200">
                        <td className="px-3 py-3">
                          <span className={`inline-flex rounded-full px-2.5 py-1 text-xs font-semibold ${row.status === 'AUTO_MATCHED' ? 'bg-emerald-100 text-emerald-700' : row.status === 'REVIEW_REQUIRED' ? 'bg-amber-100 text-amber-700' : row.status === 'SPLIT_MATCHED' ? 'bg-sky-100 text-sky-700' : row.status === 'UNMATCHED' ? 'bg-red-100 text-red-700' : 'bg-slate-100 text-slate-700'}`}>
                            {row.status === 'AUTO_MATCHED' ? '자동확정' : row.status === 'REVIEW_REQUIRED' ? '확인필요' : row.status === 'SPLIT_MATCHED' ? '분할배분 확정' : row.status === 'UNMATCHED' ? '미매칭' : row.status === 'CONFIRMED' ? '사용자확정' : '제외'}
                          </span>
                        </td>
                        <td className="px-3 py-3">{row.householdName}</td>
                        <td className="px-3 py-3">{row.matchedDeposit?.address ?? '-'}</td>
                        <td className="px-3 py-3">{row.matchedDeposit ? `${row.matchedDeposit.amount.toLocaleString()}원` : '-'}</td>
                        <td className="px-3 py-3">{row.matchedDeposit?.depositorName ?? '-'}</td>
                        <td className="px-3 py-3">{row.matchedDeposit?.amount ? `${row.matchedDeposit.amount.toLocaleString()}원` : '-'}</td>
                        <td className="px-3 py-3">{row.matchedDeposit?.source ?? '-'}</td>
                        <td className="px-3 py-3">{row.score}점</td>
                        <td className="px-3 py-3 text-xs text-slate-600">{row.reasons.join(' / ')}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </section>

            {splitCandidates.length > 0 ? (
              <section className="mt-8 rounded-2xl border border-sky-200 bg-sky-50 p-4 shadow-sm">
                <div className="mb-4 flex items-center justify-between gap-3">
                  <div>
                    <h3 className="text-lg font-semibold text-sky-900">분할 배분 확인</h3>
                    <p className="mt-1 text-sm text-sky-800">기존 W열의 분할 근거를 참고 후보로 표시합니다. 자동확정하지 않으며, 배분 합계가 정확히 맞을 때만 승인됩니다.</p>
                  </div>
                  <span className="rounded-full bg-white px-3 py-1 text-sm font-semibold text-sky-700">{splitCandidates.length}건</span>
                </div>
                <div className="space-y-4">
                  {splitCandidates.map((candidate) => (
                    <div key={candidate.id} className="rounded-2xl border border-sky-200 bg-white p-4">
                      <div className="flex flex-wrap items-start justify-between gap-3">
                        <div>
                          <div className="text-sm font-semibold text-slate-900">{candidate.formula} · {candidate.deposit.source}</div>
                          <div className="mt-1 text-sm text-slate-600">{candidate.deposit.depositorName || '-'} / 원입금 {candidate.deposit.amount.toLocaleString()}원 / {candidate.households.length}세대 후보</div>
                        </div>
                        <button type="button" onClick={() => handleSplitDecision(candidate)} className="rounded-lg bg-sky-600 px-3 py-2 text-sm font-semibold text-white hover:bg-sky-500">배분 확정</button>
                      </div>
                      <div className="mt-3 grid gap-2 md:grid-cols-2">
                        {candidate.households.map((household) => (
                          <div key={household.id} className="rounded-xl border border-slate-200 bg-slate-50 p-3 text-sm">
                            <div className="font-semibold text-slate-900">{household.name} · {household.allocatedAmount.toLocaleString()}원</div>
                            <div className="mt-1 text-slate-600">{household.address || '-'}</div>
                          </div>
                        ))}
                      </div>
                      <div className="mt-3 text-right text-xs text-slate-500">배분 합계 {candidate.allocations.reduce((sum, allocation) => sum + allocation.allocatedAmount, 0).toLocaleString()}원</div>
                    </div>
                  ))}
                </div>
              </section>
            ) : null}

            {reviewTotal > 0 ? (
              <section className="mt-8 rounded-2xl border border-amber-200 bg-amber-50 p-4 shadow-sm">
                <div className="mb-4 flex flex-wrap items-center justify-between gap-3 text-amber-800">
                  <div className="flex items-center gap-2">
                    <AlertTriangle size={18} />
                    <h3 className="text-lg font-semibold">확인필요 항목</h3>
                    <span className="rounded-full bg-white px-2 py-0.5 text-xs font-semibold">{reviewTotal}건</span>
                  </div>
                  <div className="flex items-center gap-2 text-sm">
                    <button type="button" onClick={() => { setReviewQueueFilter('ALL'); setReviewPage(0); }} className={`rounded-full px-3 py-1.5 ${reviewQueueFilter === 'ALL' ? 'bg-amber-700 text-white' : 'bg-white text-amber-800'}`}>전체 {reviewTotal}</button>
                    <button type="button" onClick={() => { setReviewQueueFilter('DUPLICATE_NAME'); setReviewPage(0); }} className={`rounded-full px-3 py-1.5 ${reviewQueueFilter === 'DUPLICATE_NAME' ? 'bg-amber-700 text-white' : 'bg-white text-amber-800'}`}>동명이인 {results.filter((row) => row.status === 'REVIEW_REQUIRED' && getReviewCategory(row) === 'DUPLICATE_NAME').length}</button>
                    <button type="button" onClick={() => { setReviewQueueFilter('AMOUNT'); setReviewPage(0); }} className={`rounded-full px-3 py-1.5 ${reviewQueueFilter === 'AMOUNT' ? 'bg-amber-700 text-white' : 'bg-white text-amber-800'}`}>금액 {results.filter((row) => row.status === 'REVIEW_REQUIRED' && getReviewCategory(row) === 'AMOUNT').length}</button>
                    <button type="button" onClick={() => { setReviewQueueFilter('FUZZY'); setReviewPage(0); }} className={`rounded-full px-3 py-1.5 ${reviewQueueFilter === 'FUZZY' ? 'bg-amber-700 text-white' : 'bg-white text-amber-800'}`}>유사명 {results.filter((row) => row.status === 'REVIEW_REQUIRED' && getReviewCategory(row) === 'FUZZY').length}</button>
                  </div>
                </div>
                <div className="space-y-4">
                  {reviewItems.map((item) => (
                    <div key={item.householdId} onClick={() => setActiveReviewId(item.householdId)} className={`rounded-2xl border bg-white p-4 ${activeReviewId === item.householdId ? 'border-sky-400 ring-2 ring-sky-100' : 'border-amber-200'}`}>
                      <div className="mb-3 overflow-x-auto rounded-xl border border-slate-200">
                        <table className="min-w-full text-left text-xs">
                          <thead className="bg-slate-50 text-slate-600">
                            <tr>
                              <th className="px-2 py-2">세대주</th>
                              <th className="px-2 py-2">입금자</th>
                              <th className="px-2 py-2">입금금액</th>
                              <th className="px-2 py-2">기준금액</th>
                              <th className="px-2 py-2">거래일</th>
                              <th className="px-2 py-2">주소</th>
                              <th className="px-2 py-2">점수</th>
                            </tr>
                          </thead>
                          <tbody>
                            {item.candidates.map((candidate: DepositRow) => (
                              <tr key={`${item.householdId}-${candidate.id}`} className="border-t border-slate-200">
                                <td className="px-2 py-2 font-semibold">{item.householdName}</td>
                                <td className="px-2 py-2">{candidate.depositorName || '-'}</td>
                                <td className="px-2 py-2">{candidate.amount.toLocaleString()}원</td>
                                <td className="px-2 py-2">{(item.householdAmount ?? 0).toLocaleString()}원</td>
                                <td className="px-2 py-2 whitespace-nowrap">{getCandidateDate(candidate)}</td>
                                <td className="max-w-56 px-2 py-2">{candidate.address || '-'}</td>
                                <td className="px-2 py-2 font-semibold text-sky-700">{getCandidateScore({ name: item.householdName, address: item.householdAddress, amount: item.householdAmount ?? 0 }, candidate)}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                      <div className="grid gap-4 md:grid-cols-2">
                        <div>
                          <div className="text-xs uppercase tracking-wide text-slate-500">신청세대</div>
                          <div className="mt-2 text-xl font-bold">{item.householdName}</div>
                          <div className="mt-2 text-sm text-slate-600">자부담: {item.householdAmount ? `${item.householdAmount.toLocaleString()}원` : '미기재'}</div>
                          <div className="mt-1 text-sm text-slate-600">주소: {item.householdAddress || '-'}</div>
                        </div>
                        <div>
                          <div className="text-xs uppercase tracking-wide text-slate-500">입금내역 후보</div>
                          <div className="mt-2 space-y-2">
                            {item.candidates.length ? item.candidates.map((candidate: any) => (
                              <div key={candidate.id} className="rounded-xl border border-slate-200 bg-slate-50 p-2 text-sm text-slate-700">
                                <div className="flex items-center justify-between gap-3">
                                  <span className="font-semibold">{candidate.depositorName}</span>
                                  <span className="rounded-full bg-slate-200 px-2 py-0.5 text-[10px] font-medium text-slate-700">{candidate.source}</span>
                                </div>
                                <div className="mt-1">금액: {candidate.amount.toLocaleString()}원</div>
                                <div className="mt-0.5">주소: {candidate.address || '-'}</div>
                                <div className="mt-0.5">거래일: {getCandidateDate(candidate)}</div>
                                <div className="mt-0.5 font-semibold text-sky-700">후보 점수: {getCandidateScore({ name: item.householdName, address: item.householdAddress, amount: item.householdAmount ?? 0 }, candidate)}점</div>
                              </div>
                            )) : (
                              <div className="rounded-xl border border-slate-200 bg-slate-50 p-2 text-sm text-slate-600">후보 없음</div>
                            )}
                          </div>
                        </div>
                      </div>
                      <div className="mt-4 rounded-xl border border-amber-100 bg-amber-50 p-3 text-sm text-amber-800">
                        {item.reasons.join(' / ')}
                      </div>
                      <div className="mt-3 flex gap-2">
                        <button type="button" onClick={() => decideAndAdvance(item.householdId, 'CONFIRMED')} className="rounded-lg bg-slate-900 px-3 py-2 text-sm font-medium text-white">입금확정</button>
                        <button type="button" onClick={() => decideAndAdvance(item.householdId, 'UNMATCHED')} className="rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm font-medium text-slate-700">매칭 안 함</button>
                        <button type="button" onClick={() => moveToNextReview(item.householdId)} className="rounded-lg border border-sky-300 bg-sky-50 px-3 py-2 text-sm font-medium text-sky-700">다음 검토</button>
                      </div>
                    </div>
                  ))}
                </div>
                <div className="mt-4 flex items-center justify-between text-sm text-amber-800">
                  <span>{reviewFilteredTotal}건 중 {reviewPage * 20 + 1}-{Math.min((reviewPage + 1) * 20, reviewFilteredTotal)}건 표시</span>
                  <div className="flex gap-2">
                    <button type="button" disabled={reviewPage === 0} onClick={() => setReviewPage((page) => Math.max(page - 1, 0))} className="rounded-lg bg-white px-3 py-1.5 disabled:opacity-40">이전</button>
                    <span className="px-2 py-1.5">{reviewPage + 1} / {reviewPageCount}</span>
                    <button type="button" disabled={reviewPage + 1 >= reviewPageCount} onClick={() => setReviewPage((page) => Math.min(page + 1, reviewPageCount - 1))} className="rounded-lg bg-white px-3 py-1.5 disabled:opacity-40">다음</button>
                  </div>
                </div>
              </section>
            ) : null}

            <div className="mt-8 flex justify-end">
              <button type="button" onClick={downloadWorkbook} className="inline-flex items-center gap-2 rounded-xl bg-emerald-600 px-5 py-3 text-sm font-semibold text-white shadow-sm hover:bg-emerald-500">
                <Download size={16} /> 결과 엑셀 다운로드
              </button>
            </div>
          </>
        ) : null}
      </div>
    </main>
  );
}
