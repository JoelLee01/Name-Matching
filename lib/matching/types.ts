export type MatchStatus =
  | 'UNMATCHED'
  | 'AUTO_MATCHED'
  | 'REVIEW_REQUIRED'
  | 'SPLIT_MATCHED'
  | 'CONFIRMED'
  | 'IGNORED';

export type DepositSource = 'GIRO' | 'POST';

export interface HouseholdRow {
  id: string;
  rowIndex: number;
  name: string;
  requiredAmount: number;
  address?: string;
  normalizedName: string;
  normalizedAddress: string;
  originalData: Record<string, unknown>;
}

export interface DepositRow {
  id: string;
  rowIndex: number;
  source: DepositSource;
  depositorName: string;
  amount: number;
  address?: string;
  normalizedDepositorName: string;
  normalizedAddress: string;
  originalData: Record<string, unknown>;
}

export interface MatchResult {
  householdId: string;
  householdName: string;
  depositIds: string[];
  status: MatchStatus;
  score: number;
  reasons: string[];
  matchedDeposit: DepositRow | null;
  candidates: DepositRow[];
  allocations?: Array<{
    depositId: string;
    householdId: string;
    allocatedAmount: number;
  }>;
  householdAddress?: string;
  householdAmount?: number;
}
