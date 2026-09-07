import type { DepositRow, HouseholdRow } from './types';
import type { ExistingFormulaReference } from './validation';

export interface DepositAllocation {
  depositId: string;
  householdId: string;
  allocatedAmount: number;
}

export interface SplitCandidate {
  id: string;
  formula: string;
  deposit: DepositRow;
  households: Array<HouseholdRow & { allocatedAmount: number }>;
  allocations: DepositAllocation[];
}

export function validateAllocations(depositAmount: number, allocations: DepositAllocation[]) {
  const total = allocations.reduce((sum, item) => sum + item.allocatedAmount, 0);
  return {
    valid: total === depositAmount,
    total,
    remaining: depositAmount - total,
  };
}

export function buildSplitCandidates(
  references: ExistingFormulaReference[],
  households: HouseholdRow[],
  deposits: DepositRow[],
): SplitCandidate[] {
  const householdByRow = new Map(households.map((household) => [household.rowIndex, household]));
  const depositByKey = new Map(deposits.map((deposit) => [`${deposit.source}:${deposit.rowIndex}`, deposit]));
  const grouped = new Map<string, SplitCandidate>();

  references
    .filter((reference) => reference.divisor > 1)
    .forEach((reference) => {
      const deposit = depositByKey.get(`${reference.source}:${reference.depositRow}`);
      const household = householdByRow.get(reference.householdRow);
      if (!deposit || !household) return;

      const id = `${deposit.id}:${reference.depositRows.join('+')}`;
      const existing = grouped.get(id);
      const allocatedAmount = Math.round(deposit.amount / reference.divisor);
      const allocation = { depositId: deposit.id, householdId: household.id, allocatedAmount };
      if (existing) {
        existing.households.push({ ...household, allocatedAmount });
        existing.allocations.push(allocation);
        return;
      }

      grouped.set(id, {
        id,
        formula: reference.formula,
        deposit,
        households: [{ ...household, allocatedAmount }],
        allocations: [allocation],
      });
    });

  return [...grouped.values()];
}
