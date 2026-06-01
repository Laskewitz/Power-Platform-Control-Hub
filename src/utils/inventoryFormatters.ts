import type { InventorySharingSummary } from '../types/inventory.ts';

/** Returns true if value is a non-empty string. */
export function hasText(value: unknown): value is string {
  return typeof value === 'string' && value.trim().length > 0;
}

/** Formats an InventorySharingSummary into a human-readable string. */
export function formatSharedSummary(summary?: InventorySharingSummary): string | null {
  if (!summary) return null;
  if (summary.entireTenant) return 'Entire tenant';

  const parts: string[] = [];
  if (typeof summary.userCount === 'number') {
    parts.push(`${summary.userCount} user${summary.userCount === 1 ? '' : 's'}`);
  }
  if (typeof summary.groupCount === 'number') {
    parts.push(`${summary.groupCount} group${summary.groupCount === 1 ? '' : 's'}`);
  }

  return parts.length > 0 ? parts.join(', ') : null;
}

/** Safely casts an unknown value to a string array, filtering blanks. */
export function getStringArray(value: unknown): string[] {
  if (!Array.isArray(value)) return [];
  return value.filter((item): item is string => typeof item === 'string' && item.trim().length > 0);
}

/** Extracts [name, count] pairs from a capabilitiesCounts-style object. */
export function getCapabilityEntries(value: unknown): Array<[string, number]> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return [];
  return Object.entries(value).filter((entry): entry is [string, number] => typeof entry[1] === 'number');
}
