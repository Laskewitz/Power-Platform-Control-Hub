/**
 * Tombstone service — persists deleted resource IDs so they stay hidden
 * across all users and sessions until the Inventory API refreshes.
 *
 * PRIMARY: Dataverse table `ppa_resourcetombstone` (shared across all users).
 * FALLBACK: localStorage (per user/browser) when Dataverse is unavailable
 *           (e.g. running locally without a live connection).
 *
 * To upgrade another tenant: import the solution in `solution/`, then run:
 *   npx power-apps add-data-source --api-id dataverse \
 *     --resource-name ppa_resourcetombstone --org-url <your-org-url>
 */

import { Ppa_resourcetombstonesService } from '../generated/services/Ppa_resourcetombstonesService.ts';

const LS_KEY = 'ppa_resourceTombstones';
const MAX_AGE_DAYS = 30;

export interface TombstoneEntry {
  resourceId: string;
  resourceType: string;
  environmentId: string;
  displayName: string;
  deletedBy?: string;
}

// Local cache: resourceId → Dataverse record GUID (for fast deletes)
const dvIdCache = new Map<string, string>();

// ── localStorage fallback ─────────────────────────────────────────────────

interface LocalEntry extends TombstoneEntry { deletedOn: string }

function lsLoad(): LocalEntry[] {
  try {
    const raw = localStorage.getItem(LS_KEY);
    if (!raw) return [];
    const entries = JSON.parse(raw) as LocalEntry[];
    const cutoff = Date.now() - MAX_AGE_DAYS * 86_400_000;
    return entries.filter((e) => new Date(e.deletedOn).getTime() > cutoff);
  } catch { return []; }
}

function lsSave(entries: LocalEntry[]): void {
  try { localStorage.setItem(LS_KEY, JSON.stringify(entries)); } catch { /* quota */ }
}

function lsGetIds(): Set<string> {
  return new Set(lsLoad().map((e) => e.resourceId));
}

function lsAdd(entry: TombstoneEntry): void {
  const entries = lsLoad();
  if (entries.some((e) => e.resourceId === entry.resourceId)) return;
  entries.push({ ...entry, deletedOn: new Date().toISOString() });
  lsSave(entries);
}

function lsRemove(resourceId: string): void {
  lsSave(lsLoad().filter((e) => e.resourceId !== resourceId));
}

// ── Dataverse operations ──────────────────────────────────────────────────

async function dvGetIds(): Promise<Set<string>> {
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - MAX_AGE_DAYS);

  const result = await Ppa_resourcetombstonesService.getAll({
    select: ['ppa_resourceid', 'ppa_resourcetombstoneid', 'ppa_deletedon'],
    filter: `ppa_deletedon ge ${cutoff.toISOString()}`,
  });

  if (!result.success || !result.data) return new Set();

  const ids = new Set<string>();
  for (const record of result.data) {
    if (!record.ppa_resourceid) continue;
    ids.add(record.ppa_resourceid);
    if (record.ppa_resourcetombstoneid) {
      dvIdCache.set(record.ppa_resourceid, record.ppa_resourcetombstoneid);
    }
  }
  return ids;
}

async function dvAdd(entry: TombstoneEntry): Promise<void> {
  // ownerid and statecode are intentionally omitted — Dataverse assigns the
  // current user as owner and defaults statecode to 0 (Active) automatically.
  const payload = {
    ppa_resourceid: entry.resourceId,
    ppa_resourcetype: entry.resourceType,
    ppa_environmentid: entry.environmentId,
    ppa_displayname: entry.displayName,
    ppa_deletedby: entry.deletedBy ?? '',
    ppa_deletedon: new Date().toISOString(),
  } as Parameters<typeof Ppa_resourcetombstonesService.create>[0];

  const result = await Ppa_resourcetombstonesService.create(payload);
  if (result.success && result.data?.ppa_resourcetombstoneid) {
    dvIdCache.set(entry.resourceId, result.data.ppa_resourcetombstoneid);
  } else if (!result.success) {
    console.error('[Tombstone] Dataverse create failed:', result.error);
    throw new Error(result.error?.message ?? 'Dataverse create failed');
  }
}

async function dvRemove(resourceId: string): Promise<void> {
  let id = dvIdCache.get(resourceId);
  if (!id) {
    const lookup = await Ppa_resourcetombstonesService.getAll({
      select: ['ppa_resourcetombstoneid'],
      filter: `ppa_resourceid eq '${resourceId.replace(/'/g, "''")}'`,
    });
    id = lookup.data?.[0]?.ppa_resourcetombstoneid;
  }
  if (id) {
    await Ppa_resourcetombstonesService.delete(id);
    dvIdCache.delete(resourceId);
  }
}

// ── Public API (async) ────────────────────────────────────────────────────

/**
 * Fetches all tombstoned resource IDs.
 * Tries Dataverse first; falls back to localStorage on error.
 */
export async function fetchTombstonedIds(): Promise<Set<string>> {
  try {
    const ids = await dvGetIds();
    // Merge with any local entries written while offline
    for (const id of lsGetIds()) ids.add(id);
    return ids;
  } catch {
    return lsGetIds();
  }
}

/**
 * Records a resource as deleted. Writes to both Dataverse and localStorage
 * so the entry is available even before the next fetch completes.
 */
export async function addTombstone(entry: TombstoneEntry): Promise<void> {
  lsAdd(entry); // instant local write first
  try {
    await dvAdd(entry);
  } catch {
    // Dataverse unavailable — localStorage write above is the fallback
  }
}

/**
 * Removes a tombstone (rollback on failed delete).
 */
export async function removeTombstone(resourceId: string): Promise<void> {
  lsRemove(resourceId);
  try {
    await dvRemove(resourceId);
  } catch {
    // Best effort
  }
}
