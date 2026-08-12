import { useEffect, useReducer } from 'react';
import { AadusersService } from '../generated/services/AadusersService.ts';
import { getEnvironmentDataverseInfo } from './copilotStudioService.ts';
import {
  listAllDataverseSystemUsers,
  listDataverseSystemUsers,
} from './dataverseConnectorService.ts';

// GUIDs whose first three groups are all zeros are system/application accounts
const SYSTEM_PATTERN = /^0{8}-0{4}-0{4}/i;
const GUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

const cache = new Map<string, string>();

// Pre-seed known system accounts
cache.set('00000000-0000-0000-0000-000000000000', 'SYSTEM');
cache.set('system', 'SYSTEM');

function isSystemGuid(guid: string): boolean {
  return SYSTEM_PATTERN.test(guid);
}

// Listeners so hooks can react when a new GUID resolves
const listeners = new Set<() => void>();
let directoryLoad: Promise<void> | null = null;
function notify() {
  for (const fn of listeners) fn();
}

function cacheDirectoryUser(user: {
  aaduserid: string;
  id: string;
  displayname?: string;
  mail?: string;
  userprincipalname?: string;
}): void {
  const name = user.displayname ?? user.mail ?? user.userprincipalname;
  if (!name) return;
  cache.set(user.id.toLowerCase(), name);
  cache.set(user.aaduserid.toLowerCase(), name);
}

function loadDirectory(): Promise<void> {
  directoryLoad ??= AadusersService.getAll({
    select: ['aaduserid', 'id', 'displayname', 'mail', 'userprincipalname'],
    top: 5000,
  }).then((result) => {
    if (result.success) result.data.forEach(cacheDirectoryUser);
  }).catch(() => {
    directoryLoad = null;
  });
  return directoryLoad;
}

async function resolveOwners(
  guids: (string | undefined | null)[],
  environmentId?: string,
): Promise<void> {
  const normalized = guids
    .filter((guid): guid is string => Boolean(guid))
    .map((guid) => guid.trim().replace(/^\{|\}$/g, '').toLowerCase());
  let cacheChanged = false;
  for (const guid of normalized) {
    if (isSystemGuid(guid) && !cache.has(guid)) {
      cache.set(guid, 'SYSTEM');
      cacheChanged = true;
    }
  }

  const unresolved = [...new Set(
    normalized
      .filter((guid) => GUID_PATTERN.test(guid))
      .filter((guid) => !cache.has(guid)),
  )];
  if (unresolved.length === 0) {
    if (cacheChanged) notify();
    return;
  }

  const GLOBAL_BATCH_SIZE = 40;
  for (let index = 0; index < unresolved.length; index += GLOBAL_BATCH_SIZE) {
    const batch = unresolved.slice(index, index + GLOBAL_BATCH_SIZE);
    try {
      const byDirectoryId = await AadusersService.getAll({
        filter: batch.map((guid) => `id eq '${guid}'`).join(' or '),
        select: ['aaduserid', 'id', 'displayname', 'mail', 'userprincipalname'],
        top: batch.length,
      });
      if (byDirectoryId.success) {
        byDirectoryId.data.forEach(cacheDirectoryUser);
      }

      const remaining = batch.filter((guid) => !cache.has(guid));
      if (remaining.length > 0) {
        const byAadUserId = await AadusersService.getAll({
          filter: remaining.map((guid) => `aaduserid eq ${guid}`).join(' or '),
          select: ['aaduserid', 'id', 'displayname', 'mail', 'userprincipalname'],
          top: remaining.length,
        });
        if (byAadUserId.success) {
          byAadUserId.data.forEach(cacheDirectoryUser);
        }
      }
    } catch {
      // Environment-specific resolution below remains available.
    }
  }

  let dataverseCandidates = unresolved.filter((guid) => !cache.has(guid));
  if (dataverseCandidates.length > 0) {
    await loadDirectory();
    dataverseCandidates = dataverseCandidates.filter((guid) => !cache.has(guid));
  }

  if (environmentId && dataverseCandidates.length > 0) {
    const environment = await getEnvironmentDataverseInfo(environmentId);
    if (environment.instanceUrl) {
      const BATCH_SIZE = 20;
      for (let index = 0; index < dataverseCandidates.length; index += BATCH_SIZE) {
        const batch = dataverseCandidates.slice(index, index + BATCH_SIZE);
        try {
          const result = await listDataverseSystemUsers(environment.instanceUrl, batch);
          if (!result.success) continue;
          const cacheSystemUsers = (users: typeof result.data.value): void => {
            for (const user of users ?? []) {
              const name = user.fullname ?? user.internalemailaddress ?? user.domainname;
              if (!name) continue;
              if (user.systemuserid) cache.set(user.systemuserid.toLowerCase(), name);
              if (user.azureactivedirectoryobjectid) {
                cache.set(user.azureactivedirectoryobjectid.toLowerCase(), name);
              }
            }
          };
          cacheSystemUsers(result.data.value);

          if (batch.some((guid) => !cache.has(guid))) {
            const directory = await listAllDataverseSystemUsers(environment.instanceUrl);
            if (directory.success) cacheSystemUsers(directory.data.value);
          }
        } catch {
          // Keep the raw ID visible; reopening the detail view can retry.
        }
      }
    }
  }

  notify();
}

export async function resolveOwner(
  guid: string | undefined | null,
  environmentId?: string,
): Promise<string> {
  if (!guid) return '—';
  const key = guid.toLowerCase();

  if (cache.has(key)) return cache.get(key)!;

  if (isSystemGuid(key)) {
    cache.set(key, 'SYSTEM');
    notify();
    return 'SYSTEM';
  }

  await resolveOwners([guid], environmentId);
  return cache.get(key) ?? guid;
}

/** Resolve multiple GUIDs in parallel, no-op for already-cached ones. */
export function prefetchOwners(
  guids: (string | undefined | null)[],
  environmentId?: string,
): void {
  void resolveOwners(guids, environmentId);
}

/**
 * Hook: resolves all provided GUIDs and returns a Map<guid, displayName>.
 * Re-renders when any GUID in the list resolves.
 */
export function useOwners(
  guids: (string | undefined | null)[],
  environmentId?: string,
): Map<string, string> {
  // forceUpdate trick — increment to trigger re-render when cache changes
  const [, forceUpdate] = useReducer((x: number) => x + 1, 0);

  useEffect(() => {
    listeners.add(forceUpdate);
    return () => { listeners.delete(forceUpdate); };
  }, []);

  useEffect(() => {
    prefetchOwners(guids, environmentId);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [guids.join(','), environmentId]);

  const result = new Map<string, string>();
  for (const guid of guids) {
    if (!guid) continue;
    const key = guid.toLowerCase();
    result.set(key, cache.get(key) ?? guid);
  }
  return result;
}

export interface EnvironmentOwnerRequest {
  guid: string;
  environmentId?: string;
}

/** Resolve a tenant resource list while preserving each owner's environment context. */
export function useEnvironmentOwners(requests: EnvironmentOwnerRequest[]): Map<string, string> {
  const [, forceUpdate] = useReducer((x: number) => x + 1, 0);
  const signature = requests
    .map(({ guid, environmentId }) => `${guid.toLowerCase()}@${environmentId ?? ''}`)
    .sort()
    .join(',');

  useEffect(() => {
    listeners.add(forceUpdate);
    return () => { listeners.delete(forceUpdate); };
  }, []);

  useEffect(() => {
    const unique = new Map<string, EnvironmentOwnerRequest>();
    for (const request of requests) {
      const key = request.guid.toLowerCase();
      if (!unique.has(key)) unique.set(key, request);
    }

    const byEnvironment = new Map<string | undefined, string[]>();
    for (const request of unique.values()) {
      const guids = byEnvironment.get(request.environmentId) ?? [];
      guids.push(request.guid);
      byEnvironment.set(request.environmentId, guids);
    }
    for (const [environmentId, guids] of byEnvironment) {
      prefetchOwners(guids, environmentId);
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [signature]);

  const result = new Map<string, string>();
  for (const { guid } of requests) {
    const key = guid.toLowerCase();
    result.set(key, cache.get(key) ?? guid);
  }
  return result;
}

/** Convenience hook for a single GUID. */
export function useOwner(guid: string | undefined | null, environmentId?: string): string {
  const map = useOwners(guid ? [guid] : [], environmentId);
  if (!guid) return '—';
  return map.get(guid.toLowerCase()) ?? guid;
}
