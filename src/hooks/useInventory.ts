import { useCallback, useEffect, useState } from 'react';
import {
  computeResourceCounts,
  fetchEnvironments,
  fetchResources,
} from '../services/inventoryApi.ts';
import type { Resource, ResourceCounts } from '../types/inventory.ts';
import { extractMessage } from '../utils/errorUtils.ts';
import { isGuid, resolveUserIds } from '../services/userService.ts';

export interface UseInventoryResult {
  resources: Resource[];
  environments: Resource[];
  counts: ResourceCounts | null;
  isLoading: boolean;
  loadingLabel: string | null;
  error: string | null;
  refresh: () => Promise<void>;
}

/** Extract the owner GUID from a resource, if it is a bare GUID rather than a display name. */
function extractOwnerGuid(r: Resource): string | null {
  const p = r.properties as Record<string, unknown>;
  if (p.owner && typeof p.owner === 'object') {
    const o = p.owner as { displayName?: string; id?: string };
    if (!o.displayName && o.id && isGuid(o.id)) return o.id;
  }
  if (typeof p.ownerId === 'string' && isGuid(p.ownerId)) return p.ownerId;
  if (typeof p.createdBy === 'string' && isGuid(p.createdBy)) return p.createdBy;
  return null;
}

export function useInventory(): UseInventoryResult {
  const [resources, setResources] = useState<Resource[]>([]);
  const [environments, setEnvironments] = useState<Resource[]>([]);
  const [counts, setCounts] = useState<ResourceCounts | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [loadingLabel, setLoadingLabel] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);

  const refresh = useCallback(async () => {
    setIsLoading(true);
    setLoadingLabel('Loading resources…');
    setError(null);

    try {
      const [fetchedResources, fetchedEnvironments] = await Promise.all([
        fetchResources((page, count) => {
          if (page > 1) setLoadingLabel(`Loading resources… (${count.toLocaleString()} so far)`);
        }),
        fetchEnvironments(),
      ]);

      // Build a map of envId (lower-cased) → instanceUrl from the full environment
      // properties. The JOIN projection in fetchResources may not support deeply
      // nested paths, so this is the reliable source of truth.
      const instanceUrlByEnv = new Map<string, string>();
      for (const env of fetchedEnvironments) {
        const key = env.name.toLowerCase();
        const p = env.properties as Record<string, unknown>;
        const linked = p.linkedEnvironmentMetadata as { instanceUrl?: string } | undefined;
        if (linked?.instanceUrl) instanceUrlByEnv.set(key, linked.instanceUrl);
      }

      // Collect unique owner GUIDs and resolve them to display names
      setLoadingLabel('Resolving owner names…');
      const ownerGuids = [...new Set(fetchedResources.map(extractOwnerGuid).filter(Boolean) as string[])];
      const nameMap = await resolveUserIds(ownerGuids);

      // Enrich resources with resolved owner name + reliable environmentInstanceUrl
      const enriched = fetchedResources.map((r) => {
        const guid = extractOwnerGuid(r);
        const resolved = guid ? nameMap.get(guid) : undefined;
        const envUrl = r.joinKey ? instanceUrlByEnv.get(r.joinKey) : undefined;

        if ((!resolved || resolved === guid) && !envUrl) return r;
        return {
          ...r,
          ...(envUrl ? { environmentInstanceUrl: envUrl } : {}),
          properties: {
            ...r.properties,
            ...(resolved && resolved !== guid ? { resolvedOwnerName: resolved } : {}),
          },
        };
      });

      setResources(enriched);
      setEnvironments(fetchedEnvironments);
      setCounts(computeResourceCounts(enriched));
    } catch (e: unknown) {
      setError(
        e instanceof Error ? extractMessage(e.message) : 'Failed to load Power Platform inventory.',
      );
    } finally {
      setIsLoading(false);
      setLoadingLabel(null);
    }
  }, []);

  useEffect(() => {
    void refresh();
  }, [refresh]);

  return { resources, environments, counts, isLoading, loadingLabel, error, refresh };
}
