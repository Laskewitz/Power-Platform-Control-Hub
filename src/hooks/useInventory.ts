import { useCallback, useEffect, useRef, useState } from 'react';
import {
  computeResourceCounts,
  fetchEnvironments,
  fetchResources,
} from '../services/inventoryApi.ts';
import type { Resource, ResourceCounts } from '../types/inventory.ts';
import { extractMessage } from '../utils/errorUtils.ts';
import { fetchPowerPagesWebsitesForEnvironments } from '../services/adminApi.ts';

export interface UseInventoryResult {
  resources: Resource[];
  environments: Resource[];
  counts: ResourceCounts | null;
  isLoading: boolean;
  loadingLabel: string | null;
  error: string | null;
  refresh: () => Promise<void>;
}

export function useInventory(): UseInventoryResult {
  const [resources, setResources] = useState<Resource[]>([]);
  const [environments, setEnvironments] = useState<Resource[]>([]);
  const [counts, setCounts] = useState<ResourceCounts | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [loadingLabel, setLoadingLabel] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const refreshGeneration = useRef(0);

  const refresh = useCallback(async () => {
    const generation = ++refreshGeneration.current;
    setIsLoading(true);
    setLoadingLabel('Loading resources…');
    setError(null);

    try {
      const [fetchedResources, fetchedEnvironments] = await Promise.all([
        fetchResources((page, count) => {
          if (refreshGeneration.current === generation && page > 1) {
            setLoadingLabel(`Loading resources… (${count.toLocaleString()} so far)`);
          }
        }),
        fetchEnvironments(),
      ]);

      if (refreshGeneration.current !== generation) return;

      // Make the core inventory interactive before optional enrichment completes.
      setResources(fetchedResources);
      setEnvironments(fetchedEnvironments);
      setCounts(computeResourceCounts(fetchedResources));
      setIsLoading(false);
      setLoadingLabel(null);

      void fetchPowerPagesWebsitesForEnvironments(fetchedEnvironments).then((websiteResources) => {
        if (refreshGeneration.current !== generation) return;
        const allResources = [...fetchedResources, ...websiteResources];
        setResources(allResources);
        setCounts(computeResourceCounts(allResources));
      }).catch(() => {
        // Power Pages enrichment is optional and must not block inventory.
      });
    } catch (e: unknown) {
      if (refreshGeneration.current === generation) {
        setError(
          e instanceof Error ? extractMessage(e.message) : 'Failed to load Power Platform inventory.',
        );
      }
    } finally {
      if (refreshGeneration.current === generation) {
        setIsLoading(false);
        setLoadingLabel(null);
      }
    }
  }, []);

  useEffect(() => {
    void refresh();
    return () => {
      refreshGeneration.current += 1;
    };
  }, [refresh]);

  return { resources, environments, counts, isLoading, loadingLabel, error, refresh };
}
