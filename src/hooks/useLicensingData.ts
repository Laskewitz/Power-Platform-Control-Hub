import { useCallback, useEffect, useState } from 'react';
import { fetchLicensingSnapshot } from '../services/adminApi.ts';
import type { LicensingSnapshot } from '../types/admin.ts';
import { extractMessage } from '../utils/errorUtils.ts';

export interface UseLicensingDataResult {
  snapshot: LicensingSnapshot | null;
  isLoading: boolean;
  error: string | null;
  refresh: () => Promise<void>;
  loadPeriod: (startDate: string, endDate: string) => Promise<void>;
}

function createCurrentPeriod(): { startDate: string; endDate: string } {
  const end = new Date();
  const start = new Date(end);
  start.setUTCDate(start.getUTCDate() - 29);
  return {
    startDate: start.toISOString().slice(0, 10),
    endDate: end.toISOString().slice(0, 10),
  };
}

export function useLicensingData(): UseLicensingDataResult {
  const [period, setPeriod] = useState(createCurrentPeriod);
  const [snapshot, setSnapshot] = useState<LicensingSnapshot | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const loadPeriod = useCallback(async (startDate: string, endDate: string) => {
    setIsLoading(true);
    setError(null);
    try {
      setSnapshot(await fetchLicensingSnapshot(startDate, endDate));
      setPeriod({ startDate, endDate });
    } catch (reason: unknown) {
      setSnapshot(null);
      setError(
        reason instanceof Error
          ? extractMessage(reason.message)
          : 'Failed to load licensing and capacity data.',
      );
    } finally {
      setIsLoading(false);
    }
  }, []);

  const refresh = useCallback(
    () => loadPeriod(period.startDate, period.endDate),
    [loadPeriod, period.endDate, period.startDate],
  );

  useEffect(() => {
    void loadPeriod(period.startDate, period.endDate);
    // The initial reporting period is fixed for the lifetime of this mount.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [loadPeriod]);

  return { snapshot, isLoading, error, refresh, loadPeriod };
}
