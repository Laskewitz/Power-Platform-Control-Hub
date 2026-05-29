import type { IOperationResult } from '@microsoft/power-apps/data';
import { PowerPlatformforAdminsV2Service } from '../generated/services/PowerPlatformforAdminsV2Service.ts';

const API = '2024-10-01';

function unwrapOperationResult<T>(result: IOperationResult<T>): T {
  if (!result.success || result.error) {
    throw new Error(result.error?.message ?? 'Operation failed');
  }
  return result.data;
}

export async function enableEnvironment(environmentId: string): Promise<void> {
  const result = await PowerPlatformforAdminsV2Service.EnableEnvironment(environmentId, API);
  unwrapOperationResult(result);
}

export async function disableEnvironment(environmentId: string, reason = 'UserRequested'): Promise<void> {
  const result = await PowerPlatformforAdminsV2Service.DisableEnvironment(environmentId, API, undefined, undefined, { reason });
  unwrapOperationResult(result);
}

export async function enableManagedEnvironment(environmentId: string): Promise<void> {
  const result = await PowerPlatformforAdminsV2Service.EnableManagedEnvironment(environmentId, API);
  unwrapOperationResult(result);
}

export async function disableManagedEnvironment(environmentId: string): Promise<void> {
  const result = await PowerPlatformforAdminsV2Service.DisableManagedEnvironment(environmentId, API);
  unwrapOperationResult(result);
}

export async function createEnvironmentBackup(environmentId: string, notes: string): Promise<void> {
  const result = await PowerPlatformforAdminsV2Service.CreateEnvironmentBackup(
    environmentId,
    API,
    { label: notes },
  );
  unwrapOperationResult(result);
}
