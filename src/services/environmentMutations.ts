import type { IOperationResult } from '@microsoft/power-apps/data';
import { getContext } from '@microsoft/power-apps/app';
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

/**
 * Returns the roleAssignmentId if the current user has an applied Environment Admin
 * role on the given environment (granted via ApplyAdminRole), or null if not.
 */
export async function getCurrentUserRoleAssignmentId(environmentId: string): Promise<string | null> {
  if (!environmentId) return null;
  const [ctx, result] = await Promise.all([
    getContext(),
    PowerPlatformforAdminsV2Service.ListEnvironmentRoleAssignments(environmentId, API),
  ]);
  if (!result.success) return null;
  const objectId = ctx.user.objectId?.toLowerCase();
  if (!objectId) return null;
  const assignment = result.data?.value?.find(
    (a) => a.principalObjectId?.toLowerCase() === objectId,
  );
  return assignment?.roleAssignmentId ?? null;
}

/**
 * Adds the currently signed-in user as System Administrator on the given environment
 * using the dedicated ApplyAdminRole API (no role definition lookup required).
 */
export async function addSelfAsEnvironmentAdmin(environmentId: string): Promise<void> {
  const result = await PowerPlatformforAdminsV2Service.ApplyAdminRole(environmentId, API);
  unwrapOperationResult(result);
}

/**
 * Removes the current user's applied admin role from the given environment.
 * Looks up the user's role assignment ID first, then deletes it.
 */
export async function removeSelfAdminRole(environmentId: string): Promise<void> {
  const assignmentId = await getCurrentUserRoleAssignmentId(environmentId);
  if (!assignmentId) throw new Error('No admin role assignment found to remove.');
  const result = await PowerPlatformforAdminsV2Service.DeleteEnvironmentRoleAssignment(
    environmentId,
    assignmentId,
    API,
  );
  unwrapOperationResult(result);
}
