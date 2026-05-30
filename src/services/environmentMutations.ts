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

/** Module-level cache so we don't call ListRoleDefinitions more than once. */
let cachedSysAdminRoleId: string | undefined;

async function getSysAdminRoleId(): Promise<string> {
  if (cachedSysAdminRoleId) return cachedSysAdminRoleId;
  const rolesResult = await PowerPlatformforAdminsV2Service.ListRoleDefinitions(API);
  const roles = unwrapOperationResult(rolesResult);
  const role = roles.value?.find((r) => r.roleDefinitionName?.toLowerCase() === 'system administrator');
  if (!role?.roleDefinitionId) throw new Error('Could not find System Administrator role definition.');
  cachedSysAdminRoleId = role.roleDefinitionId;
  return cachedSysAdminRoleId;
}

/**
 * Returns true if the currently signed-in user has the System Administrator
 * role on the given environment.
 */
export async function checkCurrentUserIsEnvAdmin(environmentId: string): Promise<boolean> {
  if (!environmentId) return false;
  const [ctx, sysAdminRoleId] = await Promise.all([getContext(), getSysAdminRoleId()]);
  const objectId = ctx.user.objectId?.toLowerCase();
  if (!objectId) return false;

  const result = await PowerPlatformforAdminsV2Service.ListEnvironmentRoleAssignments(environmentId, API);
  // If the call fails (no access), treat as "unknown" — don't surface a banner
  if (!result.success) return true;
  const assignments = result.data;
  return (
    assignments?.value?.some(
      (a) =>
        a.principalObjectId?.toLowerCase() === objectId &&
        a.roleDefinitionId === sysAdminRoleId,
    ) ?? false
  );
}

/**
 * Adds the currently signed-in user as System Administrator on the given environment.
 * Looks up the System Administrator role definition by name from ListRoleDefinitions,
 * then calls CreateEnvironmentRoleAssignment.
 */
export async function addSelfAsEnvironmentAdmin(environmentId: string): Promise<void> {
  const ctx = await getContext();
  const objectId = ctx.user.objectId;
  if (!objectId) throw new Error('Could not determine current user object ID.');

  const sysAdminRoleId = await getSysAdminRoleId();

  const assignResult = await PowerPlatformforAdminsV2Service.CreateEnvironmentRoleAssignment(
    environmentId,
    API,
    {
      principalObjectId: objectId,
      roleDefinitionId: sysAdminRoleId,
      principalType: 'User',
    },
  );
  unwrapOperationResult(assignResult);
}
