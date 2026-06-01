import type {
  AdminFlowWithConnectionReferences,
  ResourceArray_FlowPermission,
  ModifyFlowPermissionPayload,
} from '../generated/models/PowerAutomateManagementModel.ts';
import { PowerAutomateManagementService } from '../generated/services/PowerAutomateManagementService.ts';
import type { IOperationResult } from '@microsoft/power-apps/data';
import type { FlowDefinition } from './flowAnalyzer.ts';

function unwrap<T>(result: IOperationResult<T>): T {
  if (!result.success || result.error) {
    const rawMsg = result.error?.message ?? 'Operation failed';
    let friendlyMsg = rawMsg;
    try {
      const parsed = JSON.parse(rawMsg) as {
        error?: { code?: string; message?: string };
        message?: string;
        code?: string;
      };
      if (parsed?.error?.message) {
        friendlyMsg = parsed.error.message;
        if (parsed.error.code) friendlyMsg = `[${parsed.error.code}] ${friendlyMsg}`;
      } else if (parsed?.message) {
        friendlyMsg = parsed.message;
        if (parsed.code) friendlyMsg = `[${parsed.code}] ${friendlyMsg}`;
      }
    } catch {
      // Not JSON — use as-is
    }
    throw new Error(friendlyMsg);
  }
  return result.data;
}

/** AdminGetFlow response extended with the full flow definition (present when includeFlowDefinition=true). */
export interface AdminFlowWithDefinition extends AdminFlowWithConnectionReferences {
  properties: AdminFlowWithConnectionReferences['properties'] & {
    definition?: FlowDefinition;
  };
}

/**
 * Fetch basic flow details (no definition) — fast path for initial render.
 */
export async function getFlowDetailsOnly(
  environmentName: string,
  flowName: string,
): Promise<AdminFlowWithDefinition> {
  const result = await PowerAutomateManagementService.AdminGetFlow(environmentName, flowName, false);
  return unwrap(result) as AdminFlowWithDefinition;
}

/**
 * Fetch the flow definition for analysis. Returns null if not available
 * (some flow types do not expose their definition via the admin API).
 */
export async function getFlowDefinitionForAnalysis(
  environmentName: string,
  flowName: string,
): Promise<FlowDefinition | null> {
  try {
    const result = await PowerAutomateManagementService.AdminGetFlow(environmentName, flowName, true);
    const data = unwrap(result) as AdminFlowWithDefinition;
    return data.properties?.definition ?? null;
  } catch {
    return null;
  }
}

/**
 * @deprecated Use getFlowDetailsOnly + getFlowDefinitionForAnalysis for better performance.
 * Fetch full flow details as admin, always including the flow definition
 * so static analysis and best-practice checks can be performed on it.
 * Falls back to fetching without the definition if the API rejects the request
 * (e.g. HTTP-triggered flows with token security return 400 for definition requests).
 */
export async function getFlowAsAdmin(
  environmentName: string,
  flowName: string,
): Promise<AdminFlowWithDefinition> {
  try {
    const result = await PowerAutomateManagementService.AdminGetFlow(environmentName, flowName, true);
    return unwrap(result) as AdminFlowWithDefinition;
  } catch {
    // Some flow types reject includeFlowDefinition=true — retry without it
    const result = await PowerAutomateManagementService.AdminGetFlow(environmentName, flowName, false);
    return unwrap(result) as AdminFlowWithDefinition;
  }
}

export async function enableFlow(environmentName: string, flowName: string): Promise<void> {
  const result = await PowerAutomateManagementService.StartFlow(environmentName, flowName);
  unwrap(result);
}

export async function disableFlow(environmentName: string, flowName: string): Promise<void> {
  const result = await PowerAutomateManagementService.StopFlow(environmentName, flowName);
  unwrap(result);
}

export async function deleteFlow(environmentName: string, flowName: string): Promise<void> {
  const result = await PowerAutomateManagementService.DeleteFlow(environmentName, flowName);
  unwrap(result);
}

export async function listFlowOwners(
  environmentName: string,
  flowName: string,
): Promise<ResourceArray_FlowPermission> {
  const result = await PowerAutomateManagementService.ListFlowOwners(environmentName, flowName);
  return unwrap(result);
}

export async function modifyFlowOwners(
  environmentName: string,
  flowName: string,
  payload: ModifyFlowPermissionPayload,
): Promise<void> {
  const result = await PowerAutomateManagementService.AdminModifyFlowOwners(environmentName, flowName, payload);
  unwrap(result);
}

export async function listRunOnlyUsers(
  environmentName: string,
  flowName: string,
): Promise<ResourceArray_FlowPermission> {
  const result = await PowerAutomateManagementService.ListFlowUsers(environmentName, flowName);
  return unwrap(result);
}

export async function modifyRunOnlyUsers(
  environmentName: string,
  flowName: string,
  payload: ModifyFlowPermissionPayload,
): Promise<void> {
  const result = await PowerAutomateManagementService.ModifyRunOnlyUsers(environmentName, flowName, payload);
  unwrap(result);
}
