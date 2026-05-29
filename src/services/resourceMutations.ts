import type { IOperationResult } from '@microsoft/power-apps/data';
import { PowerPlatformforAdminsV2Service } from '../generated/services/PowerPlatformforAdminsV2Service.ts';

const API = '2024-10-01';

function unwrapOperationResult<T>(result: IOperationResult<T>): T {
  if (!result.success || result.error) {
    throw new Error(result.error?.message ?? 'Operation failed');
  }
  return result.data;
}

export async function deleteFlow(flowId: string): Promise<void> {
  const result = await PowerPlatformforAdminsV2Service.DeleteFlow(flowId, API);
  unwrapOperationResult(result);
}

export async function deleteCopilotAgent(environmentId: string, botId: string): Promise<void> {
  const result = await PowerPlatformforAdminsV2Service.DeleteCopilotAgent(environmentId, botId, API);
  if (!result.success) {
    // BotNotFound (ErrorCode 4105) means already deleted — treat as success
    const msg = result.error?.message ?? '';
    if (msg.includes('BotNotFound') || msg.includes('4105') || msg.includes('Bot not found')) return;
    throw new Error(msg || 'Operation failed');
  }
}
