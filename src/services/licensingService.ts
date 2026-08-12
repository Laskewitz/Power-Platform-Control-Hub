import type { IOperationResult } from '@microsoft/power-apps/data';
import { PowerPlatformforAdminsV2Service } from '../generated/services/PowerPlatformforAdminsV2Service.ts';
import type { Resource } from '../types/inventory.ts';

const API_VERSION = '2024-10-01';
const COPILOT_CREDIT_CURRENCY = 'MCSMessages';
const CONNECTOR_TIMEOUT_MS = 15_000;

export interface CopilotCreditSummary {
  purchased: number;
  allocated: number;
  consumed: number;
  lastUpdated?: string;
}

export interface HarnessEnvironmentLicensing {
  environmentId: string;
  environmentName: string;
  environmentType: string;
  agentCount: number;
  ownerCount: number;
  allocatedCredits?: number;
  billingPolicyName?: string;
  billingPolicyEnabled: boolean;
  allocationError?: string;
  billingPolicyError?: string;
}

function operationError<T>(result: IOperationResult<T>): string | undefined {
  if (result.success && !result.error) return undefined;
  return result.error?.message ?? 'The connector action did not complete successfully.';
}

interface ConnectorOutcome<T> {
  result?: IOperationResult<T>;
  error?: string;
}

async function settleConnectorRequest<T>(
  request: Promise<IOperationResult<T>>,
): Promise<ConnectorOutcome<T>> {
  let timeoutId: ReturnType<typeof setTimeout> | undefined;
  const timeout = new Promise<ConnectorOutcome<T>>((resolve) => {
    timeoutId = setTimeout(
      () => resolve({ error: 'This information is taking too long to load. Try refreshing.' }),
      CONNECTOR_TIMEOUT_MS,
    );
  });

  try {
    return await Promise.race([
      request.then((result) => ({ result })).catch((reason: unknown) => ({
        error: reason instanceof Error ? reason.message : 'The connector request failed.',
      })),
      timeout,
    ]);
  } finally {
    if (timeoutId) clearTimeout(timeoutId);
  }
}

function isHarnessAgent(resource: Resource): boolean {
  const value = resource.properties.isCLIAgent;
  return value === true || (typeof value === 'string' && value.toLowerCase() === 'true');
}

export async function fetchCopilotCreditSummary(): Promise<CopilotCreditSummary> {
  const result = await PowerPlatformforAdminsV2Service.ListCurrencyReports(API_VERSION, true, true);
  const error = operationError(result);
  if (error) throw new Error(error);

  const report = result.data.find((item) => item.currencyType === COPILOT_CREDIT_CURRENCY);
  return {
    purchased: report?.purchased ?? 0,
    allocated: report?.allocated ?? 0,
    consumed: report?.consumed?.unitsConsumed ?? 0,
    lastUpdated: report?.consumed?.lastUpdatedDay,
  };
}

export function getHarnessAgents(resources: Resource[]): Resource[] {
  return resources.filter(
    (resource) =>
      resource.type.toLowerCase() === 'microsoft.copilotstudio/agents' &&
      isHarnessAgent(resource),
  );
}

export async function fetchHarnessEnvironmentLicensing(
  environments: Resource[],
  resources: Resource[],
): Promise<HarnessEnvironmentLicensing[]> {
  const harnessAgents = getHarnessAgents(resources);
  const agentsByEnvironment = new Map<string, Resource[]>();

  for (const agent of harnessAgents) {
    const environmentId = agent.properties.environmentId;
    if (typeof environmentId !== 'string' || !environmentId) continue;
    const key = environmentId.toLowerCase();
    agentsByEnvironment.set(key, [...(agentsByEnvironment.get(key) ?? []), agent]);
  }

  return Promise.all(
    [...agentsByEnvironment.entries()].map(async ([key, agents]) => {
      const environment = environments.find((item) => item.name.toLowerCase() === key);
      const environmentId = environment?.name ?? String(agents[0]?.properties.environmentId ?? '');
      const [allocationOutcome, billingPolicyOutcome] = await Promise.all([
        settleConnectorRequest(
          PowerPlatformforAdminsV2Service.GetCurrencyAllocationByEnvironment(environmentId, API_VERSION),
        ),
        settleConnectorRequest(
          PowerPlatformforAdminsV2Service.GetEnvironmentBillingPolicy(environmentId, API_VERSION),
        ),
      ]);

      const allocationError = allocationOutcome.error
        ?? (allocationOutcome.result ? operationError(allocationOutcome.result) : 'Reserved credits are unavailable.');
      const billingPolicyError = billingPolicyOutcome.error
        ?? (billingPolicyOutcome.result ? operationError(billingPolicyOutcome.result) : 'Billing information is unavailable.');
      const allocation = allocationOutcome.result?.data?.currencyAllocations?.find(
        (item) => item.currencyType === COPILOT_CREDIT_CURRENCY,
      );
      const ownerIds = new Set(
        agents
          .map((agent) => agent.properties.ownerId)
          .filter((ownerId): ownerId is string => typeof ownerId === 'string' && ownerId.length > 0),
      );

      return {
        environmentId,
        environmentName:
          environment?.properties.displayName ??
          agents[0]?.environmentName ??
          environmentId,
        environmentType:
          environment?.properties.environmentType ??
          environment?.environmentType ??
          'Unknown',
        agentCount: agents.length,
        ownerCount: ownerIds.size,
        allocatedCredits: allocationError ? undefined : allocation?.allocated ?? 0,
        billingPolicyName: billingPolicyError ? undefined : billingPolicyOutcome.result?.data?.name,
        billingPolicyEnabled:
          !billingPolicyError && billingPolicyOutcome.result?.data?.status === 'Enabled',
        allocationError,
        billingPolicyError,
      };
    }),
  );
}

export async function updateCopilotCreditAllocation(
  environmentId: string,
  allocatedCredits: number,
): Promise<number> {
  const result = await PowerPlatformforAdminsV2Service.PatchCurrencyAllocationByEnvironment(
    environmentId,
    API_VERSION,
    {
      currencyAllocations: [
        {
          currencyType: COPILOT_CREDIT_CURRENCY,
          allocated: allocatedCredits,
        },
      ],
    },
  );
  const error = operationError(result);
  if (error) throw new Error(error);

  return result.data.currencyAllocations?.find(
    (item) => item.currencyType === COPILOT_CREDIT_CURRENCY,
  )?.allocated ?? allocatedCredits;
}
