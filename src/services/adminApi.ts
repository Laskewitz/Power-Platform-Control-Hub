import type { IOperationResult } from '@microsoft/power-apps/data';
import type {
  AdvisorRecommendation as GeneratedAdvisorRecommendation,
  BillingPolicyResponseModel,
  Connection as GeneratedConnection,
  CrossTenantConnectionReport,
  EnvironmentGroup as GeneratedEnvironmentGroup,
  GetConnectorByIdResponse,
  Principal,
  WebsiteDto,
} from '../generated/models/PowerPlatformforAdminsV2Model.ts';
import { PowerPlatformforAdminsV2Service } from '../generated/services/PowerPlatformforAdminsV2Service.ts';
import type {
  AdvisorRecommendation,
  BillingPolicy,
  Connection,
  CrossTenantReport,
  EnvironmentGroup,
  PowerPagesWebsite,
  RoleAssignment,
} from '../types/admin.ts';

const API_VERSION = '2024-10-01';

function unwrapOperationResult<T>(result: IOperationResult<T>): T {
  if (!result.success || result.error) {
    throw new Error(result.error?.message ?? 'Operation failed');
  }
  return result.data;
}

function mapPrincipal(principal?: Principal) {
  return {
    id: principal?.id ?? '',
    displayName: principal?.displayName,
    email: principal?.email,
    type: principal?.type ?? '',
    tenantId: principal?.tenantId,
  };
}

function mapRecommendation(recommendation: GeneratedAdvisorRecommendation): AdvisorRecommendation {
  return {
    scenario: recommendation.scenario ?? '',
    details: {
      resourceCount: recommendation.details?.resourceCount ?? 0,
      lastRefreshedTimestamp: recommendation.details?.lastRefreshedTimestamp ?? '',
      expectedNextRefreshTimestamp: recommendation.details?.expectedNextRefreshTimestamp ?? '',
    },
  };
}

function mapEnvironmentGroup(group: GeneratedEnvironmentGroup): EnvironmentGroup {
  return {
    id: group.id ?? '',
    displayName: group.displayName ?? '',
    description: group.description ?? '',
    parentGroupId: group.parentGroupId,
    childrenGroupIds: group.childrenGroupIds,
    createdTime: group.createdTime ?? '',
    createdBy: mapPrincipal(group.createdBy),
    lastModifiedTime: group.lastModifiedTime ?? '',
    lastModifiedBy: mapPrincipal(group.lastModifiedBy),
  };
}

function mapBillingPolicy(policy: BillingPolicyResponseModel): BillingPolicy {
  return {
    id: policy.id ?? '',
    name: policy.name ?? '',
    status: policy.status ?? 'Unknown',
    location: policy.location ?? '',
    billingInstrument: {
      subscriptionId: policy.billingInstrument?.subscriptionId ?? '',
      resourceGroup: policy.billingInstrument?.resourceGroup ?? '',
      id: policy.billingInstrument?.id ?? '',
    },
    createdOn: policy.createdOn ?? '',
    createdBy: {
      id: policy.createdBy?.id ?? '',
      type: policy.createdBy?.type ?? '',
    },
    lastModifiedOn: policy.lastModifiedOn ?? '',
  };
}

function mapCrossTenantReport(report: CrossTenantConnectionReport): CrossTenantReport {
  return {
    tenantId: report.tenantId,
    reportId: report.reportId,
    requestDate: report.requestDate,
    startDate: report.startDate,
    endDate: report.endDate,
    status: report.status,
    connections: report.connections?.map((connection) => ({
      tenantId: connection.tenantId,
      connectionType: connection.connectionType,
    })),
  };
}

function mapConnection(item: GeneratedConnection): Connection {
  return {
    name: item.name ?? '',
    id: item.id ?? '',
    type: item.type ?? '',
    properties: {
      displayName: item.properties?.displayName ?? '',
      iconUri: item.properties?.iconUri,
      createdTime: item.properties?.createdTime ?? '',
      changedTime: item.properties?.lastModifiedTime ?? '',
      description:
        typeof item.properties?.metadata?.description === 'string'
          ? item.properties.metadata.description
          : undefined,
    },
  };
}

function mapConnector(item: GetConnectorByIdResponse): Connection {
  return {
    name: item.name ?? '',
    id: item.id ?? '',
    type: item.type ?? '',
    properties: {
      displayName: item.properties?.displayName ?? '',
      iconUri: item.properties?.iconUri,
      iconBrandColor: item.properties?.iconBrandColor,
      isCustomApi: item.properties?.isCustomApi,
      description: item.properties?.description,
      createdTime: item.properties?.createdTime ?? '',
      changedTime: item.properties?.changedTime ?? '',
      tier: item.properties?.tier,
      publisher: item.properties?.publisher,
    },
  };
}

function mapWebsite(website: WebsiteDto): PowerPagesWebsite {
  return {
    id: website.id ?? '',
    name: website.name ?? '',
    createdOn: website.createdOn ?? '',
    templateName: website.templateName ?? '',
    websiteUrl: website.websiteUrl ?? '',
    tenantId: website.tenantId ?? '',
    dataverseInstanceUrl: website.dataverseInstanceUrl ?? '',
    environmentName: website.environmentName ?? '',
    environmentId: website.environmentId ?? '',
    selectedBaseLanguage: String(website.selectedBaseLanguage ?? ''),
    customHostNames: website.customHostNames,
    subdomain: website.subdomain,
    packageInstallStatus: website.packageInstallStatus ?? '',
    type: website.type ?? '',
    trialExpiringInDays: website.trialExpiringInDays,
    packageVersion: website.packageVersion,
    siteVisibility: website.siteVisibility ?? '',
    status: website.status ?? '',
  };
}

export async function fetchAdvisorRecommendations(): Promise<AdvisorRecommendation[]> {
  const result = await PowerPlatformforAdminsV2Service.GetRecommendations(API_VERSION);
  const recs = (unwrapOperationResult(result).value ?? []).map(mapRecommendation);

  // Fetch allowed actions for each scenario in parallel (best-effort — ignore failures)
  await Promise.all(
    recs.map(async (rec) => {
      try {
        const actionsResult = await PowerPlatformforAdminsV2Service.GetScenarioActions(rec.scenario, API_VERSION);
        if (actionsResult.success && !actionsResult.error) {
          rec.details.actions = (actionsResult.data ?? []).map((a) => ({
            actionType: a.actionType,
            actionName: a.actionName,
          }));
        }
      } catch {
        // Actions not available for this scenario — leave empty
      }
    }),
  );

  return recs;
}

export async function fetchRoleAssignments(): Promise<RoleAssignment[]> {
  const result = await PowerPlatformforAdminsV2Service.ListRoleAssignments(API_VERSION);
  return (unwrapOperationResult(result).value ?? []).map((assignment) => ({
    roleDefinitionId: assignment.roleDefinitionId ?? '',
    roleDefinitionName: assignment.roleDefinitionId ?? '',
    permissions: [],
  }));
}

export async function fetchEnvironmentGroups(): Promise<EnvironmentGroup[]> {
  const result = await PowerPlatformforAdminsV2Service.ListEnvironmentGroups(API_VERSION);
  return (unwrapOperationResult(result).value ?? []).map(mapEnvironmentGroup);
}

export async function fetchBillingPolicies(): Promise<BillingPolicy[]> {
  const result = await PowerPlatformforAdminsV2Service.ListBillingPolicies(API_VERSION);
  return (unwrapOperationResult(result).value ?? []).map(mapBillingPolicy);
}

export async function fetchCrossTenantReports(): Promise<CrossTenantReport[]> {
  const result = await PowerPlatformforAdminsV2Service.ListCrossTenantConnectionReports(API_VERSION);
  return (unwrapOperationResult(result).value ?? []).map(mapCrossTenantReport);
}

export async function fetchConnections(environmentId: string): Promise<Connection[]> {
  const result = await PowerPlatformforAdminsV2Service.ListConnections(environmentId, API_VERSION);
  return (unwrapOperationResult(result).value ?? []).map(mapConnection);
}

export async function fetchConnectors(environmentId: string): Promise<Connection[]> {
  const result = await PowerPlatformforAdminsV2Service.ListConnectors(environmentId, `environment eq '${environmentId}'`, API_VERSION);
  return (unwrapOperationResult(result).value ?? []).map(mapConnector);
}

export async function fetchPowerPagesWebsites(environmentId: string): Promise<PowerPagesWebsite[]> {
  const result = await PowerPlatformforAdminsV2Service.GetWebsites(environmentId, API_VERSION);
  return (unwrapOperationResult(result).value ?? []).map(mapWebsite);
}
