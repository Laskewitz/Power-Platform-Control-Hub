import type { IOperationResult } from '@microsoft/power-apps/data';
import type {
  AdvisorRecommendation as GeneratedAdvisorRecommendation,
  BillingPolicyResponseModel,
  Connection as GeneratedConnection,
  CurrencyReportV2,
  CrossTenantConnectionReport,
  EnvironmentGroup as GeneratedEnvironmentGroup,
  GetConnectorByIdResponse,
  TenantCapacityDetailsModel,
  Policy as GeneratedPolicy,
  Principal,
  RuleSetDto,
  WebsiteDto,
} from '../generated/models/PowerPlatformforAdminsV2Model.ts';
import { PowerPlatformforAdminsV2Service } from '../generated/services/PowerPlatformforAdminsV2Service.ts';
import type {
  AdvisorRecommendation,
  BillingPolicy,
  Connection,
  CrossTenantReport,
  EnvironmentGroup,
  FlowComplianceSummary,
  GovernanceRuleSet,
  LicenseEntitlement,
  LicensingSnapshot,
  PolicyRuleAssignment,
  PowerPagesWebsite,
  RoleAssignment,
  RuleBasedPolicy,
  UserCapacitySummary,
} from '../types/admin.ts';
import type { Resource } from '../types/inventory.ts';

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
    status: (policy.status ?? 'Disabled') as BillingPolicy['status'],
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
      expirationTime: item.properties?.expirationTime,
      apiId: item.properties?.apiId,
      isSsoConnection: item.properties?.isSsoConnection,
      createdBy: item.properties?.createdBy
        ? {
            displayName: item.properties.createdBy.displayName,
            email: item.properties.createdBy.email ?? item.properties.createdBy.userPrincipalName,
          }
        : undefined,
      statuses: item.properties?.statuses?.map((status) => ({
        status: status.status,
        target: status.target,
        error: status.error
          ? {
              code: status.error.code,
              message: status.error.message,
            }
          : undefined,
      })),
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
    type: (website.type ?? 'Trial') as 'Trial' | 'Production',
    trialExpiringInDays: website.trialExpiringInDays,
    packageVersion: website.packageVersion,
    siteVisibility: (website.siteVisibility ?? 'private') as 'public' | 'private',
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

function mapEntitlements(details: TenantCapacityDetailsModel): LicenseEntitlement[] {
  return (details.tenantCapacities ?? []).flatMap((capacity) =>
    (capacity.capacityEntitlements ?? []).flatMap((entitlement) =>
      (entitlement.licenses ?? []).map((license) => ({
        capacityType: capacity.capacityType ?? entitlement.capacityType ?? 'Unknown',
        capacitySubType: entitlement.capacitySubType ?? 'Standard',
        entitlementCode: license.entitlementCode ?? '',
        displayName: license.displayName ?? license.entitlementCode ?? 'Unnamed entitlement',
        skuId: license.skuId,
        servicePlanId: license.servicePlanId,
        paidEnabled: license.paid?.enabled ?? 0,
        trialEnabled: license.trial?.enabled ?? 0,
        totalCapacity: license.totalCapacity ?? entitlement.totalCapacity ?? 0,
        isTemporary: license.isTemporaryLicense ?? false,
        expiryDate: license.temporaryLicenseExpiryDate,
        nextLifecycleDate: license.nextLifecycleDate,
        capabilityStatus: license.capabilityStatus ?? 'Unknown',
      })),
    ),
  );
}

function mapCurrencies(reports: CurrencyReportV2[]) {
  return reports.map((report) => ({
    currencyType: report.currencyType ?? 'Unknown',
    purchased: report.purchased ?? 0,
    allocated: report.allocated ?? 0,
    consumed: report.consumed?.unitsConsumed ?? 0,
    lastUpdatedDay: report.consumed?.lastUpdatedDay,
  }));
}

export async function fetchLicensingSnapshot(
  startDate: string,
  endDate: string,
): Promise<LicensingSnapshot> {
  const [capacityResult, currencyResult, complianceResult, usersResult] = await Promise.allSettled([
    PowerPlatformforAdminsV2Service.GetTenantCapacityDetails(API_VERSION),
    PowerPlatformforAdminsV2Service.ListCurrencyReports(API_VERSION, true, true),
    PowerPlatformforAdminsV2Service.GetUserPerFlowCapacitySourceTenantContextSummary(
      startDate,
      API_VERSION,
      endDate,
    ),
    PowerPlatformforAdminsV2Service.GetUserPerFlowCapacitySourceUserContextSummary(
      startDate,
      API_VERSION,
      endDate,
      1,
      200,
    ),
  ]);

  const warnings: string[] = [];
  const readSettled = <T>(
    result: PromiseSettledResult<IOperationResult<T>>,
    label: string,
  ): T | undefined => {
    if (result.status === 'rejected') {
      warnings.push(`${label}: ${result.reason instanceof Error ? result.reason.message : 'request failed'}`);
      return undefined;
    }
    if (!result.value.success || result.value.error) {
      warnings.push(`${label}: ${result.value.error?.message ?? 'request failed'}`);
      return undefined;
    }
    return result.value.data;
  };

  const capacity = readSettled(capacityResult, 'Tenant capacity');
  const currencies = readSettled(currencyResult, 'Currency report');
  const compliance = readSettled(complianceResult, 'Licensing compliance');
  const users = readSettled(usersResult, 'User capacity');

  if (!capacity && !currencies && !compliance && !users) {
    throw new Error('The licensing and capacity APIs did not return data for this tenant.');
  }

  const mappedCompliance: FlowComplianceSummary[] = (compliance ?? []).map((item) => ({
    flowContext: item.flowContext ?? 'All flow contexts',
    usersInCompliance: item.countOfUsersInCompliance ?? 0,
    usersExceedingCapacity: item.countOfUsersExceedingCapacity ?? 0,
    usersWithoutLicense: item.countOfUsersWithoutALicense ?? 0,
    usersWithoutPremiumLicense: item.countOfUsersWithoutPremiumLicenseUsingPremiumFeatures ?? 0,
  }));
  const mappedUsers: UserCapacitySummary[] = (users?.records ?? []).map((item) => ({
    userId: item.userId ?? '',
    flowContext: item.flowContext ?? 'Unknown',
    licenseCategorization: item.flowLicenseCategorization ?? 'Unknown',
    totalConsumption: item.totalConsumption ?? 0,
    totalCapacity: item.totalCapacity ?? 0,
    totalFlows: item.totalFlows ?? 0,
  }));

  return {
    tenantId: capacity?.tenantId ?? '',
    licenseModelType: capacity?.licenseModelType ?? 'Unknown',
    capacityStatus: capacity?.capacitySummary?.status ?? 'Unknown',
    capacityStatusMessage: capacity?.capacitySummary?.statusMessage ?? '',
    finOpsStatus: capacity?.capacitySummary?.finOpsStatus ?? 'Unknown',
    finOpsStatusMessage: capacity?.capacitySummary?.finOpsStatusMessage ?? '',
    hasTemporaryLicense: capacity?.temporaryLicenseInfo?.hasTemporaryLicense ?? false,
    temporaryLicenseExpiryDate: capacity?.temporaryLicenseInfo?.temporaryLicenseExpiryDate,
    periodStart: startDate,
    periodEnd: endDate,
    capacities: (capacity?.tenantCapacities ?? []).map((item) => ({
      capacityType: item.capacityType ?? 'Unknown',
      units: item.capacityUnits ?? '',
      totalCapacity: item.totalCapacity ?? 0,
      maxCapacity: item.maxCapacity ?? item.totalCapacity ?? 0,
      consumed: item.consumption?.actual ?? 0,
      ratedConsumption: item.consumption?.rated,
      status: item.status ?? 'Unknown',
      updatedOn: item.consumption?.actualUpdatedOn ?? item.consumption?.ratedUpdatedOn,
    })),
    entitlements: capacity ? mapEntitlements(capacity) : [],
    currencies: mapCurrencies(currencies ?? []),
    compliance: mappedCompliance,
    users: mappedUsers,
    warnings,
  };
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

const POWER_PAGES_RESOURCE_TYPE = 'microsoft.powerpages/websites';

function mapWebsiteToResource(website: PowerPagesWebsite, environment?: Resource): Resource {
  return {
    id: website.id,
    name: website.name,
    type: POWER_PAGES_RESOURCE_TYPE,
    location: environment?.environmentRegion ?? environment?.location,
    properties: {
      displayName: website.name,
      createdAt: website.createdOn,
      environmentId: website.environmentId,
      websiteUrl: website.websiteUrl,
      siteVisibility: website.siteVisibility,
      websiteType: website.type,
      status: website.status,
      packageVersion: website.packageVersion,
    },
    environmentName: website.environmentName || environment?.properties.displayName,
    environmentRegion: environment?.environmentRegion ?? environment?.location,
  };
}

/**
 * Fetches Power Pages websites across every supplied environment and normalizes
 * them into the shared Resource shape so they can appear alongside other
 * resources in the inventory. Failures for individual environments are ignored
 * so a single inaccessible environment doesn't break the whole load.
 */
export async function fetchPowerPagesWebsitesForEnvironments(environments: Resource[]): Promise<Resource[]> {
  const envByName = new Map(environments.map((environment) => [environment.name, environment]));
  const settled = await Promise.allSettled(
    environments.map((environment) => fetchPowerPagesWebsites(environment.name)),
  );

  const resources: Resource[] = [];
  settled.forEach((result, index) => {
    if (result.status !== 'fulfilled') return;
    const fallbackEnv = environments[index];
    for (const website of result.value) {
      resources.push(mapWebsiteToResource(website, envByName.get(website.environmentId) ?? fallbackEnv));
    }
  });

  return resources;
}

export async function addEnvironmentToGroup(groupId: string, environmentId: string): Promise<void> {
  const result = await PowerPlatformforAdminsV2Service.AddEnvironmentToGroup(groupId, environmentId, API_VERSION);
  unwrapOperationResult(result);
}

export async function removeEnvironmentFromGroup(groupId: string, environmentId: string): Promise<void> {
  const result = await PowerPlatformforAdminsV2Service.RemoveEnvironmentFromGroup(groupId, environmentId, API_VERSION);
  unwrapOperationResult(result);
}

function mapRuleBasedPolicy(p: GeneratedPolicy): RuleBasedPolicy {
  return {
    id: p.id ?? '',
    name: p.name ?? '',
    lastModified: p.lastModified ?? '',
    ruleSetCount: p.ruleSetCount ?? 0,
    ruleSets: (p.ruleSets ?? []).map((rs) => ({
      id: rs.id ?? '',
      version: rs.version ?? '',
      inputs: rs.inputs ?? {},
    })),
  };
}

export async function fetchRuleBasedPolicies(): Promise<RuleBasedPolicy[]> {
  const result = await PowerPlatformforAdminsV2Service.ListRuleBasedPolicies(API_VERSION);
  return (unwrapOperationResult(result).value ?? []).map(mapRuleBasedPolicy);
}

export async function fetchAllRuleAssignments(): Promise<PolicyRuleAssignment[]> {
  const result = await PowerPlatformforAdminsV2Service.ListRuleAssignments(true, API_VERSION);
  return (unwrapOperationResult(result).value ?? [])
    .filter((a) => a.policyId && a.resourceId)
    .map((a) => ({
      policyId: a.policyId!,
      resourceId: a.resourceId!,
      resourceType: (a.resourceType ?? 'NotSpecified') as PolicyRuleAssignment['resourceType'],
      ruleSetCount: a.ruleSetCount ?? 0,
    }));
}

function mapGovernanceRuleSet(rs: RuleSetDto): GovernanceRuleSet {
  return {
    id: rs.id ?? '',
    lastModified: rs.lastModified ?? '',
    environmentFilterType: rs.environmentFilter?.type ?? '',
    environmentFilterValues: (rs.environmentFilter?.values ?? []).map((v) => ({
      id: v.id ?? '',
      type: v.type ?? '',
      name: v.name ?? '',
    })),
    parameters: (rs.parameters ?? []).map((p) => ({
      type: p.type ?? '',
      resourceType: p.resourceType ?? '',
      rules: (p.value ?? []).map((r) => ({ id: r.id ?? '', value: r.value ?? '' })),
    })),
  };
}

export async function fetchRuleSets(): Promise<GovernanceRuleSet[]> {
  const result = await PowerPlatformforAdminsV2Service.GetRuleSetListForTenant(API_VERSION);
  return (unwrapOperationResult(result).value ?? []).map(mapGovernanceRuleSet);
}
