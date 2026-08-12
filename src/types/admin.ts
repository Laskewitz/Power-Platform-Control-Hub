export interface ODataList<T> {
  value: T[];
  nextLink?: string;
  '@odata.nextLink'?: string;
}

export interface AdvisorRecommendation {
  scenario: string;
  details: {
    resourceCount: number;
    lastRefreshedTimestamp: string;
    expectedNextRefreshTimestamp: string;
    actions?: { actionType?: string; actionName?: string }[];
  };
}

export interface RoleAssignment {
  roleDefinitionId: string;
  roleDefinitionName: string;
  permissions: string[];
}

export interface AuditActor {
  id: string;
  displayName?: string;
  email?: string;
  type: string;
  tenantId?: string;
}

export interface EnvironmentGroup {
  id: string;
  displayName: string;
  description: string;
  parentGroupId?: string;
  childrenGroupIds?: string[];
  createdTime: string;
  createdBy: AuditActor;
  lastModifiedTime: string;
  lastModifiedBy: AuditActor;
}

export interface BillingPolicyActor {
  id: string;
  type: string;
}

export interface BillingPolicy {
  id: string;
  name: string;
  status: 'Enabled' | 'Disabled';
  location: string;
  billingInstrument: {
    subscriptionId: string;
    resourceGroup: string;
    id: string;
  };
  createdOn: string;
  createdBy: BillingPolicyActor;
  lastModifiedOn: string;
}

export interface CrossTenantConnection {
  tenantId: string;
  connectionType: 'Inbound' | 'Outbound';
}

export interface RuleBasedPolicy {
  id: string;
  name: string;
  lastModified: string;
  ruleSetCount: number;
  ruleSets: PolicyRuleSet[];
}

export interface PolicyRuleSet {
  id: string;
  version: string;
  inputs: Record<string, unknown>;
}

export interface GovernanceRuleSet {
  id: string;
  lastModified: string;
  environmentFilterType: string;
  environmentFilterValues: { id: string; type: string; name: string }[];
  parameters: { type: string; resourceType: string; rules: { id: string; value: string }[] }[];
}

export interface PolicyRuleAssignment {
  policyId: string;
  resourceId: string;
  resourceType: 'NotSpecified' | 'EnvironmentGroup' | 'Environment';
  ruleSetCount: number;
}

export interface CrossTenantReport {
  tenantId: string;
  reportId: string;
  requestDate: string;
  startDate: string;
  endDate: string;
  status: 'Received' | 'InProgress' | 'Completed' | 'Failed';
  connections?: CrossTenantConnection[];
}

export interface Connection {
  name: string;
  id: string;
  type: string;
  properties: {
    displayName: string;
    iconUri?: string;
    iconBrandColor?: string;
    isCustomApi?: boolean;
    description?: string;
    createdTime: string;
    changedTime: string;
    expirationTime?: string;
    apiId?: string;
    isSsoConnection?: boolean;
    createdBy?: {
      displayName?: string;
      email?: string;
    };
    statuses?: {
      status?: string;
      target?: string;
      error?: {
        code?: string;
        message?: string;
      };
    }[];
    tier?: string;
    publisher?: string;
  };
}

export interface PowerPagesWebsite {
  id: string;
  name: string;
  createdOn: string;
  templateName: string;
  websiteUrl: string;
  tenantId: string;
  dataverseInstanceUrl: string;
  environmentName: string;
  environmentId: string;
  selectedBaseLanguage: string;
  customHostNames?: string[];
  subdomain?: string;
  packageInstallStatus: string;
  type: 'Trial' | 'Production';
  trialExpiringInDays?: number;
  packageVersion?: string;
  siteVisibility: 'public' | 'private';
  status: string;
}

export interface CapacityRegisterEntry {
  capacityType: string;
  units: string;
  totalCapacity: number;
  maxCapacity: number;
  consumed: number;
  ratedConsumption?: number;
  status: string;
  updatedOn?: string;
}

export interface LicenseEntitlement {
  capacityType: string;
  capacitySubType: string;
  entitlementCode: string;
  displayName: string;
  skuId?: string;
  servicePlanId?: string;
  paidEnabled: number;
  trialEnabled: number;
  totalCapacity: number;
  isTemporary: boolean;
  expiryDate?: string;
  nextLifecycleDate?: string;
  capabilityStatus: string;
}

export interface CurrencyRegisterEntry {
  currencyType: string;
  purchased: number;
  allocated: number;
  consumed: number;
  lastUpdatedDay?: string;
}

export interface FlowComplianceSummary {
  flowContext: string;
  usersInCompliance: number;
  usersExceedingCapacity: number;
  usersWithoutLicense: number;
  usersWithoutPremiumLicense: number;
}

export interface UserCapacitySummary {
  userId: string;
  flowContext: string;
  licenseCategorization: string;
  totalConsumption: number;
  totalCapacity: number;
  totalFlows: number;
}

export interface LicensingSnapshot {
  tenantId: string;
  licenseModelType: string;
  capacityStatus: string;
  capacityStatusMessage: string;
  finOpsStatus: string;
  finOpsStatusMessage: string;
  hasTemporaryLicense: boolean;
  temporaryLicenseExpiryDate?: string;
  periodStart: string;
  periodEnd: string;
  capacities: CapacityRegisterEntry[];
  entitlements: LicenseEntitlement[];
  currencies: CurrencyRegisterEntry[];
  compliance: FlowComplianceSummary[];
  users: UserCapacitySummary[];
  warnings: string[];
}
