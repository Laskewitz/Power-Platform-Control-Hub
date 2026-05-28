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
