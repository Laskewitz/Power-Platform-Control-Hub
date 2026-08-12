export interface InventoryActor {
  id?: string;
  displayName?: string;
  email?: string;
  userPrincipalName?: string;
}

export interface InventorySharingSummary {
  userCount?: number;
  groupCount?: number;
  entireTenant?: boolean;
}

export interface ResourceProperties {
  displayName?: string;
  createdAt?: string;
  createdBy?: string | InventoryActor;
  lastModifiedAt?: string;
  lastModifiedBy?: string | InventoryActor;
  /** Deprecated alias retained for compatibility with older callers. */
  modifiedAt?: string;
  ownerId?: string;
  environmentId?: string;
  isQuarantined?: boolean;
  subType?: string;
  appModuleId?: string;
  logicalName?: string;
  workflowEntityId?: string;
  lastPublishedAt?: string;
  createdIn?: 'Copilot Studio' | 'Microsoft 365 Copilot Agent Builder' | string;
  schemaName?: string;
  isManaged?: boolean;
  quarantinedAt?: string;
  botId?: string;
  entraAppId?: string;
  entraAgentId?: string;
  entraAgentBlueprintId?: string;
  isCLIAgent?: boolean | string;
  orchestration?: 'Generative' | 'Classic' | string;
  model?: string;
  authentication?: 'Microsoft Entra' | 'None' | 'Generic OAuth 2.0' | string;
  channels?: string[];
  sharedWithViewers?: InventorySharingSummary;
  sharedWithEditors?: InventorySharingSummary;
  capabilitiesCounts?: Record<string, number>;
  environmentType?: string;
  environmentGroup?: string;
  environmentGroupId?: string;
  /** Pre-resolved AAD display name for the owner GUID (populated by useInventory). */
  resolvedOwnerName?: string;
  [key: string]: unknown;
}

export interface Resource {
  id?: string;
  name: string;
  type: string;
  location?: string;
  properties: ResourceProperties;
  /** Computed join key — lower-cased environment ID. */
  joinKey?: string;
  /** Joined from the environment record. */
  environmentName?: string;
  environmentRegion?: string;
  environmentType?: string;
  isManagedEnvironment?: boolean | string;
}

export type AgentHarness = 'github-copilot' | 'standard' | 'unknown';

export function getAgentHarness(properties: ResourceProperties): AgentHarness {
  const entry = Object.entries(properties).find(([key]) => {
    const normalizedKey = key.toLowerCase();
    return normalizedKey === 'iscliagent' || normalizedKey === 'cliagent';
  });
  const value = entry?.[1];

  if (value === true || value === 1 || (typeof value === 'string' && value.toLowerCase() === 'true')) {
    return 'github-copilot';
  }
  if (value === false || value === 0 || (typeof value === 'string' && value.toLowerCase() === 'false')) {
    return 'standard';
  }
  return 'unknown';
}

export const AGENT_HARNESS_LABELS: Partial<Record<AgentHarness, string>> = {
  'github-copilot': 'GitHub Copilot Agent',
  standard: 'Standard Harness Agent',
};

export interface InventoryQueryResult {
  totalRecords: number;
  count: number;
  /** 0 = complete, 1 = truncated — use skipToken to page. */
  resultTruncated: number;
  skipToken?: string;
  data: Resource[];
}

export interface ResourceCounts {
  canvasApps: number;
  modelDrivenApps: number;
  cloudFlows: number;
  agents: number;
  agentFlows: number;
  appBuilderApps: number;
  m365AgentFlows: number;
  codeApps: number;
  total: number;
}

export const RESOURCE_TYPE_LABELS: Record<string, string> = {
  'microsoft.powerapps/canvasapps': 'Canvas App',
  'microsoft.powerapps/modeldrivenapps': 'Model-Driven App',
  'microsoft.powerautomate/cloudflows': 'Cloud Flow',
  'microsoft.copilotstudio/agents': 'Copilot Studio Agent',
  'microsoft.powerautomate/agentflows': 'Agent Flow',
  'microsoft.powerapps/apps': 'App Builder App',
  'microsoft.powerautomate/m365agentflows': 'M365 Agent Flow',
  'microsoft.powerapps/codeapps': 'Code App',
  'microsoft.powerpages/websites': 'Power Pages Website',
  'microsoft.powerplatform/environments': 'Environment',
};

/** Shorter badge labels so they always fit on one line. */
export const RESOURCE_TYPE_SHORT_LABELS: Record<string, string> = {
  'microsoft.powerapps/canvasapps': 'Canvas App',
  'microsoft.powerapps/modeldrivenapps': 'Model-Driven',
  'microsoft.powerautomate/cloudflows': 'Cloud Flow',
  'microsoft.copilotstudio/agents': 'Agent',
  'microsoft.powerautomate/agentflows': 'Agent Flow',
  'microsoft.powerapps/apps': 'App Builder',
  'microsoft.powerautomate/m365agentflows': 'M365 Flow',
  'microsoft.powerapps/codeapps': 'Code App',
  'microsoft.powerpages/websites': 'Power Pages',
};

type BadgeColor = 'brand' | 'informative' | 'success' | 'warning' | 'severe' | 'important' | 'danger';

export function getTypeBadgeColor(type: string): BadgeColor {
  const t = type.toLowerCase();
  if (t === 'microsoft.copilotstudio/agents') return 'success';
  if (t === 'microsoft.powerpages/websites') return 'important';
  if (t.includes('flow')) return 'informative';
  if (t.includes('apps') || t.includes('codeapps') || t.includes('canvasapps') || t.includes('modeldrivenapps')) return 'brand';
  return 'brand';
}

export const RESOURCE_TYPES_FILTER = [
  { key: 'all', label: 'All Types' },
  { key: 'microsoft.powerapps/canvasapps', label: 'Canvas Apps' },
  { key: 'microsoft.powerapps/modeldrivenapps', label: 'Model-Driven Apps' },
  { key: 'microsoft.powerautomate/cloudflows', label: 'Cloud Flows' },
  { key: 'microsoft.copilotstudio/agents', label: 'Agents' },
  { key: 'microsoft.powerautomate/agentflows', label: 'Agent Flows' },
  { key: 'microsoft.powerapps/apps', label: 'App Builder Apps' },
  { key: 'microsoft.powerautomate/m365agentflows', label: 'M365 Agent Flows' },
  { key: 'microsoft.powerapps/codeapps', label: 'Code Apps' },
  { key: 'microsoft.powerpages/websites', label: 'Power Pages Websites' },
] as const;
