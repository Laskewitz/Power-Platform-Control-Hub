import type { Resource } from '../types/inventory.ts';

export type ConnectorUsageCategory = 'apps' | 'flows' | 'agents';

export interface ConnectorUsageEntry {
  id: string;
  name: string;
  resourceCount: number;
}

export interface ConnectorUsageScanResult {
  source: 'inventory-v2';
  scannedAt: string;
  rankings: Record<ConnectorUsageCategory, ConnectorUsageEntry[]>;
  coverage: Record<ConnectorUsageCategory, { scanned: number; failed: number }>;
}

interface InventoryConnector {
  connectorId?: string;
  displayName?: string;
  name?: string;
}

interface ScanTarget {
  category: ConnectorUsageCategory;
  resource: Resource;
}

function getCategory(typeValue: string): ConnectorUsageCategory | null {
  const type = typeValue.toLowerCase();
  if (
    type === 'microsoft.powerapps/canvasapps'
    || type === 'microsoft.powerapps/modeldrivenapps'
    || type === 'microsoft.powerapps/apps'
    || type === 'microsoft.powerapps/codeapps'
  ) return 'apps';
  if (type === 'microsoft.powerautomate/cloudflows') return 'flows';
  if (
    type === 'microsoft.copilotstudio/agents'
    || type === 'microsoft.powerautomate/agentflows'
    || type === 'microsoft.powerautomate/m365agentflows'
  ) return 'agents';
  return null;
}

function normalizeConnectorId(value?: string): string {
  return (value ?? '').split('/').filter(Boolean).at(-1)?.toLowerCase() ?? '';
}

function connectorNameFromId(id: string): string {
  const knownNames: Record<string, string> = {
    shared_commondataserviceforapps: 'Microsoft Dataverse',
    shared_agentnode: 'Agents',
    shared_advancedapprovals: 'Human review',
    shared_computeroperator: 'Computer operator',
    shared_logicflows: 'Power Automate Management',
    shared_microsoftcopilotstudio: 'Microsoft Copilot Studio',
    shared_office365: 'Office 365 Outlook',
    shared_office365users: 'Office 365 Users',
    shared_onedriveforbusiness: 'OneDrive for Business',
    shared_onenote: 'OneNote',
    shared_powerappsforappmakers: 'Power Apps for Makers',
    shared_powerplatformadminv2: 'Power Platform for Admins V2',
    shared_sharepointonline: 'SharePoint',
    shared_teams: 'Microsoft Teams',
  };
  if (knownNames[id]) return knownNames[id];
  return id
    .replace(/^shared_/, '')
    .split(/[_-]/)
    .filter(Boolean)
    .map((part) => part[0]?.toUpperCase() + part.slice(1))
    .join(' ');
}

function getTargets(resources: Resource[]): ScanTarget[] {
  return resources.flatMap((resource) => {
    const category = getCategory(resource.type);
    return category ? [{ category, resource }] : [];
  });
}

function getInventoryConnectors(resource: Resource): InventoryConnector[] {
  const value = resource.properties.powerPlatformConnectors;
  if (!Array.isArray(value)) return [];
  return value.flatMap((connector) => {
    if (typeof connector === 'string') return [{ connectorId: connector }];
    if (connector && typeof connector === 'object') return [connector as InventoryConnector];
    return [];
  });
}

export function getConnectorUsageTargetCount(resources: Resource[]): number {
  return getTargets(resources).length;
}

export async function scanTenantConnectorUsage(
  resources: Resource[],
  onProgress: (completed: number, total: number) => void,
): Promise<ConnectorUsageScanResult> {
  const targets = getTargets(resources);
  const usage = {
    apps: new Map<string, ConnectorUsageEntry>(),
    flows: new Map<string, ConnectorUsageEntry>(),
    agents: new Map<string, ConnectorUsageEntry>(),
  };
  const coverage = {
    apps: { scanned: 0, failed: 0 },
    flows: { scanned: 0, failed: 0 },
    agents: { scanned: 0, failed: 0 },
  };

  onProgress(0, targets.length);
  for (let index = 0; index < targets.length; index++) {
    const { category, resource } = targets[index];
    const connectors = getInventoryConnectors(resource);
    coverage[category].scanned += 1;
    const unique = new Map<string, InventoryConnector>();
    for (const connector of connectors) {
      const id = normalizeConnectorId(connector.connectorId ?? connector.name);
      if (id) unique.set(id, connector);
    }
    for (const [id, connector] of unique) {
      const existing = usage[category].get(id);
      usage[category].set(id, {
        id,
        name: connector.displayName ?? existing?.name ?? connectorNameFromId(id),
        resourceCount: (existing?.resourceCount ?? 0) + 1,
      });
    }

    if ((index + 1) % 100 === 0) await Promise.resolve();
    onProgress(index + 1, targets.length);
  }

  const sortEntries = (entries: Map<string, ConnectorUsageEntry>): ConnectorUsageEntry[] => (
    [...entries.values()].sort((left, right) => (
      right.resourceCount - left.resourceCount || left.name.localeCompare(right.name)
    ))
  );

  return {
    source: 'inventory-v2',
    scannedAt: new Date().toISOString(),
    rankings: {
      apps: sortEntries(usage.apps),
      flows: sortEntries(usage.flows),
      agents: sortEntries(usage.agents),
    },
    coverage,
  };
}
