import { useCallback, useEffect, useMemo, useState } from 'react';
import type { ReactElement } from 'react';
import {
  Accordion,
  AccordionHeader,
  AccordionItem,
  AccordionPanel,
  Badge,
  Button,
  Card,
  Checkbox,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Dropdown,
  Field,
  Input,
  MessageBar,
  MessageBarBody,
  Option,
  Spinner,
  Text,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import {
  ArrowLeftRegular,
  AddRegular,
  DeleteRegular,
  EditRegular,
  WarningRegular,
  ProhibitedRegular,
  CheckmarkCircleRegular,
  InfoRegular,
  CheckmarkRegular,
  ShieldLockRegular,
  PlugConnectedRegular,
  GlobeRegular,
} from '@fluentui/react-icons';
import type {
  ManagedPolicyV2,
  ManagedPolicyV2defaultConnectorsClassification,
  ManagedPolicyV2environmentType,
  PolicyV2,
} from '../generated/models/PowerPlatformforAdminsModel.ts';
import type { Connection } from '../types/admin.ts';
import type { Resource } from '../types/inventory.ts';
import { fetchConnectors } from '../services/adminApi.ts';
import { createDlpPolicy, deleteDlpPolicy, updateDlpPolicy } from '../services/dlpService.ts';
import { useMutation } from '../hooks/useMutation.tsx';
import ConfirmDialog from './ConfirmDialog.tsx';

interface DlpPoliciesViewProps {
  dlpPolicies: PolicyV2[];
  isLoading: boolean;
  environments: Resource[];
  resources: Resource[];
  onRefresh: () => Promise<void>;
}

type DlpPage =
  | { type: 'list' }
  | { type: 'create' }
  | { type: 'edit'; policy: PolicyV2 }
  | { type: 'detail'; policy: PolicyV2 }
  | { type: 'impact'; returnToEdit?: PolicyV2 };

type ConnectorClassification = ManagedPolicyV2defaultConnectorsClassification;

type ConnectorItem = {
  id: string;
  name: string;
  type: string;
  classification: ConnectorClassification;
};

const ENV_TYPE_LABELS: Record<ManagedPolicyV2environmentType, string> = {
  AllEnvironments: 'All Environments',
  OnlyEnvironments: 'Only Environments',
  ExceptEnvironments: 'Except Environments',
  SingleEnvironment: 'Single Environment',
};

const CLASSIFICATION_LABELS: Record<ConnectorClassification, string> = {
  Confidential: 'Confidential (Business)',
  General: 'General (Non-Business)',
  Blocked: 'Blocked',
};

const CLASSIFICATION_COLORS: Record<ConnectorClassification, 'success' | 'warning' | 'danger' | 'informative'> = {
  Confidential: 'warning',
  General: 'success',
  Blocked: 'danger',
};

const CONNECTOR_CLASSIFICATIONS: ConnectorClassification[] = ['Confidential', 'General', 'Blocked'];

const HIDDEN_CONNECTORS: Array<Omit<ConnectorItem, 'classification'>> = [
  {
    id: '/providers/Microsoft.PowerApps/apis/shared_microsoftspatialservices',
    name: 'Spatial Services',
    type: 'Microsoft.PowerApps/apis',
  },
  {
    id: 'HttpRequestReceived',
    name: 'When a HTTP request is received',
    type: 'Microsoft.PowerApps/apis',
  },
  {
    id: 'HttpWebhook',
    name: 'HTTP Webhook',
    type: 'Microsoft.PowerApps/apis',
  },
  {
    id: 'Http',
    name: 'HTTP',
    type: 'Microsoft.PowerApps/apis',
  },
];

interface AdvisoryRule { id: string; name: string; reason: string }
interface Advisory { connectorName: string; currentClassification: ConnectorClassification; recommendedClassification: ConnectorClassification; reason: string }

const SHOULD_BE_CONFIDENTIAL: AdvisoryRule[] = [
  { id: 'shared_sharepointonline',        name: 'SharePoint',               reason: 'Contains business documents and sensitive organisational data' },
  { id: 'shared_onedriveforbusiness',     name: 'OneDrive for Business',    reason: 'Stores business files — should stay within corporate boundary' },
  { id: 'shared_dynamicscrmonline',       name: 'Dynamics 365',             reason: 'Contains CRM and business-critical operational data' },
  { id: 'shared_commondataserviceforapps',name: 'Microsoft Dataverse',      reason: 'Primary store for structured business data in Power Platform' },
  { id: 'shared_commondataservice',       name: 'Common Data Service',      reason: 'Stores structured business data' },
  { id: 'shared_sql',                     name: 'SQL Server',               reason: 'Often contains sensitive business and customer records' },
  { id: 'shared_teams',                   name: 'Microsoft Teams',          reason: 'Contains internal business communications and meetings' },
  { id: 'shared_office365',              name: 'Office 365 Outlook',       reason: 'Business email — mixing with non-business connectors risks data leaks' },
  { id: 'shared_office365users',          name: 'Office 365 Users',         reason: 'Exposes organisational user directory information' },
  { id: 'shared_azuread',                 name: 'Azure Active Directory',   reason: 'Identity and access management data should be treated as confidential' },
  { id: 'shared_powerbi',                 name: 'Power BI',                 reason: 'Contains business intelligence reports and underlying datasets' },
  { id: 'shared_keyvault',                name: 'Azure Key Vault',          reason: 'Stores secrets, certificates, and credentials' },
  { id: 'shared_azureblob',               name: 'Azure Blob Storage',       reason: 'May contain sensitive business files and backups' },
  { id: 'shared_visualstudioteamservices',name: 'Azure DevOps',             reason: 'Contains source code, pipelines, and project data' },
  { id: 'shared_microsoftforms',          name: 'Microsoft Forms',          reason: 'May collect sensitive business information via forms' },
  { id: 'shared_microsoftgraph',          name: 'Microsoft Graph',          reason: 'Broad access to Microsoft 365 data across the organisation' },
  { id: 'shared_approvals',              name: 'Approvals',                reason: 'Used for business workflows — should be in the same group as other business connectors' },
];

const SHOULD_BE_BLOCKED: AdvisoryRule[] = [
  { id: 'shared_dropbox',    name: 'Dropbox',        reason: 'Personal cloud storage — high risk of uncontrolled data exfiltration' },
  { id: 'shared_twitter',    name: 'Twitter / X',    reason: 'Public social media — business data could be posted publicly' },
  { id: 'shared_gmail',      name: 'Gmail',          reason: 'Personal email — bypasses corporate email DLP controls' },
  { id: 'shared_googledrive',name: 'Google Drive',   reason: 'Unmanaged cloud storage — business data can leave the corporate tenant' },
  { id: 'shared_instagram',  name: 'Instagram',      reason: 'Consumer social media platform — not appropriate for business data' },
  { id: 'shared_youtube',    name: 'YouTube',        reason: 'Consumer platform — no business justification for data integration' },
];

function connectorMatchesAdvisory(connectorId: string, advisoryId: string): boolean {
  const norm = connectorId.toLowerCase();
  const adv = advisoryId.toLowerCase();
  return norm === adv || norm.endsWith('/' + adv) || norm.includes(adv);
}

function computeAdvisories(policy: PolicyV2): Advisory[] {
  // Build id→classification and name→classification maps
  const byId = new Map<string, ConnectorClassification>();
  const byName = new Map<string, ConnectorClassification>();
  for (const group of policy.connectorGroups ?? []) {
    const cls = group.classification as ConnectorClassification;
    for (const c of group.connectors ?? []) {
      if (c.id)   byId.set(c.id.toLowerCase(), cls);
      if (c.name) byName.set(c.name.toLowerCase(), cls);
    }
  }
  const defaultClass = (policy.defaultConnectorsClassification ?? 'General') as ConnectorClassification;
  const getClass = (rule: AdvisoryRule): ConnectorClassification => {
    for (const [id, cls] of byId.entries()) if (connectorMatchesAdvisory(id, rule.id)) return cls;
    return byName.get(rule.name.toLowerCase()) ?? defaultClass;
  };

  const advisories: Advisory[] = [];
  for (const rule of SHOULD_BE_CONFIDENTIAL) {
    const current = getClass(rule);
    if (current !== 'Confidential')
      advisories.push({ connectorName: rule.name, currentClassification: current, recommendedClassification: 'Confidential', reason: rule.reason });
  }
  for (const rule of SHOULD_BE_BLOCKED) {
    const current = getClass(rule);
    if (current !== 'Blocked')
      advisories.push({ connectorName: rule.name, currentClassification: current, recommendedClassification: 'Blocked', reason: rule.reason });
  }
  return advisories;
}

function computeAdvisoriesFromItems(items: ConnectorItem[]): Advisory[] {
  const byId = new Map<string, ConnectorClassification>();
  for (const item of items) byId.set(item.id.toLowerCase(), item.classification);

  const getClass = (rule: AdvisoryRule): ConnectorClassification | null => {
    for (const [id, cls] of byId.entries())
      if (connectorMatchesAdvisory(id, rule.id)) return cls;
    return null;
  };

  const advisories: Advisory[] = [];
  for (const rule of SHOULD_BE_CONFIDENTIAL) {
    const current = getClass(rule);
    if (current !== null && current !== 'Confidential')
      advisories.push({ connectorName: rule.name, currentClassification: current, recommendedClassification: 'Confidential', reason: rule.reason });
  }
  for (const rule of SHOULD_BE_BLOCKED) {
    const current = getClass(rule);
    if (current !== null && current !== 'Blocked')
      advisories.push({ connectorName: rule.name, currentClassification: current, recommendedClassification: 'Blocked', reason: rule.reason });
  }
  return advisories;
}

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalL,
    flex: 1,
    minHeight: 0,
    overflow: 'hidden',
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalXL}`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    flexShrink: 0,
    backgroundColor: tokens.colorNeutralBackground1,
    flexWrap: 'wrap',
  },
  headerMeta: {
    display: 'flex',
    flexDirection: 'column',
    gap: '2px',
    flex: 1,
    minWidth: 0,
  },
  detailPageTitle: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  titleRow: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: tokens.spacingHorizontalM,
    flexWrap: 'wrap',
  },
  backBtn: {
    alignSelf: 'flex-start',
  },
  detailActionBar: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalXL}`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    flexShrink: 0,
    flexWrap: 'wrap',
  },
  detailLabel: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground3,
    textTransform: 'uppercase' as const,
    letterSpacing: '0.04em',
    paddingTop: '2px',
  },
  detailValue: {
    fontSize: tokens.fontSizeBase300,
    color: tokens.colorNeutralForeground1,
  },
  detailGrid: {
    display: 'grid',
    gridTemplateColumns: '180px 1fr',
    gap: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalL}`,
    alignItems: 'start',
  },
  body: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalL,
    flex: 1,
    minHeight: 0,
    overflowY: 'auto',
    paddingRight: tokens.spacingHorizontalXS,
  },
  form: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalL,
    width: '100%',
  },
  formSection: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
  },
  sectionTitle: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
  },
  infoGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))',
    gap: tokens.spacingHorizontalM,
  },
  infoCell: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXS,
    padding: tokens.spacingHorizontalM,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground2,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  tableWrapper: {
    flex: 1,
    overflowY: 'auto',
    overflowX: 'auto',
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  table: {
    width: '100%',
    borderCollapse: 'collapse',
    minWidth: '760px',
  },
  th: {
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    textAlign: 'left',
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground3,
    textTransform: 'uppercase',
    letterSpacing: '0.04em',
    borderBottom: `2px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground3,
    whiteSpace: 'nowrap',
  },
  td: {
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    fontSize: tokens.fontSizeBase300,
    color: tokens.colorNeutralForeground1,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    verticalAlign: 'middle',
  },
  emptyState: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    textAlign: 'center',
    gap: tokens.spacingVerticalS,
    padding: tokens.spacingHorizontalXXL,
    minHeight: '240px',
    color: tokens.colorNeutralForeground3,
  },
  connectorTag: {
    display: 'inline-flex',
  },
  actionBar: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    justifyContent: 'flex-end',
    flexWrap: 'wrap',
  },
  stageActionBar: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    flexWrap: 'wrap',
  },
  nameButton: {
    justifyContent: 'flex-start',
    paddingLeft: '0',
    minWidth: '0',
  },
  pageTitle: {
    fontSize: tokens.fontSizeHero700,
    fontWeight: tokens.fontWeightSemibold,
  },
  subtitle: {
    color: tokens.colorNeutralForeground3,
  },
  titleMeta: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    flexWrap: 'wrap',
  },
  listHeader: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXS,
  },
  helperText: {
    color: tokens.colorNeutralForeground3,
  },
  errorText: {
    color: tokens.colorPaletteRedForeground1,
  },
  connectorList: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: tokens.spacingHorizontalS,
    paddingRight: tokens.spacingHorizontalXS,
  },
  accordion: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
  },
  bucketHeader: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    flexWrap: 'wrap',
  },
  connectorChipList: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
  },
  connectorChip: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: tokens.spacingHorizontalM,
    flexWrap: 'wrap',
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalM}`,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground2,
  },
  connectorChipName: {
    flex: 1,
    minWidth: '220px',
  },
  connectorChipSelect: {
    minWidth: '220px',
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalS}`,
    borderRadius: tokens.borderRadiusSmall,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorNeutralForeground1,
    fontSize: tokens.fontSizeBase200,
  },
  divider: {
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    paddingBottom: tokens.spacingVerticalM,
  },
  advisoryRow: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    padding: `${tokens.spacingVerticalS} 0`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    ':last-child': { borderBottom: 'none' },
  },
  advisoryBadges: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    flexShrink: 0,
  },
  advisoryReason: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    flex: 1,
  },
  advisoryHeader: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    marginBottom: tokens.spacingVerticalS,
  },
  impactSection: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
  },
  impactMetricRow: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))',
    gap: tokens.spacingHorizontalM,
  },
  impactMetric: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXS,
    padding: tokens.spacingHorizontalM,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground2,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  impactMetricValue: {
    fontSize: tokens.fontSizeHero700,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorBrandForeground1,
  },
  impactMetricLabel: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
  },
  riskRow: {
    display: 'flex',
    alignItems: 'flex-start',
    gap: tokens.spacingHorizontalM,
    padding: `${tokens.spacingVerticalS} 0`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    ':last-child': { borderBottom: 'none' },
  },
  riskConnectorName: {
    fontWeight: tokens.fontWeightSemibold,
    flex: 1,
  },
  riskReason: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
  },
});

function formatDate(value?: string): string {
  if (!value) return '—';
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? value : date.toLocaleString();
}

function getEnvironmentScopeLabel(environmentType?: string): string {
  if (!environmentType) return '—';
  return ENV_TYPE_LABELS[environmentType as ManagedPolicyV2environmentType] ?? environmentType;
}

function getClassificationColor(classification?: string): 'success' | 'warning' | 'danger' | 'informative' {
  if (!classification) return 'informative';
  return CLASSIFICATION_COLORS[classification as ConnectorClassification] ?? 'informative';
}

function getClassificationLabel(classification?: string): string {
  if (!classification) return 'Unknown';
  return CLASSIFICATION_LABELS[classification as ConnectorClassification] ?? classification;
}

function getInventoryEnvironmentLabel(environment: Resource): string {
  return environment.properties.displayName ?? environment.name;
}

function getTotalConnectors(policy: PolicyV2): number {
  return policy.connectorGroups?.reduce((total, group) => total + (group.connectors?.length ?? 0), 0) ?? 0;
}

function getPrincipalName(principal?: { displayName?: string; id?: string }): string {
  return principal?.displayName || principal?.id || '—';
}

function getPolicyEnvironmentBadgeLabel(environment: { id: string; name: string; type: string }): string {
  return environment.id || environment.name || environment.type;
}

function connectorMatchesSearch(connector: ConnectorItem, filter: string): boolean {
  if (!filter) return true;
  const query = filter.toLowerCase();
  return connector.name.toLowerCase().includes(query) || connector.id.toLowerCase().includes(query);
}

function mapConnectionToConnectorItem(
  connector: Connection,
  classification: ConnectorClassification,
): ConnectorItem {
  const id = connector.name || connector.id;
  return {
    id,
    name: connector.properties.displayName || id,
    type: connector.type,
    classification,
  };
}

function getConnectorGroups(connectors: ConnectorItem[]): ManagedPolicyV2['connectorGroups'] {
  return CONNECTOR_CLASSIFICATIONS.map((classification) => ({
    classification,
    connectors: connectors
      .filter((connector) => connector.classification === classification)
      .sort((left, right) => left.name.localeCompare(right.name))
      .map(({ id, name, type }) => ({ id, name, type })),
  }));
}

export default function DlpPoliciesView({
  dlpPolicies,
  isLoading,
  environments,
  resources,
  onRefresh,
}: DlpPoliciesViewProps): ReactElement {
  const styles = useStyles();
  const [page, setPage] = useState<DlpPage>({ type: 'list' });
  const [displayName, setDisplayName] = useState('');
  const [envType, setEnvType] = useState<ManagedPolicyV2environmentType>('AllEnvironments');
  const [defaultClass, setDefaultClass] = useState<ConnectorClassification>('General');
  const [selectedEnvironmentId, setSelectedEnvironmentId] = useState('');
  const [connectorItems, setConnectorItems] = useState<ConnectorItem[]>([]);
  const [connectorSearch, setConnectorSearch] = useState('');
  const [connectorSourceEnvironmentId, setConnectorSourceEnvironmentId] = useState<string | null>(null);
  const [connectorLoadError, setConnectorLoadError] = useState('');
  const [isLoadingConnectors, setIsLoadingConnectors] = useState(false);
  const [displayNameError, setDisplayNameError] = useState('');
  const [deleteTarget, setDeleteTarget] = useState<PolicyV2 | null>(null);
  const [showBestPracticesConfirm, setShowBestPracticesConfirm] = useState(false);

  const sortedPolicies = useMemo(
    () => [...dlpPolicies].sort((left, right) => (left.displayName || left.name).localeCompare(right.displayName || right.name)),
    [dlpPolicies],
  );

  const currentPolicy = page.type === 'detail'
    ? sortedPolicies.find((policy) => policy.name === page.policy.name) ?? page.policy
    : null;

  const environmentOptions = useMemo(
    () => [...environments]
      .map((environment) => ({
        id: environment.name,
        label: getInventoryEnvironmentLabel(environment),
      }))
      .sort((left, right) => left.label.localeCompare(right.label)),
    [environments],
  );

  const selectedEnvironmentLabel = useMemo(
    () => environmentOptions.find((option) => option.id === selectedEnvironmentId)?.label ?? '',
    [environmentOptions, selectedEnvironmentId],
  );

  const connectorsByClassification = useMemo(() => ({
    Confidential: connectorItems.filter((connector) => connector.classification === 'Confidential'),
    General: connectorItems.filter((connector) => connector.classification === 'General'),
    Blocked: connectorItems.filter((connector) => connector.classification === 'Blocked'),
  }), [connectorItems]);

  const filteredConnectorsByClassification = useMemo(() => ({
    Confidential: connectorsByClassification.Confidential.filter((connector) => connectorMatchesSearch(connector, connectorSearch)),
    General: connectorsByClassification.General.filter((connector) => connectorMatchesSearch(connector, connectorSearch)),
    Blocked: connectorsByClassification.Blocked.filter((connector) => connectorMatchesSearch(connector, connectorSearch)),
  }), [connectorSearch, connectorsByClassification]);

  useEffect(() => {
    if (page.type !== 'create') return;
    if (!environmentOptions.length) {
      setSelectedEnvironmentId('');
      return;
    }

    if (!selectedEnvironmentId || !environmentOptions.some((option) => option.id === selectedEnvironmentId)) {
      setSelectedEnvironmentId(environmentOptions[0].id);
    }
  }, [environmentOptions, page.type, selectedEnvironmentId]);

  function resetForm() {
    setDisplayName('');
    setEnvType('AllEnvironments');
    setDefaultClass('General');
    setSelectedEnvironmentId('');
    setConnectorItems([]);
    setConnectorSearch('');
    setConnectorSourceEnvironmentId(null);
    setConnectorLoadError('');
    setIsLoadingConnectors(false);
    setDisplayNameError('');
  }

  function handleBackToList() {
    resetForm();
    setPage({ type: 'list' });
  }

  function handleOpenCreate() {
    resetForm();
    setPage({ type: 'create' });
  }

  const handleLoadConnectors = useCallback(async (environmentId = selectedEnvironmentId) => {
    if (!environmentId) {
      setConnectorLoadError('Select an environment to load connectors from.');
      return;
    }

    setIsLoadingConnectors(true);
    setConnectorLoadError('');

    try {
      const apiConnectors = await fetchConnectors(environmentId);
      const mergedConnectors = [
        ...apiConnectors.map((connector) => mapConnectionToConnectorItem(connector, defaultClass)),
        ...HIDDEN_CONNECTORS.map((connector) => ({ ...connector, classification: defaultClass })),
      ];
      const uniqueConnectors = new Map<string, ConnectorItem>();

      for (const connector of mergedConnectors) {
        if (!connector.id) continue;
        const key = connector.id.toLowerCase();
        if (!uniqueConnectors.has(key)) {
          uniqueConnectors.set(key, {
            ...connector,
            name: connector.name.trim() || connector.id,
          });
        }
      }

      setConnectorItems(Array.from(uniqueConnectors.values()).sort((left, right) => left.name.localeCompare(right.name)));
      setConnectorSearch('');
      setConnectorSourceEnvironmentId(environmentId);
    } catch (error: unknown) {
      setConnectorItems([]);
      setConnectorSourceEnvironmentId(null);
      setConnectorLoadError(error instanceof Error ? error.message : 'Failed to load connectors.');
    } finally {
      setIsLoadingConnectors(false);
    }
  }, [defaultClass, selectedEnvironmentId]);

  useEffect(() => {
    if (page.type !== 'create') return;
    if (environmentOptions.length !== 1) return;
    if (!selectedEnvironmentId || connectorSourceEnvironmentId === selectedEnvironmentId || isLoadingConnectors) return;
    void handleLoadConnectors(selectedEnvironmentId);
  }, [
    connectorSourceEnvironmentId,
    environmentOptions.length,
    handleLoadConnectors,
    isLoadingConnectors,
    page.type,
    selectedEnvironmentId,
  ]);

  function handleConnectorClassificationChange(connectorId: string, classification: ConnectorClassification) {
    setConnectorItems((current) => current.map((connector) => (
      connector.id === connectorId
        ? { ...connector, classification }
        : connector
    )));
  }

  const { execute: execCreateDlp, isLoading: isCreating } = useMutation(createDlpPolicy, {
    successMessage: 'DLP policy created.',
    onSuccess: () => {
      handleBackToList();
      void onRefresh();
    },
  });

  const { execute: execDeleteDlp, isLoading: isDeleting } = useMutation(deleteDlpPolicy, {
    successMessage: 'DLP policy deleted.',
    onSuccess: () => {
      setDeleteTarget(null);
      setPage({ type: 'list' });
      void onRefresh();
    },
  });

  const { execute: execUpdateDlp, isLoading: isUpdating } = useMutation(
    (args: { name: string; data: ManagedPolicyV2 }) => updateDlpPolicy(args.name, args.data),
    {
      successMessage: 'DLP policy updated.',
      onSuccess: () => {
        setPage((p) => {
          if (p.type === 'edit') return { type: 'detail', policy: p.policy };
          if (p.type === 'impact' && p.returnToEdit) return { type: 'detail', policy: p.returnToEdit };
          return p;
        });
        void onRefresh();
      },
    },
  );

  function initEditMode(policy: PolicyV2) {
    setDisplayName(policy.displayName ?? '');
    setEnvType((policy.environmentType as ManagedPolicyV2environmentType | undefined) ?? 'AllEnvironments');
    setDefaultClass((policy.defaultConnectorsClassification as ConnectorClassification | undefined) ?? 'General');
    // Populate connectorItems from the policy's existing connectorGroups
    const items: ConnectorItem[] = [];
    for (const group of policy.connectorGroups ?? []) {
      const cls = group.classification as ConnectorClassification;
      for (const c of group.connectors ?? []) {
        const id = c.id ?? '';
        if (!id) continue;
        items.push({ id, name: c.name || id, type: c.type ?? 'Microsoft.PowerApps/apis', classification: cls });
      }
    }
    items.sort((a, b) => a.name.localeCompare(b.name));
    setConnectorItems(items);
    setConnectorSearch('');
    setConnectorSourceEnvironmentId('policy');
    setConnectorLoadError('');
    setDisplayNameError('');
    setPage({ type: 'edit', policy });
  }

  function handleEditSubmit() {
    const trimmed = displayName.trim();
    if (!trimmed) { setDisplayNameError('Display name is required.'); return; }
    setDisplayNameError('');
    const target = page.type === 'edit'
      ? page.policy
      : page.type === 'impact'
        ? page.returnToEdit ?? null
        : null;
    if (!target) return;
    const payload: ManagedPolicyV2 = {
      displayName: trimmed,
      environmentType: envType,
      defaultConnectorsClassification: defaultClass,
      environments: target.environments ?? [],
      connectorGroups: getConnectorGroups(connectorItems),
    };
    void execUpdateDlp({ name: target.name, data: payload });
  }

  function handleCreateSubmit() {
    const trimmedDisplayName = displayName.trim();
    if (!trimmedDisplayName) {
      setDisplayNameError('Display name is required.');
      return;
    }

    if (!connectorItems.length) {
      setConnectorLoadError('Load connectors before creating the policy.');
      return;
    }

    setDisplayNameError('');
    setConnectorLoadError('');

    const payload: ManagedPolicyV2 = {
      displayName: trimmedDisplayName,
      environmentType: envType,
      defaultConnectorsClassification: defaultClass,
      environments: [],
      connectorGroups: getConnectorGroups(connectorItems),
    };

    void execCreateDlp(payload);
  }

  let renderedPage: ReactElement;

  if (page.type === 'create' || page.type === 'edit') {
    const isEdit = page.type === 'edit';
    const editPolicy = isEdit ? page.policy : null;
    renderedPage = (
      <>
        <div className={styles.header}>
          <Button className={styles.backBtn} appearance="subtle" icon={<ArrowLeftRegular />}
            onClick={isEdit ? () => setPage({ type: 'detail', policy: editPolicy! }) : handleBackToList}>
            {isEdit ? 'Back to Policy' : 'Back to DLP Policies'}
          </Button>
          <div className={styles.listHeader}>
            <Text className={styles.pageTitle}>{isEdit ? `Edit: ${editPolicy?.displayName ?? editPolicy?.name}` : 'New DLP Policy'}</Text>
          </div>
        </div>

        <div className={styles.body}>
          <div className={styles.form}>
            <Card>
              <div className={styles.formSection}>
                <Text className={styles.sectionTitle}>Basic Settings</Text>
                <Field
                  label="Display Name"
                  required
                  validationState={displayNameError ? 'error' : 'none'}
                  validationMessage={displayNameError || undefined}
                >
                  <Input
                    value={displayName}
                    onChange={(_, data) => {
                      setDisplayName(data.value);
                      if (displayNameError && data.value.trim()) setDisplayNameError('');
                    }}
                    placeholder="My DLP Policy"
                  />
                </Field>
                <Field label="Environment Scope" required>
                  <Dropdown
                    value={getEnvironmentScopeLabel(envType)}
                    selectedOptions={[envType]}
                    onOptionSelect={(_, data) => setEnvType((data.optionValue as ManagedPolicyV2environmentType | undefined) ?? 'AllEnvironments')}
                  >
                    {Object.entries(ENV_TYPE_LABELS).map(([value, label]) => (
                      <Option key={value} value={value}>{label}</Option>
                    ))}
                  </Dropdown>
                </Field>
                <Field label="Default Connector Classification" required>
                  <Dropdown
                    value={getClassificationLabel(defaultClass)}
                    selectedOptions={[defaultClass]}
                    onOptionSelect={(_, data) => setDefaultClass((data.optionValue as ConnectorClassification | undefined) ?? 'General')}
                  >
                    {CONNECTOR_CLASSIFICATIONS.map((classification) => (
                      <Option key={classification} value={classification}>{getClassificationLabel(classification)}</Option>
                    ))}
                  </Dropdown>
                </Field>
                <Field label="Load connectors from environment" required>
                  <Dropdown
                    placeholder={environmentOptions.length ? 'Select an environment' : 'No environments available'}
                    value={selectedEnvironmentLabel}
                    selectedOptions={selectedEnvironmentId ? [selectedEnvironmentId] : []}
                    onOptionSelect={(_, data) => {
                      const nextEnvironmentId = data.optionValue ?? '';
                      setSelectedEnvironmentId(nextEnvironmentId);
                      setConnectorItems([]);
                      setConnectorSearch('');
                      setConnectorSourceEnvironmentId(null);
                      setConnectorLoadError('');
                    }}
                  >
                    {environmentOptions.map((environment) => (
                      <Option key={environment.id} value={environment.id}>{environment.label}</Option>
                    ))}
                  </Dropdown>
                </Field>
                <Text className={styles.helperText}>
                  Load every connector from one environment, then move connectors into the Confidential, General, and Blocked buckets.
                </Text>
                {connectorLoadError ? <Text className={styles.errorText}>{connectorLoadError}</Text> : null}
                <div className={styles.stageActionBar}>
                  <Button
                    appearance="secondary"
                    disabled={!selectedEnvironmentId || !environmentOptions.length || isLoadingConnectors}
                    icon={isLoadingConnectors ? <Spinner size="tiny" /> : undefined}
                    onClick={() => void handleLoadConnectors()}
                  >
                    {isLoadingConnectors ? 'Loading…' : connectorSourceEnvironmentId ? 'Reload Connectors' : 'Load Connectors'}
                  </Button>
                  {connectorSourceEnvironmentId ? (
                    <Text className={styles.helperText}>
                      {connectorItems.length} connector{connectorItems.length === 1 ? '' : 's'} loaded from {selectedEnvironmentLabel}.
                    </Text>
                  ) : null}
                </div>
              </div>
            </Card>

            {connectorSourceEnvironmentId ? (
              <Card>
                <div className={styles.formSection}>
                  <div className={styles.listHeader}>
                    <Text className={styles.sectionTitle}>Connector Classification</Text>
                    <Text className={styles.helperText}>
                      All connectors started in the {getClassificationLabel(defaultClass)} bucket when they were loaded.
                    </Text>
                  </div>
                  <Field label="Filter connectors">
                    <Input
                      value={connectorSearch}
                      onChange={(_, data) => setConnectorSearch(data.value)}
                      placeholder="Search by connector name or id"
                    />
                  </Field>
                  <Accordion className={styles.accordion} multiple collapsible defaultOpenItems={CONNECTOR_CLASSIFICATIONS}>
                    {CONNECTOR_CLASSIFICATIONS.map((classification) => {
                      const connectors = connectorsByClassification[classification];
                      const filteredConnectors = filteredConnectorsByClassification[classification];

                      return (
                        <AccordionItem key={classification} value={classification}>
                          <AccordionHeader>
                            <div className={styles.bucketHeader}>
                              <Text weight="semibold">{getClassificationLabel(classification)}</Text>
                              <Badge appearance="tint" color={getClassificationColor(classification)}>{connectors.length}</Badge>
                            </div>
                          </AccordionHeader>
                          <AccordionPanel>
                            {filteredConnectors.length === 0 ? (
                              <Text className={styles.helperText}>
                                {connectorSearch ? 'No connectors match the current filter.' : 'No connectors assigned to this bucket.'}
                              </Text>
                            ) : (
                              <div className={styles.connectorChipList}>
                                {filteredConnectors.map((connector) => (
                                  <div key={connector.id} className={styles.connectorChip}>
                                    <Text className={styles.connectorChipName}>{connector.name}</Text>
                                    <select
                                      className={styles.connectorChipSelect}
                                      value={connector.classification}
                                      onChange={(event) => handleConnectorClassificationChange(
                                        connector.id,
                                        event.target.value as ConnectorClassification,
                                      )}
                                    >
                                      {CONNECTOR_CLASSIFICATIONS.map((option) => (
                                        <option key={option} value={option}>{getClassificationLabel(option)}</option>
                                      ))}
                                    </select>
                                  </div>
                                ))}
                              </div>
                            )}
                          </AccordionPanel>
                        </AccordionItem>
                      );
                    })}
                  </Accordion>
                </div>
              </Card>
            ) : null}

            <div className={styles.actionBar}>
              <Button appearance="primary"
                icon={(isEdit ? isUpdating : isCreating) ? <Spinner size="tiny" /> : (isEdit ? <EditRegular /> : <AddRegular />)}
                disabled={isEdit ? isUpdating : isCreating}
                onClick={isEdit ? handleEditSubmit : handleCreateSubmit}>
                {isEdit ? (isUpdating ? 'Saving…' : 'Save Changes') : (isCreating ? 'Creating…' : 'Create Policy')}
              </Button>
              <Button
                appearance="subtle"
                icon={<CheckmarkRegular />}
                disabled={(isEdit ? isUpdating : isCreating) || connectorItems.length === 0}
                onClick={() => setShowBestPracticesConfirm(true)}
              >
                Apply Best Practices
              </Button>
              <Button appearance="secondary" disabled={isEdit ? isUpdating : isCreating}
                onClick={isEdit ? () => setPage({ type: 'detail', policy: editPolicy! }) : handleBackToList}>
                Cancel
              </Button>
            </div>
          </div>
        </div>

        {/* Best Practices confirmation dialog */}
        {(() => {
          const pendingChanges = computeAdvisoriesFromItems(connectorItems);
          return (
            <Dialog open={showBestPracticesConfirm} onOpenChange={(_, d) => { if (!d.open) setShowBestPracticesConfirm(false); }}>
              <DialogSurface style={{ maxWidth: '560px', width: '100%' }}>
                <DialogBody>
                  <DialogTitle>Apply Best Practices</DialogTitle>
                  <DialogContent>
                    {pendingChanges.length === 0 ? (
                      <Text style={{ color: tokens.colorNeutralForeground2 }}>
                        ✅ All loaded connectors already follow best practices — no changes needed.
                      </Text>
                    ) : (
                      <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM }}>
                        <Text style={{ color: tokens.colorNeutralForeground2 }}>
                          The following {pendingChanges.length} connector{pendingChanges.length !== 1 ? 's' : ''} will be reclassified:
                        </Text>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXS, maxHeight: '320px', overflowY: 'auto' }}>
                          {pendingChanges.map((adv, i) => (
                            <div key={i} style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS, padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalS}`, borderRadius: tokens.borderRadiusMedium, backgroundColor: tokens.colorNeutralBackground2 }}>
                              <Badge appearance="tint" color={adv.currentClassification === 'Confidential' ? 'warning' : adv.currentClassification === 'Blocked' ? 'danger' : 'informative'} size="small">
                                {adv.currentClassification}
                              </Badge>
                              <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>→</Text>
                              <Badge appearance="filled" color={adv.recommendedClassification === 'Confidential' ? 'warning' : adv.recommendedClassification === 'Blocked' ? 'danger' : 'informative'} size="small">
                                {adv.recommendedClassification}
                              </Badge>
                              <Text style={{ flex: 1, fontSize: tokens.fontSizeBase300, fontWeight: tokens.fontWeightSemibold }}>{adv.connectorName}</Text>
                            </div>
                          ))}
                        </div>
                        <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>
                          These changes will be applied to the editor. You can review and adjust before saving.
                        </Text>
                      </div>
                    )}
                  </DialogContent>
                  <DialogActions>
                    <Button appearance="secondary" onClick={() => setShowBestPracticesConfirm(false)}>Cancel</Button>
                    {pendingChanges.length > 0 && (
                      <Button
                        appearance="primary"
                        icon={<CheckmarkRegular />}
                        onClick={() => {
                          setConnectorItems(prev => prev.map(item => {
                            for (const rule of SHOULD_BE_CONFIDENTIAL)
                              if (connectorMatchesAdvisory(item.id, rule.id) && item.classification !== 'Confidential')
                                return { ...item, classification: 'Confidential' as ConnectorClassification };
                            for (const rule of SHOULD_BE_BLOCKED)
                              if (connectorMatchesAdvisory(item.id, rule.id) && item.classification !== 'Blocked')
                                return { ...item, classification: 'Blocked' as ConnectorClassification };
                            return item;
                          }));
                          setShowBestPracticesConfirm(false);
                        }}
                      >
                        Apply {pendingChanges.length} Change{pendingChanges.length !== 1 ? 's' : ''}
                      </Button>
                    )}
                  </DialogActions>
                </DialogBody>
              </DialogSurface>
            </Dialog>
          );
        })()}
      </>
    );
  } else if (page.type === 'impact') {
    const returnToEdit = page.returnToEdit;
    const isFromEdit = Boolean(returnToEdit);
    const blockedItems = connectorItems.filter((connector) => connector.classification === 'Blocked');
    const confidentialItems = connectorItems.filter((connector) => connector.classification === 'Confidential');
    const generalItems = connectorItems.filter((connector) => connector.classification === 'General');
    const hasIsolationRisk = confidentialItems.length > 0 && generalItems.length > 0;
    const liveAdvisories = computeAdvisoriesFromItems(connectorItems);

    const affectedEnvIds = new Set(
      envType === 'AllEnvironments'
        ? environments.map((environment) => environment.name)
        : envType === 'OnlyEnvironments' || envType === 'SingleEnvironment'
          ? [selectedEnvironmentId].filter(Boolean)
          : environments.map((environment) => environment.name)
    );

    const affectedFlows = resources.filter((resource) =>
      resource.type.toLowerCase().includes('cloudflow')
      || resource.type.toLowerCase().includes('agentflow')
      || resource.type.toLowerCase().includes('m365agent')
    ).filter((resource) => {
      const environmentId = resource.properties.environmentId as string | undefined;
      return !environmentId || affectedEnvIds.has(environmentId);
    });

    const affectedApps = resources.filter((resource) =>
      resource.type.toLowerCase().includes('canvasapp')
      || resource.type.toLowerCase().includes('modeldriven')
    ).filter((resource) => {
      const environmentId = resource.properties.environmentId as string | undefined;
      return !environmentId || affectedEnvIds.has(environmentId);
    });

    renderedPage = (
      <>
        <div className={styles.header}>
          <Button className={styles.backBtn} appearance="subtle" icon={<ArrowLeftRegular />}
            onClick={() => setPage(isFromEdit ? { type: 'edit', policy: returnToEdit! } : { type: 'create' })}>
            {isFromEdit ? 'Back to Edit Policy' : 'Back to New Policy'}
          </Button>
          <div className={styles.listHeader}>
            <Text className={styles.pageTitle}>Impact Analysis</Text>
            <Text className={styles.subtitle}>
              Review how "{displayName || 'this policy'}" will affect your tenant before applying it.
            </Text>
          </div>
        </div>

        <div className={styles.body}>
          <div className={styles.form}>
            <Card>
              <div className={styles.impactSection}>
                <div className={styles.advisoryHeader}>
                  <InfoRegular style={{ color: tokens.colorBrandForeground1 }} />
                  <Text className={styles.sectionTitle}>Coverage</Text>
                </div>
                <div className={styles.impactMetricRow}>
                  <div className={styles.impactMetric}>
                    <Text className={styles.impactMetricValue}>{affectedEnvIds.size}</Text>
                    <Text className={styles.impactMetricLabel}>Environments in scope</Text>
                  </div>
                  <div className={styles.impactMetric}>
                    <Text className={styles.impactMetricValue}>{affectedFlows.length}</Text>
                    <Text className={styles.impactMetricLabel}>Flows in scope</Text>
                  </div>
                  <div className={styles.impactMetric}>
                    <Text className={styles.impactMetricValue}>{affectedApps.length}</Text>
                    <Text className={styles.impactMetricLabel}>Apps in scope</Text>
                  </div>
                  <div className={styles.impactMetric}>
                    <Text className={styles.impactMetricValue}>{connectorItems.length}</Text>
                    <Text className={styles.impactMetricLabel}>Connectors classified</Text>
                  </div>
                </div>
                <Text className={styles.helperText}>
                  Scope: <strong>{getEnvironmentScopeLabel(envType)}</strong> · Default classification: <strong>{getClassificationLabel(defaultClass)}</strong>
                </Text>
              </div>
            </Card>

            {(blockedItems.length > 0 || hasIsolationRisk) && (
              <Card>
                <div className={styles.impactSection}>
                  <Text className={styles.sectionTitle}>Risk Summary</Text>
                  {blockedItems.length > 0 && (
                    <MessageBar intent="error">
                      <MessageBarBody>
                        <strong>{blockedItems.length} connector{blockedItems.length !== 1 ? 's' : ''} will be blocked.</strong>{' '}
                        Any flow or app using these connectors will be suspended in environments covered by this policy.
                      </MessageBarBody>
                    </MessageBar>
                  )}
                  {hasIsolationRisk && (
                    <MessageBar intent="warning">
                      <MessageBarBody>
                        <strong>Isolation risk detected.</strong>{' '}
                        This policy has both Confidential ({confidentialItems.length}) and General ({generalItems.length}) connectors.
                        Flows that use connectors from both groups simultaneously will be suspended.
                      </MessageBarBody>
                    </MessageBar>
                  )}
                </div>
              </Card>
            )}

            {blockedItems.length > 0 && (
              <Card>
                <div className={styles.impactSection}>
                  <div className={styles.advisoryHeader}>
                    <ProhibitedRegular style={{ color: tokens.colorPaletteRedForeground1 }} />
                    <Text className={styles.sectionTitle}>Blocked Connectors ({blockedItems.length})</Text>
                  </div>
                  <Text className={styles.helperText}>
                    Flows and apps using any of these connectors will be suspended.
                  </Text>
                  <div>
                    {blockedItems.map((connector) => (
                      <div key={connector.id} className={styles.riskRow}>
                        <div style={{ flex: 1 }}>
                          <Text className={styles.riskConnectorName}>{connector.name}</Text>
                        </div>
                        <Badge appearance="tint" color="danger" size="small">Blocked</Badge>
                      </div>
                    ))}
                  </div>
                </div>
              </Card>
            )}

            {confidentialItems.length > 0 && (
              <Card>
                <div className={styles.impactSection}>
                  <div className={styles.advisoryHeader}>
                    <WarningRegular style={{ color: tokens.colorPaletteGoldForeground2 }} />
                    <Text className={styles.sectionTitle}>Confidential Connectors ({confidentialItems.length})</Text>
                  </div>
                  <Text className={styles.helperText}>
                    These connectors can only be used with other Confidential connectors in the same flow or app.
                    {hasIsolationRisk ? ' Flows mixing these with General connectors will be suspended.' : ''}
                  </Text>
                  <div>
                    {confidentialItems.map((connector) => (
                      <div key={connector.id} className={styles.riskRow}>
                        <Text className={styles.riskConnectorName} style={{ flex: 1 }}>{connector.name}</Text>
                        <Badge appearance="tint" color="warning" size="small">Confidential</Badge>
                      </div>
                    ))}
                  </div>
                </div>
              </Card>
            )}

            {liveAdvisories.length > 0 && (
              <Card>
                <div className={styles.impactSection}>
                  <div className={styles.advisoryHeader}>
                    <InfoRegular style={{ color: tokens.colorBrandForeground1 }} />
                    <Text className={styles.sectionTitle}>Best Practice Advisories ({liveAdvisories.length})</Text>
                  </div>
                  <Text className={styles.helperText}>
                    These connectors are not classified per best practices. Apply them in the policy editor to reduce risk.
                  </Text>
                  <div>
                    {liveAdvisories.map((advisory, index) => (
                      <div key={`${advisory.connectorName}-${index}`} className={styles.riskRow}>
                        <div style={{ flex: 1 }}>
                          <Text className={styles.riskConnectorName}>{advisory.connectorName}</Text>
                          <Text className={styles.riskReason}>{advisory.reason}</Text>
                        </div>
                        <div style={{ display: 'flex', gap: tokens.spacingHorizontalXS, alignItems: 'center', flexShrink: 0 }}>
                          <Badge appearance="tint" color={getClassificationColor(advisory.currentClassification)} size="small">
                            {getClassificationLabel(advisory.currentClassification)}
                          </Badge>
                          <Text>→</Text>
                          <Badge appearance="tint" color={getClassificationColor(advisory.recommendedClassification)} size="small">
                            {getClassificationLabel(advisory.recommendedClassification)}
                          </Badge>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              </Card>
            )}

            {blockedItems.length === 0 && !hasIsolationRisk && liveAdvisories.length === 0 && (
              <Card>
                <div className={styles.impactSection}>
                  <div className={styles.advisoryHeader}>
                    <CheckmarkCircleRegular style={{ color: tokens.colorPaletteGreenForeground1, fontSize: '20px' }} />
                    <Text className={styles.sectionTitle}>No issues detected</Text>
                  </div>
                  <Text>This policy does not block any connectors and has no best-practice violations among the loaded connectors.</Text>
                  <div className={styles.impactSection}>
                    <Checkbox checked disabled label="No blocked connectors" />
                    <Checkbox checked disabled label="No isolation risk detected" />
                    <Checkbox checked disabled label="No best-practice advisory violations" />
                  </div>
                </div>
              </Card>
            )}

            <div className={styles.actionBar}>
              <Button appearance="primary"
                icon={(isFromEdit ? isUpdating : isCreating) ? <Spinner size="tiny" /> : (isFromEdit ? <EditRegular /> : <AddRegular />)}
                disabled={isFromEdit ? isUpdating : isCreating}
                onClick={isFromEdit ? handleEditSubmit : handleCreateSubmit}>
                {isFromEdit ? (isUpdating ? 'Saving…' : 'Save Changes') : (isCreating ? 'Creating…' : 'Create Policy')}
              </Button>
              <Button appearance="secondary"
                onClick={() => setPage(isFromEdit ? { type: 'edit', policy: returnToEdit! } : { type: 'create' })}>
                Edit Policy
              </Button>
            </div>
          </div>
        </div>
      </>
    );
  } else if (page.type === 'detail' && currentPolicy) {
    renderedPage = (
      <>
        {/* Compact header row matching Flow detail style */}
        <div className={styles.header}>
          <Button
            appearance="subtle"
            icon={<ArrowLeftRegular />}
            onClick={() => setPage({ type: 'list' })}
            size="small"
          >
            Back to DLP Policies
          </Button>
          <ShieldLockRegular fontSize={20} style={{ color: tokens.colorBrandForeground1, flexShrink: 0 }} />
          <div className={styles.headerMeta}>
            <div className={styles.titleRow}>
              <Text className={styles.detailPageTitle}>{currentPolicy.displayName || currentPolicy.name}</Text>
              <Badge appearance="tint" color="informative" size="small">{getEnvironmentScopeLabel(currentPolicy.environmentType)}</Badge>
              <Badge appearance="tint" color={getClassificationColor(currentPolicy.defaultConnectorsClassification)} size="small">
                {getClassificationLabel(currentPolicy.defaultConnectorsClassification)}
              </Badge>
            </div>
          </div>
        </div>

        {/* Action bar */}
        <div className={styles.detailActionBar}>
          <Button
            appearance="primary"
            icon={<EditRegular />}
            size="small"
            onClick={() => initEditMode(currentPolicy)}
          >
            Edit Policy
          </Button>
          <div style={{ marginLeft: 'auto' }}>
            <Button
              appearance="subtle"
              icon={<DeleteRegular />}
              size="small"
              disabled={isDeleting}
              style={{ color: tokens.colorStatusDangerForeground1 }}
              onClick={() => setDeleteTarget(currentPolicy)}
            >
              Delete
            </Button>
          </div>
        </div>

        {/* Scrollable accordion body */}
        <div className={styles.body}>
          <Accordion multiple collapsible defaultOpenItems={['details', 'connectors', 'advisories']}>

            {/* Policy Details */}
            <AccordionItem value="details">
              <AccordionHeader expandIconPosition="end" icon={<InfoRegular />}>Policy Details</AccordionHeader>
              <AccordionPanel>
                <div className={styles.detailGrid} style={{ paddingBottom: tokens.spacingVerticalL }}>
                  <span className={styles.detailLabel}>Created by</span>
                  <span className={styles.detailValue}>{getPrincipalName(currentPolicy.createdBy)}</span>
                  <span className={styles.detailLabel}>Created</span>
                  <span className={styles.detailValue}>{formatDate(currentPolicy.createdTime)}</span>
                  <span className={styles.detailLabel}>Modified by</span>
                  <span className={styles.detailValue}>{getPrincipalName(currentPolicy.lastModifiedBy)}</span>
                  <span className={styles.detailLabel}>Last modified</span>
                  <span className={styles.detailValue}>{formatDate(currentPolicy.lastModifiedTime)}</span>
                  <span className={styles.detailLabel}>Environment type</span>
                  <span className={styles.detailValue}>{getEnvironmentScopeLabel(currentPolicy.environmentType)}</span>
                  <span className={styles.detailLabel}>Legacy schema</span>
                  <span className={styles.detailValue}>{currentPolicy.isLegacySchemaVersion ? 'Yes' : 'No'}</span>
                </div>
              </AccordionPanel>
            </AccordionItem>

            {/* Connector Groups */}
            <AccordionItem value="connectors">
              <AccordionHeader expandIconPosition="end" icon={<PlugConnectedRegular />}>Connector Groups</AccordionHeader>
              <AccordionPanel>
                <Accordion multiple defaultOpenItems={CONNECTOR_CLASSIFICATIONS.filter((c) => {
                  const group = currentPolicy.connectorGroups?.find((e) => e.classification === c);
                  return (group?.connectors ?? []).length > 0;
                })}>
                  {CONNECTOR_CLASSIFICATIONS.map((classification) => {
                    const group = currentPolicy.connectorGroups?.find((entry) => entry.classification === classification);
                    const connectors = group?.connectors ?? [];
                    return (
                      <AccordionItem key={classification} value={classification}>
                        <AccordionHeader>
                          <div className={styles.titleRow}>
                            <Text weight="semibold">{getClassificationLabel(classification)}</Text>
                            <Badge appearance="tint" color={getClassificationColor(classification)} size="small">{connectors.length} connector{connectors.length === 1 ? '' : 's'}</Badge>
                          </div>
                        </AccordionHeader>
                        <AccordionPanel>
                          {connectors.length === 0 ? (
                            <Text className={styles.helperText}>No connectors assigned</Text>
                          ) : (
                            <div className={styles.connectorList}>
                              {connectors.map((connector) => (
                                <Badge key={`${classification}-${connector.id}`} className={styles.connectorTag} appearance="outline" color="informative">
                                  {connector.name || connector.id}
                                </Badge>
                              ))}
                            </div>
                          )}
                        </AccordionPanel>
                      </AccordionItem>
                    );
                  })}
                </Accordion>
              </AccordionPanel>
            </AccordionItem>

            {/* Environments (only when not AllEnvironments) */}
            {currentPolicy.environmentType !== 'AllEnvironments' && (
              <AccordionItem value="environments">
                <AccordionHeader expandIconPosition="end" icon={<GlobeRegular />}>Environments</AccordionHeader>
                <AccordionPanel>
                  {currentPolicy.environments?.length ? (
                    <div className={styles.connectorList}>
                      {currentPolicy.environments.map((environment) => (
                        <Badge key={environment.id} appearance="outline" color="brand">{getPolicyEnvironmentBadgeLabel(environment)}</Badge>
                      ))}
                    </div>
                  ) : (
                    <Text className={styles.helperText}>No environments assigned.</Text>
                  )}
                </AccordionPanel>
              </AccordionItem>
            )}

            {/* Advisories */}
            {(() => {
              const advisories = computeAdvisories(currentPolicy);
              const confidentialAdvisories = advisories.filter((a) => a.recommendedClassification === 'Confidential');
              const blockedAdvisories = advisories.filter((a) => a.recommendedClassification === 'Blocked');
              const advisoryIcon = advisories.length === 0
                ? <CheckmarkCircleRegular style={{ color: tokens.colorStatusSuccessForeground1 }} />
                : <WarningRegular style={{ color: tokens.colorStatusWarningForeground1 }} />;
              const advisoryTitle = advisories.length === 0
                ? 'Advisories — Policy looks healthy'
                : `${advisories.length} Advisor${advisories.length === 1 ? 'y' : 'ies'} — Click Edit Policy to apply`;
              return (
                <AccordionItem value="advisories">
                  <AccordionHeader expandIconPosition="end" icon={advisoryIcon}>{advisoryTitle}</AccordionHeader>
                  <AccordionPanel>
                    {advisories.length === 0 ? (
                      <Text className={styles.helperText}>No recommendations — this policy follows best practices.</Text>
                    ) : (
                      <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM }}>
                        {confidentialAdvisories.length > 0 && (
                          <Accordion multiple defaultOpenItems={['confidential']}>
                            <AccordionItem value="confidential">
                              <AccordionHeader>
                                <div className={styles.advisoryHeader}>
                                  <WarningRegular style={{ color: tokens.colorStatusWarningForeground1 }} />
                                  <Text weight="semibold">Move to Confidential (Business) — {confidentialAdvisories.length} connector{confidentialAdvisories.length !== 1 ? 's' : ''}</Text>
                                </div>
                              </AccordionHeader>
                              <AccordionPanel>
                                <Text className={styles.advisoryReason}>These Microsoft connectors handle business data and should not share a group with non-business connectors.</Text>
                                <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXS, marginTop: tokens.spacingVerticalXS }}>
                                  {confidentialAdvisories.map((a) => (
                                    <div key={a.connectorName} className={styles.advisoryRow}>
                                      <div className={styles.advisoryBadges}>
                                        <Text weight="semibold">{a.connectorName}</Text>
                                        <Badge appearance="tint" color={getClassificationColor(a.currentClassification)} size="small">{getClassificationLabel(a.currentClassification)}</Badge>
                                        <Text>→</Text>
                                        <Badge appearance="tint" color={getClassificationColor(a.recommendedClassification)} size="small">{getClassificationLabel(a.recommendedClassification)}</Badge>
                                      </div>
                                      <Text className={styles.advisoryReason}>{a.reason}</Text>
                                    </div>
                                  ))}
                                </div>
                              </AccordionPanel>
                            </AccordionItem>
                          </Accordion>
                        )}
                        {blockedAdvisories.length > 0 && (
                          <Accordion multiple defaultOpenItems={['blocked']}>
                            <AccordionItem value="blocked">
                              <AccordionHeader>
                                <div className={styles.advisoryHeader}>
                                  <ProhibitedRegular style={{ color: tokens.colorStatusDangerForeground1 }} />
                                  <Text weight="semibold">Consider Blocking — {blockedAdvisories.length} connector{blockedAdvisories.length !== 1 ? 's' : ''}</Text>
                                </div>
                              </AccordionHeader>
                              <AccordionPanel>
                                <Text className={styles.advisoryReason}>These consumer/personal connectors are commonly blocked in enterprise tenants to prevent data exfiltration.</Text>
                                <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXS, marginTop: tokens.spacingVerticalXS }}>
                                  {blockedAdvisories.map((a) => (
                                    <div key={a.connectorName} className={styles.advisoryRow}>
                                      <div className={styles.advisoryBadges}>
                                        <Text weight="semibold">{a.connectorName}</Text>
                                        <Badge appearance="tint" color={getClassificationColor(a.currentClassification)} size="small">{getClassificationLabel(a.currentClassification)}</Badge>
                                        <Text>→</Text>
                                        <Badge appearance="tint" color={getClassificationColor(a.recommendedClassification)} size="small">{getClassificationLabel(a.recommendedClassification)}</Badge>
                                      </div>
                                      <Text className={styles.advisoryReason}>{a.reason}</Text>
                                    </div>
                                  ))}
                                </div>
                              </AccordionPanel>
                            </AccordionItem>
                          </Accordion>
                        )}
                      </div>
                    )}
                  </AccordionPanel>
                </AccordionItem>
              );
            })()}
          </Accordion>
        </div>
      </>
    );
  } else {
    renderedPage = (
      <>
        <div className={styles.header}>
          <div className={styles.titleRow}>
            <div className={styles.listHeader}>
              <div className={styles.titleMeta}>
                <Text className={styles.pageTitle}>DLP Policies</Text>
                <Badge appearance="filled" color="informative">{sortedPolicies.length}</Badge>
              </div>
            </div>
            <Button appearance="primary" icon={<AddRegular />} onClick={handleOpenCreate}>
              New DLP Policy
            </Button>
          </div>
        </div>

        <div className={styles.body}>
          {isLoading ? (
            <div className={styles.emptyState}>
              <Spinner size="extra-large" label="Loading DLP policies…" />
            </div>
          ) : sortedPolicies.length === 0 ? (
            <div className={styles.emptyState}>
              <Text>No DLP policies found.</Text>
            </div>
          ) : (
            <div className={styles.tableWrapper}>
              <table className={styles.table}>
                <thead>
                  <tr>
                    <th className={styles.th}>Name ▲</th>
                    <th className={styles.th}>Scope</th>
                    <th className={styles.th}>Default Class</th>
                    <th className={styles.th}>Total Connectors</th>
                    <th className={styles.th}>Last Modified</th>
                    <th className={styles.th} style={{ width: '60px' }}></th>
                  </tr>
                </thead>
                <tbody>
                  {sortedPolicies.map((policy) => (
                    <tr key={policy.name}>
                      <td className={styles.td}>
                        <Button className={styles.nameButton} appearance="subtle" onClick={() => setPage({ type: 'detail', policy })}>
                          {policy.displayName || policy.name}
                        </Button>
                      </td>
                      <td className={styles.td}>{getEnvironmentScopeLabel(policy.environmentType)}</td>
                      <td className={styles.td}>
                        <Badge appearance="tint" color={getClassificationColor(policy.defaultConnectorsClassification)}>
                          {policy.defaultConnectorsClassification}
                        </Badge>
                      </td>
                      <td className={styles.td}>{getTotalConnectors(policy)}</td>
                      <td className={styles.td}>{formatDate(policy.lastModifiedTime)}</td>
                      <td className={styles.td}>
                        <Button
                          appearance="subtle"
                          icon={<DeleteRegular />}
                          title="Delete policy"
                          aria-label={`Delete ${policy.displayName || policy.name}`}
                          disabled={isDeleting}
                          onClick={() => setDeleteTarget(policy)}
                        />
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </>
    );
  }

  return (
    <div className={styles.root}>
      {renderedPage}
      <ConfirmDialog
        open={Boolean(deleteTarget)}
        title="Delete DLP Policy"
        message={`Delete "${deleteTarget?.displayName ?? deleteTarget?.name}"? This action cannot be undone.`}
        confirmLabel="Delete"
        isDangerous
        isLoading={isDeleting}
        onConfirm={() => { if (deleteTarget) void execDeleteDlp(deleteTarget.name); }}
        onCancel={() => setDeleteTarget(null)}
      />
    </div>
  );
}
