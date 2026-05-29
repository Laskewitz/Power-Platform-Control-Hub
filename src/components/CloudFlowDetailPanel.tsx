import { useState, useEffect, useCallback } from 'react';
import type { ReactElement, CSSProperties } from 'react';
import {
  makeStyles,
  tokens,
  Text,
  Badge,
  Button,
  Spinner,
  Accordion,
  AccordionItem,
  AccordionHeader,
  AccordionPanel,
  MessageBar,
  MessageBarBody,
  Tooltip,
  Divider,
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  Input,
} from '@fluentui/react-components';
import {
  ArrowLeftRegular,
  PlayRegular,
  StopRegular,
  DeleteRegular,
  PersonRegular,
  PeopleRegular,
  InfoRegular,
  ArrowClockwiseRegular,
  CheckmarkCircleRegular,
  DismissCircleRegular,
  CalendarRegular,
  FlowRegular,
  ShieldCheckmarkRegular,
  WarningRegular,
  ErrorCircleRegular,
  CheckmarkRegular,
  PlugConnectedRegular,
  PersonAddRegular,
  SearchRegular,
  OpenRegular,
} from '@fluentui/react-icons';
import type { Resource } from '../types/inventory.ts';
import type { FlowPermission } from '../generated/models/PowerAutomateManagementModel.ts';
import type { Aaduser } from '../generated/models/AaduserModel.ts';
import { AaduserService } from '../generated/services/AaduserService.ts';
import type { AdminFlowWithDefinition } from '../services/flowManagementService.ts';
import {
  getFlowAsAdmin,
  enableFlow,
  disableFlow,
  deleteFlow,
  listFlowOwners,
  listRunOnlyUsers,
  modifyFlowOwners,
} from '../services/flowManagementService.ts';
import { analyzeFlowDefinition } from '../services/flowAnalyzer.ts';
import type { AnalysisResult, AnalysisSeverity } from '../services/flowAnalyzer.ts';
import { useToastController, Toast, ToastTitle, ToastBody } from '@fluentui/react-components';
import ConfirmDialog from './ConfirmDialog.tsx';
import { extractMessage } from '../utils/errorUtils.ts';
import { getContext } from '@microsoft/power-apps/app';




interface CloudFlowDetailPanelProps {
  resource: Resource;
  onClose: () => void;
  onDeleted: (resourceName: string) => void;
}

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
    overflow: 'hidden',
    backgroundColor: tokens.colorNeutralBackground1,
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
    '@media (max-width: 768px)': {
      padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalM}`,
    },
  },
  headerMeta: {
    display: 'flex',
    flexDirection: 'column',
    gap: '2px',
    flex: 1,
    minWidth: 0,
  },
  titleRow: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    flexWrap: 'wrap',
  },
  flowTitle: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  envText: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
  },
  actionBar: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    flexWrap: 'wrap',
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalXL}`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    flexShrink: 0,
    '@media (max-width: 768px)': {
      padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    },
  },
  body: {
    flex: 1,
    overflowY: 'auto',
    padding: `${tokens.spacingVerticalL} ${tokens.spacingHorizontalXL}`,
    width: '100%',
    boxSizing: 'border-box',
    '@media (max-width: 768px)': {
      padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalM}`,
    },
  },
  detailGrid: {
    display: 'grid',
    gridTemplateColumns: '180px 1fr',
    gap: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalL}`,
    alignItems: 'start',
    '@media (max-width: 600px)': {
      gridTemplateColumns: '1fr',
    },
  },
  detailLabel: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground3,
    textTransform: 'uppercase',
    letterSpacing: '0.04em',
    paddingTop: '2px',
  },
  detailValue: {
    fontSize: tokens.fontSizeBase300,
    color: tokens.colorNeutralForeground1,
  },
  sectionTitle: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground2,
    marginBottom: tokens.spacingVerticalS,
  },
  permissionList: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXS,
  },
  permissionItem: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground2,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  permissionInfo: {
    display: 'flex',
    flexDirection: 'column',
    flex: 1,
    minWidth: 0,
  },
  permissionName: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  permissionEmail: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  emptyState: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    gap: tokens.spacingVerticalM,
    padding: `${tokens.spacingVerticalXXL} 0`,
    color: tokens.colorNeutralForeground3,
  },
  triggerList: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: tokens.spacingVerticalXS,
  },
  triggerItem: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    fontSize: tokens.fontSizeBase300,
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalM}`,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground2,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  connectorCard: {
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    overflow: 'hidden',
    marginBottom: tokens.spacingVerticalXS,
  },
  connectorCardHeader: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    backgroundColor: tokens.colorNeutralBackground2,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  connectorActionRow: {
    display: 'flex',
    alignItems: 'baseline',
    gap: tokens.spacingHorizontalM,
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalL}`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke3}`,
    ':last-child': { borderBottom: 'none' },
  },
  // Analysis tab
  analysisSummary: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground2,
    marginBottom: tokens.spacingVerticalM,
    flexWrap: 'wrap' as const,
  },
  analysisList: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: '0',
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    overflow: 'hidden',
  },
  analysisRow: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    cursor: 'pointer',
    userSelect: 'none' as const,
    backgroundColor: tokens.colorNeutralBackground1,
    ':hover': { backgroundColor: tokens.colorNeutralBackground1Hover },
    ':last-child': { borderBottom: 'none' },
  },
  analysisRowExpanded: {
    backgroundColor: tokens.colorNeutralBackground2,
  },
  analysisRowDetail: {
    padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalL}`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground2,
    display: 'flex',
    flexDirection: 'column' as const,
    gap: tokens.spacingVerticalS,
    ':last-child': { borderBottom: 'none' },
  },
  analysisTitle: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
    flex: 1,
  },
  analysisDesc: {
    fontSize: tokens.fontSizeBase300,
    color: tokens.colorNeutralForeground2,
    lineHeight: '1.5',
  },
  analysisRec: {
    fontSize: tokens.fontSizeBase300,
    color: tokens.colorNeutralForeground2,
    lineHeight: '1.5',
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground3,
    borderLeft: `3px solid ${tokens.colorBrandStroke1}`,
  },
  analysisAffected: {
    display: 'flex',
    flexWrap: 'wrap' as const,
    gap: tokens.spacingHorizontalXS,
  },
  analysisScore: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground2,
    marginBottom: tokens.spacingVerticalM,
  },
  userList: {
    display: 'flex',
    flexDirection: 'column',
    gap: '2px',
    maxHeight: '320px',
    overflowY: 'auto',
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    padding: '4px',
    backgroundColor: tokens.colorNeutralBackground1,
  },
  userListItem: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    borderRadius: tokens.borderRadiusSmall,
    cursor: 'pointer',
    userSelect: 'none' as const,
    ':hover': { backgroundColor: tokens.colorNeutralBackground1Hover },
  },
  userListItemSelected: {
    backgroundColor: tokens.colorBrandBackground2,
    ':hover': { backgroundColor: tokens.colorBrandBackground2Hover },
  },
  userListItemInfo: {
    display: 'flex',
    flexDirection: 'column',
    flex: 1,
    minWidth: 0,
  },
});

function formatDate(iso?: string): string {
  if (!iso) return '—';
  return new Date(iso).toLocaleString();
}

// ─── Connector / trigger display helpers ─────────────────────────────────────

const CONNECTOR_NAMES: Record<string, string> = {
  shared_commondataserviceforapps: 'Dataverse',
  shared_commondataservice: 'Dataverse (Legacy)',
  shared_office365: 'Office 365 Outlook',
  shared_office365users: 'Office 365 Users',
  shared_sharepointonline: 'SharePoint',
  shared_teams: 'Microsoft Teams',
  shared_microsoftteams: 'Microsoft Teams',
  shared_microsoftcopilotstudio: 'Copilot Studio',
  shared_powerplatformadminv2: 'Power Platform Admin V2',
  shared_flowmanagement: 'Power Automate Management',
  shared_conversionservice: 'Content Conversion',
  shared_approvals: 'Approvals',
  shared_azuread: 'Azure AD',
  shared_keyvault: 'Azure Key Vault',
  shared_servicebus: 'Service Bus',
  shared_eventhubs: 'Event Hubs',
  shared_azureblob: 'Azure Blob Storage',
  shared_sql: 'SQL Server',
  shared_documentdb: 'Azure Cosmos DB',
  shared_powerbi: 'Power BI',
  shared_planner: 'Microsoft Planner',
  shared_todo: 'Microsoft To Do',
  shared_forms: 'Microsoft Forms',
  shared_excelonlinebusiness: 'Excel Online (Business)',
  shared_wordonlinebusiness: 'Word Online (Business)',
  shared_onedriveforbusiness: 'OneDrive for Business',
  shared_onedrive: 'OneDrive',
  shared_salesforce: 'Salesforce',
  shared_dynamics365: 'Dynamics 365',
  shared_sendgrid: 'SendGrid',
  shared_twilio: 'Twilio',
  shared_slack: 'Slack',
  shared_smtp: 'SMTP',
};

const TRIGGER_TYPE_LABELS: Record<string, string> = {
  Recurrence: 'Scheduled (Recurrence)',
  Request: 'Instant / When HTTP request received',
  ApiConnection: 'Connector trigger',
  ApiConnectionNotification: 'Connector event trigger',
  ApiConnectionWebhook: 'Connector webhook trigger',
  OpenApiConnection: 'Connector trigger',
  OpenApiConnectionNotification: 'Connector event trigger',
  OpenApiConnectionWebhook: 'Connector webhook trigger',
  Manual: 'Manually triggered',
};

const CONTROL_FLOW_LABELS: Record<string, string> = {
  Foreach: 'Apply to each',
  If: 'Condition',
  Switch: 'Switch',
  Scope: 'Scope (try/catch)',
  Until: 'Do Until',
  Terminate: 'Terminate',
  Compose: 'Compose',
  ParseJson: 'Parse JSON',
  Select: 'Select (array)',
  Filter: 'Filter array',
  Join: 'Join (array)',
  Table: 'Create HTML table',
  Response: 'Response',
  InitializeVariable: 'Initialize variable',
  SetVariable: 'Set variable',
  AppendToStringVariable: 'Append to string variable',
  AppendToArrayVariable: 'Append to array variable',
  IncrementVariable: 'Increment variable',
  DecrementVariable: 'Decrement variable',
  Wait: 'Delay',
  Http: 'HTTP request',
  Workflow: 'Run a child flow',
  JavaScriptCode: 'Run JavaScript code',
};

function getConnectorDisplayName(apiName?: string): string {
  if (!apiName) return 'Unknown connector';
  const lower = apiName.toLowerCase();
  return CONNECTOR_NAMES[lower] ?? apiName.replace(/^shared_/i, '').replace(/_/g, ' ');
}

function humanizeOperationId(id?: string): string {
  if (!id) return '';
  // Split on underscores and camelCase boundaries
  return id
    .replace(/_/g, ' ')
    .replace(/([a-z])([A-Z])/g, '$1 $2')
    .replace(/\b\w/g, (c) => c.toUpperCase());
}

const TRUNCATE_LENGTH = 160;

function CollapsibleError({ error, style }: { error: string; style?: CSSProperties }): ReactElement {
  const [expanded, setExpanded] = useState(false);
  const clean = extractMessage(error);
  const isLong = clean.length > TRUNCATE_LENGTH;
  const displayed = isLong && !expanded ? clean.slice(0, TRUNCATE_LENGTH) + '…' : clean;
  return (
    <MessageBar intent="error" style={style}>
      <MessageBarBody style={{ display: 'flex', alignItems: 'flex-start', gap: tokens.spacingHorizontalS, flexWrap: 'wrap' }}>
        <span style={{ flex: 1, wordBreak: 'break-word' }}>{displayed}</span>
        {isLong && (
          <Button
            appearance="transparent"
            size="small"
            style={{ padding: 0, minWidth: 0, height: 'auto', flexShrink: 0, fontSize: tokens.fontSizeBase200 }}
            onClick={() => setExpanded((e) => !e)}
          >
            {expanded ? '▲ Show less' : '▼ Show more'}
          </Button>
        )}
      </MessageBarBody>
    </MessageBar>
  );
}

function PermissionsList({
  permissions,
  isLoading,
  error,
  emptyLabel,
  resolvedNames = {},
}: {
  permissions: FlowPermission[];
  isLoading: boolean;
  error: string | null;
  emptyLabel: string;
  resolvedNames?: Record<string, string>;
}): ReactElement {
  const styles = useStyles();
  if (isLoading) return <Spinner size="small" label="Loading…" />;
  if (error) {
    const isPermError = error.includes('InsufficientCdsPermissions') || error.includes('permissions');
    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalS }}>
        <MessageBar intent={isPermError ? 'warning' : 'error'}>
          <MessageBarBody>
            {isPermError
              ? 'This list is only available to direct owners of the flow. The admin management API does not expose a permission-neutral list endpoint for owners and run-only users.'
              : error}
          </MessageBarBody>
        </MessageBar>
      </div>
    );
  }
  if (permissions.length === 0) {
    return (
      <div className={styles.emptyState}>
        <PeopleRegular fontSize={32} />
        <Text>{emptyLabel}</Text>
      </div>
    );
  }
  return (
    <div className={styles.permissionList}>
      {permissions.map((p, i) => {
        const principal = p.properties?.principal;
        const resolvedName = principal?.id ? resolvedNames[principal.id] : undefined;
        const name = resolvedName ?? principal?.displayName ?? principal?.userPrincipalName ?? principal?.id ?? 'Unknown';
        const email = principal?.email ?? principal?.userPrincipalName ?? '';
        const role = p.properties?.roleName ?? '—';
        const isResolving = !!principal?.id && !resolvedName && !principal?.displayName;
        return (
          <div key={p.name ?? i} className={styles.permissionItem}>
            <PersonRegular fontSize={20} style={{ color: tokens.colorNeutralForeground3, flexShrink: 0 }} />
            <div className={styles.permissionInfo}>
              <span className={styles.permissionName}>
                {isResolving ? <Spinner size="extra-tiny" style={{ display: 'inline-flex' }} /> : name}
              </span>
              {!isResolving && email && email !== name && (
                <span className={styles.permissionEmail}>{email}</span>
              )}
            </div>
            <Badge appearance="tint" color="informative" size="small">{role}</Badge>
          </div>
        );
      })}
    </div>
  );
}

function severityIcon(severity: AnalysisSeverity): ReactElement {
  if (severity === 'critical') return <ErrorCircleRegular fontSize={16} style={{ color: tokens.colorStatusDangerForeground1, flexShrink: 0 }} />;
  if (severity === 'warning') return <WarningRegular fontSize={16} style={{ color: tokens.colorStatusWarningForeground1, flexShrink: 0 }} />;
  return <InfoRegular fontSize={16} style={{ color: tokens.colorBrandForeground1, flexShrink: 0 }} />;
}

const SEVERITY_LABEL: Record<AnalysisSeverity, string> = {
  critical: 'Critical',
  warning: 'Warning',
  info: 'Tip',
};

const SEVERITY_BADGE_COLOR: Record<AnalysisSeverity, 'danger' | 'warning' | 'brand'> = {
  critical: 'danger',
  warning: 'warning',
  info: 'brand',
};

function AnalysisSection({
  results,
  isLoading,
  hasDefinition,
  flowTypeLabel = 'Flow',
}: {
  results: AnalysisResult[];
  isLoading: boolean;
  hasDefinition: boolean;
  flowTypeLabel?: string;
}): ReactElement {
  const styles = useStyles();
  const [expandedIds, setExpandedIds] = useState<Set<string>>(new Set());

  if (isLoading) return <Spinner size="small" label="Analysing flow…" />;

  if (!hasDefinition) {
    return (
      <div className={styles.emptyState}>
        <InfoRegular fontSize={32} />
        <Text>{flowTypeLabel} definition not available — the API did not return it for this flow type.</Text>
      </div>
    );
  }

  const criticals = results.filter((r) => r.severity === 'critical').length;
  const warnings = results.filter((r) => r.severity === 'warning').length;
  const tips = results.filter((r) => r.severity === 'info').length;

  const toggleRow = (id: string) => {
    setExpandedIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  };

  if (results.length === 0) {
    return (
      <div className={styles.emptyState}>
        <CheckmarkRegular fontSize={40} style={{ color: tokens.colorStatusSuccessForeground1 }} />
        <Text style={{ fontWeight: tokens.fontWeightSemibold }}>All checks passed!</Text>
        <Text style={{ fontSize: tokens.fontSizeBase200, textAlign: 'center', maxWidth: '340px' }}>
          This flow appears to follow Power Automate best practices.
        </Text>
      </div>
    );
  }

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM }}>
      {/* Summary */}
      <div style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS, flexWrap: 'wrap' }}>
        {criticals > 0 && (
          <Badge appearance="filled" color="danger" size="medium">
            {criticals} critical
          </Badge>
        )}
        {warnings > 0 && (
          <Badge appearance="filled" color="warning" size="medium">
            {warnings} warning{warnings !== 1 ? 's' : ''}
          </Badge>
        )}
        {tips > 0 && (
          <Badge appearance="filled" color="brand" size="medium">
            {tips} tip{tips !== 1 ? 's' : ''}
          </Badge>
        )}
        <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3, marginLeft: tokens.spacingHorizontalS }}>
          Click a row to see details and recommendations
        </Text>
      </div>

      {/* Issue table */}
      <div className={styles.analysisList}>
        {results.map((r) => {
          const isExpanded = expandedIds.has(r.id);
          const severityBorderColor =
            r.severity === 'critical' ? tokens.colorStatusDangerForeground1
            : r.severity === 'warning' ? tokens.colorStatusWarningForeground1
            : tokens.colorBrandForeground1;
          return (
            <div key={r.id}>
              {/* Row */}
              <div
                className={styles.analysisRow}
                style={{ borderLeft: `3px solid ${severityBorderColor}`, backgroundColor: isExpanded ? tokens.colorNeutralBackground2 : undefined }}
                onClick={() => toggleRow(r.id)}
                role="button"
                tabIndex={0}
                onKeyDown={(e) => e.key === 'Enter' || e.key === ' ' ? toggleRow(r.id) : undefined}
                aria-expanded={isExpanded}
              >
                {severityIcon(r.severity)}
                <Text className={styles.analysisTitle} style={{ flex: 1 }}>{r.title}</Text>
                <Badge appearance="filled" color={SEVERITY_BADGE_COLOR[r.severity]} size="small">
                  {SEVERITY_LABEL[r.severity]}
                </Badge>
                <ChevronIcon expanded={isExpanded} />
              </div>

              {/* Expanded detail */}
              {isExpanded && (
                <div className={styles.analysisRowDetail} style={{ borderLeft: `3px solid ${severityBorderColor}` }}>
                  <Text className={styles.analysisDesc}>{r.description}</Text>
                  <div className={styles.analysisRec}>
                    <Text style={{ fontSize: tokens.fontSizeBase200, fontWeight: tokens.fontWeightSemibold, color: tokens.colorNeutralForeground2 }}>
                      💡 Recommendation
                    </Text>
                    <Text style={{ fontSize: tokens.fontSizeBase300, display: 'block', marginTop: '4px' }}>{r.recommendation}</Text>
                  </div>
                  {r.affectedItems && r.affectedItems.length > 0 && (
                    <div>
                      <Text style={{ fontSize: tokens.fontSizeBase200, fontWeight: tokens.fontWeightSemibold, color: tokens.colorNeutralForeground2, display: 'block', marginBottom: '6px' }}>
                        Affected actions
                      </Text>
                      <div className={styles.analysisAffected}>
                        {r.affectedItems.map((item) => (
                          <Badge key={item} appearance="tint" color="informative" size="small">{item}</Badge>
                        ))}
                      </div>
                    </div>
                  )}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

function ChevronIcon({ expanded }: { expanded: boolean }): ReactElement {
  return (
    <span style={{
      display: 'inline-flex',
      fontSize: '12px',
      color: tokens.colorNeutralForeground3,
      transition: 'transform 0.15s',
      transform: expanded ? 'rotate(90deg)' : 'rotate(0deg)',
      flexShrink: 0,
    }}>▶</span>
  );
}

function AddOwnerDialog({
  open,
  onClose,
  onAdded,
  envId,
  flowName,
}: {
  open: boolean;
  onClose: () => void;
  onAdded: () => void;
  envId: string;
  flowName: string;
}): ReactElement {
  const styles = useStyles();
  const { dispatchToast } = useToastController('coe-toaster');
  const [users, setUsers] = useState<Aaduser[]>([]);
  const [loadingUsers, setLoadingUsers] = useState(false);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [search, setSearch] = useState('');
  const [selectedUser, setSelectedUser] = useState<Aaduser | null>(null);
  const [saving, setSaving] = useState(false);
  const [saveError, setSaveError] = useState<string | null>(null);

  useEffect(() => {
    if (!open) return;
    setSearch('');
    setSelectedUser(null);
    setSaveError(null);
    setLoadError(null);
    setLoadingUsers(true);
    // Load without server-side filter — the #EXT# contains() filter causes OData encoding issues
    // Filter out external/guest accounts client-side instead
    AaduserService.getAll({ top: 500, orderBy: ['displayname asc'] })
      .then((res) => {
        if (res.success && res.data) {
          setUsers(res.data.filter(u => !u.userprincipalname?.includes('#EXT#')));
        } else {
          setLoadError(res.error?.message ?? 'Failed to load users');
        }
      })
      .catch((e: unknown) => {
        setLoadError(e instanceof Error ? e.message : 'Failed to load users');
      })
      .finally(() => setLoadingUsers(false));
  }, [open]);

  const filteredUsers = users.filter((u) => {
    if (!search.trim()) return true;
    const q = search.toLowerCase();
    return (u.displayname ?? '').toLowerCase().includes(q)
      || (u.mail ?? '').toLowerCase().includes(q)
      || (u.userprincipalname ?? '').toLowerCase().includes(q);
  });

  async function handleAdd() {
    if (!selectedUser) return;
    setSaving(true);
    setSaveError(null);
    try {
      await modifyFlowOwners(envId, flowName, {
        put: [{ properties: { principal: { id: selectedUser.aaduserid, type: 'User' } } }],
      });
      dispatchToast(
        <Toast>
          <ToastTitle>Owner added</ToastTitle>
          <ToastBody>{selectedUser.displayname ?? selectedUser.mail} has been added as an owner.</ToastBody>
        </Toast>,
        { intent: 'success' }
      );
      onAdded();
      onClose();
    } catch (e) {
      setSaveError(e instanceof Error ? e.message : 'Failed to add owner');
    } finally {
      setSaving(false);
    }
  }

  return (
    <Dialog open={open} onOpenChange={(_, d) => { if (!d.open) onClose(); }}>
      <DialogSurface style={{ maxWidth: '520px', width: '100%' }}>
        <DialogBody>
          <DialogTitle>Add Owner</DialogTitle>
          <DialogContent>
            <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM }}>
              <Text style={{ fontSize: tokens.fontSizeBase300, color: tokens.colorNeutralForeground2 }}>
                Select a user from your organization to add as an owner of this flow.
              </Text>
              <Input
                contentBefore={<SearchRegular />}
                placeholder="Search by name or email…"
                value={search}
                onChange={(_, d) => { setSearch(d.value); setSelectedUser(null); }}
                style={{ width: '100%' }}
              />
              {loadingUsers ? (
                <Spinner size="small" label="Loading users…" />
              ) : loadError ? (
                <MessageBar intent="error">
                  <MessageBarBody>Failed to load users: {loadError}</MessageBarBody>
                </MessageBar>
              ) : (
                <div className={styles.userList}>
                  {filteredUsers.length === 0 ? (
                    <div style={{ padding: tokens.spacingVerticalM, textAlign: 'center', color: tokens.colorNeutralForeground3 }}>
                      <Text style={{ fontSize: tokens.fontSizeBase200 }}>
                        {search.trim() ? `No users found matching "${search}"` : 'No users found'}
                      </Text>
                    </div>
                  ) : filteredUsers.map((u) => {
                    const isSelected = selectedUser?.aaduserid === u.aaduserid;
                    return (
                      <div
                        key={u.aaduserid}
                        className={`${styles.userListItem}${isSelected ? ` ${styles.userListItemSelected}` : ''}`}
                        onClick={() => setSelectedUser(u)}
                        role="option"
                        aria-selected={isSelected}
                        tabIndex={0}
                        onKeyDown={(e) => (e.key === 'Enter' || e.key === ' ') && setSelectedUser(u)}
                      >
                        <PersonRegular fontSize={20} style={{ color: isSelected ? tokens.colorBrandForeground1 : tokens.colorNeutralForeground3, flexShrink: 0 }} />
                        <div className={styles.userListItemInfo}>
                          <Text style={{ fontSize: tokens.fontSizeBase300, fontWeight: tokens.fontWeightSemibold, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                            {u.displayname ?? u.userprincipalname ?? u.aaduserid}
                          </Text>
                          {(u.mail ?? u.userprincipalname) && (
                            <Text style={{ fontSize: tokens.fontSizeBase200, color: isSelected ? tokens.colorBrandForeground2 : tokens.colorNeutralForeground3, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                              {u.mail ?? u.userprincipalname}
                            </Text>
                          )}
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
              {saveError && (
                <MessageBar intent="error"><MessageBarBody>{saveError}</MessageBarBody></MessageBar>
              )}
            </div>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={onClose} disabled={saving}>Cancel</Button>
            <Button
              appearance="primary"
              disabled={!selectedUser || saving}
              icon={saving ? <Spinner size="tiny" /> : undefined}
              onClick={() => void handleAdd()}
            >
              {saving ? 'Adding…' : 'Add as Owner'}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

export default function CloudFlowDetailPanel({
  resource,
  onClose,
  onDeleted,
}: CloudFlowDetailPanelProps): ReactElement {
  const styles = useStyles();
  const { dispatchToast } = useToastController('coe-toaster');

  const displayName = resource.properties.displayName ?? resource.name;
  const envId = resource.properties.environmentId ?? '';
  const flowName = resource.name;

  // Derive a human-readable label from the resource type
  const flowTypeLabel = (() => {
    switch (resource.type.toLowerCase()) {
      case 'microsoft.powerautomate/agentflows': return 'Agent Flow';
      case 'microsoft.powerautomate/m365agentflows': return 'M365 Agent Flow';
      default: return 'Cloud Flow';
    }
  })();

  // Build the Power Automate deep-link URL.
  // envId may be "Default-<tenantGuid>" or a bare GUID — the make.powerautomate.com URL needs just the GUID.
  const flowUrl = (() => {
    const envGuid = envId.startsWith('Default-') ? envId.slice('Default-'.length) : envId;
    if (!envGuid || !flowName) return null;
    return `https://make.powerautomate.com/environments/${envGuid}/flows/${flowName}`;
  })();

  const [openSections, setOpenSections] = useState<string[]>(['details', 'analysis']);
  const [flowDetails, setFlowDetails] = useState<AdminFlowWithDefinition | null>(null);
  const [detailsLoading, setDetailsLoading] = useState(true);
  const [detailsError, setDetailsError] = useState<string | null>(null);
  const [analysisResults, setAnalysisResults] = useState<AnalysisResult[]>([]);
  const [creatorDisplayName, setCreatorDisplayName] = useState<string | null>(null);
  const [creatorEmail, setCreatorEmail] = useState<string | null>(null);

  const [owners, setOwners] = useState<FlowPermission[]>([]);
  const [ownersLoading, setOwnersLoading] = useState(false);
  const [ownersError, setOwnersError] = useState<string | null>(null);
  const [ownersFetched, setOwnersFetched] = useState(false);
  const [ownerNames, setOwnerNames] = useState<Record<string, string>>({});
  const [showAddOwner, setShowAddOwner] = useState(false);

  const [runOnlyUsers, setRunOnlyUsers] = useState<FlowPermission[]>([]);
  const [runOnlyLoading, setRunOnlyLoading] = useState(false);
  const [runOnlyError, setRunOnlyError] = useState<string | null>(null);
  const [runOnlyFetched, setRunOnlyFetched] = useState(false);

  const [actionLoading, setActionLoading] = useState<string | null>(null);
  const [confirmDelete, setConfirmDelete] = useState(false);

  // Derived status from flowDetails or fall back to inventory data
  const flowState = flowDetails?.properties?.state ?? null;
  const isStarted = flowState === 'Started';
  const isStopped = flowState === 'Stopped';

  const loadDetails = useCallback(async () => {
    setDetailsLoading(true);
    setDetailsError(null);
    setCreatorDisplayName(null);
    setCreatorEmail(null);
    try {
      const data = await getFlowAsAdmin(envId, flowName);
      setFlowDetails(data);
      if (data.properties?.definition) {
        setAnalysisResults(analyzeFlowDefinition(data.properties.definition));
      } else {
        setAnalysisResults([]);
      }
      // Resolve creator identity
      const userId = data.properties?.creator?.userId ?? data.properties?.creator?.objectId;
      if (userId) {
        AaduserService.get(userId).then((res) => {
          if (res.success && res.data) {
            setCreatorDisplayName(res.data.displayname ?? null);
            setCreatorEmail(res.data.mail ?? res.data.userprincipalname ?? null);
          }
        }).catch(() => { /* graceful fallback */ });
      }
    } catch (e) {
      setDetailsError(e instanceof Error ? e.message : 'Failed to load flow details');
    } finally {
      setDetailsLoading(false);
    }
  }, [envId, flowName]);

  useEffect(() => { void loadDetails(); }, [loadDetails]);

  // Resolve owner GUIDs to display names when owners list changes
  useEffect(() => {
    if (owners.length === 0) return;
    const toResolve = owners
      .map((o) => o.properties?.principal?.id)
      .filter((id): id is string => !!id);
    if (toResolve.length === 0) return;
    let cancelled = false;
    void (async () => {
      const entries = await Promise.all(toResolve.map(async (id) => {
        try {
          const res = await AaduserService.get(id);
          if (res.success && res.data) {
            return [id, res.data.displayname ?? res.data.userprincipalname ?? id] as const;
          }
        } catch { /* fallback to GUID */ }
        return [id, id] as const;
      }));
      if (!cancelled) {
        setOwnerNames((prev) => {
          const next = { ...prev };
          for (const [id, name] of entries) { if (!next[id]) next[id] = name; }
          return next;
        });
      }
    })();
    return () => { cancelled = true; };
  }, [owners]);

  const refreshOwners = useCallback(() => {
    setOwnersLoading(true);
    setOwnersError(null);
    setOwnersFetched(false);
    setOwnerNames({});
    listFlowOwners(envId, flowName)
      .then((d) => { setOwners(d.value ?? []); setOwnersFetched(true); })
      .catch((e) => setOwnersError(e instanceof Error ? e.message : 'Failed to load owners'))
      .finally(() => setOwnersLoading(false));
  }, [envId, flowName]);

  const handleSectionToggle = useCallback((_: unknown, data: { openItems: string[] }) => {
    const newOpen = data.openItems;
    setOpenSections(newOpen);
    if (newOpen.includes('owners') && !ownersFetched) {
      setOwnersLoading(true);
      setOwnersError(null);
      listFlowOwners(envId, flowName)
        .then((d) => { setOwners(d.value ?? []); setOwnersFetched(true); })
        .catch((e) => setOwnersError(e instanceof Error ? e.message : 'Failed to load owners'))
        .finally(() => setOwnersLoading(false));
    }
    if (newOpen.includes('runonlyusers') && !runOnlyFetched) {
      setRunOnlyLoading(true);
      setRunOnlyError(null);
      listRunOnlyUsers(envId, flowName)
        .then((d) => { setRunOnlyUsers(d.value ?? []); setRunOnlyFetched(true); })
        .catch((e) => setRunOnlyError(e instanceof Error ? e.message : 'Failed to load run-only users'))
        .finally(() => setRunOnlyLoading(false));
    }
  }, [ownersFetched, runOnlyFetched, envId, flowName]);

  async function handleEnable() {
    setActionLoading('enable');
    try {
      await enableFlow(envId, flowName);
      setFlowDetails((prev) => prev ? { ...prev, properties: { ...prev.properties, state: 'Started' } } : prev);
      dispatchToast(
        <Toast><ToastTitle>Flow enabled</ToastTitle><ToastBody>"{displayName}" is now running.</ToastBody></Toast>,
        { intent: 'success' }
      );
    } catch (e) {
      dispatchToast(
        <Toast><ToastTitle>Failed to enable flow</ToastTitle><ToastBody>{e instanceof Error ? e.message : 'Unknown error'}</ToastBody></Toast>,
        { intent: 'error' }
      );
    } finally {
      setActionLoading(null);
    }
  }

  async function handleDisable() {
    setActionLoading('disable');
    try {
      await disableFlow(envId, flowName);
      setFlowDetails((prev) => prev ? { ...prev, properties: { ...prev.properties, state: 'Stopped' } } : prev);
      dispatchToast(
        <Toast><ToastTitle>Flow disabled</ToastTitle><ToastBody>"{displayName}" has been stopped.</ToastBody></Toast>,
        { intent: 'success' }
      );
    } catch (e) {
      dispatchToast(
        <Toast><ToastTitle>Failed to disable flow</ToastTitle><ToastBody>{e instanceof Error ? e.message : 'Unknown error'}</ToastBody></Toast>,
        { intent: 'error' }
      );
    } finally {
      setActionLoading(null);
    }
  }

  async function handleDelete() {
    setConfirmDelete(false);
    setActionLoading('delete');
    try {
      await deleteFlow(envId, flowName);
      dispatchToast(
        <Toast><ToastTitle>Flow deleted</ToastTitle><ToastBody>"{displayName}" has been deleted.</ToastBody></Toast>,
        { intent: 'success' }
      );
      onDeleted(flowName);
      onClose();
    } catch (e) {
      dispatchToast(
        <Toast><ToastTitle>Failed to delete flow</ToastTitle><ToastBody>{e instanceof Error ? e.message : 'Unknown error'}</ToastBody></Toast>,
        { intent: 'error' }
      );
    } finally {
      setActionLoading(null);
    }
  }

  async function handleAddSelfAsOwner() {
    const ctx = await getContext();
    const objectId = ctx.user.objectId;
    if (!objectId) {
      dispatchToast(
        <Toast><ToastTitle>Cannot determine your user identity</ToastTitle><ToastBody>Your AAD object ID is not available in this session.</ToastBody></Toast>,
        { intent: 'warning' }
      );
      return;
    }
    setActionLoading('addSelf');
    try {
      await modifyFlowOwners(envId, flowName, {
        put: [{ properties: { principal: { id: objectId, type: 'User' } } }],
      });
      dispatchToast(
        <Toast><ToastTitle>Added as owner</ToastTitle><ToastBody>You have been added as an owner of "{displayName}".</ToastBody></Toast>,
        { intent: 'success' }
      );
      refreshOwners();
    } catch (e) {
      dispatchToast(
        <Toast><ToastTitle>Failed to add yourself as owner</ToastTitle><ToastBody>{e instanceof Error ? e.message : 'Unknown error'}</ToastBody></Toast>,
        { intent: 'error' }
      );
    } finally {
      setActionLoading(null);
    }
  }

  const props = flowDetails?.properties;
  const triggersSummary = props?.definitionSummary?.triggers ?? [];
  const actionsSummary = props?.definitionSummary?.actions ?? [];
  const creator = props?.creator;
  const connectionRefs = props?.connectionReferences ?? [];

  // Build a map from api internal name → display name using connectionReferences
  const connectorNameMap: Record<string, string> = {};
  for (const ref of connectionRefs) {
    const apiName = ref.apiDefinition?.name ?? '';
    const displayName = ref.apiDefinition?.properties?.displayName ?? ref.displayName ?? '';
    if (apiName && displayName) connectorNameMap[apiName.toLowerCase()] = displayName;
  }
  function resolveConnectorName(apiName?: string): string {
    if (!apiName) return 'Unknown connector';
    return connectorNameMap[apiName.toLowerCase()] ?? getConnectorDisplayName(apiName);
  }

  // Build rich connector groups: prefer full definition (has user action names + operationId)
  type RichAction = { name: string; operationId?: string };
  type RichConnectorGroup = { connector: string; actions: RichAction[] };
  const richGroups: Record<string, RichConnectorGroup> = {};
  const richControlFlow: Array<{ type: string; name: string }> = [];

  const defActions = props?.definition?.actions ?? {};
  const hasFullDef = Object.keys(defActions).length > 0;

  if (hasFullDef) {
    for (const [actionName, action] of Object.entries(defActions)) {
      const displayName = humanizeOperationId(actionName);
      const connTypes = ['OpenApiConnection', 'ApiConnection', 'OpenApiConnectionWebhook',
                         'ApiConnectionNotification', 'OpenApiConnectionNotification'];
      if (action.type && connTypes.includes(action.type)) {
        const inputs = action.inputs as Record<string, unknown> | undefined;
        const host = inputs?.host as Record<string, unknown> | undefined;
        const connName = (host?.connectionName as string) ?? '';
        const operationId = humanizeOperationId((host?.operationId as string) ?? '');
        const key = connName.toLowerCase() || 'unknown';
        if (!richGroups[key]) {
          richGroups[key] = { connector: resolveConnectorName(connName), actions: [] };
        }
        richGroups[key].actions.push({ name: displayName, operationId: operationId || undefined });
      } else if (action.type) {
        const typeLabel = CONTROL_FLOW_LABELS[action.type] ?? humanizeOperationId(action.type);
        richControlFlow.push({ type: typeLabel, name: displayName });
      }
    }
  } else {
    // Fallback to summary data
    for (const a of actionsSummary) {
      if (a.api?.name) {
        const key = a.api.name.toLowerCase();
        if (!richGroups[key]) {
          richGroups[key] = { connector: resolveConnectorName(a.api.name), actions: [] };
        }
        if (a.swaggerOperationId) {
          richGroups[key].actions.push({ name: humanizeOperationId(a.swaggerOperationId) });
        }
      } else if (a.type) {
        richControlFlow.push({ type: CONTROL_FLOW_LABELS[a.type] ?? a.type, name: '' });
      }
    }
  }

  const totalActions = Object.values(richGroups).reduce((sum, g) => sum + g.actions.length, 0)
    + richControlFlow.length;

  // Recurrence schedule display from full definition
  const defTriggers = props?.definition?.triggers ?? {};
  const firstTriggerKey = Object.keys(defTriggers)[0];
  const firstTriggerDef = firstTriggerKey ? defTriggers[firstTriggerKey] : undefined;
  const recurrence = firstTriggerDef?.recurrence;
  const recurrenceLabel = recurrence
    ? `Every ${recurrence.interval ?? 1} ${recurrence.frequency ?? ''}${(recurrence.interval ?? 1) !== 1 ? 's' : ''}`
    : null;

  return (
    <>
      <div className={styles.root}>
        {/* Header */}
        <div className={styles.header}>
          <Button
            appearance="subtle"
            icon={<ArrowLeftRegular />}
            onClick={onClose}
            size="small"
          >
            Back to Resources
          </Button>
          <FlowRegular fontSize={20} style={{ color: tokens.colorBrandForeground1, flexShrink: 0 }} />
          <div className={styles.headerMeta}>
            <div className={styles.titleRow}>
              <Tooltip content={displayName} relationship="label">
                <Text className={styles.flowTitle}>{displayName}</Text>
              </Tooltip>
              <Badge appearance="tint" color="brand" size="small">{flowTypeLabel}</Badge>
              {flowState && (
                <Badge
                  appearance="tint"
                  color={isStarted ? 'success' : 'warning'}
                  size="small"
                  icon={isStarted ? <CheckmarkCircleRegular /> : <DismissCircleRegular />}
                >
                  {isStarted ? 'Enabled' : 'Disabled'}
                </Badge>
              )}
            </div>
            {resource.environmentName && (
              <Text className={styles.envText}>🌐 {resource.environmentName}</Text>
            )}
          </div>
        </div>

        {/* Action bar */}
        <div className={styles.actionBar}>
          {isStopped && (
            <Button
              appearance="primary"
              icon={actionLoading === 'enable' ? <Spinner size="tiny" /> : <PlayRegular />}
              disabled={actionLoading !== null}
              onClick={() => void handleEnable()}
              size="small"
            >
              Enable Flow
            </Button>
          )}
          {isStarted && (
            <Button
              appearance="secondary"
              icon={actionLoading === 'disable' ? <Spinner size="tiny" /> : <StopRegular />}
              disabled={actionLoading !== null}
              onClick={() => void handleDisable()}
              size="small"
            >
              Disable Flow
            </Button>
          )}
          <Button
            appearance="subtle"
            icon={<ArrowClockwiseRegular />}
            disabled={actionLoading !== null || detailsLoading}
            onClick={() => void loadDetails()}
            size="small"
            title="Refresh"
          />
          <Button
            appearance="subtle"
            icon={actionLoading === 'addSelf' ? <Spinner size="tiny" /> : <PersonAddRegular />}
            disabled={actionLoading !== null}
            onClick={() => void handleAddSelfAsOwner()}
            size="small"
            title="Add yourself as owner of this flow"
          >
            Add Yourself As Owner
          </Button>
          <Button
            appearance="subtle"
            icon={<PersonAddRegular />}
            disabled={actionLoading !== null}
            onClick={() => setShowAddOwner(true)}
            size="small"
          >
            Add Owner
          </Button>
          <div style={{ marginLeft: 'auto' }}>
            {flowUrl && (
              <Button
                appearance="subtle"
                icon={<OpenRegular />}
                as="a"
                href={flowUrl}
                target="_blank"
                rel="noopener noreferrer"
                size="small"
                title={`Open ${flowTypeLabel} in Power Automate`}
              >
                Open in Power Automate
              </Button>
            )}
            <Button
              appearance="subtle"
              icon={actionLoading === 'delete' ? <Spinner size="tiny" /> : <DeleteRegular />}
              disabled={actionLoading !== null}
              onClick={() => setConfirmDelete(true)}
              size="small"
              style={{ color: tokens.colorStatusDangerForeground1 }}
            >
              Delete
            </Button>
          </div>
        </div>

        {/* Accordion sections */}
        <div className={styles.body}>
          {detailsLoading && <Spinner size="small" label="Loading flow details…" style={{ marginBottom: tokens.spacingVerticalM }} />}
          {detailsError && (
            <CollapsibleError error={detailsError} style={{ marginBottom: tokens.spacingVerticalM }} />
          )}
          <Accordion
            multiple
            collapsible
            openItems={openSections}
            onToggle={handleSectionToggle as (e: unknown, d: { openItems: string[] }) => void}
          >
            {/* ── Flow Details ── */}
            <AccordionItem value="details">
              <AccordionHeader expandIconPosition="end" icon={<InfoRegular />}>
                Flow Details
              </AccordionHeader>
              <AccordionPanel>
                {!detailsLoading && !detailsError && (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalL, paddingBottom: tokens.spacingVerticalL }}>
                    {/* Metadata grid */}
                    <div className={styles.detailGrid}>
                      <span className={styles.detailLabel}>Status</span>
                      <span className={styles.detailValue}>
                        {flowState
                          ? <Badge appearance="tint" color={isStarted ? 'success' : 'warning'} size="small"
                              icon={isStarted ? <CheckmarkCircleRegular /> : <DismissCircleRegular />}>
                              {isStarted ? 'Enabled' : 'Disabled'}
                            </Badge>
                          : '—'}
                      </span>

                      <span className={styles.detailLabel}>Created</span>
                      <span className={styles.detailValue}>
                        <CalendarInlineIcon /> {formatDate(props?.createdTime ?? resource.properties.createdAt)}
                      </span>

                      <span className={styles.detailLabel}>Last Modified</span>
                      <span className={styles.detailValue}>
                        <CalendarInlineIcon /> {formatDate(props?.lastModifiedTime ?? resource.properties.modifiedAt)}
                      </span>

                      {creator && (
                        <>
                          <span className={styles.detailLabel}>Creator</span>
                          <span className={styles.detailValue}>
                            {creatorDisplayName
                              ? <span>
                                  <span style={{ fontWeight: tokens.fontWeightSemibold }}>{creatorDisplayName}</span>
                                  {creatorEmail && <span style={{ color: tokens.colorNeutralForeground3, marginLeft: '8px', fontSize: tokens.fontSizeBase200 }}>{creatorEmail}</span>}
                                </span>
                              : <span style={{ color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 }}>
                                  {creator.userId ?? creator.objectId ?? '—'}
                                </span>
                            }
                          </span>
                        </>
                      )}

                      <span className={styles.detailLabel}>Environment</span>
                      <span className={styles.detailValue}>
                        {resource.environmentName
                          ? <span>
                              <span style={{ fontWeight: tokens.fontWeightSemibold }}>{resource.environmentName}</span>
                              <span style={{ display: 'block', color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200, wordBreak: 'break-all' }}>{envId}</span>
                            </span>
                          : <span style={{ color: tokens.colorNeutralForeground2, fontSize: tokens.fontSizeBase200, wordBreak: 'break-all' }}>{envId}</span>
                        }
                      </span>

                      <span className={styles.detailLabel}>Flow ID</span>
                      <span className={styles.detailValue} style={{ wordBreak: 'break-all', fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>
                        {flowName}
                      </span>
                    </div>

                    {/* Trigger */}
                    {triggersSummary.length > 0 && (
                      <>
                        <Divider />
                        <div>
                          <Text className={styles.sectionTitle}>Trigger</Text>
                          <div className={styles.triggerList}>
                            {triggersSummary.map((t, i) => {
                              const label = TRIGGER_TYPE_LABELS[t.type ?? ''] ?? t.type ?? 'Unknown trigger';
                              const isConnectorTrigger = (t.type ?? '').toLowerCase().includes('apiconnection') || (t.type ?? '').toLowerCase().includes('openapiconnection');
                              const isRecurrence = t.type === 'Recurrence';
                              return (
                                <div key={i} className={styles.triggerItem}>
                                  <FlowRegular fontSize={16} style={{ color: tokens.colorBrandForeground1, flexShrink: 0 }} />
                                  <div style={{ display: 'flex', flexDirection: 'column', gap: '2px' }}>
                                    <Text style={{ fontWeight: tokens.fontWeightSemibold }}>{label}</Text>
                                    {isRecurrence && recurrenceLabel && (
                                      <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>
                                        ⏱ {recurrenceLabel}
                                      </Text>
                                    )}
                                    {isConnectorTrigger && t.kind && (
                                      <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>
                                        {`Trigger kind: ${t.kind}`}
                                      </Text>
                                    )}
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      </>
                    )}

                    {/* Connectors & Actions */}
                    {(Object.keys(richGroups).length > 0 || richControlFlow.length > 0) && (
                      <>
                        <Divider />
                        <div>
                          <Text className={styles.sectionTitle}>
                            Connectors &amp; Actions ({totalActions} action{totalActions !== 1 ? 's' : ''})
                          </Text>
                          <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXS, marginTop: tokens.spacingVerticalXS }}>
                            {/* Connector groups */}
                            {Object.values(richGroups).map((group, i) => (
                              <div key={i} className={styles.connectorCard}>
                                <div className={styles.connectorCardHeader}>
                                  <PlugConnectorIcon />
                                  <Text style={{ fontWeight: tokens.fontWeightSemibold, flex: 1 }}>{group.connector}</Text>
                                  <Badge appearance="filled" color="informative" size="small">
                                    {group.actions.length} action{group.actions.length !== 1 ? 's' : ''}
                                  </Badge>
                                </div>
                                {group.actions.map((action, j) => (
                                  <div key={j} className={styles.connectorActionRow}>
                                    <Text style={{ flex: 1, fontSize: tokens.fontSizeBase300 }}>{action.name}</Text>
                                    {action.operationId && (
                                      <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3, fontFamily: 'monospace', flexShrink: 0 }}>
                                        {action.operationId}
                                      </Text>
                                    )}
                                  </div>
                                ))}
                              </div>
                            ))}
                            {/* Control flow */}
                            {richControlFlow.length > 0 && (
                              <div className={styles.connectorCard}>
                                <div className={styles.connectorCardHeader}>
                                  <FlowRegular fontSize={14} style={{ color: tokens.colorNeutralForeground3 }} />
                                  <Text style={{ fontWeight: tokens.fontWeightSemibold, color: tokens.colorNeutralForeground2, flex: 1 }}>Control flow</Text>
                                  <Badge appearance="tint" color="subtle" size="small">
                                   {richControlFlow.length} action{richControlFlow.length !== 1 ? 's' : ''}
                                  </Badge>
                                </div>
                                {[...new Map(richControlFlow.map((a) => [`${a.type}|${a.name}`, a])).values()].map((a, j) => (
                                  <div key={j} className={styles.connectorActionRow}>
                                    <Text style={{ flex: 1, fontSize: tokens.fontSizeBase300 }}>{a.name || a.type}</Text>
                                    {a.name && a.type && a.name !== a.type && (
                                      <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3, flexShrink: 0 }}>{a.type}</Text>
                                    )}
                                  </div>
                                ))}
                              </div>
                            )}
                          </div>
                        </div>
                      </>
                    )}
                  </div>
                )}
              </AccordionPanel>
            </AccordionItem>

            {/* ── Best Practice Analysis ── */}
            <AccordionItem value="analysis">
              <AccordionHeader expandIconPosition="end" icon={<ShieldCheckmarkRegular />}>
                <span style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS }}>
                  Best Practice Analysis
                  {analysisResults.length > 0 && !detailsLoading && (
                    <Badge
                      appearance="filled"
                      size="small"
                      color={analysisResults.some((r) => r.severity === 'critical') ? 'danger'
                        : analysisResults.some((r) => r.severity === 'warning') ? 'warning'
                        : 'brand'}
                    >
                      {analysisResults.length}
                    </Badge>
                  )}
                </span>
              </AccordionHeader>
              <AccordionPanel>
                <div style={{ paddingBottom: tokens.spacingVerticalL }}>
                  <AnalysisSection
                    results={analysisResults}
                    isLoading={detailsLoading}
                    hasDefinition={Boolean(flowDetails?.properties?.definition)}
                    flowTypeLabel={flowTypeLabel}
                  />
                </div>
              </AccordionPanel>
            </AccordionItem>

            {/* ── Owners ── */}
            <AccordionItem value="owners">
              <AccordionHeader expandIconPosition="end" icon={<PersonRegular />}>Owners</AccordionHeader>
              <AccordionPanel>
                <div style={{ paddingBottom: tokens.spacingVerticalL }}>
                  <PermissionsList
                    permissions={owners}
                    isLoading={ownersLoading}
                    error={ownersError}
                    emptyLabel="No owners found for this flow."
                    resolvedNames={ownerNames}
                  />
                </div>
              </AccordionPanel>
            </AccordionItem>

            {/* ── Run-Only Users ── */}
            <AccordionItem value="runonlyusers">
              <AccordionHeader expandIconPosition="end" icon={<PeopleRegular />}>Run-Only Users</AccordionHeader>
              <AccordionPanel>
                <div style={{ paddingBottom: tokens.spacingVerticalL }}>
                  <PermissionsList
                    permissions={runOnlyUsers}
                    isLoading={runOnlyLoading}
                    error={runOnlyError}
                    emptyLabel="No run-only users found for this flow."
                  />
                </div>
              </AccordionPanel>
            </AccordionItem>
          </Accordion>
        </div>
      </div>

      <ConfirmDialog
        open={confirmDelete}
        title={`Delete ${flowTypeLabel}`}
        message={`Delete "${displayName}"? This action cannot be undone.`}
        confirmLabel="Delete"
        isDangerous
        isLoading={actionLoading === 'delete'}
        onConfirm={() => void handleDelete()}
        onCancel={() => setConfirmDelete(false)}
      />

      <AddOwnerDialog
        open={showAddOwner}
        onClose={() => setShowAddOwner(false)}
        onAdded={refreshOwners}
        envId={envId}
        flowName={flowName}
      />
    </>
  );
}

// Tiny inline icon helper
function CalendarInlineIcon(): ReactElement {
  return <CalendarRegular style={{ display: 'inline', verticalAlign: 'middle', marginRight: '4px', fontSize: '14px', color: tokens.colorNeutralForeground3 }} />;
}
function PlugConnectorIcon(): ReactElement {
  return <PlugConnectedRegular style={{ fontSize: '14px', color: tokens.colorBrandForeground1, flexShrink: 0 }} />;
}
