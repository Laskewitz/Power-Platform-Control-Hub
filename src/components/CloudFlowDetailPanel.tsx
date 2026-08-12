import { useState, useEffect, useCallback, useMemo } from 'react';
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
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  Input,
  Tab,
  TabList,
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
  PlugConnectedRegular,
  PersonAddRegular,
  SearchRegular,
  OpenRegular,
  AppsListRegular,
} from '@fluentui/react-icons';
import type { Resource } from '../types/inventory.ts';
import type { FlowPermission } from '../generated/models/PowerAutomateManagementModel.ts';
import type { Aaduser } from '../generated/models/AaduserModel.ts';
import { AaduserService } from '../generated/services/AaduserService.ts';
import type { AdminFlowWithDefinition } from '../services/flowManagementService.ts';
import {
  getFlowDetailsOnly,
  getFlowDefinitionForAnalysis,
  enableFlow,
  disableFlow,
  deleteFlow,
  listFlowOwners,
  listRunOnlyUsers,
  modifyFlowOwners,
} from '../services/flowManagementService.ts';
import { analyzeFlowDefinition } from '../services/flowAnalyzer.ts';
import type { AnalysisResult } from '../services/flowAnalyzer.ts';
import { useToastController, Toast, ToastTitle, ToastBody } from '@fluentui/react-components';
import ConfirmDialog from './ConfirmDialog.tsx';
import AddSelfAsAdminBanner from './AddSelfAsAdminBanner.tsx';
import { extractMessage } from '../utils/errorUtils.ts';
import { getContext } from '@microsoft/power-apps/app';
import { formatSharedSummary } from '../utils/inventoryFormatters.ts';
import AnalysisPosture from './AnalysisPosture.tsx';
import { useOwner, useOwners } from '../services/ownerCache.ts';




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
    color: '#E5EEF5',
    backgroundColor: '#060A0F',
    fontFamily: '"Aptos", "Segoe UI", sans-serif',
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    padding: `14px ${tokens.spacingHorizontalXL}`,
    borderBottom: '1px solid #29404F',
    flexShrink: 0,
    backgroundColor: '#0C141D',
    boxShadow: 'inset 3px 0 0 #FFB547',
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
    fontSize: '20px',
    fontWeight: tokens.fontWeightSemibold,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
    minWidth: 0,
    flex: '1 1 0',
    color: '#F5FAFD',
  },
  envText: {
    fontSize: '11px',
    color: '#9CB0BF',
    textTransform: 'uppercase',
    letterSpacing: '0.08em',
  },
  actionBar: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalXL}`,
    borderBottom: '1px solid #20313E',
    backgroundColor: '#111827',
    flexShrink: 0,
    flexWrap: 'wrap',
    '& button': {
      color: '#43D9FF',
      borderRadius: 0,
    },
    '@media (max-width: 768px)': {
      padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    },
  },
  body: {
    flex: 1,
    overflowY: 'auto',
    padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalXL}`,
    width: '100%',
    boxSizing: 'border-box',
    '@media (max-width: 768px)': {
      padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalM}`,
    },
  },
  detailGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(2, 1fr)',
    gap: 0,
    alignItems: 'start',
    borderTop: '1px solid #20313E',
    borderLeft: '1px solid #20313E',
    '@media (max-width: 768px)': {
      gridTemplateColumns: '1fr',
    },
  },
  detailItem: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: '4px',
    minWidth: 0,
    minHeight: '58px',
    padding: '10px 12px',
    borderRight: '1px solid #20313E',
    borderBottom: '1px solid #20313E',
    backgroundColor: '#0C141D',
  },
  detailItemWide: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: '4px',
    minWidth: 0,
    gridColumn: 'span 2',
    minHeight: '58px',
    padding: '10px 12px',
    borderRight: '1px solid #20313E',
    borderBottom: '1px solid #20313E',
    backgroundColor: '#0C141D',
    '@media (max-width: 768px)': {
      gridColumn: 'span 1',
    },
  },
  detailLabel: {
    fontSize: '10px',
    fontWeight: tokens.fontWeightSemibold,
    color: '#8DA5B5',
    textTransform: 'uppercase',
    letterSpacing: '0.1em',
  },
  detailValue: {
    fontSize: tokens.fontSizeBase300,
    color: '#E5EEF5',
    fontVariantNumeric: 'tabular-nums',
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
    backgroundColor: '#0C141D',
    border: '1px solid #20313E',
    ':hover': {
      backgroundColor: '#111827',
      boxShadow: 'inset 1px 0 0 #43D9FF',
    },
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
    color: '#9CB0BF',
    border: '1px solid #20313E',
    backgroundColor: '#0C141D',
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
    backgroundColor: '#0C141D',
    border: '1px solid #20313E',
  },
  connectorCard: {
    border: '1px solid #29404F',
    overflow: 'hidden',
    marginBottom: tokens.spacingVerticalXS,
  },
  connectorCardHeader: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    backgroundColor: '#111827',
    borderBottom: '1px solid #20313E',
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
    backgroundColor: '#111827',
    border: '1px solid #29404F',
    marginBottom: tokens.spacingVerticalM,
    flexWrap: 'wrap' as const,
  },
  analysisList: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: '0',
    border: '1px solid #29404F',
    overflow: 'hidden',
  },
  analysisRow: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    borderBottom: '1px solid #20313E',
    cursor: 'pointer',
    userSelect: 'none' as const,
    backgroundColor: '#0C141D',
    ':hover': { backgroundColor: '#111827', boxShadow: 'inset 1px 0 0 #43D9FF' },
    ':last-child': { borderBottom: 'none' },
  },
  analysisRowExpanded: {
    backgroundColor: '#111827',
  },
  analysisRowDetail: {
    padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalL}`,
    borderBottom: '1px solid #20313E',
    backgroundColor: '#111827',
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
    backgroundColor: '#0C141D',
    border: '1px solid #29404F',
    borderLeft: '1px solid #43D9FF',
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
    backgroundColor: '#111827',
    border: '1px solid #29404F',
    marginBottom: tokens.spacingVerticalM,
  },
  userList: {
    display: 'flex',
    flexDirection: 'column',
    gap: '2px',
    maxHeight: '320px',
    overflowY: 'auto',
    border: '1px solid #29404F',
    padding: '4px',
    backgroundColor: '#0C141D',
  },
  userListItem: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    cursor: 'pointer',
    userSelect: 'none' as const,
    ':hover': { backgroundColor: '#111827' },
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
  accordionCard: {
    border: '1px solid #29404F',
    overflow: 'hidden',
    marginBottom: '8px',
    backgroundColor: '#0C141D',
  },
  accordionHeaderTinted: {
    minHeight: '42px',
    color: '#E5EEF5',
    backgroundColor: '#111827',
    borderBottom: '1px solid #20313E',
    ':hover': {
      color: '#43D9FF',
      backgroundColor: '#14212D',
    },
  },
  actionRegister: {
    marginLeft: 'auto',
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    flexWrap: 'wrap',
  },
  contentTabs: {
    padding: `0 ${tokens.spacingHorizontalXL}`,
    borderBottom: '1px solid #29404F',
    backgroundColor: '#060A0F',
    flexShrink: 0,
  },
  sectionBody: {
    padding: `10px ${tokens.spacingHorizontalM} ${tokens.spacingVerticalM}`,
  },
  stack: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
  },
  flowTreeRow: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalM}`,
    border: '1px solid #20313E',
    marginBottom: tokens.spacingVerticalXXS,
    backgroundColor: '#0C141D',
  },
  flowTreeConnector: {
    backgroundColor: '#111827',
  },
  flowTreeContainer: {
    border: '1px solid #29404F',
    marginBottom: tokens.spacingVerticalXS,
    overflow: 'hidden',
    backgroundColor: '#0C141D',
  },
  flowTreeHeader: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    backgroundColor: '#111827',
    borderBottom: '1px solid #20313E',
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

// ─── Flow Tree Types & Builder ────────────────────────────────────────────────

type ActionDef = {
  type?: string;
  inputs?: Record<string, unknown>;
  actions?: Record<string, ActionDef>;
  else?: { actions?: Record<string, ActionDef> };
  cases?: Record<string, { actions?: Record<string, ActionDef> }>;
  default?: { actions?: Record<string, ActionDef> };
};

type RichNode =
  | { kind: 'connector'; name: string; connector: string; operationId: string }
  | { kind: 'control'; name: string; label: string }
  | { kind: 'condition'; name: string; trueNodes: RichNode[]; falseNodes: RichNode[] }
  | { kind: 'foreach'; name: string; children: RichNode[] }
  | { kind: 'scope'; name: string; children: RichNode[] }
  | { kind: 'switch'; name: string; cases: { label: string; nodes: RichNode[] }[]; defaultNodes: RichNode[] }
  | { kind: 'until'; name: string; children: RichNode[] };

const CONN_TYPES = new Set([
  'OpenApiConnection', 'ApiConnection', 'OpenApiConnectionWebhook',
  'ApiConnectionNotification', 'OpenApiConnectionNotification',
]);

function buildFlowTree(
  actions: Record<string, ActionDef> | undefined,
  resolveConnector: (connName: string) => string,
): RichNode[] {
  if (!actions) return [];
  return Object.entries(actions).map(([actionName, action]): RichNode => {
    const displayName = humanizeOperationId(actionName);
    const type = action.type ?? '';

    if (CONN_TYPES.has(type)) {
      const host = action.inputs?.host as Record<string, unknown> | undefined;
      const connName = (host?.connectionName as string) ?? '';
      const opId = humanizeOperationId((host?.operationId as string) ?? '');
      return { kind: 'connector', name: displayName, connector: resolveConnector(connName), operationId: opId };
    }
    if (type === 'If') {
      return {
        kind: 'condition', name: displayName,
        trueNodes: buildFlowTree(action.actions, resolveConnector),
        falseNodes: buildFlowTree(action.else?.actions, resolveConnector),
      };
    }
    if (type === 'Foreach') {
      return { kind: 'foreach', name: displayName, children: buildFlowTree(action.actions, resolveConnector) };
    }
    if (type === 'Scope') {
      return { kind: 'scope', name: displayName, children: buildFlowTree(action.actions, resolveConnector) };
    }
    if (type === 'Until') {
      return { kind: 'until', name: displayName, children: buildFlowTree(action.actions, resolveConnector) };
    }
    if (type === 'Switch') {
      const cases = Object.entries(action.cases ?? {}).map(([label, c]) => ({
        label, nodes: buildFlowTree(c.actions, resolveConnector),
      }));
      return {
        kind: 'switch', name: displayName, cases,
        defaultNodes: buildFlowTree(action.default?.actions, resolveConnector),
      };
    }
    const label = CONTROL_FLOW_LABELS[type] ?? humanizeOperationId(type);
    return { kind: 'control', name: displayName, label };
  });
}

function countNodes(nodes: RichNode[]): number {
  let n = nodes.length;
  for (const node of nodes) {
    if (node.kind === 'condition') n += countNodes(node.trueNodes) + countNodes(node.falseNodes);
    else if (node.kind === 'foreach' || node.kind === 'scope' || node.kind === 'until') n += countNodes(node.children);
    else if (node.kind === 'switch') {
      for (const c of node.cases) n += countNodes(c.nodes);
      n += countNodes(node.defaultNodes);
    }
  }
  return n;
}

function FlowTreeNode({ node }: { node: RichNode }): ReactElement {
  const styles = useStyles();

  if (node.kind === 'connector') {
    return (
      <div className={`${styles.flowTreeRow} ${styles.flowTreeConnector}`}>
        <PlugConnectedRegular fontSize={14} style={{ color: tokens.colorBrandForeground1, flexShrink: 0 }} />
        <Text style={{ fontWeight: tokens.fontWeightSemibold, fontSize: tokens.fontSizeBase300 }}>{node.connector}</Text>
        <Text style={{ color: tokens.colorNeutralForeground3 }}>·</Text>
        <Text style={{ flex: 1, fontSize: tokens.fontSizeBase300 }}>{node.name}</Text>
        {node.operationId && (
          <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3, fontFamily: 'monospace', flexShrink: 0 }}>{node.operationId}</Text>
        )}
      </div>
    );
  }

  if (node.kind === 'control') {
    return (
      <div className={styles.flowTreeRow}>
        <FlowRegular fontSize={14} style={{ color: tokens.colorNeutralForeground3, flexShrink: 0 }} />
        <Text style={{ color: tokens.colorNeutralForeground2, fontSize: tokens.fontSizeBase200, flexShrink: 0 }}>{node.label}</Text>
        {node.name && node.name !== node.label && (
          <Text style={{ flex: 1, fontSize: tokens.fontSizeBase300 }}>{node.name}</Text>
        )}
      </div>
    );
  }

  if (node.kind === 'condition') {
    return (
      <div className={styles.flowTreeContainer}>
        <div className={styles.flowTreeHeader}>
          <FlowRegular fontSize={14} style={{ color: tokens.colorBrandForeground1, flexShrink: 0 }} />
          <Text style={{ fontWeight: tokens.fontWeightSemibold, fontSize: tokens.fontSizeBase300 }}>Condition</Text>
          <Text style={{ color: tokens.colorNeutralForeground3 }}>·</Text>
          <Text style={{ fontSize: tokens.fontSizeBase300 }}>{node.name}</Text>
        </div>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr' }}>
          <div style={{ padding: tokens.spacingHorizontalM, borderRight: `1px solid ${tokens.colorNeutralStroke2}` }}>
            <Text style={{ fontSize: tokens.fontSizeBase200, fontWeight: tokens.fontWeightSemibold, color: tokens.colorStatusSuccessForeground1, display: 'block', marginBottom: tokens.spacingVerticalXS }}>✓ True</Text>
            {node.trueNodes.length > 0
              ? node.trueNodes.map((n, i) => <FlowTreeNode key={i} node={n} />)
              : <Text style={{ color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 }}>No actions</Text>}
          </div>
          <div style={{ padding: tokens.spacingHorizontalM }}>
            <Text style={{ fontSize: tokens.fontSizeBase200, fontWeight: tokens.fontWeightSemibold, color: tokens.colorStatusDangerForeground1, display: 'block', marginBottom: tokens.spacingVerticalXS }}>✗ False</Text>
            {node.falseNodes.length > 0
              ? node.falseNodes.map((n, i) => <FlowTreeNode key={i} node={n} />)
              : <Text style={{ color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 }}>No actions</Text>}
          </div>
        </div>
      </div>
    );
  }

  if (node.kind === 'switch') {
    return (
      <div className={styles.flowTreeContainer}>
        <div className={styles.flowTreeHeader}>
          <FlowRegular fontSize={14} style={{ color: tokens.colorBrandForeground1, flexShrink: 0 }} />
          <Text style={{ fontWeight: tokens.fontWeightSemibold, fontSize: tokens.fontSizeBase300 }}>Switch</Text>
          <Text style={{ color: tokens.colorNeutralForeground3 }}>·</Text>
          <Text style={{ fontSize: tokens.fontSizeBase300 }}>{node.name}</Text>
        </div>
        {node.cases.map((c, i) => (
          <div key={i} style={{ padding: tokens.spacingHorizontalM, borderBottom: `1px solid ${tokens.colorNeutralStroke2}` }}>
            <Text style={{ fontSize: tokens.fontSizeBase200, fontWeight: tokens.fontWeightSemibold, color: tokens.colorNeutralForeground2, display: 'block', marginBottom: tokens.spacingVerticalXS }}>Case: {c.label}</Text>
            {c.nodes.map((n, j) => <FlowTreeNode key={j} node={n} />)}
          </div>
        ))}
        {node.defaultNodes.length > 0 && (
          <div style={{ padding: tokens.spacingHorizontalM }}>
            <Text style={{ fontSize: tokens.fontSizeBase200, fontWeight: tokens.fontWeightSemibold, color: tokens.colorNeutralForeground2, display: 'block', marginBottom: tokens.spacingVerticalXS }}>Default</Text>
            {node.defaultNodes.map((n, i) => <FlowTreeNode key={i} node={n} />)}
          </div>
        )}
      </div>
    );
  }

  // foreach, scope, until
  const containerLabel = node.kind === 'foreach' ? 'For Each' : node.kind === 'until' ? 'Do Until' : 'Scope';
  const children = node.children;
  return (
    <div className={styles.flowTreeContainer}>
      <div className={styles.flowTreeHeader}>
        <FlowRegular fontSize={14} style={{ color: tokens.colorBrandForeground1, flexShrink: 0 }} />
        <Text style={{ fontWeight: tokens.fontWeightSemibold, fontSize: tokens.fontSizeBase300 }}>{containerLabel}</Text>
        <Text style={{ color: tokens.colorNeutralForeground3 }}>·</Text>
        <Text style={{ fontSize: tokens.fontSizeBase300 }}>{node.name}</Text>
      </div>
      <div style={{ padding: tokens.spacingHorizontalM, display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXXS }}>
        {children.length > 0
          ? children.map((n, i) => <FlowTreeNode key={i} node={n} />)
          : <Text style={{ color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 }}>No actions</Text>}
      </div>
    </div>
  );
}

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
        const cachedName = principal?.id ? resolvedNames[principal.id.toLowerCase()] : undefined;
        const resolvedName = cachedName && cachedName.toLowerCase() !== principal?.id?.toLowerCase()
          ? cachedName
          : undefined;
        const name = principal?.displayName ?? resolvedName ?? principal?.userPrincipalName ?? principal?.id ?? 'Unknown';
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
  return (
    <AnalysisPosture
      results={results}
      title={`${flowTypeLabel} posture`}
      description="Governance signals derived from the flow definition, actions, connections, and error handling."
      isLoading={isLoading}
      loadingLabel={`Evaluating ${flowTypeLabel.toLowerCase()} posture…`}
      unavailableMessage={!hasDefinition
        ? `${flowTypeLabel} definition is unavailable because the API did not return it for this flow type.`
        : undefined}
      emptyDescription={`This ${flowTypeLabel.toLowerCase()} follows the evaluated governance practices.`}
    />
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

  // Reset state when dialog opens
  useEffect(() => {
    if (!open) return;
    setSearch('');
    setSelectedUser(null);
    setSaveError(null);
    setLoadError(null);
    setUsers([]);
  }, [open]);

  // Debounced server-side search — fires when user types ≥2 chars
  useEffect(() => {
    const q = search.trim();
    if (q.length < 2) {
      setUsers([]);
      return;
    }
    setLoadingUsers(true);
    setLoadError(null);
    const escaped = q.replace(/'/g, "''");
    const filter = `startswith(displayname,'${escaped}') or startswith(userprincipalname,'${escaped}') or startswith(mail,'${escaped}')`;
    const timer = setTimeout(() => {
      AaduserService.getAll({ filter, top: 50, orderBy: ['displayname asc'] })
        .then((res) => {
          if (res.success && res.data) {
            // Filter out external/guest accounts client-side
            setUsers(res.data.filter(u => !u.userprincipalname?.includes('#EXT#')));
          } else {
            setLoadError(res.error?.message ?? 'Failed to search users');
          }
        })
        .catch((e: unknown) => {
          setLoadError(e instanceof Error ? e.message : 'Failed to search users');
        })
        .finally(() => setLoadingUsers(false));
    }, 350);
    return () => clearTimeout(timer);
  }, [search]);

  const filteredUsers = users;

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
                <Spinner size="small" label="Searching users…" />
              ) : loadError ? (
                <MessageBar intent="error">
                  <MessageBarBody>Search failed: {loadError}</MessageBarBody>
                </MessageBar>
              ) : (
                <div className={styles.userList}>
                  {search.trim().length < 2 ? (
                    <div style={{ padding: tokens.spacingVerticalM, textAlign: 'center', color: tokens.colorNeutralForeground3 }}>
                      <Text style={{ fontSize: tokens.fontSizeBase200 }}>Type at least 2 characters to search</Text>
                    </div>
                  ) : filteredUsers.length === 0 ? (
                    <div style={{ padding: tokens.spacingVerticalM, textAlign: 'center', color: tokens.colorNeutralForeground3 }}>
                      <Text style={{ fontSize: tokens.fontSizeBase200 }}>No users found matching &ldquo;{search}&rdquo;</Text>
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
  // The make.powerautomate.com URL uses the full environment ID as returned by the API
  // (e.g. "Default-<tenantGuid>" for default envs, bare GUID for others) plus /details.
  const flowUrl = (() => {
    if (!envId || !flowName) return null;
    return `https://make.powerautomate.com/environments/${envId}/flows/${flowName}/details`;
  })();

  const [openSections, setOpenSections] = useState<string[]>(['details', 'triggers']);
  const [contentTab, setContentTab] = useState<'overview' | 'analysis'>('overview');
  const [flowDetails, setFlowDetails] = useState<AdminFlowWithDefinition | null>(null);
  const [detailsLoading, setDetailsLoading] = useState(true);
  const [detailsError, setDetailsError] = useState<string | null>(null);
  const [analysisResults, setAnalysisResults] = useState<AnalysisResult[]>([]);
  const [analysisLoading, setAnalysisLoading] = useState(false);

  const [owners, setOwners] = useState<FlowPermission[]>([]);
  const [ownersLoading, setOwnersLoading] = useState(false);
  const [ownersError, setOwnersError] = useState<string | null>(null);
  const [ownersFetched, setOwnersFetched] = useState(false);
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
  const isStopped = flowState !== null && !isStarted;

  const loadDetails = useCallback(async () => {
    setDetailsLoading(true);
    setAnalysisLoading(true); // keep true until Phase 2 finishes to avoid premature "not available" flash
    setDetailsError(null);
    setAnalysisResults([]);
    let basicData: AdminFlowWithDefinition | null = null;

    // Phase 1: load basic flow info without definition (fast)
    try {
      basicData = await getFlowDetailsOnly(envId, flowName);
      setFlowDetails(basicData);
    } catch (e) {
      setDetailsError(e instanceof Error ? e.message : 'Failed to load flow details');
      return;
    } finally {
      setDetailsLoading(false); // show basic content immediately
    }

    // Phase 2: load definition for analysis in background
    try {
      const definition = await getFlowDefinitionForAnalysis(envId, flowName);
      if (definition) {
        setFlowDetails(prev => prev
          ? { ...prev, properties: { ...prev.properties, definition } }
          : prev);
        setAnalysisResults(analyzeFlowDefinition(definition));
      }
    } finally {
      setAnalysisLoading(false);
    }
  }, [envId, flowName]);

  useEffect(() => { void loadDetails(); }, [loadDetails]);

  const refreshOwners = useCallback(() => {
    setOwnersLoading(true);
    setOwnersError(null);
    setOwnersFetched(false);
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
  const creator = props?.creator;
  const creatorDisplayName = useOwner(creator?.userId ?? creator?.objectId, envId);
  const permissionPrincipalIds = useMemo(
    () => [...owners, ...runOnlyUsers]
      .map((permission) => permission.properties?.principal?.id)
      .filter((id): id is string => Boolean(id)),
    [owners, runOnlyUsers],
  );
  const permissionPrincipalMap = useOwners(permissionPrincipalIds, envId);
  const permissionPrincipalNames = useMemo(
    () => Object.fromEntries(permissionPrincipalMap),
    [permissionPrincipalMap],
  );
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

  // Build recursive flow tree from definition
  const flowTree: RichNode[] = useMemo(() => {
    const defActions = props?.definition?.actions as Record<string, ActionDef> | undefined;
    if (!defActions || Object.keys(defActions).length === 0) return [];
    return buildFlowTree(defActions, resolveConnectorName);
  }, [props, resolveConnectorName]);

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
              <div style={{ overflow: 'hidden', minWidth: 0, flex: '1 1 auto' }}>
                <Tooltip content={displayName} relationship="label">
                  <Text className={styles.flowTitle}>{displayName}</Text>
                </Tooltip>
              </div>
              <Badge appearance="tint" color="brand" size="small" style={{ flexShrink: 0 }}>{flowTypeLabel}</Badge>
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
              <Text className={styles.envText}>{resource.environmentName}</Text>
            )}
          </div>
        </div>

        {/* Action bar */}
        <div className={styles.actionBar}>
          <Button
            appearance="subtle"
            icon={detailsLoading || analysisLoading ? <Spinner size="tiny" /> : <ArrowClockwiseRegular />}
            disabled={actionLoading !== null || detailsLoading || analysisLoading}
            onClick={() => void loadDetails()}
            size="small"
          >
            {detailsLoading || analysisLoading ? 'Analyzing…' : 'Re-analyze'}
          </Button>
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
          <AddSelfAsAdminBanner environmentId={envId} variant="inline" onChanged={() => void loadDetails()} />
          <div className={styles.actionRegister}>
            {isStopped && (
              <Button
                appearance="subtle"
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
                appearance="subtle"
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

        <div className={styles.contentTabs}>
          <TabList
            selectedValue={contentTab}
            onTabSelect={(_, data) => setContentTab(data.value as 'overview' | 'analysis')}
          >
            <Tab value="overview">Overview</Tab>
            <Tab value="analysis" icon={<ShieldCheckmarkRegular />}>Analysis</Tab>
          </TabList>
        </div>

        {/* Accordion sections */}
        <div className={styles.body}>
          {detailsError && (() => {
            const isDataverseErr = detailsError.includes('InsufficientCdsPermissions');
            if (isDataverseErr) {
              return (
                <MessageBar intent="warning" style={{ marginBottom: tokens.spacingVerticalM }}>
                  <MessageBarBody>
                    This M365 Agent Flow requires Dataverse access to load its details. Use <strong>Apply admin access</strong> in the toolbar above to gain access, then refresh.
                  </MessageBarBody>
                </MessageBar>
              );
            }
            return <CollapsibleError error={detailsError} style={{ marginBottom: tokens.spacingVerticalM }} />;
          })()}
          {contentTab === 'analysis' ? (
            <div className={styles.sectionBody}>
              <AnalysisSection
                results={analysisResults}
                isLoading={analysisLoading}
                hasDefinition={Boolean(flowDetails?.properties?.definition)}
                flowTypeLabel={flowTypeLabel}
              />
            </div>
          ) : (
          <Accordion
            multiple
            collapsible
            openItems={openSections}
            onToggle={handleSectionToggle as (e: unknown, d: { openItems: string[] }) => void}
          >
            {/* ── Flow Details ── */}
            <AccordionItem value="details" className={styles.accordionCard}>
              <AccordionHeader expandIconPosition="end" icon={<InfoRegular />} className={styles.accordionHeaderTinted}>
                Flow Details
              </AccordionHeader>
              <AccordionPanel>
                <div className={styles.sectionBody}>
                  {detailsLoading && <Spinner size="small" label="Loading flow details…" />}
                  {!detailsLoading && !detailsError && (
                  <div className={styles.stack}>
                    {/* Metadata grid */}
                    <div className={styles.detailGrid}>
                      <div className={styles.detailItem}>
                        <span className={styles.detailLabel}>Status</span>
                        <span className={styles.detailValue}>
                          {flowState
                            ? <Badge appearance="tint" color={isStarted ? 'success' : 'warning'} size="small"
                                icon={isStarted ? <CheckmarkCircleRegular /> : <DismissCircleRegular />}>
                                {isStarted ? 'Enabled' : 'Disabled'}
                              </Badge>
                            : '—'}
                        </span>
                      </div>

                      <div className={styles.detailItem}>
                        <span className={styles.detailLabel}>Environment</span>
                        <span className={styles.detailValue}>
                          {resource.environmentName ?? <span style={{ color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200, wordBreak: 'break-all' }}>{envId}</span>}
                        </span>
                      </div>

                      <div className={styles.detailItem}>
                        <span className={styles.detailLabel}>Created</span>
                        <span className={styles.detailValue}>
                          <CalendarInlineIcon /> {formatDate(props?.createdTime ?? resource.properties.createdAt)}
                        </span>
                      </div>

                      <div className={styles.detailItem}>
                        <span className={styles.detailLabel}>Last Modified</span>
                        <span className={styles.detailValue}>
                          <CalendarInlineIcon /> {formatDate(props?.lastModifiedTime ?? resource.properties.lastModifiedAt ?? resource.properties.modifiedAt)}
                        </span>
                      </div>

                      {creator && (
                        <div className={styles.detailItem}>
                          <span className={styles.detailLabel}>Creator</span>
                          <span className={styles.detailValue}>
                            {creatorDisplayName}
                          </span>
                        </div>
                      )}

                      <div className={styles.detailItem}>
                        <span className={styles.detailLabel}>Flow ID</span>
                        <span className={styles.detailValue} style={{ wordBreak: 'break-all', fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3, fontFamily: 'Consolas, "Courier New", monospace' }}>
                          {flowName}
                        </span>
                      </div>

                      {resource.properties.workflowEntityId && (
                        <div className={styles.detailItem}>
                          <span className={styles.detailLabel}>Workflow Entity ID</span>
                          <span className={styles.detailValue} style={{ wordBreak: 'break-all', fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3, fontFamily: 'Consolas, "Courier New", monospace' }}>
                            {resource.properties.workflowEntityId}
                          </span>
                        </div>
                      )}
                    </div>
                  </div>
                  )}
                </div>
              </AccordionPanel>
            </AccordionItem>

            {/* ── Triggers & Actions ── */}
            <AccordionItem value="triggers" className={styles.accordionCard}>
              <AccordionHeader expandIconPosition="end" icon={<FlowRegular />} className={styles.accordionHeaderTinted}>
                <span style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS }}>
                  Triggers &amp; Actions
                  {!detailsLoading && (triggersSummary.length + countNodes(flowTree)) > 0 && (
                    <Badge appearance="filled" size="small" color="informative">{triggersSummary.length + countNodes(flowTree)}</Badge>
                  )}
                </span>
              </AccordionHeader>
              <AccordionPanel>
                <div className={styles.sectionBody}>
                    {/* Trigger */}
                    {(Object.keys(defTriggers).length > 0 || triggersSummary.length > 0) && (
                      <div>
                        <Text className={styles.sectionTitle}>Trigger</Text>
                        <div className={styles.triggerList}>
                          {Object.keys(defTriggers).length > 0
                            ? Object.entries(defTriggers).map(([triggerKey, t], i) => {
                                const type = t.type ?? '';
                                const isConnectorTrigger = CONN_TYPES.has(type);
                                const isRecurrence = type === 'Recurrence';
                                const host = (t.inputs as Record<string, unknown> | undefined)?.host as Record<string, unknown> | undefined;
                                const connName = (host?.connectionName as string) ?? '';
                                const operationId = humanizeOperationId((host?.operationId as string) ?? '');
                                const connectorLabel = connName ? resolveConnectorName(connName) : null;
                                const triggerName = humanizeOperationId(triggerKey);
                                return (
                                  <div key={i} className={styles.triggerItem}>
                                    <PlugConnectedRegular fontSize={16} style={{ color: tokens.colorBrandForeground1, flexShrink: 0 }} />
                                    <div style={{ display: 'flex', flexDirection: 'column', gap: '2px' }}>
                                      {isConnectorTrigger && connectorLabel ? (
                                        <>
                                          <Text style={{ fontWeight: tokens.fontWeightSemibold }}>{connectorLabel}</Text>
                                          {operationId && (
                                            <Text style={{ fontSize: tokens.fontSizeBase300 }}>{operationId}</Text>
                                          )}
                                          <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>{triggerName}</Text>
                                        </>
                                      ) : (
                                        <>
                                          <Text style={{ fontWeight: tokens.fontWeightSemibold }}>
                                            {TRIGGER_TYPE_LABELS[type] ?? triggerName}
                                          </Text>
                                          {isRecurrence && recurrenceLabel && (
                                            <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>⏱ {recurrenceLabel}</Text>
                                          )}
                                          {!isRecurrence && t.kind && (
                                            <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>Kind: {t.kind}</Text>
                                          )}
                                        </>
                                      )}
                                    </div>
                                  </div>
                                );
                              })
                            : triggersSummary.map((t, i) => {
                                const label = TRIGGER_TYPE_LABELS[t.type ?? ''] ?? t.type ?? 'Unknown trigger';
                                return (
                                  <div key={i} className={styles.triggerItem}>
                                    <FlowRegular fontSize={16} style={{ color: tokens.colorBrandForeground1, flexShrink: 0 }} />
                                    <div style={{ display: 'flex', flexDirection: 'column', gap: '2px' }}>
                                      <Text style={{ fontWeight: tokens.fontWeightSemibold }}>{label}</Text>
                                      {t.kind && (
                                        <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>Kind: {t.kind}</Text>
                                      )}
                                    </div>
                                  </div>
                                );
                              })
                          }
                        </div>
                      </div>
                    )}

                    {/* Flow tree (recursive) */}
                    {flowTree.length > 0 && (
                      <div>
                        <Text className={styles.sectionTitle}>
                          Actions ({countNodes(flowTree)})
                        </Text>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXXS, marginTop: tokens.spacingVerticalXS }}>
                          {flowTree.map((node, i) => <FlowTreeNode key={i} node={node} />)}
                        </div>
                      </div>
                    )}

                    {!detailsLoading && triggersSummary.length === 0 && flowTree.length === 0 && (
                      <Text style={{ color: tokens.colorNeutralForeground3 }}>No trigger or action data available.</Text>
                    )}
                </div>
              </AccordionPanel>
            </AccordionItem>

            {/* ── Inventory ── */}
            {(() => {
              const props = resource.properties;
              const viewersSummary = props?.sharedWithViewers ? formatSharedSummary(props.sharedWithViewers) : null;
              const editorsSummary = props?.sharedWithEditors ? formatSharedSummary(props.sharedWithEditors) : null;
              const hasData = viewersSummary || editorsSummary;
              if (!hasData) return null;
              return (
                <AccordionItem value="inventory" className={styles.accordionCard}>
                  <AccordionHeader expandIconPosition="end" icon={<AppsListRegular />} className={styles.accordionHeaderTinted}>Inventory</AccordionHeader>
                  <AccordionPanel>
                    <div className={styles.detailGrid}>
                      {viewersSummary && (
                        <div className={styles.detailItem}>
                          <Text size={200} weight="semibold">Run-Only Users (Inventory)</Text>
                          <Text size={300}>{viewersSummary}</Text>
                        </div>
                      )}
                      {editorsSummary && (
                        <div className={styles.detailItem}>
                          <Text size={200} weight="semibold">Co-Owners (Inventory)</Text>
                          <Text size={300}>{editorsSummary}</Text>
                        </div>
                      )}
                    </div>
                  </AccordionPanel>
                </AccordionItem>
              );
            })()}

            {/* ── Owners ── */}
            <AccordionItem value="owners" className={styles.accordionCard}>
              <AccordionHeader expandIconPosition="end" icon={<PersonRegular />} className={styles.accordionHeaderTinted}>Owners</AccordionHeader>
              <AccordionPanel>
                <div className={styles.sectionBody}>
                  <PermissionsList
                    permissions={owners}
                    isLoading={ownersLoading}
                    error={ownersError}
                    emptyLabel="No owners found for this flow."
                    resolvedNames={permissionPrincipalNames}
                  />
                </div>
              </AccordionPanel>
            </AccordionItem>

            {/* ── Run-Only Users ── */}
            <AccordionItem value="runonlyusers" className={styles.accordionCard}>
              <AccordionHeader expandIconPosition="end" icon={<PeopleRegular />} className={styles.accordionHeaderTinted}>Run-Only Users</AccordionHeader>
              <AccordionPanel>
                <div className={styles.sectionBody}>
                  <PermissionsList
                    permissions={runOnlyUsers}
                    isLoading={runOnlyLoading}
                    error={runOnlyError}
                    emptyLabel="No run-only users found for this flow."
                    resolvedNames={permissionPrincipalNames}
                  />
                </div>
              </AccordionPanel>
            </AccordionItem>
          </Accordion>
          )}
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
