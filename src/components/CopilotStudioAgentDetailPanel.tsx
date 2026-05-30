import { useState, useEffect } from 'react';
import type { ReactElement } from 'react';
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
  MessageBarActions,
  Tooltip,
  useToastController,
  Toast,
  ToastTitle,
  ToastBody,
} from '@fluentui/react-components';
import {
  ArrowLeftRegular,
  DeleteRegular,
  ArrowClockwiseRegular,
  InfoRegular,
  ShieldCheckmarkRegular,
  WarningRegular,
  ErrorCircleRegular,
  CheckmarkRegular,
  CheckmarkCircleRegular,
  DismissCircleRegular,
  BotRegular,
  LockClosedRegular,
  LockOpenRegular,
  CalendarRegular,
  BrainRegular,
  CodeRegular,
  AppsListRegular,
  ShieldPersonRegular,
  BookOpenRegular,
  PersonAddRegular,
} from '@fluentui/react-icons';
import type { InventorySharingSummary, Resource } from '../types/inventory.ts';
import type { Bots } from '../generated/models/BotsModel.ts';
import type { BotComponent } from '../services/dataverseConnectorService.ts';
import { COMPONENT_TYPE_LABELS } from '../services/dataverseConnectorService.ts';
import {
  fetchBotDetails,
  fetchBotComponents,
  getEnvironmentDataverseInfo,
  deleteCopilotAgent,
  getBotQuarantineStatus,
  quarantineBot,
  unquarantineBot,
} from '../services/copilotStudioService.ts';
import type { BotEnvironmentInfo } from '../services/copilotStudioService.ts';
import type { AnalysisResult, AnalysisSeverity } from '../services/flowAnalyzer.ts';
import { resolveUserIds } from '../services/userService.ts';
import { extractMessage } from '../utils/errorUtils.ts';
import { checkCurrentUserIsEnvAdmin, addSelfAsEnvironmentAdmin } from '../services/environmentMutations.ts';
import ConfirmDialog from './ConfirmDialog.tsx';

interface Props {
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
  title: {
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
    gridTemplateColumns: 'repeat(2, 1fr)',
    gap: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalXXL}`,
    alignItems: 'start',
    '@media (max-width: 768px)': {
      gridTemplateColumns: '1fr',
    },
  },
  detailItem: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: tokens.spacingVerticalXS,
    minWidth: 0,
  },
  detailItemWide: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: tokens.spacingVerticalXS,
    minWidth: 0,
    gridColumn: 'span 2',
  },
  detailLabel: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground3,
    textTransform: 'uppercase',
    letterSpacing: '0.04em',
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
  jsonBox: {
    fontFamily: 'Consolas, "Courier New", monospace',
    fontSize: tokens.fontSizeBase200,
    backgroundColor: tokens.colorNeutralBackground2,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    padding: tokens.spacingHorizontalM,
    overflowX: 'auto',
    whiteSpace: 'pre',
    lineHeight: '1.5',
    color: tokens.colorNeutralForeground1,
    maxHeight: '480px',
    overflowY: 'auto',
  },
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
  accordionCard: {
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    overflow: 'hidden',
    marginBottom: tokens.spacingVerticalS,
  },
  accordionHeaderTinted: {
    backgroundColor: tokens.colorNeutralBackground2,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
});

function formatDate(iso?: string): string {
  if (!iso) return '—';
  try { return new Date(iso).toLocaleString(); } catch { return iso; }
}

function hasText(value: unknown): value is string {
  return typeof value === 'string' && value.trim().length > 0;
}

function formatSharedSummary(summary?: InventorySharingSummary): string | null {
  if (!summary) return null;
  if (summary.entireTenant) return 'Entire tenant';

  const parts: string[] = [];
  if (typeof summary.userCount === 'number') {
    parts.push(`${summary.userCount} user${summary.userCount === 1 ? '' : 's'}`);
  }
  if (typeof summary.groupCount === 'number') {
    parts.push(`${summary.groupCount} group${summary.groupCount === 1 ? '' : 's'}`);
  }

  return parts.length > 0 ? parts.join(', ') : null;
}

function getStringArray(value: unknown): string[] {
  if (!Array.isArray(value)) return [];
  return value.filter((item): item is string => typeof item === 'string' && item.trim().length > 0);
}

function getCapabilityEntries(value: unknown): Array<[string, number]> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return [];
  return Object.entries(value).filter((entry): entry is [string, number] => typeof entry[1] === 'number');
}

function severityIcon(severity: AnalysisSeverity): ReactElement {
  if (severity === 'critical') return <ErrorCircleRegular style={{ color: tokens.colorStatusDangerForeground1, flexShrink: 0 }} fontSize={16} />;
  if (severity === 'warning') return <WarningRegular style={{ color: tokens.colorStatusWarningForeground1, flexShrink: 0 }} fontSize={16} />;
  return <CheckmarkRegular style={{ color: tokens.colorStatusSuccessForeground1, flexShrink: 0 }} fontSize={16} />;
}

// ── Best practice analysis for Copilot Studio agents ──────────────────────────

const ACCESS_CONTROL_LABELS: Record<number, string> = {
  0: 'Any', 1: 'Copilot readers', 2: 'Group membership', 3: 'Any (multi-tenant)',
};

function analyzeCopilotAgent(bot: Bots | null, components: BotComponent[]): AnalysisResult[] {
  const results: AnalysisResult[] = [];
  if (!bot) return results;

  // 1. Agent is inactive
  if (Number(bot.statecode) !== 0) {
    results.push({
      id: 'inactive-agent',
      title: 'Agent is inactive',
      description: 'This Copilot Studio agent is currently in an inactive state and will not respond to users.',
      recommendation: 'Activate the agent if it should be serving users, or delete it if it is no longer needed.',
      severity: 'warning',
    });
  }

  // 2. Never published
  if (!bot.publishedon) {
    results.push({
      id: 'never-published',
      title: 'Agent has never been published',
      description: 'This agent has no published version. Users cannot interact with an unpublished agent.',
      recommendation: 'Publish the agent from Copilot Studio to make it available to end users.',
      severity: 'warning',
    });
  }

  // 3. Authentication mode: Unspecified or None
  const authMode = Number(bot.authenticationmode);
  if (authMode === 0 || authMode === 1) {
    results.push({
      id: 'auth-mode-none',
      title: 'Authentication mode is None or Unspecified',
      description: `This agent's authentication mode is set to "${authMode === 0 ? 'Unspecified' : 'None'}". Anyone can interact with it without signing in.`,
      recommendation: 'Consider enabling authentication (AAD, Custom) if this agent handles sensitive data or should only be accessible to authorised users.',
      severity: 'warning',
    });
  }

  // 4. No configuration data
  if (!bot.configuration || bot.configuration === '{}') {
    results.push({
      id: 'empty-configuration',
      title: 'No configuration data found',
      description: 'The agent\'s configuration field in Dataverse is empty. The agent may not have been set up or the data may be incomplete.',
      recommendation: 'Verify the agent is correctly saved in Copilot Studio and that the Dataverse bots table is accessible.',
      severity: 'warning',
    });
  }

  // 5. Language not set
  if (!bot.language) {
    results.push({
      id: 'no-language',
      title: 'No primary language configured',
      description: 'The agent does not have a primary language set in Dataverse.',
      recommendation: 'Ensure the agent has a configured primary language for accurate language routing.',
      severity: 'warning',
    });
  }

  // 6. Access control policy open to everyone
  const acp = Number(bot.accesscontrolpolicy);
  if (acp === 0) {
    results.push({
      id: 'access-control-open',
      title: 'Access control allows anyone',
      description: 'The access control policy is set to "Any" — anyone can interact with this agent without group or reader restrictions.',
      recommendation: 'If this agent is for internal use only, consider setting access control to "Copilot readers" or "Group membership" to restrict access.',
      severity: 'info',
    });
  }

  // 7. Group membership policy but no security groups configured
  if (acp === 2 && !bot.authorizedsecuritygroupids) {
    results.push({
      id: 'group-membership-no-groups',
      title: 'Group membership access control has no groups configured',
      description: 'Access control is set to "Group membership" but no AAD security groups are configured. No users will be able to access this agent.',
      recommendation: 'Add the required AAD security group IDs to the Authorized Security Groups field in Copilot Studio.',
      severity: 'critical',
    });
  }

  // Component-based checks (only when components were loaded)
  if (components.length > 0) {
    // 8. Inactive custom topics
    const inactiveTopics = components.filter(
      c => (c.componenttype === 0 || c.componenttype === 9) && Number(c.statecode) !== 0,
    );
    if (inactiveTopics.length > 0) {
      results.push({
        id: 'inactive-topics',
        title: `${inactiveTopics.length} inactive topic${inactiveTopics.length !== 1 ? 's' : ''} found`,
        description: `${inactiveTopics.length} topic${inactiveTopics.length !== 1 ? 's' : ''} ${inactiveTopics.length === 1 ? 'is' : 'are'} disabled: ${inactiveTopics.map(t => t.name ?? 'Unknown').slice(0, 5).join(', ')}${inactiveTopics.length > 5 ? '…' : ''}.`,
        recommendation: 'Disabled topics are not active. Review these topics and either enable, update or delete them to keep the agent clean.',
        severity: 'warning',
      });
    }

    // 9. No knowledge sources
    const knowledgeSources = components.filter(c => c.componenttype === 16);
    if (knowledgeSources.length === 0) {
      results.push({
        id: 'no-knowledge-sources',
        title: 'No knowledge sources configured',
        description: 'This agent has no knowledge sources. Without grounding data, the agent relies solely on its topics and generative AI defaults.',
        recommendation: 'Consider adding knowledge sources (SharePoint, websites, uploaded files) to ground the agent\'s responses in your organisation\'s data.',
        severity: 'info',
      });
    }

    // 10. No test cases
    const testCases = components.filter(c => c.componenttype === 19);
    if (testCases.length === 0) {
      results.push({
        id: 'no-test-cases',
        title: 'No test cases defined',
        description: 'This agent has no test cases. Test cases help validate that the agent responds correctly after changes.',
        recommendation: 'Add test cases in Copilot Studio to catch regressions and verify agent behaviour before publishing.',
        severity: 'warning',
      });
    }
  }

  return results;
}

// ── Component types shown in the governance view ─────────────────────────────
const GOVERNANCE_TYPES = new Set([0, 1, 9, 13, 15, 16, 17, 18, 19]);

function ComponentsView({ components }: { components: BotComponent[] }): ReactElement {
  // Group by componenttype
  const groups = new Map<number, BotComponent[]>();
  for (const c of components) {
    const t = c.componenttype ?? -1;
    if (!groups.has(t)) groups.set(t, []);
    groups.get(t)!.push(c);
  }

  // Sort groups: governance types first, then the rest
  const sortedTypes = [...groups.keys()].sort((a, b) => {
    const aGov = GOVERNANCE_TYPES.has(a) ? 0 : 1;
    const bGov = GOVERNANCE_TYPES.has(b) ? 0 : 1;
    return aGov !== bGov ? aGov - bGov : a - b;
  });

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM }}>
      {/* Summary table */}
      <div style={{
        display: 'grid',
        gridTemplateColumns: '1fr auto auto auto',
        gap: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalM}`,
        fontSize: tokens.fontSizeBase200,
        borderRadius: tokens.borderRadiusMedium,
        border: `1px solid ${tokens.colorNeutralStroke2}`,
        overflow: 'hidden',
      }}>
        <div style={{ padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`, fontWeight: tokens.fontWeightSemibold, backgroundColor: tokens.colorNeutralBackground2, borderBottom: `1px solid ${tokens.colorNeutralStroke2}` }}>Type</div>
        <div style={{ padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`, fontWeight: tokens.fontWeightSemibold, backgroundColor: tokens.colorNeutralBackground2, textAlign: 'right', borderBottom: `1px solid ${tokens.colorNeutralStroke2}` }}>Total</div>
        <div style={{ padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`, fontWeight: tokens.fontWeightSemibold, backgroundColor: tokens.colorNeutralBackground2, textAlign: 'right', borderBottom: `1px solid ${tokens.colorNeutralStroke2}` }}>Active</div>
        <div style={{ padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`, fontWeight: tokens.fontWeightSemibold, backgroundColor: tokens.colorNeutralBackground2, textAlign: 'right', borderBottom: `1px solid ${tokens.colorNeutralStroke2}` }}>Inactive</div>
        {sortedTypes.map((type, idx) => {
          const items = groups.get(type)!;
          const active = items.filter(c => Number(c.statecode) === 0).length;
          const inactive = items.length - active;
          const isLast = idx === sortedTypes.length - 1;
          const borderStyle = isLast ? 'none' : `1px solid ${tokens.colorNeutralStroke2}`;
          return (
            <>
              <div key={`label-${type}`} style={{ padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalM}`, borderBottom: borderStyle }}>
                {GOVERNANCE_TYPES.has(type) && <BookOpenRegular fontSize={12} style={{ marginRight: 4, verticalAlign: 'middle', color: tokens.colorBrandForeground1 }} />}
                {COMPONENT_TYPE_LABELS[type] ?? `Type ${type}`}
              </div>
              <div key={`total-${type}`} style={{ padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalM}`, textAlign: 'right', borderBottom: borderStyle }}>{items.length}</div>
              <div key={`active-${type}`} style={{ padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalM}`, textAlign: 'right', color: tokens.colorStatusSuccessForeground1, borderBottom: borderStyle }}>{active > 0 ? active : '—'}</div>
              <div key={`inactive-${type}`} style={{ padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalM}`, textAlign: 'right', color: inactive > 0 ? tokens.colorStatusWarningForeground1 : tokens.colorNeutralForeground3, borderBottom: borderStyle }}>
                {inactive > 0 ? inactive : '—'}
              </div>
            </>
          );
        })}
      </div>

      {/* Detailed topic list (governance types only) */}
      {sortedTypes.filter(t => GOVERNANCE_TYPES.has(t)).map(type => {
        const items = groups.get(type)!;
        return (
          <div key={type}>
            <Text style={{ fontWeight: tokens.fontWeightSemibold, fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground2, display: 'block', marginBottom: tokens.spacingVerticalXS }}>
              {COMPONENT_TYPE_LABELS[type] ?? `Type ${type}`} ({items.length})
            </Text>
            <div style={{
              borderRadius: tokens.borderRadiusMedium,
              border: `1px solid ${tokens.colorNeutralStroke2}`,
              overflow: 'hidden',
            }}>
              {items.map((comp, i) => {
                const isActive = Number(comp.statecode) === 0;
                const isLast = i === items.length - 1;
                return (
                  <div key={comp.botcomponentid ?? i} style={{
                    display: 'flex',
                    alignItems: 'center',
                    gap: tokens.spacingHorizontalS,
                    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalM}`,
                    borderBottom: isLast ? 'none' : `1px solid ${tokens.colorNeutralStroke2}`,
                    fontSize: tokens.fontSizeBase200,
                  }}>
                    {isActive
                      ? <CheckmarkCircleRegular fontSize={14} style={{ color: tokens.colorStatusSuccessForeground1, flexShrink: 0 }} />
                      : <DismissCircleRegular fontSize={14} style={{ color: tokens.colorStatusWarningForeground1, flexShrink: 0 }} />
                    }
                    <span style={{ flex: 1, wordBreak: 'break-word' }}>{comp.name ?? '(unnamed)'}</span>
                    {!isActive && (
                      <Badge appearance="tint" color="warning" size="tiny">Inactive</Badge>
                    )}
                    {comp.category && (
                      <span style={{ color: tokens.colorNeutralForeground3 }}>{comp.category}</span>
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        );
      })}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────

export default function CopilotStudioAgentDetailPanel({ resource, onClose, onDeleted }: Props): ReactElement {
  const styles = useStyles();
  const { dispatchToast } = useToastController();

  const displayName = (resource.properties.displayName as string) ?? resource.name;
  const envId = (resource.properties.environmentId as string) ?? '';
  const botName = resource.name; // GUID or internal name from Inventory API

  const [bot, setBot] = useState<Bots | null>(null);
  const [botLoading, setBotLoading] = useState(false);
  const [botError, setBotError] = useState<string | null>(null);

  const [instanceUrl, setInstanceUrl] = useState<string | null>(null);
  const [dataverseError, setDataverseError] = useState<string | null>(null);
  const [isQuarantined, setIsQuarantined] = useState<boolean | null>(null);

  const [components, setComponents] = useState<BotComponent[]>([]);
  const [componentsLoading, setComponentsLoading] = useState(false);
  const [resolvedOwner, setResolvedOwner] = useState<string | null>(null);

  const [actionLoading, setActionLoading] = useState<string | null>(null);
  const [confirmDelete, setConfirmDelete] = useState(false);
  const [expandedAnalysis, setExpandedAnalysis] = useState<Set<string>>(new Set());
  const [openSections, setOpenSections] = useState<string[]>(['details', 'inventory', 'analysis']);

  // Admin access gate — check before attempting any Dataverse call
  const [adminCheck, setAdminCheck] = useState<'loading' | 'admin' | 'notAdmin'>('loading');
  const [addingAdmin, setAddingAdmin] = useState(false);
  const [addAdminError, setAddAdminError] = useState<string | null>(null);

  function handleSectionToggle(_: unknown, data: { openItems: string[] }) {
    setOpenSections(data.openItems);
  }

  async function loadDetails() {
    setBotLoading(true);
    setBotError(null);
    setComponents([]);
    setResolvedOwner(null);
    setDataverseError(null);
    try {
      let envInstanceUrl: string | null = null;
      if (envId) {
        const envInfo = await getEnvironmentDataverseInfo(envId);
        envInstanceUrl = envInfo.instanceUrl ?? null;
        if (!envInstanceUrl && envInfo.dataverseError) {
          setDataverseError(envInfo.dataverseError);
        }
      }
      setInstanceUrl(envInstanceUrl);

      const envInfo: BotEnvironmentInfo = { instanceUrl: envInstanceUrl ?? undefined };
      const result = await fetchBotDetails(botName, envInfo);
      setBot(result.bot);
      if (!result.bot && result.crossEnvError) {
        setBotError(result.crossEnvError);
      }

      // Resolve owner GUID to display name (best-effort)
      if (result.bot?._ownerid_value) {
        try {
          const names = await resolveUserIds([result.bot._ownerid_value]);
          setResolvedOwner(names.get(result.bot._ownerid_value) ?? null);
        } catch {
          // non-critical
        }
      }

      // Fetch bot components from cross-env Dataverse (best-effort, non-blocking display)
      if (result.bot?.botid && envInstanceUrl) {
        setComponentsLoading(true);
        try {
          const comps = await fetchBotComponents(envInstanceUrl, result.bot.botid);
          setComponents(comps);
        } catch {
          // components are optional
        } finally {
          setComponentsLoading(false);
        }
      }

      // Fetch quarantine status (best-effort)
      try {
        const q = await getBotQuarantineStatus(envId, botName);
        setIsQuarantined(q);
      } catch {
        setIsQuarantined(null);
      }
    } catch (e) {
      setBotError(extractMessage(e instanceof Error ? e.message : 'Failed to load agent details'));
    } finally {
      setBotLoading(false);
    }
  }

  useEffect(() => {
    if (!envId) { setAdminCheck('admin'); return; }
    checkCurrentUserIsEnvAdmin(envId)
      .then(isAdmin => setAdminCheck(isAdmin ? 'admin' : 'notAdmin'))
      .catch(() => setAdminCheck('admin')); // on check failure, proceed anyway
  }, []);

  // Only load Dataverse data once admin access is confirmed
  useEffect(() => {
    if (adminCheck === 'admin') void loadDetails();
  }, [adminCheck]);

  async function handleMakeAdmin() {
    setAddingAdmin(true);
    setAddAdminError(null);
    try {
      await addSelfAsEnvironmentAdmin(envId);
      setAdminCheck('admin'); // triggers loadDetails via useEffect
    } catch (e) {
      setAddAdminError(extractMessage(String(e)));
    } finally {
      setAddingAdmin(false);
    }
  }

  async function handleDelete() {
    setActionLoading('delete');
    try {
      await deleteCopilotAgent(envId, botName);
      dispatchToast(
        <Toast><ToastTitle>Agent deleted</ToastTitle><ToastBody>"{displayName}" has been deleted.</ToastBody></Toast>,
        { intent: 'success' },
      );
      onDeleted(resource.name);
    } catch (e) {
      dispatchToast(
        <Toast><ToastTitle>Delete failed</ToastTitle><ToastBody>{extractMessage(e instanceof Error ? e.message : 'Unknown error')}</ToastBody></Toast>,
        { intent: 'error' },
      );
    } finally {
      setActionLoading(null);
    }
  }

  async function handleQuarantine() {
    const action = isQuarantined ? 'unquarantine' : 'quarantine';
    setActionLoading(action);
    try {
      if (isQuarantined) {
        await unquarantineBot(envId, botName);
        setIsQuarantined(false);
        dispatchToast(<Toast><ToastTitle>Agent unquarantined</ToastTitle></Toast>, { intent: 'success' });
      } else {
        await quarantineBot(envId, botName);
        setIsQuarantined(true);
        dispatchToast(<Toast><ToastTitle>Agent quarantined</ToastTitle></Toast>, { intent: 'success' });
      }
    } catch (e) {
      dispatchToast(
        <Toast><ToastTitle>Action failed</ToastTitle><ToastBody>{extractMessage(e instanceof Error ? e.message : 'Unknown error')}</ToastBody></Toast>,
        { intent: 'error' },
      );
    } finally {
      setActionLoading(null);
    }
  }

  function toggleAnalysis(id: string) {
    setExpandedAnalysis((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  }

  const analysis = analyzeCopilotAgent(bot, components);
  const issueCount = analysis.filter((r) => r.severity !== 'info').length;
  const hasIssues = issueCount > 0;

  // Parse configuration JSON for display
  let configDisplay = '—';
  if (bot?.configuration) {
    try {
      configDisplay = JSON.stringify(JSON.parse(bot.configuration) as unknown, null, 2);
    } catch {
      configDisplay = bot.configuration;
    }
  }

  const isActive = Number(bot?.statecode) === 0;

  const authModeLabels: Record<number, string> = {
    0: 'Unspecified', 1: 'None', 2: 'Integrated', 3: 'Custom',
    5: 'Azure AD v2', 6: 'Azure AD v2 (Certificate)', 10: 'Generic OAuth 2',
  };
  const inventoryProps = resource.properties;
  const inventoryChannels = getStringArray(inventoryProps.channels);
  const inventoryViewers = formatSharedSummary(inventoryProps.sharedWithViewers);
  const inventoryEditors = formatSharedSummary(inventoryProps.sharedWithEditors);
  const inventoryCapabilities = getCapabilityEntries(inventoryProps.capabilitiesCounts);
  const hasInventoryDetails = [
    hasText(inventoryProps.orchestration),
    hasText(inventoryProps.model),
    hasText(inventoryProps.authentication),
    hasText(inventoryProps.createdIn),
    inventoryChannels.length > 0,
    Boolean(inventoryViewers),
    Boolean(inventoryEditors),
    hasText(inventoryProps.entraAppId),
    hasText(inventoryProps.entraAgentId),
    hasText(inventoryProps.entraAgentBlueprintId),
    inventoryCapabilities.length > 0,
  ].some(Boolean);

  return (
    <>
      <div className={styles.root}>
        {/* Header */}
        <div className={styles.header}>
          <Button appearance="subtle" icon={<ArrowLeftRegular />} onClick={onClose} size="small">
            Back to Resources
          </Button>
          <BotRegular fontSize={20} style={{ color: tokens.colorPaletteGreenForeground1, flexShrink: 0 }} />
          <div className={styles.headerMeta}>
            <div className={styles.titleRow}>
              <Tooltip content={displayName} relationship="label">
                <Text className={styles.title}>{displayName}</Text>
              </Tooltip>
              <Badge appearance="tint" color="success" size="small">Copilot Studio Agent</Badge>
              {bot && (
                <Badge
                  appearance="tint"
                  color={isActive ? 'success' : 'warning'}
                  size="small"
                  icon={isActive ? <CheckmarkCircleRegular /> : <DismissCircleRegular />}
                >
                  {isActive ? 'Active' : 'Inactive'}
                </Badge>
              )}
              {isQuarantined === true && (
                <Badge appearance="tint" color="danger" size="small" icon={<LockClosedRegular />}>
                  Quarantined
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
          <Button
            appearance="subtle"
            icon={actionLoading === 'quarantine' || actionLoading === 'unquarantine'
              ? <Spinner size="tiny" />
              : isQuarantined ? <LockOpenRegular /> : <LockClosedRegular />}
            disabled={actionLoading !== null || isQuarantined === null}
            onClick={() => void handleQuarantine()}
            size="small"
          >
            {isQuarantined ? 'Unquarantine' : 'Quarantine'}
          </Button>
          <Button
            appearance="subtle"
            icon={<ArrowClockwiseRegular />}
            disabled={actionLoading !== null || botLoading}
            onClick={() => void loadDetails()}
            size="small"
            title="Refresh"
          />
          <div style={{ marginLeft: 'auto' }}>
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

        {/* Body */}
        <div className={styles.body}>
          {/* Access gate: check admin before loading any Dataverse data */}
          {adminCheck === 'loading' && (
            <Spinner size="small" label="Checking environment access…" style={{ marginBottom: tokens.spacingVerticalM }} />
          )}
          {adminCheck === 'notAdmin' && (
            <MessageBar intent={addAdminError ? 'error' : 'warning'} style={{ marginBottom: tokens.spacingVerticalM, flexShrink: 0 }}>
              <MessageBarBody>
                {addAdminError
                  ? `Failed to add you as System Administrator: ${addAdminError}`
                  : 'You are not a System Administrator on this environment. Bot details require Dataverse access.'}
              </MessageBarBody>
              {!addAdminError && (
                <MessageBarActions>
                  <Button
                    size="small"
                    appearance="primary"
                    icon={addingAdmin ? <Spinner size="tiny" /> : <PersonAddRegular />}
                    disabled={addingAdmin}
                    onClick={() => void handleMakeAdmin()}
                  >
                    {addingAdmin ? 'Adding…' : 'Add me as System Administrator'}
                  </Button>
                </MessageBarActions>
              )}
            </MessageBar>
          )}
          {adminCheck === 'admin' && (
            <>
          {botLoading && (
            <Spinner size="small" label="Loading agent details…" style={{ marginBottom: tokens.spacingVerticalM }} />
          )}
          {botError && (
            <MessageBar intent="warning" style={{ marginBottom: tokens.spacingVerticalM }}>
              <MessageBarBody>
                <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXS }}>
                  <span>
                    Could not load bot record from Dataverse.{' '}
                    <a
                      href={`https://copilotstudio.microsoft.com/environments/${envId}/bots/${botName}/overview`}
                      target="_blank"
                      rel="noopener noreferrer"
                      style={{ color: tokens.colorBrandForegroundLink }}
                    >
                      View in Copilot Studio ↗
                    </a>
                  </span>
                  <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>
                    {botError}
                  </Text>
                </div>
              </MessageBarBody>
            </MessageBar>
          )}

          <Accordion
            multiple
            collapsible
            openItems={openSections}
            onToggle={handleSectionToggle as (e: unknown, d: { openItems: string[] }) => void}
          >
            {/* ── Agent Details ── */}
            <AccordionItem value="details" className={styles.accordionCard}>
              <AccordionHeader expandIconPosition="end" icon={<InfoRegular />} className={styles.accordionHeaderTinted}>
                Agent Details
              </AccordionHeader>
              <AccordionPanel>
                <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalL, paddingBottom: tokens.spacingVerticalL }}>
                  <div className={styles.detailGrid}>
                    <div className={styles.detailItem}>
                      <span className={styles.detailLabel}>Display Name</span>
                      <span className={styles.detailValue}>{displayName}</span>
                    </div>

                    {bot && (
                      <div className={styles.detailItem}>
                        <span className={styles.detailLabel}>Status</span>
                        <span className={styles.detailValue}>
                          <Badge appearance="tint" color={isActive ? 'success' : 'warning'} size="small"
                            icon={isActive ? <CheckmarkCircleRegular /> : <DismissCircleRegular />}>
                            {isActive ? 'Active' : 'Inactive'}
                          </Badge>
                        </span>
                      </div>
                    )}

                    {bot?.language !== undefined && bot.language !== null && (
                      <div className={styles.detailItem}>
                        <span className={styles.detailLabel}>Language (LCID)</span>
                        <span className={styles.detailValue}>{bot.language}</span>
                      </div>
                    )}

                    {bot?.authenticationmode !== undefined && (
                      <div className={styles.detailItem}>
                        <span className={styles.detailLabel}>Authentication</span>
                        <span className={styles.detailValue}>
                          {authModeLabels[Number(bot.authenticationmode)] ?? String(bot.authenticationmode)}
                        </span>
                      </div>
                    )}

                    {bot?.accesscontrolpolicy !== undefined && (
                      <div className={styles.detailItem}>
                        <span className={styles.detailLabel}>Access Control</span>
                        <span className={styles.detailValue}>
                          <ShieldPersonRegular fontSize={14} style={{ marginRight: 4, verticalAlign: 'middle' }} />
                          {ACCESS_CONTROL_LABELS[Number(bot.accesscontrolpolicy)] ?? String(bot.accesscontrolpolicy)}
                        </span>
                      </div>
                    )}

                    {isQuarantined !== null && (
                      <div className={styles.detailItem}>
                        <span className={styles.detailLabel}>Quarantine</span>
                        <span className={styles.detailValue}>
                          <Badge
                            appearance="tint"
                            color={isQuarantined ? 'danger' : 'success'}
                            size="small"
                            icon={isQuarantined ? <LockClosedRegular /> : <LockOpenRegular />}
                          >
                            {isQuarantined ? 'Quarantined' : 'Not Quarantined'}
                          </Badge>
                        </span>
                      </div>
                    )}

                    {bot?.publishedon && (
                      <div className={styles.detailItem}>
                        <span className={styles.detailLabel}>Last Published</span>
                        <span className={styles.detailValue}>
                          <CalendarRegular fontSize={14} style={{ marginRight: 4, verticalAlign: 'middle' }} />
                          {formatDate(bot.publishedon)}
                          {bot.publishedby && <span style={{ color: tokens.colorNeutralForeground3, marginLeft: 6 }}>by {bot.publishedby}</span>}
                        </span>
                      </div>
                    )}

                    {(bot?.createdon ?? resource.properties.createdAt) && (
                      <div className={styles.detailItem}>
                        <span className={styles.detailLabel}>Created</span>
                        <span className={styles.detailValue}>
                          <CalendarRegular fontSize={14} style={{ marginRight: 4, verticalAlign: 'middle' }} />
                          {formatDate((bot?.createdon ?? resource.properties.createdAt) as string)}
                        </span>
                      </div>
                    )}

                    {bot?.modifiedon && (
                      <div className={styles.detailItem}>
                        <span className={styles.detailLabel}>Last Modified</span>
                        <span className={styles.detailValue}>
                          <CalendarRegular fontSize={14} style={{ marginRight: 4, verticalAlign: 'middle' }} />
                          {formatDate(bot.modifiedon)}
                        </span>
                      </div>
                    )}

                    {(bot?.owneridname ?? resolvedOwner ?? bot?._ownerid_value) && (
                      <div className={styles.detailItem}>
                        <span className={styles.detailLabel}>Owner</span>
                        <span className={styles.detailValue}>{bot?.owneridname ?? resolvedOwner ?? bot?._ownerid_value}</span>
                      </div>
                    )}

                    {bot?.createdbyname && (
                      <div className={styles.detailItem}>
                        <span className={styles.detailLabel}>Created By</span>
                        <span className={styles.detailValue}>{bot.createdbyname}</span>
                      </div>
                    )}

                    {bot?.schemaname && (
                      <div className={styles.detailItem}>
                        <span className={styles.detailLabel}>Schema Name</span>
                        <span className={styles.detailValue} style={{ fontSize: tokens.fontSizeBase200, wordBreak: 'break-all', color: tokens.colorNeutralForeground3 }}>
                          {bot.schemaname}
                        </span>
                      </div>
                    )}

                    {bot?.authorizedsecuritygroupids && (
                      <div className={styles.detailItem}>
                        <span className={styles.detailLabel}>Authorized Groups</span>
                        <span className={styles.detailValue} style={{ fontSize: tokens.fontSizeBase200, wordBreak: 'break-all', color: tokens.colorNeutralForeground3 }}>
                          {bot.authorizedsecuritygroupids}
                        </span>
                      </div>
                    )}

                    <div className={styles.detailItem}>
                      <span className={styles.detailLabel}>Environment</span>
                      <span className={styles.detailValue}>
                        {resource.environmentName ?? envId}
                      </span>
                    </div>

                    {instanceUrl && (
                      <div className={styles.detailItem}>
                        <span className={styles.detailLabel}>Dataverse URL</span>
                        <span className={styles.detailValue} style={{ fontSize: tokens.fontSizeBase200, wordBreak: 'break-all', color: tokens.colorNeutralForeground3 }}>
                          {instanceUrl}
                        </span>
                      </div>
                    )}

                    <div className={styles.detailItem}>
                      <span className={styles.detailLabel}>Agent ID</span>
                      <span className={styles.detailValue} style={{ fontSize: tokens.fontSizeBase200, wordBreak: 'break-all', color: tokens.colorNeutralForeground3 }}>
                        {bot?.botid ?? botName}
                      </span>
                    </div>
                  </div>

                  {!bot && !botLoading && !botError && (
                    <MessageBar intent="info">
                      <MessageBarBody>
                        {instanceUrl
                          ? <>Bot record not found in Dataverse for this environment.{' '}</>
                          : <>Unable to resolve Dataverse instance URL for this environment.{dataverseError ? <><br /><span style={{ fontSize: tokens.fontSizeBase200, opacity: 0.8 }}>{dataverseError}</span></> : ' Bot record unavailable.'}{' '}</>
                        }
                        <a
                          href={`https://copilotstudio.microsoft.com/environments/${envId}/bots/${botName}/overview`}
                          target="_blank"
                          rel="noopener noreferrer"
                          style={{ color: tokens.colorBrandForegroundLink }}
                        >
                          View in Copilot Studio ↗
                        </a>
                      </MessageBarBody>
                    </MessageBar>
                  )}
                </div>
              </AccordionPanel>
            </AccordionItem>

            {/* ── Inventory ── */}
            <AccordionItem value="inventory" className={styles.accordionCard}>
              <AccordionHeader expandIconPosition="end" icon={<AppsListRegular />} className={styles.accordionHeaderTinted}>
                Inventory
              </AccordionHeader>
              <AccordionPanel>
                <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalL, paddingBottom: tokens.spacingVerticalL }}>
                  {hasInventoryDetails ? (
                    <div className={styles.detailGrid}>
                      {hasText(inventoryProps.orchestration) && (
                        <div className={styles.detailItem}>
                          <span className={styles.detailLabel}>Orchestration</span>
                          <span className={styles.detailValue}>{inventoryProps.orchestration}</span>
                        </div>
                      )}

                      {hasText(inventoryProps.model) && (
                        <div className={styles.detailItem}>
                          <span className={styles.detailLabel}>AI Model</span>
                          <span className={styles.detailValue}>{inventoryProps.model}</span>
                        </div>
                      )}

                      {hasText(inventoryProps.authentication) && (
                        <div className={styles.detailItem}>
                          <span className={styles.detailLabel}>Authentication</span>
                          <span className={styles.detailValue}>{inventoryProps.authentication}</span>
                        </div>
                      )}

                      {hasText(inventoryProps.createdIn) && (
                        <div className={styles.detailItem}>
                          <span className={styles.detailLabel}>Created In</span>
                          <span className={styles.detailValue}>{inventoryProps.createdIn}</span>
                        </div>
                      )}

                      {inventoryViewers && (
                        <div className={styles.detailItem}>
                          <span className={styles.detailLabel}>Shared With Viewers</span>
                          <span className={styles.detailValue}>{inventoryViewers}</span>
                        </div>
                      )}

                      {inventoryEditors && (
                        <div className={styles.detailItem}>
                          <span className={styles.detailLabel}>Shared With Editors</span>
                          <span className={styles.detailValue}>{inventoryEditors}</span>
                        </div>
                      )}

                      {inventoryChannels.length > 0 && (
                        <div className={styles.detailItemWide}>
                          <span className={styles.detailLabel}>Channels</span>
                          <span className={styles.detailValue} style={{ display: 'flex', flexWrap: 'wrap', gap: tokens.spacingHorizontalXS }}>
                            {inventoryChannels.map((channel) => (
                              <Badge key={channel} appearance="tint" color="informative" size="small">{channel}</Badge>
                            ))}
                          </span>
                        </div>
                      )}

                      {hasText(inventoryProps.entraAppId) && (
                        <div className={styles.detailItemWide}>
                          <span className={styles.detailLabel}>Entra App ID</span>
                          <span className={styles.detailValue} style={{ fontSize: tokens.fontSizeBase200, wordBreak: 'break-all', fontFamily: 'Consolas, "Courier New", monospace', color: tokens.colorNeutralForeground3 }}>
                            {inventoryProps.entraAppId}
                          </span>
                        </div>
                      )}

                      {hasText(inventoryProps.entraAgentId) && (
                        <div className={styles.detailItemWide}>
                          <span className={styles.detailLabel}>Entra Agent ID</span>
                          <span className={styles.detailValue} style={{ fontSize: tokens.fontSizeBase200, wordBreak: 'break-all', fontFamily: 'Consolas, "Courier New", monospace', color: tokens.colorNeutralForeground3 }}>
                            {inventoryProps.entraAgentId}
                          </span>
                        </div>
                      )}

                      {hasText(inventoryProps.entraAgentBlueprintId) && (
                        <div className={styles.detailItemWide}>
                          <span className={styles.detailLabel}>Entra Blueprint ID</span>
                          <span className={styles.detailValue} style={{ fontSize: tokens.fontSizeBase200, wordBreak: 'break-all', fontFamily: 'Consolas, "Courier New", monospace', color: tokens.colorNeutralForeground3 }}>
                            {inventoryProps.entraAgentBlueprintId}
                          </span>
                        </div>
                      )}

                      {inventoryCapabilities.length > 0 && (
                        <div className={styles.detailItemWide}>
                          <span className={styles.detailLabel}>Capabilities</span>
                          <span className={styles.detailValue} style={{ display: 'flex', flexWrap: 'wrap', gap: tokens.spacingHorizontalXS }}>
                            {inventoryCapabilities.map(([name, count]) => (
                              <Badge key={name} appearance="tint" color="brand" size="small">{name}: {count}</Badge>
                            ))}
                          </span>
                        </div>
                      )}
                    </div>
                  ) : (
                    <Text style={{ color: tokens.colorNeutralForeground3 }}>No inventory-only agent fields available.</Text>
                  )}
                </div>
              </AccordionPanel>
            </AccordionItem>

            {/* ── Configuration / Definition ── */}
            <AccordionItem value="configuration" className={styles.accordionCard}>
              <AccordionHeader expandIconPosition="end" icon={<CodeRegular />} className={styles.accordionHeaderTinted}>
                <span style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS }}>
                  <BrainRegular fontSize={16} />
                  Definition (Configuration)
                </span>
              </AccordionHeader>
              <AccordionPanel>
                <div style={{ paddingBottom: tokens.spacingVerticalL }}>
                  {botLoading ? (
                    <Spinner size="tiny" label="Loading configuration…" />
                  ) : bot?.configuration ? (
                    <>
                      <Text className={styles.sectionTitle} style={{ marginBottom: tokens.spacingVerticalS, display: 'block' }}>
                        Raw configuration from Dataverse <code>bots.configuration</code> column
                      </Text>
                      <div className={styles.jsonBox}>{configDisplay}</div>
                    </>
                  ) : (
                    <Text style={{ color: tokens.colorNeutralForeground3 }}>
                      {bot
                        ? 'No configuration data in this bot record.'
                        : <>Bot record not available — configuration cannot be shown.{' '}
                            <a
                              href={`https://copilotstudio.microsoft.com/environments/${envId}/bots/${botName}/overview`}
                              target="_blank"
                              rel="noopener noreferrer"
                              style={{ color: tokens.colorBrandForegroundLink }}
                            >
                              View in Copilot Studio ↗
                            </a>
                          </>
                      }
                    </Text>
                  )}
                </div>
              </AccordionPanel>
            </AccordionItem>

            {/* ── Best Practice Analysis ── */}
            <AccordionItem value="analysis" className={styles.accordionCard}>
              <AccordionHeader expandIconPosition="end" icon={<ShieldCheckmarkRegular />} className={styles.accordionHeaderTinted}>
                <span style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS }}>
                  Best Practice Analysis
                  {!botLoading && analysis.length > 0 && (
                    <Badge appearance="filled" size="small" color={hasIssues ? 'warning' : 'success'}>
                      {hasIssues ? `${issueCount} issue${issueCount !== 1 ? 's' : ''}` : 'OK'}
                    </Badge>
                  )}
                </span>
              </AccordionHeader>
              <AccordionPanel>
                <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalS, paddingBottom: tokens.spacingVerticalL }}>
                  {botLoading ? (
                    <Spinner size="tiny" label="Running analysis…" />
                  ) : analysis.length === 0 ? (
                    <div className={styles.analysisSummary}>
                      <CheckmarkCircleRegular fontSize={20} style={{ color: tokens.colorStatusSuccessForeground1 }} />
                      <Text>
                        {bot
                          ? 'All checks passed. No issues found.'
                          : 'Dataverse record not available — analysis requires bot data from Dataverse.'
                        }
                      </Text>
                    </div>
                  ) : (
                    <>
                      <div className={styles.analysisSummary}>
                        {hasIssues
                          ? <WarningRegular fontSize={20} style={{ color: tokens.colorStatusWarningForeground1 }} />
                          : <CheckmarkCircleRegular fontSize={20} style={{ color: tokens.colorStatusSuccessForeground1 }} />
                        }
                        <Text>
                          {hasIssues
                            ? `${issueCount} issue${issueCount !== 1 ? 's' : ''} found`
                            : 'All checks passed'}
                          {analysis.length > issueCount && ` (${analysis.length - issueCount} info)`}
                        </Text>
                      </div>
                      <div className={styles.analysisList}>
                        {[...analysis].sort((a, b) => {
                          const order = { critical: 0, warning: 1, info: 2 };
                          return (order[a.severity] ?? 3) - (order[b.severity] ?? 3);
                        }).map((item) => {
                          const isExpanded = expandedAnalysis.has(item.id);
                          return (
                            <div key={item.id}>
                              <div
                                className={`${styles.analysisRow}${isExpanded ? ` ${styles.analysisRowExpanded}` : ''}`}
                                onClick={() => toggleAnalysis(item.id)}
                              >
                                {severityIcon(item.severity)}
                                <Text className={styles.analysisTitle}>{item.title}</Text>
                                <Badge appearance="tint" color={
                                  item.severity === 'critical' ? 'danger' :
                                  item.severity === 'warning' ? 'warning' : 'informative'
                                } size="small">
                                  {item.severity}
                                </Badge>
                              </div>
                              {isExpanded && (
                                <div className={styles.analysisRowDetail}>
                                  <Text className={styles.analysisDesc}>{item.description}</Text>
                                  {item.recommendation && (
                                    <Text className={styles.analysisRec}>💡 {item.recommendation}</Text>
                                  )}
                                </div>
                              )}
                            </div>
                          );
                        })}
                      </div>
                    </>
                  )}
                </div>
              </AccordionPanel>
            </AccordionItem>

            {/* ── Components ── */}
            <AccordionItem value="components" className={styles.accordionCard}>
              <AccordionHeader expandIconPosition="end" icon={<AppsListRegular />} className={styles.accordionHeaderTinted}>
                <span style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS }}>
                  Components
                  {components.length > 0 && (
                    <Badge appearance="tint" color="informative" size="small">{components.length}</Badge>
                  )}
                </span>
              </AccordionHeader>
              <AccordionPanel>
                <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM, paddingBottom: tokens.spacingVerticalL }}>
                  {(botLoading || componentsLoading) ? (
                    <Spinner size="tiny" label="Loading components…" />
                  ) : !bot ? (
                    <Text style={{ color: tokens.colorNeutralForeground3 }}>
                      Bot record not available — components cannot be loaded.
                    </Text>
                  ) : components.length === 0 ? (
                    <Text style={{ color: tokens.colorNeutralForeground3 }}>No components found for this agent.</Text>
                  ) : (
                    <ComponentsView components={components} />
                  )}
                </div>
              </AccordionPanel>
            </AccordionItem>
          </Accordion>
            </>
          )}
        </div>
      </div>

      <ConfirmDialog
        open={confirmDelete}
        title="Delete Agent?"
        message={`Are you sure you want to delete "${displayName}"? This action cannot be undone.`}
        confirmLabel="Delete"
        isDangerous
        onConfirm={() => {
          setConfirmDelete(false);
          void handleDelete();
        }}
        onCancel={() => setConfirmDelete(false)}
      />
    </>
  );
}
