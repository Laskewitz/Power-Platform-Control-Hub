import { useState, useEffect, useRef } from 'react';
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
  Tab,
  TabList,
  useToastController,
  Toast,
  ToastTitle,
  ToastBody,
} from '@fluentui/react-components';
import {
  ArrowLeftRegular,
  DeleteRegular,
  ArrowClockwiseRegular,
  InfoFilled,
  ShieldCheckmarkRegular,
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
import type { Resource } from '../types/inventory.ts';
import { AGENT_HARNESS_LABELS, getAgentHarness } from '../types/inventory.ts';
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
import type { AnalysisResult } from '../services/flowAnalyzer.ts';
import { resolveOwner } from '../services/ownerCache.ts';
import { extractMessage } from '../utils/errorUtils.ts';
import { lcidToLabel } from '../utils/lcidUtils.ts';
import { addSelfAsEnvironmentAdmin } from '../services/environmentMutations.ts';
import AddSelfAsAdminBanner from './AddSelfAsAdminBanner.tsx';
import ConfirmDialog from './ConfirmDialog.tsx';
import AnalysisPosture from './AnalysisPosture.tsx';
import { hasText, formatSharedSummary, getStringArray, getCapabilityEntries } from '../utils/inventoryFormatters.ts';

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
    backgroundColor: '#0C141D',
    boxShadow: 'inset 3px 0 0 #FFB547',
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
    fontSize: '20px',
    fontWeight: tokens.fontWeightSemibold,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
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
  postureStrip: {
    display: 'grid',
    gridTemplateColumns: 'minmax(220px, 1fr) repeat(3, minmax(110px, auto))',
    gap: '1px',
    marginBottom: '12px',
    border: '1px solid #29404F',
    backgroundColor: '#20313E',
    '@media (max-width: 820px)': {
      gridTemplateColumns: 'repeat(2, minmax(0, 1fr))',
    },
  },
  postureLead: {
    display: 'grid',
    gap: '4px',
    alignContent: 'center',
    minHeight: '76px',
    padding: '12px 14px',
    backgroundColor: '#0C141D',
  },
  postureTitle: {
    color: '#F5FAFD',
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
  },
  postureSummary: {
    color: '#9CB0BF',
    fontSize: tokens.fontSizeBase200,
  },
  postureMetric: {
    display: 'grid',
    gap: '2px',
    alignContent: 'center',
    minHeight: '76px',
    padding: '12px 14px',
    backgroundColor: '#0C141D',
  },
  postureValue: {
    color: '#F5FAFD',
    fontSize: '22px',
    lineHeight: '24px',
    fontWeight: tokens.fontWeightBold,
    fontVariantNumeric: 'tabular-nums',
  },
  postureLabel: {
    color: '#8DA5B5',
    fontSize: '10px',
    fontWeight: tokens.fontWeightSemibold,
    letterSpacing: '0.08em',
    textTransform: 'uppercase',
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
  jsonBox: {
    fontFamily: 'Consolas, "Courier New", monospace',
    fontSize: tokens.fontSizeBase200,
    backgroundColor: '#060A0F',
    border: '1px solid #29404F',
    padding: tokens.spacingHorizontalM,
    overflowX: 'auto',
    whiteSpace: 'pre',
    lineHeight: '1.5',
    color: '#BEEFFF',
    maxHeight: '480px',
    overflowY: 'auto',
  },
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
    width: '100%',
    color: '#E5EEF5',
    borderTop: 0,
    borderRight: 0,
    borderLeft: 0,
    font: 'inherit',
    textAlign: 'left',
    ':hover': { backgroundColor: '#111827', boxShadow: 'inset 1px 0 0 #43D9FF' },
    ':focus-visible': { outline: '2px solid #43D9FF', outlineOffset: '-2px' },
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
  analysisAction: {
    width: 'fit-content',
    color: '#62E6FF',
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    textDecorationLine: 'none',
    ':hover': { textDecorationLine: 'underline' },
  },
  analysisChevron: {
    flexShrink: 0,
    color: '#8DA5B5',
    transitionProperty: 'transform',
    transitionDuration: '160ms',
    transitionTimingFunction: 'cubic-bezier(.16, 1, .3, 1)',
  },
  analysisChevronExpanded: {
    transform: 'rotate(90deg)',
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
  sectionBody: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
    paddingBottom: tokens.spacingVerticalM,
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
  componentStack: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
  },
  componentRegister: {
    display: 'grid',
    gridTemplateColumns: 'minmax(160px, 1fr) repeat(3, minmax(64px, auto))',
    gap: 0,
    fontSize: tokens.fontSizeBase200,
    borderTop: '1px solid #29404F',
    borderLeft: '1px solid #29404F',
    overflowX: 'auto',
  },
  componentHeaderCell: {
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    fontWeight: tokens.fontWeightSemibold,
    color: '#9CB0BF',
    backgroundColor: '#111827',
    borderRight: '1px solid #29404F',
    borderBottom: '1px solid #29404F',
    textTransform: 'uppercase',
    letterSpacing: '0.08em',
  },
  componentCell: {
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalM}`,
    borderRight: '1px solid #20313E',
    borderBottom: '1px solid #20313E',
    fontVariantNumeric: 'tabular-nums',
  },
  componentDirectory: {
    borderTop: '1px solid #29404F',
    borderLeft: '1px solid #29404F',
  },
  componentRow: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalM}`,
    borderRight: '1px solid #20313E',
    borderBottom: '1px solid #20313E',
    fontSize: tokens.fontSizeBase200,
  },
});

function formatDate(iso?: string): string {
  if (!iso) return '—';
  try { return new Date(iso).toLocaleString(); } catch { return iso; }
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
      severity: 'info',
    });
  }

  // 5. Language not set
  if (!bot.language) {
    results.push({
      id: 'no-language',
      title: 'No primary language configured',
      description: 'The agent does not have a primary language set in Dataverse.',
      recommendation: 'Ensure the agent has a configured primary language for accurate language routing.',
      severity: 'info',
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

  // 8. Multi-tenant access
  if (acp === 3) {
    results.push({
      id: 'agent-multitenant-access',
      title: 'Agent allows multi-tenant access',
      description: 'The access control policy is set to "Any (multi-tenant)" — users from any Azure AD tenant can interact with this agent.',
      recommendation: 'Only enable multi-tenant access for agents intentionally serving external organisations. Restrict to your own tenant for internal agents.',
      severity: 'warning',
    });
  }

  // 9. Stale agent (published more than 6 months ago)
  if (bot.publishedon) {
    const daysSince = (Date.now() - new Date(bot.publishedon).getTime()) / (1000 * 60 * 60 * 24);
    if (daysSince > 180) {
      results.push({
        id: 'agent-stale',
        title: `Agent not re-published in ${Math.floor(daysSince / 30)} months`,
        description: `The agent was last published ${Math.floor(daysSince)} days ago. Outdated agents may have stale topics, broken knowledge sources, or missing improvements.`,
        recommendation: "Review the agent's topics and knowledge sources, make any necessary updates, and re-publish to ensure users receive the latest experience.",
        severity: 'info',
      });
    }
  }

  // Component-based checks (only when components were loaded)
  if (components.length > 0) {
    // 10. Inactive custom topics
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

    // 11. High inactive topic ratio
    const allTopics = components.filter(c => c.componenttype === 0 || c.componenttype === 9);
    if (allTopics.length >= 5 && inactiveTopics.length > 0) {
      const ratio = inactiveTopics.length / allTopics.length;
      if (ratio > 0.5) {
        results.push({
          id: 'agent-high-inactive-ratio',
          title: `${Math.round(ratio * 100)}% of topics are disabled`,
          description: `${inactiveTopics.length} of ${allTopics.length} topics are inactive. A high proportion of disabled topics suggests the agent may be poorly maintained or partially decommissioned.`,
          recommendation: 'Review all disabled topics. Enable the ones that should be active, and delete any that are obsolete.',
          severity: 'warning',
        });
      }
    }

    // 12. No knowledge sources
    const knowledgeSources = components.filter(c => c.componenttype === 16);
    if (knowledgeSources.length === 0) {
      results.push({
        id: 'no-knowledge-sources',
        title: 'No knowledge sources configured',
        description: 'This agent has no knowledge sources. Without grounding data, the agent relies solely on its topics and generative AI defaults.',
        recommendation: "Consider adding knowledge sources (SharePoint, websites, uploaded files) to ground the agent's responses in your organisation's data.",
        severity: 'info',
      });
    }

    // 13. No test cases
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

const CAPABILITY_LABELS: Record<string, string> = {
  distinctPowerPlatformConnectorsOperations: 'Connector operations',
  distinctPowerPlatformConnectors: 'Connectors',
  distinctFlows: 'Flows',
};

function formatCapabilityLabel(value: string): string {
  return CAPABILITY_LABELS[value]
    ?? value
      .replace(/^distinct/i, '')
      .replace(/([A-Z])([A-Z][a-z])/g, '$1 $2')
      .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
      .trim();
}

function ComponentsView({ components }: { components: BotComponent[] }): ReactElement {
  const styles = useStyles();
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
    <div className={styles.componentStack}>
      {/* Summary table */}
      <div className={styles.componentRegister}>
        <div className={styles.componentHeaderCell}>Type</div>
        <div className={styles.componentHeaderCell}>Total</div>
        <div className={styles.componentHeaderCell}>Active</div>
        <div className={styles.componentHeaderCell}>Inactive</div>
        {sortedTypes.map((type, idx) => {
          const items = groups.get(type)!;
          const active = items.filter(c => Number(c.statecode) === 0).length;
          const inactive = items.length - active;
          const isLast = idx === sortedTypes.length - 1;
          const borderStyle = isLast ? 'none' : `1px solid ${tokens.colorNeutralStroke2}`;
          return (
            <>
              <div key={`label-${type}`} className={styles.componentCell} style={{ borderBottom: borderStyle }}>
                {GOVERNANCE_TYPES.has(type) && <BookOpenRegular fontSize={12} style={{ marginRight: 4, verticalAlign: 'middle', color: tokens.colorBrandForeground1 }} />}
                {COMPONENT_TYPE_LABELS[type] ?? `Type ${type}`}
              </div>
              <div key={`total-${type}`} className={styles.componentCell} style={{ textAlign: 'right', borderBottom: borderStyle }}>{items.length}</div>
              <div key={`active-${type}`} className={styles.componentCell} style={{ textAlign: 'right', color: tokens.colorStatusSuccessForeground1, borderBottom: borderStyle }}>{active > 0 ? active : '—'}</div>
              <div key={`inactive-${type}`} className={styles.componentCell} style={{ textAlign: 'right', color: inactive > 0 ? tokens.colorStatusWarningForeground1 : tokens.colorNeutralForeground3, borderBottom: borderStyle }}>
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
            <div className={styles.componentDirectory}>
              {items.map((comp, i) => {
                const isActive = Number(comp.statecode) === 0;
                const isLast = i === items.length - 1;
                return (
                  <div key={comp.botcomponentid ?? i} className={styles.componentRow} style={{ borderBottom: isLast ? 'none' : `1px solid ${tokens.colorNeutralStroke2}` }}>
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
  const bodyRef = useRef<HTMLDivElement>(null);

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
  const [confirmQuarantine, setConfirmQuarantine] = useState(false);
  const [openSections, setOpenSections] = useState<string[]>(['details']);
  const [contentTab, setContentTab] = useState<'overview' | 'analysis'>('overview');
  const [componentsError, setComponentsError] = useState<string | null>(null);

  const [addingAdmin, setAddingAdmin] = useState(false);
  const [addAdminError, setAddAdminError] = useState<string | null>(null);

  function handleSectionToggle(_: unknown, data: { openItems: string[] }) {
    setOpenSections(data.openItems);
  }

  async function loadDetails() {
    setBotLoading(true);
    setBotError(null);
    setComponents([]);
    setComponentsError(null);
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
          setResolvedOwner(await resolveOwner(result.bot._ownerid_value, envId));
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
        } catch (error) {
          setComponentsError(
            extractMessage(error instanceof Error ? error.message : 'Agent components could not be loaded.'),
          );
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

  // Load immediately — privilege errors are detected from the actual Dataverse call result
  useEffect(() => {
    void loadDetails();
  }, []);

  useEffect(() => {
    bodyRef.current?.scrollTo({ top: 0 });
  }, [resource.id]);

  async function handleMakeAdmin() {
    setAddingAdmin(true);
    setAddAdminError(null);
    try {
      await addSelfAsEnvironmentAdmin(envId);
      setBotError(null);
      void loadDetails();
      dispatchToast(
        <Toast><ToastTitle>Admin access applied</ToastTitle><ToastBody>Your admin role has been applied for this environment.</ToastBody></Toast>,
        { intent: 'success' },
      );
    } catch (e) {
      setAddAdminError(extractMessage(String(e)));
      dispatchToast(
        <Toast><ToastTitle>Failed to apply admin access</ToastTitle><ToastBody>{extractMessage(String(e))}</ToastBody></Toast>,
        { intent: 'error' },
      );
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

  const analysis = analyzeCopilotAgent(bot, components);

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
  const agentHarness = getAgentHarness(inventoryProps);
  const agentHarnessLabel = AGENT_HARNESS_LABELS[agentHarness];
  const inventoryChannels = getStringArray(inventoryProps.channels);
  const inventoryViewers = formatSharedSummary(inventoryProps.sharedWithViewers);
  const inventoryEditors = formatSharedSummary(inventoryProps.sharedWithEditors);
  const inventoryCapabilities = getCapabilityEntries(inventoryProps.capabilitiesCounts);
  const meaningfulCapabilities = inventoryCapabilities.filter(([, count]) => Number(count) > 0);
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
  const copilotStudioUrl = `https://copilotstudio.microsoft.com/environments/${envId}/bots/${botName}/overview`;

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
              <Badge appearance="tint" color="informative" size="small">Copilot Studio Agent</Badge>
              {agentHarnessLabel && (
                <Badge
                  appearance="tint"
                  color={agentHarness === 'github-copilot' ? 'brand' : 'success'}
                  size="small"
                >
                  {agentHarnessLabel}
                </Badge>
              )}
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
              <Text className={styles.envText}>{resource.environmentName}</Text>
            )}
          </div>
        </div>

        {/* Action bar */}
        <div className={styles.actionBar}>
          <Button
            appearance="subtle"
            icon={botLoading ? <Spinner size="tiny" /> : <ArrowClockwiseRegular />}
            disabled={actionLoading !== null || botLoading}
            onClick={() => void loadDetails()}
            size="small"
          >
            {botLoading ? 'Refreshing…' : 'Refresh agent'}
          </Button>
          <AddSelfAsAdminBanner environmentId={envId} variant="inline" onChanged={() => void loadDetails()} />
          <div className={styles.actionRegister}>
            <Button
              appearance="subtle"
              icon={actionLoading === 'quarantine' || actionLoading === 'unquarantine'
                ? <Spinner size="tiny" />
                : isQuarantined ? <LockOpenRegular /> : <LockClosedRegular />}
              disabled={actionLoading !== null || isQuarantined === null}
              onClick={() => {
                if (isQuarantined) {
                  void handleQuarantine();
                } else {
                  setConfirmQuarantine(true);
                }
              }}
              title={isQuarantined === null ? 'Quarantine status is unavailable for this agent.' : undefined}
              size="small"
            >
              {isQuarantined ? 'Unquarantine' : 'Quarantine'}
            </Button>
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

        {/* Body */}
        <div ref={bodyRef} className={styles.body}>
          {botLoading && (
            <Spinner size="small" label="Loading agent details…" style={{ marginBottom: tokens.spacingVerticalM }} />
          )}
          {botError && (() => {
            const isPrivilegeError =
              botError.includes('0x80040220') ||
              botError.includes('prvRead') ||
              (botError.includes('missing') && botError.includes('privilege'));
            return (
              <MessageBar intent={isPrivilegeError ? 'error' : 'warning'} style={{ marginBottom: tokens.spacingVerticalM }}>
                <MessageBarBody>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXS }}>
                    <span>
                      {isPrivilegeError
                        ? 'You are missing the required Dataverse privilege to load this bot record. Apply your admin role to gain access.'
                        : 'Could not load bot record from Dataverse.'}{' '}
                      <a
                        href={`https://copilotstudio.microsoft.com/environments/${envId}/bots/${botName}/overview`}
                        target="_blank"
                        rel="noopener noreferrer"
                        style={{ color: tokens.colorBrandForegroundLink }}
                      >
                        View in Copilot Studio ↗
                      </a>
                    </span>
                    {addAdminError && (
                      <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorStatusDangerForeground1 }}>
                        {addAdminError}
                      </Text>
                    )}
                    {!isPrivilegeError && (
                      <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>
                        {botError}
                      </Text>
                    )}
                  </div>
                </MessageBarBody>
                {isPrivilegeError && (
                  <MessageBarActions>
                    <Button
                      size="small"
                      appearance="primary"
                      icon={addingAdmin ? <Spinner size="tiny" /> : <PersonAddRegular />}
                      disabled={addingAdmin}
                      onClick={() => void handleMakeAdmin()}
                    >
                      {addingAdmin ? 'Applying…' : 'Apply admin access'}
                    </Button>
                  </MessageBarActions>
                )}
              </MessageBar>
            );
          })()}

          {contentTab === 'analysis' ? (
            <div className={styles.sectionBody}>
              <AnalysisPosture
                results={analysis}
                title="Copilot Studio agent posture"
                description="Governance signals derived from access, publishing, testing, grounding, and component configuration."
                isLoading={botLoading}
                emptyDescription="This agent follows the evaluated governance practices."
                action={{ label: 'Resolve in Copilot Studio', href: copilotStudioUrl }}
              />
            </div>
          ) : (
          <Accordion
            multiple
            collapsible
            openItems={openSections}
            onToggle={handleSectionToggle as (e: unknown, d: { openItems: string[] }) => void}
          >
            {/* ── Agent Details ── */}
            <AccordionItem value="details" className={styles.accordionCard}>
              <AccordionHeader expandIconPosition="end" icon={<InfoFilled />} className={styles.accordionHeaderTinted}>
                Agent Details
              </AccordionHeader>
              <AccordionPanel>
                <div className={styles.sectionBody}>
                  <div className={styles.detailGrid}>
                    <div className={styles.detailItem}>
                      <span className={styles.detailLabel}>Display Name</span>
                      <span className={styles.detailValue}>{displayName}</span>
                    </div>

                    {agentHarnessLabel && (
                      <div className={styles.detailItem}>
                        <span className={styles.detailLabel}>Agent Harness</span>
                        <span className={styles.detailValue}>{agentHarnessLabel}</span>
                      </div>
                    )}

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
                        <span className={styles.detailLabel}>Language</span>
                        <span className={styles.detailValue}>{lcidToLabel(Number(bot.language))}</span>
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

                    {(bot?.owneridname ?? resolvedOwner ?? resource.properties.resolvedOwnerName ?? bot?._ownerid_value) && (
                      <div className={styles.detailItem}>
                        <span className={styles.detailLabel}>Owner</span>
                        <span className={styles.detailValue}>{bot?.owneridname ?? resolvedOwner ?? resource.properties.resolvedOwnerName ?? bot?._ownerid_value}</span>
                      </div>
                    )}

                    {bot?.createdbyname && (
                      <div className={styles.detailItem}>
                        <span className={styles.detailLabel}>Created By</span>
                        <span className={styles.detailValue}>{bot.createdbyname}</span>
                      </div>
                    )}

                    <div className={styles.detailItem}>
                      <span className={styles.detailLabel}>Environment</span>
                      <span className={styles.detailValue}>
                        {resource.environmentName ?? envId}
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
                <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM, paddingBottom: tokens.spacingVerticalM }}>
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

                      {inventoryCapabilities.length > 0 && (
                        <div className={styles.detailItemWide}>
                          <span className={styles.detailLabel}>Connected Resources</span>
                          <span className={styles.detailValue} style={{ display: 'flex', flexWrap: 'wrap', gap: tokens.spacingHorizontalXS }}>
                            {meaningfulCapabilities.length > 0
                              ? meaningfulCapabilities.map(([name, count]) => (
                                  <Badge key={name} appearance="tint" color="brand" size="small">
                                    {formatCapabilityLabel(name)}: {count}
                                  </Badge>
                                ))
                              : 'No connector or flow dependencies detected.'}
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
            {(!botError || bot) && (
            <AccordionItem value="configuration" className={styles.accordionCard}>
              <AccordionHeader expandIconPosition="end" icon={<CodeRegular />} className={styles.accordionHeaderTinted}>
                <span style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS }}>
                  <BrainRegular fontSize={16} />
                  Technical details
                </span>
              </AccordionHeader>
              <AccordionPanel>
                <div className={styles.sectionBody}>
                  <div className={styles.detailGrid}>
                    <div className={styles.detailItem}>
                      <span className={styles.detailLabel}>Agent ID</span>
                      <span className={styles.detailValue} style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: tokens.fontSizeBase200, wordBreak: 'break-all' }}>
                        {bot?.botid ?? botName}
                      </span>
                    </div>
                    {bot?.schemaname && (
                      <div className={styles.detailItem}>
                        <span className={styles.detailLabel}>Schema Name</span>
                        <span className={styles.detailValue} style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: tokens.fontSizeBase200, wordBreak: 'break-all' }}>{bot.schemaname}</span>
                      </div>
                    )}
                    {instanceUrl && (
                      <div className={styles.detailItemWide}>
                        <span className={styles.detailLabel}>Dataverse URL</span>
                        <span className={styles.detailValue} style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: tokens.fontSizeBase200, wordBreak: 'break-all' }}>{instanceUrl}</span>
                      </div>
                    )}
                    {hasText(inventoryProps.entraAppId) && (
                      <div className={styles.detailItem}>
                        <span className={styles.detailLabel}>Entra App ID</span>
                        <span className={styles.detailValue} style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: tokens.fontSizeBase200, wordBreak: 'break-all' }}>{inventoryProps.entraAppId}</span>
                      </div>
                    )}
                    {hasText(inventoryProps.entraAgentId) && (
                      <div className={styles.detailItem}>
                        <span className={styles.detailLabel}>Entra Agent ID</span>
                        <span className={styles.detailValue} style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: tokens.fontSizeBase200, wordBreak: 'break-all' }}>{inventoryProps.entraAgentId}</span>
                      </div>
                    )}
                    {hasText(inventoryProps.entraAgentBlueprintId) && (
                      <div className={styles.detailItem}>
                        <span className={styles.detailLabel}>Entra Blueprint ID</span>
                        <span className={styles.detailValue} style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: tokens.fontSizeBase200, wordBreak: 'break-all' }}>{inventoryProps.entraAgentBlueprintId}</span>
                      </div>
                    )}
                    {bot?.authorizedsecuritygroupids && (
                      <div className={styles.detailItemWide}>
                        <span className={styles.detailLabel}>Authorized Security Group IDs</span>
                        <span className={styles.detailValue} style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: tokens.fontSizeBase200, wordBreak: 'break-all' }}>{bot.authorizedsecuritygroupids}</span>
                      </div>
                    )}
                  </div>
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
            )}

            {/* ── Components ── */}
            {(!botError || bot) && (
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
                <div className={styles.sectionBody}>
                  {(botLoading || componentsLoading) ? (
                    <Spinner size="tiny" label="Loading components…" />
                  ) : componentsError ? (
                    <MessageBar intent="warning">
                      <MessageBarBody>Agent components could not be loaded: {componentsError}</MessageBarBody>
                    </MessageBar>
                  ) : components.length === 0 ? (
                    <Text style={{ color: tokens.colorNeutralForeground3 }}>No components found for this agent.</Text>
                  ) : (
                    <ComponentsView components={components} />
                  )}
                </div>
              </AccordionPanel>
            </AccordionItem>
            )}
          </Accordion>
          )}
        </div>
      </div>

      <ConfirmDialog
        open={confirmQuarantine}
        title={`Quarantine ${displayName}?`}
        message="Quarantining prevents this agent from serving users until an administrator removes the quarantine. The agent and its configuration are retained."
        confirmLabel="Quarantine agent"
        isDangerous
        onConfirm={() => {
          setConfirmQuarantine(false);
          void handleQuarantine();
        }}
        onCancel={() => setConfirmQuarantine(false)}
      />

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
