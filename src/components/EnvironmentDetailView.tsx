import { useEffect, useMemo, useState } from 'react';
import type { ReactElement } from 'react';
import {
  makeStyles,
  shorthands,
  tokens,
  Text,
  Badge,
  Button,
  Input,
  Dropdown,
  Option,
  Menu,
  MenuTrigger,
  MenuList,
  MenuItem,
  MenuPopover,
  Tab,
  TabList,
} from '@fluentui/react-components';
import {
  SearchRegular,
  GlobeRegular,
  CalendarRegular,
  PersonRegular,
  KeyRegular,
  PlayRegular,
  StopRegular,
  ShieldRegular,
  ShieldDismissRegular,
  SaveRegular,
  ChevronRightRegular,
  ChevronDownRegular,
  BoxRegular,
  AppsFilled,
  FlowRegular,
  BotRegular,
  DatabaseRegular,
  LayerRegular,
  SubtractCircleRegular,
  ShieldCheckmarkRegular,
  OpenRegular,
  DocumentSearchRegular,
} from '@fluentui/react-icons';
import type { Resource } from '../types/inventory.ts';
import type { EnvironmentGroup } from '../types/admin.ts';
import type { AnalysisResult } from '../services/flowAnalyzer.ts';
import { RESOURCE_TYPE_LABELS } from '../types/inventory.ts';
import EmptyState from './EmptyState.tsx';
import AnalysisPosture from './AnalysisPosture.tsx';

const DETAIL_PANEL_TYPES = new Set([
  'microsoft.powerautomate/cloudflows',
  'microsoft.powerautomate/agentflows',
  'microsoft.powerautomate/m365agentflows',
  'microsoft.powerapps/apps',
  'microsoft.powerapps/canvasapps',
  'microsoft.copilotstudio/agents',
]);
import ConfirmDialog from './ConfirmDialog.tsx';
import { useMutation } from '../hooks/useMutation.tsx';
import {
  enableEnvironment,
  disableEnvironment,
  enableManagedEnvironment,
  disableManagedEnvironment,
  createEnvironmentBackup,
} from '../services/environmentMutations.ts';
import { PowerPlatformforAdminsV2Service } from '../generated/services/PowerPlatformforAdminsV2Service.ts';
import BackupDialog from './BackupDialog.tsx';
import EnvironmentGroupDialog from './EnvironmentGroupDialog.tsx';
import { useOwners } from '../services/ownerCache.ts';
import AddSelfAsAdminBanner from './AddSelfAsAdminBanner.tsx';
import { formatDate } from '../utils/formatDate.ts';

interface EnvironmentDetailViewProps {
  environment: Resource;
  resources: Resource[];
  envGroups?: EnvironmentGroup[];
  onBack: () => void;
  onRefreshEnvironments?: () => Promise<void>;
  onResourceSelect?: (resource: Resource) => void;
}

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    gap: '16px',
    height: '100%',
    padding: '22px 28px 28px',
    overflow: 'hidden',
    backgroundColor: tokens.colorNeutralBackground2,
    backgroundImage: 'linear-gradient(rgba(67, 217, 255, 0.025) 1px, transparent 1px), linear-gradient(90deg, rgba(67, 217, 255, 0.025) 1px, transparent 1px)',
    backgroundSize: '32px 32px',
    '@media (max-width: 768px)': {
      padding: '16px',
    },
  },

  // ── Hero ──────────────────────────────────────────────────────────────────
  hero: {
    position: 'relative',
    overflow: 'hidden',
    color: '#F4FBFD',
    border: '1px solid #29404F',
    backgroundColor: '#0B121A',
    backgroundImage: 'linear-gradient(90deg, rgba(67, 217, 255, 0.055) 1px, transparent 1px)',
    backgroundSize: '40px 100%',
    boxShadow: '0 12px 30px rgba(0, 0, 0, 0.32)',
    flexShrink: 0,
    display: 'flex',
    flexDirection: 'column',
    '::before': {
      content: '""',
      position: 'absolute',
      inset: '0 0 0 auto',
      width: '15%',
      backgroundColor: '#10232F',
      clipPath: 'polygon(32% 0, 100% 0, 100% 100%, 0 100%)',
    },
    '::after': {
      content: '""',
      position: 'absolute',
      right: '15%',
      bottom: 0,
      width: '68px',
      height: '5px',
      backgroundColor: '#FFB547',
    },
  },
  breadcrumb: {
    position: 'relative',
    zIndex: 1,
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    minHeight: '38px',
    padding: '0 20px',
    color: '#91A8B5',
    fontSize: tokens.fontSizeBase200,
    borderBottom: '1px solid #20313E',
    backgroundColor: 'rgba(6, 10, 15, 0.48)',
  },
  breadcrumbLink: {
    color: tokens.colorBrandForeground1,
    cursor: 'pointer',
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase200,
    ':hover': { textDecoration: 'underline' },
  },
  heroBody: {
    position: 'relative',
    zIndex: 1,
    display: 'flex',
    alignItems: 'center',
    gap: '14px',
    flexWrap: 'wrap',
    padding: '20px 22px 14px',
  },
  heroIcon: {
    fontSize: '2rem',
    color: '#43D9FF',
    flexShrink: 0,
  },
  heroName: {
    color: '#F4FBFD',
    fontSize: '30px',
    lineHeight: '34px',
    letterSpacing: '-0.035em',
    fontWeight: tokens.fontWeightSemibold,
    flex: 1,
    minWidth: 0,
  },
  heroBadges: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    flexWrap: 'wrap',
  },
  heroMeta: {
    position: 'relative',
    zIndex: 1,
    display: 'grid',
    gridTemplateColumns: 'repeat(4, minmax(0, 1fr))',
    borderTop: '1px solid #20313E',
    '@media (max-width: 900px)': {
      gridTemplateColumns: 'repeat(2, minmax(0, 1fr))',
    },
    '@media (max-width: 560px)': {
      gridTemplateColumns: '1fr',
    },
  },
  heroMetaItem: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    minWidth: 0,
    minHeight: '42px',
    padding: '9px 16px',
    fontSize: tokens.fontSizeBase200,
    color: '#91A8B5',
    borderRight: '1px solid #20313E',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
    ':last-child': { borderRight: 0 },
  },

  // ── Content ───────────────────────────────────────────────────────────────
  content: {
    display: 'flex',
    flexDirection: 'column',
    flex: 1,
    minHeight: 0,
    overflow: 'hidden',
    gap: '14px',
  },
  // Settings tab uses the content area as the scroll container directly
  contentSettings: {
    display: 'flex',
    flexDirection: 'column',
    flex: 1,
    minHeight: 0,
    overflowY: 'auto',
    padding: `${tokens.spacingVerticalL} ${tokens.spacingHorizontalXL}`,
    gap: tokens.spacingVerticalM,
  },

  // ── Info strip ────────────────────────────────────────────────────────────
  infoStrip: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fill, minmax(180px, 1fr))',
    gap: tokens.spacingHorizontalM,
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: tokens.borderRadiusMedium,
    padding: tokens.spacingVerticalM,
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
  },
  infoItem: {
    display: 'flex',
    flexDirection: 'column',
    gap: '2px',
    padding: `0 ${tokens.spacingHorizontalM}`,
    borderRightWidth: '1px',
    borderRightStyle: 'solid',
    borderRightColor: tokens.colorNeutralStroke2,
    ':last-child': { borderRight: 'none' },
  },
  infoLabel: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    textTransform: 'uppercase',
    letterSpacing: '0.04em',
    fontWeight: tokens.fontWeightSemibold,
  },
  infoValue: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
    maxWidth: '100%',
  },

  // ── Stats row ─────────────────────────────────────────────────────────────
  statsRow: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(170px, 1fr))',
    gap: '1px',
    flexShrink: 0,
    border: '1px solid #29404F',
    backgroundColor: '#29404F',
  },
  statPill: {
    display: 'grid',
    gridTemplateColumns: 'auto auto minmax(0, 1fr)',
    alignItems: 'center',
    gap: '9px',
    minHeight: '50px',
    padding: '10px 14px',
    color: '#B6D2D8',
    border: 0,
    borderRadius: 0,
    backgroundColor: '#0C141D',
    cursor: 'pointer',
    textAlign: 'left',
    ':hover': {
      color: '#F4FBFD',
      backgroundColor: '#10232F',
    },
    ':focus-visible': {
      outline: '2px solid #43D9FF',
      outlineOffset: '-2px',
    },
    '&[aria-pressed="true"]': {
      color: '#F4FBFD',
      backgroundColor: '#0B1D26',
      boxShadow: 'inset 0 -2px 0 #43D9FF',
    },
  },
  statCount: {
    color: '#43D9FF',
    fontSize: '20px',
    lineHeight: 1,
    fontWeight: tokens.fontWeightSemibold,
    fontVariantNumeric: 'tabular-nums',
  },
  statLabel: {
    fontSize: tokens.fontSizeBase200,
    color: 'inherit',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },

  // ── Resources table ───────────────────────────────────────────────────────
  tableSection: {
    display: 'flex',
    flexDirection: 'column',
    flex: 1,
    overflow: 'hidden',
    backgroundColor: '#0C141D',
    border: '1px solid #29404F',
    boxShadow: '0 12px 28px rgba(0, 0, 0, 0.28)',
  },
  tableHeader: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    minHeight: '52px',
    padding: '10px 14px',
    borderBottomWidth: '1px',
    borderBottomStyle: 'solid',
    borderBottomColor: '#29404F',
    flexShrink: 0,
    gap: tokens.spacingHorizontalM,
  },
  tableTitle: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
  },
  tableControls: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    flexWrap: 'wrap',
  },
  tableWrapper: {
    overflow: 'auto',
    flex: 1,
  },
  resourceHeader: {
    display: 'grid',
    gridTemplateColumns: 'minmax(280px, 1.45fr) minmax(180px, 0.8fr) 110px 110px minmax(190px, 0.9fr) 38px',
    gap: '12px',
    alignItems: 'center',
    minWidth: '1050px',
    minHeight: '34px',
    padding: '0 12px',
    color: '#78909C',
    fontSize: '10px',
    fontWeight: tokens.fontWeightSemibold,
    letterSpacing: '0.08em',
    textTransform: 'uppercase',
    borderBottom: '1px solid #29404F',
    backgroundColor: '#111827',
  },
  resourceRow: {
    display: 'grid',
    gridTemplateColumns: 'minmax(280px, 1.45fr) minmax(180px, 0.8fr) 110px 110px minmax(190px, 0.9fr) 38px',
    gap: '12px',
    alignItems: 'center',
    minWidth: '1050px',
    minHeight: '52px',
    padding: '8px 12px',
    borderBottom: '1px solid #20313E',
    backgroundColor: '#0C141D',
    transitionProperty: 'background-color, box-shadow',
    transitionDuration: '160ms',
    ':hover': {
      backgroundColor: '#101D27',
      boxShadow: 'inset 2px 0 0 #43D9FF',
    },
  },
  resourceIdentity: {
    display: 'flex',
    alignItems: 'center',
    gap: '10px',
    minWidth: 0,
  },
  resourceIcon: {
    display: 'grid',
    placeItems: 'center',
    width: '28px',
    height: '28px',
    flexShrink: 0,
    color: '#43D9FF',
    border: '1px solid #29404F',
    backgroundColor: '#0B1D26',
  },
  resourcePrimary: {
    color: '#F4FBFD',
    fontWeight: tokens.fontWeightSemibold,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  resourceSecondary: {
    color: '#91A8B5',
    fontSize: tokens.fontSizeBase200,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },

  // ── Content tabs ──────────────────────────────────────────────────────────
  contentTabs: {
    backgroundColor: 'transparent',
    borderBottomWidth: '1px',
    borderBottomStyle: 'solid',
    borderBottomColor: tokens.colorNeutralStroke2,
    paddingLeft: 0,
    flexShrink: 0,
  },

  // ── Settings panel ────────────────────────────────────────────────────────
  settingsOuter: {
    display: 'flex',
    flexDirection: 'column',
    flex: 1,
    minHeight: 0,
    overflow: 'hidden',
    gap: tokens.spacingVerticalM,
  },
  settingsScroll: {
    flex: 1,
    minHeight: 0,
    overflowY: 'auto',
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
    paddingRight: tokens.spacingHorizontalXS,
  },
  settingsGroup: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: tokens.borderRadiusMedium,
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
    overflow: 'hidden',
  },
  settingsGroupHeader: {
    display: 'flex',
    alignItems: 'center',
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalL}`,
    backgroundColor: tokens.colorNeutralBackground3,
    borderBottomWidth: '1px',
    borderBottomStyle: 'solid',
    borderBottomColor: tokens.colorNeutralStroke2,
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    textTransform: 'uppercase',
    letterSpacing: '0.05em',
    color: tokens.colorNeutralForeground3,
  },
  settingRow: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: tokens.spacingHorizontalL,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalL}`,
    borderBottomWidth: '1px',
    borderBottomStyle: 'solid',
    borderBottomColor: tokens.colorNeutralStroke2,
    ':last-child': { borderBottom: 'none' },
  },
  settingLabel: {
    fontSize: tokens.fontSizeBase300,
    color: tokens.colorNeutralForeground1,
    flex: 1,
  },
  settingsActions: {
    display: 'flex',
    justifyContent: 'flex-end',
    gap: tokens.spacingHorizontalM,
    flexShrink: 0,
    position: 'sticky',
    bottom: `-${tokens.spacingVerticalL}`,
    backgroundColor: tokens.colorNeutralBackground2,
    paddingTop: tokens.spacingVerticalM,
    paddingBottom: tokens.spacingVerticalM,
    marginTop: 'auto',
    borderTopWidth: '1px',
    borderTopStyle: 'solid',
    borderTopColor: tokens.colorNeutralStroke2,
  },

  // ── BPA analysis ─────────────────────────────────────────────────────────
  analysisScroll: {
    flex: 1,
    overflowY: 'auto',
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
  },
  analysisSummary: {
    display: 'grid',
    gridTemplateColumns: 'minmax(260px, 1fr) repeat(3, minmax(110px, 0.28fr))',
    border: '1px solid #29404F',
    backgroundColor: '#29404F',
    gap: '1px',
    '@media (max-width: 760px)': {
      gridTemplateColumns: '1fr',
    },
  },
  analysisLead: {
    display: 'grid',
    gap: '3px',
    padding: '14px 16px',
    backgroundColor: '#0B1D26',
  },
  analysisMetric: {
    display: 'grid',
    alignContent: 'center',
    gap: '2px',
    padding: '10px 14px',
    backgroundColor: '#0C141D',
  },
  analysisMetricValue: {
    color: '#F4FBFD',
    fontSize: '20px',
    lineHeight: 1,
    fontWeight: tokens.fontWeightSemibold,
    fontVariantNumeric: 'tabular-nums',
  },
  analysisMetricLabel: {
    color: '#78909C',
    fontSize: '10px',
    fontWeight: tokens.fontWeightSemibold,
    letterSpacing: '0.08em',
    textTransform: 'uppercase',
  },
  analysisRow: {
    display: 'grid',
    gridTemplateColumns: '32px minmax(0, 1fr) auto 24px',
    alignItems: 'center',
    gap: '12px',
    minHeight: '60px',
    padding: '10px 12px',
    borderRadius: 0,
    backgroundColor: '#0C141D',
    border: '1px solid #29404F',
    cursor: 'pointer',
    ':hover': { backgroundColor: '#101D27' },
    ':focus-visible': {
      outline: '2px solid #43D9FF',
      outlineOffset: '-2px',
    },
  },
  findingIndex: {
    color: '#78909C',
    fontSize: '11px',
    fontWeight: tokens.fontWeightSemibold,
    fontVariantNumeric: 'tabular-nums',
    textAlign: 'center',
  },
  findingTitle: {
    color: '#F4FBFD',
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
  },
  findingCopy: {
    display: 'grid',
    gap: '3px',
    minWidth: 0,
  },
  findingChevron: {
    color: '#91A8B5',
    fontSize: '14px',
    transitionProperty: 'transform',
    transitionDuration: '160ms',
  },
  analysisRowDetail: {
    padding: '12px 14px 16px',
    backgroundColor: '#09121A',
    borderLeft: '1px solid #29404F',
    borderRight: '1px solid #29404F',
    borderBottom: '1px solid #29404F',
    borderRadius: 0,
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
  },
  findingDetailGrid: {
    display: 'grid',
    gridTemplateColumns: 'minmax(0, 0.8fr) minmax(0, 1.2fr)',
    gap: '1px',
    border: '1px solid #20313E',
    backgroundColor: '#20313E',
    '@media (max-width: 720px)': {
      gridTemplateColumns: '1fr',
    },
  },
  findingDetailSection: {
    display: 'grid',
    alignContent: 'start',
    gap: '8px',
    padding: '14px',
    color: '#B6D2D8',
    backgroundColor: '#0C141D',
  },
  findingDetailLabel: {
    color: '#78909C',
    fontSize: '10px',
    fontWeight: tokens.fontWeightSemibold,
    letterSpacing: '0.08em',
    textTransform: 'uppercase',
  },
  recommendationHeading: {
    display: 'flex',
    alignItems: 'center',
    gap: '7px',
    color: '#43D9FF',
  },
});

function analyzeEnvironment(
  envType: string,
  isManaged: boolean,
  envGroupId: string | undefined,
  resourceCount: number,
  domainName: string | undefined,
): AnalysisResult[] {
  const results: AnalysisResult[] = [];

  if (envType.toLowerCase() === 'default') {
    results.push({
      id: 'env-default',
      severity: 'warning',
      title: 'This is the Default environment',
      description: 'The Default environment is shared by all users in the tenant and cannot be deleted. It is often used for ad-hoc development, making governance difficult. Resources created here are visible to all makers.',
      recommendation: 'Avoid using the Default environment for production workloads or sensitive data. Create dedicated environments with appropriate DLP policies for specific use cases.',
    });
  }

  if (envType.toLowerCase() === 'trial') {
    results.push({
      id: 'env-trial',
      severity: 'warning',
      title: 'Trial environment will expire',
      description: 'Trial environments have a limited lifespan (typically 30 days). Any resources in this environment will be deleted when it expires.',
      recommendation: 'Migrate production workloads to a permanent environment before the trial expires. Convert the environment to Sandbox or Production if the work should be preserved.',
    });
  }

  if (!isManaged) {
    results.push({
      id: 'env-not-managed',
      severity: 'info',
      title: 'Not a Managed Environment',
      description: 'Managed Environment features are not enabled. Managed Environments provide additional governance controls such as solution checker enforcement, sharing limits, usage insights, and weekly maker digests.',
      recommendation: 'Consider enabling Managed Environments for better governance visibility. This requires a Power Apps / Power Automate Premium licence for makers in the environment.',
    });
  }

  if (!envGroupId) {
    results.push({
      id: 'env-no-group',
      severity: 'info',
      title: 'Not assigned to an Environment Group',
      description: 'This environment is not part of any Environment Group. Environment groups allow you to apply policies and settings consistently across multiple environments.',
      recommendation: 'Assign this environment to an appropriate Environment Group so it inherits shared governance policies.',
    });
  }

  if (resourceCount > 200) {
    results.push({
      id: 'env-large',
      severity: 'info',
      title: `Large environment (${resourceCount} resources)`,
      description: `This environment contains ${resourceCount} resources. Large, unmanaged environments are difficult to govern and may contain stale or unused resources.`,
      recommendation: 'Review the resources in this environment periodically. Remove or move unused resources to reduce complexity and licensing costs.',
    });
  }

  // Auto-generated domain names start with "org" followed by 7-12 lowercase alphanumeric chars
  // with no hyphens or underscores — e.g. "org1234567a". A manually set URL is meaningful.
  if (domainName && /^org[a-z0-9]{6,12}$/.test(domainName)) {
    results.push({
      id: 'env-autogenerated-url',
      severity: 'info',
      title: 'Environment URL appears auto-generated',
      description: `The environment URL domain is "${domainName}", which looks auto-generated by Power Platform. Auto-generated URLs are hard to remember and don't communicate the environment's purpose.`,
      recommendation: 'Set a meaningful, descriptive environment URL (e.g. "contoso-production" or "hr-sandbox") via the Power Platform Admin Center under Environment > Edit. Note: the URL can only be changed once.',
    });
  }

  return results;
}

function envTypeColor(envType: string): 'brand' | 'success' | 'warning' | 'important' | 'informative' {
  switch (envType.toLowerCase()) {
    case 'production': return 'brand';
    case 'sandbox': return 'warning';
    case 'developer': return 'success';
    case 'default': return 'informative';
    default: return 'important';
  }
}

function resourceTypeIcon(type: string): ReactElement {
  if (type.includes('cloudflows') || type.includes('agentflows') || type.includes('m365agent')) return <FlowRegular />;
  if (type.includes('agents')) return <BotRegular />;
  if (type.includes('powerapps') || type.includes('codeapps')) return <AppsFilled />;
  if (type.includes('powerpages')) return <DatabaseRegular />;
  return <BoxRegular />;
}

function EnvironmentBpaSection({
  envType, isManaged, envGroupId, resourceCount, domainName,
}: {
  envType: string;
  isManaged: boolean;
  envGroupId: string | undefined;
  resourceCount: number;
  domainName: string | undefined;
}): ReactElement {
  const results = analyzeEnvironment(envType, isManaged, envGroupId, resourceCount, domainName);
  return (
    <AnalysisPosture
      results={results}
      title="Environment posture"
      description="Governance signals derived from environment configuration and inventory."
      emptyDescription="This environment follows the evaluated governance practices."
    />
  );
}

export default function EnvironmentDetailView({
  environment: env,
  resources,
  envGroups = [],
  onBack,
  onRefreshEnvironments,
  onResourceSelect,
}: EnvironmentDetailViewProps): ReactElement {
  const styles = useStyles();
  const [search, setSearch] = useState('');
  const [typeFilter, setTypeFilter] = useState('all');
  const [confirmAction, setConfirmAction] = useState<{ type: 'disable' | 'disableManaged' } | null>(null);
  const [showBackup, setShowBackup] = useState(false);
  const [groupDialogMode, setGroupDialogMode] = useState<'add' | 'remove' | null>(null);
  const [isPending, setIsPending] = useState(false);
  const [contentTab, setContentTab] = useState<'resources' | 'analysis'>('resources');
  const [domainName, setDomainName] = useState<string | undefined>(undefined);

  // Lazy-fetch domain name from Admin V2 API when Analysis tab first opens
  useEffect(() => {
    if (contentTab !== 'analysis' || domainName !== undefined) return;
    PowerPlatformforAdminsV2Service.GetEnvironmentByIdForUser(env.name, '2024-10-01')
      .then((result) => {
        if (result.success && result.data) {
          setDomainName(result.data.domainName ?? '');
        }
      })
      .catch(() => setDomainName(''));
  }, [contentTab, env.name, domainName]);

  const displayName = env.properties.displayName ?? env.name;
  const envType = (env.properties.environmentType ?? 'Unknown') as string;
  const isManaged = env.properties.isManaged === true;
  const region = env.location ?? '—';
  const createdAt = env.properties.createdAt
    ? formatDate(env.properties.createdAt as string)
    : '—';
  const envGroupId = env.properties['environmentGroupId'] as string | undefined;
  const groupMap = useMemo(() => new Map(envGroups.map((g) => [g.id, g.displayName])), [envGroups]);

  // Runtime state from Inventory API: properties.states.runtime.id = 'Enabled' | 'Disabled'
  const runtimeState = ((env.properties['states'] as Record<string, unknown> | undefined)
    ?.['runtime'] as Record<string, unknown> | undefined)
    ?.['id'] as string | undefined;
  const isDisabled = runtimeState?.toLowerCase() === 'disabled';
  const isEnabled = !runtimeState || runtimeState?.toLowerCase() === 'enabled';

  const { execute: execEnable } = useMutation(enableEnvironment, {
    successMessage: 'Enable environment request submitted.',
    onSuccess: () => void onRefreshEnvironments?.(),
  });
  const { execute: execDisable } = useMutation(disableEnvironment, {
    successMessage: 'Disable environment request submitted.',
    onSuccess: () => void onRefreshEnvironments?.(),
  });
  const { execute: execEnableManaged } = useMutation(enableManagedEnvironment, {
    successMessage: 'Enable managed environment request submitted.',
    onSuccess: () => void onRefreshEnvironments?.(),
  });
  const { execute: execDisableManaged } = useMutation(disableManagedEnvironment, {
    successMessage: 'Disable managed environment request submitted.',
    onSuccess: () => void onRefreshEnvironments?.(),
  });
  const { execute: execBackup, isLoading: isBackupLoading } = useMutation(createEnvironmentBackup, {
    successMessage: 'Backup request submitted.',
    onSuccess: () => setShowBackup(false),
  });

  async function runAction(action: () => Promise<unknown>) {
    setIsPending(true);
    await action();
    setIsPending(false);
  }

  const envResources = useMemo(() => {
    const envNameLower = displayName.toLowerCase();
    return resources.filter(
      (r) => (r.environmentName ?? '').toLowerCase() === envNameLower && !r.type.includes('environments'),
    );
  }, [resources, displayName]);

  const typeCounts = useMemo(() => {
    const map = new Map<string, number>();
    for (const r of envResources) map.set(r.type, (map.get(r.type) ?? 0) + 1);
    return map;
  }, [envResources]);

  const filteredResources = useMemo(() => {
    const term = search.toLowerCase();
    return envResources.filter((resource) => {
      const matchesType = typeFilter === 'all' || resource.type === typeFilter;
      const matchesSearch = !term
        || (resource.properties.displayName ?? resource.name).toLowerCase().includes(term)
        || (RESOURCE_TYPE_LABELS[resource.type] ?? resource.type).toLowerCase().includes(term);
      return matchesType && matchesSearch;
    });
  }, [envResources, search, typeFilter]);

  const environmentCreatorId = typeof env.properties.createdBy === 'string'
    ? env.properties.createdBy
    : env.properties.createdBy?.id;
  const ownerGuids = useMemo(
    () => [
      environmentCreatorId,
      ...envResources.map((resource) => {
        const owner = resource.properties.createdBy ?? resource.properties.ownerId;
        return typeof owner === 'string' ? owner : owner?.id;
      }),
    ],
    [env, envResources],
  );
  const ownerNames = useOwners(ownerGuids, env.name);

  return (
    <div className={styles.root}>
      {/* Hero */}
      <div className={styles.hero}>
        {/* Breadcrumb */}
        <div className={styles.breadcrumb}>
          <Button
            appearance="subtle"
            size="small"
            onClick={onBack}
            style={{ color: tokens.colorBrandForeground1, padding: `0 ${tokens.spacingHorizontalXS}`, minWidth: 0 }}
          >Environments</Button>
          <ChevronRightRegular style={{ fontSize: '0.7rem' }} />
          <Text style={{ fontSize: tokens.fontSizeBase200 }}>{displayName}</Text>
        </div>

        {/* Title row */}
        <div className={styles.heroBody}>
          <GlobeRegular className={styles.heroIcon} />
          <Text className={styles.heroName}>{displayName}</Text>
          <div className={styles.heroBadges}>
            <Badge appearance="filled" color={envTypeColor(envType)}>{envType}</Badge>
            {isDisabled && <Badge appearance="filled" color="danger" icon={<StopRegular />}>Disabled</Badge>}
            {isManaged && <Badge appearance="tint" color="success">Managed</Badge>}
            {envGroupId && (
              <Badge appearance="outline" color="informative" icon={<LayerRegular />}>
                {groupMap.get(envGroupId) ?? 'In Group'}
              </Badge>
            )}
            {isPending && <Badge appearance="tint" color="informative">Working…</Badge>}
          </div>
          <Menu>
            <MenuTrigger>
              <Button appearance="outline" size="small" disabled={isPending} icon={<ChevronDownRegular />} iconPosition="after">Actions</Button>
            </MenuTrigger>
            <MenuPopover>
              <MenuList>
                {isDisabled && <MenuItem icon={<PlayRegular />} onClick={() => void runAction(() => execEnable(env.name))}>Enable</MenuItem>}
                {isEnabled && <MenuItem icon={<StopRegular />} onClick={() => setConfirmAction({ type: 'disable' })}>Disable</MenuItem>}
                {isManaged
                  ? <MenuItem icon={<ShieldDismissRegular />} onClick={() => setConfirmAction({ type: 'disableManaged' })}>Disable Managed</MenuItem>
                  : <MenuItem icon={<ShieldRegular />} onClick={() => void runAction(() => execEnableManaged(env.name))}>Enable Managed</MenuItem>
                }
                <MenuItem icon={<SaveRegular />} onClick={() => setShowBackup(true)}>Create Backup</MenuItem>
                <AddSelfAsAdminBanner
                  variant="menu"
                  environmentId={env.name}
                  onChanged={() => void onRefreshEnvironments?.()}
                />
                {!envGroupId && (
                  <MenuItem icon={<LayerRegular />} onClick={() => setGroupDialogMode('add')}>Add to Group</MenuItem>
                )}
                {envGroupId && (
                  <MenuItem icon={<SubtractCircleRegular />} onClick={() => setGroupDialogMode('remove')}>Remove from Group</MenuItem>
                )}
              </MenuList>
            </MenuPopover>
          </Menu>
        </div>

        {/* Sub-line meta */}
        <div className={styles.heroMeta}>
          <span className={styles.heroMetaItem}><GlobeRegular style={{ fontSize: '0.85rem' }} />{region}</span>
          <span className={styles.heroMetaItem}><CalendarRegular style={{ fontSize: '0.85rem' }} />Created {createdAt}</span>
          <span className={styles.heroMetaItem} title={env.name}><KeyRegular style={{ fontSize: '0.85rem' }} />ID: {env.name}</span>
          {environmentCreatorId && (
            <span className={styles.heroMetaItem}>
              <PersonRegular style={{ fontSize: '0.85rem' }} />
              {ownerNames.get(environmentCreatorId.toLowerCase()) ?? environmentCreatorId}
            </span>
          )}
        </div>
      </div>

      <div className={styles.contentTabs}>
        <TabList selectedValue={contentTab} onTabSelect={(_, d) => setContentTab(d.value as 'resources' | 'analysis')}>
          <Tab value="resources">Resources</Tab>
          <Tab value="analysis" icon={<ShieldCheckmarkRegular />}>Analysis</Tab>
        </TabList>
      </div>

      <div className={styles.content}>
        {contentTab === 'analysis' ? (
          <EnvironmentBpaSection
            envType={envType}
            isManaged={isManaged}
            envGroupId={envGroupId}
            resourceCount={envResources.length}
            domainName={domainName}
          />
        ) : (
          <>
            {typeCounts.size > 0 && (
              <div className={styles.statsRow}>
                {Array.from(typeCounts.entries())
                  .sort((left, right) => right[1] - left[1])
                  .map(([type, count]) => (
                    <button
                      key={type}
                      type="button"
                      className={styles.statPill}
                      aria-pressed={typeFilter === type}
                      onClick={() => setTypeFilter((current) => current === type ? 'all' : type)}
                    >
                      {resourceTypeIcon(type)}
                      <Text className={styles.statCount}>{count}</Text>
                      <Text className={styles.statLabel}>{RESOURCE_TYPE_LABELS[type] ?? type}</Text>
                    </button>
                  ))}
              </div>
            )}

            <div className={styles.tableSection}>
              <div className={styles.tableHeader}>
                <div>
                  <Text className={styles.tableTitle}>Resource directory</Text>
                  <Text className={styles.resourceSecondary}>
                    {filteredResources.length} of {envResources.length} resources in scope
                  </Text>
                </div>
                <div className={styles.tableControls}>
                  <Dropdown
                    value={typeFilter === 'all' ? 'All resource types' : RESOURCE_TYPE_LABELS[typeFilter] ?? typeFilter}
                    selectedOptions={[typeFilter]}
                    onOptionSelect={(_, data) => setTypeFilter(data.optionValue ?? 'all')}
                    size="small"
                  >
                    <Option value="all">All resource types</Option>
                    {Array.from(typeCounts.keys())
                      .sort((left, right) => (
                        (RESOURCE_TYPE_LABELS[left] ?? left).localeCompare(RESOURCE_TYPE_LABELS[right] ?? right)
                      ))
                      .map((type) => (
                        <Option key={type} value={type}>{RESOURCE_TYPE_LABELS[type] ?? type}</Option>
                      ))}
                  </Dropdown>
                  <Input
                    placeholder="Search this environment"
                    value={search}
                    onChange={(_, data) => setSearch(data.value)}
                    contentBefore={<SearchRegular />}
                    size="small"
                    style={{ minWidth: '240px' }}
                  />
                </div>
              </div>
              <div className={styles.tableWrapper}>
                <div className={styles.resourceHeader}>
                  <span>Resource</span>
                  <span>Type</span>
                  <span>Created</span>
                  <span>Modified</span>
                  <span>Owner</span>
                  <span>Open</span>
                </div>
                {filteredResources.length === 0 ? (
                  <EmptyState
                    icon={<DocumentSearchRegular />}
                    title="No resources found"
                    subtitle={search || typeFilter !== 'all'
                      ? 'Clear the active filters to restore the environment directory.'
                      : 'This environment has no inventoried resources.'}
                    action={search || typeFilter !== 'all'
                      ? {
                        label: 'Clear filters',
                        onClick: () => {
                          setSearch('');
                          setTypeFilter('all');
                        },
                      }
                      : undefined}
                  />
                ) : filteredResources.map((resource, index) => {
                  const name = resource.properties.displayName ?? resource.name;
                  const typeLabel = RESOURCE_TYPE_LABELS[resource.type] ?? resource.type;
                  const created = formatDate(resource.properties.createdAt as string | undefined);
                  const modified = formatDate((
                    resource.properties.lastModifiedAt
                    ?? resource.properties.modifiedAt
                    ?? resource.properties.lastPublishedAt
                  ) as string | undefined);
                  const rawOwner = resource.properties.createdBy ?? resource.properties.ownerId;
                  const ownerGuid = typeof rawOwner === 'string' ? rawOwner : rawOwner?.id ?? '';
                  const owner = ownerNames.get(ownerGuid.toLowerCase()) ?? (ownerGuid || '—');
                  const canOpen = Boolean(
                    onResourceSelect && DETAIL_PANEL_TYPES.has(resource.type.toLowerCase()),
                  );
                  return (
                    <div className={styles.resourceRow} key={resource.id ?? `${resource.name}-${index}`}>
                      <div className={styles.resourceIdentity}>
                        <span className={styles.resourceIcon}>{resourceTypeIcon(resource.type)}</span>
                        <Text className={styles.resourcePrimary} title={name}>{name}</Text>
                      </div>
                      <Text className={styles.resourceSecondary} title={typeLabel}>{typeLabel}</Text>
                      <Text className={styles.resourceSecondary}>{created}</Text>
                      <Text className={styles.resourceSecondary}>{modified}</Text>
                      <Text className={styles.resourceSecondary} title={owner}>{owner}</Text>
                      <Button
                        appearance="subtle"
                        size="small"
                        icon={<OpenRegular />}
                        disabled={!canOpen}
                        onClick={() => onResourceSelect?.(resource)}
                        title={canOpen ? `Open ${name} details` : 'Detail view unavailable'}
                        aria-label={canOpen ? `Open ${name} details` : `${name} details unavailable`}
                      />
                    </div>
                  );
                })}
              </div>
            </div>
            </>
          )}
      </div>

      <ConfirmDialog
        open={confirmAction?.type === 'disable'}
        title="Disable Environment"
        message={`Disable "${displayName}"? Users will lose access until re-enabled.`}
        confirmLabel="Disable"
        isDangerous
        onConfirm={() => { setConfirmAction(null); void runAction(() => execDisable(env.name)); }}
        onCancel={() => setConfirmAction(null)}
      />
      <ConfirmDialog
        open={confirmAction?.type === 'disableManaged'}
        title="Disable Managed Environment"
        message={`Remove managed environment features from "${displayName}"?`}
        confirmLabel="Disable Managed"
        isDangerous
        onConfirm={() => { setConfirmAction(null); void runAction(() => execDisableManaged(env.name)); }}
        onCancel={() => setConfirmAction(null)}
      />
      <BackupDialog
        open={showBackup}
        environmentName={displayName}
        isLoading={isBackupLoading}
        onConfirm={(notes) => void execBackup(env.name, notes)}
        onCancel={() => setShowBackup(false)}
      />
      <EnvironmentGroupDialog
        open={groupDialogMode !== null}
        mode={groupDialogMode ?? 'add'}
        environmentName={displayName}
        environmentId={env.name}
        preselectedGroupId={groupDialogMode === 'remove' ? envGroupId : undefined}
        onClose={() => setGroupDialogMode(null)}
        onSuccess={() => void onRefreshEnvironments?.()}
      />
    </div>
  );
}