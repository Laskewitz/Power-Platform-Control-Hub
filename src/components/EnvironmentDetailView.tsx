import { useEffect, useMemo, useState } from 'react';
import type { ReactElement } from 'react';
import {
  makeStyles,
  shorthands,
  tokens,
  Text,
  Badge,
  Button,
  Table,
  TableHeader,
  TableHeaderCell,
  TableBody,
  TableRow,
  TableCell,
  Input,
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
  BoxRegular,
  AppsFilled,
  FlowRegular,
  BotRegular,
  DatabaseRegular,
  LayerRegular,
  SubtractCircleRegular,
  ShieldCheckmarkRegular,
  ErrorCircleRegular,
  WarningRegular,
  InfoRegular,
  CheckmarkCircleRegular,
  OpenRegular,
} from '@fluentui/react-icons';
import type { Resource } from '../types/inventory.ts';
import type { EnvironmentGroup } from '../types/admin.ts';
import type { AnalysisResult } from '../services/flowAnalyzer.ts';
import { RESOURCE_TYPE_LABELS } from '../types/inventory.ts';

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
    height: '100%',
    overflow: 'hidden',
    backgroundColor: tokens.colorNeutralBackground2,
  },

  // ── Hero ──────────────────────────────────────────────────────────────────
  hero: {
    backgroundColor: tokens.colorNeutralBackground1,
    borderBottomWidth: '1px',
    borderBottomStyle: 'solid',
    borderBottomColor: tokens.colorNeutralStroke2,
    padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalXL}`,
    flexShrink: 0,
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
  },
  breadcrumb: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
  breadcrumbLink: {
    color: tokens.colorBrandForeground1,
    cursor: 'pointer',
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase200,
    ':hover': { textDecoration: 'underline' },
  },
  heroBody: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    flexWrap: 'wrap',
  },
  heroIcon: {
    fontSize: '1.75rem',
    color: tokens.colorBrandForeground1,
    flexShrink: 0,
  },
  heroName: {
    fontSize: tokens.fontSizeBase600,
    fontWeight: tokens.fontWeightBold,
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
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalL,
    flexWrap: 'wrap',
    paddingLeft: `calc(1.75rem + ${tokens.spacingHorizontalM})`,
  },
  heroMetaItem: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
  },

  // ── Content ───────────────────────────────────────────────────────────────
  content: {
    display: 'flex',
    flexDirection: 'column',
    flex: 1,
    minHeight: 0,
    overflow: 'hidden',
    padding: `${tokens.spacingVerticalL} ${tokens.spacingHorizontalXL}`,
    gap: tokens.spacingVerticalL,
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
    display: 'flex',
    gap: tokens.spacingHorizontalM,
    flexWrap: 'wrap',
    flexShrink: 0,
  },
  statPill: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
    borderRadius: tokens.borderRadiusCircular,
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalM}`,
  },
  statCount: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightBold,
    color: tokens.colorBrandForeground1,
  },
  statLabel: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },

  // ── Resources table ───────────────────────────────────────────────────────
  tableSection: {
    display: 'flex',
    flexDirection: 'column',
    flex: 1,
    overflow: 'hidden',
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: tokens.borderRadiusMedium,
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
  },
  tableHeader: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalL}`,
    borderBottomWidth: '1px',
    borderBottomStyle: 'solid',
    borderBottomColor: tokens.colorNeutralStroke2,
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
  },
  tableWrapper: {
    overflowY: 'auto',
    flex: 1,
  },

  // ── Content tabs ──────────────────────────────────────────────────────────
  contentTabs: {
    backgroundColor: tokens.colorNeutralBackground1,
    borderBottomWidth: '1px',
    borderBottomStyle: 'solid',
    borderBottomColor: tokens.colorNeutralStroke2,
    paddingLeft: tokens.spacingHorizontalXL,
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
    padding: `${tokens.spacingVerticalL} ${tokens.spacingHorizontalXL}`,
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
  },
  analysisRow: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground1,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    cursor: 'pointer',
    ':hover': { backgroundColor: tokens.colorNeutralBackground2 },
  },
  analysisRowDetail: {
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM} ${tokens.spacingVerticalM}`,
    backgroundColor: tokens.colorNeutralBackground1,
    borderLeft: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRight: `1px solid ${tokens.colorNeutralStroke2}`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: `0 0 ${tokens.borderRadiusMedium} ${tokens.borderRadiusMedium}`,
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
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

type BpaStyles = ReturnType<typeof useStyles>;

const SEV_ICON: Record<string, ReactElement> = {
  critical: <ErrorCircleRegular style={{ color: tokens.colorStatusDangerForeground1, flexShrink: 0 }} />,
  warning: <WarningRegular style={{ color: tokens.colorStatusWarningForeground1, flexShrink: 0 }} />,
  info: <InfoRegular style={{ color: tokens.colorStatusSuccessForeground1, flexShrink: 0 }} />,
};
const SEV_COLOR: Record<string, 'danger' | 'warning' | 'success'> = {
  critical: 'danger', warning: 'warning', info: 'success',
};
const SEV_LABEL: Record<string, string> = {
  critical: 'Critical', warning: 'Warning', info: 'Info',
};

function EnvironmentBpaSection({
  envType, isManaged, envGroupId, resourceCount, domainName, styles,
}: {
  envType: string;
  isManaged: boolean;
  envGroupId: string | undefined;
  resourceCount: number;
  domainName: string | undefined;
  styles: BpaStyles;
}): ReactElement {
  const results = analyzeEnvironment(envType, isManaged, envGroupId, resourceCount, domainName);
  const [expanded, setExpanded] = useState<Set<string>>(new Set());
  const toggle = (id: string) => setExpanded(prev => {
    const next = new Set(prev);
    if (next.has(id)) next.delete(id); else next.add(id);
    return next;
  });

  if (results.length === 0) {
    return (
      <div className={styles.analysisScroll}>
        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: tokens.spacingVerticalM, padding: tokens.spacingVerticalXL }}>
          <CheckmarkCircleRegular fontSize={40} style={{ color: tokens.colorStatusSuccessForeground1 }} />
          <Text style={{ fontWeight: tokens.fontWeightSemibold }}>No issues found</Text>
          <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>
            This environment follows governance best practices.
          </Text>
        </div>
      </div>
    );
  }

  return (
    <div className={styles.analysisScroll}>
      {results.map(r => {
        const isOpen = expanded.has(r.id);
        const borderColor =
          r.severity === 'critical' ? tokens.colorStatusDangerForeground1
          : r.severity === 'warning' ? tokens.colorStatusWarningForeground1
          : tokens.colorStatusSuccessForeground1;
        return (
          <div key={r.id}>
            <div
              className={styles.analysisRow}
              style={{ borderLeft: `3px solid ${borderColor}`, borderRadius: isOpen ? `${tokens.borderRadiusMedium} ${tokens.borderRadiusMedium} 0 0` : undefined }}
              onClick={() => toggle(r.id)}
              role="button"
              tabIndex={0}
              aria-expanded={isOpen}
              onKeyDown={e => (e.key === 'Enter' || e.key === ' ') && toggle(r.id)}
            >
              {SEV_ICON[r.severity]}
              <Text style={{ flex: 1, fontSize: tokens.fontSizeBase300, fontWeight: tokens.fontWeightSemibold }}>{r.title}</Text>
              <Badge appearance="filled" color={SEV_COLOR[r.severity]} size="small">{SEV_LABEL[r.severity]}</Badge>
              <span style={{ fontSize: '12px', color: tokens.colorNeutralForeground3, transform: isOpen ? 'rotate(90deg)' : undefined, display: 'inline-flex' }}>▶</span>
            </div>
            {isOpen && (
              <div className={styles.analysisRowDetail} style={{ borderLeft: `3px solid ${borderColor}` }}>
                <Text style={{ fontSize: tokens.fontSizeBase300, color: tokens.colorNeutralForeground2 }}>{r.description}</Text>
                <div style={{ backgroundColor: tokens.colorNeutralBackground3, borderRadius: tokens.borderRadiusMedium, padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalS}` }}>
                  <Text style={{ fontSize: tokens.fontSizeBase200, fontWeight: tokens.fontWeightSemibold, color: tokens.colorNeutralForeground2 }}>💡 Recommendation</Text>
                  <Text style={{ fontSize: tokens.fontSizeBase300, display: 'block', marginTop: '4px' }}>{r.recommendation}</Text>
                </div>
              </div>
            )}
          </div>
        );
      })}
    </div>
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
    if (!term) return envResources;
    return envResources.filter((r) =>
      (r.properties.displayName ?? r.name).toLowerCase().includes(term) ||
      (RESOURCE_TYPE_LABELS[r.type] ?? r.type).toLowerCase().includes(term),
    );
  }, [envResources, search]);

  const ownerGuids = useMemo(
    () => [
      env.properties.createdBy as string | undefined,
      ...envResources.map((r) => (r.properties.createdBy ?? r.properties.ownerId) as string | undefined),
    ],
    [env, envResources],
  );
  const ownerNames = useOwners(ownerGuids);

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
              <Button appearance="outline" size="small" disabled={isPending}>Actions ▾</Button>
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
          {env.properties.createdBy && (
            <span className={styles.heroMetaItem}>
              <PersonRegular style={{ fontSize: '0.85rem' }} />
              {ownerNames.get(String(env.properties.createdBy).toLowerCase()) ?? String(env.properties.createdBy)}
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
            styles={styles}
          />
        ) : (
          <>
            {typeCounts.size > 0 && (
          <div className={styles.statsRow}>
            {Array.from(typeCounts.entries()).map(([type, count]) => (
              <div key={type} className={styles.statPill}>
                {resourceTypeIcon(type)}
                <Text className={styles.statCount}>{count}</Text>
                <Text className={styles.statLabel}>{RESOURCE_TYPE_LABELS[type] ?? type}</Text>
              </div>
            ))}
          </div>
        )}

        {/* Resources table */}
        <div className={styles.tableSection}>
          <div className={styles.tableHeader}>
            <Text className={styles.tableTitle}>Resources</Text>
            <div className={styles.tableControls}>
              <Input
                placeholder="Search resources…"
                value={search}
                onChange={(_, data) => setSearch(data.value)}
                contentBefore={<SearchRegular />}
                size="small"
                style={{ minWidth: '220px' }}
              />
              <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3, whiteSpace: 'nowrap' }}>
                {filteredResources.length} of {envResources.length}
              </Text>
            </div>
          </div>
          <div className={styles.tableWrapper}>
            <Table size="small">
              <TableHeader>
                <TableRow>
                  <TableHeaderCell>Name</TableHeaderCell>
                  <TableHeaderCell style={{ width: '200px', minWidth: '200px', whiteSpace: 'nowrap' }}>Type</TableHeaderCell>
                  <TableHeaderCell style={{ width: '100px' }}>Created</TableHeaderCell>
                  <TableHeaderCell style={{ width: '100px' }}>Modified</TableHeaderCell>
                  <TableHeaderCell>Owner</TableHeaderCell>
                </TableRow>
              </TableHeader>
              <TableBody>
                {filteredResources.length === 0 ? (
                  <TableRow>
                    <TableCell colSpan={5}>
                      <Text style={{ color: tokens.colorNeutralForeground3 }}>No resources found.</Text>
                    </TableCell>
                  </TableRow>
                ) : (
                  filteredResources.map((r, i) => {
                    const name = r.properties.displayName ?? r.name;
                    const typeLabel = RESOURCE_TYPE_LABELS[r.type] ?? r.type;
                    const created = formatDate(r.properties.createdAt as string | undefined);
                    const modified = formatDate((r.properties.lastModifiedAt ?? r.properties.modifiedAt ?? r.properties.lastPublishedAt) as string | undefined);
                    const ownerGuid = (r.properties.createdBy ?? r.properties.ownerId ?? '') as string;
                    const owner = ownerNames.get(ownerGuid.toLowerCase()) ?? (ownerGuid || '—');
                    return (
                      <TableRow key={r.id ?? `${r.name}-${i}`}>
                        <TableCell>
                          <span style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalXS }}>
                            <Text style={{ fontWeight: tokens.fontWeightSemibold }}>{name}</Text>
                            {onResourceSelect && DETAIL_PANEL_TYPES.has(r.type.toLowerCase()) && (
                              <Button
                                appearance="subtle"
                                size="small"
                                icon={<OpenRegular />}
                                onClick={() => onResourceSelect(r)}
                                style={{ minWidth: 0, padding: `0 ${tokens.spacingHorizontalXS}` }}
                                title={`Open ${name} details`}
                              />
                            )}
                          </span>
                        </TableCell>
                        <TableCell>
                          <span style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalXS, whiteSpace: 'nowrap' }}>
                            {resourceTypeIcon(r.type)}
                            <Text>{typeLabel}</Text>
                          </span>
                        </TableCell>
                        <TableCell>{created}</TableCell>
                        <TableCell>{modified}</TableCell>
                        <TableCell>
                          <Text style={{ color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 }}>
                            {owner}
                          </Text>
                        </TableCell>
                      </TableRow>
                    );
                  })
                )}
              </TableBody>
            </Table>
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