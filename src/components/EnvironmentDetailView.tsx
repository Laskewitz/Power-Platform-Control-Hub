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
  Switch,
  Spinner,
  MessageBar,
  MessageBarBody,
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
  SettingsRegular,
  PersonAddRegular,
} from '@fluentui/react-icons';
import type { Resource } from '../types/inventory.ts';
import type { EnvironmentGroup } from '../types/admin.ts';
import type { EnvironmentManagementSetting } from '../generated/models/PowerPlatformforAdminsV2Model.ts';
import { RESOURCE_TYPE_LABELS } from '../types/inventory.ts';
import ConfirmDialog from './ConfirmDialog.tsx';
import { useMutation } from '../hooks/useMutation.tsx';
import {
  enableEnvironment,
  disableEnvironment,
  enableManagedEnvironment,
  disableManagedEnvironment,
  createEnvironmentBackup,
  addSelfAsEnvironmentAdmin,
} from '../services/environmentMutations.ts';
import { fetchEnvironmentSettings, updateEnvironmentSettings } from '../services/settingsService.ts';
import BackupDialog from './BackupDialog.tsx';
import EnvironmentGroupDialog from './EnvironmentGroupDialog.tsx';
import { useOwners } from '../services/ownerCache.ts';

interface EnvironmentDetailViewProps {
  environment: Resource;
  resources: Resource[];
  envGroups?: EnvironmentGroup[];
  onBack: () => void;
  onRefreshEnvironments?: () => Promise<void>;
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
    overflow: 'hidden',
    padding: `${tokens.spacingVerticalL} ${tokens.spacingHorizontalXL}`,
    gap: tokens.spacingVerticalL,
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
  settingsScroll: {
    flex: 1,
    overflowY: 'auto',
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXL,
    padding: `${tokens.spacingVerticalL} ${tokens.spacingHorizontalXL}`,
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
  },
});

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

type SettingFieldDef = { key: keyof EnvironmentManagementSetting; label: string; type: 'boolean' | 'string' };
type SettingGroupDef = { label: string; fields: SettingFieldDef[] };

const SETTING_GROUPS: SettingGroupDef[] = [
  {
    label: 'Power Apps',
    fields: [
      { key: 'powerApps_CopilotChat', label: 'Copilot Chat', type: 'boolean' },
      { key: 'powerApps_NLSearch', label: 'Natural Language Search', type: 'boolean' },
      { key: 'powerApps_AllowCodeApps', label: 'Code Apps', type: 'boolean' },
      { key: 'powerApps_EnableFormInsights', label: 'Form Insights', type: 'boolean' },
      { key: 'powerApps_ChartVisualization', label: 'Chart Visualization AI', type: 'boolean' },
      { key: 'powerApps_FormPredictSmartPaste', label: 'Smart Paste', type: 'boolean' },
      { key: 'powerApps_FormPredictAutomatic', label: 'Automatic Form Predictions', type: 'boolean' },
    ],
  },
  {
    label: 'Copilot Studio',
    fields: [
      { key: 'copilotStudio_ConnectedAgents', label: 'Connected Agents', type: 'boolean' },
      { key: 'copilotStudio_CodeInterpreter', label: 'Code Interpreter', type: 'boolean' },
      { key: 'copilotStudio_ConversationAuditLoggingEnabled', label: 'Conversation Audit Logging', type: 'boolean' },
      { key: 'copilotStudio_ComputerUseSharedMachines', label: 'Computer Use – Shared Machines', type: 'boolean' },
      { key: 'copilotStudio_ComputerUseCredentialsAllowed', label: 'Computer Use – Credentials Allowed', type: 'boolean' },
      { key: 'copilotStudio_ComputerUseAppAllowlist', label: 'Computer Use – App Allowlist', type: 'string' },
      { key: 'copilotStudio_ComputerUseWebAllowlist', label: 'Computer Use – Web Allowlist', type: 'string' },
    ],
  },
  {
    label: 'Power Pages',
    fields: [
      { key: 'powerPages_AllowMakerCopilotsForNewSites', label: 'Maker Copilots – New Sites', type: 'string' },
      { key: 'powerPages_AllowMakerCopilotsForExistingSites', label: 'Maker Copilots – Existing Sites', type: 'string' },
      { key: 'powerPages_AllowProDevCopilotsForSites', label: 'ProDev Copilots for Sites', type: 'string' },
      { key: 'powerPages_AllowSiteCopilotForSites', label: 'Site Copilot', type: 'string' },
      { key: 'powerPages_AllowNonProdPublicSites', label: 'Non-Prod Public Sites', type: 'string' },
      { key: 'powerPages_AllowProDevCopilotsForEnvironment', label: 'ProDev Copilots for Environment', type: 'string' },
    ],
  },
  {
    label: 'Dynamics 365',
    fields: [
      { key: 'd365CustomerService_Copilot', label: 'Customer Service Copilot', type: 'boolean' },
      { key: 'd365CustomerService_AIAgents', label: 'AI Agents', type: 'boolean' },
    ],
  },
  {
    label: 'Security',
    fields: [
      { key: 'enableIpBasedStorageAccessSignatureRule', label: 'IP-Based Storage Access Signature', type: 'boolean' },
      { key: 'loggingEnabledForIpBasedStorageAccessSignature', label: 'Logging for IP-Based Storage Access', type: 'boolean' },
      { key: 'allowedIpRangeForStorageAccessSignatures', label: 'Allowed IP Range for Storage', type: 'string' },
    ],
  },
];

export default function EnvironmentDetailView({
  environment: env,
  resources,
  envGroups = [],
  onBack,
  onRefreshEnvironments,
}: EnvironmentDetailViewProps): ReactElement {
  const styles = useStyles();
  const [search, setSearch] = useState('');
  const [confirmAction, setConfirmAction] = useState<{ type: 'disable' | 'disableManaged' } | null>(null);
  const [showBackup, setShowBackup] = useState(false);
  const [groupDialogMode, setGroupDialogMode] = useState<'add' | 'remove' | null>(null);
  const [isPending, setIsPending] = useState(false);
  const [contentTab, setContentTab] = useState<'resources' | 'settings'>('resources');

  // Settings state
  const [settings, setSettings] = useState<EnvironmentManagementSetting | null>(null);
  const [pendingSettings, setPendingSettings] = useState<EnvironmentManagementSetting>({});
  const [settingsLoading, setSettingsLoading] = useState(false);
  const [settingsError, setSettingsError] = useState<string | null>(null);
  const [isSavingSettings, setIsSavingSettings] = useState(false);

  // Fetch settings lazily when the Settings tab is first opened
  useEffect(() => {
    if (contentTab !== 'settings' || settings !== null || settingsLoading) return;
    setSettingsLoading(true);
    setSettingsError(null);
    fetchEnvironmentSettings(env.name)
      .then((s) => { setSettings(s ?? {}); setPendingSettings(s ?? {}); })
      .catch((e: unknown) => setSettingsError(String(e)))
      .finally(() => setSettingsLoading(false));
  }, [contentTab, env.name, settings, settingsLoading]);

  const displayName = env.properties.displayName ?? env.name;
  const envType = (env.properties.environmentType ?? 'Unknown') as string;
  const isManaged = env.properties.isManaged === true;
  const region = env.location ?? '—';
  const createdAt = env.properties.createdAt
    ? new Date(env.properties.createdAt as string).toLocaleDateString()
    : '—';
  const shortId = env.name.length > 20 ? `${env.name.slice(0, 8)}…${env.name.slice(-8)}` : env.name;
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
  const { execute: execAddSelfAsAdmin, isLoading: isAddingAsAdmin } = useMutation(addSelfAsEnvironmentAdmin, {
    successMessage: 'You have been added as System Administrator. Refreshing…',
    onSuccess: () => void onRefreshEnvironments?.(),
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
          <Text
            className={styles.breadcrumbLink}
            onClick={onBack}
            role="button"
            tabIndex={0}
            aria-label="Back to Environments list"
            onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); onBack(); } }}
          >Environments</Text>
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
                <MenuItem
                  icon={isAddingAsAdmin ? <Spinner size="tiny" /> : <PersonAddRegular />}
                  disabled={isAddingAsAdmin || isPending}
                  onClick={() => void execAddSelfAsAdmin(env.name)}
                >
                  Add yourself as Environment Admin
                </MenuItem>
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
          <span className={styles.heroMetaItem}><KeyRegular style={{ fontSize: '0.85rem' }} title={env.name}>{shortId}</KeyRegular></span>
          {env.properties.createdBy && (
            <span className={styles.heroMetaItem}>
              <PersonRegular style={{ fontSize: '0.85rem' }} />
              {ownerNames.get(String(env.properties.createdBy).toLowerCase()) ?? String(env.properties.createdBy)}
            </span>
          )}
        </div>
      </div>

      <div className={styles.contentTabs}>
        <TabList selectedValue={contentTab} onTabSelect={(_, d) => setContentTab(d.value as 'resources' | 'settings')}>
          <Tab value="resources">Resources</Tab>
          <Tab value="settings" icon={<SettingsRegular />}>Settings</Tab>
        </TabList>
      </div>

      <div className={styles.content}>
        {contentTab === 'settings' ? (
          settingsLoading ? (
            <div style={{ display: 'flex', justifyContent: 'center', padding: tokens.spacingVerticalXL }}>
              <Spinner label="Loading settings…" />
            </div>
          ) : settingsError ? (
            <MessageBar intent="error"><MessageBarBody>{settingsError}</MessageBarBody></MessageBar>
          ) : (
            <div className={styles.settingsScroll}>
              {SETTING_GROUPS.map((group) => (
                <div key={group.label} className={styles.settingsGroup}>
                  <div className={styles.settingsGroupHeader}>{group.label}</div>
                  {group.fields.map((field) => {
                    const current = pendingSettings[field.key];
                    return (
                      <div key={field.key} className={styles.settingRow}>
                        <Text className={styles.settingLabel}>{field.label}</Text>
                        {field.type === 'boolean' ? (
                          <Switch
                            checked={Boolean(current)}
                            onChange={(_, d) => setPendingSettings((prev) => ({ ...prev, [field.key]: d.checked }))}
                          />
                        ) : (
                          <Input
                            size="small"
                            style={{ minWidth: '260px' }}
                            value={String(current ?? '')}
                            placeholder="Not set"
                            onChange={(_, d) => setPendingSettings((prev) => ({ ...prev, [field.key]: d.value }))}
                          />
                        )}
                      </div>
                    );
                  })}
                </div>
              ))}
              <div className={styles.settingsActions}>
                <Button
                  appearance="secondary"
                  disabled={isSavingSettings}
                  onClick={() => setPendingSettings(settings ?? {})}
                >
                  Reset
                </Button>
                <Button
                  appearance="primary"
                  icon={isSavingSettings ? <Spinner size="tiny" /> : <SaveRegular />}
                  disabled={isSavingSettings}
                  onClick={async () => {
                    setIsSavingSettings(true);
                    try {
                      await updateEnvironmentSettings(env.name, pendingSettings);
                      setSettings(pendingSettings);
                    } catch (e: unknown) {
                      setSettingsError(String(e));
                    } finally {
                      setIsSavingSettings(false);
                    }
                  }}
                >
                  {isSavingSettings ? 'Saving…' : 'Save Settings'}
                </Button>
              </div>
            </div>
          )
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
                  <TableHeaderCell style={{ width: '160px' }}>Type</TableHeaderCell>
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
                    const created = r.properties.createdAt ? new Date(r.properties.createdAt as string).toLocaleDateString() : '—';
                    const modified = r.properties.modifiedAt ? new Date(r.properties.modifiedAt as string).toLocaleDateString() : '—';
                    const ownerGuid = (r.properties.createdBy ?? r.properties.ownerId ?? '') as string;
                    const owner = ownerNames.get(ownerGuid.toLowerCase()) ?? (ownerGuid || '—');
                    return (
                      <TableRow key={r.id ?? `${r.name}-${i}`}>
                        <TableCell><Text style={{ fontWeight: tokens.fontWeightSemibold }}>{name}</Text></TableCell>
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