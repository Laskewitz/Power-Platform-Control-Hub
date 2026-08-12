import { useMemo, useState } from 'react';
import type { ReactElement } from 'react';
import {
  makeStyles,
  tokens,
  Text,
  Card,
  Badge,
  Input,
  MessageBar,
  MessageBarBody,
  Button,
  Menu,
  MenuTrigger,
  MenuList,
  MenuItem,
  MenuPopover,
  Dropdown,
  Option,
  Spinner,
} from '@fluentui/react-components';
import {
  SearchRegular,
  GlobeRegular,
  LockClosedRegular,
  AppsListRegular,
  BoxRegular,
  MoreHorizontalRegular,
  PlayRegular,
  StopRegular,
  ShieldRegular,
  ShieldDismissRegular,
  SaveRegular,
  LayerRegular,
  SubtractCircleRegular,
  FilterRegular,
  ArrowClockwiseRegular,
} from '@fluentui/react-icons';
import type { Resource } from '../types/inventory.ts';
import type { EnvironmentGroup } from '../types/admin.ts';
import ConfirmDialog from './ConfirmDialog.tsx';
import BackupDialog from './BackupDialog.tsx';
import EnvironmentDetailView from './EnvironmentDetailView.tsx';
import EnvironmentGroupDialog from './EnvironmentGroupDialog.tsx';
import CloudFlowDetailPanel from './CloudFlowDetailPanel.tsx';
import CanvasAppDetailPanel from './CanvasAppDetailPanel.tsx';
import CopilotStudioAgentDetailPanel from './CopilotStudioAgentDetailPanel.tsx';
import EmptyState from './EmptyState.tsx';
import { useMutation } from '../hooks/useMutation.tsx';
import { formatDate } from '../utils/formatDate.ts';
import { OperationsSkeleton, PageHeader } from './ui.tsx';
import {
  enableEnvironment,
  disableEnvironment,
  enableManagedEnvironment,
  disableManagedEnvironment,
  createEnvironmentBackup,
} from '../services/environmentMutations.ts';

interface EnvironmentsViewProps {
  environments: Resource[];
  resources: Resource[];
  envGroups: EnvironmentGroup[];
  isLoading: boolean;
  error: string | null;
  onRefreshEnvironments?: () => Promise<void>;
}

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    gap: '18px',
    padding: '28px 32px 32px',
    height: '100%',
    overflow: 'hidden',
    '@media (max-width: 768px)': {
      padding: tokens.spacingHorizontalM,
    },
    directoryRegion: {
      display: 'grid',
      gap: '2px',
      minWidth: 0,
      '@media (max-width: 980px)': {
        display: 'none',
      },
    },
    directoryMetric: {
      display: 'grid',
      justifyItems: 'end',
      gap: '2px',
    },
    directoryValue: {
      color: tokens.colorNeutralForeground1,
      fontSize: '22px',
      lineHeight: 1,
      fontWeight: tokens.fontWeightSemibold,
      fontVariantNumeric: 'tabular-nums',
    },
    directoryLabel: {
      color: tokens.colorNeutralForeground3,
      fontSize: '9px',
      fontWeight: tokens.fontWeightSemibold,
      letterSpacing: '0.08em',
      textTransform: 'uppercase',
    },
    actionCell: {
      display: 'grid',
      placeItems: 'center',
    },
  },
  controlRail: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    flexShrink: 0,
    flexWrap: 'wrap',
    padding: '0 0 16px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
    backgroundColor: 'transparent',
  },
  atlasSummary: {
    display: 'grid',
    gridTemplateColumns: 'repeat(3, minmax(120px, 1fr))',
    maxWidth: '860px',
    color: '#C9D8DE',
    border: '1px solid #29404F',
    backgroundColor: '#0C141D',
    boxShadow: '0 10px 24px rgba(0, 0, 0, 0.28)',
    '@media (max-width: 620px)': {
      gridTemplateColumns: '1fr',
    },
  },
  summaryCell: {
    display: 'grid',
    gridTemplateColumns: 'auto minmax(0, 1fr)',
    alignItems: 'baseline',
    gap: tokens.spacingHorizontalS,
    padding: '12px 16px',
    borderRight: '1px solid #29404F',
    ':first-child': {
      color: '#43D9FF',
      backgroundColor: '#0B1D26',
    },
    ':last-child': {
      color: '#FFB547',
      backgroundColor: '#1A1710',
      borderRight: 0,
    },
    '@media (max-width: 620px)': {
      borderRight: 0,
      borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    },
  },
  summaryValue: {
    color: 'inherit',
    fontSize: '24px',
    lineHeight: 1,
    fontWeight: tokens.fontWeightSemibold,
    fontVariantNumeric: 'tabular-nums',
  },
  summaryLabel: {
    color: 'inherit',
    opacity: 0.78,
    fontSize: tokens.fontSizeBase200,
  },
  count: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    whiteSpace: 'nowrap',
    marginLeft: 'auto',
  },
  grid: {
    display: 'flex',
    flexDirection: 'column',
    gap: 0,
    borderTop: '1px solid #29404F',
    borderBottom: '1px solid #29404F',
    overflowX: 'hidden',
    overflowY: 'auto',
    flex: 1,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  card: {
    position: 'relative',
    display: 'grid',
    gridTemplateColumns: 'minmax(300px, 1.6fr) minmax(210px, 0.9fr) 150px 180px',
    alignItems: 'center',
    padding: '14px 16px',
    gap: '18px',
    minWidth: 0,
    boxSizing: 'border-box',
    minHeight: '92px',
    border: 0,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: 0,
    backgroundColor: tokens.colorNeutralBackground1,
    boxShadow: 'none',
    cursor: 'pointer',
    transitionProperty: 'background-color, box-shadow, transform',
    transitionDuration: '180ms',
    transitionTimingFunction: 'cubic-bezier(.16, 1, .3, 1)',
    ':hover': {
      zIndex: 2,
      backgroundColor: tokens.colorBrandBackground2,
      boxShadow: 'inset 0 -2px 0 #43D9FF',
      transform: 'translateX(4px)',
    },
    '@media (max-width: 980px)': {
      gridTemplateColumns: 'minmax(240px, 1fr) minmax(180px, auto) 110px 150px',
    },
    '@media (max-width: 680px)': {
      gridTemplateColumns: 'minmax(0, 1fr)',
      gap: '10px',
    },
    ':focus-visible': {
      outline: `2px solid ${tokens.colorBrandStroke1}`,
      outlineOffset: '2px',
    },
  },
  coordinate: {
    color: tokens.colorNeutralForeground3,
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: '10px',
    letterSpacing: '0.08em',
    textTransform: 'uppercase',
  },
  cardTop: {
    display: 'flex',
    alignItems: 'flex-start',
    gap: tokens.spacingHorizontalS,
    minWidth: 0,
  },
  cardIcon: {
    fontSize: '1.15rem',
    color: tokens.colorBrandForeground1,
    flexShrink: 0,
    marginTop: '2px',
  },
  cardTitles: {
    display: 'flex',
    flexDirection: 'column',
    gap: '2px',
    flex: 1,
    minWidth: 0,
  },
  cardName: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase400,
    letterSpacing: '-0.015em',
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
  },
  cardRegion: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    textTransform: 'uppercase',
    letterSpacing: '0.06em',
  },
  badgeRow: {
    display: 'flex',
    gap: tokens.spacingHorizontalS,
    flexWrap: 'wrap',
    paddingTop: 0,
    '@media (max-width: 680px)': {
      gridColumn: '1 / -1',
    },
  },
  divider: {
    marginTop: tokens.spacingVerticalS,
    marginBottom: tokens.spacingVerticalXS,
  },
  disabledCard: {
    backgroundImage: 'repeating-linear-gradient(135deg, transparent, transparent 8px, rgba(122, 74, 0, 0.05) 8px, rgba(122, 74, 0, 0.05) 9px)',
  },
  metaRow: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },
  metaIcon: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
  },
  centered: {
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    height: '100%',
  },
});

function envTypeColor(
  envType: string,
): 'brand' | 'success' | 'warning' | 'important' | 'informative' {
  switch (envType.toLowerCase()) {
    case 'production': return 'brand';
    case 'sandbox': return 'warning';
    case 'developer': return 'success';
    case 'default': return 'informative';
    default: return 'important';
  }
}

export default function EnvironmentsView({
  environments,
  resources,
  envGroups,
  isLoading,
  error,
  onRefreshEnvironments,
}: EnvironmentsViewProps): ReactElement {
  const styles = useStyles();
  const [selectedEnv, setSelectedEnv] = useState<Resource | null>(null);
  const [detailResource, setDetailResource] = useState<Resource | null>(null);
  const [search, setSearch] = useState('');
  const [typeFilter, setTypeFilter] = useState('all');
  const [regionFilter, setRegionFilter] = useState('all');
  const [confirmAction, setConfirmAction] = useState<{
    type: 'disable' | 'disableManaged';
    env: Resource;
  } | null>(null);
  const [backupEnv, setBackupEnv] = useState<Resource | null>(null);
  const [pendingEnvId, setPendingEnvId] = useState<string | null>(null);
  const [groupDialog, setGroupDialog] = useState<{ mode: 'add' | 'remove'; env: Resource; groupId?: string } | null>(null);

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
    successMessage: 'Backup request submitted. It may take a few minutes to complete.',
    onSuccess: () => setBackupEnv(null),
  });

  async function runAction(envId: string, action: () => Promise<unknown>) {
    setPendingEnvId(envId);
    try {
      await action();
    } finally {
      setPendingEnvId(null);
    }
  }

  const groupMap = useMemo(
    () => new Map(envGroups.map((g) => [g.id, g.displayName])),
    [envGroups],
  );

  // Derive unique type and region values from data
  const typeOptions = useMemo(() => {
    const types = new Set(environments.map((e) => (e.properties.environmentType as string | undefined) ?? 'Unknown'));
    return ['all', ...Array.from(types).sort()];
  }, [environments]);

  const regionOptions = useMemo(() => {
    const regions = new Set(environments.map((e) => e.location ?? 'Unknown'));
    return ['all', ...Array.from(regions).sort()];
  }, [environments]);

  // Pre-compute resource counts per environment name (lower-cased for matching).
  const countByEnv = useMemo(() => {
    const map = new Map<string, number>();
    for (const r of resources) {
      const key = (r.environmentName ?? '').toLowerCase();
      map.set(key, (map.get(key) ?? 0) + 1);
    }
    return map;
  }, [resources]);

  const filtered = useMemo(() => {
    const term = search.toLowerCase();
    return environments.filter((e) => {
      const name = (e.properties.displayName ?? e.name).toLowerCase();
      const matchesSearch = !term || name.includes(term) || (e.location ?? '').toLowerCase().includes(term);
      const envType = (e.properties.environmentType as string | undefined) ?? 'Unknown';
      const matchesType = typeFilter === 'all' || envType === typeFilter;
      const matchesRegion = regionFilter === 'all' || (e.location ?? 'Unknown') === regionFilter;
      return matchesSearch && matchesType && matchesRegion;
    });
  }, [environments, search, typeFilter, regionFilter]);

  const environmentSummary = useMemo(() => ({
    managed: environments.filter((environment) => environment.properties.isManaged === true).length,
    disabled: environments.filter((environment) => {
      const runtime = ((environment.properties.states as Record<string, unknown> | undefined)
        ?.runtime as Record<string, unknown> | undefined)?.id;
      return typeof runtime === 'string' && runtime.toLowerCase() === 'disabled';
    }).length,
    resources: Array.from(countByEnv.values()).reduce((total, count) => total + count, 0),
  }), [countByEnv, environments]);

  if (isLoading) {
    return <OperationsSkeleton />;
  }

  if (selectedEnv) {
    // If a resource detail panel is open on top of the environment view
    if (detailResource) {
      const typeLower = detailResource.type.toLowerCase();
      if (typeLower.includes('powerapps') || typeLower.includes('canvasapps')) {
        return (
          <CanvasAppDetailPanel
            resource={detailResource}
            onClose={() => setDetailResource(null)}
          />
        );
      }
      if (typeLower.includes('cloudflows') || typeLower.includes('agentflows') || typeLower.includes('m365agent')) {
        return (
          <CloudFlowDetailPanel
            resource={detailResource}
            onClose={() => setDetailResource(null)}
            onDeleted={() => setDetailResource(null)}
          />
        );
      }
      if (typeLower.includes('agents')) {
        return (
          <CopilotStudioAgentDetailPanel
            resource={detailResource}
            onClose={() => setDetailResource(null)}
            onDeleted={() => setDetailResource(null)}
          />
        );
      }
    }
    return (
      <EnvironmentDetailView
        environment={selectedEnv}
        resources={resources}
        envGroups={envGroups}
        onBack={() => setSelectedEnv(null)}
        onRefreshEnvironments={onRefreshEnvironments}
        onResourceSelect={(r) => setDetailResource(r)}
      />
    );
  }

  return (
    <div className={styles.root}>
      <PageHeader
        title="Environments"
        description="Monitor capacity, lifecycle state, and governance posture for every environment."
        actions={
          onRefreshEnvironments && (
            <Button appearance="secondary" icon={<ArrowClockwiseRegular />} onClick={() => void onRefreshEnvironments()}>
              Refresh
            </Button>
          )
        }
      />

      <div className={styles.atlasSummary} aria-label="Environment scope summary">
        <div className={styles.summaryCell}>
          <Text className={styles.summaryValue}>{environmentSummary.managed}</Text>
          <Text className={styles.summaryLabel}>Managed environments</Text>
        </div>
        <div className={styles.summaryCell}>
          <Text className={styles.summaryValue}>{environmentSummary.disabled}</Text>
          <Text className={styles.summaryLabel}>Disabled environments</Text>
        </div>
        <div className={styles.summaryCell}>
          <Text className={styles.summaryValue}>{environmentSummary.resources}</Text>
          <Text className={styles.summaryLabel}>Indexed resources</Text>
        </div>
      </div>

      <div className={styles.controlRail}>
        <Input
          placeholder="Search environments…"
          value={search}
          onChange={(_, data) => setSearch(data.value)}
          contentBefore={<SearchRegular />}
          size="small"
          style={{ minWidth: '200px' }}
        />
        <Dropdown
          value={typeFilter === 'all' ? 'All Types' : typeFilter}
          selectedOptions={[typeFilter]}
          onOptionSelect={(_, data) => setTypeFilter(data.optionValue ?? 'all')}
          size="small"
          style={{ minWidth: '140px' }}
        >
          {typeOptions.map((t) => (
            <Option key={t} value={t}>{t === 'all' ? 'All Types' : t}</Option>
          ))}
        </Dropdown>
        <Dropdown
          value={regionFilter === 'all' ? 'All Regions' : regionFilter}
          selectedOptions={[regionFilter]}
          onOptionSelect={(_, data) => setRegionFilter(data.optionValue ?? 'all')}
          size="small"
          style={{ minWidth: '160px' }}
        >
          {regionOptions.map((r) => (
            <Option key={r} value={r}>{r === 'all' ? 'All Regions' : r}</Option>
          ))}
        </Dropdown>
        <Text className={styles.count}>{filtered.length} environment(s)</Text>
      </div>

      {error && (
        <MessageBar intent="error">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}

      <div className={styles.grid}>
        {filtered.length === 0 ? (
          <EmptyState
            icon={<FilterRegular />}
            title="No environments match your filters"
            subtitle="Try adjusting your search term, type, or region filter."
            action={search || typeFilter !== 'all' || regionFilter !== 'all'
              ? { label: 'Clear filters', onClick: () => { setSearch(''); setTypeFilter('all'); setRegionFilter('all'); } }
              : undefined}
          />
        ) : (
          filtered.map((e, i) => {
            const displayName = e.properties.displayName ?? e.name;
            const envType = (e.properties.environmentType ?? 'Unknown') as string;
            const isManaged = e.properties.isManaged === true;
            const region = e.location ?? '—';
            const resourceCount = countByEnv.get(displayName.toLowerCase()) ?? 0;
            const createdAt = formatDate(e.properties.createdAt as string | undefined);
            const isPending = pendingEnvId === e.name;
            const envGroupId = e.properties['environmentGroupId'] as string | undefined;
            const runtimeState = ((e.properties['states'] as Record<string, unknown> | undefined)
              ?.['runtime'] as Record<string, unknown> | undefined)
              ?.['id'] as string | undefined;
            const envIsDisabled = runtimeState?.toLowerCase() === 'disabled';
            const envIsEnabled = !runtimeState || runtimeState?.toLowerCase() === 'enabled';

            return (
              <Card
                key={e.id ?? `env-${e.name}-${i}`}
                className={`${styles.card} ${envIsDisabled ? styles.disabledCard : ''}`}
                onClick={() => setSelectedEnv(e)}
                tabIndex={0}
                role="button"
                aria-label={`Open details for ${displayName}`}
                onKeyDown={(ev) => { if (ev.target !== ev.currentTarget) return; if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); setSelectedEnv(e); } }}
              >
                <div className={styles.cardTop}>
                  <GlobeRegular className={styles.cardIcon} />
                  <div className={styles.cardTitles}>
                    <Text className={styles.coordinate}>ENV / {e.name.slice(-8)}</Text>
                    <Text className={styles.cardName} title={displayName}>{displayName}</Text>
                    <Text className={styles.cardRegion}>{region}</Text>
                  </div>
                  <Menu>
                    <MenuTrigger>
                      <Button
                        appearance="subtle"
                        size="small"
                        icon={<MoreHorizontalRegular />}
                        disabled={isPending}
                        title="Actions"
                        style={{ marginLeft: 'auto', flexShrink: 0 }}
                        onClick={(ev) => ev.stopPropagation()}
                      />
                    </MenuTrigger>
                    <MenuPopover>
                      <MenuList>
                        {envIsDisabled && (
                          <MenuItem
                            icon={<PlayRegular />}
                            onClick={(ev) => { ev.stopPropagation(); void runAction(e.name, () => execEnable(e.name)); }}
                          >
                            Enable
                          </MenuItem>
                        )}
                        {envIsEnabled && (
                          <MenuItem
                            icon={<StopRegular />}
                            onClick={(ev) => { ev.stopPropagation(); setConfirmAction({ type: 'disable', env: e }); }}
                          >
                            Disable
                          </MenuItem>
                        )}
                        {isManaged ? (
                          <MenuItem
                            icon={<ShieldDismissRegular />}
                            onClick={(ev) => { ev.stopPropagation(); setConfirmAction({ type: 'disableManaged', env: e }); }}
                          >
                            Disable Managed
                          </MenuItem>
                        ) : (
                          <MenuItem
                            icon={<ShieldRegular />}
                            onClick={(ev) => { ev.stopPropagation(); void runAction(e.name, () => execEnableManaged(e.name)); }}
                          >
                            Enable Managed
                          </MenuItem>
                        )}
                        <MenuItem
                          icon={<SaveRegular />}
                          onClick={(ev) => { ev.stopPropagation(); setBackupEnv(e); }}
                        >
                          Create Backup
                        </MenuItem>
                        {!envGroupId && (
                          <MenuItem
                            icon={<LayerRegular />}
                            onClick={(ev) => { ev.stopPropagation(); setGroupDialog({ mode: 'add', env: e }); }}
                          >
                            Add to Group
                          </MenuItem>
                        )}
                        {envGroupId && (
                          <MenuItem
                            icon={<SubtractCircleRegular />}
                            onClick={(ev) => { ev.stopPropagation(); setGroupDialog({ mode: 'remove', env: e, groupId: envGroupId }); }}
                          >
                            Remove from Group
                          </MenuItem>
                        )}
                      </MenuList>
                    </MenuPopover>
                  </Menu>
                </div>

                <div className={styles.badgeRow}>
                  <Badge appearance="filled" color={envTypeColor(envType)} size="small">
                    {envType}
                  </Badge>
                  {envIsDisabled && (
                    <Badge appearance="tint" color="warning" size="small">
                      Disabled
                    </Badge>
                  )}
                  {isManaged && (
                    <Badge appearance="outline" color="success" size="small" icon={<LockClosedRegular />}>
                      Managed
                    </Badge>
                  )}
                  {envGroupId && (
                    <Badge appearance="outline" color="informative" size="small" icon={<LayerRegular />}>
                      {groupMap.get(envGroupId) ?? 'In Group'}
                    </Badge>
                  )}
                  {isPending && <Spinner size="tiny" />}
                </div>

                <div className={styles.metaRow}>
                  <span className={styles.metaIcon}>
                    <BoxRegular style={{ fontSize: '0.9rem' }} />
                    <Text>Resources</Text>
                  </span>
                  <Text style={{ fontWeight: tokens.fontWeightSemibold }}>{resourceCount}</Text>
                </div>

                {createdAt !== '—' && (
                  <div className={styles.metaRow}>
                    <span className={styles.metaIcon}>
                      <AppsListRegular style={{ fontSize: '0.9rem' }} />
                      <Text>Created</Text>
                    </span>
                    <Text>{createdAt}</Text>
                  </div>
                )}
              </Card>
            );
          })
        )}
      </div>

      {/* Disable environment confirmation */}
      <ConfirmDialog
        open={confirmAction?.type === 'disable'}
        title="Disable Environment"
        message={`Disable "${confirmAction?.env.properties.displayName ?? confirmAction?.env.name}"? Users will lose access until it is re-enabled.`}
        confirmLabel="Disable"
        isDangerous
        onConfirm={() => {
          if (confirmAction) {
            void runAction(confirmAction.env.name, () => execDisable(confirmAction.env.name));
          }
          setConfirmAction(null);
        }}
        onCancel={() => setConfirmAction(null)}
      />

      {/* Disable managed environment confirmation */}
      <ConfirmDialog
        open={confirmAction?.type === 'disableManaged'}
        title="Disable Managed Environment"
        message={`Remove managed status from "${confirmAction?.env.properties.displayName ?? confirmAction?.env.name}"? Managed environment policies will no longer apply.`}
        confirmLabel="Disable Managed"
        isDangerous
        onConfirm={() => {
          if (confirmAction) {
            void runAction(confirmAction.env.name, () => execDisableManaged(confirmAction.env.name));
          }
          setConfirmAction(null);
        }}
        onCancel={() => setConfirmAction(null)}
      />

      {/* Backup dialog */}
      {backupEnv && (
        <BackupDialog
          open
          environmentName={backupEnv.properties.displayName ?? backupEnv.name}
          isLoading={isBackupLoading}
          onConfirm={(notes) => void execBackup(backupEnv.name, notes)}
          onCancel={() => setBackupEnv(null)}
        />
      )}

      {groupDialog && (
        <EnvironmentGroupDialog
          open
          mode={groupDialog.mode}
          environmentId={groupDialog.env.name}
          environmentName={groupDialog.env.properties.displayName ?? groupDialog.env.name}
          preselectedGroupId={groupDialog.groupId}
          onClose={() => setGroupDialog(null)}
        />
      )}
    </div>
  );
}
