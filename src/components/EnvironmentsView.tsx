import { useMemo, useState } from 'react';
import type { ReactElement } from 'react';
import {
  makeStyles,
  tokens,
  Text,
  Card,
  Badge,
  Input,
  Spinner,
  MessageBar,
  MessageBarBody,
  Divider,
  Button,
  Menu,
  MenuTrigger,
  MenuList,
  MenuItem,
  MenuPopover,
  Dropdown,
  Option,
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
} from '@fluentui/react-icons';
import type { Resource } from '../types/inventory.ts';
import type { EnvironmentGroup } from '../types/admin.ts';
import ConfirmDialog from './ConfirmDialog.tsx';
import BackupDialog from './BackupDialog.tsx';
import EnvironmentDetailView from './EnvironmentDetailView.tsx';
import EnvironmentGroupDialog from './EnvironmentGroupDialog.tsx';
import { useMutation } from '../hooks/useMutation.tsx';
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
    gap: tokens.spacingVerticalL,
    padding: tokens.spacingHorizontalXL,
    height: '100%',
    overflow: 'hidden',
    '@media (max-width: 768px)': {
      padding: tokens.spacingHorizontalM,
    },
  },
  toolbar: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    flexShrink: 0,
    flexWrap: 'wrap',
  },
  title: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
    marginRight: 'auto',
  },
  count: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
  },
  grid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fill, minmax(380px, 1fr))',
    gap: tokens.spacingHorizontalXL,
    overflowY: 'auto',
    flex: 1,
    paddingBottom: tokens.spacingVerticalL,
    alignContent: 'start',
    alignItems: 'start',
    '@media (max-width: 480px)': {
      gridTemplateColumns: '1fr',
    },
  },
  card: {
    display: 'flex',
    flexDirection: 'column',
    padding: tokens.spacingVerticalM,
    gap: tokens.spacingVerticalS,
    minWidth: 0,
    boxSizing: 'border-box',
    minHeight: '160px',
  },
  cardTop: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    minWidth: 0,
  },
  cardIcon: {
    fontSize: '1.25rem',
    color: tokens.colorBrandForeground1,
    flexShrink: 0,
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
    fontSize: tokens.fontSizeBase300,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
  },
  cardRegion: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
  },
  badgeRow: {
    display: 'flex',
    gap: tokens.spacingHorizontalS,
    flexWrap: 'wrap',
    paddingTop: tokens.spacingVerticalXS,
  },
  divider: {
    marginTop: tokens.spacingVerticalS,
    marginBottom: tokens.spacingVerticalXS,
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

  if (isLoading) {
    return (
      <div className={styles.centered}>
        <Spinner size="extra-large" label="Loading environments…" />
      </div>
    );
  }

  if (selectedEnv) {
    return (
      <EnvironmentDetailView
        environment={selectedEnv}
        resources={resources}
        envGroups={envGroups}
        onBack={() => setSelectedEnv(null)}
        onRefreshEnvironments={onRefreshEnvironments}
      />
    );
  }

  return (
    <div className={styles.root}>
      <div className={styles.toolbar}>
        <Text className={styles.title}>Environments</Text>
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
          <Text style={{ color: tokens.colorNeutralForeground3 }}>
            No environments found.
          </Text>
        ) : (
          filtered.map((e, i) => {
            const displayName = e.properties.displayName ?? e.name;
            const envType = (e.properties.environmentType ?? 'Unknown') as string;
            const isManaged = e.properties.isManaged === true;
            const region = e.location ?? '—';
            const resourceCount = countByEnv.get(displayName.toLowerCase()) ?? 0;
            const createdAt = e.properties.createdAt
              ? new Date(e.properties.createdAt as string).toLocaleDateString()
              : null;
            const isPending = pendingEnvId === e.name;
            const envGroupId = e.properties['environmentGroupId'] as string | undefined;

            return (
              <Card
                key={e.id ?? `env-${e.name}-${i}`}
                className={styles.card}
                onClick={() => setSelectedEnv(e)}
                style={{ cursor: 'pointer' }}
              >
                <div className={styles.cardTop}>
                  <GlobeRegular className={styles.cardIcon} />
                  <div className={styles.cardTitles}>
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
                        <MenuItem
                          icon={<PlayRegular />}
                          onClick={(ev) => { ev.stopPropagation(); void runAction(e.name, () => execEnable(e.name)); }}
                        >
                          Enable
                        </MenuItem>
                        <MenuItem
                          icon={<StopRegular />}
                          onClick={(ev) => { ev.stopPropagation(); setConfirmAction({ type: 'disable', env: e }); }}
                        >
                          Disable
                        </MenuItem>
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

                <Divider className={styles.divider} />

                <div className={styles.metaRow}>
                  <span className={styles.metaIcon}>
                    <BoxRegular style={{ fontSize: '0.9rem' }} />
                    <Text>Resources</Text>
                  </span>
                  <Text style={{ fontWeight: tokens.fontWeightSemibold }}>{resourceCount}</Text>
                </div>

                {createdAt && (
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
