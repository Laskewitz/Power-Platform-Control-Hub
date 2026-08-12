import { useState, useMemo, useEffect, useRef } from 'react';
import type { ReactElement } from 'react';
import {
  makeStyles,
  tokens,
  Text,
  Input,
  Dropdown,
  Option,
  OptionGroup,
  Badge,
  Button,
  MessageBar,
  MessageBarBody,
} from '@fluentui/react-components';
import {
  DeleteRegular,
  SearchRegular,
  ArrowClockwiseRegular,
  OpenRegular,
  FilterRegular,
  LockClosedRegular,
  ArrowSortUpRegular,
  ArrowSortDownRegular,
  ArrowSortRegular,
} from '@fluentui/react-icons';
import CloudFlowDetailPanel from './CloudFlowDetailPanel.tsx';
import CanvasAppDetailPanel from './CanvasAppDetailPanel.tsx';
import CopilotStudioAgentDetailPanel from './CopilotStudioAgentDetailPanel.tsx';
import EmptyState from './EmptyState.tsx';
import type { Resource } from '../types/inventory.ts';
import { RESOURCE_TYPE_LABELS, RESOURCE_TYPE_SHORT_LABELS, RESOURCE_TYPES_FILTER, getTypeBadgeColor } from '../types/inventory.ts';
import ConfirmDialog from './ConfirmDialog.tsx';
import { extractMessage } from '../utils/errorUtils.ts';
import { useMutation } from '../hooks/useMutation.tsx';
import { deleteCopilotAgent } from '../services/resourceMutations.ts';
import { fetchTombstonedIds, addTombstone, removeTombstone } from '../services/tombstoneService.ts';
import { formatDate } from '../utils/formatDate.ts';
import { OperationsSkeleton, PageHeader } from './ui.tsx';
import { useEnvironmentOwners } from '../services/ownerCache.ts';
type SortField = 'name' | 'type' | 'environment' | 'region' | 'owner' | 'created' | 'lastModified';
type SortDir = 'asc' | 'desc';

interface ResourcesViewProps {
  resources: Resource[];
  isLoading: boolean;
  error: string | null;
  onRefresh: () => Promise<void>;
  initialTypeFilter?: string;
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
  },
  controlRail: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    flexWrap: 'wrap',
    flexShrink: 0,
    padding: '0 0 16px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
    backgroundColor: 'transparent',
  },
  count: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    whiteSpace: 'nowrap',
    marginLeft: 'auto',
  },
  board: {
    flex: 1,
    overflowY: 'auto',
    overflowX: 'auto',
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: 0,
    backgroundColor: tokens.colorNeutralBackground1,
    boxShadow: tokens.shadow4,
  },
  table: {
    width: '100%',
    borderCollapse: 'collapse',
    tableLayout: 'fixed',
    minWidth: '640px',
  },
  th: {
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    textAlign: 'left',
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground3,
    textTransform: 'uppercase',
    letterSpacing: '0.04em',
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
    backgroundColor: tokens.colorNeutralBackground3,
    cursor: 'pointer',
    userSelect: 'none',
    whiteSpace: 'nowrap',
    ':hover': {
      color: tokens.colorNeutralForeground1,
      backgroundColor: tokens.colorNeutralBackground3Hover,
    },
  },
  thInner: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXXS,
  },
  thIcon: {
    display: 'inline-flex',
    opacity: 0.55,
    fontSize: '12px',
  },
  td: {
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    fontSize: tokens.fontSizeBase300,
    color: tokens.colorNeutralForeground1,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    verticalAlign: 'middle',
    overflow: 'hidden',
  },
  tdText: {
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
    display: 'block',
  },
  nameCell: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    minWidth: 0,
  },
  tr: {
    cursor: 'pointer',
    ':hover td': {
      backgroundColor: tokens.colorNeutralBackground1Hover,
    },
    ':focus-visible': {
      outline: `2px solid ${tokens.colorBrandStroke1}`,
      outlineOffset: '-2px',
    },
    ':focus-within td': {
      backgroundColor: tokens.colorNeutralBackground1Hover,
    },
  },
  trNonClickable: {
    ':hover td': {
      backgroundColor: tokens.colorNeutralBackground1Hover,
    },
  },
  centered: {
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    height: '100%',
  },
});

function getOwnerId(resource: Resource): string | undefined {
  const properties = resource.properties as Record<string, unknown>;
  if (properties.owner && typeof properties.owner === 'object') {
    return (properties.owner as { id?: string }).id;
  }
  if (typeof properties.createdBy === 'string') return properties.createdBy;
  if (properties.createdBy && typeof properties.createdBy === 'object') {
    return (properties.createdBy as { id?: string }).id;
  }
  return typeof properties.ownerId === 'string' ? properties.ownerId : undefined;
}

function getOwnerDisplay(r: Resource, resolvedOwners?: Map<string, string>): string {
  const p = r.properties as Record<string, unknown>;
  // Use pre-resolved AAD display name if available
  if (typeof p.resolvedOwnerName === 'string') return p.resolvedOwnerName;
  if (p.owner && typeof p.owner === 'object') {
    const o = p.owner as { displayName?: string; email?: string; id?: string };
    const resolved = o.id ? resolvedOwners?.get(o.id.toLowerCase()) : undefined;
    return o.displayName ?? o.email ?? resolved ?? o.id ?? '—';
  }
  if (p.createdBy) {
    if (typeof p.createdBy === 'string') {
      return resolvedOwners?.get(p.createdBy.toLowerCase()) ?? p.createdBy;
    }
    const cb = p.createdBy as { displayName?: string; email?: string; id?: string };
    const resolved = cb.id ? resolvedOwners?.get(cb.id.toLowerCase()) : undefined;
    return cb.displayName ?? cb.email ?? resolved ?? cb.id ?? '—';
  }
  if (typeof p.ownerId === 'string') {
    return resolvedOwners?.get(p.ownerId.toLowerCase()) ?? p.ownerId;
  }
  return '—';
}

function getFieldValue(r: Resource, field: SortField, resolvedOwners: Map<string, string>): string {
  switch (field) {
    case 'name': return (r.properties.displayName ?? r.name).toLowerCase();
    case 'type': return r.type.toLowerCase();
    case 'environment': return (r.environmentName ?? '').toLowerCase();
    case 'region': return (r.environmentRegion ?? r.location ?? '').toLowerCase();
    case 'owner': return getOwnerDisplay(r, resolvedOwners).toLowerCase();
    case 'created': return r.properties.createdAt ?? '';
    case 'lastModified': return r.properties.lastModifiedAt ?? r.properties.modifiedAt ?? r.properties.lastPublishedAt ?? '';
  }
}

const DELETABLE_TYPES = new Set<string>([]);

const DETAIL_PANEL_TYPES = new Set([
  'microsoft.powerautomate/cloudflows',
  'microsoft.powerautomate/agentflows',
  'microsoft.powerautomate/m365agentflows',
  'microsoft.powerapps/apps',
  'microsoft.powerapps/canvasapps',
  'microsoft.copilotstudio/agents',
]);

export default function ResourcesView({
  resources,
  isLoading,
  error,
  onRefresh,
  initialTypeFilter,
}: ResourcesViewProps): ReactElement {
  const styles = useStyles();
  const [search, setSearch] = useState('');
  const [typeFilter, setTypeFilter] = useState(initialTypeFilter ?? 'all');
  const [sortField, setSortField] = useState<SortField>('created');
  const [sortDir, setSortDir] = useState<SortDir>('desc');
  const [confirmDelete, setConfirmDelete] = useState<Resource | null>(null);
  const [pendingResourceName, setPendingResourceName] = useState<string | null>(null);
  const [deletedNames, setDeletedNames] = useState<Set<string>>(new Set());
  const [detailResource, setDetailResource] = useState<Resource | null>(null);
  const [ownerFilter, setOwnerFilter] = useState('all');
  const [envFilter, setEnvFilter] = useState('all');
  const ownerRequests = useMemo(
    () => resources.flatMap((resource) => {
      const guid = getOwnerId(resource);
      if (!guid) return [];
      const environmentId = typeof resource.properties.environmentId === 'string'
        ? resource.properties.environmentId
        : undefined;
      return [{ guid, environmentId }];
    }),
    [resources],
  );
  const resolvedOwners = useEnvironmentOwners(ownerRequests);

  // Load tombstones from Dataverse (+ localStorage fallback) on mount
  useEffect(() => {
    void fetchTombstonedIds().then(setDeletedNames);
  }, []);
  const pendingDeleteRef = useRef<string | null>(null);

  const { execute: execDeleteAgent } = useMutation(deleteCopilotAgent, {
    successMessage: 'Copilot agent deleted.',
    onSuccess: () => setPendingResourceName(null),
    onError: () => {
      if (pendingDeleteRef.current) {
        removeTombstone(pendingDeleteRef.current);
        setDeletedNames((prev) => { const n = new Set(prev); n.delete(pendingDeleteRef.current!); return n; });
        pendingDeleteRef.current = null;
      }
      setPendingResourceName(null);
    },
  });

  // Sync when navigating from Overview tile
  useEffect(() => {
    setTypeFilter(initialTypeFilter ?? 'all');
  }, [initialTypeFilter]);

  const filtered = useMemo(() => {
    const term = search.toLowerCase();
    return resources.filter((r) => {
      if (deletedNames.has(r.name)) return false;
      const matchesType = typeFilter === 'all' || r.type.toLowerCase() === typeFilter;
      const name = (r.properties.displayName ?? r.name).toLowerCase();
      const matchesSearch = !term || name.includes(term);
      const matchesOwner = ownerFilter === 'all' || getOwnerDisplay(r, resolvedOwners) === ownerFilter;
      const matchesEnv = envFilter === 'all' || (r.environmentName ?? '') === envFilter;
      return matchesType && matchesSearch && matchesOwner && matchesEnv;
    });
  }, [resources, search, typeFilter, ownerFilter, envFilter, deletedNames, resolvedOwners]);

  const sorted = useMemo(() => {
    return [...filtered].sort((a, b) => {
      const av = getFieldValue(a, sortField, resolvedOwners);
      const bv = getFieldValue(b, sortField, resolvedOwners);
      const cmp = av.localeCompare(bv);
      return sortDir === 'asc' ? cmp : -cmp;
    });
  }, [filtered, sortField, sortDir, resolvedOwners]);

  const uniqueOwners = useMemo(() => {
    const seen = new Set<string>();
    for (const r of resources) {
      const o = getOwnerDisplay(r, resolvedOwners);
      if (o !== '—') seen.add(o);
    }
    return Array.from(seen).sort((a, b) => a.localeCompare(b));
  }, [resources, resolvedOwners]);

  const uniqueEnvironments = useMemo(() => {
    const seen = new Set<string>();
    for (const r of resources) {
      const e = r.environmentName ?? '';
      if (e) seen.add(e);
    }
    return Array.from(seen).sort((a, b) => a.localeCompare(b));
  }, [resources]);

  function handleSort(field: SortField) {
    if (field === sortField) {
      setSortDir((d) => (d === 'asc' ? 'desc' : 'asc'));
    } else {
      setSortField(field);
      setSortDir('asc');
    }
  }

  function sortIndicator(field: SortField): ReactElement {
    if (field !== sortField) return <ArrowSortRegular className={styles.thIcon} aria-hidden="true" />;
    return sortDir === 'asc'
      ? <ArrowSortUpRegular className={styles.thIcon} aria-hidden="true" />
      : <ArrowSortDownRegular className={styles.thIcon} aria-hidden="true" />;
  }

  function ariaSortAttr(field: SortField): 'ascending' | 'descending' | 'none' {
    if (field !== sortField) return 'none';
    return sortDir === 'asc' ? 'ascending' : 'descending';
  }

  const selectedLabel =
    RESOURCE_TYPES_FILTER.find((t) => t.key === typeFilter)?.label ?? 'All Types';

  if (isLoading) {
    return <OperationsSkeleton />;
  }

  // Full-page detail view
  if (detailResource) {
    const typeLower = detailResource.type.toLowerCase();
    if (typeLower === 'microsoft.powerapps/apps' || typeLower === 'microsoft.powerapps/canvasapps') {
      return (
        <CanvasAppDetailPanel
          resource={detailResource}
          onClose={() => setDetailResource(null)}
        />
      );
    }
    // Cloud flows, agent flows, and M365 agent flows all use the same panel
    if (
      typeLower === 'microsoft.powerautomate/cloudflows' ||
      typeLower === 'microsoft.powerautomate/agentflows' ||
      typeLower === 'microsoft.powerautomate/m365agentflows'
    ) {
      return (
        <CloudFlowDetailPanel
          resource={detailResource}
          onClose={() => setDetailResource(null)}
          onDeleted={(name) => {
            setDeletedNames((prev) => new Set([...prev, name]));
            setDetailResource(null);
          }}
        />
      );
    }
    // Copilot Studio agents
    if (typeLower === 'microsoft.copilotstudio/agents') {
      return (
        <CopilotStudioAgentDetailPanel
          resource={detailResource}
          onClose={() => setDetailResource(null)}
          onDeleted={(name) => {
            setDeletedNames((prev) => new Set([...prev, name]));
            setDetailResource(null);
          }}
        />
      );
    }
  }

  return (
    <div className={styles.root}>
      <PageHeader
        title="Resources"
        description="Search, filter, and manage apps, flows, and agents registered across the tenant."
        actions={
          <Button appearance="secondary" icon={<ArrowClockwiseRegular />} onClick={() => void onRefresh()}>
            Refresh
          </Button>
        }
      />

      <div className={styles.controlRail}>
        <Input
          placeholder="Search by name…"
          value={search}
          onChange={(_, data) => setSearch(data.value)}
          contentBefore={<SearchRegular />}
          size="small"
          style={{ minWidth: '220px' }}
        />
        <Dropdown
          value={envFilter === 'all' ? 'All Environments' : envFilter}
          selectedOptions={[envFilter]}
          onOptionSelect={(_, data) => setEnvFilter(data.optionValue ?? 'all')}
          size="small"
          style={{ minWidth: '180px' }}
        >
          <Option value="all">All Environments</Option>
          {uniqueEnvironments.map((e) => (
            <Option key={e} value={e}>{e}</Option>
          ))}
        </Dropdown>
        <Dropdown
          value={selectedLabel}
          selectedOptions={[typeFilter]}
          onOptionSelect={(_, data) => setTypeFilter(data.optionValue ?? 'all')}
          size="small"
          style={{ minWidth: '160px' }}
        >
          <OptionGroup label="Overall">
            <Option value="all">All Types</Option>
          </OptionGroup>
          <OptionGroup label="By Type">
            {[...RESOURCE_TYPES_FILTER.filter(t => t.key !== 'all')]
              .sort((a, b) => a.label.localeCompare(b.label))
              .map((t) => (
                <Option key={t.key} value={t.key}>{t.label}</Option>
              ))
            }
          </OptionGroup>
        </Dropdown>
        <Dropdown
          value={ownerFilter === 'all' ? 'All Owners' : ownerFilter}
          selectedOptions={[ownerFilter]}
          onOptionSelect={(_, data) => setOwnerFilter(data.optionValue ?? 'all')}
          size="small"
          style={{ minWidth: '160px' }}
        >
          <Option value="all">All Owners</Option>
          {uniqueOwners.map((o) => (
            <Option key={o} value={o}>{o}</Option>
          ))}
        </Dropdown>
        <Text className={styles.count}>{sorted.length} result(s)</Text>
      </div>

      {error && (
        <MessageBar intent="error">
          <MessageBarBody>{extractMessage(error)}</MessageBarBody>
        </MessageBar>
      )}

      <div className={styles.board}>
        <table className={styles.table}>
          <colgroup>
            <col style={{ width: '22%' }} />
            <col style={{ width: '110px' }} />
            <col style={{ width: '20%' }} />
            <col style={{ width: '10%' }} />
            <col style={{ width: '16%' }} />
            <col style={{ width: '96px' }} />
            <col style={{ width: '96px' }} />
            <col style={{ width: '44px' }} />
          </colgroup>
          <thead>
            <tr>
              {(
                [
                  ['name', 'Name'],
                  ['type', 'Type'],
                  ['environment', 'Environment'],
                  ['region', 'Region'],
                  ['owner', 'Owner'],
                  ['created', 'Created'],
                  ['lastModified', 'Modified'],
                ] as const
              ).map(([field, label]) => (
                <th
                  key={field}
                  className={styles.th}
                  scope="col"
                  aria-sort={ariaSortAttr(field)}
                >
                  <button
                    type="button"
                    className={styles.thInner}
                    onClick={() => handleSort(field)}
                    style={{ background: 'none', border: 0, padding: 0, font: 'inherit', color: 'inherit', cursor: 'pointer' }}
                  >
                    {label}
                    {sortIndicator(field)}
                  </button>
                </th>
              ))}
              <th className={styles.th} scope="col" aria-label="Actions" />
            </tr>
          </thead>
          <tbody>
            {sorted.length === 0 ? (
              <tr>
                <td colSpan={8} style={{ padding: 0 }}>
                  <EmptyState
                    icon={<FilterRegular />}
                    title="No resources match your filters"
                    subtitle="Try adjusting your search term, environment, type, or owner filter."
                    action={search || typeFilter !== 'all' || envFilter !== 'all' || ownerFilter !== 'all'
                      ? { label: 'Clear filters', onClick: () => { setSearch(''); setTypeFilter('all'); setEnvFilter('all'); setOwnerFilter('all'); } }
                      : undefined}
                  />
                </td>
              </tr>
            ) : (
              sorted.map((r, i) => {
                const displayName = r.properties.displayName ?? r.name;
                const envName = r.environmentName;
                const typeLower = r.type.toLowerCase();
                const isClickable = DETAIL_PANEL_TYPES.has(typeLower);
                return (
                  <tr
                    key={r.id ?? `${r.type}-${r.name}-${i}`}
                    className={isClickable ? styles.tr : styles.trNonClickable}
                    onClick={isClickable ? () => setDetailResource(r) : undefined}
                    tabIndex={isClickable ? 0 : undefined}
                    onKeyDown={isClickable ? (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); setDetailResource(r); } } : undefined}
                    aria-label={isClickable ? `Open details for ${displayName}` : undefined}
                  >
                    <td className={styles.td} title={displayName}>
                      <span className={styles.nameCell}>
                        <span className={styles.tdText}>{displayName}</span>
                        {r.properties.isQuarantined === true && (
                          <Badge appearance="tint" color="danger" size="small" icon={<LockClosedRegular />}>
                            Quarantined
                          </Badge>
                        )}
                      </span>
                    </td>
                    <td className={styles.td}>
                      <Badge
                        appearance="tint"
                        color={getTypeBadgeColor(typeLower)}
                        size="small"
                        style={{ whiteSpace: 'nowrap' }}
                      >
                        {RESOURCE_TYPE_SHORT_LABELS[typeLower] ?? RESOURCE_TYPE_LABELS[typeLower] ?? r.type}
                      </Badge>
                    </td>
                    <td className={styles.td} title={envName ?? 'Environment not available in API for this resource type'}>
                      {envName
                        ? <span className={styles.tdText}>{envName}</span>
                        : <span className={styles.tdText} style={{ color: tokens.colorNeutralForeground4 }}>Not available</span>
                      }
                    </td>
                    <td className={styles.td}>
                      <span className={styles.tdText}>{r.environmentRegion ?? r.location ?? '—'}</span>
                    </td>
                    <td className={styles.td} title={getOwnerDisplay(r, resolvedOwners)}>
                      <span className={styles.tdText}>{getOwnerDisplay(r, resolvedOwners)}</span>
                    </td>
                    <td className={styles.td}>
                      <span className={styles.tdText}>{formatDate(r.properties.createdAt)}</span>
                    </td>
                    <td className={styles.td}>
                      <span className={styles.tdText}>{formatDate((r.properties.lastModifiedAt ?? r.properties.modifiedAt ?? r.properties.lastPublishedAt) as string | undefined)}</span>
                    </td>
                    <td className={styles.td} onClick={(e) => e.stopPropagation()}>
                      {DELETABLE_TYPES.has(typeLower) && (
                        <Button
                          appearance="subtle"
                          size="small"
                          icon={<DeleteRegular />}
                          title="Delete"
                          disabled={pendingResourceName === r.name}
                          style={{ color: tokens.colorStatusDangerForeground1 }}
                          onClick={() => setConfirmDelete(r)}
                        />
                      )}
                      {isClickable && (
                        <Button
                          appearance="subtle"
                          size="small"
                          icon={<OpenRegular />}
                          title="View details"
                          onClick={(e) => { e.stopPropagation(); setDetailResource(r); }}
                        />
                      )}
                    </td>
                  </tr>
                );
              })
            )}
          </tbody>
        </table>
      </div>

      <ConfirmDialog
        open={Boolean(confirmDelete)}
        title="Delete Copilot Agent"
        message={`Delete "${confirmDelete?.properties.displayName ?? confirmDelete?.name}"? This action cannot be undone.`}
        confirmLabel="Delete"
        isDangerous
        isLoading={pendingResourceName !== null}
        onConfirm={() => {
          if (confirmDelete) {
            const name = confirmDelete.name;
            const displayName = confirmDelete.properties.displayName ?? name;
            setPendingResourceName(name);
            pendingDeleteRef.current = name;
            addTombstone({
              resourceId: name,
              resourceType: confirmDelete.type,
              environmentId: confirmDelete.properties.environmentId ?? '',
              displayName,
              deletedBy: '',
            });
            setDeletedNames((prev) => new Set([...prev, name]));
            void execDeleteAgent(confirmDelete.properties.environmentId ?? '', name);
          }
          setConfirmDelete(null);
        }}
        onCancel={() => setConfirmDelete(null)}
      />
    </div>
  );
}
