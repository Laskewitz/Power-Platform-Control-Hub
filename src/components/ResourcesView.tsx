import { useState, useMemo, useEffect, useRef } from 'react';
import type { ReactElement } from 'react';
import {
  makeStyles,
  tokens,
  Text,
  Input,
  Dropdown,
  Option,
  Badge,
  Button,
  Spinner,
  MessageBar,
  MessageBarBody,
} from '@fluentui/react-components';
import { DeleteRegular, SearchRegular, ArrowClockwiseRegular } from '@fluentui/react-icons';
import type { Resource } from '../types/inventory.ts';
import { RESOURCE_TYPE_LABELS, RESOURCE_TYPES_FILTER } from '../types/inventory.ts';
import ConfirmDialog from './ConfirmDialog.tsx';
import { useMutation } from '../hooks/useMutation.tsx';
import { deleteCopilotAgent } from '../services/resourceMutations.ts';
import { fetchTombstonedIds, addTombstone, removeTombstone } from '../services/tombstoneService.ts';
type SortField = 'name' | 'type' | 'environment' | 'region' | 'created';
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
    gap: tokens.spacingVerticalM,
    padding: tokens.spacingHorizontalXL,
    height: '100%',
    overflow: 'hidden',
  },
  toolbar: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    flexWrap: 'wrap',
    flexShrink: 0,
  },
  title: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
    marginRight: 'auto',
  },
  count: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    whiteSpace: 'nowrap',
  },
  tableWrapper: {
    flex: 1,
    overflowY: 'auto',
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  table: {
    width: '100%',
    borderCollapse: 'collapse',
  },
  th: {
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    textAlign: 'left',
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground3,
    textTransform: 'uppercase',
    letterSpacing: '0.04em',
    borderBottom: `2px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground3,
    cursor: 'pointer',
    userSelect: 'none',
    whiteSpace: 'nowrap',
  },
  td: {
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    fontSize: tokens.fontSizeBase300,
    color: tokens.colorNeutralForeground1,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    verticalAlign: 'middle',
  },
  centered: {
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    height: '100%',
  },
});

function getFieldValue(r: Resource, field: SortField): string {
  switch (field) {
    case 'name': return (r.properties.displayName ?? r.name).toLowerCase();
    case 'type': return r.type.toLowerCase();
    case 'environment': return (r.environmentName ?? '').toLowerCase();
    case 'region': return (r.environmentRegion ?? r.location ?? '').toLowerCase();
    case 'created': return r.properties.createdAt ?? '';
  }
}

const DELETABLE_TYPES = new Set([
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
      const matchesSearch = !term || name.includes(term) || (r.environmentName ?? '').toLowerCase().includes(term);
      return matchesType && matchesSearch;
    });
  }, [resources, search, typeFilter, deletedNames]);

  const sorted = useMemo(() => {
    return [...filtered].sort((a, b) => {
      const av = getFieldValue(a, sortField);
      const bv = getFieldValue(b, sortField);
      const cmp = av.localeCompare(bv);
      return sortDir === 'asc' ? cmp : -cmp;
    });
  }, [filtered, sortField, sortDir]);

  function handleSort(field: SortField) {
    if (field === sortField) {
      setSortDir((d) => (d === 'asc' ? 'desc' : 'asc'));
    } else {
      setSortField(field);
      setSortDir('asc');
    }
  }

  function sortIndicator(field: SortField) {
    if (field !== sortField) return ' ↕';
    return sortDir === 'asc' ? ' ↑' : ' ↓';
  }

  const selectedLabel =
    RESOURCE_TYPES_FILTER.find((t) => t.key === typeFilter)?.label ?? 'All Types';

  if (isLoading) {
    return (
      <div className={styles.centered}>
        <Spinner size="extra-large" label="Loading resources…" />
      </div>
    );
  }

  return (
    <div className={styles.root}>
      <div className={styles.toolbar}>
        <Text className={styles.title}>Resources</Text>
        <Input
          placeholder="Search by name or environment…"
          value={search}
          onChange={(_, data) => setSearch(data.value)}
          contentBefore={<SearchRegular />}
          size="small"
          style={{ minWidth: '220px' }}
        />
        <Dropdown
          value={selectedLabel}
          selectedOptions={[typeFilter]}
          onOptionSelect={(_, data) => setTypeFilter(data.optionValue ?? 'all')}
          size="small"
          style={{ minWidth: '160px' }}
        >
          {RESOURCE_TYPES_FILTER.map((t) => (
            <Option key={t.key} value={t.key}>{t.label}</Option>
          ))}
        </Dropdown>
        <Text className={styles.count}>{sorted.length} result(s)</Text>
        <Button
          appearance="subtle"
          icon={<ArrowClockwiseRegular />}
          size="small"
          onClick={() => void onRefresh()}
          title="Refresh"
        />
      </div>

      {error && (
        <MessageBar intent="error">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}

      <div className={styles.tableWrapper}>
        <table className={styles.table}>
          <thead>
            <tr>
              <th className={styles.th} onClick={() => handleSort('name')}>
                Name{sortIndicator('name')}
              </th>
              <th className={styles.th} onClick={() => handleSort('type')}>
                Type{sortIndicator('type')}
              </th>
              <th className={styles.th} onClick={() => handleSort('environment')}>
                Environment{sortIndicator('environment')}
              </th>
              <th className={styles.th} onClick={() => handleSort('region')}>
                Region{sortIndicator('region')}
              </th>
              <th className={styles.th} onClick={() => handleSort('created')}>
                Created{sortIndicator('created')}
              </th>
              <th className={styles.th} style={{ width: '60px' }}></th>
            </tr>
          </thead>
          <tbody>
            {sorted.length === 0 ? (
              <tr>
                <td
                  className={styles.td}
                  colSpan={6}
                  style={{ textAlign: 'center', color: tokens.colorNeutralForeground3 }}
                >
                  No resources match your filters.
                </td>
              </tr>
            ) : (
              sorted.map((r, i) => (
                <tr key={r.id ?? `${r.type}-${r.name}-${i}`}>
                  <td className={styles.td}>
                    {r.properties.displayName ?? r.name}
                  </td>
                  <td className={styles.td}>
                    <Badge appearance="outline" size="small">
                      {RESOURCE_TYPE_LABELS[r.type.toLowerCase()] ?? r.type}
                    </Badge>
                  </td>
                  <td className={styles.td}>{r.environmentName ?? '—'}</td>
                  <td className={styles.td}>
                    {r.environmentRegion ?? r.location ?? '—'}
                  </td>
                  <td className={styles.td}>
                    {r.properties.createdAt
                      ? new Date(r.properties.createdAt).toLocaleDateString()
                      : '—'}
                  </td>
                  <td className={styles.td}>
                    {DELETABLE_TYPES.has(r.type.toLowerCase()) && (
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
                  </td>
                </tr>
              ))
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
