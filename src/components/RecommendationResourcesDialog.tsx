import { useEffect, useState } from 'react';
import type { ReactElement } from 'react';
import {
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  Button,
  Spinner,
  Table,
  TableHeader,
  TableHeaderCell,
  TableBody,
  TableRow,
  TableCell,
  Text,
  Badge,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import { PlayRegular, PersonRegular, GlobeRegular, LightbulbRegular, ErrorCircleRegular } from '@fluentui/react-icons';
import { PowerPlatformforAdminsV2Service } from '../generated/services/PowerPlatformforAdminsV2Service.ts';
import type { AdvisorRecommendationResource } from '../generated/models/PowerPlatformforAdminsV2Model.ts';
import { useMutation } from '../hooks/useMutation.tsx';
import EmptyState from './EmptyState.tsx';

const API_VERSION = '2024-10-01';

interface Action {
  actionType?: string;
  actionName?: string;
}

interface RecommendationResourcesDialogProps {
  open: boolean;
  scenario: string;
  scenarioDisplayName: string;
  actions: Action[];
  onClose: () => void;
}

const useStyles = makeStyles({
  tableWrapper: {
    overflowY: 'auto',
    overflowX: 'auto',
    maxHeight: '50vh',
    minHeight: '120px',
  },
  statusBadge: {
    textTransform: 'capitalize',
  },
  actionsCell: {
    display: 'flex',
    gap: tokens.spacingHorizontalXS,
    flexWrap: 'wrap',
  },
  metaText: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
  },
  footer: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    flex: 1,
    flexWrap: 'nowrap',
  },
  bulkButton: {
    whiteSpace: 'nowrap',
    flexShrink: 0,
  },
  errorRow: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    color: tokens.colorStatusDangerForeground1,
    padding: `${tokens.spacingVerticalM} 0`,
  },
});

function statusColor(status?: string): 'brand' | 'success' | 'warning' | 'danger' | 'informative' {
  if (!status) return 'informative';
  const s = status.toLowerCase();
  if (s.includes('complete') || s.includes('success')) return 'success';
  if (s.includes('fail') || s.includes('error')) return 'danger';
  if (s.includes('pending') || s.includes('progress')) return 'warning';
  return 'informative';
}

function formatDate(value?: string): string {
  if (!value) return '—';
  const d = new Date(value);
  return Number.isNaN(d.getTime()) ? value : d.toLocaleDateString();
}

async function executeForResource(scenario: string, actionName: string, resourceId: string) {
  // actionParameters must be a JSON array of parameter dicts per API contract
  const body = { scenario, actionParameters: [{ resourceId }] } as unknown as Parameters<typeof PowerPlatformforAdminsV2Service.ExecuteRecommendationAction>[0];
  const result = await PowerPlatformforAdminsV2Service.ExecuteRecommendationAction(body, actionName, API_VERSION);
  if (!result.success || result.error) throw new Error(result.error?.message ?? 'Action failed');
  return result.data;
}

async function executeBulk(scenario: string, actionName: string) {
  // Empty array = apply action to all resources
  const body = { scenario, actionParameters: [] } as unknown as Parameters<typeof PowerPlatformforAdminsV2Service.ExecuteRecommendationAction>[0];
  const result = await PowerPlatformforAdminsV2Service.ExecuteRecommendationAction(body, actionName, API_VERSION);
  if (!result.success || result.error) throw new Error(result.error?.message ?? 'Action failed');
  return result.data;
}

export default function RecommendationResourcesDialog({
  open,
  scenario,
  scenarioDisplayName,
  actions,
  onClose,
}: RecommendationResourcesDialogProps): ReactElement {
  const styles = useStyles();
  const [resources, setResources] = useState<AdvisorRecommendationResource[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [fetchError, setFetchError] = useState<string | null>(null);
  const [pendingResourceId, setPendingResourceId] = useState<string | null>(null);
  const [bulkLoadingAction, setBulkLoadingAction] = useState<string | null>(null);

  const { execute: execResource } = useMutation(
    (resId: string, actionName: string) => executeForResource(scenario, actionName, resId),
    { successMessage: 'Action submitted for resource.' },
  );

  const { execute: execBulk } = useMutation(
    (actionName: string) => executeBulk(scenario, actionName),
    {
      successMessage: 'Bulk action submitted for all resources.',
      onSuccess: () => setBulkLoadingAction(null),
      onError: () => setBulkLoadingAction(null),
    },
  );

  useEffect(() => {
    if (!open) { setResources([]); setFetchError(null); return; }
    setIsLoading(true);
    PowerPlatformforAdminsV2Service.GetRecommendationResources(scenario, API_VERSION)
      .then((result) => {
        if (!result.success || result.error) throw new Error(result.error?.message ?? 'Failed to load resources');
        setResources(result.data?.value ?? []);
      })
      .catch((err: unknown) => setFetchError(String(err)))
      .finally(() => setIsLoading(false));
  }, [open, scenario]);

  function bulkActionLabel(actionName?: string): string {
    if (!actionName) return 'Execute All';
    // "Enable X for agents" / "Enable X on canvas apps" → "Enable X All"
    const prepMatch = actionName.match(/^(.+?)\s+(?:for|on|to)\s+.+$/i);
    if (prepMatch) return `${prepMatch[1]} All`;
    // "Delete App" → "Delete All Apps", "Quarantine App" → "Quarantine All Apps"
    if (/\bApps?$/i.test(actionName)) return actionName.replace(/\bApps?$/i, 'All Apps');
    return `${actionName} All`;
  }

  return (
    <Dialog open={open} onOpenChange={(_, d) => { if (!d.open) onClose(); }}>
      <DialogSurface style={{ maxWidth: '900px', width: '90vw' }}>
        <DialogBody>
          <DialogTitle action={<LightbulbRegular aria-hidden="true" style={{ color: tokens.colorBrandForeground1, fontSize: '20px' }} />}>
            {scenarioDisplayName}
          </DialogTitle>
          <DialogContent>
            {isLoading ? (
              <div style={{ display: 'flex', justifyContent: 'center', padding: tokens.spacingVerticalXL }}>
                <Spinner label="Loading resources…" />
              </div>
            ) : fetchError ? (
              <div className={styles.errorRow} role="alert">
                <ErrorCircleRegular aria-hidden="true" />
                <Text>{fetchError}</Text>
              </div>
            ) : resources.length === 0 ? (
              <EmptyState
                icon={<LightbulbRegular />}
                title="No resources found"
                subtitle="No resources are currently associated with this recommendation."
              />
            ) : (
              <div className={styles.tableWrapper}>
                <Table size="small">
                  <TableHeader>
                    <TableRow>
                      <TableHeaderCell>Name</TableHeaderCell>
                      <TableHeaderCell>Type</TableHeaderCell>
                      <TableHeaderCell>Owner</TableHeaderCell>
                      <TableHeaderCell>Environment</TableHeaderCell>
                      <TableHeaderCell>Last Accessed</TableHeaderCell>
                      <TableHeaderCell>Status</TableHeaderCell>
                      {actions.length > 0 && <TableHeaderCell>Actions</TableHeaderCell>}
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {resources.map((r, i) => {
                      const key = r.resourceId ?? `r-${i}`;
                      const isPending = pendingResourceId === key;
                      return (
                        <TableRow key={key}>
                          <TableCell>
                            <Text style={{ fontWeight: tokens.fontWeightSemibold }}>
                              {r.resourceName ?? r.resourceId ?? '—'}
                            </Text>
                            {r.resourceDescription && (
                              <Text style={{ display: 'block', fontSize: tokens.fontSizeBase100, color: tokens.colorNeutralForeground3 }}>
                                {r.resourceDescription}
                              </Text>
                            )}
                          </TableCell>
                          <TableCell>
                            <Text style={{ fontSize: tokens.fontSizeBase200 }}>
                              {r.resourceSubType ?? r.resourceType ?? '—'}
                            </Text>
                          </TableCell>
                          <TableCell>
                            <span className={styles.metaText}>
                              <PersonRegular style={{ fontSize: '0.85rem' }} />
                              {r.resourceOwner ?? r.resourceOwnerId ?? '—'}
                            </span>
                          </TableCell>
                          <TableCell>
                            <span className={styles.metaText}>
                              <GlobeRegular style={{ fontSize: '0.85rem' }} />
                              {r.environmentName ?? r.environmentId ?? '—'}
                            </span>
                          </TableCell>
                          <TableCell>{formatDate(r.lastAccessedDate ?? r.lastModifiedDate)}</TableCell>
                          <TableCell>
                            {r.resourceActionStatus ? (
                              <Badge
                                appearance="tint"
                                color={statusColor(r.resourceActionStatus)}
                                size="small"
                                className={styles.statusBadge}
                              >
                                {r.resourceActionStatus}
                              </Badge>
                            ) : '—'}
                          </TableCell>
                          {actions.length > 0 && (
                            <TableCell>
                              <div className={styles.actionsCell}>
                                {actions.map((a) => (
                                  <Button
                                    key={a.actionType ?? a.actionName}
                                    appearance="outline"
                                    size="small"
                                    icon={isPending ? <Spinner size="tiny" /> : <PlayRegular />}
                                    disabled={isPending || bulkLoadingAction !== null}
                                    onClick={async () => {
                                      setPendingResourceId(key);
                                      await execResource(r.resourceId ?? '', a.actionType ?? a.actionName ?? '');
                                      setPendingResourceId(null);
                                    }}
                                  >
                                    {a.actionName ?? a.actionType}
                                  </Button>
                                ))}
                              </div>
                            </TableCell>
                          )}
                        </TableRow>
                      );
                    })}
                  </TableBody>
                </Table>
              </div>
            )}
          </DialogContent>
          <DialogActions>
            <div className={styles.footer}>
              {!isLoading && resources.length > 0 && (
                <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3, flexShrink: 0 }}>
                  {resources.length} resource(s)
                </Text>
              )}
              <div style={{ marginLeft: 'auto', display: 'flex', gap: tokens.spacingHorizontalM, flexShrink: 0, alignItems: 'center' }}>
                <Button appearance="secondary" onClick={onClose}>Close</Button>
                {resources.length > 0 && actions.map((action, i) => {
                  const actionKey = action.actionType ?? action.actionName ?? String(i);
                  const isThisLoading = bulkLoadingAction === actionKey;
                  const anyBulkLoading = bulkLoadingAction !== null;
                  return (
                    <Button
                      key={actionKey}
                      appearance={i === 0 ? 'primary' : 'outline'}
                      icon={isThisLoading ? <Spinner size="tiny" /> : <PlayRegular />}
                      disabled={anyBulkLoading || isLoading}
                      className={styles.bulkButton}
                      onClick={() => {
                        setBulkLoadingAction(actionKey);
                        void execBulk(actionKey);
                      }}
                    >
                      {isThisLoading ? 'Executing…' : bulkActionLabel(action.actionName ?? action.actionType)}
                    </Button>
                  );
                })}
              </div>
            </div>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
