import { useMemo, useState } from 'react';
import type { ReactElement } from 'react';
import {
  Badge,
  Button,
  Card,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Field,
  Input,
  MessageBar,
  MessageBarBody,
  Spinner,
  Tab,
  TabList,
  Text,
  Textarea,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import { AddRegular, DeleteRegular, EditRegular, PeopleRegular, ShieldRegular } from '@fluentui/react-icons';
import type {
  BillingPolicy,
  CrossTenantReport,
  EnvironmentGroup,
  RoleAssignment,
} from '../types/admin.ts';
import ConfirmDialog from './ConfirmDialog.tsx';
import { useMutation } from '../hooks/useMutation.tsx';
import { useOwners } from '../services/ownerCache.ts';
import {
  createEnvironmentGroup,
  deleteEnvironmentGroup,
  triggerCrossTenantReport,
  updateEnvironmentGroup,
} from '../services/governanceMutations.ts';

interface GovernanceViewProps {
  roleAssignments: RoleAssignment[];
  envGroups: EnvironmentGroup[];
  billingPolicies: BillingPolicy[];
  crossTenantReports: CrossTenantReport[];
  isLoading: boolean;
  error: string | null;
  onRefreshAdmin?: () => Promise<void>;
}

type GovernanceTab = 'environmentGroups' | 'billingPolicies' | 'crossTenantReports';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalL,
    padding: tokens.spacingHorizontalXL,
    height: '100%',
    overflow: 'hidden',
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalL,
    flexWrap: 'wrap',
    flexShrink: 0,
  },
  titleBlock: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXS,
    marginRight: 'auto',
  },
  title: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
  },
  subtitle: {
    color: tokens.colorNeutralForeground3,
  },
  metricCard: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    padding: tokens.spacingVerticalM,
    minWidth: '220px',
  },
  metricLabel: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    color: tokens.colorNeutralForeground2,
  },
  metricValue: {
    fontSize: tokens.fontSizeHero700,
    fontWeight: tokens.fontWeightBold,
    lineHeight: tokens.lineHeightHero700,
  },
  body: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
    minHeight: 0,
    flex: 1,
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

interface GroupFormDialogProps {
  open: boolean;
  initial?: EnvironmentGroup;
  isLoading: boolean;
  onSubmit: (displayName: string, description: string) => void;
  onCancel: () => void;
}

function GroupFormDialog({ open, initial, isLoading, onSubmit, onCancel }: GroupFormDialogProps): ReactElement {
  const [displayName, setDisplayName] = useState(initial?.displayName ?? '');
  const [description, setDescription] = useState(initial?.description ?? '');
  const isEdit = Boolean(initial);
  return (
    <Dialog open={open} onOpenChange={(_, data) => { if (!data.open) onCancel(); }}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>{isEdit ? 'Edit Environment Group' : 'New Environment Group'}</DialogTitle>
          <DialogContent style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM }}>
            <Field label="Display name" required>
              <Input value={displayName} onChange={(_, d) => setDisplayName(d.value)} placeholder="My Group" />
            </Field>
            <Field label="Description">
              <Textarea value={description} onChange={(_, d) => setDescription(d.value)} rows={3} />
            </Field>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" disabled={isLoading} onClick={onCancel}>Cancel</Button>
            <Button appearance="primary" disabled={isLoading || !displayName.trim()} onClick={() => onSubmit(displayName.trim(), description.trim())}>
              {isLoading ? 'Saving…' : isEdit ? 'Save' : 'Create'}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

function formatDate(value?: string): string {
  if (!value) return '—';
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? value : date.toLocaleString();
}

function getBillingBadgeColor(status: BillingPolicy['status']): 'success' | 'warning' {
  return status === 'Enabled' ? 'success' : 'warning';
}

function getReportBadgeColor(
  status: CrossTenantReport['status'],
): 'success' | 'informative' | 'danger' | 'subtle' {
  switch (status) {
    case 'Completed':
      return 'success';
    case 'InProgress':
      return 'informative';
    case 'Failed':
      return 'danger';
    case 'Received':
    default:
      return 'subtle';
  }
}

function renderEnvironmentGroupsTable(
  styles: ReturnType<typeof useStyles>,
  envGroups: EnvironmentGroup[],
  pendingGroupId: string | null,
  ownerNames: Map<string, string>,
  onEdit: (group: EnvironmentGroup) => void,
  onDelete: (group: EnvironmentGroup) => void,
): ReactElement {
  return (
    <div className={styles.tableWrapper}>
      <table className={styles.table}>
        <thead>
          <tr>
            <th className={styles.th}>Display Name</th>
            <th className={styles.th}>Description</th>
            <th className={styles.th}>Created By</th>
            <th className={styles.th}>Created Time</th>
            <th className={styles.th}>Child Count</th>
            <th className={styles.th} style={{ width: '80px' }}></th>
          </tr>
        </thead>
        <tbody>
          {envGroups.length === 0 ? (
            <tr>
              <td className={styles.td} colSpan={6} style={{ textAlign: 'center' }}>
                No environment groups found.
              </td>
            </tr>
          ) : (
            envGroups.map((group) => (
              <tr key={group.id}>
                <td className={styles.td}>{group.displayName}</td>
                <td className={styles.td}>{group.description || '—'}</td>
                <td className={styles.td}>{ownerNames.get(group.createdBy.id?.toLowerCase()) ?? group.createdBy.displayName ?? group.createdBy.email ?? group.createdBy.id}</td>
                <td className={styles.td}>{formatDate(group.createdTime)}</td>
                <td className={styles.td}>{group.childrenGroupIds?.length ?? 0}</td>
                <td className={styles.td}>
                  <span style={{ display: 'flex', gap: tokens.spacingHorizontalXS }}>
                    <Button
                      appearance="subtle" size="small"
                      icon={<EditRegular />}
                      disabled={pendingGroupId === group.id}
                      title="Edit"
                      onClick={() => onEdit(group)}
                    />
                    <Button
                      appearance="subtle" size="small"
                      icon={<DeleteRegular />}
                      disabled={pendingGroupId === group.id}
                      title="Delete"
                      style={{ color: tokens.colorStatusDangerForeground1 }}
                      onClick={() => onDelete(group)}
                    />
                  </span>
                </td>
              </tr>
            ))
          )}
        </tbody>
      </table>
    </div>
  );
}

function renderBillingPoliciesTable(
  styles: ReturnType<typeof useStyles>,
  billingPolicies: BillingPolicy[],
): ReactElement {
  return (
    <div className={styles.tableWrapper}>
      <table className={styles.table}>
        <thead>
          <tr>
            <th className={styles.th}>Name</th>
            <th className={styles.th}>Status</th>
            <th className={styles.th}>Location</th>
            <th className={styles.th}>Subscription ID</th>
            <th className={styles.th}>Created On</th>
          </tr>
        </thead>
        <tbody>
          {billingPolicies.length === 0 ? (
            <tr>
              <td className={styles.td} colSpan={5} style={{ textAlign: 'center' }}>
                No billing policies found.
              </td>
            </tr>
          ) : (
            billingPolicies.map((policy) => (
              <tr key={policy.id}>
                <td className={styles.td}>{policy.name}</td>
                <td className={styles.td}>
                  <Badge appearance="filled" color={getBillingBadgeColor(policy.status)}>
                    {policy.status}
                  </Badge>
                </td>
                <td className={styles.td}>{policy.location}</td>
                <td className={styles.td}>{policy.billingInstrument.subscriptionId}</td>
                <td className={styles.td}>{formatDate(policy.createdOn)}</td>
              </tr>
            ))
          )}
        </tbody>
      </table>
    </div>
  );
}

function renderCrossTenantReportsTable(
  styles: ReturnType<typeof useStyles>,
  crossTenantReports: CrossTenantReport[],
  isTriggerLoading: boolean,
  onTrigger: () => void,
): ReactElement {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM, flex: 1, minHeight: 0 }}>
      <div style={{ display: 'flex', justifyContent: 'flex-end' }}>
        <Button
          appearance="primary"
          size="small"
          icon={<AddRegular />}
          disabled={isTriggerLoading}
          onClick={onTrigger}
        >
          {isTriggerLoading ? 'Submitting…' : 'Trigger New Report'}
        </Button>
      </div>
      <div className={styles.tableWrapper}>
        <table className={styles.table}>
          <thead>
            <tr>
              <th className={styles.th}>Report ID</th>
              <th className={styles.th}>Request Date</th>
              <th className={styles.th}>Status</th>
              <th className={styles.th}>Inbound</th>
              <th className={styles.th}>Outbound</th>
            </tr>
          </thead>
          <tbody>
            {crossTenantReports.length === 0 ? (
              <tr>
                <td className={styles.td} colSpan={5} style={{ textAlign: 'center' }}>
                  No cross-tenant reports found.
                </td>
              </tr>
            ) : (
              crossTenantReports.map((report) => {
                const inboundCount = report.status === 'Completed'
                  ? report.connections?.filter((connection) => connection.connectionType === 'Inbound').length ?? 0
                  : null;
                const outboundCount = report.status === 'Completed'
                  ? report.connections?.filter((connection) => connection.connectionType === 'Outbound').length ?? 0
                  : null;

                return (
                  <tr key={report.reportId}>
                    <td className={styles.td}>{report.reportId}</td>
                    <td className={styles.td}>{formatDate(report.requestDate)}</td>
                    <td className={styles.td}>
                      <Badge appearance="filled" color={getReportBadgeColor(report.status)}>
                        {report.status}
                      </Badge>
                    </td>
                    <td className={styles.td}>{inboundCount ?? '—'}</td>
                    <td className={styles.td}>{outboundCount ?? '—'}</td>
                  </tr>
                );
              })
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

export default function GovernanceView({
  roleAssignments,
  envGroups,
  billingPolicies,
  crossTenantReports,
  isLoading,
  error,
  onRefreshAdmin,
}: GovernanceViewProps): ReactElement {
  const styles = useStyles();
  const [activeTab, setActiveTab] = useState<GovernanceTab>('environmentGroups');
  const [groupFormOpen, setGroupFormOpen] = useState(false);
  const [editingGroup, setEditingGroup] = useState<EnvironmentGroup | null>(null);
  const [deleteGroup, setDeleteGroup] = useState<EnvironmentGroup | null>(null);
  const [pendingGroupId, setPendingGroupId] = useState<string | null>(null);

  const groupOwnerGuids = useMemo(
    () => envGroups.map((g) => g.createdBy?.id as string | undefined),
    [envGroups],
  );
  const groupOwnerNames = useOwners(groupOwnerGuids);

  const { execute: execCreate, isLoading: isCreating } = useMutation(createEnvironmentGroup, {
    successMessage: 'Environment group created.',
    onSuccess: () => { setGroupFormOpen(false); void onRefreshAdmin?.(); },
  });
  const { execute: execUpdate, isLoading: isUpdating } = useMutation(updateEnvironmentGroup, {
    successMessage: 'Environment group updated.',
    onSuccess: () => { setEditingGroup(null); void onRefreshAdmin?.(); },
  });
  const { execute: execDelete } = useMutation(deleteEnvironmentGroup, {
    successMessage: 'Environment group deleted.',
    onSuccess: () => { setPendingGroupId(null); void onRefreshAdmin?.(); },
  });
  const { execute: execTriggerReport, isLoading: isTriggeringReport } = useMutation(triggerCrossTenantReport, {
    successMessage: 'Cross-tenant report request submitted. Results will appear when processing completes.',
    onSuccess: () => void onRefreshAdmin?.(),
  });

  const content = useMemo(() => {
    if (isLoading) {
      return (
        <div className={styles.centered}>
          <Spinner size="extra-large" label="Loading governance data…" />
        </div>
      );
    }

    if (error) {
      return (
        <MessageBar intent="error">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      );
    }

    switch (activeTab) {
      case 'billingPolicies':
        return renderBillingPoliciesTable(styles, billingPolicies);
      case 'crossTenantReports':
        return renderCrossTenantReportsTable(styles, crossTenantReports, isTriggeringReport, () => void execTriggerReport());
      case 'environmentGroups':
      default:
        return (
          <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM, flex: 1, minHeight: 0 }}>
            <div style={{ display: 'flex', justifyContent: 'flex-end' }}>
              <Button appearance="primary" size="small" icon={<AddRegular />} onClick={() => setGroupFormOpen(true)}>
                New Group
              </Button>
            </div>
            {renderEnvironmentGroupsTable(styles, envGroups, pendingGroupId, groupOwnerNames, (g) => setEditingGroup(g), (g) => setDeleteGroup(g))}
          </div>
        );
    }
  }, [activeTab, billingPolicies, crossTenantReports, envGroups, error, execTriggerReport, isTriggeringReport, isLoading, pendingGroupId, styles]);

  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <div className={styles.titleBlock}>
          <Text className={styles.title}>Governance</Text>
          <Text className={styles.subtitle}>Tenant-wide governance and authorization insights.</Text>
        </div>

        <Card className={styles.metricCard}>
          <Text className={styles.metricLabel}>
            <PeopleRegular /> Tenant role assignments
          </Text>
          <Text className={styles.metricValue}>{roleAssignments.length}</Text>
        </Card>
      </div>

      <div className={styles.body}>
        <TabList
          selectedValue={activeTab}
          onTabSelect={(_, data) => setActiveTab(data.value as GovernanceTab)}
        >
          <Tab value="environmentGroups" icon={<ShieldRegular />}>
            Environment Groups
          </Tab>
          <Tab value="billingPolicies">Billing Policies</Tab>
          <Tab value="crossTenantReports">Cross-Tenant Reports</Tab>
        </TabList>

        {content}
      </div>

      {/* Create group dialog */}
      <GroupFormDialog
        open={groupFormOpen}
        isLoading={isCreating}
        onSubmit={(name, desc) => void execCreate(name, desc)}
        onCancel={() => setGroupFormOpen(false)}
      />

      {/* Edit group dialog */}
      {editingGroup && (
        <GroupFormDialog
          open
          initial={editingGroup}
          isLoading={isUpdating}
          onSubmit={(name, desc) => void execUpdate(editingGroup.id, name, desc)}
          onCancel={() => setEditingGroup(null)}
        />
      )}

      {/* Delete group confirmation */}
      <ConfirmDialog
        open={Boolean(deleteGroup)}
        title="Delete Environment Group"
        message={`Delete group "${deleteGroup?.displayName}"? This action cannot be undone.`}
        confirmLabel="Delete"
        isDangerous
        onConfirm={() => {
          if (deleteGroup) {
            setPendingGroupId(deleteGroup.id);
            void execDelete(deleteGroup.id);
          }
          setDeleteGroup(null);
        }}
        onCancel={() => setDeleteGroup(null)}
      />
    </div>
  );
}
