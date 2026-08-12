import { useMemo, useState } from 'react';
import type { ReactElement } from 'react';
import {
  Badge,
  Button,
  MessageBar,
  MessageBarBody,
  Spinner,
  Tab,
  TabList,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import { AddRegular, DocumentRibbonRegular, ShieldLockRegular, GlobeRegular } from '@fluentui/react-icons';
import type { BillingPolicy, CrossTenantReport } from '../types/admin.ts';
import type { Resource } from '../types/inventory.ts';
import type { PolicyV2 } from '../services/dlpService.ts';
import { useMutation } from '../hooks/useMutation.tsx';
import { triggerCrossTenantReport } from '../services/governanceMutations.ts';
import DlpPoliciesView from './DlpPoliciesView.tsx';
import EmptyState from './EmptyState.tsx';
import { PageHeader } from './ui.tsx';

interface GovernanceViewProps {
  billingPolicies: BillingPolicy[];
  crossTenantReports: CrossTenantReport[];
  dlpPolicies: PolicyV2[];
  environments: Resource[];
  resources: Resource[];
  isLoading: boolean;
  error: string | null;
  onRefreshAdmin?: () => Promise<void>;
}

type GovernanceTab = 'billingPolicies' | 'crossTenantReports' | 'dlpPolicies';

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
  body: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
    minHeight: 0,
    flex: 1,
    overflow: 'hidden',
  },
  toolbar: {
    display: 'flex',
    justifyContent: 'flex-end',
  },
  board: {
    flex: 1,
    minHeight: 0,
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
    minWidth: '600px',
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
    whiteSpace: 'nowrap',
  },
  td: {
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    fontSize: tokens.fontSizeBase300,
    color: tokens.colorNeutralForeground1,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    verticalAlign: 'middle',
  },
  tabular: {
    fontVariantNumeric: 'tabular-nums',
  },
  centered: {
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    height: '100%',
  },
});

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

function renderBillingPoliciesTable(
  styles: ReturnType<typeof useStyles>,
  billingPolicies: BillingPolicy[],
): ReactElement {
  if (billingPolicies.length === 0) {
    return (
      <EmptyState
        icon={<DocumentRibbonRegular />}
        title="No billing policies found"
        subtitle="Billing policies created for pay-as-you-go environments will appear here."
      />
    );
  }
  return (
    <div className={styles.board}>
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
          {billingPolicies.map((policy) => (
            <tr key={policy.id}>
              <td className={styles.td}>{policy.name}</td>
              <td className={styles.td}>
                <Badge appearance="filled" color={getBillingBadgeColor(policy.status)}>
                  {policy.status}
                </Badge>
              </td>
              <td className={styles.td}>{policy.location}</td>
              <td className={styles.td} style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: tokens.fontSizeBase200 }}>
                {policy.billingInstrument.subscriptionId}
              </td>
              <td className={styles.td}>{formatDate(policy.createdOn)}</td>
            </tr>
          ))}
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
      <div className={styles.toolbar}>
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
      {crossTenantReports.length === 0 ? (
        <EmptyState
          icon={<GlobeRegular />}
          title="No cross-tenant reports found"
          subtitle="Trigger a new report to inspect inbound and outbound connections with other tenants."
          action={{ label: 'Trigger New Report', onClick: onTrigger }}
        />
      ) : (
        <div className={styles.board}>
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
              {crossTenantReports.map((report) => {
                const inboundCount = report.status === 'Completed'
                  ? report.connections?.filter((connection) => connection.connectionType === 'Inbound').length ?? 0
                  : null;
                const outboundCount = report.status === 'Completed'
                  ? report.connections?.filter((connection) => connection.connectionType === 'Outbound').length ?? 0
                  : null;

                return (
                  <tr key={report.reportId}>
                    <td className={styles.td} style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: tokens.fontSizeBase200 }}>
                      {report.reportId}
                    </td>
                    <td className={styles.td}>{formatDate(report.requestDate)}</td>
                    <td className={styles.td}>
                      <Badge appearance="filled" color={getReportBadgeColor(report.status)}>
                        {report.status}
                      </Badge>
                    </td>
                    <td className={`${styles.td} ${styles.tabular}`}>{inboundCount ?? '—'}</td>
                    <td className={`${styles.td} ${styles.tabular}`}>{outboundCount ?? '—'}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

export default function GovernanceView({
  billingPolicies,
  crossTenantReports,
  dlpPolicies,
  environments,
  resources,
  isLoading,
  error,
  onRefreshAdmin,
}: GovernanceViewProps): ReactElement {
  const styles = useStyles();
  const [activeTab, setActiveTab] = useState<GovernanceTab>('dlpPolicies');

  const { execute: execTriggerReport, isLoading: isTriggeringReport } = useMutation(triggerCrossTenantReport, {
    successMessage: 'Cross-tenant report request submitted. Results will appear when processing completes.',
    onSuccess: () => void onRefreshAdmin?.(),
  });

  const refreshDlpPolicies = onRefreshAdmin ?? (async () => {});

  const content = useMemo(() => {
    if (error) {
      return (
        <MessageBar intent="error">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      );
    }

    switch (activeTab) {
      case 'crossTenantReports':
        if (isLoading) {
          return (
            <div className={styles.centered}>
              <Spinner size="extra-large" label="Loading cross-tenant reports…" />
            </div>
          );
        }
        return renderCrossTenantReportsTable(styles, crossTenantReports, isTriggeringReport, () => void execTriggerReport());
      case 'billingPolicies':
        if (isLoading) {
          return (
            <div className={styles.centered}>
              <Spinner size="extra-large" label="Loading billing policies…" />
            </div>
          );
        }
        return renderBillingPoliciesTable(styles, billingPolicies);
      case 'dlpPolicies':
      default:
        return <DlpPoliciesView dlpPolicies={dlpPolicies} environments={environments} resources={resources} isLoading={isLoading} onRefresh={refreshDlpPolicies} />;
    }
  }, [
    activeTab,
    billingPolicies,
    crossTenantReports,
    dlpPolicies,
    environments,
    error,
    execTriggerReport,
    isLoading,
    isTriggeringReport,
    refreshDlpPolicies,
    resources,
    styles,
  ]);

  return (
    <div className={styles.root}>
      <PageHeader
        title="Tenant Policies"
        description="DLP policies, billing policies, and cross-tenant governance insights for the tenant."
      />

      <div className={styles.body}>
        <TabList
          selectedValue={activeTab}
          onTabSelect={(_, data) => setActiveTab(data.value as GovernanceTab)}
        >
          <Tab value="dlpPolicies" icon={<ShieldLockRegular />}>DLP Policies</Tab>
          <Tab value="billingPolicies">Billing Policies</Tab>
          <Tab value="crossTenantReports">Cross-Tenant Reports</Tab>
        </TabList>

        {content}
      </div>
    </div>
  );
}
