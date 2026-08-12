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
  Text,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import { AddRegular, ShieldLockRegular } from '@fluentui/react-icons';
import type { BillingPolicy, CrossTenantReport } from '../types/admin.ts';
import type { Resource } from '../types/inventory.ts';
import type { PolicyV2 } from '../services/dlpService.ts';
import { useMutation } from '../hooks/useMutation.tsx';
import { triggerCrossTenantReport } from '../services/governanceMutations.ts';
import DlpPoliciesView from './DlpPoliciesView.tsx';
import LicensingView from './LicensingView.tsx';

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

type GovernanceTab = 'licensing' | 'crossTenantReports' | 'dlpPolicies';

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
  body: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
    minHeight: 0,
    flex: 1,
    overflow: 'hidden',
  },
  tableWrapper: {
    flex: 1,
    overflowY: 'auto',
    overflowX: 'auto',
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground1,
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

function formatDate(value?: string): string {
  if (!value) return '—';
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? value : date.toLocaleString();
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
              <Spinner size="extra-large" label="Loading data…" />
            </div>
          );
        }
        return renderCrossTenantReportsTable(styles, crossTenantReports, isTriggeringReport, () => void execTriggerReport());
      case 'licensing':
        return (
          <LicensingView
            billingPolicies={billingPolicies}
            environments={environments}
            resources={resources}
          />
        );
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
      <div className={styles.header}>
        <div className={styles.titleBlock}>
          <Text className={styles.title}>Tenant Policies</Text>
          <Text className={styles.subtitle}>DLP, licensing, and cross-tenant governance controls.</Text>
        </div>
      </div>

      <div className={styles.body}>
        <TabList
          selectedValue={activeTab}
          onTabSelect={(_, data) => setActiveTab(data.value as GovernanceTab)}
        >
          <Tab value="dlpPolicies" icon={<ShieldLockRegular />}>DLP Policies</Tab>
          <Tab value="licensing">Licensing</Tab>
          <Tab value="crossTenantReports">Cross-Tenant Reports</Tab>
        </TabList>

        {content}
      </div>

    </div>
  );
}
