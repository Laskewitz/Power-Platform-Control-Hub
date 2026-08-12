import { useCallback, useEffect, useMemo, useState } from 'react';
import type { ReactElement } from 'react';
import {
  Badge,
  Button,
  Input,
  MessageBar,
  MessageBarBody,
  Skeleton,
  SkeletonItem,
  Text,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import { ArrowClockwiseRegular, SaveRegular } from '@fluentui/react-icons';
import type { BillingPolicy } from '../types/admin.ts';
import type { Resource } from '../types/inventory.ts';
import {
  fetchCopilotCreditSummary,
  fetchHarnessEnvironmentLicensing,
  getHarnessAgents,
  updateCopilotCreditAllocation,
} from '../services/licensingService.ts';
import type {
  CopilotCreditSummary,
  HarnessEnvironmentLicensing,
} from '../services/licensingService.ts';

interface LicensingViewProps {
  billingPolicies: BillingPolicy[];
  environments: Resource[];
  resources: Resource[];
}

const numberFormatter = new Intl.NumberFormat();

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalL,
    minHeight: 0,
    flex: 1,
    overflowY: 'auto',
    paddingBottom: tokens.spacingVerticalL,
  },
  section: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
  },
  sectionHeader: {
    display: 'flex',
    alignItems: 'flex-start',
    justifyContent: 'space-between',
    gap: tokens.spacingHorizontalM,
  },
  sectionTitle: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
  },
  sectionDescription: {
    color: tokens.colorNeutralForeground3,
    maxWidth: '72ch',
  },
  summary: {
    display: 'grid',
    gridTemplateColumns: 'repeat(4, minmax(120px, 1fr))',
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground1,
    '@media (max-width: 768px)': {
      gridTemplateColumns: 'repeat(2, minmax(120px, 1fr))',
    },
  },
  summaryItem: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXXS,
    padding: tokens.spacingHorizontalM,
    borderRight: `1px solid ${tokens.colorNeutralStroke2}`,
    ':last-child': {
      borderRight: 'none',
    },
    '@media (max-width: 768px)': {
      borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    },
  },
  summaryLabel: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
  summaryValue: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
    fontVariantNumeric: 'tabular-nums',
  },
  tableWrapper: {
    overflowX: 'auto',
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  table: {
    width: '100%',
    borderCollapse: 'collapse',
    minWidth: '860px',
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
  muted: {
    color: tokens.colorNeutralForeground3,
  },
  allocationControl: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
  },
  allocationInput: {
    width: '128px',
  },
  empty: {
    padding: tokens.spacingHorizontalXL,
    textAlign: 'center',
    color: tokens.colorNeutralForeground3,
  },
  skeleton: {
    display: 'grid',
    gap: tokens.spacingVerticalS,
  },
});

function formatNumber(value: number): string {
  return numberFormatter.format(Math.round(value));
}

function formatDate(value?: string): string {
  if (!value) return 'Not reported';
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? value : date.toLocaleDateString();
}

function getBillingBadgeColor(status: BillingPolicy['status']): 'success' | 'warning' {
  return status === 'Enabled' ? 'success' : 'warning';
}

export default function LicensingView({
  billingPolicies,
  environments,
  resources,
}: LicensingViewProps): ReactElement {
  const styles = useStyles();
  const [summary, setSummary] = useState<CopilotCreditSummary | null>(null);
  const [environmentLicensing, setEnvironmentLicensing] = useState<HarnessEnvironmentLicensing[]>([]);
  const [allocationDrafts, setAllocationDrafts] = useState<Record<string, string>>({});
  const [savingEnvironmentId, setSavingEnvironmentId] = useState<string>();
  const [isLoading, setIsLoading] = useState(true);
  const [error, setError] = useState<string>();
  const harnessAgents = useMemo(() => getHarnessAgents(resources), [resources]);

  const loadLicensing = useCallback(async () => {
    setIsLoading(true);
    setError(undefined);
    try {
      const [nextSummary, nextEnvironmentLicensing] = await Promise.all([
        fetchCopilotCreditSummary(),
        fetchHarnessEnvironmentLicensing(environments, resources),
      ]);
      setSummary(nextSummary);
      setEnvironmentLicensing(nextEnvironmentLicensing);
      setAllocationDrafts(
        Object.fromEntries(
          nextEnvironmentLicensing.map((item) => [
            item.environmentId,
            item.allocatedCredits === undefined ? '' : String(item.allocatedCredits),
          ]),
        ),
      );
    } catch (cause: unknown) {
      setError(cause instanceof Error ? cause.message : 'Licensing data could not be loaded.');
    } finally {
      setIsLoading(false);
    }
  }, [environments, resources]);

  useEffect(() => {
    void loadLicensing();
  }, [loadLicensing]);

  async function saveAllocation(item: HarnessEnvironmentLicensing): Promise<void> {
    const draft = allocationDrafts[item.environmentId]?.trim() ?? '';
    const allocatedCredits = Number(draft);
    if (!Number.isInteger(allocatedCredits) || allocatedCredits < 0) {
      setError('Reserved Copilot Credits must be a whole number of zero or greater.');
      return;
    }

    setSavingEnvironmentId(item.environmentId);
    setError(undefined);
    try {
      const savedAllocation = await updateCopilotCreditAllocation(item.environmentId, allocatedCredits);
      setEnvironmentLicensing((current) =>
        current.map((environment) =>
          environment.environmentId === item.environmentId
            ? { ...environment, allocatedCredits: savedAllocation, allocationError: undefined }
            : environment,
        ),
      );
      setAllocationDrafts((current) => ({
        ...current,
        [item.environmentId]: String(savedAllocation),
      }));
      const nextSummary = await fetchCopilotCreditSummary();
      setSummary(nextSummary);
    } catch (cause: unknown) {
      setError(cause instanceof Error ? cause.message : 'The allocation could not be updated.');
    } finally {
      setSavingEnvironmentId(undefined);
    }
  }

  const availableCredits = Math.max(0, (summary?.purchased ?? 0) - (summary?.consumed ?? 0));

  return (
    <div className={styles.root}>
      {error && (
        <MessageBar intent="error">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}

      <section className={styles.section} aria-labelledby="copilot-credit-heading">
        <div className={styles.sectionHeader}>
          <div>
            <Text as="h2" id="copilot-credit-heading" className={styles.sectionTitle}>
              Copilot Credit governance
            </Text>
            <div>
              <Text className={styles.sectionDescription}>
                Track tenant consumption, find GitHub Copilot harness agents, and reserve capacity for the environments that contain them.
              </Text>
            </div>
          </div>
          <Button
            appearance="subtle"
            icon={<ArrowClockwiseRegular />}
            onClick={() => void loadLicensing()}
            disabled={isLoading}
          >
            Refresh
          </Button>
        </div>

        {isLoading ? (
          <Skeleton className={styles.skeleton} aria-label="Loading Copilot Credit totals">
            <SkeletonItem size={48} />
            <SkeletonItem size={48} />
          </Skeleton>
        ) : (
          <div className={styles.summary}>
            <div className={styles.summaryItem}>
              <Text className={styles.summaryLabel}>Purchased</Text>
              <Text className={styles.summaryValue}>{formatNumber(summary?.purchased ?? 0)}</Text>
            </div>
            <div className={styles.summaryItem}>
              <Text className={styles.summaryLabel}>Reserved to environments</Text>
              <Text className={styles.summaryValue}>{formatNumber(summary?.allocated ?? 0)}</Text>
            </div>
            <div className={styles.summaryItem}>
              <Text className={styles.summaryLabel}>Consumed</Text>
              <Text className={styles.summaryValue}>{formatNumber(summary?.consumed ?? 0)}</Text>
            </div>
            <div className={styles.summaryItem}>
              <Text className={styles.summaryLabel}>Available tenant capacity</Text>
              <Text className={styles.summaryValue}>{formatNumber(availableCredits)}</Text>
              <Text className={styles.muted}>Updated {formatDate(summary?.lastUpdated)}</Text>
            </div>
          </div>
        )}

        <MessageBar intent="info">
          <MessageBarBody>
            Classify each environment as maker development or funded production before changing capacity. Development environments need a deliberate spending boundary; production limits should reflect funding, ownership, expected usage, and service criticality.
          </MessageBarBody>
        </MessageBar>
      </section>

      <section className={styles.section} aria-labelledby="harness-heading">
        <div>
          <Text as="h2" id="harness-heading" className={styles.sectionTitle}>
            GitHub Copilot harness environments
          </Text>
          <div>
            <Text className={styles.sectionDescription}>
              {harnessAgents.length} harness {harnessAgents.length === 1 ? 'agent' : 'agents'} detected from Power Platform Inventory. Review the purpose and accountable cost owner before reserving credits.
            </Text>
          </div>
        </div>
        <div className={styles.tableWrapper}>
          <table className={styles.table}>
            <thead>
              <tr>
                <th className={styles.th}>Environment</th>
                <th className={styles.th}>Type</th>
                <th className={styles.th}>Harness agents</th>
                <th className={styles.th}>Owners</th>
                <th className={styles.th}>Reserved credits</th>
                <th className={styles.th}>Pay-as-you-go</th>
              </tr>
            </thead>
            <tbody>
              {!isLoading && environmentLicensing.length === 0 ? (
                <tr>
                  <td className={styles.empty} colSpan={6}>
                    No GitHub Copilot harness agents were found in the current inventory.
                  </td>
                </tr>
              ) : (
                environmentLicensing.map((item) => (
                  <tr key={item.environmentId}>
                    <td className={styles.td}>{item.environmentName}</td>
                    <td className={styles.td}>{item.environmentType}</td>
                    <td className={styles.td}>{item.agentCount}</td>
                    <td className={styles.td}>{item.ownerCount || 'Not identified'}</td>
                    <td className={styles.td}>
                      <div className={styles.allocationControl} title={item.allocationError}>
                        <Input
                          className={styles.allocationInput}
                          type="number"
                          min={0}
                          step={1}
                          value={allocationDrafts[item.environmentId] ?? ''}
                          placeholder={item.allocationError ? 'Not reserved' : undefined}
                          aria-label={`Reserved Copilot Credits for ${item.environmentName}`}
                          onChange={(_, data) =>
                            setAllocationDrafts((current) => ({
                              ...current,
                              [item.environmentId]: data.value,
                            }))
                          }
                        />
                        <Button
                          appearance="subtle"
                          size="small"
                          icon={<SaveRegular />}
                          aria-label={`Save reserved Copilot Credits for ${item.environmentName}`}
                          disabled={savingEnvironmentId !== undefined}
                          onClick={() => void saveAllocation(item)}
                        />
                      </div>
                    </td>
                    <td className={styles.td}>
                      {item.billingPolicyError ? (
                        <Badge
                          appearance="tint"
                          color="subtle"
                          title={item.billingPolicyError}
                        >
                          Not linked or unavailable
                        </Badge>
                      ) : (
                        <Badge
                          appearance="tint"
                          color={item.billingPolicyEnabled ? 'success' : 'warning'}
                        >
                          {item.billingPolicyName ?? 'Disabled'}
                        </Badge>
                      )}
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </section>

      <section className={styles.section} aria-labelledby="billing-heading">
        <div>
          <Text as="h2" id="billing-heading" className={styles.sectionTitle}>Billing policies</Text>
          <div>
            <Text className={styles.sectionDescription}>
              Pay-as-you-go policies provide approved continuity when reserved or tenant capacity is exhausted.
            </Text>
          </div>
        </div>
        <div className={styles.tableWrapper}>
          <table className={styles.table}>
            <thead>
              <tr>
                <th className={styles.th}>Name</th>
                <th className={styles.th}>Status</th>
                <th className={styles.th}>Location</th>
                <th className={styles.th}>Subscription ID</th>
                <th className={styles.th}>Created on</th>
              </tr>
            </thead>
            <tbody>
              {billingPolicies.length === 0 ? (
                <tr>
                  <td className={styles.empty} colSpan={5}>No billing policies found.</td>
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
      </section>
    </div>
  );
}
