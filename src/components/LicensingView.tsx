import { useCallback, useEffect, useMemo, useState } from 'react';
import type { ReactElement } from 'react';
import {
  Badge,
  Button,
  Input,
  MessageBar,
  MessageBarBody,
  Tab,
  TabList,
  Text,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import {
  ArrowClockwiseRegular,
  BotRegular,
  ChevronLeftRegular,
  ChevronRightRegular,
  DatabaseRegular,
  KeyRegular,
  MoneyRegular,
  PersonWarningRegular,
  SearchRegular,
  SaveRegular,
} from '@fluentui/react-icons';
import type { LicensingSnapshot } from '../types/admin.ts';
import type { Resource } from '../types/inventory.ts';
import {
  fetchHarnessEnvironmentLicensing,
  getHarnessAgents,
  updateCopilotCreditAllocation,
} from '../services/licensingService.ts';
import type { HarnessEnvironmentLicensing } from '../services/licensingService.ts';
import EmptyState from './EmptyState.tsx';
import { OperationsSkeleton, PageHeader } from './ui.tsx';

interface LicensingViewProps {
  snapshot: LicensingSnapshot | null;
  isLoading: boolean;
  error: string | null;
  onRefresh: () => Promise<void>;
  onPeriodChange: (startDate: string, endDate: string) => Promise<void>;
  environments: Resource[];
  resources: Resource[];
}

type LicensingTab = 'capacity' | 'harness' | 'entitlements' | 'compliance';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    gap: '18px',
    height: '100%',
    padding: '28px 32px 32px',
    overflowY: 'auto',
    overflowX: 'hidden',
    '@media (max-width: 768px)': { padding: tokens.spacingHorizontalM },
  },
  scrollBody: {
    display: 'flex',
    flexDirection: 'column',
    gap: '18px',
    flexShrink: 0,
  },
  statusDeck: {
    display: 'grid',
    gridTemplateColumns: 'minmax(260px, 1.4fr) repeat(3, minmax(150px, 0.65fr))',
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    backgroundColor: tokens.colorNeutralStroke2,
    gap: '1px',
    boxShadow: tokens.shadow4,
    '@media (max-width: 980px)': { gridTemplateColumns: 'repeat(2, minmax(0, 1fr))' },
    '@media (max-width: 560px)': { gridTemplateColumns: '1fr' },
  },
  posture: {
    display: 'grid',
    alignContent: 'center',
    gap: '7px',
    minHeight: '126px',
    padding: '20px 22px',
    backgroundColor: tokens.colorNeutralBackground1,
    backgroundImage: 'linear-gradient(90deg, rgba(67, 217, 255, 0.06), transparent 72%)',
  },
  postureLine: { display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' },
  postureTitle: {
    color: tokens.colorNeutralForeground1,
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
  },
  postureMessage: { color: tokens.colorNeutralForeground3, lineHeight: '20px', maxWidth: '62ch' },
  stat: {
    display: 'grid',
    alignContent: 'center',
    gap: '4px',
    minHeight: '126px',
    padding: '18px',
    backgroundColor: tokens.colorNeutralBackground1,
  },
  statValue: {
    color: tokens.colorNeutralForeground1,
    fontSize: '28px',
    lineHeight: '32px',
    fontWeight: tokens.fontWeightBold,
    fontVariantNumeric: 'tabular-nums',
    letterSpacing: '-0.03em',
  },
  statLabel: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase100,
    fontWeight: tokens.fontWeightSemibold,
    letterSpacing: '0.07em',
    textTransform: 'uppercase',
  },
  tabBar: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: '16px',
    flexWrap: 'wrap',
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  periodControl: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    marginLeft: 'auto',
    paddingBottom: '6px',
  },
  periodCopy: {
    display: 'grid',
    gap: '1px',
    minWidth: '240px',
    textAlign: 'right',
    '@media (max-width: 620px)': {
      minWidth: 0,
      textAlign: 'left',
    },
  },
  periodLabel: {
    color: tokens.colorNeutralForeground1,
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
  },
  periodHint: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase100,
  },
  boardGrid: {
    display: 'grid',
    gridTemplateColumns: 'minmax(0, 1fr)',
    gap: '18px',
    alignItems: 'start',
  },
  monitor: {
    minWidth: 0,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    backgroundColor: tokens.colorNeutralBackground1,
    boxShadow: tokens.shadow4,
  },
  monitorHeader: {
    display: 'flex',
    alignItems: 'center',
    gap: '9px',
    minHeight: '46px',
    padding: '0 14px',
    color: tokens.colorNeutralForeground1,
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
    backgroundColor: tokens.colorNeutralBackground3,
    fontWeight: tokens.fontWeightSemibold,
  },
  tableScroll: { width: '100%' },
  table: { width: '100%', tableLayout: 'fixed', borderCollapse: 'collapse' },
  th: {
    padding: '9px 13px',
    color: tokens.colorNeutralForeground3,
    backgroundColor: tokens.colorNeutralBackground2,
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
    textAlign: 'left',
    fontSize: tokens.fontSizeBase100,
    fontWeight: tokens.fontWeightSemibold,
    letterSpacing: '0.06em',
    textTransform: 'uppercase',
    whiteSpace: 'nowrap',
    '@media (max-width: 700px)': {
      padding: '8px',
      whiteSpace: 'normal',
    },
  },
  td: {
    padding: '10px 13px',
    color: tokens.colorNeutralForeground2,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    fontSize: tokens.fontSizeBase200,
    verticalAlign: 'middle',
    overflowWrap: 'anywhere',
    '@media (max-width: 700px)': {
      padding: '9px 8px',
      fontSize: tokens.fontSizeBase100,
    },
  },
  strong: { color: tokens.colorNeutralForeground1, fontWeight: tokens.fontWeightSemibold },
  numeric: { fontVariantNumeric: 'tabular-nums', whiteSpace: 'nowrap' },
  secondary: { display: 'block', color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase100, marginTop: '2px' },
  toolbar: { display: 'flex', alignItems: 'center', gap: '12px', justifyContent: 'space-between', flexWrap: 'wrap' },
  summaryRows: { display: 'grid' },
  summaryRow: {
    display: 'grid',
    gridTemplateColumns: 'minmax(160px, 1fr) repeat(4, minmax(90px, auto))',
    gap: '12px',
    alignItems: 'center',
    minHeight: '58px',
    padding: '9px 13px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    '@media (max-width: 700px)': {
      gridTemplateColumns: 'repeat(2, minmax(0, 1fr))',
      '& > :first-child': { gridColumn: '1 / -1' },
    },
  },
  summaryValue: { display: 'grid', gap: '1px', fontVariantNumeric: 'tabular-nums' },
  ledgerGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(3, minmax(0, 1fr))',
    gap: '1px',
    backgroundColor: tokens.colorNeutralStroke2,
    '@media (max-width: 1100px)': { gridTemplateColumns: 'repeat(2, minmax(0, 1fr))' },
    '@media (max-width: 650px)': { gridTemplateColumns: '1fr' },
  },
  ledgerCell: {
    display: 'grid',
    gap: '16px',
    minHeight: '126px',
    padding: '16px',
    backgroundColor: tokens.colorNeutralBackground1,
  },
  ledgerTitle: {
    display: 'grid',
    gap: '3px',
    minWidth: 0,
  },
  ledgerName: {
    color: tokens.colorNeutralForeground1,
    fontWeight: tokens.fontWeightSemibold,
    overflowWrap: 'anywhere',
  },
  ledgerDate: { color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase100 },
  ledgerValues: {
    display: 'grid',
    gridTemplateColumns: 'repeat(3, minmax(0, 1fr))',
    gap: '10px',
  },
  ledgerValue: {
    display: 'grid',
    gap: '2px',
    minWidth: 0,
    color: tokens.colorNeutralForeground1,
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    fontVariantNumeric: 'tabular-nums',
  },
  ledgerLabel: {
    color: tokens.colorNeutralForeground3,
    fontSize: '10px',
    fontWeight: tokens.fontWeightSemibold,
    letterSpacing: '0.06em',
    textTransform: 'uppercase',
  },
  warning: { color: tokens.colorStatusWarningForeground1, fontWeight: tokens.fontWeightSemibold },
  danger: { color: tokens.colorStatusDangerForeground1, fontWeight: tokens.fontWeightSemibold },
  emptyInline: { padding: '26px' },
  allocationControl: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },
  allocationInput: { width: '132px' },
  guidance: {
    maxWidth: '78ch',
    color: tokens.colorNeutralForeground2,
    lineHeight: '20px',
  },
});

function labelize(value: string): string {
  return value
    .replace(/([A-Z])([A-Z][a-z])/g, '$1 $2')
    .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
    .replace(/[_-]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

const CURRENCY_LABELS: Record<string, string> = {
  AI: 'AI credits',
  MCSMessages: 'Copilot Credits',
  PAUnattendedRPA: 'Power Automate unattended RPA',
  ProcessMiningDataStorage: 'Process Mining data storage',
  TenantM365Copilot: 'Microsoft 365 Copilot',
  W365APAYGO: 'Windows 365 pay-as-you-go',
};

function formatNumber(value: number): string {
  return new Intl.NumberFormat(undefined, { maximumFractionDigits: value < 10 ? 2 : 0 }).format(value);
}

function formatDate(value?: string): string {
  if (!value) return '—';
  const date = new Date(value);
  return Number.isNaN(date.getTime())
    ? value
    : date.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' });
}

function shiftPeriod(startDate: string, direction: -1 | 1): { startDate: string; endDate: string } {
  const start = new Date(`${startDate}T00:00:00Z`);
  start.setUTCDate(start.getUTCDate() + direction * 30);
  const end = new Date(start);
  end.setUTCDate(end.getUTCDate() + 29);
  return {
    startDate: start.toISOString().slice(0, 10),
    endDate: end.toISOString().slice(0, 10),
  };
}

function badgeColor(status: string): 'success' | 'warning' | 'danger' | 'informative' {
  const normalized = status.toLowerCase();
  if (normalized.includes('available') || normalized.includes('healthy') || normalized.includes('enabled')) return 'success';
  if (normalized.includes('critical') || normalized.includes('exceed') || normalized.includes('unavailable')) return 'danger';
  if (normalized.includes('warning') || normalized.includes('limited')) return 'warning';
  return 'informative';
}

export default function LicensingView({
  snapshot,
  isLoading,
  error,
  onRefresh,
  onPeriodChange,
  environments,
  resources,
}: LicensingViewProps): ReactElement {
  const styles = useStyles();
  const [activeTab, setActiveTab] = useState<LicensingTab>('capacity');
  const [search, setSearch] = useState('');
  const [harnessEnvironments, setHarnessEnvironments] = useState<HarnessEnvironmentLicensing[]>([]);
  const [allocationDrafts, setAllocationDrafts] = useState<Record<string, string>>({});
  const [harnessError, setHarnessError] = useState<string | null>(null);
  const [isHarnessLoading, setIsHarnessLoading] = useState(false);
  const [savingEnvironmentId, setSavingEnvironmentId] = useState<string>();
  const harnessAgents = useMemo(() => getHarnessAgents(resources), [resources]);

  const loadHarnessEnvironments = useCallback(async () => {
    setIsHarnessLoading(true);
    setHarnessError(null);
    try {
      const next = await fetchHarnessEnvironmentLicensing(environments, resources);
      setHarnessEnvironments(next);
      setAllocationDrafts(Object.fromEntries(
        next.map((item) => [
          item.environmentId,
          item.allocatedCredits === undefined ? '' : String(item.allocatedCredits),
        ]),
      ));
    } catch (reason: unknown) {
      setHarnessError(reason instanceof Error ? reason.message : 'Harness environment licensing could not be loaded.');
    } finally {
      setIsHarnessLoading(false);
    }
  }, [environments, resources]);

  useEffect(() => {
    void loadHarnessEnvironments();
  }, [loadHarnessEnvironments]);

  async function saveAllocation(item: HarnessEnvironmentLicensing): Promise<void> {
    const allocatedCredits = Number(allocationDrafts[item.environmentId]?.trim() ?? '');
    if (!Number.isInteger(allocatedCredits) || allocatedCredits < 0) {
      setHarnessError('Reserved Copilot Credits must be a whole number of zero or greater.');
      return;
    }

    setSavingEnvironmentId(item.environmentId);
    setHarnessError(null);
    try {
      const saved = await updateCopilotCreditAllocation(item.environmentId, allocatedCredits);
      setHarnessEnvironments((current) => current.map((environment) =>
        environment.environmentId === item.environmentId
          ? { ...environment, allocatedCredits: saved, allocationError: undefined }
          : environment,
      ));
      setAllocationDrafts((current) => ({ ...current, [item.environmentId]: String(saved) }));
      await onRefresh();
    } catch (reason: unknown) {
      setHarnessError(reason instanceof Error ? reason.message : 'The Copilot Credit allocation could not be updated.');
    } finally {
      setSavingEnvironmentId(undefined);
    }
  }

  const filteredEntitlements = useMemo(() => {
    if (!snapshot) return [];
    const term = search.trim().toLowerCase();
    if (!term) return snapshot.entitlements;
    return snapshot.entitlements.filter((item) =>
      [item.displayName, item.entitlementCode, item.capacityType, item.skuId]
        .some((value) => value?.toLowerCase().includes(term)),
    );
  }, [search, snapshot]);

  const riskTotals = useMemo(() => (
    (snapshot?.compliance ?? []).reduce(
      (totals, row) => ({
        exceeding: totals.exceeding + row.usersExceedingCapacity,
        unlicensed: totals.unlicensed + row.usersWithoutLicense + row.usersWithoutPremiumLicense,
      }),
      { exceeding: 0, unlicensed: 0 },
    )
  ), [snapshot]);
  const today = new Date().toISOString().slice(0, 10);

  if (isLoading && !snapshot) return <OperationsSkeleton />;

  return (
    <div className={styles.root}>
      <PageHeader
        title="Licensing & Capacity"
        description="Tenant capacity, Copilot Credit governance, entitlements, and per-flow compliance through Power Platform Admin V2 connector actions."
        actions={(
          <Button
            appearance="primary"
            icon={<ArrowClockwiseRegular />}
            disabled={isLoading || isHarnessLoading}
            onClick={() => void Promise.all([onRefresh(), loadHarnessEnvironments()])}
          >
            {isLoading || isHarnessLoading ? 'Refreshing…' : 'Refresh licensing'}
          </Button>
        )}
      />

      <div className={styles.scrollBody}>
        {error && (
          <MessageBar intent="error">
            <MessageBarBody>{error}</MessageBarBody>
          </MessageBar>
        )}
        {snapshot?.warnings.length ? (
          <MessageBar intent="warning">
            <MessageBarBody>
              Some licensing signals are unavailable: {snapshot.warnings.join(' · ')}
            </MessageBarBody>
          </MessageBar>
        ) : null}
        {harnessError && (
          <MessageBar intent="warning">
            <MessageBarBody>{harnessError}</MessageBarBody>
          </MessageBar>
        )}

        {!snapshot ? (
          <EmptyState
            icon={<KeyRegular />}
            title="Licensing signal unavailable"
            subtitle="The tenant did not return capacity or entitlement data. Confirm that the Admin V2 connection can access tenant capacity APIs."
            action={{ label: 'Retry licensing', onClick: () => void onRefresh() }}
          />
        ) : (
          <>
            <section className={styles.statusDeck} aria-label="Licensing posture">
              <div className={styles.posture}>
                <div className={styles.postureLine}>
                  <Text className={styles.postureTitle}>Tenant capacity {labelize(snapshot.capacityStatus)}</Text>
                  <Badge appearance="filled" color={badgeColor(snapshot.capacityStatus)}>
                    {labelize(snapshot.capacityStatus)}
                  </Badge>
                </div>
                <Text className={styles.postureMessage}>
                  {snapshot.capacityStatusMessage || `${snapshot.licenseModelType} license model reporting across ${snapshot.capacities.length} capacity families.`}
                </Text>
              </div>
              <div className={styles.stat}>
                <Text className={styles.statValue}>{snapshot.entitlements.length}</Text>
                <Text className={styles.statLabel}>Entitlements indexed</Text>
              </div>
              <div className={styles.stat}>
                <Text className={`${styles.statValue} ${riskTotals.exceeding ? styles.warning : ''}`}>{riskTotals.exceeding}</Text>
                <Text className={styles.statLabel}>Users over capacity</Text>
              </div>
              <div className={styles.stat}>
                <Text className={`${styles.statValue} ${riskTotals.unlicensed ? styles.danger : ''}`}>{riskTotals.unlicensed}</Text>
                <Text className={styles.statLabel}>License gaps</Text>
              </div>
            </section>

            <div className={styles.tabBar}>
              <TabList
                selectedValue={activeTab}
                onTabSelect={(_, data) => setActiveTab(data.value as LicensingTab)}
                aria-label="Licensing registers"
              >
                <Tab value="capacity" icon={<DatabaseRegular />}>Capacity</Tab>
                <Tab value="harness" icon={<BotRegular />}>Harness agents</Tab>
                <Tab value="entitlements" icon={<KeyRegular />}>Entitlements</Tab>
                <Tab value="compliance" icon={<PersonWarningRegular />}>Compliance</Tab>
              </TabList>
              <div className={styles.periodControl}>
                <Button
                  appearance="subtle"
                  size="small"
                  icon={<ChevronLeftRegular />}
                  aria-label="Load previous 30-day compliance window"
                  title="Previous 30 days"
                  disabled={isLoading}
                  onClick={() => {
                    const period = shiftPeriod(snapshot.periodStart, -1);
                    void onPeriodChange(period.startDate, period.endDate);
                  }}
                />
                <div className={styles.periodCopy}>
                  <Text className={styles.periodLabel}>
                    {formatDate(snapshot.periodStart)}–{formatDate(snapshot.periodEnd)}
                  </Text>
                  <Text className={styles.periodHint}>
                    Rolling 30-day compliance window · not the billing cycle
                  </Text>
                </div>
                <Button
                  appearance="subtle"
                  size="small"
                  icon={<ChevronRightRegular />}
                  aria-label="Load next 30-day compliance window"
                  title="Next 30 days"
                  disabled={isLoading || snapshot.periodEnd >= today}
                  onClick={() => {
                    const period = shiftPeriod(snapshot.periodStart, 1);
                    const endDate = period.endDate > today ? today : period.endDate;
                    const end = new Date(`${endDate}T00:00:00Z`);
                    const start = new Date(end);
                    start.setUTCDate(start.getUTCDate() - 29);
                    void onPeriodChange(start.toISOString().slice(0, 10), endDate);
                  }}
                />
              </div>
            </div>

            {activeTab === 'capacity' && (
              <div className={styles.boardGrid}>
                <section className={styles.monitor}>
                  <div className={styles.monitorHeader}><DatabaseRegular /> Capacity register</div>
                  {snapshot.capacities.length ? (
                    <div className={styles.tableScroll}>
                      <table className={styles.table}>
                        <thead>
                          <tr><th className={styles.th}>Capacity</th><th className={styles.th}>Status</th><th className={styles.th}>Available</th><th className={styles.th}>Consumed</th><th className={styles.th}>Updated</th></tr>
                        </thead>
                        <tbody>
                          {snapshot.capacities.map((item) => (
                            <tr key={`${item.capacityType}-${item.units}`}>
                              <td className={`${styles.td} ${styles.strong}`}>{labelize(item.capacityType)}<span className={styles.secondary}>{item.units || 'Units not specified'}</span></td>
                              <td className={styles.td}><Badge appearance="tint" color={badgeColor(item.status)}>{labelize(item.status)}</Badge></td>
                              <td className={`${styles.td} ${styles.numeric}`}>{formatNumber(item.totalCapacity)}</td>
                              <td className={`${styles.td} ${styles.numeric}`}>
                                {formatNumber(item.consumed)}
                                {item.ratedConsumption !== undefined && item.ratedConsumption > 0 && (
                                  <span
                                    className={styles.secondary}
                                    title="Service-adjusted consumption reported by Microsoft for capacity accounting."
                                  >
                                    Rated usage {formatNumber(item.ratedConsumption)}
                                  </span>
                                )}

                              </td>
                              <td className={styles.td}>{formatDate(item.updatedOn)}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  ) : <div className={styles.emptyInline}>Capacity details were not returned.</div>}
                </section>

                <section className={styles.monitor}>
                  <div className={styles.monitorHeader}><MoneyRegular /> Currency ledger</div>
                  {snapshot.currencies.length ? (
                    <div className={styles.ledgerGrid}>
                      {snapshot.currencies.map((item) => (
                        <article className={styles.ledgerCell} key={item.currencyType}>
                          <div className={styles.ledgerTitle}>
                            <Text className={styles.ledgerName}>
                              {CURRENCY_LABELS[item.currencyType] ?? labelize(item.currencyType)}
                            </Text>
                            {item.lastUpdatedDay && (
                              <Text className={styles.ledgerDate}>Updated {formatDate(item.lastUpdatedDay)}</Text>
                            )}
                          </div>
                          <div className={styles.ledgerValues}>
                            <span className={styles.ledgerValue}>
                              {formatNumber(item.purchased)}
                              <small className={styles.ledgerLabel}>Purchased</small>
                            </span>
                            <span className={styles.ledgerValue}>
                              {formatNumber(item.allocated)}
                              <small className={styles.ledgerLabel}>Allocated</small>
                            </span>
                            <span className={styles.ledgerValue}>
                              {formatNumber(item.consumed)}
                              <small className={styles.ledgerLabel}>Consumed</small>
                            </span>
                          </div>
                        </article>
                      ))}
                    </div>
                  ) : <div className={styles.emptyInline}>Currency reporting is not enabled for this tenant.</div>}
                </section>
              </div>
            )}

            {activeTab === 'harness' && (
              <div className={styles.boardGrid}>
                <MessageBar intent="info">
                  <MessageBarBody>
                    Classify each affected environment as maker development or funded production before changing capacity. Development needs a deliberate spending boundary; production controls should reflect accountable ownership, funding, expected usage, and service criticality.
                  </MessageBarBody>
                </MessageBar>
                <section className={styles.monitor}>
                  <div className={styles.monitorHeader}>
                    <BotRegular /> GitHub Copilot harness environments
                  </div>
                  <div className={styles.emptyInline}>
                    <Text className={styles.guidance}>
                      {harnessAgents.length} harness {harnessAgents.length === 1 ? 'agent' : 'agents'} detected through the inventory <code>isCLIAgent</code> property. Reserved credits and linked pay-as-you-go policies are read with Admin V2 connector actions.
                    </Text>
                  </div>
                  {harnessEnvironments.length ? (
                    <div className={styles.tableScroll}>
                      <table className={styles.table}>
                        <thead>
                          <tr>
                            <th className={styles.th}>Environment</th>
                            <th className={styles.th}>Purpose signal</th>
                            <th className={styles.th}>Agents</th>
                            <th className={styles.th}>Owners</th>
                            <th className={styles.th}>Reserved credits</th>
                            <th className={styles.th}>Pay-as-you-go</th>
                          </tr>
                        </thead>
                        <tbody>
                          {harnessEnvironments.map((item) => (
                            <tr key={item.environmentId}>
                              <td className={`${styles.td} ${styles.strong}`}>
                                {item.environmentName}
                                <span className={styles.secondary}>{item.environmentId}</span>
                              </td>
                              <td className={styles.td}>{labelize(item.environmentType)}</td>
                              <td className={`${styles.td} ${styles.numeric}`}>{item.agentCount}</td>
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
                                    onChange={(_, data) => setAllocationDrafts((current) => ({
                                      ...current,
                                      [item.environmentId]: data.value,
                                    }))}
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
                                <Badge
                                  appearance="tint"
                                  color={item.billingPolicyEnabled ? 'success' : 'subtle'}
                                  title={item.billingPolicyError}
                                >
                                  {item.billingPolicyName ?? (item.billingPolicyError ? 'Not linked or unavailable' : 'Disabled')}
                                </Badge>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  ) : (
                    <div className={styles.emptyInline}>
                      {isHarnessLoading
                        ? 'Loading harness environment controls…'
                        : 'No GitHub Copilot harness agents were found in the current inventory.'}
                    </div>
                  )}
                </section>
              </div>
            )}

            {activeTab === 'entitlements' && (
              <>
                <div className={styles.toolbar}>
                  <Input
                    contentBefore={<SearchRegular />}
                    placeholder="Search license, SKU, or capacity…"
                    value={search}
                    onChange={(_, data) => setSearch(data.value)}
                    style={{ minWidth: '280px' }}
                  />
                  <Text size={200}>{filteredEntitlements.length} of {snapshot.entitlements.length} entitlements</Text>
                </div>
                <section className={styles.monitor}>
                  {filteredEntitlements.length ? (
                    <div className={styles.tableScroll}>
                      <table className={styles.table}>
                        <thead>
                          <tr><th className={styles.th}>Entitlement</th><th className={styles.th}>Capacity</th><th className={styles.th}>Paid</th><th className={styles.th}>Trial</th><th className={styles.th}>Status</th><th className={styles.th}>Lifecycle</th></tr>
                        </thead>
                        <tbody>
                          {filteredEntitlements.map((item, index) => (
                            <tr key={`${item.entitlementCode}-${item.skuId ?? index}`}>
                              <td className={`${styles.td} ${styles.strong}`}>{item.displayName}<span className={styles.secondary}>{item.entitlementCode || item.skuId || 'No entitlement code'}</span></td>
                              <td className={styles.td}>{labelize(item.capacityType)}<span className={styles.secondary}>{labelize(item.capacitySubType)} · {formatNumber(item.totalCapacity)}</span></td>
                              <td className={`${styles.td} ${styles.numeric}`}>{formatNumber(item.paidEnabled)}</td>
                              <td className={`${styles.td} ${styles.numeric}`}>{formatNumber(item.trialEnabled)}</td>
                              <td className={styles.td}><Badge appearance="tint" color={badgeColor(item.capabilityStatus)}>{labelize(item.capabilityStatus)}</Badge>{item.isTemporary && <span className={styles.secondary}>Temporary</span>}</td>
                              <td className={styles.td}>{formatDate(item.expiryDate ?? item.nextLifecycleDate)}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  ) : (
                    <EmptyState
                      icon={<KeyRegular />}
                      title={search ? 'No entitlements match your search' : 'No license entitlements returned'}
                      subtitle={search ? 'Search by another license, SKU, or capacity family.' : 'The tenant capacity response did not include embedded license details.'}
                      action={search ? { label: 'Clear search', onClick: () => setSearch('') } : undefined}
                    />
                  )}
                </section>
              </>
            )}

            {activeTab === 'compliance' && (
              <div className={styles.boardGrid}>
                <section className={styles.monitor}>
                  <div className={styles.monitorHeader}><PersonWarningRegular /> Tenant compliance summary</div>
                  <div className={styles.tableScroll}>
                    {snapshot.compliance.length ? (
                      <div className={styles.summaryRows}>
                        {snapshot.compliance.map((item) => (
                          <div className={styles.summaryRow} key={item.flowContext}>
                            <Text className={styles.strong}>{labelize(item.flowContext)}</Text>
                            <span className={styles.summaryValue}><strong>{item.usersInCompliance}</strong><small>Compliant</small></span>
                            <span className={`${styles.summaryValue} ${item.usersExceedingCapacity ? styles.warning : ''}`}><strong>{item.usersExceedingCapacity}</strong><small>Over capacity</small></span>
                            <span className={`${styles.summaryValue} ${item.usersWithoutLicense ? styles.danger : ''}`}><strong>{item.usersWithoutLicense}</strong><small>No license</small></span>
                            <span className={`${styles.summaryValue} ${item.usersWithoutPremiumLicense ? styles.danger : ''}`}><strong>{item.usersWithoutPremiumLicense}</strong><small>Premium gap</small></span>
                          </div>
                        ))}
                      </div>
                    ) : <div className={styles.emptyInline}>No per-flow compliance summary was returned for this period.</div>}
                  </div>
                </section>

                <section className={styles.monitor}>
                  <div className={styles.monitorHeader}><PersonWarningRegular /> User utilization · first 200 records</div>
                  {snapshot.users.length ? (
                    <div className={styles.tableScroll}>
                      <table className={styles.table}>
                        <thead>
                          <tr><th className={styles.th}>User ID</th><th className={styles.th}>Context</th><th className={styles.th}>Consumption</th><th className={styles.th}>Flows</th></tr>
                        </thead>
                        <tbody>
                          {snapshot.users.map((item, index) => {
                            const overCapacity = item.totalCapacity > 0 && item.totalConsumption > item.totalCapacity;
                            return (
                              <tr key={`${item.userId}-${item.flowContext}-${index}`}>
                                <td className={`${styles.td} ${styles.strong}`}><span style={{ fontFamily: tokens.fontFamilyMonospace }}>{item.userId || 'Unknown user'}</span><span className={styles.secondary}>{labelize(item.licenseCategorization)}</span></td>
                                <td className={styles.td}>{labelize(item.flowContext)}</td>
                                <td className={`${styles.td} ${styles.numeric} ${overCapacity ? styles.warning : ''}`}>{formatNumber(item.totalConsumption)} / {formatNumber(item.totalCapacity)}</td>
                                <td className={`${styles.td} ${styles.numeric}`}>{item.totalFlows}</td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  ) : <div className={styles.emptyInline}>No user utilization records were returned for this period.</div>}
                </section>
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
}
