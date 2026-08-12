import { useState, useEffect, useMemo } from 'react';
import type { ReactElement } from 'react';
import {
  makeStyles,
  tokens,
  Text,
  Badge,
  Button,
  Spinner,
  Accordion,
  AccordionItem,
  AccordionHeader,
  AccordionPanel,
  MessageBar,
  MessageBarBody,
  Tooltip,
  Tab,
  TabList,
} from '@fluentui/react-components';
import {
  ArrowLeftRegular,
  ArrowClockwiseRegular,
  AppGenericRegular,
  InfoRegular,
  ShieldCheckmarkRegular,
  WarningFilled,
  PeopleRegular,
  PlugConnectedRegular,
  PersonRegular,
  LockClosedRegular,
  LockOpenRegular,
  AppsListRegular,
} from '@fluentui/react-icons';
import type { Resource } from '../types/inventory.ts';
import type { CanvasAppAdminInfo, AppRoleAssignment } from '../services/canvasAppAdminService.ts';
import { getCanvasAppAdminInfo, getAppRoleAssignments, setAppQuarantineState } from '../services/canvasAppAdminService.ts';
import { analyzeCanvasApp } from '../services/canvasAppAnalyzer.ts';
import type { AnalysisResult } from '../services/flowAnalyzer.ts';
import { extractMessage } from '../utils/errorUtils.ts';
import AddSelfAsAdminBanner from './AddSelfAsAdminBanner.tsx';
import { formatDateTime } from '../utils/formatDate.ts';
import { formatSharedSummary } from '../utils/inventoryFormatters.ts';
import AnalysisPosture from './AnalysisPosture.tsx';
import { useOwners } from '../services/ownerCache.ts';

interface Props {
  resource: Resource;
  onClose: () => void;
}

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
    overflow: 'hidden',
    color: '#E5EEF5',
    backgroundColor: '#060A0F',
    fontFamily: '"Aptos", "Segoe UI", sans-serif',
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    padding: `14px ${tokens.spacingHorizontalXL}`,
    borderBottom: '1px solid #29404F',
    backgroundColor: '#0C141D',
    boxShadow: 'inset 3px 0 0 #FFB547',
    flexShrink: 0,
    flexWrap: 'wrap',
  },
  headerMeta: {
    display: 'flex',
    flexDirection: 'column',
    gap: '2px',
    flex: 1,
    minWidth: 0,
  },
  titleRow: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    flexWrap: 'wrap',
  },
  title: {
    fontSize: '20px',
    fontWeight: tokens.fontWeightSemibold,
    color: '#F5FAFD',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  envText: {
    fontSize: '11px',
    color: '#9CB0BF',
    textTransform: 'uppercase',
    letterSpacing: '0.08em',
  },
  actionBar: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    padding: `8px ${tokens.spacingHorizontalXL}`,
    borderBottom: '1px solid #20313E',
    backgroundColor: '#111827',
    flexShrink: 0,
    flexWrap: 'wrap',
    '& button': {
      color: '#43D9FF',
      borderRadius: 0,
    },
  },
  actionRegister: {
    marginLeft: 'auto',
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    flexWrap: 'wrap',
  },
  contentTabs: {
    padding: `0 ${tokens.spacingHorizontalXL}`,
    borderBottom: '1px solid #29404F',
    backgroundColor: '#060A0F',
    flexShrink: 0,
  },
  body: {
    flex: 1,
    overflowY: 'auto',
    padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalXL}`,
    width: '100%',
    boxSizing: 'border-box',
  },
  detailGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(2, 1fr)',
    gap: 0,
    alignItems: 'start',
    borderTop: '1px solid #20313E',
    borderLeft: '1px solid #20313E',
    '@media (max-width: 768px)': {
      gridTemplateColumns: '1fr',
    },
  },
  detailItem: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: '4px',
    minWidth: 0,
    minHeight: '58px',
    padding: '10px 12px',
    borderRight: '1px solid #20313E',
    borderBottom: '1px solid #20313E',
    backgroundColor: '#0C141D',
  },
  detailItemWide: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: '4px',
    minWidth: 0,
    gridColumn: 'span 2',
    minHeight: '58px',
    padding: '10px 12px',
    borderRight: '1px solid #20313E',
    borderBottom: '1px solid #20313E',
    backgroundColor: '#0C141D',
    '@media (max-width: 768px)': {
      gridColumn: 'span 1',
    },
  },
  detailLabel: {
    fontSize: '10px',
    fontWeight: tokens.fontWeightSemibold,
    color: '#8DA5B5',
    textTransform: 'uppercase',
    letterSpacing: '0.1em',
  },
  detailValue: {
    fontSize: tokens.fontSizeBase300,
    color: '#E5EEF5',
    fontVariantNumeric: 'tabular-nums',
  },
  dsItem: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalM}`,
    backgroundColor: '#111827',
    borderBottom: '1px solid #20313E',
  },
  emptyState: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    gap: tokens.spacingVerticalM,
    padding: `${tokens.spacingVerticalXXL} 0`,
    color: '#9CB0BF',
    border: '1px solid #20313E',
    backgroundColor: '#0C141D',
  },
  analysisList: {
    display: 'flex',
    flexDirection: 'column',
    gap: '0',
    border: '1px solid #29404F',
    overflow: 'hidden',
  },
  analysisRow: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    borderBottom: '1px solid #20313E',
    cursor: 'pointer',
    userSelect: 'none' as const,
    backgroundColor: '#0C141D',
    ':hover': { backgroundColor: '#111827', boxShadow: 'inset 1px 0 0 #43D9FF' },
    ':last-child': { borderBottom: 'none' },
  },
  analysisRowDetail: {
    padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalL}`,
    borderBottom: '1px solid #20313E',
    backgroundColor: '#111827',
    display: 'flex',
    flexDirection: 'column' as const,
    gap: tokens.spacingVerticalS,
    ':last-child': { borderBottom: 'none' },
  },
  analysisRec: {
    fontSize: tokens.fontSizeBase300,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    backgroundColor: '#0C141D',
    border: '1px solid #29404F',
    borderLeft: '1px solid #43D9FF',
  },
  analysisAffected: {
    display: 'flex',
    flexWrap: 'wrap' as const,
    gap: tokens.spacingHorizontalXS,
  },
  accordionCard: {
    border: '1px solid #29404F',
    overflow: 'hidden',
    marginBottom: '8px',
    backgroundColor: '#0C141D',
  },
  accordionHeaderTinted: {
    minHeight: '42px',
    color: '#E5EEF5',
    backgroundColor: '#111827',
    borderBottom: '1px solid #20313E',
    ':hover': {
      color: '#43D9FF',
      backgroundColor: '#14212D',
    },
  },
  sectionBody: {
    padding: `10px ${tokens.spacingHorizontalM} ${tokens.spacingVerticalM}`,
  },
  stack: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
  },
  metricGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(2, minmax(100px, 160px)) minmax(180px, 1fr)',
    gap: 0,
    borderTop: '1px solid #29404F',
    borderLeft: '1px solid #29404F',
    '@media (max-width: 768px)': {
      gridTemplateColumns: '1fr 1fr',
    },
  },
  metricCell: {
    display: 'flex',
    flexDirection: 'column',
    justifyContent: 'center',
    padding: '12px 14px',
    minHeight: '72px',
    borderRight: '1px solid #29404F',
    borderBottom: '1px solid #29404F',
    backgroundColor: '#111827',
    fontVariantNumeric: 'tabular-nums',
  },
  attentionCell: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    padding: '12px 14px',
    borderRight: '1px solid #29404F',
    borderBottom: '1px solid #29404F',
    color: '#FFB547',
    backgroundColor: '#171A1D',
  },
});

function AnalysisSection({ results }: { results: AnalysisResult[] }): ReactElement {
  return (
    <AnalysisPosture
      results={results}
      title="Canvas app posture"
      description="Governance signals derived from app configuration, ownership, sharing, and connections."
      emptyDescription="This canvas app follows the evaluated governance practices."
    />
  );
}

export default function CanvasAppDetailPanel({ resource, onClose }: Props): ReactElement {
  const styles = useStyles();
  const displayName = resource.properties.displayName ?? resource.name;
  const envId = resource.properties.environmentId ?? '';
  const appId = resource.name;

  const isAppBuilderApp = resource.properties.subType === 'appBuilderApp';

  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [adminInfo, setAdminInfo] = useState<CanvasAppAdminInfo | null>(null);
  const [roleAssignments, setRoleAssignments] = useState<AppRoleAssignment[]>([]);
  const [quarantining, setQuarantining] = useState(false);
  const [analysisResults, setAnalysisResults] = useState<AnalysisResult[]>([]);
  const [contentTab, setContentTab] = useState<'overview' | 'analysis'>('overview');

  async function loadAnalysis() {
    setLoading(true);
    setError(null);
    setAdminInfo(null);
    setRoleAssignments([]);
    setAnalysisResults([]);
    try {
      const [adminResult, rolesResult] = await Promise.allSettled([
        getCanvasAppAdminInfo(envId, appId),
        getAppRoleAssignments(envId, appId),
      ]);

      const admin = adminResult.status === 'fulfilled' ? adminResult.value : null;

      if (adminResult.status === 'rejected') {
        setError(extractMessage((adminResult.reason as Error).message ?? 'Could not load admin data.'));
      }
      if (rolesResult.status === 'fulfilled') {
        setRoleAssignments(rolesResult.value);
      }

      setAdminInfo(admin);
      setAnalysisResults(analyzeCanvasApp(
        admin ?? undefined,
        rolesResult.status === 'fulfilled' ? rolesResult.value : undefined,
      ));
    } finally {
      setLoading(false);
    }
  }

  async function handleQuarantine(quarantine: boolean) {
    setQuarantining(true);
    try {
      await setAppQuarantineState(envId, appId, quarantine);
    } catch {
      // error is shown inline
    } finally {
      setQuarantining(false);
    }
  }

  useEffect(() => { void loadAnalysis(); }, [envId, appId]);

  const props = resource.properties;
  const principalIds = useMemo(() => [
    adminInfo?.owner?.id,
    adminInfo?.createdBy?.id,
    adminInfo?.lastModifiedBy?.id,
    ...roleAssignments.map((assignment) => assignment.principalId),
  ], [adminInfo, roleAssignments]);
  const principalNames = useOwners(principalIds, envId);
  const principalName = (id?: string): string | undefined => (
    id ? principalNames.get(id.toLowerCase()) : undefined
  );

  return (
    <div className={styles.root}>
      {/* Header */}
      <div className={styles.header}>
        <Button appearance="subtle" icon={<ArrowLeftRegular />} onClick={onClose} size="small">
          Back to Resources
        </Button>
        <AppGenericRegular fontSize={20} style={{ color: tokens.colorBrandForeground1, flexShrink: 0 }} />
        <div className={styles.headerMeta}>
          <div className={styles.titleRow}>
            <Tooltip content={displayName} relationship="label">
              <Text className={styles.title}>{displayName}</Text>
            </Tooltip>
            <Badge appearance="tint" color="brand" size="small" style={{ flexShrink: 0 }}>Canvas App</Badge>
            {props.isQuarantined && (
              <Badge appearance="tint" color="danger" size="small" icon={<LockClosedRegular />}>Quarantined</Badge>
            )}
          </div>
          {resource.environmentName && (
            <Text className={styles.envText}>{resource.environmentName}</Text>
          )}
        </div>
      </div>

      {/* Action bar */}
      <div className={styles.actionBar}>
        <Button
          appearance="subtle"
          icon={loading ? <Spinner size="tiny" /> : <ArrowClockwiseRegular />}
          disabled={loading}
          onClick={() => void loadAnalysis()}
          size="small"
        >
          {loading ? 'Analyzing…' : 'Re-analyze'}
        </Button>
        {!isAppBuilderApp && <AddSelfAsAdminBanner environmentId={envId} variant="inline" />}
        <div className={styles.actionRegister}>
          {adminInfo && (
            <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>
              Shared with {adminInfo.sharedUsersCount} user{adminInfo.sharedUsersCount !== 1 ? 's' : ''}
              {adminInfo.sharedGroupsCount > 0 ? ` + ${adminInfo.sharedGroupsCount} group${adminInfo.sharedGroupsCount !== 1 ? 's' : ''}` : ''}
            </Text>
          )}
          {!isAppBuilderApp && !props.isQuarantined && (
            <Button
              appearance="subtle"
              icon={quarantining ? <Spinner size="tiny" /> : <LockClosedRegular />}
              disabled={quarantining || loading}
              onClick={() => void handleQuarantine(true)}
              size="small"
            >
              Quarantine
            </Button>
          )}
          {!isAppBuilderApp && props.isQuarantined && (
            <Button
              appearance="subtle"
              icon={quarantining ? <Spinner size="tiny" /> : <LockOpenRegular />}
              disabled={quarantining || loading}
              onClick={() => void handleQuarantine(false)}
              size="small"
            >
              Unquarantine
            </Button>
          )}
        </div>
      </div>

      <div className={styles.contentTabs}>
        <TabList
          selectedValue={contentTab}
          onTabSelect={(_, data) => setContentTab(data.value as 'overview' | 'analysis')}
        >
          <Tab value="overview">Overview</Tab>
          <Tab value="analysis" icon={<ShieldCheckmarkRegular />}>Analysis</Tab>
        </TabList>
      </div>

      {/* Body */}
      <div className={styles.body}>
        {error && (() => {
          const isM365Blocked = error.includes('PermissionBlockedByM365CopilotApp');
          return (
            <MessageBar
              intent={isM365Blocked ? 'info' : 'warning'}
              style={{ marginBottom: tokens.spacingVerticalS }}
            >
              <MessageBarBody style={{ wordBreak: 'break-word' }}>
                {isM365Blocked
                  ? 'Admin data is restricted for AI Builder apps (appBuilderApp). This is a platform limitation — some details and sharing information are not available for this app type.'
                  : `Admin API: ${error}`}
              </MessageBarBody>
            </MessageBar>
          );
        })()}

        {contentTab === 'analysis' ? (
          <div className={styles.sectionBody}>
            {loading && <Spinner size="small" label="Analyzing app…" />}
            {!loading && adminInfo && <AnalysisSection results={analysisResults} />}
            {!loading && !adminInfo && (
              <div className={styles.emptyState}>
                <InfoRegular fontSize={32} />
                <Text>No data loaded yet.</Text>
              </div>
            )}
          </div>
        ) : (
        <Accordion multiple collapsible defaultOpenItems={['details', 'governance', 'roles']}>
          {/* ── App Details ── */}
          <AccordionItem value="details" className={styles.accordionCard}>
            <AccordionHeader expandIconPosition="end" icon={<InfoRegular />} className={styles.accordionHeaderTinted}>App Details</AccordionHeader>
            <AccordionPanel>
              <div className={styles.sectionBody}>
                {loading && !adminInfo && <Spinner size="small" />}
                <div className={styles.detailGrid}>
                  <div className={styles.detailItem}>
                    <span className={styles.detailLabel}>Name</span>
                    <span className={styles.detailValue}>{displayName}</span>
                  </div>

                  <div className={styles.detailItem}>
                    <span className={styles.detailLabel}>Environment</span>
                    <span className={styles.detailValue}>{resource.environmentName ?? '—'}</span>
                  </div>

                  {adminInfo?.owner && (
                    <div className={styles.detailItem}>
                      <span className={styles.detailLabel}>Owner</span>
                      <span className={styles.detailValue}>
                        {adminInfo.owner.displayName ?? adminInfo.owner.email ?? principalName(adminInfo.owner.id) ?? adminInfo.owner.id ?? '—'}
                        {adminInfo.owner.email && adminInfo.owner.displayName && (
                          <span style={{ color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200, display: 'block' }}>{adminInfo.owner.email}</span>
                        )}
                      </span>
                    </div>
                  )}

                  {adminInfo?.createdBy && (
                    <div className={styles.detailItem}>
                      <span className={styles.detailLabel}>Created by</span>
                      <span className={styles.detailValue}>{adminInfo.createdBy.displayName ?? adminInfo.createdBy.email ?? principalName(adminInfo.createdBy.id) ?? adminInfo.createdBy.id ?? '—'}</span>
                    </div>
                  )}

                  <div className={styles.detailItem}>
                    <span className={styles.detailLabel}>Created</span>
                    <span className={styles.detailValue}>
                      {formatDateTime(adminInfo?.createdTime ?? props.createdAt)}
                    </span>
                  </div>

                  <div className={styles.detailItem}>
                    <span className={styles.detailLabel}>Last Modified</span>
                    <span className={styles.detailValue}>
                      {formatDateTime(adminInfo?.lastModifiedTime ?? props.lastModifiedAt ?? props.modifiedAt)}
                    </span>
                  </div>

                  {adminInfo?.lastModifiedBy && (
                    <div className={styles.detailItem}>
                      <span className={styles.detailLabel}>Modified by</span>
                      <span className={styles.detailValue}>{adminInfo.lastModifiedBy.displayName ?? adminInfo.lastModifiedBy.email ?? principalName(adminInfo.lastModifiedBy.id) ?? adminInfo.lastModifiedBy.id ?? '—'}</span>
                    </div>
                  )}

                  {props.subType && (
                    <div className={styles.detailItem}>
                      <span className={styles.detailLabel}>Sub Type</span>
                      <span className={styles.detailValue}>{props.subType}</span>
                    </div>
                  )}

                  {props.isQuarantined !== undefined && (
                    <div className={styles.detailItem}>
                      <span className={styles.detailLabel}>Quarantine Status</span>
                      <span className={styles.detailValue}>
                        <Badge appearance="tint" color={props.isQuarantined ? 'danger' : 'success'} size="small">
                          {props.isQuarantined ? 'Quarantined' : 'Not Quarantined'}
                        </Badge>
                      </span>
                    </div>
                  )}

                  {adminInfo?.tags.primaryFormFactor && (
                    <div className={styles.detailItem}>
                      <span className={styles.detailLabel}>Form factor</span>
                      <span className={styles.detailValue}>
                        {adminInfo.tags.primaryFormFactor}
                        {adminInfo.tags.supportsPortrait === 'true' && ' · Portrait'}
                        {adminInfo.tags.supportsLandscape === 'true' && ' · Landscape'}
                      </span>
                    </div>
                  )}

                  {adminInfo?.appVersion && (
                    <div className={styles.detailItem}>
                      <span className={styles.detailLabel}>App version</span>
                      <span className={styles.detailValue} style={{ fontFamily: 'monospace', fontSize: tokens.fontSizeBase200 }}>{adminInfo.appVersion}</span>
                    </div>
                  )}

                  {adminInfo?.bypassConsent && (
                    <div className={styles.detailItem}>
                      <span className={styles.detailLabel}>Consent bypass</span>
                      <span className={styles.detailValue}>
                        <Badge appearance="filled" color="warning" size="small">Enabled</Badge>
                      </span>
                    </div>
                  )}

                  <div className={styles.detailItem}>
                    <span className={styles.detailLabel}>App ID</span>
                    <span className={styles.detailValue} style={{ fontFamily: 'monospace', fontSize: tokens.fontSizeBase200, wordBreak: 'break-all', color: tokens.colorNeutralForeground3 }}>{appId}</span>
                  </div>

                  {adminInfo?.description && (
                    <div className={styles.detailItemWide}>
                      <span className={styles.detailLabel}>Description</span>
                      <span className={styles.detailValue}>{adminInfo.description}</span>
                    </div>
                  )}
                </div>
              </div>
            </AccordionPanel>
          </AccordionItem>

          {/* ── Governance & Sharing ── */}
          <AccordionItem value="governance" className={styles.accordionCard}>
            <AccordionHeader expandIconPosition="end" icon={<PeopleRegular />} className={styles.accordionHeaderTinted}>
              Governance &amp; Sharing
            </AccordionHeader>
            <AccordionPanel>
              <div className={styles.sectionBody}>
                {loading && !adminInfo && <Spinner size="small" label="Loading…" />}
                {adminInfo && (
                  <div className={styles.stack}>
                    {/* Sharing stats */}
                    <div className={styles.metricGrid}>
                      <div className={styles.metricCell}>
                        <Text style={{ fontSize: tokens.fontSizeHero700, fontWeight: tokens.fontWeightBold, color: adminInfo.sharedUsersCount > 500 ? tokens.colorStatusDangerForeground1 : adminInfo.sharedUsersCount > 100 ? tokens.colorStatusWarningForeground1 : tokens.colorNeutralForeground1 }}>
                          {adminInfo.sharedUsersCount.toLocaleString()}
                        </Text>
                        <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>Users</Text>
                      </div>
                      <div className={styles.metricCell}>
                        <Text style={{ fontSize: tokens.fontSizeHero700, fontWeight: tokens.fontWeightBold }}>
                          {adminInfo.sharedGroupsCount}
                        </Text>
                        <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>Groups</Text>
                      </div>
                      {adminInfo.bypassConsent && (
                        <div className={styles.attentionCell}>
                          <WarningFilled fontSize={16} style={{ color: tokens.colorStatusWarningForeground1 }} />
                          <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorStatusWarningForeground3 }}>Consent bypass enabled</Text>
                        </div>
                      )}
                    </div>

                    {/* Connection references from admin API */}
                    {adminInfo.connectionReferences.length > 0 && (
                      <div>
                        <Text style={{ fontSize: tokens.fontSizeBase200, fontWeight: tokens.fontWeightSemibold, color: tokens.colorNeutralForeground3, textTransform: 'uppercase', letterSpacing: '0.04em', display: 'block', marginBottom: tokens.spacingVerticalS }}>
                          Active connections ({adminInfo.connectionReferences.length})
                        </Text>
                        {adminInfo.connectionReferences.map((conn, i) => (
                          <div key={conn.id ?? i} className={styles.dsItem}>
                            <PlugConnectedRegular fontSize={16} style={{ color: tokens.colorBrandForeground1, flexShrink: 0 }} />
                            <Text style={{ flex: 1, fontSize: tokens.fontSizeBase300 }}>{conn.displayName ?? conn.id}</Text>
                            {conn.apiTier && (
                              <Badge appearance="tint" color={conn.apiTier.toLowerCase() === 'premium' ? 'warning' : 'subtle'} size="small">
                                {conn.apiTier}
                              </Badge>
                            )}
                            {conn.isOnPremiseConnection && (
                              <Badge appearance="tint" color="informative" size="small">On-Premises</Badge>
                            )}
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                )}
                {!loading && !adminInfo && (
                  <div className={styles.emptyState}>
                    <PeopleRegular fontSize={32} />
                    <Text>Governance data unavailable.</Text>
                  </div>
                )}
              </div>
            </AccordionPanel>
          </AccordionItem>

          {/* ── Role Assignments ── */}
          <AccordionItem value="roles" className={styles.accordionCard}>
            <AccordionHeader expandIconPosition="end" icon={<PersonRegular />} className={styles.accordionHeaderTinted}>
              Role Assignments
              {roleAssignments.length > 0 && (
                <Badge appearance="tint" color="informative" size="small" style={{ marginLeft: tokens.spacingHorizontalS }}>
                  {roleAssignments.length}
                </Badge>
              )}
            </AccordionHeader>
            <AccordionPanel>
              <div style={{ padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalM} ${tokens.spacingVerticalL}` }}>
                {loading && <Spinner size="small" label="Loading…" />}
                {!loading && roleAssignments.length === 0 && (
                  <div className={styles.emptyState}>
                    <PersonRegular fontSize={32} />
                    <Text>No role assignments found or data unavailable.</Text>
                  </div>
                )}
                {!loading && roleAssignments.length > 0 && (
                  <div className={styles.analysisList}>
                    {roleAssignments.map(ra => (
                      <div key={ra.id} style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalM, padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`, borderBottom: `1px solid ${tokens.colorNeutralStroke2}` }}>
                        <PersonRegular fontSize={16} style={{ color: tokens.colorBrandForeground1, flexShrink: 0 }} />
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <Text style={{ fontWeight: tokens.fontWeightSemibold, fontSize: tokens.fontSizeBase300 }}>
                            {ra.principalDisplayName ?? ra.principalEmail ?? principalName(ra.principalId) ?? ra.principalId}
                          </Text>
                          {ra.principalEmail && ra.principalDisplayName && (
                            <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3, display: 'block' }}>{ra.principalEmail}</Text>
                          )}
                        </div>
                        <Badge
                          appearance="tint"
                          color={ra.roleName === 'Owner' ? 'warning' : ra.roleName === 'CanEdit' ? 'informative' : 'subtle'}
                          size="small"
                        >
                          {ra.roleName}
                        </Badge>
                        <Badge appearance="tint" color="subtle" size="small">{ra.principalType}</Badge>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </AccordionPanel>
          </AccordionItem>

          {/* ── Inventory ── */}
          {(() => {
            const props = resource.properties;
            const viewersSummary = props?.sharedWithViewers ? formatSharedSummary(props.sharedWithViewers) : null;
            const editorsSummary = props?.sharedWithEditors ? formatSharedSummary(props.sharedWithEditors) : null;
            const hasData = viewersSummary || editorsSummary || props?.lastPublishedAt || props?.createdIn;
            if (!hasData) return null;
            return (
              <AccordionItem value="inventory" className={styles.accordionCard}>
                <AccordionHeader expandIconPosition="end" icon={<AppsListRegular />} className={styles.accordionHeaderTinted}>Inventory</AccordionHeader>
                <AccordionPanel>
                  <div className={styles.detailGrid}>
                    {props?.createdIn && (
                      <div className={styles.detailItem}>
                        <Text size={200} weight="semibold">Created In</Text>
                        <Text size={300}>{props.createdIn}</Text>
                      </div>
                    )}
                    {props?.lastPublishedAt && (
                      <div className={styles.detailItem}>
                        <Text size={200} weight="semibold">Last Published</Text>
                        <Text size={300}>{formatDateTime(props.lastPublishedAt)}</Text>
                      </div>
                    )}
                    {viewersSummary && (
                      <div className={styles.detailItem}>
                        <Text size={200} weight="semibold">Viewers (Inventory)</Text>
                        <Text size={300}>{viewersSummary}</Text>
                      </div>
                    )}
                    {editorsSummary && (
                      <div className={styles.detailItem}>
                        <Text size={200} weight="semibold">Editors (Inventory)</Text>
                        <Text size={300}>{editorsSummary}</Text>
                      </div>
                    )}
                  </div>
                </AccordionPanel>
              </AccordionItem>
            );
          })()}

        </Accordion>
        )}
      </div>
    </div>
  );
}
