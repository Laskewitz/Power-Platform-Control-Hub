import { useState, useEffect } from 'react';
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
} from '@fluentui/react-components';
import {
  ArrowLeftRegular,
  ArrowClockwiseRegular,
  AppGenericRegular,
  InfoRegular,
  ShieldCheckmarkRegular,
  ErrorCircleFilled,
  WarningFilled,
  InfoFilled,
  CheckmarkRegular,
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
import type { AnalysisResult, AnalysisSeverity } from '../services/flowAnalyzer.ts';
import { extractMessage } from '../utils/errorUtils.ts';
import AddSelfAsAdminBanner from './AddSelfAsAdminBanner.tsx';
import { formatDateTime } from '../utils/formatDate.ts';
import { formatSharedSummary } from '../utils/inventoryFormatters.ts';

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
    backgroundColor: tokens.colorNeutralBackground1,
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalXL}`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
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
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  envText: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
  },
  actionBar: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalXL}`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    flexShrink: 0,
  },
  body: {
    flex: 1,
    overflowY: 'auto',
    padding: `${tokens.spacingVerticalL} ${tokens.spacingHorizontalXL}`,
    width: '100%',
    boxSizing: 'border-box',
  },
  detailGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(2, 1fr)',
    gap: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalXL}`,
    alignItems: 'start',
    '@media (max-width: 768px)': {
      gridTemplateColumns: '1fr',
    },
  },
  detailItem: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: tokens.spacingVerticalXS,
    minWidth: 0,
  },
  detailItemWide: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: tokens.spacingVerticalXS,
    minWidth: 0,
    gridColumn: 'span 2',
  },
  detailLabel: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground3,
    textTransform: 'uppercase',
    letterSpacing: '0.04em',
  },
  detailValue: {
    fontSize: tokens.fontSizeBase300,
    color: tokens.colorNeutralForeground1,
  },
  dsItem: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalM}`,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground2,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    marginBottom: tokens.spacingVerticalXS,
  },
  emptyState: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    gap: tokens.spacingVerticalM,
    padding: `${tokens.spacingVerticalXXL} 0`,
    color: tokens.colorNeutralForeground3,
  },
  analysisList: {
    display: 'flex',
    flexDirection: 'column',
    gap: '0',
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    overflow: 'hidden',
  },
  analysisRow: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    cursor: 'pointer',
    userSelect: 'none' as const,
    backgroundColor: tokens.colorNeutralBackground1,
    ':hover': { backgroundColor: tokens.colorNeutralBackground1Hover },
    ':last-child': { borderBottom: 'none' },
  },
  analysisRowDetail: {
    padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalL}`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground2,
    display: 'flex',
    flexDirection: 'column' as const,
    gap: tokens.spacingVerticalS,
    ':last-child': { borderBottom: 'none' },
  },
  analysisRec: {
    fontSize: tokens.fontSizeBase300,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground3,
    borderLeft: `3px solid ${tokens.colorBrandStroke1}`,
  },
  analysisAffected: {
    display: 'flex',
    flexWrap: 'wrap' as const,
    gap: tokens.spacingHorizontalXS,
  },
  accordionCard: {
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    overflow: 'hidden',
    marginBottom: tokens.spacingVerticalS,
  },
  accordionHeaderTinted: {
    backgroundColor: tokens.colorNeutralBackground2,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
});

const SEV_ICON: Record<AnalysisSeverity, ReactElement> = {
  critical: <ErrorCircleFilled fontSize={16} style={{ color: tokens.colorStatusDangerForeground1, flexShrink: 0 }} />,
  warning: <WarningFilled fontSize={16} style={{ color: tokens.colorStatusWarningForeground1, flexShrink: 0 }} />,
  info: <InfoFilled fontSize={16} style={{ color: tokens.colorStatusSuccessForeground1, flexShrink: 0 }} />,
};
const SEV_LABEL: Record<AnalysisSeverity, string> = { critical: 'Error', warning: 'Warning', info: 'Info' };
const SEV_COLOR: Record<AnalysisSeverity, 'danger' | 'warning' | 'success'> = {
  critical: 'danger', warning: 'warning', info: 'success',
};

function AnalysisSection({ results }: { results: AnalysisResult[] }): ReactElement {
  const styles = useStyles();
  const [expandedIds, setExpandedIds] = useState<Set<string>>(new Set());

  const toggle = (id: string) => setExpandedIds(prev => {
    const next = new Set(prev);
    if (next.has(id)) next.delete(id); else next.add(id);
    return next;
  });

  if (results.length === 0) {
    return (
      <div className={styles.emptyState}>
        <CheckmarkRegular fontSize={40} style={{ color: tokens.colorStatusSuccessForeground1 }} />
        <Text style={{ fontWeight: tokens.fontWeightSemibold }}>All checks passed!</Text>
        <Text style={{ fontSize: tokens.fontSizeBase200, textAlign: 'center' }}>
          This canvas app appears to follow best practices.
        </Text>
      </div>
    );
  }

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM }}>
      <div className={styles.analysisList}>
        {results.map(r => {
          const expanded = expandedIds.has(r.id);
          const borderColor =
            r.severity === 'critical' ? tokens.colorStatusDangerForeground1
            : r.severity === 'warning' ? tokens.colorStatusWarningForeground1
            : tokens.colorStatusSuccessForeground1;
          return (
            <div key={r.id}>
              <div
                className={styles.analysisRow}
                style={{ borderLeft: `3px solid ${borderColor}`, backgroundColor: expanded ? tokens.colorNeutralBackground2 : undefined }}
                onClick={() => toggle(r.id)}
                role="button"
                tabIndex={0}
                aria-expanded={expanded}
                onKeyDown={e => (e.key === 'Enter' || e.key === ' ') && toggle(r.id)}
              >
                {SEV_ICON[r.severity]}
                <Text style={{ flex: 1, fontSize: tokens.fontSizeBase300, fontWeight: tokens.fontWeightSemibold }}>{r.title}</Text>
                <Badge appearance="filled" color={SEV_COLOR[r.severity]} size="small">{SEV_LABEL[r.severity]}</Badge>
                <span style={{ fontSize: '12px', color: tokens.colorNeutralForeground3, transform: expanded ? 'rotate(90deg)' : undefined, display: 'inline-flex' }}>▶</span>
              </div>
              {expanded && (
                <div className={styles.analysisRowDetail} style={{ borderLeft: `3px solid ${borderColor}` }}>
                  <Text style={{ fontSize: tokens.fontSizeBase300, color: tokens.colorNeutralForeground2 }}>{r.description}</Text>
                  <div className={styles.analysisRec}>
                    <Text style={{ fontSize: tokens.fontSizeBase200, fontWeight: tokens.fontWeightSemibold, color: tokens.colorNeutralForeground2 }}>💡 Recommendation</Text>
                    <Text style={{ fontSize: tokens.fontSizeBase300, display: 'block', marginTop: '4px' }}>{r.recommendation}</Text>
                  </div>
                  {r.affectedItems && r.affectedItems.length > 0 && (
                    <div>
                      <Text style={{ fontSize: tokens.fontSizeBase200, fontWeight: tokens.fontWeightSemibold, color: tokens.colorNeutralForeground2, display: 'block', marginBottom: '6px' }}>Affected items</Text>
                      <div className={styles.analysisAffected}>
                        {r.affectedItems.map(item => (
                          <Badge key={item} appearance="tint" color="informative" size="small">{item}</Badge>
                        ))}
                      </div>
                    </div>
                  )}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
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
            <Text className={styles.envText}>🌐 {resource.environmentName}</Text>
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
        <div style={{ marginLeft: 'auto', display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS }}>
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

        <Accordion multiple collapsible defaultOpenItems={['details', 'governance', 'roles', 'analysis']}>
          {/* ── App Details ── */}
          <AccordionItem value="details" className={styles.accordionCard}>
            <AccordionHeader expandIconPosition="end" icon={<InfoRegular />} className={styles.accordionHeaderTinted}>App Details</AccordionHeader>
            <AccordionPanel>
              <div style={{ padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM} ${tokens.spacingVerticalM}` }}>
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
                        {adminInfo.owner.displayName ?? adminInfo.owner.email ?? adminInfo.owner.id ?? '—'}
                        {adminInfo.owner.email && adminInfo.owner.displayName && (
                          <span style={{ color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200, display: 'block' }}>{adminInfo.owner.email}</span>
                        )}
                      </span>
                    </div>
                  )}

                  {adminInfo?.createdBy && (
                    <div className={styles.detailItem}>
                      <span className={styles.detailLabel}>Created by</span>
                      <span className={styles.detailValue}>{adminInfo.createdBy.displayName ?? adminInfo.createdBy.email ?? '—'}</span>
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
                      <span className={styles.detailValue}>{adminInfo.lastModifiedBy.displayName ?? adminInfo.lastModifiedBy.email ?? '—'}</span>
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
              <div style={{ padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalM} ${tokens.spacingVerticalL}` }}>
                {loading && !adminInfo && <Spinner size="small" label="Loading…" />}
                {adminInfo && (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM }}>
                    {/* Sharing stats */}
                    <div style={{ display: 'flex', gap: tokens.spacingHorizontalL, flexWrap: 'wrap' }}>
                      <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalXL}`, borderRadius: tokens.borderRadiusMedium, border: `1px solid ${tokens.colorNeutralStroke2}`, backgroundColor: tokens.colorNeutralBackground2, minWidth: '100px' }}>
                        <Text style={{ fontSize: tokens.fontSizeHero700, fontWeight: tokens.fontWeightBold, color: adminInfo.sharedUsersCount > 500 ? tokens.colorStatusDangerForeground1 : adminInfo.sharedUsersCount > 100 ? tokens.colorStatusWarningForeground1 : tokens.colorNeutralForeground1 }}>
                          {adminInfo.sharedUsersCount.toLocaleString()}
                        </Text>
                        <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>Users</Text>
                      </div>
                      <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalXL}`, borderRadius: tokens.borderRadiusMedium, border: `1px solid ${tokens.colorNeutralStroke2}`, backgroundColor: tokens.colorNeutralBackground2, minWidth: '100px' }}>
                        <Text style={{ fontSize: tokens.fontSizeHero700, fontWeight: tokens.fontWeightBold }}>
                          {adminInfo.sharedGroupsCount}
                        </Text>
                        <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>Groups</Text>
                      </div>
                      {adminInfo.bypassConsent && (
                        <div style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS, padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalM}`, borderRadius: tokens.borderRadiusMedium, border: `1px solid ${tokens.colorStatusWarningBorder1}`, backgroundColor: tokens.colorStatusWarningBackground1 }}>
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
                            {ra.principalDisplayName ?? ra.principalEmail ?? ra.principalId}
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

          {/* ── Best Practice Analysis ── */}
          <AccordionItem value="analysis" className={styles.accordionCard}>
            <AccordionHeader expandIconPosition="end" icon={<ShieldCheckmarkRegular />} className={styles.accordionHeaderTinted}>
              <span style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS }}>
                Best Practice Analysis
                {!loading && analysisResults.length > 0 && (() => {
                  const errs = analysisResults.filter(r => r.severity === 'critical').length;
                  const warns = analysisResults.filter(r => r.severity === 'warning').length;
                  const infs = analysisResults.filter(r => r.severity === 'info').length;
                  return (
                    <>
                      <Badge appearance="filled" color="subtle" size="small">{errs} error{errs !== 1 ? 's' : ''}</Badge>
                      <Badge appearance="filled" color="subtle" size="small">{warns} warning{warns !== 1 ? 's' : ''}</Badge>
                      <Badge appearance="filled" color="subtle" size="small">{infs} info</Badge>
                    </>
                  );
                })()}
              </span>
            </AccordionHeader>
            <AccordionPanel>
              <div style={{ padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalM} ${tokens.spacingVerticalL}` }}>
                {loading && <Spinner size="small" label="Analyzing app…" />}
                {!loading && adminInfo && <AnalysisSection results={analysisResults} />}
                {!loading && !adminInfo && (
                  <div className={styles.emptyState}>
                    <InfoRegular fontSize={32} />
                    <Text>No data loaded yet.</Text>
                  </div>
                )}
              </div>
            </AccordionPanel>
          </AccordionItem>

        </Accordion>
      </div>
    </div>
  );
}
