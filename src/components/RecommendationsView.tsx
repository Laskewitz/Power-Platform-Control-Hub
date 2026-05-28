import { useMemo, useState } from 'react';
import type { ReactElement } from 'react';
import {
  Button,
  Input,
  MessageBar,
  MessageBarBody,
  Spinner,
  Text,
  makeStyles,
  shorthands,
  tokens,
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
} from '@fluentui/react-components';
import {
  SearchRegular,
  LightbulbRegular,
  ClockRegular,
  ArrowClockwiseRegular,
  BoxRegular,
  InfoRegular,
  DismissRegular,
  PlayRegular,
} from '@fluentui/react-icons';
import type { AdvisorRecommendation } from '../types/admin.ts';
import { useMutation } from '../hooks/useMutation.tsx';
import { PowerPlatformforAdminsV2Service } from '../generated/services/PowerPlatformforAdminsV2Service.ts';

const API_VERSION = '2024-10-01';

interface RecommendationsViewProps {
  recommendations: AdvisorRecommendation[];
  isLoading: boolean;
  error: string | null;
}

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalL,
    padding: tokens.spacingHorizontalXL,
    height: '100%',
    overflow: 'hidden',
  },
  notice: {
    display: 'flex',
    alignItems: 'flex-start',
    gap: tokens.spacingHorizontalS,
    backgroundColor: tokens.colorNeutralBackground3,
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
    borderRadius: tokens.borderRadiusMedium,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    flexShrink: 0,
  },
  noticeText: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    flex: 1,
  },
  toolbar: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    flexShrink: 0,
  },
  title: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
    flex: 1,
  },
  count: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    whiteSpace: 'nowrap',
  },
  grid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fill, minmax(300px, 1fr))',
    gap: tokens.spacingHorizontalL,
    overflowY: 'auto',
    flex: 1,
    paddingBottom: tokens.spacingVerticalL,
    alignContent: 'start',
    alignItems: 'start',
  },
  card: {
    display: 'flex',
    flexDirection: 'column',
    gap: '0',
    overflow: 'hidden',
    borderRadius: tokens.borderRadiusMedium,
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
    backgroundColor: tokens.colorNeutralBackground1,
    ':hover': { boxShadow: tokens.shadow8 },
  },
  cardAccent: {
    height: '4px',
    flexShrink: 0,
  },
  cardBody: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
    padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalM}`,
  },
  cardTitle: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    lineHeight: tokens.lineHeightBase400,
  },
  cardCount: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    backgroundColor: tokens.colorBrandBackground2,
    borderRadius: tokens.borderRadiusMedium,
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalM}`,
    width: 'fit-content',
  },
  cardCountNum: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightBold,
    color: tokens.colorBrandForeground1,
  },
  cardCountLabel: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorBrandForeground2,
  },
  cardDates: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXS,
    borderTopWidth: '1px',
    borderTopStyle: 'solid',
    borderTopColor: tokens.colorNeutralStroke2,
    paddingTop: tokens.spacingVerticalS,
  },
  dateRow: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
  },
  dateValue: {
    marginLeft: 'auto',
    color: tokens.colorNeutralForeground2,
    fontWeight: tokens.fontWeightMedium,
  },
  cardActions: {
    borderTopWidth: '1px',
    borderTopStyle: 'solid',
    borderTopColor: tokens.colorNeutralStroke2,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    display: 'flex',
    gap: tokens.spacingHorizontalS,
    flexWrap: 'wrap',
  },
  centered: {
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    height: '100%',
  },
  emptyState: {
    color: tokens.colorNeutralForeground3,
  },
  resultList: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXS,
    maxHeight: '300px',
    overflowY: 'auto',
    fontSize: tokens.fontSizeBase200,
  },
  resultRow: {
    display: 'flex',
    justifyContent: 'space-between',
    padding: `${tokens.spacingVerticalXS} 0`,
    borderBottomWidth: '1px',
    borderBottomStyle: 'solid',
    borderBottomColor: tokens.colorNeutralStroke2,
  },
});

function formatScenarioName(value: string): string {
  return value
    .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
    .replace(/([A-Z])([A-Z][a-z])/g, '$1 $2')
    .trim();
}

function formatDate(value?: string): string {
  if (!value) return '—';
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? value : date.toLocaleString(undefined, {
    month: 'short', day: 'numeric', year: 'numeric',
    hour: '2-digit', minute: '2-digit',
  });
}

function accentColor(resourceCount: number): string {
  if (resourceCount >= 20) return tokens.colorStatusDangerBackground3;
  if (resourceCount >= 5) return tokens.colorStatusWarningBackground3;
  return tokens.colorBrandBackground;
}

async function executeAction(scenario: string, actionName: string) {
  const result = await PowerPlatformforAdminsV2Service.ExecuteRecommendationAction(
    { scenario, actionParameters: {} },
    actionName,
    API_VERSION,
  );
  if (!result.success || result.error) throw new Error(result.error?.message ?? 'Action failed');
  return result.data;
}

interface ActionConfirm {
  scenario: string;
  actionName: string;
  displayName: string;
  resourceCount: number;
}

export default function RecommendationsView({
  recommendations,
  isLoading,
  error,
}: RecommendationsViewProps): ReactElement {
  const styles = useStyles();
  const [search, setSearch] = useState('');
  const [noticeVisible, setNoticeVisible] = useState(true);
  const [actionConfirm, setActionConfirm] = useState<ActionConfirm | null>(null);
  const [actionResult, setActionResult] = useState<{ succeeded: number; failed: number } | null>(null);

  const { execute: execAction, isLoading: isActionLoading } = useMutation(
    (scenario: string, actionName: string) => executeAction(scenario, actionName),
    {
      successMessage: 'Recommendation action submitted successfully.',
      onSuccess: (data) => {
        const results = (data as { results?: { statusCode?: number }[] })?.results ?? [];
        const succeeded = results.filter((r) => (r.statusCode ?? 200) < 300).length;
        const failed = results.length - succeeded;
        setActionResult({ succeeded, failed });
        setActionConfirm(null);
      },
    },
  );

  const filteredRecommendations = useMemo(() => {
    const term = search.trim().toLowerCase();
    if (!term) return recommendations;
    return recommendations.filter((r) =>
      formatScenarioName(r.scenario).toLowerCase().includes(term) ||
      r.scenario.toLowerCase().includes(term),
    );
  }, [recommendations, search]);

  if (isLoading) {
    return (
      <div className={styles.centered}>
        <Spinner size="extra-large" label="Loading advisor recommendations…" />
      </div>
    );
  }

  return (
    <div className={styles.root}>
      {noticeVisible && (
        <div className={styles.notice}>
          <InfoRegular style={{ fontSize: '0.9rem', flexShrink: 0, marginTop: '1px', color: tokens.colorNeutralForeground3 }} />
          <Text className={styles.noticeText}>
            Some permissions (Monitor.Alerts, Monitor.Metrics, Security.Recommendations, CopilotGovernance) are registered but their API endpoints are not yet publicly available.
          </Text>
          <Button
            appearance="subtle"
            size="small"
            icon={<DismissRegular />}
            onClick={() => setNoticeVisible(false)}
            style={{ flexShrink: 0 }}
          />
        </div>
      )}

      <div className={styles.toolbar}>
        <LightbulbRegular style={{ fontSize: '1.25rem', color: tokens.colorBrandForeground1 }} />
        <Text className={styles.title}>Recommendations</Text>
        <Input
          placeholder="Search scenarios…"
          value={search}
          onChange={(_, data) => setSearch(data.value)}
          contentBefore={<SearchRegular />}
          size="small"
          style={{ minWidth: '220px' }}
        />
        <Text className={styles.count}>{filteredRecommendations.length} scenario(s)</Text>
      </div>

      {error && (
        <MessageBar intent="error">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}

      <div className={styles.grid}>
        {filteredRecommendations.length === 0 ? (
          <Text className={styles.emptyState}>No advisor recommendations found.</Text>
        ) : (
          filteredRecommendations.map((rec) => (
            <div key={rec.scenario} className={styles.card}>
              <div className={styles.cardAccent} style={{ backgroundColor: accentColor(rec.details.resourceCount) }} />
              <div className={styles.cardBody}>
                <Text className={styles.cardTitle}>{formatScenarioName(rec.scenario)}</Text>

                <div className={styles.cardCount}>
                  <BoxRegular style={{ fontSize: '1rem', color: tokens.colorBrandForeground1 }} />
                  <Text className={styles.cardCountNum}>{rec.details.resourceCount}</Text>
                  <Text className={styles.cardCountLabel}>
                    {rec.details.resourceCount === 1 ? 'resource affected' : 'resources affected'}
                  </Text>
                </div>

                <div className={styles.cardDates}>
                  <div className={styles.dateRow}>
                    <ClockRegular style={{ fontSize: '0.85rem' }} />
                    <Text style={{ fontSize: tokens.fontSizeBase200 }}>Last refreshed</Text>
                    <Text className={styles.dateValue}>{formatDate(rec.details.lastRefreshedTimestamp)}</Text>
                  </div>
                  <div className={styles.dateRow}>
                    <ArrowClockwiseRegular style={{ fontSize: '0.85rem' }} />
                    <Text style={{ fontSize: tokens.fontSizeBase200 }}>Next refresh</Text>
                    <Text className={styles.dateValue}>{formatDate(rec.details.expectedNextRefreshTimestamp)}</Text>
                  </div>
                </div>
              </div>

              {rec.details.actions && rec.details.actions.length > 0 && (
                <div className={styles.cardActions}>
                  {rec.details.actions.map((action) => (
                    <Button
                      key={action.actionType ?? action.actionName}
                      appearance="outline"
                      size="small"
                      icon={<PlayRegular />}
                      onClick={() => setActionConfirm({
                        scenario: rec.scenario,
                        actionName: action.actionType ?? action.actionName ?? '',
                        displayName: action.actionName ?? action.actionType ?? 'Action',
                        resourceCount: rec.details.resourceCount,
                      })}
                    >
                      {action.actionName ?? action.actionType}
                    </Button>
                  ))}
                </div>
              )}
            </div>
          ))
        )}
      </div>

      {/* Confirm execute action */}
      <Dialog open={actionConfirm !== null} onOpenChange={(_, d) => { if (!d.open) setActionConfirm(null); }}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>Execute Recommendation Action</DialogTitle>
            <DialogContent>
              <Text>
                Run <strong>{actionConfirm?.displayName}</strong> on all{' '}
                <strong>{actionConfirm?.resourceCount}</strong> resource(s) affected by{' '}
                <strong>{actionConfirm ? formatScenarioName(actionConfirm.scenario) : ''}</strong>?
              </Text>
              <Text style={{ display: 'block', marginTop: tokens.spacingVerticalS, fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>
                This action will be applied to every resource in this recommendation.
              </Text>
            </DialogContent>
            <DialogActions>
              <Button appearance="secondary" disabled={isActionLoading} onClick={() => setActionConfirm(null)}>Cancel</Button>
              <Button
                appearance="primary"
                disabled={isActionLoading}
                icon={isActionLoading ? <Spinner size="tiny" /> : <PlayRegular />}
                onClick={() => {
                  if (actionConfirm) {
                    void execAction(actionConfirm.scenario, actionConfirm.actionName);
                  }
                }}
              >
                {isActionLoading ? 'Executing…' : 'Execute'}
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>

      {/* Action result summary */}
      <Dialog open={actionResult !== null} onOpenChange={(_, d) => { if (!d.open) setActionResult(null); }}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>Action Results</DialogTitle>
            <DialogContent>
              <div className={styles.resultList}>
                <div className={styles.resultRow}>
                  <Text>✅ Succeeded</Text>
                  <Text style={{ fontWeight: tokens.fontWeightSemibold }}>{actionResult?.succeeded ?? 0}</Text>
                </div>
                <div className={styles.resultRow}>
                  <Text>❌ Failed</Text>
                  <Text style={{ fontWeight: tokens.fontWeightSemibold }}>{actionResult?.failed ?? 0}</Text>
                </div>
              </div>
            </DialogContent>
            <DialogActions>
              <Button appearance="primary" onClick={() => setActionResult(null)}>Close</Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </div>
  );
}
