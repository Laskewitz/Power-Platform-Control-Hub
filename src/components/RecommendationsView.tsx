import { useMemo, useState } from 'react';
import type { ReactElement } from 'react';
import {
  Button,
  Input,
  MessageBar,
  MessageBarBody,
  Text,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import {
  SearchRegular,
  ClockRegular,
  ListRegular,
  LightbulbRegular,
} from '@fluentui/react-icons';
import type { AdvisorRecommendation } from '../types/admin.ts';
import RecommendationResourcesDialog from './RecommendationResourcesDialog.tsx';
import EmptyState from './EmptyState.tsx';
import { OperationsSkeleton, PageHeader } from './ui.tsx';

interface RecommendationsViewProps {
  recommendations: AdvisorRecommendation[];
  isLoading: boolean;
  error: string | null;
}

type Severity = 'high' | 'medium' | 'low';

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
  toolbar: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    flexShrink: 0,
    flexWrap: 'wrap',
  },
  count: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    whiteSpace: 'nowrap',
    marginLeft: 'auto',
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
  strip: {
    display: 'grid',
    gridTemplateColumns: '14px minmax(220px, 1.6fr) minmax(120px, 0.6fr) minmax(160px, 0.7fr) auto',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    minWidth: '760px',
    minHeight: '52px',
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    ':last-child': { borderBottom: 'none' },
    ':hover': { backgroundColor: tokens.colorNeutralBackground2 },
  },
  rail: {
    width: '8px',
    height: '8px',
    justifySelf: 'center',
    borderRadius: '50%',
    boxShadow: `0 0 0 2px ${tokens.colorNeutralBackground1}`,
  },
  scenarioCell: {
    display: 'flex',
    flexDirection: 'column',
    gap: '2px',
    minWidth: 0,
  },
  scenarioName: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  scenarioId: {
    fontSize: tokens.fontSizeBase100,
    fontFamily: tokens.fontFamilyMonospace,
    color: tokens.colorNeutralForeground3,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  countCell: {
    display: 'flex',
    flexDirection: 'column',
    gap: '2px',
  },
  countValue: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightBold,
    fontVariantNumeric: 'tabular-nums',
  },
  countLabel: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    textTransform: 'uppercase',
    letterSpacing: '0.04em',
  },
  refreshCell: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    whiteSpace: 'nowrap',
  },
  centered: {
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    height: '100%',
  },
});

function formatScenarioName(value: string): string {
  return value
    // Replace underscores/hyphens with spaces
    .replace(/[_-]/g, ' ')
    // Split lowercase prepositions embedded between words (e.g. "AccessforAgents")
    .replace(/([A-Za-z])(for|on|with|by|in|to|at|of)([A-Z])/g, '$1 $2 $3')
    // Standard camelCase split: lowercase→uppercase and UPPER→Uppercase
    .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
    .replace(/([A-Z])([A-Z][a-z])/g, '$1 $2')
    // Collapse multiple spaces
    .replace(/\s+/g, ' ')
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

function getSeverity(resourceCount: number): Severity {
  if (resourceCount >= 20) return 'high';
  if (resourceCount >= 5) return 'medium';
  return 'low';
}

function severityRailColor(severity: Severity): string {
  if (severity === 'high') return tokens.colorStatusDangerBackground3;
  if (severity === 'medium') return tokens.colorStatusWarningBackground3;
  return tokens.colorBrandBackground;
}

function severityTextColor(severity: Severity): string {
  if (severity === 'high') return tokens.colorStatusDangerForeground1;
  if (severity === 'medium') return tokens.colorStatusWarningForeground1;
  return tokens.colorBrandForeground1;
}

export default function RecommendationsView({
  recommendations,
  isLoading,
  error,
}: RecommendationsViewProps): ReactElement {
  const styles = useStyles();
  const [search, setSearch] = useState('');
  const [resourcesDialog, setResourcesDialog] = useState<{ scenario: string; name: string; actions: AdvisorRecommendation['details']['actions'] } | null>(null);

  const filteredRecommendations = useMemo(() => {
    const term = search.trim().toLowerCase();
    if (!term) return recommendations;
    return recommendations.filter((r) =>
      formatScenarioName(r.scenario).toLowerCase().includes(term) ||
      r.scenario.toLowerCase().includes(term),
    );
  }, [recommendations, search]);

  if (isLoading) {
    return <OperationsSkeleton />;
  }

  return (
    <div className={styles.root}>
      <PageHeader
        title="Recommendations"
        description="Advisor scenarios ranked by the number of resources they affect across the tenant."
      />

      <div className={styles.toolbar}>
        <Input
          placeholder="Search scenarios…"
          value={search}
          onChange={(_, data) => setSearch(data.value)}
          contentBefore={<SearchRegular />}
          size="small"
          style={{ minWidth: '220px' }}
        />
        <Text className={styles.count}>{filteredRecommendations.length} scenario{filteredRecommendations.length === 1 ? '' : 's'}</Text>
      </div>

      {error && (
        <MessageBar intent="error">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}

      {filteredRecommendations.length === 0 ? (
        <EmptyState
          icon={<LightbulbRegular />}
          title={search ? 'No scenarios match your search' : 'No advisor recommendations found'}
          subtitle={search ? 'Try a different search term.' : 'The tenant currently has no outstanding advisor scenarios.'}
          action={search ? { label: 'Clear search', onClick: () => setSearch('') } : undefined}
        />
      ) : (
        <div className={styles.board} role="table" aria-label="Advisor recommendations">
          {filteredRecommendations.map((rec) => {
            const severity = getSeverity(rec.details.resourceCount);
            return (
              <div key={rec.scenario} className={styles.strip} role="row">
                <span className={styles.rail} style={{ backgroundColor: severityRailColor(severity) }} aria-hidden="true" />
                <div className={styles.scenarioCell}>
                  <Text className={styles.scenarioName}>{formatScenarioName(rec.scenario)}</Text>
                  <Text className={styles.scenarioId}>{rec.scenario}</Text>
                </div>
                <div className={styles.countCell}>
                  <Text className={styles.countValue} style={{ color: severityTextColor(severity) }}>
                    {rec.details.resourceCount}
                  </Text>
                  <Text className={styles.countLabel}>
                    {rec.details.resourceCount === 1 ? 'resource' : 'resources'}
                  </Text>
                </div>
                <div className={styles.refreshCell}>
                  <ClockRegular fontSize={14} />
                  <span>{formatDate(rec.details.lastRefreshedTimestamp)}</span>
                </div>
                <Button
                  appearance="subtle"
                  size="small"
                  icon={<ListRegular />}
                  onClick={() => setResourcesDialog({
                    scenario: rec.scenario,
                    name: formatScenarioName(rec.scenario),
                    actions: rec.details.actions ?? [],
                  })}
                >
                  View resources
                </Button>
              </div>
            );
          })}
        </div>
      )}

      <RecommendationResourcesDialog
        open={resourcesDialog !== null}
        scenario={resourcesDialog?.scenario ?? ''}
        scenarioDisplayName={(resourcesDialog?.name ?? '').replace(/([^ ])for /gi, '$1 for ')}
        actions={resourcesDialog?.actions ?? []}
        onClose={() => setResourcesDialog(null)}
      />
    </div>
  );
}
