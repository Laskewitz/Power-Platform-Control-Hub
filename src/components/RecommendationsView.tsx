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
} from '@fluentui/react-components';
import {
  SearchRegular,
  LightbulbRegular,
  ClockRegular,
  BoxRegular,
  ListRegular,
} from '@fluentui/react-icons';
import type { AdvisorRecommendation } from '../types/admin.ts';
import RecommendationResourcesDialog from './RecommendationResourcesDialog.tsx';

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
    '@media (max-width: 480px)': {
      gridTemplateColumns: '1fr',
    },
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
    wordBreak: 'normal',
    overflowWrap: 'normal',
    hyphens: 'none',
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

function accentColor(resourceCount: number): string {
  if (resourceCount >= 20) return tokens.colorStatusDangerBackground3;
  if (resourceCount >= 5) return tokens.colorStatusWarningBackground3;
  return tokens.colorBrandBackground;
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
    return (
      <div className={styles.centered}>
        <Spinner size="extra-large" label="Loading advisor recommendations…" />
      </div>
    );
  }

  return (
    <div className={styles.root}>
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
                </div>
              </div>

              <div className={styles.cardActions}>
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
            </div>
          ))
        )}
      </div>

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
