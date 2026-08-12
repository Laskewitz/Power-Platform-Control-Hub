import { useEffect, useState } from 'react';
import type { ReactElement } from 'react';
import { Badge, Spinner, Text, makeStyles, tokens } from '@fluentui/react-components';
import {
  CheckmarkCircleRegular,
  ChevronRightRegular,
  LightbulbRegular,
} from '@fluentui/react-icons';
import type { AnalysisResult, AnalysisSeverity } from '../services/flowAnalyzer.ts';
import EmptyState from './EmptyState.tsx';

interface AnalysisPostureProps {
  results: AnalysisResult[];
  title: string;
  description: string;
  isLoading?: boolean;
  loadingLabel?: string;
  unavailableMessage?: string;
  emptyTitle?: string;
  emptyDescription?: string;
  action?: {
    label: string;
    href: string;
  };
}

const SEVERITY_LABEL: Record<AnalysisSeverity, string> = {
  critical: 'Critical',
  warning: 'Attention',
  info: 'Guidance',
};

const SEVERITY_COLOR: Record<AnalysisSeverity, 'danger' | 'warning' | 'informative'> = {
  critical: 'danger',
  warning: 'warning',
  info: 'informative',
};

const useStyles = makeStyles({
  root: {
    display: 'grid',
    gap: '8px',
  },
  summary: {
    display: 'grid',
    gridTemplateColumns: 'minmax(240px, 1fr) repeat(3, minmax(92px, auto))',
    gap: '1px',
    border: '1px solid #29404F',
    backgroundColor: '#20313E',
    '@media (max-width: 720px)': {
      gridTemplateColumns: 'repeat(3, minmax(0, 1fr))',
    },
  },
  lead: {
    display: 'grid',
    gap: '4px',
    alignContent: 'center',
    minHeight: '82px',
    padding: '14px',
    backgroundColor: '#0C141D',
    '@media (max-width: 720px)': {
      gridColumn: '1 / -1',
    },
  },
  title: {
    color: '#F4FBFD',
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
  },
  description: {
    color: '#91A8B5',
    fontSize: tokens.fontSizeBase200,
    lineHeight: '18px',
  },
  metric: {
    display: 'grid',
    gap: '2px',
    alignContent: 'center',
    justifyItems: 'center',
    minHeight: '82px',
    padding: '12px',
    backgroundColor: '#0C141D',
  },
  metricValue: {
    color: '#F4FBFD',
    fontSize: '22px',
    lineHeight: '24px',
    fontWeight: tokens.fontWeightBold,
    fontVariantNumeric: 'tabular-nums',
  },
  metricLabel: {
    color: '#78909C',
    fontSize: '10px',
    fontWeight: tokens.fontWeightSemibold,
    letterSpacing: '0.08em',
    textTransform: 'uppercase',
  },
  row: {
    display: 'grid',
    gridTemplateColumns: '32px minmax(0, 1fr) auto 24px',
    alignItems: 'center',
    gap: '12px',
    width: '100%',
    minHeight: '60px',
    padding: '10px 12px',
    color: '#E5EEF5',
    backgroundColor: '#0C141D',
    borderTop: '1px solid #29404F',
    borderRight: '1px solid #29404F',
    borderBottom: '1px solid #29404F',
    borderLeft: '1px solid #29404F',
    font: 'inherit',
    textAlign: 'left',
    cursor: 'pointer',
    ':hover': { backgroundColor: '#101D27' },
    ':focus-visible': {
      outline: '2px solid #43D9FF',
      outlineOffset: '-2px',
    },
  },
  index: {
    color: '#78909C',
    fontSize: '11px',
    fontWeight: tokens.fontWeightSemibold,
    fontVariantNumeric: 'tabular-nums',
    textAlign: 'center',
  },
  copy: {
    display: 'grid',
    gap: '3px',
    minWidth: 0,
  },
  findingTitle: {
    color: '#F4FBFD',
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
  },
  findingDescription: {
    color: '#91A8B5',
    fontSize: tokens.fontSizeBase200,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  chevron: {
    color: '#91A8B5',
    fontSize: '14px',
    transitionProperty: 'transform',
    transitionDuration: '160ms',
    transitionTimingFunction: 'cubic-bezier(.16, 1, .3, 1)',
  },
  detail: {
    padding: '12px 14px 16px',
    backgroundColor: '#09121A',
    borderRight: '1px solid #29404F',
    borderBottom: '1px solid #29404F',
    borderLeft: '1px solid #29404F',
  },
  detailGrid: {
    display: 'grid',
    gridTemplateColumns: 'minmax(0, 0.8fr) minmax(0, 1.2fr)',
    gap: '1px',
    border: '1px solid #20313E',
    backgroundColor: '#20313E',
    '@media (max-width: 720px)': {
      gridTemplateColumns: '1fr',
    },
  },
  detailSection: {
    display: 'grid',
    alignContent: 'start',
    gap: '8px',
    padding: '14px',
    color: '#B6D2D8',
    backgroundColor: '#0C141D',
  },
  detailLabel: {
    color: '#78909C',
    fontSize: '10px',
    fontWeight: tokens.fontWeightSemibold,
    letterSpacing: '0.08em',
    textTransform: 'uppercase',
  },
  recommendationHeading: {
    display: 'flex',
    alignItems: 'center',
    gap: '7px',
    color: '#43D9FF',
  },
  affected: {
    display: 'flex',
    gap: '6px',
    flexWrap: 'wrap',
  },
  action: {
    width: 'fit-content',
    marginTop: '4px',
    color: '#62E6FF',
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    textDecorationLine: 'none',
    ':hover': { textDecorationLine: 'underline' },
  },
  loading: {
    display: 'grid',
    placeItems: 'center',
    minHeight: '180px',
    border: '1px solid #29404F',
    backgroundColor: '#0C141D',
  },
});

export default function AnalysisPosture({
  results,
  title,
  description,
  isLoading = false,
  loadingLabel = 'Evaluating governance posture…',
  unavailableMessage,
  emptyTitle = 'No issues found',
  emptyDescription = 'This resource follows the evaluated governance practices.',
  action,
}: AnalysisPostureProps): ReactElement {
  const styles = useStyles();
  const [expanded, setExpanded] = useState<Set<string>>(
    () => new Set(results.map((result) => result.id)),
  );
  const resultIds = results.map((result) => result.id).join('|');

  useEffect(() => {
    setExpanded(new Set(results.map((result) => result.id)));
  // Re-open the complete signal set only when the analysis result set changes.
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [resultIds]);

  if (isLoading) {
    return <div className={styles.loading}><Spinner size="small" label={loadingLabel} /></div>;
  }

  if (unavailableMessage) {
    return (
      <EmptyState
        icon={<LightbulbRegular />}
        title="Analysis unavailable"
        subtitle={unavailableMessage}
      />
    );
  }

  if (results.length === 0) {
    return (
      <EmptyState
        icon={<CheckmarkCircleRegular />}
        title={emptyTitle}
        subtitle={emptyDescription}
      />
    );
  }

  const counts = {
    critical: results.filter((result) => result.severity === 'critical').length,
    warning: results.filter((result) => result.severity === 'warning').length,
    info: results.filter((result) => result.severity === 'info').length,
  };

  return (
    <div className={styles.root}>
      <div className={styles.summary}>
        <div className={styles.lead}>
          <Text className={styles.title}>{title}</Text>
          <Text className={styles.description}>{description}</Text>
        </div>
        {(['critical', 'warning', 'info'] as const).map((severity) => (
          <div className={styles.metric} key={severity}>
            <Text className={styles.metricValue}>{counts[severity]}</Text>
            <Text className={styles.metricLabel}>{SEVERITY_LABEL[severity]}</Text>
          </div>
        ))}
      </div>

      {results.map((result, index) => {
        const isOpen = expanded.has(result.id);
        const detailId = `analysis-${result.id}`;
        const borderColor =
          result.severity === 'critical' ? tokens.colorStatusDangerForeground1
          : result.severity === 'warning' ? tokens.colorStatusWarningForeground1
          : tokens.colorBrandForeground1;
        return (
          <div key={result.id}>
            <button
              type="button"
              className={styles.row}
              style={{ borderLeftColor: borderColor }}
              aria-expanded={isOpen}
              aria-controls={detailId}
              onClick={() => setExpanded((current) => {
                const next = new Set(current);
                if (next.has(result.id)) next.delete(result.id); else next.add(result.id);
                return next;
              })}
            >
              <Text className={styles.index}>{String(index + 1).padStart(2, '0')}</Text>
              <span className={styles.copy}>
                <Text className={styles.findingTitle}>{result.title}</Text>
                <Text className={styles.findingDescription}>{result.description}</Text>
              </span>
              <Badge appearance="filled" color={SEVERITY_COLOR[result.severity]} size="small">
                {SEVERITY_LABEL[result.severity]}
              </Badge>
              <ChevronRightRegular
                className={styles.chevron}
                style={{ transform: isOpen ? 'rotate(90deg)' : undefined }}
                aria-hidden="true"
              />
            </button>
            {isOpen && (
              <div id={detailId} className={styles.detail} style={{ borderLeftColor: borderColor }}>
                <div className={styles.detailGrid}>
                  <div className={styles.detailSection}>
                    <Text className={styles.detailLabel}>Observed condition</Text>
                    <Text>{result.description}</Text>
                    {result.affectedItems?.length ? (
                      <div className={styles.affected}>
                        {result.affectedItems.map((item) => (
                          <Badge key={item} appearance="tint" color="informative" size="small">{item}</Badge>
                        ))}
                      </div>
                    ) : null}
                  </div>
                  <div className={styles.detailSection}>
                    <div className={styles.recommendationHeading}>
                      <LightbulbRegular aria-hidden="true" fontSize={16} />
                      <Text className={styles.detailLabel}>Recommended action</Text>
                    </div>
                    <Text>{result.recommendation}</Text>
                    {action && (
                      <a className={styles.action} href={action.href} target="_blank" rel="noopener noreferrer">
                        {action.label} ↗
                      </a>
                    )}
                  </div>
                </div>
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}
