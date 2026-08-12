import { useMemo } from 'react';
import type { ReactElement } from 'react';
import {
  makeStyles,
  tokens,
  Text,
  MessageBar,
  MessageBarBody,
  Button,
} from '@fluentui/react-components';
import { extractMessage } from '../utils/errorUtils.ts';
import {
  GridRegular,
  DatabaseRegular,
  ArrowRepeatAllRegular,
  BotRegular,
  BoxRegular,
  CodeRegular,
  SettingsRegular,
  GlobeRegular,
  LightbulbRegular,
  ArrowRightRegular,
} from '@fluentui/react-icons';
import type { Resource, ResourceCounts } from '../types/inventory.ts';
import type { AdvisorRecommendation } from '../types/admin.ts';
import { OperationsSkeleton } from './ui.tsx';

interface DashboardProps {
  resources: Resource[];
  counts: ResourceCounts | null;
  environmentsCount: number;
  recommendations: AdvisorRecommendation[];
  isLoading: boolean;
  error: string | null;
  onNavigateToResources: (typeKey: string) => void;
  onNavigateToRecommendations: () => void;
}

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
    overflow: 'hidden',
  },
  scrollable: {
    flex: 1,
    minHeight: 0,
    display: 'flex',
    flexDirection: 'column',
    overflow: 'hidden',
  },
  topSection: {
    flex: 1,
    minHeight: 0,
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalL,
    padding: tokens.spacingHorizontalXL,
    overflowY: 'auto',
    '@media (max-width: 768px)': {
      padding: tokens.spacingHorizontalM,
    },
  },
  tableSection: {
    flex: 1,
    minHeight: 0,
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
    padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalXL} ${tokens.spacingHorizontalXL}`,
    '@media (max-width: 768px)': {
      padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalM} ${tokens.spacingHorizontalM}`,
    },
  },
  sectionTitle: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
  },
  operationsHero: {
    position: 'relative',
    display: 'grid',
    gridTemplateColumns: 'minmax(0, 1fr) 360px',
    minHeight: '220px',
    overflow: 'hidden',
    color: '#F4FBFD',
    border: '1px solid #29404F',
    backgroundColor: '#09121A',
    backgroundImage: 'linear-gradient(90deg, rgba(67, 217, 255, 0.045) 1px, transparent 1px)',
    backgroundSize: '48px 100%',
    boxShadow: '0 12px 30px rgba(0, 0, 0, 0.32)',
    '@media (max-width: 840px)': {
      gridTemplateColumns: '1fr',
    },
  },
  heroCopy: {
    position: 'relative',
    zIndex: 1,
    display: 'flex',
    flexDirection: 'column',
    justifyContent: 'center',
    gap: '16px',
    padding: '34px 40px',
  },
  heroTitle: {
    maxWidth: '640px',
    color: '#FFFFFF',
    fontSize: '54px',
    lineHeight: '54px',
    fontWeight: tokens.fontWeightSemibold,
    letterSpacing: '-0.045em',
    textWrap: 'balance',
    '@media (max-width: 620px)': {
      fontSize: '36px',
      lineHeight: '39px',
    },
  },
  heroDescription: {
    maxWidth: '58ch',
    color: '#9BB0BC',
    fontSize: tokens.fontSizeBase400,
    lineHeight: '24px',
  },
  heroRegister: {
    display: 'flex',
    alignItems: 'center',
    gap: '10px',
    flexWrap: 'wrap',
  },
  registerItem: {
    padding: '5px 9px',
    color: '#A9F0FF',
    border: '1px solid #2C5A69',
    backgroundColor: '#0B1D26',
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: '10px',
    letterSpacing: '0.06em',
    textTransform: 'uppercase',
  },
  signalMap: {
    position: 'relative',
    display: 'grid',
    placeItems: 'center',
    minHeight: '220px',
    overflow: 'hidden',
    borderLeft: '1px solid #29404F',
    backgroundColor: '#0D1821',
    backgroundImage: 'linear-gradient(rgba(67, 217, 255, 0.05) 1px, transparent 1px), linear-gradient(90deg, rgba(67, 217, 255, 0.05) 1px, transparent 1px)',
    backgroundSize: '24px 24px',
    '@media (max-width: 840px)': {
      display: 'none',
    },
  },
  signalCore: {
    display: 'grid',
    placeItems: 'center',
    width: '168px',
    height: '124px',
    borderTop: '1px solid #43D9FF',
    borderBottom: '1px solid #43D9FF',
  },
  signalValue: {
    color: '#43D9FF',
    fontSize: '72px',
    lineHeight: 1,
    fontWeight: tokens.fontWeightSemibold,
    letterSpacing: '-0.04em',
    fontVariantNumeric: 'tabular-nums',
  },
  signalLabel: {
    color: '#8FB2BF',
    fontSize: '9px',
    letterSpacing: '0.1em',
    textTransform: 'uppercase',
  },
  orbitNode: {
    position: 'absolute',
    display: 'grid',
    gap: '1px',
    minWidth: '88px',
    padding: '7px 9px',
    color: '#FFCF87',
    border: '1px solid #6B512C',
    backgroundColor: '#1A1710',
  },
  orbitTop: {
    top: '22px',
    right: '18px',
  },
  orbitBottom: {
    bottom: '22px',
    left: '18px',
  },
  orbitValue: {
    color: '#FFB547',
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    fontVariantNumeric: 'tabular-nums',
  },
  orbitLabel: {
    color: '#BFA77E',
    fontSize: '9px',
    letterSpacing: '0.08em',
    textTransform: 'uppercase',
  },
  operationsGrid: {
    display: 'grid',
    gridTemplateColumns: 'minmax(0, 1.6fr) minmax(300px, 0.8fr)',
    gap: tokens.spacingHorizontalL,
    '@media (max-width: 980px)': {
      gridTemplateColumns: '1fr',
    },
  },
  movementBoard: {
    display: 'grid',
    gridTemplateColumns: 'minmax(180px, 0.45fr) minmax(0, 1.55fr)',
    minHeight: '190px',
    overflow: 'hidden',
    padding: 0,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    backgroundColor: tokens.colorNeutralBackground1,
    boxShadow: '7px 9px 24px rgba(5, 36, 49, 0.18), 0 0 0 1px rgba(117, 224, 209, 0.12)',
    '@media (max-width: 680px)': {
      gridTemplateColumns: '1fr',
    },
  },
  tenantSummary: {
    display: 'flex',
    flexDirection: 'column',
    justifyContent: 'space-between',
    gap: tokens.spacingVerticalL,
    padding: tokens.spacingHorizontalL,
    color: '#EAF8FB',
    backgroundColor: '#0C2631',
    backgroundImage: 'none',
  },
  summaryLabel: {
    color: '#A9BDC8',
    fontSize: tokens.fontSizeBase100,
    fontWeight: tokens.fontWeightSemibold,
    letterSpacing: '0.08em',
    textTransform: 'uppercase',
  },
  resourceTotal: {
    display: 'block',
    color: '#FFFFFF',
    fontSize: '52px',
    lineHeight: 1,
    fontWeight: tokens.fontWeightSemibold,
    letterSpacing: '-0.04em',
  },
  summaryMeta: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    color: '#C9D7DF',
    fontSize: tokens.fontSizeBase200,
  },
  cardsGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(2, minmax(0, 1fr))',
    alignContent: 'stretch',
    gap: 0,
    backgroundColor: tokens.colorNeutralBackground1,
    '@media (max-width: 480px)': {
      gridTemplateColumns: '1fr',
    },
  },
  metricCard: {
    display: 'grid',
    gridTemplateColumns: '24px minmax(0, 1fr) auto',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    minHeight: '46px',
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    color: tokens.colorNeutralForeground1,
    textAlign: 'left',
    font: 'inherit',
    border: 0,
    borderRight: `1px solid ${tokens.colorNeutralStroke2}`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
    cursor: 'pointer',
    transitionProperty: 'color, background-color, box-shadow, transform',
    transitionDuration: '180ms',
    transitionTimingFunction: 'cubic-bezier(.16, 1, .3, 1)',
    ':hover': {
      zIndex: 1,
      color: tokens.colorBrandForeground1,
      backgroundColor: tokens.colorBrandBackground2,
      boxShadow: `inset 0 0 0 1px ${tokens.colorBrandStroke1}`,
      transform: 'translateY(-2px)',
    },
  },
  metricIcon: {
    display: 'grid',
    placeItems: 'center',
    fontSize: '18px',
    color: tokens.colorBrandForeground1,
  },
  metricCount: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightBold,
    lineHeight: tokens.lineHeightBase400,
    color: tokens.colorNeutralForeground1,
  },
  metricLabel: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },
  advisoryBoard: {
    display: 'flex',
    flexDirection: 'column',
    minHeight: '190px',
    padding: 0,
    overflow: 'hidden',
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    backgroundColor: tokens.colorNeutralBackground1,
    boxShadow: '7px 9px 24px rgba(5, 36, 49, 0.16)',
  },
  advisoryHeader: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    minHeight: '44px',
    padding: `0 ${tokens.spacingHorizontalM}`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    color: '#081018',
    backgroundColor: '#FFB547',
    backgroundImage: 'none',
  },
  advisoryCount: {
    marginLeft: 'auto',
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
  advisoryList: {
    display: 'grid',
    flex: 1,
  },
  advisoryRow: {
    display: 'grid',
    gridTemplateColumns: 'minmax(0, 1fr) auto',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    minHeight: '46px',
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  advisoryName: {
    overflow: 'hidden',
    color: tokens.colorNeutralForeground1,
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    whiteSpace: 'nowrap',
    textOverflow: 'ellipsis',
  },
  affected: {
    color: tokens.colorStatusWarningForeground1,
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    fontVariantNumeric: 'tabular-nums',
  },
  advisoryAction: {
    justifyContent: 'flex-start',
    borderRadius: 0,
  },
  tableCard: {
    flex: 1,
    minHeight: 0,
    overflow: 'hidden',
    overflowX: 'auto',
    overflowY: 'auto',
    padding: '0',
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    boxShadow: tokens.shadow4,
  },
  table: {
    width: '100%',
    borderCollapse: 'collapse',
    tableLayout: 'fixed',
    minWidth: '580px',
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
  },
  thSortable: {
    cursor: 'pointer',
    userSelect: 'none',
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground3Hover,
      color: tokens.colorNeutralForeground2,
    },
  },
  thInner: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
  },
  td: {
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    fontSize: tokens.fontSizeBase300,
    color: tokens.colorNeutralForeground1,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    verticalAlign: 'middle',
    overflow: 'hidden',
    whiteSpace: 'nowrap',
    textOverflow: 'ellipsis',
  },
  tdLast: {
    borderBottom: 'none',
  },
  sentinel: {
    height: '1px',
  },
  loadingMore: {
    display: 'flex',
    justifyContent: 'center',
    padding: tokens.spacingVerticalM,
  },
  centered: {
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    height: '100%',
  },
  toolbar: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    flexWrap: 'wrap',
    flexShrink: 0,
  },
  count: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    whiteSpace: 'nowrap',
    marginLeft: 'auto',
  },
  tdText: {
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
    display: 'block',
  },
});

const METRIC_ITEMS = [
  {
    key: 'canvasApps' as const,
    typeKey: 'microsoft.powerapps/canvasapps',
    label: 'Canvas Apps',
    icon: <GridRegular />,
  },
  {
    key: 'modelDrivenApps' as const,
    typeKey: 'microsoft.powerapps/modeldrivenapps',
    label: 'Model-Driven Apps',
    icon: <DatabaseRegular />,
  },
  {
    key: 'cloudFlows' as const,
    typeKey: 'microsoft.powerautomate/cloudflows',
    label: 'Cloud Flows',
    icon: <ArrowRepeatAllRegular />,
  },
  {
    key: 'agents' as const,
    typeKey: 'microsoft.copilotstudio/agents',
    label: 'Agents',
    icon: <BotRegular />,
  },
  {
    key: 'agentFlows' as const,
    typeKey: 'microsoft.powerautomate/agentflows',
    label: 'Agent Flows',
    icon: <ArrowRepeatAllRegular />,
  },
  {
    key: 'appBuilderApps' as const,
    typeKey: 'microsoft.powerapps/apps',
    label: 'App Builder Apps',
    icon: <BoxRegular />,
  },
  {
    key: 'm365AgentFlows' as const,
    typeKey: 'microsoft.powerautomate/m365agentflows',
    label: 'M365 Agent Flows',
    icon: <SettingsRegular />,
  },
  {
    key: 'codeApps' as const,
    typeKey: 'microsoft.powerapps/codeapps',
    label: 'Code Apps',
    icon: <CodeRegular />,
  },
];

export default function Dashboard({
  resources,
  counts,
  environmentsCount,
  recommendations,
  isLoading,
  error,
  onNavigateToResources,
  onNavigateToRecommendations,
}: DashboardProps): ReactElement {
  const styles = useStyles();

  const topRecommendations = useMemo(
    () => [...recommendations]
      .sort((a, b) => b.details.resourceCount - a.details.resourceCount)
      .slice(0, 3),
    [recommendations],
  );

  if (isLoading) {
    return <OperationsSkeleton />;
  }

  if (error) {
    return (
      <div className={styles.root}>
        <MessageBar intent="error">
          <MessageBarBody>{extractMessage(error)}</MessageBarBody>
        </MessageBar>
      </div>
    );
  }

  return (
    <div className={styles.root}>
      <div className={styles.scrollable}>
        <div className={styles.topSection}>
          <section className={styles.operationsHero} aria-labelledby="tenant-operations-title">
            <div className={styles.heroCopy}>
              <Text as="h1" id="tenant-operations-title" className={styles.heroTitle}>
                Tenant command, live.
              </Text>
              <Text className={styles.heroDescription}>
                Inventory, environment state, and governance guidance monitored from one operational surface.
              </Text>
              <div className={styles.heroRegister} aria-label="Tenant index status">
                <Text className={styles.registerItem}>Live inventory</Text>
                <Text className={styles.registerItem}>{environmentsCount.toLocaleString()} environments</Text>
                <Text className={styles.registerItem}>{recommendations.length.toLocaleString()} active signals</Text>
              </div>
            </div>
            <div className={styles.signalMap} aria-label={`${resources.length} indexed resources`}>
              <div className={styles.signalCore}>
                <Text className={styles.signalValue}>{resources.length.toLocaleString()}</Text>
                <Text className={styles.signalLabel}>Resources</Text>
              </div>
              <div className={`${styles.orbitNode} ${styles.orbitTop}`}>
                <Text className={styles.orbitValue}>{recommendations.length.toLocaleString()}</Text>
                <Text className={styles.orbitLabel}>Signals</Text>
              </div>
              <div className={`${styles.orbitNode} ${styles.orbitBottom}`}>
                <Text className={styles.orbitValue}>{environmentsCount.toLocaleString()}</Text>
                <Text className={styles.orbitLabel}>Environments</Text>
              </div>
            </div>
          </section>

          <div className={styles.operationsGrid}>
            <section className={styles.movementBoard} aria-label="Tenant inventory">
              <div className={styles.tenantSummary}>
                <div>
                  <Text className={styles.summaryLabel}>Resources in view</Text>
                  <Text className={styles.resourceTotal}>{resources.length.toLocaleString()}</Text>
                </div>
                <Text className={styles.summaryMeta}>
                  <GlobeRegular />
                  {environmentsCount.toLocaleString()} environment{environmentsCount === 1 ? '' : 's'}
                </Text>
              </div>
              <div className={styles.cardsGrid}>
                {METRIC_ITEMS.map((item) => (
                  <button
                    type="button"
                    key={item.key}
                    className={styles.metricCard}
                    onClick={() => onNavigateToResources(item.typeKey)}
                    title={`View all ${item.label}`}
                  >
                    <span className={styles.metricIcon}>{item.icon}</span>
                    <Text className={styles.metricLabel}>{item.label}</Text>
                    <Text className={styles.metricCount}>{counts ? String(counts[item.key]) : '—'}</Text>
                  </button>
                ))}
              </div>
            </section>

            <section className={styles.advisoryBoard} aria-label="Recommendations requiring review">
              <div className={styles.advisoryHeader}>
                <LightbulbRegular />
                <Text weight="semibold">Recommendations</Text>
                <Text className={styles.advisoryCount}>{recommendations.length} active</Text>
              </div>
              <div className={styles.advisoryList}>
                {topRecommendations.length === 0 ? (
                  <div className={styles.advisoryRow}>
                    <Text className={styles.metricLabel}>No recommendations returned by the admin API.</Text>
                  </div>
                ) : topRecommendations.map((recommendation) => (
                  <div className={styles.advisoryRow} key={recommendation.scenario}>
                    <Text className={styles.advisoryName} title={recommendation.scenario}>
                      {recommendation.scenario.replace(/[_-]/g, ' ').replace(/([a-z])([A-Z])/g, '$1 $2')}
                    </Text>
                    <Text className={styles.affected}>{recommendation.details.resourceCount} affected</Text>
                  </div>
                ))}
              </div>
              <Button
                className={styles.advisoryAction}
                appearance="subtle"
                icon={<ArrowRightRegular />}
                iconPosition="after"
                onClick={onNavigateToRecommendations}
              >
                Review all recommendations
              </Button>
            </section>
          </div>
        </div>

      </div>
    </div>
  );
}
