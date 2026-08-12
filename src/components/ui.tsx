import type { ReactElement, ReactNode } from 'react';
import { Skeleton, SkeletonItem, Spinner, Text, makeStyles, tokens } from '@fluentui/react-components';

const useStyles = makeStyles({
  pageHeader: {
    position: 'relative',
    display: 'flex',
    alignItems: 'flex-start',
    justifyContent: 'space-between',
    gap: '32px',
    flexWrap: 'wrap',
    flexShrink: 0,
    minHeight: '104px',
    padding: '20px 26px',
    overflow: 'hidden',
    color: '#F4FBFD',
    border: '1px solid #29404F',
    backgroundColor: '#0B121A',
    backgroundImage: 'linear-gradient(90deg, rgba(67, 217, 255, 0.055) 1px, transparent 1px)',
    backgroundSize: '40px 100%',
    boxShadow: '0 10px 28px rgba(0, 0, 0, 0.3)',
    '::before': {
      content: '""',
      position: 'absolute',
      top: 0,
      right: 0,
      bottom: 0,
      width: '12%',
      backgroundColor: '#10232F',
      clipPath: 'polygon(30% 0, 100% 0, 100% 100%, 0 100%)',
    },
    '::after': {
      content: '""',
      position: 'absolute',
      right: '12%',
      bottom: 0,
      width: '64px',
      height: '6px',
      backgroundColor: '#FFB547',
    },
    '@media (max-width: 620px)': {
      minHeight: 'auto',
      padding: '22px',
    },
  },
  headingGroup: {
    position: 'relative',
    zIndex: 1,
    display: 'grid',
    gap: '6px',
    minWidth: 0,
  },
  title: {
    fontWeight: tokens.fontWeightSemibold,
    letterSpacing: '-0.035em',
    color: '#F4FBFD',
    fontSize: '34px',
    lineHeight: '38px',
  },
  description: {
    color: '#91A8B5',
    maxWidth: '68ch',
    fontSize: tokens.fontSizeBase300,
    lineHeight: '20px',
  },
  actions: {
    position: 'relative',
    zIndex: 1,
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    flexWrap: 'wrap',
  },
  skeleton: {
    position: 'relative',
    display: 'grid',
    gridTemplateRows: '220px 1fr',
    gap: '22px',
    height: '100%',
    padding: '28px 32px 32px',
    overflow: 'hidden',
    backgroundColor: tokens.colorNeutralBackground2,
    backgroundImage: 'linear-gradient(rgba(67, 217, 255, 0.028) 1px, transparent 1px), linear-gradient(90deg, rgba(67, 217, 255, 0.028) 1px, transparent 1px)',
    backgroundSize: '32px 32px',
    '@media (max-width: 768px)': {
      padding: tokens.spacingHorizontalM,
    },
  },
  loadingHero: {
    position: 'relative',
    display: 'grid',
    gridTemplateColumns: '180px minmax(0, 1fr)',
    alignItems: 'center',
    overflow: 'hidden',
    color: '#F8FCFC',
    border: '1px solid #29404F',
    backgroundColor: '#09121A',
    backgroundImage: 'linear-gradient(110deg, #09121A 0 72%, #1A1710 72%)',
    boxShadow: '0 12px 30px rgba(0, 0, 0, 0.32)',
    '@media (max-width: 620px)': {
      gridTemplateColumns: '110px minmax(0, 1fr)',
    },
  },
  signalField: {
    position: 'relative',
    display: 'grid',
    placeItems: 'center',
    alignSelf: 'stretch',
    borderRight: '1px solid #29404F',
  },
  signalRing: {
    position: 'absolute',
    width: '98px',
    height: '98px',
    border: '1px solid rgba(67, 217, 255, 0.72)',
    borderRadius: '50%',
    boxShadow: 'inset 0 0 0 12px rgba(67, 217, 255, 0.05), 0 0 24px rgba(67, 217, 255, 0.12)',
    animation: 'signal-breathe 2.4s cubic-bezier(.16, 1, .3, 1) infinite',
  },
  signalCore: {
    position: 'relative',
    zIndex: 1,
    display: 'grid',
    placeItems: 'center',
    width: '48px',
    height: '48px',
    color: '#43D9FF',
    border: '1px solid #43D9FF',
    borderRadius: '50%',
    backgroundColor: '#0B1D26',
    boxShadow: '0 0 22px rgba(67, 217, 255, 0.28)',
  },
  loadingCopy: {
    display: 'grid',
    gap: '9px',
    padding: '28px 34px',
  },
  loadingTitle: {
    color: '#FFFFFF',
    fontSize: '32px',
    lineHeight: '36px',
    fontWeight: tokens.fontWeightSemibold,
    letterSpacing: '-0.035em',
  },
  loadingDescription: {
    maxWidth: '54ch',
    color: '#B6D2D8',
    lineHeight: '20px',
  },
  loadingStatus: {
    width: 'fit-content',
    marginTop: '8px',
    padding: '5px 9px',
    color: '#FFCF87',
    border: '1px solid #6B512C',
    fontSize: '10px',
    fontWeight: tokens.fontWeightSemibold,
    letterSpacing: '0.08em',
    textTransform: 'uppercase',
  },
  skeletonBody: {
    display: 'grid',
    gridTemplateColumns: 'repeat(2, minmax(0, 1fr))',
    gridTemplateRows: 'repeat(3, 54px)',
    gap: '1px',
    alignContent: 'start',
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    backgroundColor: tokens.colorNeutralStroke2,
    boxShadow: tokens.shadow4,
    '@media (max-width: 620px)': {
      gridTemplateColumns: '1fr',
    },
  },
  skeletonRow: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    padding: `0 ${tokens.spacingHorizontalM}`,
    backgroundColor: tokens.colorNeutralBackground1,
  },
});

interface PageHeaderProps {
  title: string;
  description?: string;
  actions?: ReactNode;
}

export function PageHeader({ title, description, actions }: PageHeaderProps): ReactElement {
  const styles = useStyles();
  return (
    <header className={styles.pageHeader}>
      <div className={styles.headingGroup}>
        <Text as="h1" className={styles.title}>{title}</Text>
        {description && <Text className={styles.description}>{description}</Text>}
      </div>
      {actions && <div className={styles.actions}>{actions}</div>}
    </header>
  );
}

export function OperationsSkeleton(): ReactElement {
  const styles = useStyles();
  return (
    <Skeleton className={styles.skeleton} aria-label="Synchronizing tenant data">
      <div className={styles.loadingHero}>
        <div className={styles.signalField} aria-hidden="true">
          <span className={styles.signalRing} />
          <span className={styles.signalCore}>
            <Spinner size="tiny" />
          </span>
        </div>
        <div className={styles.loadingCopy}>
          <Text className={styles.loadingTitle}>Establishing tenant signal</Text>
          <Text className={styles.loadingDescription}>
            Mapping environments, resources, connections, and governance guidance into the operational index.
          </Text>
          <Text className={styles.loadingStatus}>Secure synchronization in progress</Text>
        </div>
      </div>
      <div className={styles.skeletonBody}>
        {Array.from({ length: 6 }, (_, index) => (
          <div className={styles.skeletonRow} key={index}>
            <SkeletonItem size={20} shape="circle" />
            <SkeletonItem size={16} style={{ width: `${42 + (index % 3) * 12}%` }} />
          </div>
        ))}
      </div>
    </Skeleton>
  );
}
