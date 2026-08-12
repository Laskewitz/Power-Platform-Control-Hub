import type { ReactElement, ReactNode } from 'react';
import { makeStyles, tokens, Text, Button } from '@fluentui/react-components';

const useStyles = makeStyles({
  root: {
    position: 'relative',
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    gap: tokens.spacingVerticalM,
    minHeight: '220px',
    padding: '32px 28px',
    overflow: 'hidden',
    color: '#91A8B5',
    textAlign: 'center',
    border: '1px solid #29404F',
    borderRadius: 0,
    backgroundColor: '#09121A',
    backgroundImage: 'linear-gradient(rgba(67, 217, 255, 0.035) 1px, transparent 1px), linear-gradient(90deg, rgba(67, 217, 255, 0.035) 1px, transparent 1px)',
    backgroundSize: '32px 32px',
    '::after': {
      content: '""',
      position: 'absolute',
      right: 0,
      bottom: 0,
      width: '72px',
      height: '4px',
      backgroundColor: '#FFB547',
    },
  },
  icon: {
    display: 'grid',
    placeItems: 'center',
    width: '48px',
    height: '48px',
    color: '#43D9FF',
    fontSize: '25px',
    border: '1px solid #43D9FF',
    borderRadius: 0,
    backgroundColor: '#0B1D26',
    boxShadow: '0 8px 20px rgba(0, 0, 0, 0.3)',
    lineHeight: 1,
  },
  title: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    color: '#F4FBFD',
  },
  subtitle: {
    fontSize: tokens.fontSizeBase300,
    color: '#91A8B5',
    maxWidth: '360px',
  },
});

interface EmptyStateProps {
  icon: ReactNode;
  title: string;
  subtitle?: string;
  action?: { label: string; onClick: () => void };
}

export default function EmptyState({ icon, title, subtitle, action }: EmptyStateProps): ReactElement {
  const styles = useStyles();
  return (
    <div className={styles.root}>
      <div className={styles.icon}>{icon}</div>
      <Text className={styles.title}>{title}</Text>
      {subtitle && <Text className={styles.subtitle}>{subtitle}</Text>}
      {action && (
        <Button appearance="primary" size="small" onClick={action.onClick}>
          {action.label}
        </Button>
      )}
    </div>
  );
}
