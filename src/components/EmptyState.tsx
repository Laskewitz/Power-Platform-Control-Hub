import type { ReactElement, ReactNode } from 'react';
import { makeStyles, tokens, Text, Button } from '@fluentui/react-components';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    gap: tokens.spacingVerticalM,
    padding: `${tokens.spacingVerticalXXL} ${tokens.spacingHorizontalXL}`,
    color: tokens.colorNeutralForeground3,
    textAlign: 'center',
  },
  icon: {
    fontSize: '40px',
    opacity: 0.4,
    lineHeight: 1,
  },
  title: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground2,
  },
  subtitle: {
    fontSize: tokens.fontSizeBase300,
    color: tokens.colorNeutralForeground3,
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
