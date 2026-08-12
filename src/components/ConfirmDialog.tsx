import type { ReactElement } from 'react';
import {
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import { WarningRegular } from '@fluentui/react-icons';

const useStyles = makeStyles({
  surface: {
    maxWidth: '440px',
    color: '#F4FBFD',
    border: '1px solid #29404F',
    borderRadius: 0,
    backgroundColor: '#0C141D',
    boxShadow: '0 24px 64px rgba(0, 0, 0, 0.62)',
  },
  warningIcon: {
    color: tokens.colorStatusDangerForeground1,
    fontSize: '20px',
  },
});

export interface ConfirmDialogProps {
  open: boolean;
  title: string;
  message: string;
  confirmLabel?: string;
  confirmAppearance?: 'primary' | 'secondary';
  isDangerous?: boolean;
  isLoading?: boolean;
  onConfirm: () => void;
  onCancel: () => void;
}

export default function ConfirmDialog({
  open,
  title,
  message,
  confirmLabel = 'Confirm',
  confirmAppearance = 'primary',
  isDangerous = false,
  isLoading = false,
  onConfirm,
  onCancel,
}: ConfirmDialogProps): ReactElement {
  const styles = useStyles();
  return (
    <Dialog open={open} onOpenChange={(_, data) => { if (!data.open) onCancel(); }}>
      <DialogSurface className={styles.surface}>
        <DialogBody>
          <DialogTitle
            action={isDangerous
              ? <WarningRegular aria-hidden="true" className={styles.warningIcon} />
              : undefined}
          >
            {title}
          </DialogTitle>
          <DialogContent>{message}</DialogContent>
          <DialogActions>
            <Button appearance="secondary" disabled={isLoading} onClick={onCancel}>
              Cancel
            </Button>
            <Button
              appearance={isDangerous ? 'primary' : confirmAppearance}
              disabled={isLoading}
              onClick={onConfirm}
              style={isDangerous ? {
                backgroundColor: tokens.colorStatusDangerBackground3,
                borderColor: tokens.colorStatusDangerBackground3,
                color: tokens.colorNeutralForegroundOnBrand,
              } : undefined}
            >
              {isLoading ? 'Working…' : confirmLabel}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
