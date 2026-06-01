import type { ReactElement } from 'react';
import {
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  tokens,
} from '@fluentui/react-components';
import { WarningRegular } from '@fluentui/react-icons';

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
  return (
    <Dialog open={open} onOpenChange={(_, data) => { if (!data.open) onCancel(); }}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle
            action={isDangerous
              ? <WarningRegular style={{ color: tokens.colorStatusDangerForeground1, fontSize: '20px' }} />
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
