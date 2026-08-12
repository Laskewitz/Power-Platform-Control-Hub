import { useState } from 'react';
import type { ReactElement } from 'react';
import {
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  Field,
  Textarea,
  Button,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import { ArchiveRegular } from '@fluentui/react-icons';

const useStyles = makeStyles({
  surface: {
    maxWidth: '440px',
    color: '#F4FBFD',
    border: '1px solid #29404F',
    borderRadius: 0,
    backgroundColor: '#0C141D',
    boxShadow: '0 24px 64px rgba(0, 0, 0, 0.62)',
  },
  titleIcon: {
    color: '#43D9FF',
    fontSize: '20px',
  },
  field: {
    marginTop: tokens.spacingVerticalS,
  },
});

interface BackupDialogProps {
  open: boolean;
  environmentName: string;
  isLoading: boolean;
  onConfirm: (notes: string) => void;
  onCancel: () => void;
}

export default function BackupDialog({
  open,
  environmentName,
  isLoading,
  onConfirm,
  onCancel,
}: BackupDialogProps): ReactElement {
  const styles = useStyles();
  const [notes, setNotes] = useState('');
  return (
    <Dialog open={open} onOpenChange={(_, data) => { if (!data.open) onCancel(); }}>
      <DialogSurface className={styles.surface}>
        <DialogBody>
          <DialogTitle action={<ArchiveRegular aria-hidden="true" className={styles.titleIcon} />}>
            Create backup — {environmentName}
          </DialogTitle>
          <DialogContent>
            <Field label="Backup notes (optional)" className={styles.field}>
              <Textarea
                value={notes}
                onChange={(_, data) => setNotes(data.value)}
                placeholder="e.g. Pre-release snapshot"
                rows={3}
              />
            </Field>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" disabled={isLoading} onClick={onCancel}>Cancel</Button>
            <Button appearance="primary" disabled={isLoading} onClick={() => onConfirm(notes)}>
              {isLoading ? 'Submitting…' : 'Create backup'}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
