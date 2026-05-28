import { useEffect, useState } from 'react';
import type { ReactElement } from 'react';
import {
  Dialog,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogContent,
  DialogActions,
  Button,
  Dropdown,
  Option,
  Text,
  Spinner,
  tokens,
  makeStyles,
} from '@fluentui/react-components';
import { LayerRegular, SubtractCircleRegular } from '@fluentui/react-icons';
import type { EnvironmentGroup } from '../types/admin.ts';
import { fetchEnvironmentGroups, addEnvironmentToGroup, removeEnvironmentFromGroup } from '../services/adminApi.ts';
import { useMutation } from '../hooks/useMutation.tsx';

interface EnvironmentGroupDialogProps {
  open: boolean;
  mode: 'add' | 'remove';
  environmentName: string;
  /** The environment's GUID (name field from Resource). */
  environmentId: string;
  /** When provided in remove mode, the group is pre-selected and the dropdown is locked. */
  preselectedGroupId?: string;
  onClose: () => void;
  onSuccess?: () => void;
}

const useStyles = makeStyles({
  content: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
  },
  dropdownRow: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXS,
  },
  meta: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
  },
  loadingRow: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    padding: `${tokens.spacingVerticalS} 0`,
  },
  empty: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase300,
  },
});

export default function EnvironmentGroupDialog({
  open,
  mode,
  environmentName,
  environmentId,
  preselectedGroupId,
  onClose,
  onSuccess,
}: EnvironmentGroupDialogProps): ReactElement {
  const styles = useStyles();
  const [groups, setGroups] = useState<EnvironmentGroup[]>([]);
  const [loadingGroups, setLoadingGroups] = useState(false);
  const [selectedGroupId, setSelectedGroupId] = useState<string>('');

  // Lazy-load groups when dialog opens; auto-select when preselectedGroupId is set
  useEffect(() => {
    if (!open) return;
    setSelectedGroupId(preselectedGroupId ?? '');
    setLoadingGroups(true);
    fetchEnvironmentGroups()
      .then(setGroups)
      .catch(() => setGroups([]))
      .finally(() => setLoadingGroups(false));
  }, [open, preselectedGroupId]);

  const { execute: execAdd, isLoading: isAdding } = useMutation(
    (groupId: string) => addEnvironmentToGroup(groupId, environmentId),
    {
      successMessage: `Environment added to group.`,
      onSuccess: () => {
        onSuccess?.();
        onClose();
      },
    },
  );

  const { execute: execRemove, isLoading: isRemoving } = useMutation(
    (groupId: string) => removeEnvironmentFromGroup(groupId, environmentId),
    {
      successMessage: `Environment removed from group.`,
      onSuccess: () => {
        onSuccess?.();
        onClose();
      },
    },
  );

  const isBusy = isAdding || isRemoving;

  function handleConfirm() {
    if (!selectedGroupId) return;
    if (mode === 'add') void execAdd(selectedGroupId);
    else void execRemove(selectedGroupId);
  }

  const title = mode === 'add' ? 'Add to Environment Group' : 'Remove from Environment Group';
  const confirmLabel = mode === 'add' ? 'Add to Group' : 'Remove from Group';
  const confirmAppearance = mode === 'add' ? 'primary' : 'primary';
  const confirmIcon = mode === 'add' ? <LayerRegular /> : <SubtractCircleRegular />;

  const selectedGroup = groups.find((g) => g.id === selectedGroupId);
  const isLocked = mode === 'remove' && !!preselectedGroupId;

  const hint = isLocked
    ? `Remove "${environmentName}" from the group shown below.`
    : mode === 'add'
    ? `Select a group to add "${environmentName}" to.`
    : `Select the group to remove "${environmentName}" from.`;

  return (
    <Dialog open={open} onOpenChange={(_, data) => { if (!data.open) onClose(); }}>
      <DialogSurface>
        <DialogTitle>{title}</DialogTitle>
        <DialogBody>
          <DialogContent>
            <div className={styles.content}>
              <Text>{hint}</Text>

              {loadingGroups ? (
                <div className={styles.loadingRow}>
                  <Spinner size="tiny" />
                  <Text className={styles.meta}>Loading groups…</Text>
                </div>
              ) : groups.length === 0 ? (
                <Text className={styles.empty}>No environment groups found.</Text>
              ) : (
                <div className={styles.dropdownRow}>
                  <Text weight="semibold" size={200}>Group</Text>
                  <Dropdown
                    placeholder="Select a group…"
                    value={selectedGroup?.displayName ?? ''}
                    selectedOptions={selectedGroupId ? [selectedGroupId] : []}
                    onOptionSelect={(_, data) => setSelectedGroupId(data.optionValue ?? '')}
                    disabled={isBusy || isLocked}
                  >
                    {groups.map((g) => (
                      <Option key={g.id} value={g.id} text={g.displayName}>
                        <div style={{ display: 'flex', flexDirection: 'column' }}>
                          <Text weight="semibold">{g.displayName}</Text>
                          {g.description && (
                            <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>{g.description}</Text>
                          )}
                        </div>
                      </Option>
                    ))}
                  </Dropdown>
                  {isLocked && (
                    <Text className={styles.meta}>This environment is currently in this group.</Text>
                  )}
                </div>
              )}
            </div>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={onClose} disabled={isBusy}>Cancel</Button>
            <Button
              appearance={confirmAppearance}
              icon={confirmIcon}
              onClick={handleConfirm}
              disabled={isBusy || !selectedGroupId || loadingGroups}
            >
              {isBusy ? <Spinner size="tiny" /> : confirmLabel}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
