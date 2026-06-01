import { useState, useEffect } from 'react';
import type { ReactElement } from 'react';
import {
  MessageBar,
  MessageBarBody,
  MessageBarActions,
  Button,
  MenuItem,
  Spinner,
  Text,
  useToastController,
  Toast,
  ToastTitle,
  ToastBody,
} from '@fluentui/react-components';
import { PersonAddRegular, PersonDeleteRegular, ShieldPersonRegular } from '@fluentui/react-icons';
import {
  getCurrentUserRoleAssignmentId,
  addSelfAsEnvironmentAdmin,
  removeSelfAdminRole,
} from '../services/environmentMutations.ts';
import { extractMessage } from '../utils/errorUtils.ts';

interface Props {
  environmentId: string;
  /** Called after admin access is applied or removed. */
  onChanged?: () => void;
  /**
   * 'menu' – renders as MenuItems to be placed inside a MenuList.
   */
  variant?: 'banner' | 'inline' | 'menu';
}

/**
 * Shows controls to apply or remove elevated admin access for the environment.
 * - If the user has an active role assignment → shows "Remove access".
 * - If the user has no role assignment → shows "Apply admin access".
 * - Renders nothing while the initial check is in progress.
 */
export default function AddSelfAsAdminBanner({ environmentId, onChanged, variant = 'banner' }: Props): ReactElement | null {
  const { dispatchToast } = useToastController();
  const [checking, setChecking] = useState(true);
  const [assignmentId, setAssignmentId] = useState<string | null | undefined>(undefined); // undefined = not yet checked
  const [acting, setActing] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (!environmentId) { setChecking(false); return; }
    setChecking(true);
    setAssignmentId(undefined);
    setError(null);
    getCurrentUserRoleAssignmentId(environmentId)
      .then((id) => setAssignmentId(id))
      .catch(() => setAssignmentId(null))
      .finally(() => setChecking(false));
  }, [environmentId]);

  async function handleApply() {
    setActing(true);
    setError(null);
    try {
      await addSelfAsEnvironmentAdmin(environmentId);
      const id = await getCurrentUserRoleAssignmentId(environmentId);
      setAssignmentId(id);
      onChanged?.();
      dispatchToast(
        <Toast><ToastTitle>Admin access applied</ToastTitle><ToastBody>Your admin role has been applied for this environment.</ToastBody></Toast>,
        { intent: 'success' },
      );
    } catch (e: unknown) {
      setError(extractMessage(String(e)));
      dispatchToast(
        <Toast><ToastTitle>Failed to apply admin access</ToastTitle><ToastBody>{extractMessage(String(e))}</ToastBody></Toast>,
        { intent: 'error' },
      );
    } finally {
      setActing(false);
    }
  }

  async function handleRemove() {
    setActing(true);
    setError(null);
    try {
      await removeSelfAdminRole(environmentId);
      setAssignmentId(null);
      onChanged?.();
      dispatchToast(
        <Toast><ToastTitle>Admin access removed</ToastTitle><ToastBody>Your admin role has been removed from this environment.</ToastBody></Toast>,
        { intent: 'success' },
      );
    } catch (e: unknown) {
      setError(extractMessage(String(e)));
      dispatchToast(
        <Toast><ToastTitle>Failed to remove admin access</ToastTitle><ToastBody>{extractMessage(String(e))}</ToastBody></Toast>,
        { intent: 'error' },
      );
    } finally {
      setActing(false);
    }
  }

  /* ── Menu variant (inside MenuList) ── */
  if (variant === 'menu') {
    if (checking || assignmentId === undefined) {
      return (
        <MenuItem icon={<Spinner size="tiny" />} disabled>
          Checking admin access…
        </MenuItem>
      );
    }
    if (assignmentId) {
      return (
        <MenuItem
          icon={acting ? <Spinner size="tiny" /> : <PersonDeleteRegular />}
          disabled={acting}
          onClick={() => void handleRemove()}
        >
          {acting ? 'Removing access…' : 'Remove admin access'}
        </MenuItem>
      );
    }
    return (
      <MenuItem
        icon={acting ? <Spinner size="tiny" /> : <PersonAddRegular />}
        disabled={acting}
        onClick={() => void handleApply()}
      >
        {acting ? 'Applying…' : 'Apply admin access'}
      </MenuItem>
    );
  }

  if (checking || assignmentId === undefined) return null;

  /* ── Inline variant (action bar) ── */
  if (variant === 'inline') {
    if (assignmentId) {
      return (
        <>
          <div style={{ width: 1, height: 16, background: 'var(--colorNeutralStroke2)', flexShrink: 0 }} />
          <ShieldPersonRegular style={{ color: 'var(--colorBrandForeground1)', flexShrink: 0 }} fontSize={16} />
          <Text size={200} style={{ color: 'var(--colorBrandForeground1)', whiteSpace: 'nowrap' }}>
            {error ? 'Failed to remove access' : 'Elevated admin access'}
          </Text>
          <Button
            size="small"
            appearance="subtle"
            icon={acting ? <Spinner size="tiny" /> : <PersonDeleteRegular />}
            disabled={acting}
            onClick={() => void handleRemove()}
          >
            {acting ? 'Removing…' : 'Remove access'}
          </Button>
        </>
      );
    }
    return (
      <>
        <div style={{ width: 1, height: 16, background: 'var(--colorNeutralStroke2)', flexShrink: 0 }} />
        <Button
          size="small"
          appearance="subtle"
          icon={acting ? <Spinner size="tiny" /> : <PersonAddRegular />}
          disabled={acting}
          onClick={() => void handleApply()}
        >
          {acting ? 'Applying…' : 'Apply admin access'}
        </Button>
      </>
    );
  }

  /* ── Banner variant (default) ── */
  if (assignmentId) {
    return (
      <MessageBar intent={error ? 'error' : 'info'} style={{ flexShrink: 0 }}>
        <MessageBarBody>
          {error ? `Failed to remove admin access: ${error}` : 'You have elevated admin access for this environment.'}
        </MessageBarBody>
        {!error && (
          <MessageBarActions>
            <Button
              size="small"
              appearance="subtle"
              icon={acting ? <Spinner size="tiny" /> : <PersonDeleteRegular />}
              disabled={acting}
              onClick={() => void handleRemove()}
            >
              {acting ? 'Removing…' : 'Remove access'}
            </Button>
          </MessageBarActions>
        )}
      </MessageBar>
    );
  }

  return (
    <MessageBar intent={error ? 'error' : 'warning'} style={{ flexShrink: 0 }}>
      <MessageBarBody>
        {error
          ? `Failed to apply admin access: ${error}`
          : 'You are missing the required Dataverse privilege on this environment. Some data may be unavailable.'}
      </MessageBarBody>
      {!error && (
        <MessageBarActions>
          <Button
            size="small"
            appearance="primary"
            icon={acting ? <Spinner size="tiny" /> : <PersonAddRegular />}
            disabled={acting}
            onClick={() => void handleApply()}
          >
            {acting ? 'Applying…' : 'Apply admin access'}
          </Button>
        </MessageBarActions>
      )}
    </MessageBar>
  );
}


