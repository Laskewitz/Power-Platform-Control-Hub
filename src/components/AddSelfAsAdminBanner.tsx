import { useState, useEffect } from 'react';
import type { ReactElement } from 'react';
import {
  MessageBar,
  MessageBarBody,
  MessageBarActions,
  Button,
  Spinner,
} from '@fluentui/react-components';
import { PersonAddRegular } from '@fluentui/react-icons';
import {
  checkCurrentUserIsEnvAdmin,
  addSelfAsEnvironmentAdmin,
} from '../services/environmentMutations.ts';
import { extractMessage } from '../utils/errorUtils.ts';

interface Props {
  environmentId: string;
  /** Called after the user is successfully added as System Administrator. */
  onAdded?: () => void;
}

/**
 * Silently checks whether the signed-in user is a System Administrator on
 * `environmentId`. If not, renders a warning banner with a one-click button
 * to assign the role. Renders nothing while checking or if the user is
 * already an admin.
 */
export default function AddSelfAsAdminBanner({ environmentId, onAdded }: Props): ReactElement | null {
  const [checking, setChecking] = useState(true);
  const [isAdmin, setIsAdmin] = useState<boolean | null>(null);
  const [adding, setAdding] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (!environmentId) { setChecking(false); return; }
    setChecking(true);
    setIsAdmin(null);
    setError(null);
    checkCurrentUserIsEnvAdmin(environmentId)
      .then((result) => setIsAdmin(result))
      .catch(() => setIsAdmin(true)) // if check fails, stay silent
      .finally(() => setChecking(false));
  }, [environmentId]);

  async function handleAdd() {
    setAdding(true);
    setError(null);
    try {
      await addSelfAsEnvironmentAdmin(environmentId);
      setIsAdmin(true);
      onAdded?.();
    } catch (e: unknown) {
      setError(extractMessage(String(e)));
    } finally {
      setAdding(false);
    }
  }

  // Still checking or already admin — render nothing
  if (checking || isAdmin !== false) return null;

  return (
    <MessageBar intent={error ? 'error' : 'warning'} style={{ flexShrink: 0 }}>
      <MessageBarBody>
        {error
          ? `Failed to add you as System Administrator: ${error}`
          : 'You are not a System Administrator on this environment. Some data may be unavailable.'}
      </MessageBarBody>
      {!error && (
        <MessageBarActions>
          <Button
            size="small"
            appearance="primary"
            icon={adding ? <Spinner size="tiny" /> : <PersonAddRegular />}
            disabled={adding}
            onClick={() => void handleAdd()}
          >
            {adding ? 'Adding…' : 'Make me admin'}
          </Button>
        </MessageBarActions>
      )}
    </MessageBar>
  );
}
