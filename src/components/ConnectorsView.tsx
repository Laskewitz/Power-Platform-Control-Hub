import { useCallback, useEffect, useMemo, useState } from 'react';
import type { ReactElement } from 'react';
import {
  Badge,
  Button,
  Dropdown,
  MessageBar,
  MessageBarBody,
  Option,
  ProgressBar,
  Tab,
  TabList,
  Text,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import {
  ArrowSyncRegular,
  CheckmarkCircleRegular,
  ClockRegular,
  DeleteRegular,
  ErrorCircleRegular,
  PlugConnectedRegular,
  TableRegular,
  WarningRegular,
} from '@fluentui/react-icons';
import { fetchConnections, fetchConnectors } from '../services/adminApi.ts';
import { deleteConnection } from '../services/connectorMutations.ts';
import type { Connection } from '../types/admin.ts';
import type { Resource } from '../types/inventory.ts';
import { useMutation } from '../hooks/useMutation.tsx';
import {
  getConnectorUsageTargetCount,
  scanTenantConnectorUsage,
} from '../services/connectorUsageService.ts';
import type { ConnectorUsageCategory, ConnectorUsageScanResult } from '../services/connectorUsageService.ts';
import ConfirmDialog from './ConfirmDialog.tsx';
import EmptyState from './EmptyState.tsx';
import { OperationsSkeleton, PageHeader } from './ui.tsx';

interface ConnectorsViewProps {
  environments: Resource[];
  resources: Resource[];
}

type ConnectorsTab = 'health' | 'usage' | 'connectors';
type HealthLevel = 'healthy' | 'warning' | 'critical' | 'unknown';
const ALL_ENVIRONMENTS = '__all__';

interface PendingConnectionDeletion {
  connection: Connection;
  environmentId: string;
}

interface ConnectionHealth {
  level: HealthLevel;
  label: string;
  detail: string;
}

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    gap: '18px',
    height: '100%',
    padding: '28px 32px 32px',
    overflow: 'hidden',
    '@media (max-width: 768px)': {
      padding: tokens.spacingHorizontalM,
    },
  },
  controls: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    flexWrap: 'wrap',
    flexShrink: 0,
  },
  environmentSelect: {
    minWidth: '280px',
  },
  tabs: {
    marginLeft: 'auto',
  },
  body: {
    display: 'flex',
    flex: 1,
    flexDirection: 'column',
    gap: '14px',
    minHeight: 0,
  },
  signalRail: {
    display: 'grid',
    gridTemplateColumns: '1.35fr repeat(3, minmax(120px, 0.65fr))',
    minHeight: '92px',
    border: '1px solid #29404F',
    backgroundColor: '#09121A',
    '@media (max-width: 760px)': {
      gridTemplateColumns: '1fr 1fr',
    },
  },
  signalLead: {
    display: 'grid',
    alignContent: 'center',
    gap: '5px',
    padding: '16px 20px',
    borderRight: '1px solid #29404F',
    backgroundImage: 'linear-gradient(90deg, rgba(67, 217, 255, 0.07) 1px, transparent 1px)',
    backgroundSize: '32px 100%',
  },
  signalTitle: {
    color: '#F4FBFD',
    fontSize: '18px',
    fontWeight: tokens.fontWeightSemibold,
  },
  signalDescription: {
    color: '#91A8B5',
    fontSize: tokens.fontSizeBase200,
  },
  signalCell: {
    display: 'grid',
    alignContent: 'center',
    gap: '2px',
    padding: '14px 18px',
    borderRight: '1px solid #20313E',
    ':last-child': {
      borderRight: 0,
    },
  },
  signalValue: {
    color: '#F4FBFD',
    fontSize: '28px',
    lineHeight: '30px',
    fontWeight: tokens.fontWeightSemibold,
    fontVariantNumeric: 'tabular-nums',
  },
  signalLabel: {
    color: '#91A8B5',
    fontSize: '10px',
    fontWeight: tokens.fontWeightSemibold,
    letterSpacing: '0.08em',
    textTransform: 'uppercase',
  },
  list: {
    flex: 1,
    minHeight: 0,
    overflow: 'auto',
    border: '1px solid #29404F',
    backgroundColor: tokens.colorNeutralBackground1,
  },
  listHeader: {
    display: 'grid',
    gridTemplateColumns: 'minmax(190px, 1.1fr) minmax(170px, 0.9fr) minmax(160px, 0.85fr) minmax(130px, 0.65fr) minmax(150px, 0.9fr) 170px 42px',
    gap: '16px',
    alignItems: 'center',
    minWidth: '1220px',
    padding: '10px 14px',
    color: tokens.colorNeutralForeground3,
    borderBottom: '1px solid #29404F',
    backgroundColor: tokens.colorNeutralBackground3,
    fontSize: '10px',
    fontWeight: tokens.fontWeightSemibold,
    letterSpacing: '0.07em',
    textTransform: 'uppercase',
  },
  healthRow: {
    display: 'grid',
    gridTemplateColumns: 'minmax(190px, 1.1fr) minmax(170px, 0.9fr) minmax(160px, 0.85fr) minmax(130px, 0.65fr) minmax(150px, 0.9fr) 170px 42px',
    gap: '16px',
    alignItems: 'center',
    minWidth: '1220px',
    minHeight: '68px',
    padding: '10px 14px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    ':hover': {
      backgroundColor: '#101D27',
    },
  },
  identity: {
    display: 'grid',
    gap: '3px',
    minWidth: 0,
  },
  primary: {
    color: tokens.colorNeutralForeground1,
    fontWeight: tokens.fontWeightSemibold,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  secondary: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  healthStatus: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    fontWeight: tokens.fontWeightSemibold,
  },
  statusHealthy: { color: '#66E3A4' },
  statusWarning: { color: '#FFB547' },
  statusCritical: { color: '#FF8A8A' },
  statusUnknown: { color: '#91A8B5' },
  expiry: {
    display: 'grid',
    gap: '3px',
    fontVariantNumeric: 'tabular-nums',
  },
  empty: {
    display: 'grid',
    placeItems: 'center',
    minHeight: '280px',
  },
  connectorToolbar: {
    display: 'flex',
    gap: tokens.spacingHorizontalS,
    flexWrap: 'wrap',
  },
  table: {
    width: '100%',
    minWidth: '720px',
    borderCollapse: 'collapse',
  },
  th: {
    padding: '10px 14px',
    color: tokens.colorNeutralForeground3,
    borderBottom: '1px solid #29404F',
    backgroundColor: tokens.colorNeutralBackground3,
    fontSize: '10px',
    fontWeight: tokens.fontWeightSemibold,
    letterSpacing: '0.07em',
    textAlign: 'left',
    textTransform: 'uppercase',
  },
  td: {
    padding: '11px 14px',
    color: tokens.colorNeutralForeground1,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    fontSize: tokens.fontSizeBase300,
  },
  usageIntro: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    padding: '14px 18px',
    border: '1px solid #29404F',
    backgroundColor: '#09121A',
  },
  usageMeta: {
    display: 'grid',
    gap: '3px',
    marginRight: 'auto',
  },
  progressBlock: {
    display: 'grid',
    gap: '8px',
    padding: '18px',
    border: '1px solid #29404F',
    backgroundColor: '#09121A',
  },
  usageGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(3, minmax(0, 1fr))',
    gap: '1px',
    flex: 1,
    minHeight: 0,
    overflow: 'hidden',
    border: '1px solid #29404F',
    backgroundColor: '#29404F',
    '@media (max-width: 980px)': {
      gridTemplateColumns: '1fr',
      overflowY: 'auto',
    },
  },
  usageColumn: {
    display: 'flex',
    flexDirection: 'column',
    minWidth: 0,
    minHeight: 0,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  usageColumnHeader: {
    display: 'grid',
    gap: '3px',
    padding: '14px 16px',
    borderBottom: '1px solid #29404F',
    backgroundColor: tokens.colorNeutralBackground3,
  },
  usageRows: {
    overflowY: 'auto',
  },
  usageRow: {
    display: 'grid',
    gridTemplateColumns: '28px minmax(0, 1fr) auto',
    gap: '10px',
    alignItems: 'center',
    minHeight: '52px',
    padding: '8px 14px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    ':hover': {
      backgroundColor: '#101D27',
    },
  },
  usageRank: {
    color: '#6F8998',
    fontSize: tokens.fontSizeBase200,
    fontVariantNumeric: 'tabular-nums',
  },
  usageName: {
    display: 'grid',
    gap: '5px',
    minWidth: 0,
  },
  usageBarTrack: {
    height: '3px',
    overflow: 'hidden',
    backgroundColor: '#172733',
  },
  usageBar: {
    height: '100%',
    backgroundColor: '#43D9FF',
  },
  usageCount: {
    color: '#F4FBFD',
    fontWeight: tokens.fontWeightSemibold,
    fontVariantNumeric: 'tabular-nums',
  },
});

function getEnvironmentLabel(environment: Resource): string {
  return environment.properties.displayName ?? environment.name;
}

function formatDate(value?: string): string {
  if (!value) return 'Not reported';
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? value : date.toLocaleString();
}

function daysUntil(value?: string): number | null {
  if (!value) return null;
  const timestamp = new Date(value).getTime();
  if (Number.isNaN(timestamp)) return null;
  return Math.ceil((timestamp - Date.now()) / 86_400_000);
}

function getConnectionHealth(connection: Connection): ConnectionHealth {
  const statuses = connection.properties.statuses ?? [];
  const errorStatus = statuses.find((status) => status.error?.message);
  const unhealthyStatus = statuses.find((status) => {
    const value = status.status?.toLowerCase() ?? '';
    return ['error', 'failed', 'disconnected', 'invalid', 'unauthorized'].some((term) => value.includes(term));
  });

  if (errorStatus || unhealthyStatus) {
    return {
      level: 'critical',
      label: unhealthyStatus?.status || 'Connection failed',
      detail: errorStatus?.error?.message ?? `The ${unhealthyStatus?.target ?? 'connection'} requires attention.`,
    };
  }

  const expiryDays = daysUntil(connection.properties.expirationTime);
  if (expiryDays !== null && expiryDays <= 30) {
    return {
      level: 'warning',
      label: expiryDays < 0 ? 'Expired' : 'Expiring soon',
      detail: expiryDays < 0
        ? `Expired ${Math.abs(expiryDays)} day${Math.abs(expiryDays) === 1 ? '' : 's'} ago.`
        : `Expires in ${expiryDays} day${expiryDays === 1 ? '' : 's'}.`,
    };
  }

  if (statuses.length > 0) {
    return {
      level: 'healthy',
      label: statuses[0].status || 'Ready',
      detail: statuses.map((status) => status.target).filter(Boolean).join(', ') || 'No connection errors reported.',
    };
  }

  return {
    level: 'unknown',
    label: 'Status unavailable',
    detail: 'The admin API did not return a connection test result.',
  };
}

function getConnectorKey(value?: string): string {
  return (value ?? '').split('/').filter(Boolean).at(-1)?.toLowerCase() ?? '';
}

function withTimeout<T>(promise: Promise<T>, label: string, timeoutMs = 30_000): Promise<T> {
  return new Promise((resolve, reject) => {
    const timeout = window.setTimeout(
      () => reject(new Error(`${label} timed out after ${timeoutMs / 1000} seconds.`)),
      timeoutMs,
    );
    promise.then(
      (value) => {
        window.clearTimeout(timeout);
        resolve(value);
      },
      (error: unknown) => {
        window.clearTimeout(timeout);
        reject(error);
      },
    );
  });
}

function getConnectionService(connection: Connection, connectors: Connection[]): {
  name: string;
  id: string;
} {
  const apiId = connection.properties.apiId ?? '';
  const apiKey = getConnectorKey(apiId);
  const connector = connectors.find((candidate) => {
    const identifiers = [candidate.name, candidate.id, candidate.properties.apiId];
    return identifiers.some((identifier) => {
      const candidateKey = getConnectorKey(identifier);
      return candidateKey !== '' && candidateKey === apiKey;
    });
  });

  return {
    name: connector?.properties.displayName || apiKey || 'Service not reported',
    id: apiKey || 'No connector identifier',
  };
}

function StatusIcon({ level }: { level: HealthLevel }): ReactElement {
  if (level === 'healthy') return <CheckmarkCircleRegular />;
  if (level === 'warning') return <WarningRegular />;
  if (level === 'critical') return <ErrorCircleRegular />;
  return <ClockRegular />;
}

const USAGE_CACHE_KEY = 'control-hub:connector-usage:v3';
const USAGE_CATEGORY_LABELS: Record<ConnectorUsageCategory, string> = {
  apps: 'Apps',
  flows: 'Flows',
  agents: 'Agents',
};

function readUsageCache(): ConnectorUsageScanResult | null {
  const cached = localStorage.getItem(USAGE_CACHE_KEY);
  if (!cached) return null;
  try {
    return JSON.parse(cached) as ConnectorUsageScanResult;
  } catch {
    localStorage.removeItem(USAGE_CACHE_KEY);
    return null;
  }
}

export default function ConnectorsView({ environments, resources }: ConnectorsViewProps): ReactElement {
  const styles = useStyles();
  const [selectedHealthEnvironmentId, setSelectedHealthEnvironmentId] = useState(ALL_ENVIRONMENTS);
  const [selectedRegistryEnvironmentId, setSelectedRegistryEnvironmentId] = useState('');
  const [activeTab, setActiveTab] = useState<ConnectorsTab>('health');
  const [connectionsByEnvironment, setConnectionsByEnvironment] = useState<Record<string, Connection[]>>({});
  const [connectorsByEnvironment, setConnectorsByEnvironment] = useState<Record<string, Connection[]>>({});
  const [loadingKey, setLoadingKey] = useState<string | null>(null);
  const [errorByKey, setErrorByKey] = useState<Record<string, string>>({});
  const [healthLoadErrors, setHealthLoadErrors] = useState<Record<string, string>>({});
  const [healthProgress, setHealthProgress] = useState({ completed: 0, total: 0 });
  const [healthLoaded, setHealthLoaded] = useState(false);
  const [confirmDeleteConnection, setConfirmDeleteConnection] = useState<PendingConnectionDeletion | null>(null);
  const [pendingConnectionId, setPendingConnectionId] = useState<string | null>(null);
  const [publisherFilter, setPublisherFilter] = useState('');
  const [tierFilter, setTierFilter] = useState('');
  const [usageResult, setUsageResult] = useState<ConnectorUsageScanResult | null>(readUsageCache);
  const [usageProgress, setUsageProgress] = useState({ completed: 0, total: 0 });
  const [isScanningUsage, setIsScanningUsage] = useState(false);
  const [usageError, setUsageError] = useState('');

  useEffect(() => {
    if (usageResult && usageResult.source !== 'inventory-v2') {
      setUsageResult(null);
    }
  }, [usageResult]);

  const { execute: execDeleteConnection } = useMutation(deleteConnection, {
    successMessage: 'Connection deleted.',
    onSuccess: () => {
      if (confirmDeleteConnection) {
        const { connection, environmentId } = confirmDeleteConnection;
        setConnectionsByEnvironment((current) => ({
          ...current,
          [environmentId]: (current[environmentId] ?? []).filter(
            (candidate) => candidate.id !== connection.id,
          ),
        }));
      }
      setPendingConnectionId(null);
      setConfirmDeleteConnection(null);
    },
    onError: () => setPendingConnectionId(null),
  });

  const environmentOptions = useMemo(() => (
    environments
      .map((environment) => ({
        id: environment.name,
        label: getEnvironmentLabel(environment),
      }))
      .sort((left, right) => left.label.localeCompare(right.label))
  ), [environments]);
  const environmentLabelById = useMemo(
    () => new Map(environmentOptions.map((environment) => [environment.id, environment.label])),
    [environmentOptions],
  );

  useEffect(() => {
    if (!environmentOptions.length) {
      setSelectedRegistryEnvironmentId('');
      return;
    }
    if (
      !selectedRegistryEnvironmentId
      || !environmentOptions.some((option) => option.id === selectedRegistryEnvironmentId)
    ) {
      setSelectedRegistryEnvironmentId(environmentOptions[0].id);
    }
  }, [environmentOptions, selectedRegistryEnvironmentId]);

  const selectedEnvironmentLabel = useMemo(
    () => selectedHealthEnvironmentId === ALL_ENVIRONMENTS
      ? 'All environments'
      : environmentLabelById.get(selectedHealthEnvironmentId) ?? '',
    [environmentLabelById, selectedHealthEnvironmentId],
  );

  const loadSelectedData = useCallback(async () => {
    if (activeTab === 'usage') return;

    if (activeTab === 'health') {
      if (healthLoaded || environmentOptions.length === 0) return;
      const cacheKey = 'health:tenant';
      if (loadingKey === cacheKey) return;
      let cursor = 0;
      let completed = 0;
      const errors: Record<string, string> = {};

      setLoadingKey(cacheKey);
      setHealthProgress({ completed: 0, total: environmentOptions.length });
      setHealthLoadErrors({});

      async function worker(): Promise<void> {
        while (cursor < environmentOptions.length) {
          const environment = environmentOptions[cursor++];
          try {
            const [connectionsResult, connectorsResult] = await Promise.allSettled([
              withTimeout(
                fetchConnections(environment.id),
                `Connections for ${environment.label}`,
              ),
              withTimeout(
                fetchConnectors(environment.id),
                `Connector catalog for ${environment.label}`,
              ),
            ]);
            if (connectionsResult.status === 'rejected') throw connectionsResult.reason;
            setConnectionsByEnvironment((current) => ({
              ...current,
              [environment.id]: connectionsResult.value,
            }));
            setConnectorsByEnvironment((current) => ({
              ...current,
              [environment.id]: connectorsResult.status === 'fulfilled'
                ? connectorsResult.value
                : [],
            }));
          } catch (error: unknown) {
            errors[environment.id] = error instanceof Error
              ? error.message
              : 'Connection health could not be loaded.';
          } finally {
            completed += 1;
            setHealthProgress({ completed, total: environmentOptions.length });
          }
        }
      }

      await Promise.all(
        Array.from({ length: Math.min(6, environmentOptions.length) }, () => worker()),
      );
      setHealthLoadErrors(errors);
      setHealthLoaded(true);
      setLoadingKey((current) => (current === cacheKey ? null : current));
      return;
    }

    if (!selectedRegistryEnvironmentId) return;
    const cacheKey = `connectors:${selectedRegistryEnvironmentId}`;
    if (loadingKey === cacheKey) return;
    if (selectedRegistryEnvironmentId in connectorsByEnvironment) return;
    setLoadingKey(cacheKey);
    setErrorByKey((current) => ({ ...current, [cacheKey]: '' }));
    try {
      const connectors = await fetchConnectors(selectedRegistryEnvironmentId);
      setConnectorsByEnvironment((current) => ({
        ...current,
        [selectedRegistryEnvironmentId]: connectors,
      }));
    } catch (error: unknown) {
      setErrorByKey((current) => ({
        ...current,
        [cacheKey]: error instanceof Error ? error.message : 'Failed to load connectivity health.',
      }));
    } finally {
      setLoadingKey((current) => (current === cacheKey ? null : current));
    }
  }, [
    activeTab,
    connectorsByEnvironment,
    environmentOptions,
    healthLoaded,
    loadingKey,
    selectedRegistryEnvironmentId,
  ]);

  useEffect(() => {
    void loadSelectedData();
  }, [loadSelectedData]);

  const environmentScopeKey = environmentOptions.map((environment) => environment.id).join('|');
  useEffect(() => {
    setHealthLoaded(false);
  }, [environmentScopeKey]);

  useEffect(() => {
    setPublisherFilter('');
    setTierFilter('');
  }, [selectedRegistryEnvironmentId]);

  const connectors = connectorsByEnvironment[selectedRegistryEnvironmentId] ?? [];
  const connectionRows = useMemo(
    () => Object.entries(connectionsByEnvironment).flatMap(([environmentId, environmentConnections]) => {
      if (
        selectedHealthEnvironmentId !== ALL_ENVIRONMENTS
        && environmentId !== selectedHealthEnvironmentId
      ) return [];
      const environmentConnectors = connectorsByEnvironment[environmentId] ?? [];
      return environmentConnections.map((connection) => ({
        connection,
        environmentId,
        environmentLabel: environmentLabelById.get(environmentId) ?? environmentId,
        health: getConnectionHealth(connection),
        service: getConnectionService(connection, environmentConnectors),
      }));
    }).sort((left, right) => {
      const healthOrder: Record<HealthLevel, number> = {
        critical: 0,
        warning: 1,
        unknown: 2,
        healthy: 3,
      };
      return healthOrder[left.health.level] - healthOrder[right.health.level]
        || left.environmentLabel.localeCompare(right.environmentLabel)
        || (left.connection.properties.displayName || left.connection.name)
          .localeCompare(right.connection.properties.displayName || right.connection.name);
    }),
    [
      connectionsByEnvironment,
      connectorsByEnvironment,
      environmentLabelById,
      selectedHealthEnvironmentId,
    ],
  );
  const healthCounts = useMemo(() => connectionRows.reduce(
    (counts, row) => ({ ...counts, [row.health.level]: counts[row.health.level] + 1 }),
    { healthy: 0, warning: 0, critical: 0, unknown: 0 } as Record<HealthLevel, number>,
  ), [connectionRows]);
  const connectorTiers = useMemo(
    () => Array.from(new Set(
      connectors
        .map((connector) => connector.properties.tier)
        .filter((tier): tier is string => Boolean(tier)),
    )).sort(),
    [connectors],
  );
  const displayedConnectors = useMemo(() => connectors.filter((connector) => {
    const publisher = (connector.properties.publisher ?? '').toLowerCase();
    const matchesPublisher = !publisherFilter
      || (publisherFilter === 'microsoft' && publisher.includes('microsoft'))
      || (publisherFilter === 'thirdparty' && !publisher.includes('microsoft'));
    return matchesPublisher && (!tierFilter || connector.properties.tier === tierFilter);
  }), [connectors, publisherFilter, tierFilter]);

  const cacheKey = activeTab === 'connectors' && selectedRegistryEnvironmentId
    ? `connectors:${selectedRegistryEnvironmentId}`
    : activeTab === 'health'
      ? 'health:tenant'
      : '';
  const isLoading = Boolean(cacheKey) && loadingKey === cacheKey;
  const currentError = cacheKey ? errorByKey[cacheKey] : '';
  const failedEnvironmentCount = Object.keys(healthLoadErrors).length;
  const monitoredEnvironmentCount = Object.keys(connectionsByEnvironment).length;
  const attentionCount = healthCounts.critical + healthCounts.warning;
  const healthSummary = attentionCount > 0
    ? `${attentionCount} need attention`
    : healthCounts.unknown > 0
      ? `${healthCounts.unknown} unverified`
      : 'No active alerts';
  const usageTargetCount = getConnectorUsageTargetCount(resources);

  async function runUsageScan(): Promise<void> {
    setIsScanningUsage(true);
    setUsageError('');
    setUsageProgress({ completed: 0, total: usageTargetCount });
    try {
      const result = await scanTenantConnectorUsage(resources, (completed, total) => {
        setUsageProgress({ completed, total });
      });
      setUsageResult(result);
      localStorage.setItem(USAGE_CACHE_KEY, JSON.stringify(result));
    } catch (error: unknown) {
      setUsageError(error instanceof Error ? error.message : 'The connector usage scan failed.');
    } finally {
      setIsScanningUsage(false);
    }
  }

  const pageTitle = activeTab === 'usage' ? 'Tenant connector usage' : 'Connection health';
  const pageDescription = activeTab === 'usage'
    ? 'Rank the connectors used across apps, flows, and agents in every inventoried environment.'
    : 'Find authentication failures and expiring connections before apps and flows are disrupted.';

  return (
    <div className={styles.root}>
      <PageHeader
        title={pageTitle}
        description={pageDescription}
        actions={activeTab === 'usage' ? (
          <Button
            appearance="primary"
            icon={<ArrowSyncRegular />}
            disabled={isScanningUsage || usageTargetCount === 0}
            onClick={() => void runUsageScan()}
          >
            {usageResult ? 'Scan again' : 'Scan tenant'}
          </Button>
        ) : (
          <>
            <Badge
              appearance="outline"
              color={attentionCount > 0 ? 'warning' : healthCounts.unknown > 0 ? 'informative' : 'success'}
            >
              {healthSummary}
            </Badge>
            <Button
              appearance="primary"
              icon={<ArrowSyncRegular />}
              disabled={isLoading || environmentOptions.length === 0}
              onClick={() => {
                setHealthLoadErrors({});
                setHealthLoaded(false);
              }}
            >
              Refresh tenant
            </Button>
          </>
        )}
      />

      <div className={styles.controls}>
        {activeTab === 'health' && (
          <Dropdown
            className={styles.environmentSelect}
            placeholder="Filter by environment"
            value={selectedEnvironmentLabel || undefined}
            selectedOptions={[selectedHealthEnvironmentId]}
            onOptionSelect={(_, data) => (
              setSelectedHealthEnvironmentId(data.optionValue ?? ALL_ENVIRONMENTS)
            )}
          >
            <Option value={ALL_ENVIRONMENTS}>All environments</Option>
            {environmentOptions.map((environment) => (
              <Option key={environment.id} value={environment.id}>
                {environment.label}
              </Option>
            ))}
          </Dropdown>
        )}
        {activeTab === 'connectors' && (
          <Dropdown
            className={styles.environmentSelect}
            placeholder="Select an environment"
            value={environmentLabelById.get(selectedRegistryEnvironmentId) || undefined}
            selectedOptions={selectedRegistryEnvironmentId ? [selectedRegistryEnvironmentId] : []}
            onOptionSelect={(_, data) => setSelectedRegistryEnvironmentId(data.optionValue ?? '')}
          >
            {environmentOptions.map((environment) => (
              <Option key={environment.id} value={environment.id}>
                {environment.label}
              </Option>
            ))}
          </Dropdown>
        )}
        <TabList
          className={styles.tabs}
          selectedValue={activeTab}
          onTabSelect={(_, data) => setActiveTab(data.value as ConnectorsTab)}
        >
          <Tab value="health" icon={<PlugConnectedRegular />}>Health</Tab>
          <Tab value="usage" icon={<ArrowSyncRegular />}>Tenant usage</Tab>
          <Tab value="connectors" icon={<TableRegular />}>Connector registry</Tab>
        </TabList>
      </div>

      <div className={styles.body}>
        {activeTab === 'usage' ? (
          <>
            {usageError && (
              <MessageBar intent="error">
                <MessageBarBody>{usageError}</MessageBarBody>
              </MessageBar>
            )}
            {isScanningUsage && (
              <div className={styles.progressBlock}>
                <Text weight="semibold">
                  Inspecting {usageTargetCount.toLocaleString()} tenant resources
                </Text>
                <ProgressBar
                  value={usageProgress.total > 0 ? usageProgress.completed / usageProgress.total : 0}
                />
                <Text className={styles.secondary}>
                  {usageProgress.completed.toLocaleString()} of {usageProgress.total.toLocaleString()} inspected
                </Text>
              </div>
            )}
            {!isScanningUsage && !usageResult ? (
              <div className={styles.list}>
                <div className={styles.empty}>
                  <EmptyState
                    icon={<ArrowSyncRegular />}
                    title="Run the tenant usage scan"
                    subtitle={`Inspect ${usageTargetCount.toLocaleString()} apps, flows, and agents on demand. Results are cached in this browser until you scan again.`}
                    action={{ label: 'Scan tenant', onClick: () => void runUsageScan() }}
                  />
                </div>
              </div>
            ) : usageResult && (
              <>
                <div className={styles.usageIntro}>
                  <div className={styles.usageMeta}>
                    <Text weight="semibold">Tenant inventory connector usage</Text>
                    <Text className={styles.secondary}>
                      Inventory snapshot {formatDate(usageResult.scannedAt)} · counts represent resources using each connector
                    </Text>
                  </div>
                  <Badge appearance="outline" color="informative">
                    Cached result
                  </Badge>
                </div>
                <div className={styles.usageGrid}>
                  {(Object.keys(USAGE_CATEGORY_LABELS) as ConnectorUsageCategory[]).map((category) => {
                    const entries = usageResult.rankings[category].slice(0, 20);
                    const maxCount = entries[0]?.resourceCount ?? 1;
                    const coverage = usageResult.coverage[category];
                    return (
                      <section className={styles.usageColumn} key={category}>
                        <div className={styles.usageColumnHeader}>
                          <Text className={styles.signalTitle}>{USAGE_CATEGORY_LABELS[category]}</Text>
                          <Text className={styles.secondary}>
                            {coverage.scanned.toLocaleString()} scanned
                            {coverage.failed > 0 ? ` · ${coverage.failed.toLocaleString()} unavailable` : ''}
                          </Text>
                        </div>
                        <div className={styles.usageRows}>
                          {entries.length === 0 ? (
                            <div className={styles.empty}>
                              <Text className={styles.secondary}>No connector references found.</Text>
                            </div>
                          ) : entries.map((entry, index) => (
                            <div className={styles.usageRow} key={entry.id}>
                              <Text className={styles.usageRank}>{index + 1}</Text>
                              <div className={styles.usageName}>
                                <Text className={styles.primary} title={entry.id}>{entry.name}</Text>
                                <div className={styles.usageBarTrack}>
                                  <div
                                    className={styles.usageBar}
                                    style={{ width: `${Math.max(3, (entry.resourceCount / maxCount) * 100)}%` }}
                                  />
                                </div>
                              </div>
                              <Text className={styles.usageCount}>{entry.resourceCount.toLocaleString()}</Text>
                            </div>
                          ))}
                        </div>
                      </section>
                    );
                  })}
                </div>
              </>
            )}
          </>
        ) : currentError ? (
          <MessageBar intent="error">
            <MessageBarBody>{currentError}</MessageBarBody>
          </MessageBar>
        ) : isLoading && activeTab === 'connectors' ? (
          <OperationsSkeleton />
        ) : activeTab === 'connectors' && !selectedRegistryEnvironmentId ? (
          <EmptyState
            icon={<PlugConnectedRegular />}
            title="No environment selected"
            subtitle={environments.length === 0
              ? 'No environments are available from the inventory API.'
              : 'Select an environment to inspect connection health.'}
          />
        ) : activeTab === 'health' ? (
          <>
            {isLoading && (
              <div className={styles.progressBlock}>
                <Text weight="semibold">Loading tenant connection health</Text>
                <ProgressBar
                  value={healthProgress.total > 0
                    ? healthProgress.completed / healthProgress.total
                    : 0}
                />
                <Text className={styles.secondary}>
                  {healthProgress.completed.toLocaleString()} of{' '}
                  {healthProgress.total.toLocaleString()} environments inspected
                </Text>
              </div>
            )}
            {failedEnvironmentCount > 0 && (
              <MessageBar intent="warning">
                <MessageBarBody>
                  Connection health could not be loaded for {failedEnvironmentCount}{' '}
                  environment{failedEnvironmentCount === 1 ? '' : 's'}. Results from the remaining
                  environments are shown.
                </MessageBarBody>
              </MessageBar>
            )}
            <div className={styles.signalRail}>
              <div className={styles.signalLead}>
                <Text className={styles.signalTitle}>{selectedEnvironmentLabel}</Text>
                <Text className={styles.signalDescription}>
                  {connectionRows.length} connection{connectionRows.length === 1 ? '' : 's'} across{' '}
                  {selectedHealthEnvironmentId === ALL_ENVIRONMENTS
                    ? `${monitoredEnvironmentCount} environments`
                    : 'this environment'}
                </Text>
              </div>
              <div className={styles.signalCell}>
                <Text className={styles.signalValue}>{healthCounts.critical}</Text>
                <Text className={styles.signalLabel}>Failed</Text>
              </div>
              <div className={styles.signalCell}>
                <Text className={styles.signalValue}>{healthCounts.warning}</Text>
                <Text className={styles.signalLabel}>Expiring</Text>
              </div>
              <div className={styles.signalCell}>
                <Text className={styles.signalValue}>{healthCounts.healthy}</Text>
                <Text className={styles.signalLabel}>Ready</Text>
              </div>
            </div>

            {connectionRows.length === 0 ? (
              <div className={styles.list}>
                <div className={styles.empty}>
                  <EmptyState
                    icon={<PlugConnectedRegular />}
                    title="No connections found"
                    subtitle={selectedHealthEnvironmentId === ALL_ENVIRONMENTS
                      ? 'No connection instances were returned for the tenant.'
                      : 'This environment has no connection instances to monitor.'}
                  />
                </div>
              </div>
            ) : (
              <div className={styles.list}>
                <div className={styles.listHeader}>
                  <span>Connection</span>
                  <span>Service</span>
                  <span>Environment</span>
                  <span>Health</span>
                  <span>Signal</span>
                  <span>Expiry</span>
                  <span>Action</span>
                </div>
                {connectionRows.map(({
                  connection,
                  environmentId,
                  environmentLabel,
                  health,
                  service,
                }) => (
                  <div className={styles.healthRow} key={`${environmentId}:${connection.id}`}>
                    <div className={styles.identity}>
                      <Text className={styles.primary}>{connection.properties.displayName || connection.name}</Text>
                      <Text className={styles.secondary}>
                        {connection.properties.createdBy?.email
                          ?? connection.properties.createdBy?.displayName
                          ?? connection.properties.apiId
                          ?? 'Owner not reported'}
                      </Text>
                    </div>
                    <div className={styles.identity}>
                      <Text className={styles.primary}>{service.name}</Text>
                      <Text className={styles.secondary}>{service.id}</Text>
                    </div>
                    <div className={styles.identity}>
                      <Text className={styles.primary}>{environmentLabel}</Text>
                      <Text className={styles.secondary}>{environmentId}</Text>
                    </div>
                    <div className={`${styles.healthStatus} ${styles[`status${health.level[0].toUpperCase()}${health.level.slice(1)}` as keyof typeof styles]}`}>
                      <StatusIcon level={health.level} />
                      <Text>{health.label}</Text>
                    </div>
                    <Text className={styles.secondary} title={health.detail}>{health.detail}</Text>
                    <div className={styles.expiry}>
                      <Text>{formatDate(connection.properties.expirationTime)}</Text>
                      <Text className={styles.secondary}>
                        {connection.properties.isSsoConnection ? 'Single sign-on' : 'Standard authentication'}
                      </Text>
                    </div>
                    <Button
                      appearance="subtle"
                      size="small"
                      icon={<DeleteRegular />}
                      aria-label={`Delete ${connection.properties.displayName || connection.name}`}
                      title="Delete unusable connection"
                      disabled={pendingConnectionId === connection.id}
                      onClick={() => setConfirmDeleteConnection({ connection, environmentId })}
                    />
                  </div>
                ))}
              </div>
            )}
          </>
        ) : (
          <>
            <div className={styles.connectorToolbar}>
              <Dropdown
                placeholder="All publishers"
                value={publisherFilter === 'microsoft' ? 'Microsoft' : publisherFilter === 'thirdparty' ? 'Third party' : undefined}
                selectedOptions={publisherFilter ? [publisherFilter] : []}
                onOptionSelect={(_, data) => setPublisherFilter(data.optionValue ?? '')}
              >
                <Option value="microsoft">Microsoft</Option>
                <Option value="thirdparty">Third party</Option>
              </Dropdown>
              {connectorTiers.length > 0 && (
                <Dropdown
                  placeholder="All tiers"
                  value={tierFilter || undefined}
                  selectedOptions={tierFilter ? [tierFilter] : []}
                  onOptionSelect={(_, data) => setTierFilter(data.optionValue ?? '')}
                >
                  {connectorTiers.map((tier) => <Option key={tier} value={tier}>{tier}</Option>)}
                </Dropdown>
              )}
            </div>
            <div className={styles.list}>
              <table className={styles.table}>
                <thead>
                  <tr>
                    <th className={styles.th}>Connector</th>
                    <th className={styles.th}>Publisher</th>
                    <th className={styles.th}>Tier</th>
                    <th className={styles.th}>Created</th>
                  </tr>
                </thead>
                <tbody>
                  {displayedConnectors.map((connector) => (
                    <tr key={connector.id}>
                      <td className={styles.td}>{connector.properties.displayName || connector.name}</td>
                      <td className={styles.td}>{connector.properties.publisher ?? 'Not reported'}</td>
                      <td className={styles.td}>
                        {connector.properties.isCustomApi ? 'Custom API' : connector.properties.tier ?? 'Standard'}
                      </td>
                      <td className={styles.td}>{formatDate(connector.properties.createdTime)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </>
        )}
      </div>

      <ConfirmDialog
        open={Boolean(confirmDeleteConnection)}
        title="Delete connection"
        message={`Delete connection "${confirmDeleteConnection?.connection.properties.displayName}" from "${confirmDeleteConnection ? environmentLabelById.get(confirmDeleteConnection.environmentId) : ''}"? This action cannot be undone.`}
        confirmLabel="Delete"
        isDangerous
        isLoading={pendingConnectionId !== null}
        onConfirm={() => {
          if (confirmDeleteConnection) {
            setPendingConnectionId(confirmDeleteConnection.connection.id);
            void execDeleteConnection(confirmDeleteConnection.connection.id);
          }
        }}
        onCancel={() => setConfirmDeleteConnection(null)}
      />
    </div>
  );
}
