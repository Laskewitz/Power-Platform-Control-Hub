import { useCallback, useEffect, useMemo, useState } from 'react';
import type { ReactElement } from 'react';
import {
  Badge,
  Button,
  Dropdown,
  Link,
  MessageBar,
  MessageBarBody,
  Option,
  Spinner,
  Tab,
  TabList,
  Text,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import { DeleteRegular, GlobeRegular, LinkRegular, PlugConnectedRegular, TableRegular } from '@fluentui/react-icons';
import {
  fetchConnections,
  fetchConnectors,
  fetchPowerPagesWebsites,
} from '../services/adminApi.ts';
import { deleteConnection } from '../services/connectorMutations.ts';
import type { Connection, PowerPagesWebsite } from '../types/admin.ts';
import type { Resource } from '../types/inventory.ts';
import ConfirmDialog from './ConfirmDialog.tsx';
import { useMutation } from '../hooks/useMutation.tsx';

interface ConnectorsViewProps {
  environments: Resource[];
}

type ConnectorsTab = 'connections' | 'connectors' | 'websites';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalL,
    padding: tokens.spacingHorizontalXL,
    height: '100%',
    overflow: 'hidden',
  },
  toolbar: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    flexWrap: 'wrap',
    flexShrink: 0,
  },
  title: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
    marginRight: 'auto',
  },
  body: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
    minHeight: 0,
    flex: 1,
  },
  tableWrapper: {
    flex: 1,
    overflowY: 'auto',
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  table: {
    width: '100%',
    borderCollapse: 'collapse',
  },
  th: {
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    textAlign: 'left',
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground3,
    textTransform: 'uppercase',
    letterSpacing: '0.04em',
    borderBottom: `2px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground3,
    whiteSpace: 'nowrap',
  },
  td: {
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    fontSize: tokens.fontSizeBase300,
    color: tokens.colorNeutralForeground1,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    verticalAlign: 'middle',
  },
  centered: {
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    height: '100%',
  },
});

function formatDate(value?: string): string {
  if (!value) return '—';
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? value : date.toLocaleString();
}

function getEnvironmentLabel(environment: Resource): string {
  return environment.properties.displayName ?? environment.name;
}

function getWebsiteBadgeColor(type: PowerPagesWebsite['type']): 'warning' | 'brand' {
  return type === 'Trial' ? 'warning' : 'brand';
}

export default function ConnectorsView({
  environments,
}: ConnectorsViewProps): ReactElement {
  const styles = useStyles();
  const [selectedEnvironmentId, setSelectedEnvironmentId] = useState('');
  const [activeTab, setActiveTab] = useState<ConnectorsTab>('connections');
  const [connectionsByEnvironment, setConnectionsByEnvironment] = useState<Record<string, Connection[]>>({});
  const [connectorsByEnvironment, setConnectorsByEnvironment] = useState<Record<string, Connection[]>>({});
  const [websitesByEnvironment, setWebsitesByEnvironment] = useState<Record<string, PowerPagesWebsite[]>>({});
  const [loadingKey, setLoadingKey] = useState<string | null>(null);
  const [errorByKey, setErrorByKey] = useState<Record<string, string>>({});
  const [confirmDeleteConnection, setConfirmDeleteConnection] = useState<Connection | null>(null);
  const [pendingConnectionId, setPendingConnectionId] = useState<string | null>(null);
  const [publisherFilter, setPublisherFilter] = useState('');
  const [tierFilter, setTierFilter] = useState('');

  const { execute: execDeleteConnection } = useMutation(deleteConnection, {
    successMessage: 'Connection deleted.',
    onSuccess: () => {
      if (confirmDeleteConnection) {
        setConnectionsByEnvironment((prev) => ({
          ...prev,
          [selectedEnvironmentId]: (prev[selectedEnvironmentId] ?? []).filter(
            (c) => c.id !== confirmDeleteConnection.id,
          ),
        }));
      }
      setPendingConnectionId(null);
      setConfirmDeleteConnection(null);
    },
    onError: () => setPendingConnectionId(null),
  });

  const environmentOptions = useMemo(() => environments.map((environment) => ({
    id: environment.name,
    label: getEnvironmentLabel(environment),
  })), [environments]);

  useEffect(() => {
    if (!environmentOptions.length) {
      setSelectedEnvironmentId('');
      return;
    }

    if (!selectedEnvironmentId || !environmentOptions.some((option) => option.id === selectedEnvironmentId)) {
      setSelectedEnvironmentId(environmentOptions[0].id);
    }
  }, [environmentOptions, selectedEnvironmentId]);

  const selectedEnvironmentLabel = useMemo(
    () => environmentOptions.find((option) => option.id === selectedEnvironmentId)?.label ?? '',
    [environmentOptions, selectedEnvironmentId],
  );

  const loadSelectedData = useCallback(async () => {
    if (!selectedEnvironmentId) return;

    const cacheKey = `${activeTab}:${selectedEnvironmentId}`;
    const hasCachedData = activeTab === 'connections'
      ? selectedEnvironmentId in connectionsByEnvironment
      : activeTab === 'connectors'
        ? selectedEnvironmentId in connectorsByEnvironment
        : selectedEnvironmentId in websitesByEnvironment;

    if (hasCachedData) return;

    setLoadingKey(cacheKey);
    setErrorByKey((current) => ({ ...current, [cacheKey]: '' }));

    try {
      if (activeTab === 'connections') {
        const result = await fetchConnections(selectedEnvironmentId);
        setConnectionsByEnvironment((current) => ({ ...current, [selectedEnvironmentId]: result }));
      } else if (activeTab === 'connectors') {
        const result = await fetchConnectors(selectedEnvironmentId);
        setConnectorsByEnvironment((current) => ({ ...current, [selectedEnvironmentId]: result }));
      } else {
        const result = await fetchPowerPagesWebsites(selectedEnvironmentId);
        setWebsitesByEnvironment((current) => ({ ...current, [selectedEnvironmentId]: result }));
      }
    } catch (e: unknown) {
      setErrorByKey((current) => ({
        ...current,
        [cacheKey]: e instanceof Error ? e.message : 'Failed to load environment admin data.',
      }));
    } finally {
      setLoadingKey((current) => (current === cacheKey ? null : current));
    }
  }, [
    activeTab,
    connectionsByEnvironment,
    connectorsByEnvironment,
    selectedEnvironmentId,
    websitesByEnvironment,
  ]);

  useEffect(() => {
    void loadSelectedData();
  }, [loadSelectedData]);

  const currentError = selectedEnvironmentId
    ? errorByKey[`${activeTab}:${selectedEnvironmentId}`] ?? null
    : null;
  const isLoading = selectedEnvironmentId
    ? loadingKey === `${activeTab}:${selectedEnvironmentId}`
    : false;

  const currentConnections = selectedEnvironmentId
    ? activeTab === 'connections'
      ? connectionsByEnvironment[selectedEnvironmentId] ?? []
      : connectorsByEnvironment[selectedEnvironmentId] ?? []
    : [];
  const currentWebsites = selectedEnvironmentId
    ? websitesByEnvironment[selectedEnvironmentId] ?? []
    : [];

  // Reset filters when environment or tab changes
  useEffect(() => {
    setPublisherFilter('');
    setTierFilter('');
  }, [selectedEnvironmentId, activeTab]);

  const connectorPublishers = useMemo(() => {
    const pubs = new Set<string>();
    for (const c of connectorsByEnvironment[selectedEnvironmentId] ?? []) {
      if (c.properties.publisher) pubs.add(c.properties.publisher);
    }
    return Array.from(pubs).sort();
  }, [connectorsByEnvironment, selectedEnvironmentId]);

  const connectorTiers = useMemo(() => {
    const tiers = new Set<string>();
    for (const c of connectorsByEnvironment[selectedEnvironmentId] ?? []) {
      if (c.properties.tier) tiers.add(c.properties.tier);
    }
    return Array.from(tiers).sort();
  }, [connectorsByEnvironment, selectedEnvironmentId]);

  const displayedConnections = useMemo(() => {
    if (activeTab !== 'connectors') return currentConnections;
    return currentConnections.filter((c) => {
      const matchesPublisher = !publisherFilter || c.properties.publisher === publisherFilter;
      const matchesTier = !tierFilter || c.properties.tier === tierFilter;
      return matchesPublisher && matchesTier;
    });
  }, [activeTab, currentConnections, publisherFilter, tierFilter]);

  const selectedEnvironmentValue = selectedEnvironmentLabel || undefined;

  const content = useMemo(() => {
    if (!selectedEnvironmentId) {
      return (
        <div className={styles.centered}>
          <Text>No environments available.</Text>
        </div>
      );
    }

    if (isLoading) {
      return (
        <div className={styles.centered}>
          <Spinner size="extra-large" label={`Loading ${activeTab}…`} />
        </div>
      );
    }

    if (currentError) {
      return (
        <MessageBar intent="error">
          <MessageBarBody>{currentError}</MessageBarBody>
        </MessageBar>
      );
    }

    if (activeTab === 'websites') {
      return (
        <div className={styles.tableWrapper}>
          <table className={styles.table}>
            <thead>
              <tr>
                <th className={styles.th}>Name</th>
                <th className={styles.th}>URL</th>
                <th className={styles.th}>Type</th>
                <th className={styles.th}>Visibility</th>
                <th className={styles.th}>Package Version</th>
                <th className={styles.th}>Created On</th>
              </tr>
            </thead>
            <tbody>
              {currentWebsites.length === 0 ? (
                <tr>
                  <td className={styles.td} colSpan={6} style={{ textAlign: 'center' }}>
                    No Power Pages websites found for this environment.
                  </td>
                </tr>
              ) : (
                currentWebsites.map((website) => (
                  <tr key={website.id}>
                    <td className={styles.td}>{website.name}</td>
                    <td className={styles.td}>
                      <Link href={website.websiteUrl} target="_blank" rel="noreferrer">
                        {website.websiteUrl}
                      </Link>
                    </td>
                    <td className={styles.td}>
                      <Badge appearance="filled" color={getWebsiteBadgeColor(website.type)}>
                        {website.type}
                      </Badge>
                    </td>
                    <td className={styles.td}>{website.siteVisibility}</td>
                    <td className={styles.td}>{website.packageVersion ?? '—'}</td>
                    <td className={styles.td}>{formatDate(website.createdOn)}</td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      );
    }

    return (
      <div className={styles.tableWrapper}>
        <table className={styles.table}>
          <thead>
            <tr>
              <th className={styles.th}>Display Name</th>
              <th className={styles.th}>Publisher</th>
              <th className={styles.th}>Custom API</th>
              <th className={styles.th}>Created</th>
              <th className={styles.th}>Type (Tier)</th>
              {activeTab === 'connections' && <th className={styles.th} style={{ width: '60px' }}></th>}
            </tr>
          </thead>
          <tbody>
            {currentConnections.length === 0 ? (
              <tr>
                <td className={styles.td} colSpan={activeTab === 'connections' ? 6 : 5} style={{ textAlign: 'center' }}>
                  No {activeTab} found for this environment.
                </td>
              </tr>
            ) : displayedConnections.length === 0 ? (
              <tr>
                <td className={styles.td} colSpan={activeTab === 'connections' ? 6 : 5} style={{ textAlign: 'center' }}>
                  No connectors match the selected filters.
                </td>
              </tr>
            ) : (
              displayedConnections.map((item) => (
                <tr key={item.id}>
                  <td className={styles.td}>{item.properties.displayName}</td>
                  <td className={styles.td}>{item.properties.publisher ?? '—'}</td>
                  <td className={styles.td}>
                    {item.properties.isCustomApi ? (
                      <Badge appearance="outline" color="informative">
                        Custom API
                      </Badge>
                    ) : '—'}
                  </td>
                  <td className={styles.td}>{formatDate(item.properties.createdTime)}</td>
                  <td className={styles.td}>{item.properties.tier ?? item.type}</td>
                  {activeTab === 'connections' && (
                    <td className={styles.td}>
                      <Button
                        appearance="subtle"
                        size="small"
                        icon={<DeleteRegular />}
                        title="Delete connection"
                        disabled={pendingConnectionId === item.id}
                        style={{ color: tokens.colorStatusDangerForeground1 }}
                        onClick={() => setConfirmDeleteConnection(item)}
                      />
                    </td>
                  )}
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>
    );
  }, [
    activeTab,
    currentConnections,
    currentError,
    currentWebsites,
    displayedConnections,
    isLoading,
    pendingConnectionId,
    selectedEnvironmentId,
    styles,
  ]);

  return (
    <div className={styles.root}>
      <div className={styles.toolbar}>
        <Text className={styles.title}>Connectors</Text>
        {activeTab === 'connectors' && connectorPublishers.length > 0 && (
          <Dropdown
            placeholder="All publishers"
            value={publisherFilter || undefined}
            selectedOptions={publisherFilter ? [publisherFilter] : []}
            onOptionSelect={(_, data) => setPublisherFilter(data.optionValue === publisherFilter ? '' : (data.optionValue ?? ''))}
            style={{ minWidth: '180px' }}
          >
            {connectorPublishers.map((pub) => (
              <Option key={pub} value={pub}>{pub}</Option>
            ))}
          </Dropdown>
        )}
        {activeTab === 'connectors' && connectorTiers.length > 0 && (
          <Dropdown
            placeholder="All tiers"
            value={tierFilter || undefined}
            selectedOptions={tierFilter ? [tierFilter] : []}
            onOptionSelect={(_, data) => setTierFilter(data.optionValue === tierFilter ? '' : (data.optionValue ?? ''))}
            style={{ minWidth: '140px' }}
          >
            {connectorTiers.map((tier) => (
              <Option key={tier} value={tier}>{tier}</Option>
            ))}
          </Dropdown>
        )}
        <Dropdown
          placeholder="Select an environment"
          value={selectedEnvironmentValue}
          selectedOptions={selectedEnvironmentId ? [selectedEnvironmentId] : []}
          onOptionSelect={(_, data) => setSelectedEnvironmentId(data.optionValue ?? '')}
          style={{ minWidth: '260px' }}
        >
          {environmentOptions.map((environment) => (
            <Option key={environment.id} value={environment.id}>
              {environment.label}
            </Option>
          ))}
        </Dropdown>
      </div>

      <div className={styles.body}>
        <TabList
          selectedValue={activeTab}
          onTabSelect={(_, data) => setActiveTab(data.value as ConnectorsTab)}
        >
          <Tab value="connections" icon={<PlugConnectedRegular />}>
            Connections
          </Tab>
          <Tab value="connectors" icon={<TableRegular />}>
            Connectors
          </Tab>
          <Tab value="websites" icon={<GlobeRegular />}>
            Power Pages Websites
          </Tab>
        </TabList>

        <Text>
          <LinkRegular style={{ marginRight: tokens.spacingHorizontalXS }} />
          {selectedEnvironmentLabel || 'Select an environment to view connectivity resources.'}
        </Text>

        {content}
      </div>

      <ConfirmDialog
        open={Boolean(confirmDeleteConnection)}
        title="Delete Connection"
        message={`Delete connection "${confirmDeleteConnection?.properties.displayName}"? This action cannot be undone.`}
        confirmLabel="Delete"
        isDangerous
        isLoading={pendingConnectionId !== null}
        onConfirm={() => {
          if (confirmDeleteConnection) {
            setPendingConnectionId(confirmDeleteConnection.id);
            void execDeleteConnection(confirmDeleteConnection.id);
          }
        }}
        onCancel={() => setConfirmDeleteConnection(null)}
      />
    </div>
  );
}
