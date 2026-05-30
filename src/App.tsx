import { useState } from 'react';
import type { ReactElement } from 'react';
import {
  FluentProvider,
  Toaster,
  webLightTheme,
  webDarkTheme,
  makeStyles,
  tokens,
  Text,
  TabList,
  Tab,
  Button,
  Spinner,
} from '@fluentui/react-components';
import {
  TableRegular,
  GlobeRegular,
  ArrowClockwiseRegular,
  LightbulbRegular,
  ShieldRegular,
  PlugConnectedRegular,
  GridRegular,
  WeatherMoonRegular,
  WeatherSunnyRegular,
  LayerRegular,
} from '@fluentui/react-icons';
import { useAdminData } from './hooks/useAdminData.ts';
import { useInventory } from './hooks/useInventory.ts';
import Dashboard from './components/Dashboard.tsx';
import ResourcesView from './components/ResourcesView.tsx';
import EnvironmentsView from './components/EnvironmentsView.tsx';
import RecommendationsView from './components/RecommendationsView.tsx';
import GovernanceView from './components/GovernanceView.tsx';
import EnvironmentGroupsView from './components/EnvironmentGroupsView.tsx';
import ConnectorsView from './components/ConnectorsView.tsx';
import ppaLogo from './assets/ppa-logo.png?inline';

type TabValue = 'overview' | 'resources' | 'environments' | 'recommendations' | 'governance' | 'envgroups' | 'connectors';
const VALID_TABS: readonly TabValue[] = ['overview', 'resources', 'environments', 'recommendations', 'governance', 'envgroups', 'connectors'];

const useStyles = makeStyles({
  shell: {
    display: 'flex',
    flexDirection: 'column',
    height: '100vh',
    width: '100vw',
    backgroundColor: tokens.colorNeutralBackground2,
    overflow: 'hidden',
  },
  nav: {
    display: 'flex',
    alignItems: 'center',
    backgroundColor: tokens.colorNeutralBackground1,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    padding: `0 ${tokens.spacingHorizontalL}`,
    flexShrink: 0,
    gap: tokens.spacingHorizontalS,
  },
  navTabs: {
    flex: 1,
  },
  navRight: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    flexShrink: 0,
  },
  content: {
    flex: 1,
    overflow: 'hidden',
    backgroundColor: tokens.colorNeutralBackground2,
  },
  footer: {
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalL}`,
    backgroundColor: tokens.colorNeutralBackground2,
    flexShrink: 0,
  },
  footerLogo: {
    height: '16px',
    width: 'auto',
  },
  footerText: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
  },
});

export default function App(): ReactElement {
  const styles = useStyles();
  const [activeTab, setActiveTab] = useState<TabValue>('overview');
  const [resourceTypeFilter, setResourceTypeFilter] = useState<string>('all');
  const [isDark, setIsDark] = useState(false);

  function navigateToResources(typeKey: string) {
    setResourceTypeFilter(typeKey);
    setActiveTab('resources');
  }

  const {
    resources,
    environments,
    counts,
    isLoading: dataLoading,
    loadingLabel,
    error: dataError,
    refresh,
  } = useInventory();

  const {
    recommendations,
    envGroups,
    billingPolicies,
    crossTenantReports,
    ruleBasedPolicies,
    ruleAssignments,
    ruleSets,
    dlpPolicies,
    isLoading: adminLoading,
    error: adminError,
    refresh: refreshAdmin,
  } = useAdminData();

  async function refreshAll(): Promise<void> {
    await Promise.all([refresh(), refreshAdmin()]);
  }

  return (
    <FluentProvider theme={isDark ? webDarkTheme : webLightTheme}>
      <div className={styles.shell}>
        <nav className={styles.nav}>
          <div className={styles.navTabs}>
            <TabList
              selectedValue={activeTab}
              onTabSelect={(_, data) => {
                const tab = data.value as string;
                if (!VALID_TABS.includes(tab as TabValue)) return;
                if (tab === 'resources') setResourceTypeFilter('all');
                setActiveTab(tab as TabValue);
              }}
            >
              <Tab value="overview" icon={<GridRegular />}>Overview</Tab>
              <Tab value="resources" icon={<TableRegular />}>Resources</Tab>
              <Tab value="environments" icon={<GlobeRegular />}>Environments</Tab>
              <Tab value="recommendations" icon={<LightbulbRegular />}>Recommendations</Tab>
              <Tab value="governance" icon={<ShieldRegular />}>Tenant Policies</Tab>
              <Tab value="envgroups" icon={<LayerRegular />}>Environment Groups</Tab>
              <Tab value="connectors" icon={<PlugConnectedRegular />}>Connectors</Tab>
            </TabList>
          </div>
          <div className={styles.navRight}>
            {(dataLoading || adminLoading) && (
              <Spinner size="tiny" label={loadingLabel ?? undefined} labelPosition="after" />
            )}
            <Button
              appearance="subtle"
              icon={isDark ? <WeatherSunnyRegular /> : <WeatherMoonRegular />}
              onClick={() => setIsDark((d) => !d)}
              title={isDark ? 'Switch to light mode' : 'Switch to dark mode'}
            />
            <Button
              appearance="subtle"
              icon={<ArrowClockwiseRegular />}
              onClick={() => void refreshAll()}
              title="Refresh data"
            />
          </div>
        </nav>

        <main className={styles.content}>
          {activeTab === 'overview' && (
            <Dashboard
              resources={resources}
              counts={counts}
              isLoading={dataLoading}
              error={dataError}
              onNavigateToResources={navigateToResources}
            />
          )}
          {activeTab === 'resources' && (
            <ResourcesView
              resources={resources}
              isLoading={dataLoading}
              error={dataError}
              onRefresh={refresh}
              initialTypeFilter={resourceTypeFilter}
            />
          )}
          {activeTab === 'environments' && (
            <EnvironmentsView
              environments={environments}
              resources={resources}
              envGroups={envGroups}
              isLoading={dataLoading}
              error={dataError}
              onRefreshEnvironments={refresh}
            />
          )}
          {activeTab === 'recommendations' && (
            <RecommendationsView
              recommendations={recommendations}
              isLoading={adminLoading}
              error={adminError}
            />
          )}
          {activeTab === 'governance' && (
            <GovernanceView
              billingPolicies={billingPolicies}
              crossTenantReports={crossTenantReports}
              dlpPolicies={dlpPolicies}
              environments={environments}
              resources={resources}
              isLoading={adminLoading}
              error={adminError}
              onRefreshAdmin={refreshAdmin}
            />
          )}
          {activeTab === 'envgroups' && (
            <EnvironmentGroupsView
              environments={environments}
              envGroups={envGroups}
              ruleBasedPolicies={ruleBasedPolicies}
              ruleAssignments={ruleAssignments}
              ruleSets={ruleSets}
              isLoading={adminLoading}
              error={adminError}
              onRefreshAdmin={refreshAdmin}
            />
          )}
          {activeTab === 'connectors' && (
            <ConnectorsView environments={environments} />
          )}
        </main>

        <footer className={styles.footer}>
          <Text className={styles.footerText}>Made with ❤️ by the Power Platform Advocates</Text>
          <img src={ppaLogo} alt="Power Platform Advocates" className={styles.footerLogo} />
        </footer>
      </div>
      <Toaster toasterId="coe-toaster" />
    </FluentProvider>
  );
}
