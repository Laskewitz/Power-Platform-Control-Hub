import { useState } from 'react';
import type { ReactElement, ReactNode } from 'react';
import {
  Button,
  FluentProvider,
  Spinner,
  Text,
  Toaster,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import {
  ArrowClockwiseRegular,
  DismissRegular,
  GlobeRegular,
  GridRegular,
  KeyRegular,
  LightbulbRegular,
  NavigationRegular,
  PlugConnectedRegular,
  ShieldRegular,
  TableRegular,
  WeatherMoonRegular,
  WeatherSunnyRegular,
} from '@fluentui/react-icons';
import { useAdminData } from './hooks/useAdminData.ts';
import { useInventory } from './hooks/useInventory.ts';
import { useLicensingData } from './hooks/useLicensingData.ts';
import Dashboard from './components/Dashboard.tsx';
import ResourcesView from './components/ResourcesView.tsx';
import EnvironmentsView from './components/EnvironmentsView.tsx';
import RecommendationsView from './components/RecommendationsView.tsx';
import GovernanceView from './components/GovernanceView.tsx';
import ConnectorsView from './components/ConnectorsView.tsx';
import LicensingView from './components/LicensingView.tsx';
import { controlHubDarkTheme, controlHubLightTheme } from './theme.ts';
import ppaLogo from './assets/ppa-logo.png?inline';

type TabValue = 'overview' | 'resources' | 'environments' | 'recommendations' | 'governance' | 'connectors' | 'licensing';

interface NavItem {
  value: TabValue;
  label: string;
  description: string;
  icon: ReactNode;
  group: 'Operate' | 'Govern';
}

const NAV_ITEMS: readonly NavItem[] = [
  { value: 'overview', label: 'Operations', description: 'Tenant posture and activity', icon: <GridRegular />, group: 'Operate' },
  { value: 'resources', label: 'Resources', description: 'Apps, flows, and agents', icon: <TableRegular />, group: 'Operate' },
  { value: 'environments', label: 'Environments', description: 'Capacity and lifecycle', icon: <GlobeRegular />, group: 'Operate' },
  { value: 'connectors', label: 'Connections', description: 'Health and tenant usage', icon: <PlugConnectedRegular />, group: 'Operate' },
  { value: 'licensing', label: 'Licensing & capacity', description: 'Entitlements and consumption', icon: <KeyRegular />, group: 'Govern' },
  { value: 'governance', label: 'Tenant policies', description: 'DLP, billing, and reports', icon: <ShieldRegular />, group: 'Govern' },
  { value: 'recommendations', label: 'Recommendations', description: 'Prioritized admin guidance', icon: <LightbulbRegular />, group: 'Govern' },
];

const useStyles = makeStyles({
  shell: {
    display: 'grid',
    gridTemplateColumns: '232px minmax(0, 1fr)',
    width: '100%',
    height: '100dvh',
    overflow: 'hidden',
    backgroundColor: '#060A0F',
    '@media (max-width: 900px)': {
      gridTemplateColumns: '1fr',
    },
  },
  sidebar: {
    position: 'relative',
    zIndex: 20,
    display: 'flex',
    flexDirection: 'column',
    minWidth: 0,
    color: '#DDEAF0',
    backgroundColor: '#070C12',
    backgroundImage: 'linear-gradient(rgba(67, 217, 255, 0.035) 1px, transparent 1px), linear-gradient(90deg, rgba(67, 217, 255, 0.035) 1px, transparent 1px)',
    backgroundSize: '28px 28px',
    borderRight: '1px solid #20313E',
    '@media (max-width: 900px)': {
      position: 'fixed',
      inset: '0 auto 0 0',
      width: 'min(86vw, 320px)',
      boxShadow: '12px 0 36px rgba(0, 0, 0, 0.28)',
      transform: 'translateX(-105%)',
      transitionProperty: 'transform',
      transitionDuration: '180ms',
      transitionTimingFunction: 'cubic-bezier(.16, 1, .3, 1)',
    },
  },
  sidebarOpen: {
    '@media (max-width: 900px)': {
      transform: 'translateX(0)',
    },
  },
  scrim: {
    display: 'none',
    '@media (max-width: 900px)': {
      display: 'block',
      position: 'fixed',
      inset: 0,
      zIndex: 10,
      border: 0,
      backgroundColor: 'rgba(3, 13, 22, 0.58)',
    },
  },
  brand: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    minHeight: '96px',
    padding: '20px 22px',
    borderBottom: '1px solid #20313E',
  },
  brandText: {
    display: 'grid',
    gap: '1px',
    minWidth: 0,
  },
  brandName: {
    color: '#FFFFFF',
    fontSize: '22px',
    fontWeight: tokens.fontWeightSemibold,
    letterSpacing: '-0.035em',
  },
  brandSub: {
    color: '#8095A4',
    fontSize: tokens.fontSizeBase200,
  },
  closeNav: {
    display: 'none',
    marginLeft: 'auto',
    color: '#FFFFFF',
    '@media (max-width: 900px)': {
      display: 'inline-flex',
    },
  },
  nav: {
    display: 'flex',
    flexDirection: 'column',
    gap: '24px',
    flex: 1,
    padding: '20px 12px',
    overflowY: 'auto',
  },
  navGroup: {
    display: 'grid',
    gap: tokens.spacingVerticalXS,
  },
  groupLabel: {
    padding: `0 ${tokens.spacingHorizontalS}`,
    color: '#667E8E',
    fontSize: tokens.fontSizeBase100,
    fontWeight: tokens.fontWeightSemibold,
    letterSpacing: '0.08em',
    textTransform: 'uppercase',
  },
  navButton: {
    display: 'grid',
    gridTemplateColumns: '22px minmax(0, 1fr) 8px',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    width: '100%',
    minHeight: '52px',
    padding: '9px 10px',
    color: '#C6D5DC',
    font: 'inherit',
    textAlign: 'left',
    border: 0,
    borderRadius: 0,
    backgroundColor: 'transparent',
    cursor: 'pointer',
    ':hover': {
      color: '#FFFFFF',
      backgroundColor: 'rgba(67, 217, 255, 0.07)',
    },
  },
  navButtonActive: {
    color: '#F4FBFD',
    backgroundColor: '#102330',
    boxShadow: 'inset 0 0 0 1px rgba(67, 217, 255, 0.28), 0 8px 20px rgba(0, 0, 0, 0.24)',
    '& span': {
      color: '#F4FBFD',
    },
  },
  navIcon: {
    display: 'grid',
    placeItems: 'center',
    fontSize: '19px',
  },
  navCopy: {
    display: 'grid',
    gap: '1px',
    minWidth: 0,
  },
  navLabel: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
  },
  navDescription: {
    color: '#8095A4',
    fontSize: tokens.fontSizeBase100,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
  },
  activeMarker: {
    width: '8px',
    height: '8px',
    justifySelf: 'center',
    backgroundColor: '#FFB547',
    border: '1px solid #070C12',
    borderRadius: '50%',
    boxShadow: '0 0 0 2px #FFB547, 0 0 12px rgba(255, 181, 71, 0.62)',
  },
  attribution: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalL}`,
    color: '#8095A4',
    borderTop: '1px solid #20313E',
    fontSize: tokens.fontSizeBase100,
  },
  attributionLogo: {
    width: '22px',
    height: '22px',
    objectFit: 'contain',
  },
  workspace: {
    display: 'grid',
    gridTemplateRows: '64px minmax(0, 1fr)',
    minWidth: 0,
    minHeight: 0,
  },
  commandBar: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    minWidth: 0,
    padding: '0 28px',
    color: '#EAF7FA',
    backgroundColor: '#080E15',
    backgroundImage: 'linear-gradient(90deg, rgba(67, 217, 255, 0.07), transparent 40%)',
    borderBottom: '1px solid #20313E',
    boxShadow: '0 8px 24px rgba(0, 0, 0, 0.28)',
  },
  menuButton: {
    display: 'none',
    '@media (max-width: 900px)': {
      display: 'inline-flex',
    },
  },
  routeMeta: {
    display: 'grid',
    minWidth: 0,
    marginRight: 'auto',
  },
  routeTitle: {
    color: '#F4FBFD',
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase400,
    letterSpacing: '-0.02em',
  },
  routeDescription: {
    color: '#8095A4',
    fontSize: tokens.fontSizeBase100,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    '@media (max-width: 620px)': {
      display: 'none',
    },
  },
  loading: {
    color: '#43D9FF',
    '& span': {
      color: '#43D9FF',
    },
    '@media (max-width: 620px)': {
      '& > span:last-child': {
        display: 'none',
      },
    },
  },
  content: {
    position: 'relative',
    minWidth: 0,
    minHeight: 0,
    overflow: 'hidden',
    isolation: 'isolate',
    backgroundColor: '#060A0F',
    backgroundImage: 'radial-gradient(circle at 78% 8%, rgba(18, 130, 162, 0.13), transparent 26%), linear-gradient(rgba(67, 217, 255, 0.028) 1px, transparent 1px), linear-gradient(90deg, rgba(67, 217, 255, 0.028) 1px, transparent 1px)',
    backgroundSize: 'auto, 32px 32px, 32px 32px',
  },
  liveSignal: {
    display: 'flex',
    alignItems: 'center',
    gap: '7px',
    padding: '5px 9px',
    color: '#A9F0FF',
    border: '1px solid #25586A',
    backgroundColor: '#0C1C25',
    borderRadius: 0,
    fontSize: '10px',
    fontWeight: tokens.fontWeightSemibold,
    letterSpacing: '0.08em',
    textTransform: 'uppercase',
    whiteSpace: 'nowrap',
    '@media (max-width: 720px)': {
      display: 'none',
    },
  },
  signalDot: {
    width: '7px',
    height: '7px',
    borderRadius: '50%',
    backgroundColor: '#43D9FF',
    boxShadow: '0 0 0 3px rgba(67, 217, 255, 0.12), 0 0 12px rgba(67, 217, 255, 0.72)',
  },
  signalLimited: {
    backgroundColor: '#F5B942',
    boxShadow: '0 0 0 3px rgba(245, 185, 66, 0.14), 0 0 12px rgba(245, 185, 66, 0.7)',
  },
  commandAction: {
    color: '#43D9FF',
    ':hover': {
      color: '#FFFFFF',
      backgroundColor: 'rgba(67, 217, 255, 0.1)',
    },
  },
  scanline: {
    display: 'none',
  },
  viewFrame: {
    position: 'relative',
    zIndex: 1,
    width: '100%',
    height: '100%',
    animation: 'atlas-enter 420ms cubic-bezier(.16, 1, .3, 1) both',
  },
});

export default function App(): ReactElement {
  const styles = useStyles();
  const [activeTab, setActiveTab] = useState<TabValue>('overview');
  const [resourceTypeFilter, setResourceTypeFilter] = useState('all');
  const [isDark, setIsDark] = useState(() => localStorage.getItem('theme') !== 'light');
  const [navOpen, setNavOpen] = useState(false);

  const inventory = useInventory();
  const admin = useAdminData();
  const licensing = useLicensingData();
  const activeItem = NAV_ITEMS.find((item) => item.value === activeTab) ?? NAV_ITEMS[0];

  function navigate(tab: TabValue) {
    if (tab === 'resources') setResourceTypeFilter('all');
    setActiveTab(tab);
    setNavOpen(false);
  }

  function navigateToResources(typeKey: string) {
    setResourceTypeFilter(typeKey);
    setActiveTab('resources');
  }

  async function refreshAll(): Promise<void> {
    await Promise.all([inventory.refresh(), admin.refresh(), licensing.refresh()]);
  }

  return (
    <FluentProvider theme={isDark ? controlHubDarkTheme : controlHubLightTheme}>
      <div className={styles.shell}>
        {navOpen && <button className={styles.scrim} aria-label="Close navigation" onClick={() => setNavOpen(false)} />}
        <aside className={`${styles.sidebar} ${navOpen ? styles.sidebarOpen : ''}`}>
          <div className={styles.brand}>
            <div className={styles.brandText}>
              <Text className={styles.brandName}>Control Hub</Text>
              <Text className={styles.brandSub}>Tenant operations</Text>
            </div>
            <Button className={styles.closeNav} appearance="subtle" icon={<DismissRegular />} aria-label="Close navigation" onClick={() => setNavOpen(false)} />
          </div>

          <nav className={styles.nav} aria-label="Primary navigation">
            {(['Operate', 'Govern'] as const).map((group) => (
              <div className={styles.navGroup} key={group}>
                <Text className={styles.groupLabel}>{group}</Text>
                {NAV_ITEMS.filter((item) => item.group === group).map((item) => {
                  const active = item.value === activeTab;
                  return (
                    <button
                      key={item.value}
                      type="button"
                      className={`${styles.navButton} ${active ? styles.navButtonActive : ''}`}
                      aria-current={active ? 'page' : undefined}
                      onClick={() => navigate(item.value)}
                    >
                      <span className={styles.navIcon}>{item.icon}</span>
                      <span className={styles.navCopy}>
                        <span className={styles.navLabel}>{item.label}</span>
                        <span className={styles.navDescription}>{item.description}</span>
                      </span>
                      {active && <span className={styles.activeMarker} aria-hidden="true" />}
                    </button>
                  );
                })}
              </div>
            ))}
          </nav>

          <div className={styles.attribution}>
            <img src={ppaLogo} alt="" className={styles.attributionLogo} />
            <span>Built by Power Platform Advocates</span>
          </div>
        </aside>

        <section className={styles.workspace}>
          <header className={styles.commandBar}>
            <Button className={styles.menuButton} appearance="subtle" icon={<NavigationRegular />} aria-label="Open navigation" onClick={() => setNavOpen(true)} />
            <div className={styles.routeMeta}>
              <Text className={styles.routeTitle}>{activeItem.label}</Text>
              <Text className={styles.routeDescription}>{activeItem.description}</Text>
            </div>
            {(inventory.isLoading || admin.isLoading || licensing.isLoading) && (
              <Spinner className={styles.loading} size="tiny" label={inventory.loadingLabel ?? 'Refreshing tenant data'} labelPosition="after" />
            )}
            <div className={styles.liveSignal} aria-live="polite">
              <span className={`${styles.signalDot} ${inventory.error || admin.error ? styles.signalLimited : ''}`} aria-hidden="true" />
              {inventory.error || admin.error || licensing.error
                ? 'Signal limited'
                : inventory.isLoading || admin.isLoading || licensing.isLoading
                  ? 'Synchronizing'
                  : `${inventory.resources.length.toLocaleString()} indexed`}
            </div>
            <Button
              className={styles.commandAction}
              appearance="subtle"
              icon={isDark ? <WeatherSunnyRegular /> : <WeatherMoonRegular />}
              onClick={() => setIsDark((current) => {
                const next = !current;
                localStorage.setItem('theme', next ? 'dark' : 'light');
                return next;
              })}
              aria-label={isDark ? 'Switch to light mode' : 'Switch to dark mode'}
            />
            <Button className={styles.commandAction} appearance="subtle" icon={<ArrowClockwiseRegular />} onClick={() => void refreshAll()} aria-label="Refresh all tenant data" />
          </header>

          <main className={styles.content}>
            <span className={styles.scanline} aria-hidden="true" />
            <div key={activeTab} className={styles.viewFrame}>
              {activeTab === 'overview' && (
              <Dashboard
                resources={inventory.resources}
                counts={inventory.counts}
                environmentsCount={inventory.environments.length}
                recommendations={admin.recommendations}
                isLoading={inventory.isLoading}
                error={inventory.error}
                onNavigateToResources={navigateToResources}
                onNavigateToRecommendations={() => navigate('recommendations')}
              />
              )}
              {activeTab === 'resources' && (
              <ResourcesView resources={inventory.resources} isLoading={inventory.isLoading} error={inventory.error} onRefresh={inventory.refresh} initialTypeFilter={resourceTypeFilter} />
              )}
              {activeTab === 'environments' && (
              <EnvironmentsView environments={inventory.environments} resources={inventory.resources} envGroups={admin.envGroups} isLoading={inventory.isLoading} error={inventory.error} onRefreshEnvironments={inventory.refresh} />
              )}
              {activeTab === 'recommendations' && <RecommendationsView recommendations={admin.recommendations} isLoading={admin.isLoading} error={admin.error} />}
              {activeTab === 'governance' && (
              <GovernanceView billingPolicies={admin.billingPolicies} crossTenantReports={admin.crossTenantReports} dlpPolicies={admin.dlpPolicies} environments={inventory.environments} resources={inventory.resources} isLoading={admin.isLoading} error={admin.error} onRefreshAdmin={admin.refresh} />
              )}
              {activeTab === 'connectors' && (
                <ConnectorsView
                  environments={inventory.environments}
                  resources={inventory.resources}
                />
              )}
              {activeTab === 'licensing' && (
                <LicensingView
                  snapshot={licensing.snapshot}
                  isLoading={licensing.isLoading}
                  error={licensing.error}
                  onRefresh={licensing.refresh}
                  onPeriodChange={licensing.loadPeriod}
                  environments={inventory.environments}
                  resources={inventory.resources}
                />
              )}
            </div>
          </main>
        </section>
      </div>
      <Toaster toasterId="coe-toaster" />
    </FluentProvider>
  );
}
