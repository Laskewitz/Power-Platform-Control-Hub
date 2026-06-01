import type { AnalysisResult } from './flowAnalyzer.ts';
import type { CanvasAppAdminInfo, AppRoleAssignment } from './canvasAppAdminService.ts';

/**
 * Runs best-practice analysis on a canvas app using only data available from the
 * Power Apps for Admins connector (no .msapp download required).
 */
export function analyzeCanvasApp(adminInfo?: CanvasAppAdminInfo, roleAssignments?: AppRoleAssignment[]): AnalysisResult[] {
  const results: AnalysisResult[] = [];
  if (!adminInfo) return results;

  // -- 1. Missing app description
  if (!adminInfo.description?.trim()) {
    results.push({
      id: 'canvas-no-description',
      title: 'Missing app description',
      description: 'The app has no description, making it harder for admins and users to understand its purpose.',
      recommendation: 'Add a description via Settings > General > Description.',
      severity: 'info',
    });
  }

  // -- 2. Too many connection references
  if (adminInfo.connectionReferences.length > 10) {
    results.push({
      id: 'canvas-too-many-connections',
      title: 'Excessive connector usage',
      description: `The app uses ${adminInfo.connectionReferences.length} connections. Apps with many connectors authenticate all of them at startup, increasing load time.`,
      recommendation: 'Consolidate data access via Dataverse where possible. Remove unused connections from the Data panel.',
      severity: 'warning',
      affectedItems: adminInfo.connectionReferences.map(c => c.displayName ?? c.id ?? '').filter(Boolean).slice(0, 8),
    });
  }

  // -- 3. Premium connectors in use
  const premiumConns = adminInfo.connectionReferences.filter(
    c => c.apiTier?.toLowerCase() === 'premium',
  );
  if (premiumConns.length > 0) {
    results.push({
      id: 'canvas-premium-connectors',
      title: `${premiumConns.length} premium connector(s) in use`,
      description: 'Premium connectors require all users of the app to have a Power Apps Premium licence. Using them without confirming licensing can block users.',
      recommendation: 'Verify that all intended users have the appropriate Power Apps licence. Document the licence requirement in the app description.',
      severity: 'info',
      affectedItems: premiumConns.map(c => c.displayName ?? c.id ?? '').filter(Boolean).slice(0, 6),
    });
  }

  // -- 4. On-premises gateway connections
  const onPremConns = adminInfo.connectionReferences.filter(c => c.isOnPremiseConnection);
  if (onPremConns.length > 0) {
    results.push({
      id: 'canvas-on-premises-connection',
      title: `${onPremConns.length} on-premises gateway connection(s)`,
      description: 'On-premises data gateway connections add a dependency on gateway infrastructure. If the gateway goes offline, the app stops working.',
      recommendation: 'Ensure gateway infrastructure is highly available. Consider migrating to cloud equivalents where possible.',
      severity: 'info',
      affectedItems: onPremConns.map(c => c.displayName ?? c.id ?? '').filter(n => n.length > 0).slice(0, 5),
    });
  }

  // -- 5. Oversharing
  if (adminInfo.sharedUsersCount > 500) {
    results.push({
      id: 'canvas-overshared-users',
      title: `App shared with ${adminInfo.sharedUsersCount.toLocaleString()} users`,
      description: 'Sharing an app with large numbers of individual users is difficult to manage and review.',
      recommendation: 'Share via Entra ID security groups instead of individual user assignments.',
      severity: 'warning',
    });
  } else if (adminInfo.sharedUsersCount > 100) {
    results.push({
      id: 'canvas-overshared-users',
      title: `App shared with ${adminInfo.sharedUsersCount} individual users`,
      description: 'More than 100 individual user assignments can be difficult to manage. Consider group-based access.',
      recommendation: 'Consolidate access into Entra ID security groups and share the app with those groups instead.',
      severity: 'info',
    });
  }

  // -- 6. Bypass consent enabled
  if (adminInfo.bypassConsent) {
    results.push({
      id: 'canvas-bypass-consent',
      title: 'Consent bypass is enabled',
      description: 'Bypass consent allows users to run the app without an explicit consent prompt, which can expose data without user acknowledgment.',
      recommendation: 'Only enable bypass consent for internal enterprise apps where users are informed about data access through other means.',
      severity: 'warning',
    });
  }

  // -- 7. Stale app
  if (adminInfo.lastModifiedTime) {
    const daysSince = (Date.now() - new Date(adminInfo.lastModifiedTime).getTime()) / (1000 * 60 * 60 * 24);
    if (daysSince > 365) {
      results.push({
        id: 'canvas-stale-app',
        title: 'App not modified in over a year',
        description: `The app was last modified ${Math.floor(daysSince)} days ago. Stale apps may contain outdated logic, broken connectors, or unpatched security issues.`,
        recommendation: 'Verify the app is still actively used. Archive or delete it if it is no longer needed, or update it to current platform standards.',
        severity: 'warning',
      });
    } else if (daysSince > 180) {
      results.push({
        id: 'canvas-stale-app',
        title: 'App not modified in over 6 months',
        description: `The app was last modified ${Math.floor(daysSince)} days ago. Apps inactive for extended periods may have stale connections or outdated logic.`,
        recommendation: 'Review the app to confirm it is still required and working correctly, or document its status.',
        severity: 'info',
      });
    }
  }

  // -- 8. Shared with entire organisation
  if (roleAssignments) {
    const tenantShare = roleAssignments.find(r => r.principalType === 'Tenant');
    if (tenantShare) {
      results.push({
        id: 'canvas-shared-with-tenant',
        title: 'App is shared with the entire organisation',
        description: 'This app has been shared with "Everyone in the organisation". All users — including guests, depending on tenant settings — can access it.',
        recommendation: 'Only share organisation-wide if the app is genuinely intended for all users. Use Entra ID security groups for targeted access control.',
        severity: 'warning',
      });
    }
  }

  // -- 9. No co-owners (single point of failure)
  if (roleAssignments) {
    const coOwners = roleAssignments.filter(
      r => r.roleName === 'CanEdit' && r.principalType !== 'Tenant',
    );
    if (coOwners.length === 0) {
      results.push({
        id: 'canvas-no-co-owners',
        title: 'App has no co-owners (co-developers)',
        description: 'Only the original creator has edit access to this app. If that person leaves or loses access, the app cannot be maintained.',
        recommendation: 'Share the app as a co-developer with at least one additional user, group, or service principal to prevent a single point of failure.',
        severity: 'info',
      });
    }
  }

  return results;
}
