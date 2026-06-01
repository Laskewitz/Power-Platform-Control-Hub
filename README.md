# 🚀 Power Platform Control Hub

A **Center of Excellence (CoE) Starter Kit dashboard replacement** built as a [Power Apps Code App](https://learn.microsoft.com/en-us/power-apps/developer/code-apps/overview). It uses the [Power Platform Inventory API](https://learn.microsoft.com/en-us/power-platform/admin/inventory-api) and several Power Platform admin connectors — including **Microsoft Dataverse** — to surface a real-time view of all resources across your tenant: canvas apps, model-driven apps, cloud flows, agent flows, code apps, Copilot Studio agents, and environments. No CoE Starter Kit solution required.

🔐 Authentication is handled entirely by the Power Apps host. No app registration or MSAL configuration is required.

---

## ✨ Features

### 🗂️ Navigation tabs

| Tab | What it shows |
|---|---|
| 🏠 **Overview** | Metric cards per resource type + recently created resources table |
| 📋 **Resources** | Sortable, searchable, filterable table of all resources across all environments. Click any canvas app, cloud flow, or agent to open a full detail panel. |
| 🌍 **Environments** | Card grid of every environment with type badge, managed-environment indicator, region, and resource count. Click any environment to open the environment detail view. |
| 🛡️ **Tenant Policies** | DLP policies (list, create, detail), billing policies, and cross-tenant connection reports |
| 🗂️ **Environment Groups** | Environment groups, rule-based policies, and rule sets (CRUD) |
| 🔌 **Connectors** | Per-environment connections, connectors, and Power Pages websites |
| 💡 **Recommendations** | Advisor recommendations from the admin API |

---

## 🔍 Detail panels

Each supported resource type has a full-page detail panel with structured sections, action bar, and **Best Practice Analysis (BPA)** checks.

---

### 🎨 Canvas App

**Sections:** App Details · Inventory · Governance & Sharing · Role Assignments · Best Practice Analysis

**Actions:** Quarantine / Unquarantine · Add owner (AAD user search) · Elevated admin access / Remove access

**Best Practice Analysis — 9 checks:**

| # | Severity | Check |
|---|---|---|
| 1 | ℹ️ Info | Missing app description |
| 2 | ⚠️ Warning | Excessive connector usage (>10 connections) |
| 3 | ℹ️ Info | Premium connector(s) in use |
| 4 | ℹ️ Info | On-premises gateway connection(s) |
| 5 | ⚠️ Warning | App shared with >500 individual users |
| 6 | ℹ️ Info | App shared with >100 individual users |
| 7 | ⚠️ Warning | Consent bypass is enabled |
| 8 | ⚠️ Warning | App not modified in over a year (stale) |
| 9 | ℹ️ Info | App not modified in over 6 months |
| 10 | ⚠️ Warning | App is shared with the entire organisation |
| 11 | ℹ️ Info | App has no co-owners / co-developers |

---

### ⚡ Cloud Flow / Agent Flow / M365 Agent Flow

**Sections:** Flow Details · Inventory · Triggers & Actions (recursive tree) · Owners · Run-Only Users · Best Practice Analysis

**Actions:** Enable · Disable · Delete · Add owner · Elevated admin access / Remove access

The **Triggers & Actions** section renders the complete flow graph:
- Conditions with **True / False** branches side-by-side
- Loops (`Apply to each`, `Do until`), scopes, and switch/case blocks as collapsible containers
- Connector actions with resolved connector name and humanised operation label
- Trigger: connector + event, schedule details, or kind label (manual/instant)

**Best Practice Analysis — 26 checks:**

| # | Severity | Check |
|---|---|---|
| 1 | 🔴 Critical | No error handling detected |
| 2 | ℹ️ Info | No trigger conditions set |
| 3 | ⚠️ Warning | Very high-frequency recurrence trigger |
| 4 | ℹ️ Info | Trigger concurrency limit not set |
| 5 | ℹ️ Info | Many actions have default names |
| 6 | ⚠️ Warning | HTTP actions without timeout |
| 7 | 🔴 Critical | Possible sensitive data in unprotected inputs |
| 8 | ℹ️ Info | "Apply to each" loops run sequentially |
| 9 | ℹ️ Info | List actions may not retrieve all items (pagination) |
| 10 | ℹ️ Info | Flow has no description |
| 11 | ⚠️ Warning | Nested "Apply to each" loops detected |
| 12 | ⚠️ Warning | Notification or messaging actions inside a loop |
| 13 | ⚠️ Warning | HTTP actions have retries disabled |
| 14 | ℹ️ Info | Flow is very large (>50 actions) |
| 15 | ⚠️ Warning | Switch actions missing a default case |
| 16 | ℹ️ Info | Recurrence trigger has no explicit start time / time zone |
| 17 | ⚠️ Warning | "Do Until" loops without a meaningful iteration limit |
| 18 | ℹ️ Info | HTTP actions use hardcoded URLs |
| 19 | ℹ️ Info | Deeply nested actions detected |
| 20 | ℹ️ Info | Error-path Terminate actions do not use "Failed" status |
| 21 | ℹ️ Info | Trigger uses a default name |
| 22 | ℹ️ Info | Most steps have no description (comment) |
| 23 | 🔴 Critical | HTTP trigger has no Response action |
| 24 | ⚠️ Warning | Parse JSON action(s) without a schema |
| 25 | ⚠️ Warning | Variables used in a concurrent flow |
| 26 | ℹ️ Info | Empty scope(s) found |

---

### 🤖 Copilot Studio Agent

**Sections:** Agent Details · Inventory · Definition (Configuration) · Best Practice Analysis

**Actions:** Quarantine / Unquarantine · Delete · Elevated admin access / Remove access

**Agent Details** includes: display name, schema name, status, language, authentication mode, access control, quarantine state, last published, created/modified, owner, environment, Dataverse URL, and agent ID (Entra + Dataverse).

**Inventory section** shows: orchestration type, model, authentication, channels, capabilities, sharing (viewers/editors), published date, and creator tool.

**Best Practice Analysis — 13 checks:**

| # | Severity | Check |
|---|---|---|
| 1 | ⚠️ Warning | Agent is inactive |
| 2 | ⚠️ Warning | Agent has never been published |
| 3 | ⚠️ Warning | Authentication mode is None or Unspecified |
| 4 | ℹ️ Info | No configuration data found |
| 5 | ℹ️ Info | No primary language configured |
| 6 | ⚠️ Warning | Access control allows anyone |
| 7 | 🔴 Critical | Group membership access control has no groups configured |
| 8 | ⚠️ Warning | Agent allows multi-tenant access |
| 9 | ℹ️ Info | Agent not re-published in N months (stale) |
| 10 | ℹ️ Info | Inactive topic(s) found |
| 11 | ⚠️ Warning | High percentage of topics are disabled (>50%) |
| 12 | ℹ️ Info | No knowledge sources configured |
| 13 | ℹ️ Info | No test cases defined |

---

### 🌍 Environment detail view

**Sections:** Resources tab (table of all resources in the environment) · Analysis tab

**Actions (Actions menu):** Enable / Disable · Enable / Disable Managed Environment · Create Backup · Apply admin access · Add to / Remove from Group

The **Resources table** includes Name, Type, Created, Modified, Owner columns. Clicking the ↗ icon on a Canvas App, Cloud Flow, or Agent row opens that resource's full detail panel.

**Best Practice Analysis — 6 checks:**

| # | Severity | Check |
|---|---|---|
| 1 | ⚠️ Warning | This is the Default environment |
| 2 | ⚠️ Warning | Trial environment will expire |
| 3 | ℹ️ Info | Not a Managed Environment |
| 4 | ℹ️ Info | Not assigned to an Environment Group |
| 5 | ℹ️ Info | Large environment (>200 resources) |
| 6 | ℹ️ Info | Environment URL appears auto-generated |

---

### 🛡️ DLP Policy management

- Full-page list of all tenant DLP policies (V2 API)
- **➕ Create page**: two-stage flow — basic settings (name, scope, default classification) then connector classification (Confidential / General / Blocked buckets)
- **📄 Detail page**: collapsible accordion sections for Policy Details, Connector Groups, Environments, and Advisories
- **✨ Apply Best Practices**: analyses the policy against advisory rules (e.g. HTTP connector → Blocked, SharePoint → Confidential) and proposes changes before saving

### 🗂️ Environment Groups

- Create, edit, and delete environment groups
- Manage environment membership per group
- Rule-based policies: create, assign to groups, edit, extract rule sets
- Rule sets: full CRUD with JSON-based parameter editing

### 🎨 Additional UX

- 🌙 Light / dark mode toggle (preference saved to `localStorage`)
- 📱 Responsive layout (mobile hamburger menu)
- Fluent UI v9 — consistent with Microsoft 365 design language
- ♿ Accessible (WCAG-compliant contrast, ARIA labels, keyboard navigation)
- 🔔 Toast notifications for all write actions
- Inline error messages with expandable details

---

## 🏗️ Architecture

```
CoE-Code/
├── src/
│   ├── App.tsx                        # Root: tab navigation, theme, layout
│   ├── types/
│   │   ├── inventory.ts               # Resource, Environment, ResourceCounts
│   │   └── admin.ts                   # Governance, DLP, billing, connector types
│   ├── hooks/
│   │   ├── useInventory.ts            # Fetches resources & environments
│   │   └── useAdminData.ts            # Fetches admin data (DLP, groups, policies…)
│   ├── services/
│   │   ├── inventoryApi.ts            # Inventory API calls
│   │   ├── adminApi.ts                # Admin V2 API calls (connectors, groups…)
│   │   ├── dlpService.ts              # DLP policy CRUD (Power Platform for Admins)
│   │   ├── canvasAppAnalyzer.ts       # 11 best-practice checks for canvas apps
│   │   ├── canvasAppAdminService.ts   # Canvas app governance via Power Apps for Admins
│   │   ├── flowManagementService.ts   # Flow enable/disable/delete
│   │   ├── flowAnalyzer.ts            # 26 best-practice checks for cloud/agent flows
│   │   ├── copilotStudioService.ts    # Copilot Studio bot Dataverse queries + Admin V2 actions
│   │   ├── governanceMutations.ts     # Env group / policy / rule set write ops
│   │   ├── environmentMutations.ts    # Environment write ops
│   │   ├── ownerCache.ts              # AAD user display name resolution
│   │   ├── settingsService.ts         # Environment management settings (read/write)
│   │   └── tombstoneService.ts        # Soft-delete tracking for resources
│   ├── utils/
│   │   ├── errorUtils.ts              # Error message extraction helpers
│   │   ├── formatDate.ts              # Date formatting helpers
│   │   ├── inventoryFormatters.ts     # Shared inventory data formatting helpers
│   │   └── lcidUtils.ts               # LCID → language name resolution
│   └── components/
│       ├── Dashboard.tsx              # Overview / metric cards
│       ├── ResourcesView.tsx          # Resources table + detail panel routing
│       ├── CanvasAppDetailPanel.tsx   # Canvas app detail + analysis
│       ├── CloudFlowDetailPanel.tsx   # Cloud/agent flow detail + analysis
│       ├── CopilotStudioAgentDetailPanel.tsx  # Copilot Studio agent detail + analysis
│       ├── EnvironmentsView.tsx       # Environment cards
│       ├── EnvironmentDetailView.tsx  # Single environment detail + settings + analysis
│       ├── RecommendationsView.tsx    # Advisor recommendations
│       ├── GovernanceView.tsx         # Tenant Policies tab (DLP, billing, reports)
│       ├── DlpPoliciesView.tsx        # Full-page DLP list / create / detail
│       ├── EnvironmentGroupsView.tsx  # Environment Groups tab
│       ├── ConnectorsView.tsx         # Connectors / connections tab
│       ├── AddSelfAsAdminBanner.tsx   # Elevated access button (inline + menu variants)
│       └── ConfirmDialog.tsx          # Reusable confirmation dialog
├── generated/                         # Auto-generated connector clients (gitignored)
│   ├── models/
│   └── services/
├── index.html
├── vite.config.ts
├── package.json
├── tsconfig.app.json
├── power.config.json                  # Code App + connector connection references
└── deploy/
    ├── Deploy.ps1                     # Multi-environment deploy script
    └── env-config.json                # Connection IDs per environment (gitignored)
```

**Data flow:**
1. The Power Apps host handles authentication — the app renders immediately with no login screen.
2. `useInventory` and `useAdminData` call the generated connector service clients on mount.
3. Connector calls are proxied through the Power Apps host to the respective admin APIs.
4. Results are stored in React state and rendered by the view components.

---

## 🛠️ Prerequisites

| Tool | Minimum version |
|---|---|
| [Node.js](https://nodejs.org/) (LTS) | 18 |
| [Git](https://git-scm.com/) | any |
| Power Platform environment with **code apps enabled** | — |
| Power Platform **tenant admin** account | — |

---

## 🔌 Connector Setup

This app requires **five connectors**. For each one, create a connection in [make.powerapps.com](https://make.powerapps.com) (**Connections → New connection**), note the Connection ID from the URL, then run `add-data-source`.

| Connector | API ID | Docs |
|---|---|---|
| Power Platform Admin V2 | `shared_powerplatformadminv2` | [Reference](https://learn.microsoft.com/en-us/connectors/powerplatformadminv2/) |
| Power Platform for Admins | `shared_powerplatformforadmins` | [Reference](https://learn.microsoft.com/en-us/connectors/powerplatformforadmins/) |
| Power Apps for Admins | `shared_powerappsforadmins` | [Reference](https://learn.microsoft.com/en-us/connectors/powerappsforadmins/) |
| Flow Management | `shared_flowmanagement` | [Reference](https://learn.microsoft.com/en-us/connectors/flowmanagement/) |
| Microsoft Dataverse | `shared_commondataserviceforapps` | [Reference](https://learn.microsoft.com/en-us/connectors/commondataserviceforapps/) |

### Data sources by resource

Each resource type pulls from one or more connectors (and optionally a Dataverse table):

| Resource | List / inventory | Detail & mutations | Notes |
|---|---|---|---|
| **Environments** | Admin V2 → `QueryResources` | Admin V2 → `GetEnvironmentByIdForUser` | Domain/URL data lazy-fetched on Analysis tab open |
| **Canvas Apps** | Admin V2 → `QueryResources` | Power Apps for Admins → `GetApp`, `GetAppRoleAssignment` | Connection refs, owner, shared-with counts |
| **Model-Driven Apps** | Admin V2 → `QueryResources` | _(inventory metadata only)_ | No separate detail API |
| **Cloud Flows** | Admin V2 → `QueryResources` | Flow Management → `GetAdminFlowWithConnectionReferences`, `GetFlowPermissions` | Full trigger/action graph |
| **Agent Flows** | Admin V2 → `QueryResources` | Flow Management (same as Cloud Flows) | |
| **M365 Agent Flows** | Admin V2 → `QueryResources` | Flow Management (same as Cloud Flows) | |
| **Code Apps** | Admin V2 → `QueryResources` | _(inventory metadata only)_ | |
| **Copilot Studio Agents** | Admin V2 → `QueryResources` | 1. Dataverse `bot` table (admin env) via Bots connector<br>2. Cross-env Dataverse connector → `bot` + `botcomponent` tables | Falls back to cross-env query when agent is not in admin env |
| **DLP Policies** | Power Platform for Admins → `ListPoliciesV2` | Power Platform for Admins → `CreatePolicyV2 / UpdatePolicyV2 / DeletePolicyV2` | Advisory recommendations via Admin V2 |
| **Connectors & Connections** | Admin V2 → `GetConnectorById`, connections list | Admin V2 | Per-environment view |
| **Environment Groups** | Admin V2 → `ListEnvironmentGroups` | Admin V2 → create / update / delete | |
| **Billing Policies** | Admin V2 | Admin V2 | |
| **Tombstones** (soft-delete) | Dataverse `ppa_resourcetombstone` table | Dataverse `ppa_resourcetombstone` table | Falls back to `localStorage` when Dataverse is unavailable |
| **User display names** | Microsoft Dataverse → `aadusers` virtual table (AAD-backed virtual entity) | — | Used to resolve owner/created-by GUIDs across all resource types |

### Steps after cloning

```bash
# Install dependencies
npm install

# Initialise the Code App (creates power.config.json if missing)
npx power-apps init --display-name "Power Platform Control Hub"

# Add each connector (answer "No" when asked about connection references)
npx power-apps add-data-source -a shared_powerplatformadminv2    -c <connection-id>
npx power-apps add-data-source -a shared_powerplatformforadmins  -c <connection-id>
npx power-apps add-data-source -a shared_powerappsforadmins      -c <connection-id>
npx power-apps add-data-source -a shared_flowmanagement          -c <connection-id>
npx power-apps add-data-source -a shared_commondataserviceforapps -c <connection-id>
```

The `src/generated/` folder is **gitignored** — every collaborator must run `add-data-source` after cloning to regenerate the typed service clients.

---

## 💻 Local Development

```bash
npx power-apps run
```

This starts a Vite dev server and opens Power Apps in local mode. The app connects to your real tenant data with hot reload enabled.

---

## 🚢 Deploy to Power Apps

```bash
npx power-apps push --solution-id <your-solution-id>
```

The CLI prints a Power Apps URL when the push succeeds.

---

## 📦 Dataverse Solution

A companion Dataverse solution is included in the `solution/` folder. It provides the `ppa_resourcetombstone` table used for soft-delete / tombstone tracking of resources.

```powershell
# Import the solution into your environment
pac solution import --path solution\PowerPlatformControlHub_1_0_0_0.zip --environment <environment-id>
```

The solution source is unpacked in `solution/src/` and can be re-packed after changes:

```powershell
pac solution pack --zipfile solution\PowerPlatformControlHub_1_0_0_0.zip --folder solution\src --packagetype Unmanaged
```

---

## ⚙️ Enable Code Apps on Your Environment

1. Go to [Power Platform admin center](https://admin.powerplatform.microsoft.com).
2. **Manage → Environments → \<your environment\>**.
3. **Settings → Product → Features**.
4. Toggle **Enable code apps** → **Save**.

---

## 📦 Tech Stack

| Package | Purpose |
|---|---|
| [React 18](https://react.dev/) + TypeScript | UI framework |
| [Vite 6](https://vite.dev/) | Build tool / dev server |
| [@microsoft/power-apps](https://www.npmjs.com/package/@microsoft/power-apps) | Code Apps CLI (`init`, `run`, `push`, `add-data-source`) |
| [@microsoft/power-apps-vite](https://www.npmjs.com/package/@microsoft/power-apps-vite) | Vite plugin for Power Apps integration |
| [@fluentui/react-components v9](https://react.fluentui.dev/) | Fluent UI component library |
| [@fluentui/react-icons v2](https://github.com/microsoft/fluentui-system-icons) | Fluent UI icons |

---

## ⚠️ Limitations & Known Issues

- The Inventory API is automatically paginated via `skipToken` — all pages are fetched transparently, so tenants with thousands of resources are fully supported.
- The app requires the signed-in user to be a **Power Platform tenant admin** to read cross-environment data.
- The **Microsoft Dataverse connector** is used to query Copilot Studio bot records. The bundled solution (`solution/`) imports the required `ppa_resourcetombstone` table.
- **DLP policy connector groups do not auto-populate** — when creating or updating a policy, all connector assignments must be explicitly provided. The create page handles this by loading connectors from a selected environment. See [known issues](https://learn.microsoft.com/en-us/connectors/powerplatformforadmins/#known-issues-and-limitations).
- Code apps are **not** supported in the Power Apps mobile app or Power Apps for Windows.
- Code apps **do not** support Power Platform Git integration.
- The `src/generated/` folder is gitignored — every collaborator must run `npx power-apps add-data-source` for all five connectors after cloning.

---

## 📚 Related Resources

- [Power Apps Code Apps overview](https://learn.microsoft.com/en-us/power-apps/developer/code-apps/overview)
- [Power Platform Inventory API](https://learn.microsoft.com/en-us/power-platform/admin/inventory-api)
- [Power Platform Admin V2 connector](https://learn.microsoft.com/en-us/connectors/powerplatformadminv2/)
- [Power Platform for Admins connector](https://learn.microsoft.com/en-us/connectors/powerplatformforadmins/)
- [Power Apps for Admins connector](https://learn.microsoft.com/en-us/connectors/powerappsforadmins/)
- [Flow Management connector](https://learn.microsoft.com/en-us/connectors/flowmanagement/)
- [PowerAppsCodeApps GitHub repository](https://github.com/microsoft/PowerAppsCodeApps)
