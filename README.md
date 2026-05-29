# 🚀 Power Platform Control Hub

A **Center of Excellence (CoE) Starter Kit dashboard replacement** built as a [Power Apps Code App](https://learn.microsoft.com/en-us/power-apps/developer/code-apps/overview). It uses the [Power Platform Inventory API](https://learn.microsoft.com/en-us/power-platform/admin/inventory-api) and several Power Platform admin connectors to surface a real-time view of all resources across your tenant — canvas apps, model-driven apps, cloud flows, agent flows, code apps, Copilot Studio agents, and environments — without requiring Dataverse or the CoE Starter Kit solution.

🔐 Authentication is handled entirely by the Power Apps host. No app registration or MSAL configuration is required.

---

## ✨ Features

### 🗂️ Navigation tabs

| Tab | What it shows |
|---|---|
| 🏠 **Overview** | Metric cards per resource type + recently created resources table |
| 📋 **Resources** | Sortable, searchable, filterable table of all resources across all environments. Click any canvas app, cloud flow, or agent flow to open a detail panel. |
| 🌍 **Environments** | Card grid of every environment with type badge, managed-environment indicator, region, and resource count |
| 💡 **Recommendations** | Advisor recommendations from the admin API |
| 🛡️ **Tenant Policies** | DLP policies (list, create, detail), billing policies, and cross-tenant connection reports |
| 🗂️ **Environment Groups** | Environment groups, rule-based policies, and rule sets (CRUD) |
| 🔌 **Connectors** | Per-environment connections, connectors, and Power Pages websites |

### 🔍 Resource detail panels

**🎨 Canvas App detail panel** (click any canvas app in Resources):
- Best-practice analysis: 24 checks across nested galleries, delegation risks, hardcoded URLs/emails, media size, OnError handling, and more
- Governance data: sharing stats (users + groups), active connections, bypass consent flag, premium/on-premises connectors
- Role assignments list (owner, co-owner, viewer)
- 🔒 Quarantine / Unquarantine actions
- ➕ Add owner (searches AAD users)
**⚡ Cloud Flow / Agent Flow / M365 Agent Flow detail panel**:
- **Triggers & Actions** accordion section with a recursive tree view:
  - Conditions shown with **True** / **False** branches side-by-side
  - Loops (`Apply to each`, `Do until`), scopes, and switch/case blocks rendered as collapsible containers with nested children
  - Connector actions display the resolved connector name + humanised operation
  - Trigger display: connector name, human-readable trigger event, and internal trigger key (for connector triggers); schedule details (recurrence); or kind label (manual/instant)
- **Best Practice Analysis**: 22 checks across trigger naming, undocumented actions, error handling, nested conditions, hardcoded values, missing timeout/retry, and more
- ▶️ Enable / ⏹️ Disable / 🗑️ Delete actions
- Owner resolution
- ➕ Add owner action

### 🛡️ DLP Policy management

- Full-page list of all tenant DLP policies (V2 API)
- **➕ Create page**: two-stage flow — basic settings (name, scope, default classification) then a connector classification stage that loads all connectors from a selected environment, appends the 4 hardcoded connectors required by the API (Spatial Services, HTTP Request, HTTP Webhook, HTTP), and lets you classify connectors into Confidential / General / Blocked buckets
- **📄 Detail page** (Flow-style layout): compact header with policy name + badges, action bar (Edit / Delete), collapsible accordion sections for Policy Details, Connector Groups, Environments, and Advisories
- **✨ Apply Best Practices**: analyses the current policy against a set of advisory rules (e.g. HTTP connector should be Blocked, SharePoint should be Confidential) and proposes changes in a confirmation dialog; after approval the updated policy is shown for review before saving
- Advisories displayed as single-line rows (connector name + classification badges + reason)
- Follows the [known DLP API limitations](https://learn.microsoft.com/en-us/connectors/powerplatformforadmins/#known-issues-and-limitations) — connector groups are always submitted explicitly (no auto-population)

### 🗂️ Environment Groups

- Create, edit, and delete environment groups
- Manage environment membership per group
- Rule-based policies: create, assign to groups, edit, extract rule sets
- Rule sets: full CRUD with JSON-based parameter editing

### 🎨 Additional UX

- 🌙 Light / dark mode toggle
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
│   │   ├── canvasAppAnalyzer.ts       # 24 best-practice checks for canvas apps
│   │   ├── canvasAppAdminService.ts   # Canvas app governance via Power Apps for Admins
│   │   ├── flowManagementService.ts   # Flow enable/disable/delete
│   │   ├── flowAnalyzer.ts            # 22 best-practice checks for cloud/agent flows
│   │   ├── governanceMutations.ts     # Env group / policy / rule set write ops
│   │   ├── environmentMutations.ts    # Environment write ops
│   │   ├── ownerCache.ts              # AAD user display name resolution
│   │   └── settingsService.ts        # App settings persistence
│   ├── components/
│   │   ├── Dashboard.tsx              # Overview / metric cards
│   │   ├── ResourcesView.tsx          # Resources table + detail panel routing
│   │   ├── CanvasAppDetailPanel.tsx   # Canvas app detail + analysis
│   │   ├── CloudFlowDetailPanel.tsx   # Cloud/agent flow detail + analysis
│   │   ├── EnvironmentsView.tsx       # Environment cards
│   │   ├── EnvironmentDetailView.tsx  # Single environment detail
│   │   ├── RecommendationsView.tsx    # Advisor recommendations
│   │   ├── GovernanceView.tsx         # Tenant Policies tab (DLP, billing, reports)
│   │   ├── DlpPoliciesView.tsx        # Full-page DLP list / create / detail
│   │   ├── EnvironmentGroupsView.tsx  # Environment Groups tab
│   │   ├── ConnectorsView.tsx         # Connectors / connections tab
│   │   └── ConfirmDialog.tsx          # Reusable confirmation dialog
│   ├── generated/                     # Auto-generated connector clients (gitignored)
│   │   ├── models/
│   │   └── services/
│   └── utils/
│       └── errorUtils.ts              # Error message extraction helpers
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

This app requires **four connectors**. For each one, create a connection in [make.powerapps.com](https://make.powerapps.com) (**Connections → New connection**), note the Connection ID from the URL, then run `add-data-source`.

| Connector | API ID | Docs |
|---|---|---|
| Power Platform Admin V2 | `shared_powerplatformadminv2` | [Reference](https://learn.microsoft.com/en-us/connectors/powerplatformadminv2/) |
| Power Platform for Admins | `shared_powerplatformforadmins` | [Reference](https://learn.microsoft.com/en-us/connectors/powerplatformforadmins/) |
| Power Apps for Admins | `shared_powerappsforadmins` | [Reference](https://learn.microsoft.com/en-us/connectors/powerappsforadmins/) |
| Flow Management | `shared_flowmanagement` | [Reference](https://learn.microsoft.com/en-us/connectors/flowmanagement/) |

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
- **DLP policy connector groups do not auto-populate** — when creating or updating a policy, all connector assignments must be explicitly provided. The create page handles this by loading connectors from a selected environment. See [known issues](https://learn.microsoft.com/en-us/connectors/powerplatformforadmins/#known-issues-and-limitations).
- Code apps are **not** supported in the Power Apps mobile app or Power Apps for Windows.
- Code apps **do not** support Power Platform Git integration.
- The `src/generated/` folder is gitignored — every collaborator must run `npx power-apps add-data-source` for all four connectors after cloning.

---

## 📚 Related Resources

- [Power Apps Code Apps overview](https://learn.microsoft.com/en-us/power-apps/developer/code-apps/overview)
- [Power Platform Inventory API](https://learn.microsoft.com/en-us/power-platform/admin/inventory-api)
- [Power Platform Admin V2 connector](https://learn.microsoft.com/en-us/connectors/powerplatformadminv2/)
- [Power Platform for Admins connector](https://learn.microsoft.com/en-us/connectors/powerplatformforadmins/)
- [Power Apps for Admins connector](https://learn.microsoft.com/en-us/connectors/powerappsforadmins/)
- [Flow Management connector](https://learn.microsoft.com/en-us/connectors/flowmanagement/)
- [PowerAppsCodeApps GitHub repository](https://github.com/microsoft/PowerAppsCodeApps)
