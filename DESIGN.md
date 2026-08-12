# Design system

<!-- impeccable:design-schema 1 -->

## Direction

**Tenant Control Room** turns the Control Hub into a full-dark operational environment: live telemetry, clear alert state, dense monitoring panels, and fast transitions from tenant signal to governed action.

This is an **Operate** interface. The control-room atmosphere comes from real state, density, precision, and responsive feedback—not ornamental dashboards or fictional metrics.

## Color

- **Instrument black:** `#060A0F` is the primary working field.
- **Panel graphite:** `#0C141D` and `#111827` separate monitor surfaces.
- **Telemetry cyan:** `#43D9FF` marks live state, primary interaction, focus, and selected data.
- **Alert amber:** `#FFB547` identifies guidance and attention without replacing semantic warning or danger.
- **Structural steel:** `#20313E` and `#29404F` define panel boundaries and grids.
- Semantic status colors remain paired with text or icons.

## Typography

The interface uses Aptos Display, Aptos, and Segoe UI fallbacks.

- The Operations command headline uses 42–54px semibold display type.
- Route titles use 34px semibold.
- Panel and section titles use 16–20px semibold.
- Controls and data rows use 14px.
- Telemetry labels and IDs use 9–12px uppercase or monospace only when representing measured data.

## Layout

- A 232px near-black navigation rail groups **Operate** and **Govern** destinations with a typographic identity only; no avatar or monogram is used.
- A 64px instrument bar carries route state, synchronization, theme, and refresh.
- Operations is a wall of live monitors: tenant state, resource inventory, and recommendations. Recently created resources are intentionally omitted.
- Route headers are compact monitor panels with a narrow amber signal field.
- Environments use full-width operational directory feeds rather than cards.
- Dense resource and governance tables use graphite panels, cyan actions, tabular numerals, and restrained grid lines.

## Components

- **Command monitor:** semantic heading, operational description, and live data register.
- **Navigation rail:** grouped destinations, cyan selected state, and amber current marker.
- **Telemetry panel:** real counts and API state presented with monitor framing.
- **Operational feed row:** environment identity, status, region, resource count, creation date, and actions on one line.
- **Loading signal:** core inventory becomes interactive first; owner and Power Pages enrichment continue in the background.
- **Dialogs and details:** preserve Fluent semantics while inheriting the dark panel system.

## Interaction and motion

- Route changes use one 420ms exponential ease-out reveal.
- Feed rows shift laterally and gain a cyan signal line on hover.
- Loading, success, limited-signal, disabled, error, empty, selected, and focus states remain explicit.
- Reduced-motion preferences collapse transition and animation duration.

## Accessibility

- Both themes remain available; the control-room dark theme is the authored default.
- Cyan and amber accents meet contrast requirements and never carry status alone.
- Navigation exposes `aria-current`; icon-only actions retain accessible labels.
- Directory rows preserve keyboard activation and semantic labels.
- The navigation becomes a modal drawer below 900px.
