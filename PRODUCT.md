# Product

<!-- impeccable:product-schema 1 -->

## Platform

web

## Users

Power Platform tenant administrators who need to assess tenant-wide risk, investigate resources and environments, and take governance action during routine operational work.

## Product Purpose

Power Platform Control Hub turns tenant-wide inventory into prioritized governance action in one console. It gives administrators a real-time view of Power Platform resources and environments, exposes recommendations and best-practice findings, and supports direct administrative actions without requiring the CoE Starter Kit.

Success means an administrator can move quickly from a signal or inventory question to the relevant resource, understand its governance context, and complete the appropriate action without switching tools.

## Positioning

The product combines Power Platform inventory, best-practice analysis, governance configuration, and direct admin actions in a single Power Apps Code App authenticated by the Power Apps host.

## Operating Context

Administrators use the app as an operational console for:

- reviewing tenant inventory across apps, flows, agents, environments, connectors, and connections;
- investigating resource and environment details;
- triaging recommendations and best-practice findings;
- managing DLP policies, billing policies, cross-tenant reports, environment groups, rule-based policies, and rule sets;
- performing actions such as enable, disable, quarantine, delete, back up, change ownership, grant elevated access, and manage group membership.

## Capabilities and Constraints

- Preserve every existing workflow and administrative outcome.
- Preserve Microsoft Power Platform terminology.
- Preserve light and dark themes, including persisted user preference.
- Preserve Power Apps host authentication and existing connector/service contracts.
- The app is built with React, TypeScript, Fluent UI v9, and the Power Apps Code App toolchain.
- Generated connector clients and backend mutation semantics are not visual redesign targets.
- Data may be incomplete or unavailable for some resource types because of upstream API and permission limitations; the interface must communicate those limitations honestly.

## Brand Commitments

- Product name: Power Platform Control Hub.
- Preserve Power Platform Advocates attribution and the existing logo asset.
- Voice should be direct, operational, and trustworthy rather than promotional.

## Evidence on Hand

- Existing implemented workflows and product copy in `src/`.
- Existing product documentation in `README.md`.
- Existing Power Platform Advocates logo at `src/assets/ppa-logo.png`.
- Existing screenshots in `screenshots/`.
- No testimonials, customer claims, benchmarks, trend history, compliance score, or commercial claims are available and none should be invented.

## Product Principles

1. Lead from risk or operational question to a clear next action.
2. Keep dense tenant data scannable and familiar to administrators.
3. Preserve context as users move from overview to entity detail and action.
4. Make status, severity, scope, and destructive consequences unambiguous.
5. Prefer truthful incomplete states over inferred or fabricated certainty.

## Accessibility & Inclusion

The product must preserve WCAG-compliant contrast, keyboard navigation, semantic structure, visible focus, labeled controls, and status communication that does not rely on color alone across both themes and responsive layouts.
