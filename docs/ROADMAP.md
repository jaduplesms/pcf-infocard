# InfoCard Roadmap

> Sample/community roadmap. No SLAs or commitments. The current baseline is
> verified against control version **4.4.21** on `main`.

**Effort sizing:** XS = under an hour · S = a day · M = a few days · L = 1–2 weeks.

---

## Current baseline — v4.4.21

| Area | Shipped capability |
|------|--------------------|
| Layouts | Smart Card, vCard-style Contact Card, and Form Card (formerly Compact Form; configuration value `compact`); configurable collapse behavior across all three |
| Slots | 24 typed slots across header, contact, address, details, grid, and tags; 12 configuration properties |
| Data | Bound/static values, `@field`, `@lookup.field`, `@.field`, and `@.lookup.field` |
| Presets | Eight entity presets plus automatic expansion from a lookup-bound title |
| Fetching | Combined `$select`/`$expand` reads per source record with lookup/hop-two fallbacks and stale-response guards |
| Maker experience | Deterministic authoring preview, suppressed design-time network calls, and binding syntax warnings |
| Presentation | Fluent token theming, localized labels/durations, image/base64 support, lookup navigation, copy actions, and stable automation selectors |
| Quality | CI, Jest coverage, keyboard/focus behavior, 44×44 contact actions, loading shimmer/progress, and offline-aware WebAPI usage |

See [`CHANGELOG.md`](../CHANGELOG.md) for the detailed release history.

## Candidate product backlog

| # | Item | Effort | Status |
|---|------|--------|--------|
| 17 | **Distance to record (URS-aware)** — use device geolocation plus `msdyn_RetrieveDistanceMatrix` when available, with haversine fallback and configurable units/mode. | M | Candidate |
| 18 | **Static distance origin** — optional depot/resource origin coordinates instead of the device location. | S | Depends on #17 |
| 19 | **Geocode destination from address** — use `msdyn_GeocodeAddress` when latitude/longitude are unavailable. | S | Depends on #17 |
| 20 | **Live distance updates** — throttled `watchPosition` refresh after meaningful movement. | M | Depends on #17 |
| 11 | **Dataset binding option** — render a card per row from a view/subgrid rather than one card per form record. | L | Candidate |
| 25 | **Expanded maker validation** — extend the shipped syntax panel with cross-field checks such as latitude without longitude and incompatible layout/property combinations. | S | Candidate |

## Documentation and samples

| # | Item | Effort | Status |
|---|------|--------|--------|
| 13 | **Animated GIF / short video** showing all three layouts on a real form. | S | Open |
| 14 | **Sample model-driven app** demonstrating multiple entities using InfoCard. | M | Open |
| 16 | **Layout screenshots** — add `layout-smart.png`, `layout-contact.png`, and `layout-compact.png` under `docs/images/`. | XS | Open |

## Completed roadmap items

| # | Item | Completed |
|---|------|-----------|
| 1 | CI workflow: install, lint, build, and Jest on pushes/PRs | v4.1 |
| 2 | Accessibility pass: keyboard behavior, labels, focus, contrast, and touch targets | v4.1–v4.4 |
| 3 | Locale-aware date, time, integer, and duration formatting | v4.1–v4.4 |
| 4 | Seven-language PCF resource strings | v4.1 |
| 5 | Presets for Customer Asset, Work Order Service Task, Agreement, and expanded Case | v4.1 |
| 6 | Header image/avatar slot | v4.1 |
| 7 | Desktop copy-to-clipboard actions | v4.1 |
| 8 | Configurable subtitle separator | v4.1 |
| 10 | Batched related-field reads with combined projections and layered fallbacks | v4.4 |
| 12 | Initial form-designer binding diagnostics | v4.4 |
| 15 | GitHub issue/PR templates | v4.1 |
| 21 | Title prefix | v4.2 |
| 22 | Image shape preference | v4.2 |
| 23 | Configurable collapsible sections | v4.2 |
| 24 | Detail icon and label presentation options | v4.3 |

New roadmap proposals should describe the user problem, expected host support
(model-driven web/mobile/offline), and acceptance criteria. Confirm that the
capability is not already present in the manifest, control source, or changelog
before adding it as future work.
