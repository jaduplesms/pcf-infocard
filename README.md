# PCF InfoCard

![TypeScript](https://img.shields.io/badge/TypeScript-4.9-3178C6?logo=typescript&logoColor=white)
![React](https://img.shields.io/badge/React-16.8-61DAFB?logo=react&logoColor=black)
![PCF](https://img.shields.io/badge/PCF-Virtual_Control-742774)
![License](https://img.shields.io/badge/License-MIT-green)

A compact, read-only info card PCF control for Dynamics 365 and Power Apps model-driven forms. Designed for Field Service Mobile technicians who spend excessive time scrolling through forms with many read-only fields, InfoCard consolidates the most important information into a single, tappable card.

InfoCard uses a **slot-based architecture** where admins bind form fields to card zones (header, contact, address, details, grid, tags) directly in the form designer -- no code required. It supports **related entity field fetching** via `@fieldName` syntax, enabling data from lookup records to appear on the card without custom scripting.

**Version:** 4.4.23-alpha.2 (control 4.4.23) | **Namespace:** Sample | **Publisher prefix:** smp | **Control type:** Virtual (React)

> ⚠️ **This is a sample.** The control ships under the `Sample.InfoCard` namespace with publisher prefix `smp` so that it is obvious in your environment that this is reference code, not a vendor product. Before using in production, **re-namespace** the control to your own publisher (see [Re-namespacing for production](#re-namespacing-for-production)). No support, warranty, or SLA is offered.

---

## Layouts at a glance

| Layout | Best for | Presentation |
|--------|----------|--------------|
| **Smart Card** | Mobile summaries and operational records | Action bar, two-column grid, icon/text details, and optional whole-card collapse |
| **Contact Card** | People, accounts, and customer-facing records | vCard/business-card layout with stacked subtitles, top-right image, postal address, and 44×44 icon actions |
| **Form Card** | Labelled read-only form sections | Section labels and wrapping label:value rows with optional whole-card collapse |

---

## Key Features

- **3 layout modes** -- Smart Card, Contact Card, and Form Card. All three support optional whole-card collapse. Form Card retains the `compact` configuration value for existing form bindings.
- **24 configurable slot properties** across 6 groups (header, contact, address, details, grid, tags), plus **13 config properties**
- **Related entity data** via `@fieldName` (1-hop) and `@lookup.field` (2-hop chaining)
- **Current-record escape hatch** via `@.fieldName` and `@.lookup.field` for column types the PCF manifest cannot bind directly
- **Bound field or static value** per property -- select "Bind to table column" in the form designer or enter a literal value (`$columnName` is a local test-scenario convention, not form-designer syntax)
- **Localized** out of the box -- ships with English, German, French, Spanish, Italian, Dutch, and Japanese resource strings; falls back to the user's Dataverse UI language
- **Fluent UI theme integration** -- automatically adapts to Power Apps modern theming (light/dark mode)
- **Tappable action links** -- phone (`tel:`), email (`mailto:`), web (`https:`), and map links (lat/lng or address text)
- **Copy-to-clipboard buttons** on phone, email, and address chips on web (form factor 2); hidden on mobile/tablet where `tel:`/`mailto:` is primary
- **Avatar / image** -- configure `imageField` with a Dataverse image-column expression (`@.entityimage` for the current record or `@customimage` through the title lookup). Raw Dataverse base64 values render directly; URLs remain supported for compatibility
- **Title prefix** -- optional muted-colour literal before the title (e.g. `Case: ACME-001`, `Work Order: WO-12345`)
- **Configurable collapsible sections** -- choose what disappears when the card collapses: `none`, `body` (details + grid, default), `body-tags`, or `all`. Smart Card, Contact Card, and Form Card can all collapse
- **Detail row presentation** -- optionally suppress leading icons (`showDetailIcons=false`) and/or render the field's display name inline-bold or as a small caps heading (`detailLabelStyle`) so detail rows read like prose
- **Duration auto-formatting** -- numeric minutes are automatically rendered as `Xh Xm`
- **Lookup navigation** -- tap a lookup value to open the related record
- **Offline support** -- works in Field Service Mobile with offline-enabled entities; uses `context.webAPI` so cached data is honoured when offline
- **Auto-detect icons** -- address, phone, email, instructions, dates, and asset fields receive contextual icons (when `showDetailIcons` is on)
- **Layout presets** -- 8 entities auto-populate sensible default slots when fields aren't manually bound (Work Order, Bookable Resource Booking, Account, Contact, Case, Customer Asset, Work Order Service Task, Agreement)
- **Title-lookup preset expansion** -- when `titleField` is a lookup to a supported entity, unconfigured slots can auto-fill from that lookup target's preset
- **Maker diagnostics** -- authoring mode shows syntax guidance for invalid or inefficient `@` / `@.` expressions and suppresses runtime fetches
- **Batched related reads** -- fields from the same title/current record are requested in combined `$select`/`$expand` queries with layered fallbacks

---

## Slot Architecture

InfoCard exposes 24 slot properties across 6 groups, plus 13 config properties. Each slot can be bound to a table column or configured with a static value. Empty slots are automatically hidden.

### Config (13 properties)

| Property | Type | Description |
|----------|------|-------------|
| `layout` | Enum | Card layout mode: `smart` (Smart Card), `contact` (Contact Card), or `compact` (Form Card; value retained for compatibility) |
| `hideEmptyFields` | TwoOptions | Hide fields with no value (default: true) |
| `showCardBorder` | TwoOptions | Show card border and shadow (default: true) |
| `showVersionInfo` | TwoOptions | Show version badge in bottom-right corner (default: false) |
| `startExpanded` | TwoOptions | Start expanded vs. collapsed when `collapsibleSections` is enabled (default: true) |
| `showTitle` | TwoOptions | Show title field in card header (default: true) |
| `subtitleSeparator` | Text | Separator string between subtitle parts (default: `·`). Set to ` • `, ` \| `, etc. |
| `titlePrefix` | Text | Optional muted-colour literal rendered before the title (e.g. `Case: `, `Work Order: `) |
| `imageShape` | Enum | Avatar / image shape: `rounded` (default, 8 px radius), `circle`, or `square` |
| `collapsibleSections` | Enum | What collapses when the card is tapped: `none` (collapse off), `body` (details + grid, default), `body-tags`, or `all` (everything below the header) |
| `showDetailIcons` | TwoOptions | Show auto-detected leading icons on Smart/Contact detail rows (default: true). No effect on Form Card |
| `detailLabelStyle` | Enum | Render the field's display name on Smart/Contact detail rows: `none` (value only, default), `inline-bold` (`**Label:** value`), or `above` (small caps heading on its own line). No effect on Form Card |
| `fieldLabelColumns` | Text | Optional JSON mapping of Detail/Grid slots to current-table column logical names for localized labels when the host omits binding metadata. Changes labels only, not values or bindings |

### Header (5 properties)

| Property | Type | Description |
|----------|------|-------------|
| `titleField` | Bound (required) | Card header -- anchor field. Related fields (`@fieldName`) fetch from this lookup's entity. |
| `imageField` | Input | Dataverse image expression: `@.entityimage` for the current record or `@customimage` for the title lookup record. Raw base64 image values are converted to data URIs. URLs remain supported. |
| `subtitleField1` | Input | Shown under title, separator-joined. Supports `@fieldName` for related data. |
| `subtitleField2` | Input | Shown under title, separator-joined |
| `subtitleField3` | Input | Shown under title, separator-joined |

### Contact (4 properties)

| Property | Type | Description |
|----------|------|-------------|
| `phoneField1` | Input | Tap-to-call phone button |
| `phoneField2` | Input | Secondary phone button |
| `emailField` | Input | Tap-to-send email button |
| `webField` | Input | Tap-to-open URL button |

### Address (3 properties)

| Property | Type | Description |
|----------|------|-------------|
| `addressField` | Input | Tappable address row -- opens map when lat/lng are set |
| `latitudeField` | Input | Latitude for map link on address |
| `longitudeField` | Input | Longitude for map link on address |

### Details (3 properties)

| Property | Type | Description |
|----------|------|-------------|
| `detailField1` | Input | Icon + text info row |
| `detailField2` | Input | Icon + text info row |
| `detailField3` | Input | Icon + text info row |

### Grid (6 properties)

| Property | Type | Description |
|----------|------|-------------|
| `gridField1` | Input | Key:value field in 2-column grid |
| `gridField2` | Input | Key:value field in 2-column grid |
| `gridField3` | Input | Key:value field in 2-column grid |
| `gridField4` | Input | Key:value field in 2-column grid |
| `gridField5` | Input | Key:value field in 2-column grid |
| `gridField6` | Input | Key:value field in 2-column grid |

### Detail and Grid field labels

Some model-driven hosts omit the bound column name for input properties, especially numeric fields. If a bound slot shows `Detail 1` or `Grid 1`, set **Field Label Columns** to a static JSON value matching your existing column bindings:

```json
{"detailField1":"description","gridField1":"revenue","gridField2":"numberofemployees"}
```

InfoCard retrieves the columns' localized display names using the supported metadata API. The map never replaces bound values, including unsaved form values. Real binding metadata takes precedence; `@`/`@.` expressions already resolve labels from their source table and ignore these hints. Literal static values without a column mapping retain their slot-label fallback. If metadata is unavailable, explicit mappings fall back to readable logical names.

### Tags (3 properties)

| Property | Type | Description |
|----------|------|-------------|
| `tagField1` | Input | Compact chip/pill badge |
| `tagField2` | Input | Compact chip/pill badge |
| `tagField3` | Input | Compact chip/pill badge |

---

## Related Fields (@syntax)

InfoCard can fetch data from related entities using `@fieldName` syntax in property values. The **title field** serves as the anchor -- its lookup target becomes the source entity for related field resolution.

### 1-hop: Direct lookup field

```
@msdyn_serviceaccount
```

If the title field is bound to a lookup (e.g., `msdyn_serviceaccount` on a work order), prefixing a slot value with `@` tells InfoCard to fetch that field's value from the lookup entity. In this example, the control calls WebAPI to retrieve data from the account record referenced by `msdyn_serviceaccount`.

### 2-hop: Chained lookup

```
@msdyn_serviceaccount.telephone1
```

Two-hop chaining follows a lookup and then reads a specific field from the resolved entity. This fetches `telephone1` from the account record that `msdyn_serviceaccount` points to -- useful for showing a service account's phone number directly on a work order card.

### Current-record (`@.fieldName`) — escape hatch for unbindable column types

```
@.duration
@.statecode
@.ownerid
```

The PCF manifest does not allow direct binding to several common column types (`Whole.Duration`, `Status`, `Status Reason`, `Lookup.Customer`, `Lookup.Owner`, `Lookup.Regarding`, `Lookup.PartyList`, `Whole.Language`, `Whole.TimeZone`, file columns — see [the official property reference](https://learn.microsoft.com/en-us/power-apps/developer/component-framework/manifest-schema-reference/property#value-elements-that-arent-supported)). The form designer's column picker simply will not surface them.

For these columns, configure the slot as a **static input** (`type="SingleLine.Text" static="true"` in the form XML) with the value `@.columnname`. InfoCard fetches the column via WebAPI on mount and renders it just like a bound slot. This works for any column the user has read access to — `duration`, `statecode`, `customerid`, `ownerid`, etc. — and respects offline mode.

Duration columns are auto-detected (by metadata `Format=Duration` and a column-name heuristic) and rendered as `Xh Ym` rather than the locale-grouped raw integer.

### How it works

1. On `updateView`, the control scans all slot values for the `@` prefix
2. It resolves the title field's lookup target (entity type + record ID), or the current form record for `@.`
3. It issues WebAPI `retrieveRecord` calls to fetch the requested fields
4. Resolved values are merged into the card data and rendered

This enables rich, cross-entity cards without form scripts or additional customizations.

---

## Layout Presets

For known Dynamics tables, InfoCard ships with **built-in slot presets**. Drop the control on a form, bind only `titleField`, and unbound slots are auto-filled from a preset map keyed by the form's entity logical name.

| Entity | Auto-filled slots |
|--------|-------------------|
| `msdyn_workorder` | service account, primary incident type, address, summary, status, priority |
| `bookableresourcebooking` | resource, work order, status, start/end/duration |
| `account` | primary contact, industry, phones, email, web, address, status |
| `contact` | job title, parent customer, mobile/phone, email, web, address, status |
| `incident` | customer, case type, created/modified, priority, status |
| `msdyn_customerasset` | account, parent asset, product, registration, install date, location, serial number, status |
| `msdyn_workorderservicetask` | work order, task type, estimated duration, percent complete, completion notes |
| `msdyn_agreement` | service account, billing account, start/end dates, status |

**Maker bindings always win** — any slot configured in the form designer (with a `type` set) is left untouched by the preset. Preset slots that resolve to an empty column on the current record are hidden automatically.

The full map lives in `SLOT_PRESETS` in `InfoCardControl/InfoCard/index.ts`.

---

## Quick Start

### Prerequisites

- Node.js 22+ (CI uses Node.js 22)
- Power Platform CLI (`pac`) authenticated to your environment
- .NET SDK (for building the solution package)

### Build the control

```bash
cd InfoCardControl
npm ci
npm run build -- --buildMode production
npm test
```

### Build the shipping solution package

```bash
# From the repository root (PowerShell)
.\scripts\Build-Release.ps1
# artifacts\InfoCardSolution_4.4.23_unmanaged.zip
# artifacts\InfoCardSolution_4.4.23_managed.zip
# artifacts\SHA256SUMS_4.4.23.txt
```

Both zips contain only the control under `Controls/smp_Sample.InfoCard/`.
The script builds in production mode, verifies versions, managed flags and
control-only contents, and leaves source manifests unchanged. Do not use the
old, unversioned ZIPs from a previous local build. Once a release is published,
download the versioned packages from [GitHub Releases](https://github.com/jaduplesms/pcf-infocard/releases).

### Import into a Dataverse environment

```bash
pac auth create --url https://<your-org>.crm.dynamics.com
pac solution import --path .\artifacts\InfoCardSolution_4.4.23_managed.zip --publish-changes
```

After import, add the control to a form column via **Form designer → Components → Get more components → InfoCard (Sample)**.

### Uninstall

```bash
pac solution delete --solution-name InfoCardSample
```

> Deleting the managed solution removes its control after dependent form references
> are removed. Deleting an unmanaged solution removes only the solution container,
> not the installed control. Back up and remove dependencies before uninstalling.

### Regenerate types after manifest changes

```bash
npm run refreshTypes
```

---

## Re-namespacing for production

The control ships under `Sample.InfoCard` with publisher prefix `smp` so it is unmistakable that this is reference code. Before any production rollout you should re-namespace it under your own publisher. To do so:

1. Edit `InfoCardControl/InfoCard/ControlManifest.Input.xml` and change `namespace="Sample"` to your own (e.g. `namespace="Contoso"`).
2. Edit `InfoCardSolution/src/Other/Solution.xml` and replace the `<Publisher>` block with your real publisher (`UniqueName`, `CustomizationPrefix`, `CustomizationOptionValuePrefix`).
3. Bump the `<Version>` (e.g. `4.0.0.0` → `1.0.0.0` for your first internal release).
4. Run `npm run refreshTypes && npm run build` then rebuild the solution.
5. Update any form Customizations.xml in your own solution to reference the new control name `<your-prefix>_<YourNamespace>.InfoCard`.

Your customer is then importing **your** solution — not a sample — and lifecycle/upgrade rules belong to you.

---

## Supported environments

| Host | Status |
|------|--------|
| Dynamics 365 model-driven apps (online) | Tested |
| Field Service Mobile (online) | Tested |
| Field Service Mobile (offline) | Works for offline-enabled entities; `@`-related fetches require the related entity to be in the offline profile |
| Power Apps form designer (`make.powerapps.com`) | Authoring-mode preview supported |
| Canvas apps | Not supported (model-driven only) |

Tested against PCF SDK 1.x, React 16.8, Fluent UI 9 (host-provided), TypeScript 4.9.

### Compatibility and privacy limitations

Current-record presets and `@.` syntax rely on host record-context metadata
(`context.mode.contextInfo`); maker preview uses the undocumented
`context.mode.isAuthoringMode` flag with a URL fallback. These integrations may
change with the host and are not a supported-platform guarantee. Explicitly
bound slots and title-lookup `@fieldName` mappings do not require current-record
context. Validate the sample in your target environment before production use.

The control has no custom authentication or telemetry service. Call, email,
website and map actions invoke external applications; map links send the address
or coordinates to Google Maps when clicked. Remote image URLs request the image
host when rendered and may disclose network/record context. Prefer Dataverse
image columns for offline use; external images and actions are not guaranteed
offline. Do not include sensitive fields in a mobile offline profile on the
assumption that field-level security will conceal them offline.

---

## UI automation selectors

Rendered fields expose stable, localization-independent `data-automation-id` attributes based on their manifest slot names. Scope selectors to the desired PCF instance when a form contains more than one InfoCard.

| Element | Selector example |
|---------|------------------|
| Card root | `[data-automation-id="infocard"]` |
| Field slot | `[data-automation-id="infocard-gridField3"]` |
| Field action | `[data-automation-id="infocard-action-phoneField1"]` |
| Copy action | `[data-automation-id="infocard-copy-phoneField1"]` |
| Image / version | `[data-automation-id="infocard-imageField"]`, `[data-automation-id="infocard-version"]` |

---

## PCF Workbench

InfoCard is compatible with [PCF Workbench](https://github.com/jaduplesms/PCF-Workbench), which provides device emulation, network conditioning, WebAPI mocking, scenarios, lifecycle diagnostics, and headless acceptance reports.

From `InfoCardControl`, build and launch the published CLI:

```bash
npm run build
npx @pcfworkbench/cli@latest start --path .\InfoCard --port 8181
```

For a headless build/render/report pass:

```bash
npx @pcfworkbench/cli@latest loop --path .\InfoCard --out .\pcf-loop-reports
```

The tracked `test-scenarios.json` contains standalone, fictional mock scenarios
and metadata. Select one in the Scenarios panel. A local `data.json` or
`metadata.json` may override/enrich the harness; both are ignored because they
can contain environment-specific data. Do not commit live exports or include
them in public screenshots or reports.

### Previewing authoring (designer) mode locally

The browser harness has no equivalent of `context.mode.isAuthoringMode`. To exercise the maker preview path (sample data injection, suppressed WebAPI fetches) without deploying to a real environment, trigger any one of:

- Append `#authoring` to the harness URL (e.g. `http://localhost:8181/#authoring`) — easiest for ad-hoc preview
- Append `?authoring=1` to the URL
- In DevTools, run `window.__INFOCARD_AUTHORING__ = true` then re-trigger `updateView` (e.g. by reloading)

The control treats any of these as authoring mode, identical to the real Power Apps form designer. Useful for iterating on sample-data injection and snapshot-testing the designer preview output.

---

## Test Data

InfoCard ships with mock data files that power both the PCF Workbench and the test suite:

| File | Purpose |
|------|---------|
| `tests/test-data.json` | Fictional entity records imported by the Jest suite. |
| `test-scenarios.json` | Standalone mock-only scenarios with fictional records and metadata for PCF Workbench. |
| `data.json`, `metadata.json` | Optional ignored, local-only harness overrides; not required by tests or included in releases. |

Property values in test scenarios use the `$columnName` convention to indicate bound fields (e.g., `$telephone1` binds to the `telephone1` column).

---

## Tests

InfoCard includes a comprehensive test suite covering the PCF lifecycle class (`index.ts`) and the React component (`InfoCard.tsx`).

```bash
cd InfoCardControl
npm test
```

Tests use Jest with jsdom, `@testing-library/react`, and mock `ComponentFramework.Context` helpers. The suite covers:

- Slot value resolution (bound fields, formatted values, lookups)
- Related field detection and WebAPI fetch logic
- Layout mode rendering (smart, contact, compact)
- Theme resolution from Fluent Design Language tokens
- Action link generation (tel:, mailto:, https:, map links)
- Empty field hiding
- Duration formatting
- Collapsible card behavior
- Dataverse image-column base64 normalization
- vCard layout behavior, touch targets, focus states, and automation selectors
- Authoring-mode samples and binding diagnostics
- Stale async response protection

---

## Project Structure

```
InfoCard/
  CLAUDE.md                              # Development guidance
  InfoCardControl/                       # PCF control project
    InfoCardControl.pcfproj              # MSBuild project file
    package.json                         # Dependencies and scripts
    tsconfig.json                        # TypeScript configuration
    jest.config.js                       # Test configuration
    InfoCard/                            # Control source
      ControlManifest.Input.xml          # Property definitions (source of truth)
      index.ts                           # PCF lifecycle class
      InfoCard.tsx                       # React component (3 layout renderers)
      bindingValidator.ts                # @ / @. binding validation helpers
      strings/                           # Localized PCF resources (7 languages)
      test-scenarios.json                # Fictional mock harness scenarios + metadata
      thumbnail.jpg                      # Gallery thumbnail
      generated/
        ManifestTypes.d.ts               # Auto-generated types (do not edit)
      tests/
        index.test.ts                    # PCF lifecycle tests
        InfoCard.test.tsx                # React component tests
        bindingValidator.test.ts         # Binding validation tests
        test-data.json                   # Fictional Jest fixtures
  InfoCardSolution/                      # Dataverse solution wrapper
    InfoCardSolution.cdsproj             # Solution project file
    src/
      Other/
        Solution.xml                     # Solution definition
        Customizations.xml               # Solution customizations
        Relationships.xml                # Entity relationships
  InfoCardSampleForm/                    # Development-only form binding sandbox
```

---

## Roadmap

See [`docs/ROADMAP.md`](docs/ROADMAP.md) for the verified v4.4.21 baseline and the remaining candidate backlog. Issues and PRs welcome.

## Release acceptance

Before publishing a version, run `npm ci`, `npm run lint`, `npm test -- --ci`,
and `npm run build -- --buildMode production` in `InfoCardControl`, then run
`.\scripts\Build-Release.ps1` from the root. Run PCF Workbench's headless loop
against fictional scenarios for each layout, including a mobile/offline case.
CI also verifies managed and unmanaged solution packages.

Local tests and the mock harness do not prove Dataverse import, live metadata,
or mobile offline compatibility. Import the managed package into a disposable
test environment; check all layouts, lookup navigation, keyboard/screen-reader
behavior, slow-network errors, and real Field Service Mobile offline rendering.
Record the environment/client versions and results before claiming support.
Review translations with native speakers if shipping localized production use.

Version tags (`v4.4.21`, for example) trigger a workflow that validates matching
versions and creates a **draft**, not a published release, with both packages
and checksums. Review the changelog, acceptance evidence and assets before
manually publishing. Never publish local exports, debug reports or the
development-only `InfoCardSampleForm` solution.

Alpha tags such as `v4.4.23-alpha.1` use the matching numeric control version
(`4.4.23`) and create prerelease drafts. An alpha may disclose pending live
acceptance, but must pass the automated build, tests, package checks and mock
layout gates. It is not production-ready or a claim that live/offline gates
have passed; see the changelog for the exact validation limits.

---

## License

MIT License. See [LICENSE](LICENSE) for details.

### Licensing Notes

- All source code in this repository is original work
- No Microsoft source code is bundled or redistributed
- React 16 is declared as a **platform library** in `ControlManifest.Input.xml` and is provided by the Power Apps host runtime at execution time -- it is not included in the control bundle
- Fluent UI 9 is similarly provided as a platform library by the host
