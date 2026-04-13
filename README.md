# PCF InfoCard

![TypeScript](https://img.shields.io/badge/TypeScript-4.9-3178C6?logo=typescript&logoColor=white)
![React](https://img.shields.io/badge/React-16.8-61DAFB?logo=react&logoColor=black)
![PCF](https://img.shields.io/badge/PCF-Virtual_Control-742774)
![License](https://img.shields.io/badge/License-MIT-green)

A compact, read-only info card PCF control for Dynamics 365 and Power Apps model-driven forms. Designed for Field Service Mobile technicians who spend excessive time scrolling through forms with many read-only fields, InfoCard consolidates the most important information into a single, tappable card.

InfoCard uses a **slot-based architecture** where admins bind form fields to card zones (header, contact, address, details, grid, tags) directly in the form designer -- no code required. It supports **related entity field fetching** via `@fieldName` syntax, enabling data from lookup records to appear on the card without custom scripting.

**Version:** 3.0.0 | **Namespace:** Contoso | **Control type:** Virtual (React)

<!-- Screenshot placeholder -- add a screenshot of the control here -->
<!-- ![InfoCard screenshot](docs/screenshot.png) -->

---

## Key Features

- **3 layout modes** -- Smart Card (collapsible), Contact Card (full), and Compact Form (dense grid)
- **27 configurable slot properties** across 7 groups (config, header, contact, address, details, grid, tags)
- **Related entity data** via `@fieldName` (1-hop) and `@lookup.field` (2-hop chaining)
- **Bound field or static value** per property -- use `$columnName` for bound fields or enter a literal value
- **Fluent UI theme integration** -- automatically adapts to Power Apps modern theming (light/dark mode)
- **Tappable action links** -- phone (`tel:`), email (`mailto:`), web (`https:`), and map links
- **Duration auto-formatting** -- minutes are automatically rendered as `Xh Xm`
- **Lookup navigation** -- tap a lookup value to open the related record
- **Offline support** -- works in Field Service Mobile with offline-enabled entities
- **Auto-detect icons** -- address, phone, email, instructions, dates, and asset fields receive contextual icons
- **Collapsible smart card** -- preserves critical info (contact, address, tags) when collapsed

---

## Slot Architecture

InfoCard exposes 27 property slots organized into 7 groups. Each slot can be bound to a table column or configured with a static value. Empty slots are automatically hidden.

### Config (4 properties)

| Property | Type | Description |
|----------|------|-------------|
| `layout` | Enum | Card layout mode: `smart` (collapsible), `contact` (full), or `compact` (dense grid) |
| `hideEmptyFields` | TwoOptions | Hide fields with no value (default: true) |
| `showCardBorder` | TwoOptions | Show card border and shadow (default: true) |
| `showVersionInfo` | TwoOptions | Show version badge in bottom-right corner (default: false) |

### Header (4 properties)

| Property | Type | Description |
|----------|------|-------------|
| `titleField` | Bound (required) | Card header -- anchor field. Related fields (`@fieldName`) fetch from this lookup's entity. |
| `subtitleField1` | Input | Shown under title, dot-separated. Supports `@fieldName` for related data. |
| `subtitleField2` | Input | Shown under title, dot-separated |
| `subtitleField3` | Input | Shown under title, dot-separated |

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

### How it works

1. On `updateView`, the control scans all slot values for the `@` prefix
2. It resolves the title field's lookup target (entity type + record ID)
3. It issues WebAPI `retrieveRecord` calls to fetch the requested fields
4. Resolved values are merged into the card data and rendered

This enables rich, cross-entity cards without form scripts or additional customizations.

---

## Quick Start

### Prerequisites

- Node.js 18+
- Power Platform CLI (`pac`) authenticated to your environment

### Build

```bash
cd InfoCardControl
npm install
npm run build
```

### Deploy to Dataverse

```bash
pac pcf push --publisher-prefix cli
```

Or build and import the full solution:

```bash
cd InfoCardSolution
dotnet build --configuration Release
pac solution import --path bin/Release/InfoCardSolution.zip --publish-changes
```

### Regenerate types after manifest changes

```bash
npm run refreshTypes
```

---

## PCF Workbench

InfoCard is compatible with the [PCF Workbench](https://github.com/jaduplesms/PCFBuilderFramework) dev harness, which provides a gallery-based development experience with network conditioning, device emulation, performance monitoring, and hot reload.

To run InfoCard in the workbench:

```bash
cd PCFBuilderFramework/harness
PCF_CONTROL_PATH="../PowerApps/PCFGallery/InfoCard/InfoCardControl/InfoCard" npx vite --port 8181
```

The workbench loads the control's `data.json`, `test-scenarios.json`, and `metadata.json` files automatically.

---

## Test Data

InfoCard ships with mock data files that power both the PCF Workbench and the test suite:

| File | Purpose |
|------|---------|
| `data.json` | Mock entity records (contacts, accounts, bookings) for WebAPI simulation |
| `test-scenarios.json` | Pre-configured property value sets for different form configurations (e.g., "Account - Mobile Form", "Work Order - Field Service") |
| `metadata.json` | Entity metadata in Dataverse API format -- enables the workbench to resolve field display names, option set labels, and lookup targets |
| `EntityDefinitions_bookableresourcebooking.json` | Full Dataverse entity definition export for the booking entity |

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
      data.json                          # Mock entity data for harness
      test-scenarios.json                # Pre-configured test scenarios
      metadata.json                      # Entity metadata for harness
      thumbnail.jpg                      # Gallery thumbnail
      generated/
        ManifestTypes.d.ts               # Auto-generated types (do not edit)
      tests/
        index.test.ts                    # PCF lifecycle tests
        InfoCard.test.tsx                # React component tests
  InfoCardSolution/                      # Dataverse solution wrapper
    InfoCardSolution.cdsproj             # Solution project file
    src/
      Other/
        Solution.xml                     # Solution definition
        Customizations.xml               # Solution customizations
        Relationships.xml                # Entity relationships
```

---

## License

MIT License. See [LICENSE](LICENSE) for details.

### Licensing Notes

- All source code in this repository is original work
- No Microsoft source code is bundled or redistributed
- React 16 is declared as a **platform library** in `ControlManifest.Input.xml` and is provided by the Power Apps host runtime at execution time -- it is not included in the control bundle
- Fluent UI 9 is similarly provided as a platform library by the host
