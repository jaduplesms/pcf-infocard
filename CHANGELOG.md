# Changelog

All notable changes to the InfoCard PCF control are documented in this file.

The project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## Unreleased — v4.3.0 (in progress on `main`)

### Added
- **Show / hide detail icons** — new `showDetailIcons` config property (TwoOptions, default `true`). When `false`, the auto-detected leading icon (✏️ instructions, 📋 summary, 📝 notes, 📅 dates, etc.) on Smart and Contact detail rows is suppressed and only the value text renders. No effect on Compact layout (which uses label:value format).
- **Detail label style** — new `detailLabelStyle` config property (Enum: `none` / `inline-bold` / `above`, default `none`). Lets the maker render the field's display name on Smart and Contact detail rows: `inline-bold` prefixes the value with the bolded display name (e.g. **Instructions:** Access via back garden) so the row reads like prose; `above` renders the display name as a small caps heading on its own line above the value. Combine with `showDetailIcons=false` for plain-text section-style content. No effect on Compact layout.
- 9 new unit tests covering icon suppression and all three label-style modes across Smart, Contact, and Compact layouts (plus the combined "prose-style" case with both options enabled).
- Per-folder `README.md` for `InfoCardSolution/` (shipping wrapper) and `InfoCardSampleForm/` (dev sandbox) clarifying which solution is which.

### Changed
- **Folder rename:** `InfoCardTestSolution/` → `InfoCardSampleForm/` (with matching `.cdsproj`, `Solution.xml` `UniqueName`, and display name updates) so it's obvious from the folder name that this is a development reference, not a shipping artifact.
- Both solutions bumped to version `4.3.0` to track the control manifest.

### Removed
- `InfoCardControl/InfoCard/metadata.json` is no longer tracked. It is harness-only (PCF Workbench display labels), is not imported by tests, and tended to leak environment-specific custom-prefix columns. Generate one locally if you want richer labels in the workbench; an entry in `.gitignore` keeps any local copy out of commits.
- `InfoCardControl/InfoCard/EntityDefinitions_*.json` — unused dev-time exports — removed from the repo and added to `.gitignore`.

## Unreleased — v4.2.0 (in progress on `main`)

### Added
- **Configurable collapsible sections** — new `collapsibleSections` enum config property lets makers choose what disappears when the card collapses: `none` (collapse disabled, no chevron), `body` (details + grid — default, preserves prior Smart-layout behaviour), `body-tags` (details + grid + tags), or `all` (everything below the header — contact rows, body, and tags). Applies to all three layouts; Contact and Compact layouts gain the chevron + whole-card click-to-toggle for the first time when set to anything other than `none`.
- **Title prefix** — new `titlePrefix` config property (`SingleLine.Text`). Renders a muted-colour literal before the title text (e.g. `Case: ACME-001`, `Work Order: WO-12345`). Useful when the form section header doesn't already announce the record type. Plain literal — no template syntax.
- **Image shape** — new `imageShape` enum config property (`rounded` | `circle` | `square`, default `rounded`). Default changed from circle to rounded rectangle (8 px radius), which preserves entity logos (account, asset) and people photos equally well. Initials fallback follows the same shape.
- 12 new unit tests covering the three v4.2 features (titlePrefix, imageShape variants, collapsibleSections across all three layouts and all four modes).

### Changed
- **Avatar default shape**: rounded rectangle (8 px) instead of circle. Set `imageShape="circle"` to restore prior look.
- **No initials placeholder when `imageField` is unbound or empty.** Previously, leaving `imageField` blank rendered a coloured square with the record's initials. Now: no avatar renders at all unless an image URL is supplied. Initials still serve as the **fallback** when a supplied URL fails to load (e.g. 404). This matches customer expectation that "no image configured = no avatar."

## Unreleased — v4.1.0 (in progress on `main`)

### Added
- **Subtitle separator** is now configurable via the `subtitleSeparator` config property (default `·`). Makers can set any literal string (` • `, ` | `, etc.) in the form designer.
- **Copy-to-clipboard buttons** appear next to phone, email, and address chips on web (`getFormFactor() === 2`). Hidden on mobile/tablet (where `tel:`/`mailto:`/map links remain primary). Uses `navigator.clipboard.writeText` with a feature-detect; degrades silently in environments without the API. Includes a polite `aria-live` region announcing "Copied" for screen readers.
- **Four new layout presets** in `SLOT_PRESETS` so unbound slots auto-populate with sensible defaults on these entities:
  - `msdyn_customerasset` (Customer Asset)
  - `msdyn_workorderservicetask` (Work Order Service Task)
  - `msdyn_agreement` (Agreement)
  - `incident` preset expanded with phone/email/detail/grid mappings
- **GitHub community templates** under `.github/`: structured Bug Report and Feature Request issue forms, a config link to the README and roadmap, and a checklist-driven Pull Request template.
- **Localization** via standard Dataverse PCF `resx` resources (`Sample/strings/SampleInfoCard.<LCID>.resx`). Ships with English (1033), German (1031), French (1036), Spanish (1034), Italian (1040), Dutch (1043), and Japanese (1041). 16 keys cover section headers, action chip aria-labels, expand/collapse labels, copy-button labels, and duration unit suffixes.
  - Non-English translations are machine-generated; customers should have a native speaker review the resx for their language before shipping in production. The English (1033) file is the source of truth.
  - To add a new language, copy `SampleInfoCard.1033.resx`, translate the `<value>` entries, save as `SampleInfoCard.<LCID>.resx`, and register it in `ControlManifest.Input.xml` under `<resources>`.
- **Locale-aware duration formatting** — `formatDuration` now consults `context.formatting.formatInteger` (so `1h 30m` becomes locale-formatted digits where appropriate) and uses resx-driven unit suffixes (`Std/Min` for German, etc.). Falls back to plain ASCII digits if `formatInteger` is unavailable or throws.
- **Accessibility (WCAG 2.1 AA)** —
  - Lookup title and subtitle rows are now keyboard-operable (`role="button"`, `tabIndex={0}`, Enter/Space handlers, `aria-label`).
  - Action chips (phone, email, website, map link) get localized `aria-label` text (`"Call 555-1234"`, `"Open in Maps: 1 Sample Way"`, etc.).
  - Smart-card whole-card toggle exposes `aria-label` ("Expand card" / "Collapse card") in addition to the existing `aria-expanded`.
  - Decorative entity icons inside subtitle rows marked `aria-hidden="true"`.
- **CI workflow** (`.github/workflows/ci.yml`) — every PR and push to `main` runs `npm ci`, lint, build, and the full Jest suite on Node 18.
- 15 new unit tests covering localized strings, locale-aware duration, keyboard a11y, the configurable subtitle separator, and copy-to-clipboard behaviour across form factors.
- **Image / Avatar slot** in the card header. New `imageField` (`SingleLine.Text`, optional input) accepts an entity image URL (`entityimage_url`), a regular URL column, or `@`-syntax to pull from a related lookup. Renders as a 40 px circular avatar to the left of the title in Smart and Contact layouts; falls back to initials computed from the title when the URL is empty or fails to load. Compact layout omits the avatar to preserve dense form-feel. `account` and `contact` presets now bind `imageField → entityimage_url` automatically.

### Changed
- `CONTROL_VERSION` constant in `index.ts` bumped from `4.0.0` to `4.1.0` to match the manifest.
- `InfoCardProps` gains optional `strings?: Partial<InfoCardStrings>`, `subtitleSeparator?: string`, and `formFactor?: number` props. Existing call sites without these continue to work.
- `formatDuration(minutes)` signature is now `formatDuration(minutes, formatting?, strings?)` — backward-compatible (existing 1-arg callers still work).
- Centralized `MaybeFormatting` interface and `getFormatting(context)` helper (replaces the ad-hoc cast at the date-formatting site).

## 4.0.0 — 2026-05-07

First public **sample** release.

### Breaking changes
- **Namespace renamed**: `Contoso.InfoCard` → `Sample.InfoCard`. The control is now registered as `smp_Sample.InfoCard` (publisher prefix `smp`). Existing form bindings that referenced `cli_Contoso.InfoCard` must be updated.
- New shipping solution `InfoCardSolution/` (control-only, publisher `Sample` / prefix `smp`) replaces the previous `InfoCardSolution/` test wrapper. The old test-form solution has been renamed to `InfoCardSampleForm/` and is **not** intended for shipping.

### Added
- Clean shipping `InfoCardSolution/` cdsproj that produces both unmanaged and managed solution zips containing only the control (no test entity bindings).
- README sections: "Re-namespacing for production", "Supported environments", install / uninstall instructions, sample callout banner.
- `CHANGELOG.md`.
- `mergeRelatedFields` is now a top-level export of `InfoCard.tsx` for testability (no more `require()` in tests).

### Changed
- All internal documentation (`docs/*.html`, `.github/copilot-instructions.md`, README) updated to reflect the `Sample` namespace and v4.0.0 versioning.
- Test-form solution publisher and prefix standardised to `Sample` / `smp`; folder renamed to `InfoCardSampleForm/` to reflect that it is a development reference, not a shipping artifact.
- `tools/infocard-config.ts` regex updated to recognise `smp_Sample.InfoCard` form bindings.

### Fixed
- ESLint error in `tests/InfoCard.test.tsx` (top-level `require()` violating `@typescript-eslint/no-var-requires`). `npm run build` now passes the ESLint stage.
- Manifest version (was `3.9.13`), `package.json` version (was `1.0.0`), and README version (was `3.0.0`) are now consistent at `4.0.0`.

### Removed
- Stray root-level test zips: `MobileInfoCardTest_clean.zip`, `MobileInfoCardTest_export.zip`, `MobileInfoCardTest_v3.zip`.

---

## 3.x and earlier

Pre-release internal iterations. See `git log` for details.
