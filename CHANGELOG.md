# Changelog

All notable changes to the InfoCard PCF control are documented in this file.

The project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## Unreleased — v4.1.0 (in progress on `main`)

### Added
- **Localization** via standard Dataverse PCF `resx` resources (`Sample/strings/SampleInfoCard.<LCID>.resx`). Ships with English (1033), German (1031), French (1036), Spanish (1034), Italian (1040), Dutch (1043), and Japanese (1041). 14 keys cover section headers, action chip aria-labels, expand/collapse labels, and duration unit suffixes.
  - Non-English translations are machine-generated; customers should have a native speaker review the resx for their language before shipping in production. The English (1033) file is the source of truth.
  - To add a new language, copy `SampleInfoCard.1033.resx`, translate the `<value>` entries, save as `SampleInfoCard.<LCID>.resx`, and register it in `ControlManifest.Input.xml` under `<resources>`.
- **Locale-aware duration formatting** — `formatDuration` now consults `context.formatting.formatInteger` (so `1h 30m` becomes locale-formatted digits where appropriate) and uses resx-driven unit suffixes (`Std/Min` for German, etc.). Falls back to plain ASCII digits if `formatInteger` is unavailable or throws.
- **Accessibility (WCAG 2.1 AA)** —
  - Lookup title and subtitle rows are now keyboard-operable (`role="button"`, `tabIndex={0}`, Enter/Space handlers, `aria-label`).
  - Action chips (phone, email, website, map link) get localized `aria-label` text (`"Call 555-1234"`, `"Open in Maps: 1 Adams Way"`, etc.).
  - Smart-card whole-card toggle exposes `aria-label` ("Expand card" / "Collapse card") in addition to the existing `aria-expanded`.
  - Decorative entity icons inside subtitle rows marked `aria-hidden="true"`.
- **CI workflow** (`.github/workflows/ci.yml`) — every PR and push to `main` runs `npm ci`, lint, build, and the full Jest suite on Node 18.
- 9 new unit tests covering localized strings (`getStrings` translation pass-through, key-as-value fallback, throwing-getString safety), locale-aware duration (`formatInteger` substitution, throw-safety, custom suffixes), and keyboard a11y (Enter/Space on lookup title, non-lookup title not focusable, localized phone aria-label).

### Changed
- `CONTROL_VERSION` constant in `index.ts` bumped from stale `3.9.13` to `4.0.0` to match the manifest. Going forward this gets bumped alongside the manifest version on every deploy.
- `InfoCardProps` gains an optional `strings?: Partial<InfoCardStrings>` prop. Existing call sites without this prop continue to work — `DEFAULT_STRINGS` (English) is merged in at the component root.
- `formatDuration(minutes)` signature is now `formatDuration(minutes, formatting?, strings?)` — backward-compatible (existing 1-arg callers still work).
- Centralized `MaybeFormatting` interface and `getFormatting(context)` helper (replaces the ad-hoc cast at the date-formatting site).

## 4.0.0 — 2026-05-07

First public **sample** release.

### Breaking changes
- **Namespace renamed**: `Contoso.InfoCard` → `Sample.InfoCard`. The control is now registered as `smp_Sample.InfoCard` (publisher prefix `smp`). Existing form bindings that referenced `cli_Contoso.InfoCard` must be updated.
- New shipping solution `InfoCardSolution/` (control-only, publisher `Sample` / prefix `smp`) replaces the previous `InfoCardSolution/` test wrapper. The old test-form solution has been renamed to `InfoCardTestSolution/` and is **not** intended for shipping.

### Added
- Clean shipping `InfoCardSolution/` cdsproj that produces both unmanaged and managed solution zips containing only the control (no test entity bindings).
- README sections: "Re-namespacing for production", "Supported environments", install / uninstall instructions, sample callout banner.
- `CHANGELOG.md`.
- `mergeRelatedFields` is now a top-level export of `InfoCard.tsx` for testability (no more `require()` in tests).

### Changed
- All internal documentation (`docs/*.html`, `.github/copilot-instructions.md`, README) updated to reflect the `Sample` namespace and v4.0.0 versioning.
- `InfoCardTestSolution` publisher renamed from `Jdp` (prefix `jdp`) to `Sample` (prefix `smp`); unique name changed from `MobileInfoCardTest` to `InfoCardTestSolution`.
- `tools/infocard-config.ts` regex updated to recognise `smp_Sample.InfoCard` form bindings.

### Fixed
- ESLint error in `tests/InfoCard.test.tsx` (top-level `require()` violating `@typescript-eslint/no-var-requires`). `npm run build` now passes the ESLint stage.
- Manifest version (was `3.9.13`), `package.json` version (was `1.0.0`), and README version (was `3.0.0`) are now consistent at `4.0.0`.

### Removed
- Stray root-level test zips: `MobileInfoCardTest_clean.zip`, `MobileInfoCardTest_export.zip`, `MobileInfoCardTest_v3.zip`.

---

## 3.x and earlier

Pre-release internal iterations. See `git log` for details.
