# Changelog

All notable changes to the InfoCard PCF control are documented in this file.

The project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
