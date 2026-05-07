# InfoCard Roadmap

> Sample / community roadmap. No SLAs, no commitments — items here are
> ideas the maintainers and contributors think would add value next.
> Issues and PRs welcome on any line item.

**Effort sizing:** XS = under an hour · S = a day · M = a few days · L = 1–2 weeks.

---

## v4.1 — Hardening & polish

Small, high-leverage items focused on production hardening.

| # | Item | Effort |
|---|------|--------|
| 1 | **CI workflow** — `.github/workflows/ci.yml` running `npm ci`, lint, build and `jest` on every PR. | S |
| 2 | **Accessibility pass** — keyboard focus order on tappable rows, `role="button"` + `aria-label` on action icons, color-contrast audit against the `defaultTheme` fallback. | S |
| 3 | **Offline-mode UX** — when `context.client.isOffline()` is true and the related-field cache is empty, render an "Offline" indicator instead of `---` rows. Verified end-to-end in Field Service Mobile. | S |
| 4 | **Locale-aware formatting** — replace English `Xh Xm` duration and ad-hoc date strings with `context.formatting.formatInteger` / `formatDateShort` etc. | M |
| 5 | **Localized labels** — section headers (`Contact`, `Details`, `Info`), action tooltips, etc. via PCF `resx` resources so customers can translate without a fork. | M |

## v4.2 — Customer-requested features

Features anticipated from typical Field Service / model-driven app feedback.

| # | Item | Effort |
|---|------|--------|
| 6 | **More layout presets** — `msdyn_customerasset`, `msdyn_workorderservicetask`, `msdyn_agreement`, fuller `incident`. | S |
| 7 | **Header avatar / image slot** — bind `entityimage` for contact / account / resource cards. | M |
| 8 | **Copy-to-clipboard** affordance on phone / email / address rows (desktop). | S |
| 9 | **Configurable subtitle separator** (currently hardcoded `·`). | XS |
| 10 | **Telemetry hook** — optional callback prop so customers can wire fetch failures and render-time metrics to Application Insights. | M |

## v5 — Architectural

Larger investments that change the shape of the control.

| # | Item | Effort |
|---|------|--------|
| 11 | **Batch WebAPI fetches** — when multiple slots target the same related entity, collapse to a single `retrieveRecord` with a combined `$select`. Currently each slot can trigger its own fetch. | M |
| 12 | **Dataset binding option** — render a card per row from a view (subgrid / list scenarios), not just one card per form. | L |
| 13 | **Form-designer authoring helpers** — author-time validation messages (e.g. warn when `latitudeField` is bound but `longitudeField` is not). | M |

## Documentation & community

| # | Item | Effort |
|---|------|--------|
| 14 | **Animated GIF / short video** in the README showing all three layouts on a real form. | S |
| 15 | **Sample model-driven app** demonstrating 2–3 entities all using InfoCard, exportable as a solution. | M |
| 16 | **Issue / PR templates** under `.github/`. | XS |
| 17 | **Layout screenshots** — three PNGs into `docs/images/` (see `docs/images/README.md` for filenames). | XS |

---

## Recently completed

See [`CHANGELOG.md`](../CHANGELOG.md) for the full list. Highlights of v4.0.0:

- Re-namespaced from `Contoso` / `cli` (sample-y) to `Sample` / `smp` (clearly a sample).
- Clean shipping `InfoCardSolution` cdsproj producing both managed and unmanaged solution zips containing only the control.
- 23 typed slots across 6 groups + 6 config properties; layout presets for `msdyn_workorder`, `bookableresourcebooking`, `account`, `contact`, `incident`.
- Three layouts: Smart Card, Contact Card, Compact Form.
- Related-field syntax (`@field`, `@lookup.field`) for one- and two-hop fetches.
- Fluent Design Language token theming with safe fallback.
