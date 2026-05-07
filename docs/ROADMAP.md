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
| 3 | **Locale-aware formatting** — replace English `Xh Xm` duration and ad-hoc date strings with `context.formatting.formatInteger` / `formatDateShort` etc. | M |
| 4 | **Localized labels** — section headers (`Contact`, `Details`, `Info`), action tooltips, etc. via PCF `resx` resources so customers can translate without a fork. | M |

## v4.2 — Customer-requested features

Features anticipated from typical Field Service / model-driven app feedback.

| # | Item | Effort |
|---|------|--------|
| 5 | **More layout presets** — `msdyn_customerasset`, `msdyn_workorderservicetask`, `msdyn_agreement`, fuller `incident`. | S |
| 6 | **Header avatar / image slot** — bind `entityimage` for contact / account / resource cards. | M |
| 7 | **Copy-to-clipboard** affordance on phone / email / address rows (desktop). | S |
| 8 | **Configurable subtitle separator** (currently hardcoded `·`). | XS |
| 9 | **Telemetry hook** — optional callback prop so customers can wire fetch failures and render-time metrics to Application Insights. | M |

## v5 — Architectural

Larger investments that change the shape of the control.

| # | Item | Effort |
|---|------|--------|
| 10 | **Batch WebAPI fetches** — when multiple slots target the same related entity, collapse to a single `retrieveRecord` with a combined `$select`. Currently each slot can trigger its own fetch. | M |
| 11 | **Dataset binding option** — render a card per row from a view (subgrid / list scenarios), not just one card per form. | L |
| 12 | **Form-designer authoring helpers** — author-time validation messages (e.g. warn when `latitudeField` is bound but `longitudeField` is not). | M |

## Documentation & community

| # | Item | Effort |
|---|------|--------|
| 13 | **Animated GIF / short video** in the README showing all three layouts on a real form. | S |
| 14 | **Sample model-driven app** demonstrating 2–3 entities all using InfoCard, exportable as a solution. | M |
| 15 | **Issue / PR templates** under `.github/`. | XS |
| 16 | **Layout screenshots** — three PNGs into `docs/images/` (see `docs/images/README.md` for filenames). | XS |

---

## Recently completed

See [`CHANGELOG.md`](../CHANGELOG.md) for the full list. Highlights of v4.0.0:

- Re-namespaced from `Contoso` / `cli` (sample-y) to `Sample` / `smp` (clearly a sample).
- Clean shipping `InfoCardSolution` cdsproj producing both managed and unmanaged solution zips containing only the control.
- 23 typed slots across 6 groups + 6 config properties; layout presets for `msdyn_workorder`, `bookableresourcebooking`, `account`, `contact`, `incident`.
- Three layouts: Smart Card, Contact Card, Compact Form.
- Related-field syntax (`@field`, `@lookup.field`) for one- and two-hop fetches.
- Fluent Design Language token theming with safe fallback.
