# InfoCard Roadmap

> Sample / community roadmap. No SLAs, no commitments — items here are
> ideas the maintainers and contributors think would add value next.
> Issues and PRs welcome on any line item.

**Effort sizing:** XS = under an hour · S = a day · M = a few days · L = 1–2 weeks.

---

## v4.1 — Hardening & polish

Small, high-leverage items focused on production hardening.

| # | Item | Effort | Status |
|---|------|--------|--------|
| 1 | **CI workflow** — `.github/workflows/ci.yml` running `npm ci`, lint, build and `jest` on every PR. | S | ✅ done (`6582606`) |
| 2 | **Accessibility pass** — keyboard focus order on tappable rows, `role="button"` + `aria-label` on action icons, color-contrast audit against the `defaultTheme` fallback. | S | ✅ done |
| 3 | **Locale-aware formatting** — replace English `Xh Xm` duration and ad-hoc date strings with `context.formatting.formatInteger` / `formatDateShort` etc. | M | ✅ done |
| 4 | **Localized labels** — section headers (`Contact`, `Details`, `Info`), action tooltips, etc. via PCF `resx` resources so customers can translate without a fork. Ships with EN, DE, FR, ES, IT, NL, JA. Non-English translations are machine-generated and should be reviewed by a native speaker before production use. | M | ✅ done |

## v4.2 — Customer-requested features

Features anticipated from typical Field Service / model-driven app feedback.

| # | Item | Effort | Status |
|---|------|--------|--------|
| 5 | **More layout presets** — `msdyn_customerasset`, `msdyn_workorderservicetask`, `msdyn_agreement`, fuller `incident`. | S | ✅ done |
| 6 | **Header avatar / image slot** — bind `entityimage` for contact / account / resource cards. | M | ✅ done |
| 7 | **Copy-to-clipboard** affordance on phone / email / address rows (desktop). | S | ✅ done |
| 8 | **Configurable subtitle separator** (currently hardcoded `·`). | XS | ✅ done |
| 21 | **Title prefix** — muted-colour literal rendered before the title text (e.g. `Case: ACME-001`, `Work Order: WO-12345`). New `titlePrefix` (`SingleLine.Text`) input. | XS | ✅ done |
| 22 | **Avatar shape preference** — `imageShape` enum (`rounded` / `circle` / `square`). Default switched to rounded rectangle (8 px) so entity logos and people photos both look natural. | XS | ✅ done |
| 23 | **Configurable collapsible sections** — `collapsibleSections` enum (`none` / `body` / `body-tags` / `all`) lets makers choose what disappears when the card collapses. Now applies to all three layouts (Contact + Compact gain collapse). | M | ✅ done |
| 24 | **Detail row presentation options** — two new config props for the Smart / Contact detail rows: (a) `showDetailIcons` (TwoOptions, default `true`) — when `false`, the auto-detected leading icon (✏️ for instructions, 📋 for summary, etc.) is suppressed and only the value text renders; (b) `detailLabelStyle` (Enum: `none` / `inline-bold` / `above`, default `none`) — when `inline-bold`, prefixes the value with the slot's display name in bold (e.g. **Instructions:** Access via back garden); when `above`, renders the label on its own line as a small caps section header above the value. Useful when the maker wants the detail to read like prose ("Instructions: …") rather than a stand-alone icon row. Both props are independent — makers can mix-and-match (icons off + bold label, etc.). Compact layout already shows labels and is unaffected by these props. | S | ✅ done |
| 17 | **Distance to record (URS-aware)** — show how far the technician (or project resource) is from the record's address. Layered provider: when online and Universal Resource Scheduling is installed, call `msdyn_RetrieveDistanceMatrix` for true routing distance + ETA from the org's configured geospatial provider (Bing / Azure Maps / customer-registered plugin per [URS docs](https://learn.microsoft.com/en-us/dynamics365/common-scheduler/developer/use-preferred-geospatial-data-provider)); fall back to client-side haversine straight-line when offline, when URS isn't installed (e.g. Sales / Customer Service / Power Pages tenants), or when `webAPI.execute` is unavailable (canvas apps). Origin = technician device geolocation via `context.device.getCurrentPosition()` (with `navigator.geolocation` fallback for harness). Destination = existing `latitudeField` / `longitudeField` slots — no new slots needed. New config props: `showDistance` (TwoOptions), `distanceMode` (auto / straight-line / routing-only), `distanceUnit` (auto / km / mi). Display: small chip next to the map link in the contact bar — `🚗 12 km · 18 min` (routing) vs `📍 ~ 8 km away` (straight-line) so users can tell which signal they're seeing. Caches geolocation 2 min and routing results 5 min (keyed on origin/destination coord pair, module-scoped so multiple cards on a form share results). Hides silently on permission denied, missing destination coords, or service-protection limit hit. New resx keys (`Distance_Driving`, `Distance_StraightLine`, `Distance_Unit_Km`, `Distance_Unit_Mi`, `Distance_Unit_Min`) across all 7 locales. | M | |

## v4.3 — Distance follow-ups

Built on v4.2 #17 once shipped.

| # | Item | Effort | Status |
|---|------|--------|--------|
| 18 | **Static origin** — optional `originLatitudeField` / `originLongitudeField` slots so the origin can be a depot / dispatch hub / project resource home location instead of (or alongside) the technician's device geolocation. | S | |
| 19 | **Geocode destination from address text** — when `latitudeField` / `longitudeField` are unbound but `addressField` is, call `msdyn_GeocodeAddress` (URS companion to `msdyn_RetrieveDistanceMatrix`, same provider abstraction) to resolve coordinates server-side. Removes the dependency on customers pre-populating lat/lng. | S | |
| 20 | **Live distance update on move** — `watchPosition` instead of one-shot, with a battery-friendly throttle (refresh only when the technician has moved >100 m). | M | |

## v5 — Architectural

Larger investments that change the shape of the control.

| # | Item | Effort |
|---|------|--------|
| 10 | **Batch WebAPI fetches** — when multiple slots target the same related entity, collapse to a single `retrieveRecord` with a combined `$select`. Currently each slot can trigger its own fetch. | M |
| 11 | **Dataset binding option** — render a card per row from a view (subgrid / list scenarios), not just one card per form. | L |
| 12 | **Form-designer authoring helpers** — author-time validation messages (e.g. warn when `latitudeField` is bound but `longitudeField` is not). | M |

## Documentation & community

| # | Item | Effort | Status |
|---|------|--------|--------|
| 13 | **Animated GIF / short video** in the README showing all three layouts on a real form. | S | |
| 14 | **Sample model-driven app** demonstrating 2–3 entities all using InfoCard, exportable as a solution. | M | |
| 15 | **Issue / PR templates** under `.github/`. | XS | ✅ done |
| 16 | **Layout screenshots** — three PNGs into `docs/images/` (see `docs/images/README.md` for filenames). | XS | |

---

## Recently completed

See [`CHANGELOG.md`](../CHANGELOG.md) for the full list. Highlights of v4.0.0:

- Re-namespaced from `Contoso` / `cli` (sample-y) to `Sample` / `smp` (clearly a sample).
- Clean shipping `InfoCardSolution` cdsproj producing both managed and unmanaged solution zips containing only the control.
- 23 typed slots across 6 groups + 6 config properties; layout presets for `msdyn_workorder`, `bookableresourcebooking`, `account`, `contact`, `incident`.
- Three layouts: Smart Card, Contact Card, Compact Form.
- Related-field syntax (`@field`, `@lookup.field`) for one- and two-hop fetches.
- Fluent Design Language token theming with safe fallback.
