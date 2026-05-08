# InfoCardSampleForm — dev sandbox

> ⚠️ **Not for shipping.** This solution exists only to round-trip a known-good
> form configuration into a Field Service development environment for
> hands-on testing of the PCF control. Do **not** import this into a customer
> tenant — it includes a full Bookable Resource Booking form definition that
> would overwrite the customer's own.

## What's inside

- `Solution.xml` — `UniqueName: InfoCardSampleForm`, publisher
  `Sample` / prefix `smp`
- `Customizations.xml` — full Bookable Resource Booking main form XML wired
  to the InfoCard control across the mobile, tablet and desktop form factors,
  including all 24 slot bindings (mix of column-bound and `@`-prefixed
  related-field static inputs)
- `Relationships.xml` — empty
- `InfoCardSampleForm.cdsproj` — references the same
  `..\InfoCardControl\InfoCardControl.pcfproj`

## Build

```bash
dotnet build /p:SolutionPackageType=Unmanaged
# → bin/Debug/InfoCardSampleForm.zip
```

## Import (dev environments only)

```bash
pac solution import --path bin/Debug/InfoCardSampleForm.zip --publish-changes
```

## Why it exists

Re-wiring 24 slots × 3 form factors by hand in the form designer every time
you set up a fresh dev environment is tedious. This solution captures a
reference configuration so you can just import it and start testing.

For a clean **shippable** solution containing only the PCF control, see
[`../InfoCardSolution/`](../InfoCardSolution/README.md) instead.
