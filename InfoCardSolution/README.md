# InfoCardSolution — shipping wrapper

This is the **clean, control-only Dataverse solution** for the InfoCard PCF
control. Build this and import the `.zip` to deploy the control to a customer
environment.

## What's inside

- `Solution.xml` — solution manifest (`UniqueName: InfoCardSample`,
  publisher `Sample` / prefix `smp`, version aligned with the control
  manifest)
- `Customizations.xml` — control registration only (this solution intentionally ships **no form
  bindings, no entity metadata, no opinions** — only the PCF control)
- `InfoCardSolution.cdsproj` — references `..\InfoCardControl\InfoCardControl.pcfproj`
  so building this project builds the PCF and packages it into the solution

## Build

From the repository root after `npm ci` in `InfoCardControl`:

```bash
.\scripts\Build-Release.ps1
# → artifacts\InfoCardSolution_4.4.23_unmanaged.zip
# → artifacts\InfoCardSolution_4.4.23_managed.zip
# → artifacts\SHA256SUMS_4.4.23.txt
```

## Import

```bash
pac solution import --path .\artifacts\InfoCardSolution_4.4.23_managed.zip --publish-changes
```

## When to bump

Bump `<Version>` in `src/Other/Solution.xml` to match
`InfoCardControl/InfoCard/ControlManifest.Input.xml`'s `version` whenever the
control changes. Keep the two in lockstep.

The release script uses temporary copies of this wrapper for each package type,
builds a production bundle, and checks both ZIPs before copying them to
`artifacts`. It does not modify `<Managed>` in the source. Direct Release
builds default to managed packaging, so requesting unmanaged explicitly also
requires a matching `<Managed>` value.

## Re-namespacing for production

This solution ships as a **sample**. Before rolling out to a customer
production environment you should re-namespace it to your own publisher
(see the root [README → Re-namespacing for production](../README.md#re-namespacing-for-production)).
