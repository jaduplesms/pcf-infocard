# InfoCardSolution — shipping wrapper

This is the **clean, control-only Dataverse solution** for the InfoCard PCF
control. Build this and import the `.zip` to deploy the control to a customer
environment.

## What's inside

- `Solution.xml` — solution manifest (`UniqueName: InfoCardSample`,
  publisher `Sample` / prefix `smp`, version aligned with the control
  manifest)
- `Customizations.xml` — empty (this solution intentionally ships **no form
  bindings, no entity metadata, no opinions** — only the PCF control)
- `InfoCardSolution.cdsproj` — references `..\InfoCardControl\InfoCardControl.pcfproj`
  so building this project builds the PCF and packages it into the solution

## Build

From this directory:

```bash
# Unmanaged (default)
dotnet build /p:SolutionPackageType=Unmanaged
# → bin/Debug/InfoCardSolution.zip

# Managed (production deploys)
dotnet build /p:SolutionPackageType=Managed
```

## Import

```bash
pac solution import --path bin/Debug/InfoCardSolution.zip --publish-changes
```

## When to bump

Bump `<Version>` in `src/Other/Solution.xml` to match
`InfoCardControl/InfoCard/ControlManifest.Input.xml`'s `version` whenever the
control changes. Keep the two in lockstep.

## Re-namespacing for production

This solution ships as a **sample**. Before rolling out to a customer
production environment you should re-namespace it to your own publisher
(see the root [README → Re-namespacing for production](../README.md#re-namespacing-for-production)).
