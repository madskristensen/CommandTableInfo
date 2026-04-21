# Assessment: VSSDK SDK-Style Conversion

## Target Project
| Property | Value |
|----------|-------|
| Project | CommandTableInfo |
| Path | src/CommandTableInfo.csproj |
| Current TFM | net48 (v4.8) |
| Solution format | .sln |
| packages.config | No (already PackageReference) |

## VSIX Components Found
- [x] VSIX manifest (source.extension.vsixmanifest) — generator: VsixManifestGenerator, output: source.extension1.cs
- [x] VSCT command table (VSCommandTable.vsct) — generator: VsctGenerator, output: VSCommandTable.cs
- [x] Tool windows (CommandTableWindow.cs, CommandTableExplorerControl.xaml)
- [x] WPF UI (PresentationCore, PresentationFramework, System.Xaml, WindowsBase references)
- [ ] MEF exports
- [ ] Custom editors
- [ ] Language services

## Current Package References
| Package | Version | Notes |
|---------|---------|-------|
| Microsoft.VisualStudio.SDK | 17.0.32112.339 | keep |
| Microsoft.VSSDK.BuildTools | 17.14.2120 | **MUST upgrade to ≥18.5.38461** |

## Auto-Generated Files (need `Update` not `Include`)
- `source.extension1.cs` — DependentUpon source.extension.vsixmanifest
- `VSCommandTable.cs` — DependentUpon VSCommandTable.vsct

## Files to Remove
- `Properties/AssemblyInfo.cs` — SDK auto-generates assembly attributes

## Properties to Remove
- `StartAction`, `StartProgram`, `StartArguments`
- `ProjectTypeGuids`, `SchemaVersion`, `VSToolsPath`, `NuGetPackageImportStamp`
- `MinimumVisualStudioVersion`, `AppDesignerFolder`, etc.
- Legacy `<Import>` elements (Microsoft.Common.props, Microsoft.CSharp.targets, Microsoft.VsSDK.targets)

## Solution Updates Required
- Add `Deploy.0` entries for project GUID `{F484F288-F083-45BA-B96D-C4D96C22F565}` in `.sln`

## Baseline
- Project builds: Yes (has bin/Debug output present)
- Solution builds: Yes

## Key Findings
- VSSDK.BuildTools 17.14.2120 is BELOW the required minimum 18.5.38461 — must upgrade
- No packages.config — skip migration step
- WPF references must be kept explicitly (PresentationCore, PresentationFramework, etc.)
- `source.extension1.cs` is the auto-gen filename (not the standard `source.extension.cs`)
