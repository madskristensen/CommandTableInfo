# Progress: 01-convert-project

## Changes Made

### src/CommandTableInfo.csproj
- Replaced entire legacy project file with clean SDK-style format
- `<Project Sdk="Microsoft.NET.Sdk">` root element
- `<TargetFramework>net48</TargetFramework>` (unchanged from v4.8)
- Added `<UseWPF>true</UseWPF>` — required for XAML code-gen (InitializeComponent + named elements)
- Added `<VSSDKBuildToolsAutoSetup>true</VSSDKBuildToolsAutoSetup>`, `<VsixDeployOnDebug>true</VsixDeployOnDebug>`
- Kept `<GeneratePkgDefFile>true</GeneratePkgDefFile>` and `<UseCodebase>true</UseCodebase>`
- Added `<ProjectCapability Include="CreateVsixContainer" />`
- Switched `source.extension1.cs` and `VSCommandTable.cs` to `<Compile Update>` (were `Include`)
- Removed all explicit `<Compile Include>` entries — SDK globbing handles them
- Removed per-config PropertyGroups (Debug/Release) — SDK defaults
- Removed legacy `<Import>` elements
- Removed `StartAction`/`StartProgram`/`StartArguments`
- Removed redundant WPF/WindowsBase refs (UseWPF brings them in automatically)
- Upgraded `Microsoft.VSSDK.BuildTools` from `17.14.2120` → `18.5.40034` (latest; 18.x required for SDK-style; 17.x caused CreatePkgDef TypeLoadException)
- Added `ExcludeAssets="runtime"` to `Microsoft.VisualStudio.SDK`

### src/Properties/AssemblyInfo.cs
- Deleted — SDK auto-generates all standard assembly attributes

## Issues Encountered
- Initial build succeeded for compilation but `CreatePkgDef` failed with `TypeLoadException: GetGenericInstantiation` — this is a runtime incompatibility with VSSDK.BuildTools 18.5.38461
- Fixed by upgrading to 18.5.40034 (latest available)
- First build attempt failed with XAML code-behind errors — fixed by adding `<UseWPF>true</UseWPF>`

## Validation
- Build: ✅ Successful
- VSIX output: ✅ `src/bin/Debug/net48/CommandTableInfo.vsix` (131,524 bytes)
- No duplicate Compile item warnings
- No AssemblyInfo.cs conflict
