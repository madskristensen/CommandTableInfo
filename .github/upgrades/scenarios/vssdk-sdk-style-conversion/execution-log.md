
## [2026-04-21 14:06] 01-convert-project

Converted CommandTableInfo.csproj to SDK-style. Added UseWPF=true for XAML support, upgraded VSSDK.BuildTools to 18.5.40034 (18.x required for SDK-style; resolved CreatePkgDef TypeLoadException), added VsixDeployOnDebug + VSSDKBuildToolsAutoSetup, switched auto-generated files to Compile Update, removed AssemblyInfo.cs and all legacy cruft. Build successful, .vsix produced.


## [2026-04-21 14:06] 02-update-solution

Added Deploy.0 entries for Debug and Release configurations to CommandTableInfo.sln, enabling F5 VSIX deployment to the experimental VS instance.

