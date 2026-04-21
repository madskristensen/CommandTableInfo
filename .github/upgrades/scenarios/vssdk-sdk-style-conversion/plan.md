# Plan: VSSDK SDK-Style Conversion

## Tasks

### 01-convert-project
Convert `src/CommandTableInfo.csproj` from legacy to SDK-style format and apply all VSSDK-specific settings and remove legacy artifacts.

### 02-update-solution
Add `Deploy.0` entries to `CommandTableInfo.sln` for F5 debugging support.

### 03-validate
Build the project and verify .vsix output is produced.
