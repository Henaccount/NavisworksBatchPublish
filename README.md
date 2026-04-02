# Navisworks Batch Search Set Publish (Sample Code - Use at own risk)

<table border=1><tr><td><i>Disclaimer – AI-Generated Code

This code was generated in whole or in part using artificial intelligence tools. While efforts have been made to review and validate the output, AI-generated code may contain errors, omissions, or unintended behavior.

Users of this code should be aware that:

The code may not follow best practices for security, performance, or reliability.
Vulnerabilities may be present, including but not limited to injection flaws, insecure dependencies, or improper handling of data.
The code may be susceptible to issues arising from prompt manipulation (e.g., prompt injection) or unintended generation of unsafe logic.
No guarantees are made regarding correctness, completeness, or fitness for a particular purpose.

It is strongly recommended that this code be thoroughly reviewed, tested, and audited—especially before use in production or security-sensitive environments.

The authors disclaim any liability for damages or issues arising from the use of this code.</i></td></tr></table>

This Navisworks 2026 AddInPlugin exports one NWD per **saved search set**.

For every saved search set in the document, the plugin:

1. Executes the saved search.
2. Hides everything except the matched items (plus their ancestor / descendant context needed to keep them visible).
3. Publishes a new NWD with `ExcludeHiddenItems = true`.

Explicit selection sets are ignored. Search sets inside folders are processed recursively.

## What changed

This version no longer uses coordinate boxes at all.

The plugin id is still the same:

`NavisworksBatchBoxPublish.BatchExport.OAI1`

That means you can keep the same `-ExecuteAddInPlugin` value in your BAT file after rebuilding the DLL.

## Requirements

- Autodesk Navisworks 2026
- .NET Framework 4.8
- Reference to `Autodesk.Navisworks.Api.dll`

The project file assumes Navisworks Manage 2026 is installed at:

`C:\Program Files\Autodesk\Navisworks Manage 2026`

If yours is somewhere else, set `NavisworksInstallationPath` when building.

## Build

Open the `.csproj` in Visual Studio 2022 or build from a Developer Command Prompt.

Example:

```bat
msbuild NavisworksBatchBoxPublish.csproj /p:Configuration=Debug
```

## Run from command line

Recommended while testing:

```bat
"C:\Program Files\Autodesk\Navisworks Manage 2026\roamer.exe" ^
  -ShowGui ^
  -log "C:\Temp\roamer.log" ^
  -OpenFile "C:\Models\MyModel.nwd" ^
  -AddPluginAssembly "C:\Path\To\NavisworksBatchBoxPublish.dll" ^
  -ExecuteAddInPlugin "NavisworksBatchBoxPublish.BatchExport.OAI1" ^
    outdir=C:\Temp\Exports ^
    prefix=Set_ ^
    log=C:\Temp\Exports\NavisworksBatchSearchSetPublish.log ^
  -Exit
```

After it works, switch `-ShowGui` to `-NoGui`.

## Arguments

- `outdir=...` required
- `prefix=...` optional
- `log=...` optional

The parser also accepts `--outdir=...`, `--prefix=...`, and `--log=...`.

If a value contains spaces, quote the whole token, for example:

`"outdir=C:\Users\Your Name\Exports"`

## Output file names

The output file name is based on the saved search set path.

Examples:

- `Architecture/Doors/Fire Doors` -> `Architecture__Doors__Fire Doors.nwd`
- `MEP/Valves` -> `MEP__Valves.nwd`

If two names sanitize to the same file name, a numeric suffix is added.

## Logs

Normal run log default:

`<outdir>\NavisworksBatchSearchSetPublish.log`

Startup trace always goes to:

`%TEMP%\NavisworksBatchSearchSetPublish.start.log`

If the plugin fails before normal logging starts, the fallback error log goes to:

`<outdir>\NavisworksBatchSearchSetPublish.error.log`

or, if `outdir` could not be parsed:

`%TEMP%\NavisworksBatchSearchSetPublish.error.log`

## Notes

- The plugin resets hidden state before each export and again at the end.
- Empty search sets are skipped and logged.
- Only **saved search sets** are exported. Folders and explicit selection sets are not exported directly.
