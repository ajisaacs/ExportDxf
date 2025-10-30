# Repository Guidelines

## Project Structure & Module Organization
- Root solution: `ExportDXF.sln`.
- WinForms app: `ExportDXF/ExportDXF/` (.NET Framework 4.8; SolidWorks automation). Key folders: `Forms/`, `Services/`, `Models/`, `Utilities/`, `Templates/`, `Resources/`, `ItemExtractors/`, `ViewFlipDeciders/`, `Extensions/`, `Properties/`.
- Companion utilities: `EtchBendLines/` (standalone solution for DXF processing; includes vendored `netDxf`).
- Sample docs: `TestDocs/` contains drawings/fixtures useful for local validation.
- Tests: none yet; add new test projects under `tests/` (e.g., `tests/ExportDXF.Tests/`).

## Build, Test, and Development Commands
- Restore: `msbuild ExportDXF.sln /t:Restore` (or `nuget restore ExportDXF.sln`).
- Build (Release): `msbuild ExportDXF.sln /p:Configuration=Release`.
- Build single project: `msbuild ExportDXF/ExportDXF/ExportDXF.csproj /p:Configuration=Release`.
- Run app: `ExportDXF/ExportDXF/bin/Release/ExportDXF.exe` (or `bin/Debug/ExportDXF.exe` after a Debug build).
- Prereqs: Visual Studio 2019/2022 or Build Tools with .NET Framework 4.8 targeting pack; SolidWorks installed for automation features (app starts without SW but SW-dependent actions require it).
- Format: `dotnet format` from repo root if .NET SDK is installed; otherwise use VS “Format Document”.

## Coding Style & Naming Conventions
- C#: 4-space indent, braces on new lines, nullable where supported.
- Naming: PascalCase for types/methods; camelCase for locals/fields; `Async` suffix for async methods.
- Structure: Keep UI in `Forms/` thin; delegate work to `Services/` and `Utilities/`. One class per file; filename matches type (e.g., `DxfExportService.cs`).
- Interop: Isolate SolidWorks COM/interop behind interfaces for testability; prefer dependency injection where practical.

## Testing Guidelines
- Framework: xUnit recommended. Name files `*Tests.cs` and methods `MethodName_Should_DoThing`.
- Location: `tests/ExportDXF.Tests/` with reference to the app or extracted class libraries.
- Scope: Favor service- and utility-level tests; avoid UI surface. Mock/abstract SolidWorks interop.
- Run: `dotnet test` (add when tests exist).

## Commit & Pull Request Guidelines
- Messages: imperative, present tense with optional scope (e.g., `ExportDXF: improve DXF export options`). Provide rationale in body and any breaking changes.
- PRs include: clear description, linked issues (`Closes #123`), screenshots/GIFs for UI, migration notes, and local run steps.

## Security & Configuration Tips
- Do not commit secrets or license keys.
- App configuration resides in `app.config`; user/machine-specific settings should stay in user config and not be committed.
- Keep SolidWorks version/paths configurable where possible.

## Agent Tools: Roslyn Bridge (C# Analysis)
- Purpose: Use for C# code analysis, symbol queries, diagnostics, and semantic navigation via a local WebAPI bridge to Roslyn.
- Golden rules:
  - Always use this tool first for any C#/Roslyn analysis; do not guess.
  - Always check service health/instances before querying.
  - Default to `solutionName=ExportDXF` for this repo. If you intentionally open and analyze `EtchBendLines.sln`, use `solutionName=EtchBendLines`.
- Service assumptions:
  - WebAPI listens on `http://localhost:5001` when installed and running.
  - Visual Studio must have the target solution open for the instance to appear.
- Quick workflow (ExportDXF):
  - List VS instances: `curl http://localhost:5001/api/instances`
  - Health check: `curl http://localhost:5001/api/health`
  - Diagnostics summary: `curl "http://localhost:5001/api/roslyn/diagnostics/summary?solutionName=ExportDXF"`
  - Errors only: `curl "http://localhost:5001/api/roslyn/diagnostics?solutionName=ExportDXF&severity=error"`
  - Warnings only: `curl "http://localhost:5001/api/roslyn/diagnostics?solutionName=ExportDXF&severity=warning"`
  - All diagnostics: `curl "http://localhost:5001/api/roslyn/diagnostics?solutionName=ExportDXF"`
  - List projects: `curl "http://localhost:5001/api/roslyn/projects?solutionName=ExportDXF"`
  - Solution overview: `curl "http://localhost:5001/api/roslyn/solution/overview?solutionName=ExportDXF"`
  - Find symbol by name: `curl "http://localhost:5001/api/roslyn/symbol/search?solutionName=ExportDXF&symbolName=TypeOrMember"`
  - Symbol at file/position: `curl "http://localhost:5001/api/roslyn/symbol?solutionName=ExportDXF&filePath=C:/full/path/File.cs&line=10&column=5"`
  - Find references: `curl "http://localhost:5001/api/roslyn/references?solutionName=ExportDXF&filePath=C:/full/path/File.cs&line=10&column=5"`
- Notes:
  - Lines are 1-based; columns are 0-based. Use absolute file paths.
  - Get canonical file paths from the `projects` endpoint responses.
  - Prefer this tool over manual grepping for symbols/references/diagnostics.
- Troubleshooting:
  - No instances listed: ensure Visual Studio is running with the solution open; wait up to 60s for discovery.
  - Service not reachable: verify the Roslyn Bridge WebAPI service is installed and running on port 5001.

- Helper script:
  - `scripts/rb.ps1` provides short commands with defaults for this repo (defaults to `-SolutionName ExportDXF`).
  - Examples: `./scripts/rb.ps1 summary`, `./scripts/rb.ps1 projects`, `./scripts/rb.ps1 symbol -SymbolName Program`.
