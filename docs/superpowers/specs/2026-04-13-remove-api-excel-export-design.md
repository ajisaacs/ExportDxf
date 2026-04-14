# ExportDXF: Remove API, Export to Excel

**Date:** 2026-04-13
**Status:** Draft

## Goal

Remove the FabWorks API and database dependencies from ExportDXF. Replace with local file output: DXF files + an Excel workbook (BOM + Cut Templates) that downstream tools like OpenNest can import directly. Make the tool self-contained — no server setup required.

## Approach

Merge the `feature/fabworks-api` branch into `master` to get its structural improvements (service layer, models, content hashing, EtchBendLines fixes), then remove the API/DB layer and replace with Excel + log file output.

---

## Output Structure

### Output Folder

- Default location: `Templates/` folder in the same directory as the source SolidWorks file
- If `Templates/` already exists, the export writes into it (see Revision Handling below)

### Files Produced

```
C:\Projects\4321\
├── 4321 A01.SLDDRW              (source drawing)
└── Templates\
    ├── 4321 A01 PT01.dxf
    ├── 4321 A01 PT02.dxf
    ├── 4321 A01 PT03.dxf
    ├── 4321 A01 PT03 Rev2.dxf   (revised file, original kept)
    ├── 4321 A01.xlsx
    └── 4321 A01.log
```

### Excel Workbook

Written with **ClosedXML 0.104.2** (same version as OpenNest).

**"BOM" sheet** — exact copy of the SolidWorks BOM table. All visible columns and rows reproduced as-is. Only present when exporting from a Drawing that contains a BOM table. Part and Assembly exports omit this sheet.

**"Cut Templates" sheet** — one row per sheet metal part:

| Column | Description |
|--------|-------------|
| Item # | Item number from the BOM or export sequence |
| File Name | DXF filename stem (no path, no extension) |
| Revision | Revision number (1 for initial, increments on change) |
| Thickness | Sheet metal thickness |
| K-Factor | K-Factor value |
| Bend Radius | Default bend radius |
| Content Hash | SHA256 of DXF content (excluding HEADER section) |

Present for all export types (Part, Assembly, Drawing).

### Log Files

**Per-export log:** `Templates/{prefix}.log`
- Timestamped lines: `2026-04-13 14:32:05 [INFO] Exported 4321 A01 PT03.dxf`
- Levels: INFO, WARNING, ERROR
- Appends on re-export

**App-level log:** `C:\ExportDXF\ExportDXF.log`
- Rolling log capturing all exports across all drawings
- Same format as per-export log

---

## Filename Format Template

### User Interface

A single text box replaces the equipment/drawing dropdowns. The user types a format string that controls DXF and xlsx naming.

**Example:** `4321 A01 PT{item_no:2}`

### Supported Placeholders

| Placeholder | Description | Example Output |
|-------------|-------------|----------------|
| `{item_no:N}` | Item number, zero-padded to N digits | `{item_no:2}` → `03` |
| `{part_name}` | SolidWorks part name | `Bracket` |
| `{config}` | Configuration name | `Default` |
| `{material}` | Material name | `AISI 304` |

### Naming Rules

- DXF filename: full template evaluated per part → `4321 A01 PT03.dxf`
- Excel filename: literal text before the first placeholder → `4321 A01.xlsx`
- Log filename: same stem as excel → `4321 A01.log`
- Validation: template must contain `{item_no}` (or `{item_no:N}`) to avoid filename collisions
- If the template starts with a placeholder (no literal prefix), fall back to the document name for the xlsx/log filename

### Auto-Fill

On document open, the format template text box is auto-filled via `IDrawingInfoExtractor`.

---

## Drawing Info Extraction

### Interface

```csharp
public interface IDrawingInfoExtractor
{
    bool TryExtract(string documentName, out DrawingInfo info);
}

public class DrawingInfo
{
    public string EquipmentNumber { get; set; }
    public string DrawingNumber { get; set; }
    public string DefaultTemplate { get; set; }
}
```

### Implementations

**`DefaultDrawingInfoExtractor`** — uses the document name as literal prefix, appends the default suffix from app.config (`DefaultSuffix` setting, default `PT{item_no:2}`).

Example: document `MyPart.SLDPRT` → template `MyPart {item_no:2}` (or with configured suffix)

**`EquipmentDrawingInfoExtractor`** — parses the `{EquipmentNo} {DrawingNo}` pattern from document names in AJ's workplace format. Builds template like `4321 A01 PT{item_no:2}`.

### Resolution

Extractors are tried in order. First one that returns `true` wins. Falls back to `DefaultDrawingInfoExtractor`.

---

## Revision Handling

When re-exporting to an existing `Templates` folder:

1. Read the existing xlsx Cut Templates sheet (if present)
2. For each part being exported, compute the DXF content hash
3. **Hash matches existing row** → skip, leave existing DXF file untouched
4. **Hash differs from existing row** → write new DXF with revision suffix (e.g., `PT03 Rev2.dxf`), update the xlsx row with new filename, new hash, incremented revision number
5. **New part (no existing row)** → write DXF, add new row to xlsx with revision 1
6. Old revision DXF files are kept in the folder as history

The BOM sheet (if present) is fully rewritten from the current drawing's BOM table on each export.

---

## What We Keep from feature/fabworks-api

- Service layer architecture: `DxfExportService`, `IPartExporter`, `IDrawingExporter`, `IBomExtractor`
- Models: `BomItem`, `CutTemplate`, `Item`, `ExportContext`, `LogEvent`
- `ContentHasher` utility (SHA256 excluding DXF HEADER section)
- All 9 EtchBendLines submodule fixes (ACadSharp migration, bend detection, degree symbol, etc.)
- Async export flow with cancellation
- View flip decider
- SolidWorks user input disabling during export

## What We Remove

- `FabWorksApiClient` and all API DTOs
- `ExportDxfDbContext`, EF Core migrations, SQL Server connection string
- NuGet packages: `Microsoft.EntityFrameworkCore.SqlServer`, `Microsoft.EntityFrameworkCore.Tools`
- Equipment/Drawing dropdowns from MainForm
- Any FabWorks.Core / FabWorks.Api projects (if present after merge)

## What We Add

- `IDrawingInfoExtractor` interface + two implementations
- Format template text box with placeholder parsing
- `ExcelExportService` using ClosedXML — reads/writes BOM + Cut Templates sheets
- `LogFileService` — per-export and app-level log writing
- Output folder logic (Templates folder next to source, revision handling)
- NuGet package: `ClosedXML 0.104.2`
- `DefaultSuffix` setting in app.config

## UI Changes

- Replace equipment/drawing combo boxes with a single format template text box
- Keep 3 tabs: Log Events, Bill of Materials, Cut Templates
- Keep Start/Stop button
- Keep View Flip Decider dropdown
