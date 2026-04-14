# ExportDXF

A Windows desktop application that automates exporting flat pattern DXF files from SolidWorks drawings, assemblies, and parts. Built for sheet metal fabrication workflows, it extracts BOM data, generates DXF flat patterns with etch/bend line markings, and exports everything to a local folder with an Excel workbook for downstream tools like [OpenNest](https://github.com/ajisaacs/OpenNest).

## Features

- **Batch DXF export** from SolidWorks drawings, assemblies, or individual parts
- **Flat pattern generation** with automatic sheet metal detection
- **Etch line insertion** on bend-up lines for fabrication reference (via EtchBendLines/ACadSharp)
- **PDF export** of SolidWorks drawings
- **Excel BOM output** -- generates an xlsx workbook with a BOM sheet (direct copy of the SolidWorks BOM table) and a Cut Templates sheet (DXF filenames, thicknesses, K-factors, bend radii, content hashes)
- **Revision tracking** -- content hashing detects unchanged DXFs across re-exports; changed files get revision suffixes (e.g., `PT03 Rev2.dxf`)
- **Configurable filename template** -- format like `4321 A01 PT{item_no:2}` with placeholders for item number, part name, configuration, and material
- **Pluggable drawing info extraction** -- auto-fills the filename template from the document name; extensible via `IDrawingInfoExtractor`
- **View flip control** with automatic, manual, and prefer-up strategies
- **Active document tracking** -- connects to SolidWorks and updates the UI when the active document changes
- **Per-export and app-level logging** to plain text log files

## Requirements

- Windows 10/11
- .NET 8.0
- SolidWorks (installed and licensed)

## Solution Structure

```
ExportDXF.sln
  ExportDXF/          Main WinForms application
  EtchBendLines/      Library for adding etch lines to DXF files (git submodule, uses ACadSharp)
```

### Key Namespaces

| Namespace | Purpose |
|-----------|---------|
| `ExportDXF.Services` | Core services -- SolidWorks connection, DXF export, BOM extraction, PDF export, Excel output, logging, filename template parsing, drawing info extraction |
| `ExportDXF.Forms` | WinForms UI -- main export form with log, BOM, and cut templates grids |
| `ExportDXF.Models` | Data models -- BomItem, CutTemplate, ExportContext, SolidWorksDocument |
| `ExportDXF.ViewFlipDeciders` | Strategies for determining if a flat pattern view should be flipped |
| `ExportDXF.ItemExtractors` | Extract component items from BOM tables and assemblies |
| `ExportDXF.Utilities` | SolidWorks helpers, content hashing, sheet metal property extraction, text utilities |
| `ExportDXF.Extensions` | Extension methods for SolidWorks, UI, unit conversion, strings, TimeSpan |
| `EtchBendLines` | Post-processes DXF files to add etch marks on bend-up lines |

## How It Works

1. **Startup** -- Connects to a running SolidWorks instance (or launches one). Auto-fills the filename template from the active document name using pluggable extractors.

2. **Export** -- Based on the active document type:
   - **Drawing**: Copies the raw BOM table to the Excel BOM sheet, exports the drawing as PDF, then generates flat pattern DXFs for each sheet metal component.
   - **Assembly**: Extracts components from the assembly tree and generates flat pattern DXFs for each sheet metal part.
   - **Part**: Generates a single flat pattern DXF for the active part.

3. **Post-processing** -- Each exported DXF is run through the EtchBendLines library, which identifies bend-up lines and adds short etch marks at their endpoints on a dedicated `ETCH` layer. Only up bends are etched because SolidWorks automatically flips the flat pattern so the shortest flange is the first bend. The etch marks let the press brake operator verify part orientation right off the laser table without flipping.

4. **Save** -- DXF and PDF files are saved to a `Templates/` folder next to the source file. An Excel workbook is written with the BOM and Cut Templates sheets. If re-exporting, content hashes are compared against the existing workbook -- unchanged DXFs are skipped, changed DXFs get a revision suffix.

## Output

```
C:\Projects\4321\
├── 4321 A01.SLDDRW
└── Templates\
    ├── 4321 A01 PT01.dxf
    ├── 4321 A01 PT02.dxf
    ├── 4321 A01 PT03.dxf
    ├── 4321 A01 PT03 Rev2.dxf   (revised, original kept)
    ├── 4321 A01.pdf
    ├── 4321 A01.xlsx
    └── 4321 A01.log
```

### Excel Workbook

**BOM sheet** -- exact copy of all visible columns and rows from the SolidWorks BOM table (Drawing exports only).

**Cut Templates sheet**:

| Item # | File Name | Revision | Thickness | K-Factor | Bend Radius | Content Hash |
|--------|-----------|----------|-----------|----------|-------------|--------------|

## Configuration

Settings are in `App.config`:

```xml
<appSettings>
  <add key="MaxBendRadius" value="2.0" />
  <add key="DefaultSuffix" value="PT{item_no:2}" />
</appSettings>
```

### Filename Template Placeholders

| Placeholder | Description | Example |
|-------------|-------------|---------|
| `{item_no:N}` | Item number, zero-padded to N digits | `{item_no:2}` → `03` |
| `{part_name}` | SolidWorks part name | `Bracket` |
| `{config}` | Configuration name | `Default` |
| `{material}` | Material name | `AISI 304` |

## Building

Open `ExportDXF.sln` in Visual Studio and build. The EtchBendLines submodule must be initialized:

```bash
git submodule update --init --recursive
```
