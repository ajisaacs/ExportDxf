# ExportDXF

A Windows desktop application that automates exporting flat pattern DXF files from SolidWorks drawings, assemblies, and parts. Built for sheet metal fabrication workflows, it extracts BOM data, generates DXF flat patterns with etch/bend line markings, exports drawing PDFs, and uploads everything to a CutFab API for downstream cut programming.

## Features

- **Batch DXF export** from SolidWorks drawings, assemblies, or individual parts
- **Flat pattern generation** with automatic sheet metal detection
- **Etch line insertion** on bend-up lines for fabrication reference (via EtchBendLines library)
- **PDF export** of SolidWorks drawings
- **CutFab API integration** -- uploads DXFs, PDFs, and BOM data with sheet metal properties (thickness, K-factor, bend radius, material)
- **View flip control** with automatic, manual, and prefer-up strategies
- **Drawing selection UI** that connects to SolidWorks and displays the active document
- **BOM extraction** from drawing BOM tables or assembly component trees

## Requirements

- Windows 10/11
- .NET Framework 4.8
- SolidWorks (installed and licensed)
- CutFab API server (default: `http://localhost:7027`)

## Solution Structure

```
ExportDXF.sln
  ExportDXF/          Main WinForms application
  EtchBendLines/      Library for adding etch lines to DXF files (git submodule)
    netDxf/           DXF file read/write library
```

### Key Namespaces

| Namespace | Purpose |
|-----------|---------|
| `ExportDXF.Services` | Core services -- SolidWorks connection, DXF export, BOM extraction, PDF export, API client |
| `ExportDXF.Forms` | WinForms UI -- drawing selection and main export form |
| `ExportDXF.Models` | Data models -- BomItem, ExportContext, SolidWorksDocument |
| `ExportDXF.ViewFlipDeciders` | Strategies for determining if a flat pattern view should be flipped |
| `ExportDXF.ItemExtractors` | Extract component items from BOM tables and assemblies |
| `ExportDXF.Utilities` | SolidWorks helpers, sheet metal property extraction, text utilities |
| `EtchBendLines` | Post-processes DXF files to add etch marks on bend-up lines |

## How It Works

1. **Startup** -- Connects to a running SolidWorks instance (or launches one). Shows a drawing selection form that pulls equipment/drawing data from the CutFab API and displays the active SolidWorks document.

2. **Export** -- Based on the active document type:
   - **Drawing**: Extracts BOM items from BOM tables, exports the drawing as PDF, then iterates through each component to generate flat pattern DXFs.
   - **Assembly**: Extracts components from the assembly tree and generates flat pattern DXFs for each sheet metal part.
   - **Part**: Generates a single flat pattern DXF for the active part.

3. **Post-processing** -- Each exported DXF is run through the EtchBendLines library, which identifies bend-up lines and adds short etch marks at their endpoints on a dedicated `ETCH` layer. Only up bends are etched because SolidWorks automatically flips the flat pattern so the shortest flange is the first bend. The etch marks let the press brake operator verify part orientation right off the laser table without flipping -- the first bend (shortest flange) will always have an etch line.

4. **Upload** -- DXF files (zipped), PDFs, and BOM item data are uploaded to the CutFab API along with sheet metal properties. The API auto-links cut templates based on material and thickness.

## Configuration

The API base URL is configured in `App.config`:

```xml
<appSettings>
  <add key="CutFab.ApiBaseUrl" value="http://localhost:7027" />
</appSettings>
```

## DXF Filename Format

Exported files follow the pattern: `{EquipmentNo} {DrawingNo} PT{ItemNo}.dxf`

Example: `5007 A02 PT01.dxf`

## Building

Open `ExportDXF.sln` in Visual Studio and build. The EtchBendLines submodule must be initialized:

```bash
git submodule update --init --recursive
```
