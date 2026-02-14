# ExportDXF

A Windows desktop application that automates exporting flat pattern DXF files from SolidWorks drawings, assemblies, and parts. Built for sheet metal fabrication workflows, it extracts BOM data, generates DXF flat patterns with etch/bend line markings, and exports drawing PDFs to a local output folder with SQL Server tracking.

## Features

- **Batch DXF export** from SolidWorks drawings, assemblies, or individual parts
- **Flat pattern generation** with automatic sheet metal detection
- **Etch line insertion** on bend-up lines for fabrication reference (via EtchBendLines library)
- **PDF export** of SolidWorks drawings
- **Local file export** -- saves DXFs and PDFs to a configurable output folder
- **SQL Server database** -- tracks export history and BOM data via Entity Framework Core
- **View flip control** with automatic, manual, and prefer-up strategies
- **Active document tracking** -- connects to SolidWorks and displays the active document in the title bar
- **BOM extraction** from drawing BOM tables or assembly component trees
- **Drawing history dropdowns** -- equipment and drawing number filters populated from past exports

## Requirements

- Windows 10/11
- .NET 8.0
- SolidWorks (installed and licensed)
- SQL Server (default: `localhost`, database: `ExportDxfDb`, Windows auth)

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
| `ExportDXF.Services` | Core services -- SolidWorks connection, DXF export, BOM extraction, PDF export, file export |
| `ExportDXF.Forms` | WinForms UI -- main export form with log and BOM grids |
| `ExportDXF.Models` | Data models -- BomItem, ExportRecord, ExportContext, SolidWorksDocument |
| `ExportDXF.Data` | Entity Framework Core DbContext for SQL Server persistence |
| `ExportDXF.ViewFlipDeciders` | Strategies for determining if a flat pattern view should be flipped |
| `ExportDXF.ItemExtractors` | Extract component items from BOM tables and assemblies |
| `ExportDXF.Utilities` | SolidWorks helpers, sheet metal property extraction, text utilities |
| `ExportDXF.Extensions` | Extension methods for SolidWorks, UI, unit conversion, strings, TimeSpan |
| `EtchBendLines` | Post-processes DXF files to add etch marks on bend-up lines |

## How It Works

1. **Startup** -- Connects to a running SolidWorks instance (or launches one) asynchronously. Loads past export history from the database to populate equipment/drawing dropdowns. Displays the active SolidWorks document name in the title bar.

2. **Export** -- Based on the active document type:
   - **Drawing**: Extracts BOM items from BOM tables, exports the drawing as PDF, then iterates through each component to generate flat pattern DXFs.
   - **Assembly**: Extracts components from the assembly tree and generates flat pattern DXFs for each sheet metal part.
   - **Part**: Generates a single flat pattern DXF for the active part.

3. **Post-processing** -- Each exported DXF is run through the EtchBendLines library, which identifies bend-up lines and adds short etch marks at their endpoints on a dedicated `ETCH` layer. Only up bends are etched because SolidWorks automatically flips the flat pattern so the shortest flange is the first bend. The etch marks let the press brake operator verify part orientation right off the laser table without flipping -- the first bend (shortest flange) will always have an etch line.

4. **Save** -- DXF and PDF files are copied to the configured output folder. Export records and BOM items (with sheet metal properties) are saved to the SQL Server database.

## Configuration

Settings are in `App.config`:

```xml
<appSettings>
  <add key="MaxBendRadius" value="2.0" />
  <add key="ExportOutputFolder" value="C:\ExportDXF\Output" />
</appSettings>
<connectionStrings>
  <add name="ExportDxfDb"
       connectionString="Server=localhost;Database=ExportDxfDb;Trusted_Connection=True;TrustServerCertificate=True;"
       providerName="Microsoft.Data.SqlClient" />
</connectionStrings>
```

## DXF Filename Format

Exported files follow the pattern: `{DrawingNumber} PT{ItemNo}.dxf`

Example: `5007 A02 PT01.dxf`

## Building

Open `ExportDXF.sln` in Visual Studio and build. The EtchBendLines submodule must be initialized:

```bash
git submodule update --init --recursive
```
