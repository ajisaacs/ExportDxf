# Remove API, Export to Excel — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Remove FabWorks API and database dependencies from ExportDXF; export DXF files + a BOM Excel workbook to a local `Templates` folder, with content-hash-based revision tracking.

**Architecture:** Merge the `feature/fabworks-api` branch for its structural improvements (service layer, async flow, content hashing, EtchBendLines fixes), then replace the API/DB output layer with local file export + ClosedXML Excel writing. A format template system with pluggable `IDrawingInfoExtractor` handles filename generation.

**Tech Stack:** .NET 8.0, WinForms, ClosedXML 0.104.2, SolidWorks COM Interop, EtchBendLines submodule

---

## File Structure

### New Files
- `ExportDXF/Services/FilenameTemplateParser.cs` — parses `{item_no:N}`, `{part_name}`, `{config}`, `{material}` placeholders
- `ExportDXF/Services/IDrawingInfoExtractor.cs` — interface for extracting drawing info from document names
- `ExportDXF/Services/DefaultDrawingInfoExtractor.cs` — fallback: document name + configured suffix
- `ExportDXF/Services/EquipmentDrawingInfoExtractor.cs` — parses `{EquipmentNo} {DrawingNo}` pattern
- `ExportDXF/Models/DrawingInfo.cs` — result of drawing info extraction
- `ExportDXF/Services/ExcelExportService.cs` — reads/writes BOM.xlsx with ClosedXML
- `ExportDXF/Services/LogFileService.cs` — per-export and app-level log file writing
- `ExportDXF/Services/IRawBomTableReader.cs` — reads raw BOM table columns/rows from SolidWorks drawings
- `ExportDXF/Services/RawBomTableReader.cs` — implementation

### Modified Files
- `ExportDXF/ExportDXF.csproj` — remove EF Core packages, add ClosedXML
- `ExportDXF/Program.cs` — rewire DI (remove API client, DB context; add new services)
- `ExportDXF/Forms/MainForm.cs` — replace dropdowns with format template textbox
- `ExportDXF/Forms/MainForm.Designer.cs` — UI layout changes
- `ExportDXF/App.config` — remove connection string, add DefaultSuffix
- `ExportDXF/Models/CutTemplate.cs` — add Revision property
- `ExportDXF/Models/ExportContext.cs` — replace Equipment/DrawingNo/Title with template + output folder
- `ExportDXF/Services/DxfExportService.cs` — replace API calls with local file writes + Excel updates

### Deleted Files
- `ExportDXF/Services/FabWorksApiClient.cs`
- `ExportDXF/Services/IFabWorksApiClient.cs`
- `ExportDXF/Services/FabWorksApiDtos.cs`
- `ExportDXF/Data/ExportDxfDbContext.cs` (if present after merge)
- `ExportDXF/Data/Migrations/*` (if present after merge)
- `FabWorks.Core/` (entire project directory, if present after merge)
- `FabWorks.Api/` (entire project directory, if present after merge)

---

### Task 1: Merge fabworks-api branch and clean up foreign projects

**Files:**
- Modify: `ExportDXF.sln`
- Delete: `FabWorks.Core/` and `FabWorks.Api/` directories (if present after merge)

- [ ] **Step 1: Merge the branch**

```bash
git merge origin/feature/fabworks-api -m "merge: bring fabworks-api structural improvements into master"
```

Resolve any conflicts favoring the fabworks-api version for files in `ExportDXF/Services/` and `ExportDXF/Models/`. For `ExportDXF.csproj` and `Program.cs`, take fabworks-api version (we'll modify them in later tasks).

- [ ] **Step 2: Remove FabWorks projects from solution and disk**

Check what projects/directories exist after merge:

```bash
ls -d FabWorks*/ 2>/dev/null
grep -i "fabworks" ExportDXF.sln
```

Remove any FabWorks project directories and their solution references:

```bash
dotnet sln ExportDXF.sln remove FabWorks.Core/FabWorks.Core.csproj 2>/dev/null
dotnet sln ExportDXF.sln remove FabWorks.Api/FabWorks.Api.csproj 2>/dev/null
rm -rf FabWorks.Core/ FabWorks.Api/
```

- [ ] **Step 3: Commit**

```bash
git add -A
git commit -m "chore: remove FabWorks.Core and FabWorks.Api projects after merge"
```

---

### Task 2: Remove API client, database, and EF Core

**Files:**
- Delete: `ExportDXF/Services/FabWorksApiClient.cs`, `ExportDXF/Services/IFabWorksApiClient.cs`, `ExportDXF/Services/FabWorksApiDtos.cs`
- Delete: `ExportDXF/Data/` directory (DbContext + migrations)
- Modify: `ExportDXF/ExportDXF.csproj` — remove EF Core packages, add ClosedXML
- Modify: `ExportDXF/App.config` — remove connection string and FabWorksApiUrl

- [ ] **Step 1: Delete API client files**

```bash
rm -f ExportDXF/Services/FabWorksApiClient.cs
rm -f ExportDXF/Services/IFabWorksApiClient.cs
rm -f ExportDXF/Services/FabWorksApiDtos.cs
```

- [ ] **Step 2: Delete database files**

```bash
rm -rf ExportDXF/Data/
```

- [ ] **Step 3: Update ExportDXF.csproj**

Remove EF Core packages and add ClosedXML:

```bash
cd ExportDXF
dotnet remove package Microsoft.EntityFrameworkCore.SqlServer 2>/dev/null
dotnet remove package Microsoft.EntityFrameworkCore.Tools 2>/dev/null
dotnet remove package Microsoft.EntityFrameworkCore.Design 2>/dev/null
dotnet add package ClosedXML --version 0.104.2
cd ..
```

- [ ] **Step 4: Update App.config**

Remove `FabWorksApiUrl` setting and the connection string. Add `DefaultSuffix` setting.

The `appSettings` section should contain:

```xml
<appSettings>
  <add key="MaxBendRadius" value="2.0" />
  <add key="DefaultSuffix" value="PT{item_no:2}" />
</appSettings>
```

Remove any `<connectionStrings>` section entirely.

- [ ] **Step 5: Commit**

```bash
git add -A
git commit -m "refactor: remove API client, database, and EF Core dependencies"
```

---

### Task 3: Create DrawingInfo model and IDrawingInfoExtractor

**Files:**
- Create: `ExportDXF/Models/DrawingInfo.cs`
- Create: `ExportDXF/Services/IDrawingInfoExtractor.cs`
- Create: `ExportDXF/Services/DefaultDrawingInfoExtractor.cs`
- Create: `ExportDXF/Services/EquipmentDrawingInfoExtractor.cs`

- [ ] **Step 1: Create DrawingInfo model**

Create `ExportDXF/Models/DrawingInfo.cs`:

```csharp
namespace ExportDXF.Models
{
    public class DrawingInfo
    {
        public string EquipmentNumber { get; set; }
        public string DrawingNumber { get; set; }
        public string DefaultTemplate { get; set; }
    }
}
```

- [ ] **Step 2: Create IDrawingInfoExtractor interface**

Create `ExportDXF/Services/IDrawingInfoExtractor.cs`:

```csharp
using ExportDXF.Models;

namespace ExportDXF.Services
{
    public interface IDrawingInfoExtractor
    {
        bool TryExtract(string documentName, out DrawingInfo info);
    }
}
```

- [ ] **Step 3: Create EquipmentDrawingInfoExtractor**

Create `ExportDXF/Services/EquipmentDrawingInfoExtractor.cs`:

```csharp
using System.Text.RegularExpressions;
using ExportDXF.Models;

namespace ExportDXF.Services
{
    public class EquipmentDrawingInfoExtractor : IDrawingInfoExtractor
    {
        private static readonly Regex Pattern = new Regex(
            @"^(?<equip>\d+)\s+(?<drawing>[A-Za-z]\d+)",
            RegexOptions.Compiled);

        public bool TryExtract(string documentName, out DrawingInfo info)
        {
            info = null;

            // Strip file extension
            var name = Path.GetFileNameWithoutExtension(documentName);

            var match = Pattern.Match(name);
            if (!match.Success)
                return false;

            var equip = match.Groups["equip"].Value;
            var drawing = match.Groups["drawing"].Value;

            info = new DrawingInfo
            {
                EquipmentNumber = equip,
                DrawingNumber = drawing,
                DefaultTemplate = $"{equip} {drawing} PT{{item_no:2}}"
            };

            return true;
        }
    }
}
```

- [ ] **Step 4: Create DefaultDrawingInfoExtractor**

Create `ExportDXF/Services/DefaultDrawingInfoExtractor.cs`:

```csharp
using System.Configuration;
using ExportDXF.Models;

namespace ExportDXF.Services
{
    public class DefaultDrawingInfoExtractor : IDrawingInfoExtractor
    {
        public bool TryExtract(string documentName, out DrawingInfo info)
        {
            var name = Path.GetFileNameWithoutExtension(documentName);
            var suffix = ConfigurationManager.AppSettings["DefaultSuffix"] ?? "PT{item_no:2}";

            info = new DrawingInfo
            {
                EquipmentNumber = null,
                DrawingNumber = null,
                DefaultTemplate = $"{name} {suffix}"
            };

            return true;
        }
    }
}
```

- [ ] **Step 5: Commit**

```bash
git add ExportDXF/Models/DrawingInfo.cs ExportDXF/Services/IDrawingInfoExtractor.cs ExportDXF/Services/DefaultDrawingInfoExtractor.cs ExportDXF/Services/EquipmentDrawingInfoExtractor.cs
git commit -m "feat: add IDrawingInfoExtractor with equipment and default implementations"
```

---

### Task 4: Create FilenameTemplateParser

**Files:**
- Create: `ExportDXF/Services/FilenameTemplateParser.cs`

- [ ] **Step 1: Create FilenameTemplateParser**

Create `ExportDXF/Services/FilenameTemplateParser.cs`:

```csharp
using System.Text.RegularExpressions;
using ExportDXF.Models;

namespace ExportDXF.Services
{
    public static class FilenameTemplateParser
    {
        private static readonly Regex PlaceholderPattern = new Regex(
            @"\{(?<name>\w+)(?::(?<pad>\d+))?\}",
            RegexOptions.Compiled);

        /// <summary>
        /// Evaluates a template string for a given item.
        /// e.g. "4321 A01 PT{item_no:2}" with item_no=3 → "4321 A01 PT03"
        /// </summary>
        public static string Evaluate(string template, Item item)
        {
            return PlaceholderPattern.Replace(template, match =>
            {
                var name = match.Groups["name"].Value.ToLowerInvariant();
                var padStr = match.Groups["pad"].Value;
                int pad = string.IsNullOrEmpty(padStr) ? 0 : int.Parse(padStr);

                string value;
                switch (name)
                {
                    case "item_no":
                        value = item.ItemNo.ToString();
                        if (pad > 0)
                            value = value.PadLeft(pad, '0');
                        break;
                    case "part_name":
                        value = item.PartName ?? "";
                        break;
                    case "config":
                        value = item.Configuration ?? "";
                        break;
                    case "material":
                        value = item.Material ?? "";
                        break;
                    default:
                        value = match.Value; // leave unknown placeholders as-is
                        break;
                }

                return value;
            });
        }

        /// <summary>
        /// Extracts the literal prefix before the first placeholder.
        /// Used for naming the xlsx and log files.
        /// Falls back to documentName if prefix is empty.
        /// </summary>
        public static string GetPrefix(string template, string documentName)
        {
            var match = PlaceholderPattern.Match(template);
            if (!match.Success)
                return template.Trim();

            var prefix = template.Substring(0, match.Index).Trim();
            if (string.IsNullOrEmpty(prefix))
                return Path.GetFileNameWithoutExtension(documentName);

            return prefix;
        }

        /// <summary>
        /// Validates that the template contains {item_no} to prevent filename collisions.
        /// </summary>
        public static bool Validate(string template, out string error)
        {
            error = null;

            if (string.IsNullOrWhiteSpace(template))
            {
                error = "Template cannot be empty.";
                return false;
            }

            var hasItemNo = PlaceholderPattern.IsMatch(template) &&
                PlaceholderPattern.Matches(template)
                    .Cast<Match>()
                    .Any(m => m.Groups["name"].Value.ToLowerInvariant() == "item_no");

            if (!hasItemNo)
            {
                error = "Template must contain {item_no} or {item_no:N} to avoid filename collisions.";
                return false;
            }

            return true;
        }
    }
}
```

- [ ] **Step 2: Commit**

```bash
git add ExportDXF/Services/FilenameTemplateParser.cs
git commit -m "feat: add FilenameTemplateParser with placeholder evaluation and validation"
```

---

### Task 5: Update CutTemplate model with Revision

**Files:**
- Modify: `ExportDXF/Models/CutTemplate.cs`

- [ ] **Step 1: Add Revision property to CutTemplate**

Add `Revision` property (int, default 1) to `CutTemplate.cs`. Remove any EF Core attributes or database-specific annotations (foreign keys to ExportRecord, etc.):

```csharp
namespace ExportDXF.Models
{
    public class CutTemplate
    {
        public string DxfFilePath { get; set; }
        public string ContentHash { get; set; }
        public string CutTemplateName { get; set; }
        public int Revision { get; set; } = 1;
        public double Thickness { get; set; }
        public double KFactor { get; set; }
        public double DefaultBendRadius { get; set; }
        public int BomItemId { get; set; }
    }
}
```

- [ ] **Step 2: Clean up BomItem.cs**

Remove any EF Core navigation properties or database annotations. Keep the data properties:

```csharp
namespace ExportDXF.Models
{
    public class BomItem
    {
        public int ItemNo { get; set; }
        public string PartNo { get; set; }
        public int SortOrder { get; set; }
        public int Qty { get; set; }
        public int TotalQty { get; set; }
        public string Description { get; set; }
        public string PartName { get; set; }
        public string ConfigurationName { get; set; }
        public string Material { get; set; }
        public CutTemplate CutTemplate { get; set; }
    }
}
```

- [ ] **Step 3: Remove ExportRecord.cs**

The export record model was for DB persistence. Delete it — the log file replaces this purpose:

```bash
rm -f ExportDXF/Models/ExportRecord.cs
```

- [ ] **Step 4: Update ExportContext.cs**

Replace API-oriented fields with template and output folder:

```csharp
using SldWorks;

namespace ExportDXF.Models
{
    public class ExportContext
    {
        public SolidWorksDocument ActiveDocument { get; set; }
        public IViewFlipDecider ViewFlipDecider { get; set; }
        public string FilenameTemplate { get; set; }
        public string OutputFolder { get; set; }
        public CancellationToken CancellationToken { get; set; }
        public Action<string, LogLevel, string> ProgressCallback { get; set; }
        public Action<BomItem> BomItemCallback { get; set; }
        public SldWorks.SldWorks SolidWorksApp { get; set; }

        private DrawingDoc _templateDrawing;

        public DrawingDoc GetOrCreateTemplateDrawing()
        {
            if (_templateDrawing != null)
                return _templateDrawing;

            var templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "Blank.drwdot");
            var swApp = SolidWorksApp;
            int errors = 0, warnings = 0;
            var doc = swApp.NewDocument(templatePath, 0, 0, 0);
            _templateDrawing = (DrawingDoc)doc;
            return _templateDrawing;
        }

        public void CleanupTemplateDrawing()
        {
            if (_templateDrawing == null) return;
            var swApp = SolidWorksApp;
            swApp.CloseDoc(((ModelDoc2)_templateDrawing).GetTitle());
            _templateDrawing = null;
        }

        public void LogProgress(string message, LogLevel level = LogLevel.Info, string file = null)
        {
            ProgressCallback?.Invoke(message, level, file);
        }
    }
}
```

Note: Preserve the exact `GetOrCreateTemplateDrawing` and `CleanupTemplateDrawing` implementations from the fabworks-api branch — the above is approximate. The key change is replacing `Equipment`, `DrawingNo`, `Title`, `FilePrefix`, `EquipmentId` with `FilenameTemplate` and `OutputFolder`.

- [ ] **Step 5: Commit**

```bash
git add -A
git commit -m "refactor: update models — add CutTemplate.Revision, remove ExportRecord, simplify ExportContext"
```

---

### Task 6: Create ExcelExportService

**Files:**
- Create: `ExportDXF/Services/ExcelExportService.cs`

- [ ] **Step 1: Create ExcelExportService**

Create `ExportDXF/Services/ExcelExportService.cs`:

```csharp
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using ExportDXF.Models;

namespace ExportDXF.Services
{
    public class ExcelExportService
    {
        /// <summary>
        /// Reads existing Cut Templates from an xlsx file to compare content hashes.
        /// Returns empty dictionary if file doesn't exist.
        /// Key = Item #, Value = (ContentHash, Revision, FileName)
        /// </summary>
        public Dictionary<int, (string ContentHash, int Revision, string FileName)> ReadExistingCutTemplates(string xlsxPath)
        {
            var result = new Dictionary<int, (string, int, string)>();

            if (!File.Exists(xlsxPath))
                return result;

            using (var workbook = new XLWorkbook(xlsxPath))
            {
                if (!workbook.TryGetWorksheet("Cut Templates", out var ws))
                    return result;

                var headerRow = ws.Row(1);
                var headers = new Dictionary<string, int>();
                for (int col = 1; col <= ws.LastColumnUsed()?.ColumnNumber() ?? 0; col++)
                {
                    headers[headerRow.Cell(col).GetString()] = col;
                }

                if (!headers.ContainsKey("Item #") || !headers.ContainsKey("Content Hash"))
                    return result;

                for (int row = 2; row <= ws.LastRowUsed()?.RowNumber() ?? 1; row++)
                {
                    var itemNo = ws.Cell(row, headers["Item #"]).GetValue<int>();
                    var hash = ws.Cell(row, headers["Content Hash"]).GetString();
                    var revision = headers.ContainsKey("Revision")
                        ? ws.Cell(row, headers["Revision"]).GetValue<int>()
                        : 1;
                    var fileName = headers.ContainsKey("File Name")
                        ? ws.Cell(row, headers["File Name"]).GetString()
                        : "";

                    result[itemNo] = (hash, revision, fileName);
                }
            }

            return result;
        }

        /// <summary>
        /// Writes or updates the xlsx file with BOM and Cut Templates sheets.
        /// rawBomTable: list of rows where each row is a dictionary of column name → value.
        /// If null, the BOM sheet is not written (Part/Assembly exports).
        /// </summary>
        public void Write(
            string xlsxPath,
            List<Dictionary<string, string>> rawBomTable,
            List<BomItem> bomItems)
        {
            using (var workbook = File.Exists(xlsxPath)
                ? new XLWorkbook(xlsxPath)
                : new XLWorkbook())
            {
                // BOM sheet — full rewrite from drawing BOM table
                if (rawBomTable != null && rawBomTable.Count > 0)
                {
                    if (workbook.TryGetWorksheet("BOM", out var existingBom))
                        workbook.Worksheets.Delete("BOM");

                    var bomSheet = workbook.Worksheets.Add("BOM");
                    var columns = rawBomTable[0].Keys.ToList();

                    // Headers
                    for (int col = 0; col < columns.Count; col++)
                        bomSheet.Cell(1, col + 1).Value = columns[col];

                    // Data rows
                    for (int row = 0; row < rawBomTable.Count; row++)
                    {
                        for (int col = 0; col < columns.Count; col++)
                        {
                            var value = rawBomTable[row].GetValueOrDefault(columns[col], "");
                            bomSheet.Cell(row + 2, col + 1).Value = value;
                        }
                    }

                    bomSheet.Columns().AdjustToContents();
                }

                // Cut Templates sheet — full rewrite with current data
                if (workbook.TryGetWorksheet("Cut Templates", out var existingCt))
                    workbook.Worksheets.Delete("Cut Templates");

                var ctSheet = workbook.Worksheets.Add("Cut Templates");

                // Headers
                var ctHeaders = new[] { "Item #", "File Name", "Revision", "Thickness", "K-Factor", "Bend Radius", "Content Hash" };
                for (int col = 0; col < ctHeaders.Length; col++)
                    ctSheet.Cell(1, col + 1).Value = ctHeaders[col];

                // Data rows — only BOM items that have a CutTemplate
                int ctRow = 2;
                foreach (var item in bomItems.Where(b => b.CutTemplate != null).OrderBy(b => b.ItemNo))
                {
                    var ct = item.CutTemplate;
                    ctSheet.Cell(ctRow, 1).Value = item.ItemNo;
                    ctSheet.Cell(ctRow, 2).Value = ct.DxfFilePath;
                    ctSheet.Cell(ctRow, 3).Value = ct.Revision;
                    ctSheet.Cell(ctRow, 4).Value = ct.Thickness;
                    ctSheet.Cell(ctRow, 5).Value = ct.KFactor;
                    ctSheet.Cell(ctRow, 6).Value = ct.DefaultBendRadius;
                    ctSheet.Cell(ctRow, 7).Value = ct.ContentHash;
                    ctRow++;
                }

                ctSheet.Columns().AdjustToContents();

                workbook.SaveAs(xlsxPath);
            }
        }
    }
}
```

- [ ] **Step 2: Commit**

```bash
git add ExportDXF/Services/ExcelExportService.cs
git commit -m "feat: add ExcelExportService for BOM and Cut Templates xlsx output"
```

---

### Task 7: Create LogFileService

**Files:**
- Create: `ExportDXF/Services/LogFileService.cs`

- [ ] **Step 1: Create LogFileService**

Create `ExportDXF/Services/LogFileService.cs`:

```csharp
using System;
using System.IO;

namespace ExportDXF.Services
{
    public class LogFileService : IDisposable
    {
        private StreamWriter _exportLog;
        private static readonly string AppLogPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "ExportDXF", "ExportDXF.log");

        public void StartExportLog(string logFilePath)
        {
            var dir = Path.GetDirectoryName(logFilePath);
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            _exportLog = new StreamWriter(logFilePath, append: true);
        }

        public void Log(string level, string message)
        {
            var line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{level}] {message}";

            _exportLog?.WriteLine(line);
            _exportLog?.Flush();

            WriteAppLog(line);
        }

        public void LogInfo(string message) => Log("INFO", message);
        public void LogWarning(string message) => Log("WARNING", message);
        public void LogError(string message) => Log("ERROR", message);

        private void WriteAppLog(string line)
        {
            var dir = Path.GetDirectoryName(AppLogPath);
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            File.AppendAllText(AppLogPath, line + Environment.NewLine);
        }

        public void Dispose()
        {
            _exportLog?.Dispose();
            _exportLog = null;
        }
    }
}
```

- [ ] **Step 2: Commit**

```bash
git add ExportDXF/Services/LogFileService.cs
git commit -m "feat: add LogFileService for per-export and app-level logging"
```

---

### Task 8: Create RawBomTableReader

**Files:**
- Create: `ExportDXF/Services/RawBomTableReader.cs`

- [ ] **Step 1: Create RawBomTableReader**

This reads the SolidWorks BOM table as raw column/row data (all visible columns, all visible rows) for direct copy into the Excel BOM sheet.

Create `ExportDXF/Services/RawBomTableReader.cs`:

```csharp
using System.Collections.Generic;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace ExportDXF.Services
{
    public static class RawBomTableReader
    {
        /// <summary>
        /// Reads all visible columns and rows from a BOM table annotation.
        /// Returns a list of rows, each row a dictionary of column header → cell value.
        /// </summary>
        public static List<Dictionary<string, string>> Read(ITableAnnotation table)
        {
            var rows = new List<Dictionary<string, string>>();

            int colCount = table.ColumnCount;
            int rowCount = table.RowCount;

            // Build column headers (visible only)
            var columns = new List<(int Index, string Header)>();
            for (int col = 0; col < colCount; col++)
            {
                if (table.ColumnHidden[col])
                    continue;

                var header = table.get_Text(0, col)?.Trim() ?? $"Column{col}";
                columns.Add((col, header));
            }

            // Read data rows (skip header row 0, skip hidden rows)
            for (int row = 1; row < rowCount; row++)
            {
                if (table.RowHidden[row])
                    continue;

                var rowData = new Dictionary<string, string>();
                foreach (var (colIdx, header) in columns)
                {
                    rowData[header] = table.get_Text(row, colIdx)?.Trim() ?? "";
                }
                rows.Add(rowData);
            }

            return rows;
        }
    }
}
```

- [ ] **Step 2: Commit**

```bash
git add ExportDXF/Services/RawBomTableReader.cs
git commit -m "feat: add RawBomTableReader for copying SolidWorks BOM tables to Excel"
```

---

### Task 9: Update DxfExportService for local export

**Files:**
- Modify: `ExportDXF/Services/DxfExportService.cs`

This is the largest change. The service currently calls the API for file uploads, record creation, and BOM item persistence. We replace all of that with:
- Write DXF files directly to the output folder
- Use FilenameTemplateParser for naming
- Use ExcelExportService for revision tracking and xlsx output
- Use LogFileService for logging

- [ ] **Step 1: Update DxfExportService constructor and dependencies**

Replace `IFabWorksApiClient` dependency with `ExcelExportService` and `LogFileService`. Remove all API-related fields.

The constructor should accept:

```csharp
public DxfExportService(
    IPartExporter partExporter,
    IDrawingExporter drawingExporter,
    IBomExtractor bomExtractor,
    ExcelExportService excelExportService,
    LogFileService logFileService)
```

- [ ] **Step 2: Rewrite ExportAsync to use local output**

The main export flow becomes:

```csharp
public async Task ExportAsync(ExportContext context)
{
    var outputFolder = context.OutputFolder;
    if (!Directory.Exists(outputFolder))
        Directory.CreateDirectory(outputFolder);

    var prefix = FilenameTemplateParser.GetPrefix(
        context.FilenameTemplate,
        context.ActiveDocument.Title);

    var xlsxPath = Path.Combine(outputFolder, $"{prefix}.xlsx");
    var logPath = Path.Combine(outputFolder, $"{prefix}.log");

    _logFileService.StartExportLog(logPath);
    _logFileService.LogInfo($"Export started: {context.ActiveDocument.FilePath}");

    // Read existing cut templates for revision comparison
    var existingTemplates = _excelExportService.ReadExistingCutTemplates(xlsxPath);

    var bomItems = new List<BomItem>();
    List<Dictionary<string, string>> rawBomTable = null;

    try
    {
        switch (context.ActiveDocument.DocumentType)
        {
            case DocumentType.Part:
                await ExportPartAsync(context, outputFolder, existingTemplates, bomItems);
                break;
            case DocumentType.Assembly:
                await ExportAssemblyAsync(context, outputFolder, existingTemplates, bomItems);
                break;
            case DocumentType.Drawing:
                rawBomTable = await ExportDrawingAsync(context, outputFolder, existingTemplates, bomItems);
                break;
        }

        // Write Excel file
        _excelExportService.Write(xlsxPath, rawBomTable, bomItems);
        _logFileService.LogInfo($"Wrote {xlsxPath}");
    }
    catch (OperationCanceledException)
    {
        _logFileService.LogWarning("Export cancelled by user");
        throw;
    }
    catch (Exception ex)
    {
        _logFileService.LogError($"Export failed: {ex.Message}");
        throw;
    }
    finally
    {
        context.CleanupTemplateDrawing();
    }
}
```

- [ ] **Step 3: Rewrite ExportItemsAsync for local file output with revision tracking**

For each item that is sheet metal:

1. Export DXF to temp location via `_partExporter`
2. Run EtchBendLines processing
3. Compute content hash
4. Check `existingTemplates` for matching item number:
   - **Hash matches** → skip file write, reuse existing filename and revision
   - **Hash differs** → increment revision, write file with revision suffix (e.g., `PT03 Rev2.dxf`)
   - **New item** → write file, revision = 1
5. Create BomItem + CutTemplate, invoke callback

The DXF filename is generated by `FilenameTemplateParser.Evaluate(context.FilenameTemplate, item)`.

Revision suffix logic:

```csharp
private string GetDxfFileName(string baseName, int revision)
{
    if (revision <= 1)
        return $"{baseName}.dxf";
    return $"{baseName} Rev{revision}.dxf";
}
```

- [ ] **Step 4: Update ExportDrawingAsync to read raw BOM table**

Before extracting items, read the raw BOM table for the Excel BOM sheet:

```csharp
private async Task<List<Dictionary<string, string>>> ExportDrawingAsync(
    ExportContext context,
    string outputFolder,
    Dictionary<int, (string ContentHash, int Revision, string FileName)> existingTemplates,
    List<BomItem> bomItems)
{
    var drawingDoc = (DrawingDoc)context.ActiveDocument.NativeDocument;

    // Read raw BOM table for Excel output
    var rawBomTable = new List<Dictionary<string, string>>();
    var bomFeature = (BomFeature)/* get BOM feature from drawing */;
    // Use existing BomExtractor to get the table annotations, then RawBomTableReader.Read()

    // Export PDF
    await Task.Run(() => _drawingExporter.ExportToPdf(drawingDoc, outputFolder, context));

    // Extract and export items
    var items = _bomExtractor.ExtractFromDrawing(drawingDoc, context.ProgressCallback);
    await ExportItemsAsync(items, context, outputFolder, existingTemplates, bomItems);

    return rawBomTable;
}
```

Note: The exact SolidWorks API calls to get BOM table annotations should be preserved from the existing BomExtractor code. Add a call to `RawBomTableReader.Read()` alongside the existing item extraction.

- [ ] **Step 5: Remove all API client calls**

Search for any remaining references to `_apiClient`, `FabWorksApiClient`, `IFabWorksApiClient`, `ApiExportDetail`, etc. and remove them. The service should have zero references to the API.

- [ ] **Step 6: Commit**

```bash
git add -A
git commit -m "refactor: rewrite DxfExportService for local file export with revision tracking"
```

---

### Task 10: Update MainForm UI

**Files:**
- Modify: `ExportDXF/Forms/MainForm.cs`
- Modify: `ExportDXF/Forms/MainForm.Designer.cs`

- [ ] **Step 1: Update MainForm.Designer.cs**

Replace the equipment and drawing combo boxes with a single format template text box:

- Remove: `cboEquipment`, `cboDrawing`, `lblEquipment`, `lblDrawing`, `txtTitle`
- Add: `txtFilenameTemplate` (TextBox), `lblFilenameTemplate` (Label with text "Filename Template:")
- Position the template textbox where the dropdowns were, spanning the full width

- [ ] **Step 2: Update MainForm.cs constructor and fields**

Replace `IFabWorksApiClient` with `IDrawingInfoExtractor[]` (list of extractors):

```csharp
private readonly IDxfExportService _exportService;
private readonly IDrawingInfoExtractor[] _extractors;

public MainForm(
    SolidWorksService solidWorksService,
    IDxfExportService exportService,
    IDrawingInfoExtractor[] extractors)
{
    InitializeComponent();
    _solidWorksService = solidWorksService;
    _exportService = exportService;
    _extractors = extractors;
    // ... rest of init
}
```

- [ ] **Step 3: Update document change handler for auto-fill**

When the active SolidWorks document changes, try each extractor:

```csharp
private void OnActiveDocumentChanged(SolidWorksDocument doc)
{
    if (doc == null)
    {
        txtFilenameTemplate.Text = "";
        return;
    }

    foreach (var extractor in _extractors)
    {
        if (extractor.TryExtract(doc.Title, out var info))
        {
            txtFilenameTemplate.Text = info.DefaultTemplate;
            return;
        }
    }
}
```

- [ ] **Step 4: Update StartExportAsync**

Build the ExportContext with template and output folder:

```csharp
private async Task StartExportAsync()
{
    var template = txtFilenameTemplate.Text.Trim();
    if (!FilenameTemplateParser.Validate(template, out var error))
    {
        MessageBox.Show(error, "Invalid Template", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        return;
    }

    var doc = _solidWorksService.GetActiveDocument();
    var sourceDir = Path.GetDirectoryName(doc.FilePath);
    var outputFolder = Path.Combine(sourceDir, "Templates");

    var context = new ExportContext
    {
        ActiveDocument = doc,
        ViewFlipDecider = GetSelectedViewFlipDecider(),
        FilenameTemplate = template,
        OutputFolder = outputFolder,
        CancellationToken = _cancellationTokenSource.Token,
        ProgressCallback = (msg, level, file) => AddLogEvent(msg, level, file),
        BomItemCallback = bomItem => AddBomItem(bomItem),
        SolidWorksApp = _solidWorksService.Application
    };

    await _exportService.ExportAsync(context);
}
```

- [ ] **Step 5: Remove all API/DB dropdown loading methods**

Delete `LoadDrawingDropdownsAsync`, `UpdateDrawingDropdownAsync`, `LookupDrawingInfoFromHistoryAsync`, and any equipment/drawing dropdown event handlers.

- [ ] **Step 6: Commit**

```bash
git add -A
git commit -m "refactor: replace equipment/drawing dropdowns with filename template textbox"
```

---

### Task 11: Update Program.cs and wire dependencies

**Files:**
- Modify: `ExportDXF/Program.cs`

- [ ] **Step 1: Rewrite ServiceContainer**

```csharp
public class ServiceContainer
{
    public MainForm ResolveMainForm()
    {
        var swService = new SolidWorksService();
        var bomExtractor = new BomExtractor();
        var partExporter = new PartExporter();
        var drawingExporter = new DrawingExporter();
        var excelExportService = new ExcelExportService();
        var logFileService = new LogFileService();

        var exportService = new DxfExportService(
            partExporter,
            drawingExporter,
            bomExtractor,
            excelExportService,
            logFileService);

        var extractors = new IDrawingInfoExtractor[]
        {
            new EquipmentDrawingInfoExtractor(),
            new DefaultDrawingInfoExtractor()
        };

        return new MainForm(swService, exportService, extractors);
    }
}
```

- [ ] **Step 2: Remove HttpClient, API client, DB context factory creation**

Delete all code related to `HttpClient`, `FabWorksApiClient`, `FabWorksApiUrl`, `ExportDxfDbContext`, and `FileExportService` from Program.cs.

- [ ] **Step 3: Commit**

```bash
git add ExportDXF/Program.cs
git commit -m "refactor: rewire Program.cs DI — remove API/DB, add Excel and log services"
```

---

### Task 12: Build, fix compile errors, and verify

**Files:**
- Various (fixing any remaining compile errors)

- [ ] **Step 1: Build the solution**

```bash
cd ExportDXF
dotnet build ExportDXF.csproj
```

- [ ] **Step 2: Fix any remaining compile errors**

Common issues to watch for:
- Remaining references to `IFabWorksApiClient`, `FabWorksApiClient`, `ApiExportDetail`, etc.
- Remaining references to `ExportRecord`, `ExportDxfDbContext`
- Missing `using` statements for new namespaces
- `ExportContext` property name changes (`FilePrefix` → `FilenameTemplate`, `Equipment` → removed, etc.)
- `DxfExportService` method signature changes

- [ ] **Step 3: Clean build verification**

```bash
dotnet build ExportDXF.csproj --no-incremental
```

Expected: `Build succeeded. 0 Warning(s) 0 Error(s)`

- [ ] **Step 4: Final commit**

```bash
git add -A
git commit -m "fix: resolve remaining compile errors after API removal"
```

- [ ] **Step 5: Push to remote**

```bash
git push origin master
```
