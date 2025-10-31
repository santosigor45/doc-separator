# Doc Separator (.NET Edition)

High-performance PDF region extractor that produces page-by-page spreadsheets without requiring Python. The new implementation targets Windows-first deployment with native .NET executables, parallel processing, and Excel-friendly output.

## Key Capabilities
- Loads all configured regions (JSON or legacy `pdfregion.txt`) and maps each to a spreadsheet column.
- Parses single PDFs or entire directory trees, extracting every configured region for every page.
- Parallel document and page processing with minimized PDF I/O by reusing cached byte buffers.
- Writes consolidated results to `.xlsx` (ClosedXML) or `.csv` with UTF-8 BOM; column headers mirror region names.
- Emits progress, throughput, and error details through structured console logging.
- Ships without Python; publish as a self-contained Windows CLI or integrate into a desktop host.

## Prerequisites
- .NET 8 SDK (https://dotnet.microsoft.com/download).
- NuGet restore permissions to fetch:
  - `UglyToad.PdfPig` (precise region-aware PDF parsing, no OCR required).
  - `ClosedXML` (fast Excel writer) and `Microsoft.Extensions.Logging.Console`.
- Optional: Excel for validating generated spreadsheets.

## Building on Windows
```powershell
cd doc-separator\DocSeparator.Cli
dotnet restore
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true
```
The published executable lives under `DocSeparator.Cli\bin\Release\net8.0\win-x64\publish\DocSeparator.Cli.exe` and runs without bundling the .NET runtime.

## Running the CLI
```powershell
DocSeparator.Cli.exe --input "C:\Docs\source.pdf" --regions "config\regions.json" --output "C:\Export\pages.xlsx" --format xlsx --parallelism 8
```
Arguments (all `--key value` pairs):
- `--input` (required): PDF file or directory.
- `--regions`: region config path (`config/regions.json` by default, falls back to `pdfregion.txt` automatically).
- `--output`: destination file (default: `extracted.xlsx` or `.csv` under the working directory).
- `--format`: `xlsx` (default) or `csv`.
- `--parallelism`: max concurrent workers; omit to let the tool use CPU count or config override.
- `--help`: print usage.

## Region Configuration
- Primary format: `config/regions.json`
```json
{
  "parallelism": 0,
  "regions": [
    {
      "name": "Header",
      "pageScope": "1,3-5",
      "rectangle": { "left": 84, "top": 46, "right": 390, "bottom": 58 }
    },
    {
      "name": "Totals",
      "pageScope": "even",
      "rectangle": { "left": 420, "top": 680, "right": 550, "bottom": 720 }
    }
  ]
}
```
  - Coordinates keep the legacy top-left origin (pixels/points from the earlier Tk GUI).
  - `pageScope` accepts `all`, `even`, `odd`, or comma-separated ranges (e.g., `1,3,8-12`).
  - Optional `parallelism` lets you pin the default worker count (0 defers to CPU-based automatic choice).
- Legacy support: `pdfregion.txt`
  - Either `RegionName: left,top,right,bottom` per line, or the historical single line `left,top,right,bottom`.
  - Migrating to JSON unlocks page filters and explicit naming.

## Output Structure
- Column order: `Document`, `Page`, `<Region 1>`, `<Region 2>`, ...
- One row per processed page (including blank cells for filtered-out regions or missing text).
- `.xlsx` writer auto sizes columns; `.csv` writer emits UTF-8 BOM for Excel compatibility.

## Observability & Resilience
- Console logging announces active configuration, worker degree, document throughput, and writes an error summary before exiting with code `1`.
- Failures while reading a document do not corrupt others; the process stops only after reporting the fault.
- No OCR dependency: PdfPig operates on the existing text layer; upstream OCR should be performed once at source if required.

## Validation Checklist
1. Run the CLI on representative workloads and confirm runtime drops at least 50% compared to the Python GUI (parallelism engaged).
2. Inspect the spreadsheet in Excel: verify a column exists for each configured region and row counts match processed pages.
3. Spot-check cell contents against original PDF regions.
4. Exercise error paths (bad config, unreadable PDF) to ensure descriptive logging and non-zero exit codes.
5. Publish with `--self-contained true` and validate execution on a clean Windows VM with no Python installed.

## Legacy Python Code
The original `app.py` and supporting files remain for reference but are no longer required to run the extractor. New development should build against the .NET CLI; the Python GUI can be archived or removed once migration is complete.
