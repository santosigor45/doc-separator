# Architecture Overview

## Technology Selection
- Language/runtime: C# on .NET 8 (self-contained publishable for Windows x64).
- Libraries: UglyToad.PdfPig for low-level, allocation-friendly PDF parsing; ClosedXML for XLSX export; System.Text.Json for config parsing; built-in CSV writer for lightweight export.
- Rationale: Native .NET executables remove Python dependency chain, PdfPig gives direct region extraction without OCR overhead, and ClosedXML provides fast Excel output while remaining redistributable.

## Execution Flow
1. Bootstrap via CLI options (`--input`, `--regions`, `--output`, `--format`, `--parallelism`).
2. Load region configuration:
   - Preferred source: `config/regions.json`.
   - Backwards compatibility: fallback to legacy `pdfregion.txt` single-region format.
3. Enumerate PDF sources (single file or directory traversal).
4. For each PDF document:
   - Cache the raw bytes once (`File.ReadAllBytes`).
   - Dispatch extraction using `Parallel.ForEach` across page ranges; each worker opens a PdfPig document over an in-memory stream slice.
   - For every configured region, clip coordinates and extract text via PdfPig (`GetWords()` + bounding-box filtering).
   - Accumulate per-page results in-memory (`List<ExtractionRow>`).
5. After document processing, append rows to a shared `ExtractionTable` (includes document name, page number, and one column per region).
6. Persist aggregated table:
   - XLSX: single worksheet per batch using ClosedXML without keeping workbook open unnecessarily.
   - CSV: stream rows using buffered writer with invariant culture.

## Configuration Model
```json
{
  "regions": [
    {
      "name": "EmployeeName",
      "pageScope": "all",
      "rectangle": { "left": 84, "top": 46, "right": 390, "bottom": 58 }
    }
  ],
  "parallelism": 0
}
```
- `pageScope` accepts `all`, `even`, `odd`, or explicit list (e.g., `"1,3,5-7"`); the parser normalizes to predicate functions.
- Coordinates follow the legacy GUI logic (origin at top-left in PDF pixel space); runtime converts them to PdfPig bottom-left coordinates per page height.
- Optional `parallelism` overrides default (`Environment.ProcessorCount - 1` minimum 1).

## Performance Considerations
- PdfPig processing stays in-memory; workers reuse cached bytes so the source PDF is read from disk only once.
- Region extraction uses cached predicates so each page loops through configured regions once.
- Batched parallel processing reduces contention with `Partitioner.Create`.
- Spreadsheet writing happens after extraction to minimize cross-thread contention; CSV writing streams to reduce memory when requested.
- Logging uses `Microsoft.Extensions.Logging` simple console format with timestamps and throughput metrics.

## Error Handling & Telemetry
- Validate CLI arguments and configuration schema before work begins; descriptive errors with remediation tips.
- Safeguard each document inside try/catch, skip on failure but continue others, and emit summary with exit code 1 if any errors occurred.
- Structured logging to console with descriptive error messages; unreadable PDFs or invalid regions raise exceptions that bubble to `Program.Main`.
- Unit-testable components: config loader, coordinate converter, extraction pipeline, writers.
