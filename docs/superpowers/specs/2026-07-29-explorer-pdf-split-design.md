# Explorer PDF Split Design

## Goal

Add Windows Explorer PDF context-menu commands that split one or more selected PDFs by a user-supplied page interval or by page parity.

## Current State

`WindowsIntegrationService` registers one `PDF 통합...` command for `.pdf` files. `App` recognizes `--merge` and opens `MergeWindow`. `PdfMergeService` already writes a selected sequence of pages through qpdf but has no Explorer-oriented split request or collision-safe output folder policy.

## Design

`WindowsIntegrationService` will register a `PDF 분리` cascade menu alongside the existing merge command. Its children invoke the application with either `--split-interval` or `--split-parity` and preserve Windows Explorer multi-select behavior.

`App` will parse a single split mode and PDF file paths. For interval split it opens a small WPF input dialog that accepts only a positive integer, defaulting to `1`. After a valid interval is entered, it processes each selected input independently without opening the main viewer. Parity split starts immediately.

`PdfSplitPlanner` will be a pure C# service responsible for output folder allocation, interval page groups, and parity page groups. It returns a plan whose paths are all inside a newly allocated split output folder next to the source PDF. Folder names are `source_분할`, then `source_분할 (2)`, `source_분할 (3)`, and so on. The interval plan retains a final partial group. The parity plan omits a file for an empty parity.

`PdfMergeService` will execute the planned page groups through `SaveTransformedPagesAsync`, which retains qpdf as the only PDF-writing implementation. Batch execution records an outcome per input and continues after individual input failures. A completion dialog reports created files and failures.

## Interfaces

- `PdfSplitPlanner.CreateIntervalPlan(string inputPath, int pageCount, int interval)` returns a plan of consecutive page groups and creates no files.
- `PdfSplitPlanner.CreateParityPlan(string inputPath, int pageCount)` returns odd and non-empty even page groups and creates no files.
- `PdfMergeService.ExecuteSplitPlanAsync(PdfSplitPlan plan, CancellationToken)` writes every planned group and returns the generated results.
- `App` invokes the planner after resolving the selected input PDFs and presents one aggregate completion message.

## Error Handling

- An interval below `1` or non-numeric dialog input is rejected before any split work starts.
- A missing or unreadable PDF records a failure for that input and does not stop later selected files.
- Output files are never written into an existing split output folder.
- qpdf failures are reported with the source file name and error message.

## Testing

A new `test/PdfMergeTool.Tests` xUnit project links the pure planner source and verifies folder allocation, positive interval validation, partial final batches, odd/even page selection, and omission of an empty parity output. The release workflow runs this test project before packaging.

## Scope

This change is Windows Explorer only. It does not add macOS Finder integration, alter the viewer's existing page-by-page split command, or overwrite source PDFs.
