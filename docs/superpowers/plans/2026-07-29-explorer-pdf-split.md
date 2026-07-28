# Explorer PDF Split Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use `superpowers:executing-plans` to implement this plan task by task.

**Goal:** Add collision-safe PDF interval and parity split commands to the Windows Explorer context menu.

**Architecture:** A pure `PdfSplitPlanner` owns page grouping and output folder allocation. `PdfMergeService` writes the planner's outputs through qpdf. `App` maps Explorer arguments to an interval dialog or immediate parity batch and shows one aggregate result.

**Tech Stack:** .NET 8, WPF, Windows Registry, qpdf, xUnit.

## Global Constraints

- Windows Explorer `.pdf` association only.
- Never overwrite a source PDF or existing split output folder.
- Interval is a positive integer and defaults to `1`.
- Each selected source file processes independently after failures.

### Task 1: Testable Split Planning

**Files:** Create `src/PdfMergeTool/Services/PdfSplitPlanner.cs`, `test/PdfMergeTool.Tests/PdfMergeTool.Tests.csproj`, and `test/PdfMergeTool.Tests/PdfSplitPlannerTests.cs`; modify `PdfMergeTool.sln`.

**Produces:** `PdfSplitPlanner.CreateIntervalPlan(string inputPath, int pageCount, int interval)`, `PdfSplitPlanner.CreateParityPlan(string inputPath, int pageCount)`, `PdfSplitPlan`, `PdfSplitOutput`, and `PdfSplitKind`.

- [ ] Write failing planner tests.

```csharp
[Fact]
public void CreateIntervalPlan_keeps_final_partial_batch()
{
    var plan = PdfSplitPlanner.CreateIntervalPlan(@"C:\\docs\\source.pdf", 10, 3);
    Assert.Equal(["1-3", "4-6", "7-9", "10"], plan.Outputs.Select(x => x.PageRange));
}

[Fact]
public void CreateParityPlan_omits_empty_even_output()
{
    var plan = PdfSplitPlanner.CreateParityPlan(@"C:\\docs\\source.pdf", 1);
    Assert.Single(plan.Outputs);
    Assert.Equal(PdfSplitKind.Odd, plan.Outputs[0].Kind);
}
```

- [ ] Run `dotnet test test/PdfMergeTool.Tests/PdfMergeTool.Tests.csproj --filter FullyQualifiedName~PdfSplitPlannerTests`; it must fail because the planner does not exist.
- [ ] Implement the smallest planner that validates `interval >= 1`, makes interval ranges, makes odd/even ranges, and allocates `source_분할`, `source_분할 (2)`, and later folders without creating one.
- [ ] Re-run the focused tests; they must pass.

### Task 2: Execute Plans and Register Explorer Verbs

**Files:** Modify `src/PdfMergeTool/Services/PdfMergeService.cs` and `src/PdfMergeTool/Services/WindowsIntegrationService.cs`; extend `test/PdfMergeTool.Tests/PdfSplitPlannerTests.cs`.

**Consumes:** `PdfSplitPlan.Outputs` with output paths and page transforms.

**Produces:** `PdfMergeService.ExecuteSplitPlanAsync(PdfSplitPlan, CancellationToken)`, and Explorer verbs `--split-interval` and `--split-parity` below a `PDF 분리` cascade.

- [ ] Write a failing test that creates `source_분할` and asserts the next plan uses `source_분할 (2)`.
- [ ] Run that test and verify the missing collision-safe allocation failure.
- [ ] Add plan execution through `SaveTransformedPagesAsync`; preserve qpdf as the only PDF writer.
- [ ] Register `N페이지마다 분리` and `홀수/짝수 분리` command children for every existing `.pdf` shell path.
- [ ] Run all planner tests and verify pass.

### Task 3: Route Explorer Requests Through WPF

**Files:** Create `src/PdfMergeTool/SplitIntervalWindow.xaml` and `src/PdfMergeTool/SplitIntervalWindow.xaml.cs`; modify `src/PdfMergeTool/App.xaml.cs`, `src/PdfMergeTool/App.xaml`, and planner tests.

**Produces:** interval input validation and aggregate success/failure reporting for each selected source PDF.

- [ ] Write a failing parser test for `--split-interval` and `--split-parity`.
- [ ] Run it and verify failure because the parser is missing.
- [ ] Implement the dialog with a default interval of `1`; reject non-integers and values below `1` before processing.
- [ ] For each selected PDF, read its page count, create a split plan, execute it, collect outputs, and continue after exceptions.
- [ ] Run all tests plus `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\\scripts\\build.ps1` on Windows; both must pass.

### Task 4: Release Gate and User Documentation

**Files:** Modify `.github/workflows/release.yml`, `scripts/verify-stability.ps1`, and `README.md`.

- [ ] Add a failing `verify-stability.ps1` assertion that requires `dotnet test test/PdfMergeTool.Tests/PdfMergeTool.Tests.csproj` before installer packaging.
- [ ] Run the stability script and verify it fails before the workflow step is added.
- [ ] Add the release workflow test step and document both Explorer commands and output folder behavior in the README.
- [ ] Run `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\\scripts\\verify-stability.ps1` on Windows and verify it reports `stability checks passed.`

## Self-Review

- All agreed behavior is covered: interval validation, partial batch preservation, parity omission, collision-safe folders, multi-file continuation, and completion reporting.
- No macOS Finder feature, source overwrite, or viewer split rewrite is included.
