import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const read = (...parts) => readFileSync(join(root, ...parts), "utf8");

const parser = read("src", "PdfMergeTool", "Services", "PdfExplorerCommandParser.cs");
const integration = read("src", "PdfMergeTool", "Services", "WindowsIntegrationService.cs");
const app = read("src", "PdfMergeTool", "App.xaml.cs");
const xaml = read("src", "PdfMergeTool", "MainWindow.xaml");
const mainWindow = read("src", "PdfMergeTool", "MainWindow.xaml.cs");

assert.match(parser, /--split-pages/);
assert.match(parser, /PageSplit/);
assert.match(integration, /DeleteValue\(string\.Empty, throwOnMissingValue: false\)/);
assert.match(integration, /SetValue\("MUIVerb", "PDF 분리"\)/);
assert.match(integration, /CreateSubKey\("ExtendedSubCommandsKey"\)/);
assert.match(integration, /CreateSubKey\("shell"\)/);
assert.match(integration, /RegisterSplitCommand\(shellKey, "SplitByPage", "페이지별 분할", "--split-pages"/);
assert.match(integration, /RegisterSplitCommand\(shellKey, "SplitByInterval", "N페이지마다 분리", "--split-interval"/);
assert.match(integration, /RegisterSplitCommand\(shellKey, "SplitByParity", "홀수\/짝수 분리", "--split-parity"/);
assert.match(integration, /RefreshPdfContextMenuRegistration/);
assert.match(app, /RefreshPdfContextMenuRegistration\(\);/);
assert.match(app, /PdfExplorerCommand\.PageSplit => PdfSplitMode\.PageByPage/);
assert.match(xaml, /<MenuItem Header="PDF 분할">[\s\S]*?OnSplitByIntervalClick[\s\S]*?OnSplitByParityClick/);
assert.match(mainWindow, /PdfSplitPlanner\.CreateIntervalPlan\(sourcePath, pageTransforms, interval, referencePath\)/);
assert.match(mainWindow, /PdfSplitPlanner\.CreateParityPlan\(sourcePath, pageTransforms, referencePath\)/);

console.log("pdf split integration contract passed.");
