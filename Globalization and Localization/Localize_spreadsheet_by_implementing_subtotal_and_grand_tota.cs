using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

class Program
{
    static void Main(string[] args)
    {
        // Expect two arguments: input XLSX file and output XLSX file.
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: Program.exe <input.xlsx> <output.xlsx>");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];

        // Load the existing workbook (create rule is applied by the constructor).
        Workbook workbook = new Workbook(inputPath);
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // ------------------------------------------------------------
        // 1. Create settable globalization settings for normal cells.
        // ------------------------------------------------------------
        SettableGlobalizationSettings globalSettings = new SettableGlobalizationSettings();

        // Localize the labels used by the Subtotal method.
        // Example: Chinese localization.
        globalSettings.SetTotalName(ConsolidationFunction.Sum, "小计");          // "Subtotal"
        globalSettings.SetGrandTotalName(ConsolidationFunction.Sum, "合计");   // "Grand Total"

        // Assign the globalization settings to the workbook.
        workbook.Settings.GlobalizationSettings = globalSettings;

        // ------------------------------------------------------------
        // 2. Apply Subtotal to a sample range to demonstrate the labels.
        // ------------------------------------------------------------
        // Assume data exists in A1:B5 (adjust as needed).
        CellArea dataArea = CellArea.CreateCellArea(0, 0, 4, 1); // rows 0‑4, cols 0‑1 (A1:B5)
        // Subtotal by the first column (index 0), using SUM on the second column (index 1).
        // The last three booleans: replace existing, page break after each change, summary below.
        cells.Subtotal(dataArea, 0, ConsolidationFunction.Sum, new int[] { 1 }, true, false, true);

        // ------------------------------------------------------------
        // 3. If the worksheet contains a pivot table, localize its labels.
        // ------------------------------------------------------------
        if (sheet.PivotTables.Count > 0)
        {
            // Create settable pivot globalization settings.
            SettablePivotGlobalizationSettings pivotSettings = new SettablePivotGlobalizationSettings();

            // Localize the generic "Total" and "Grand Total" labels in the pivot table.
            pivotSettings.SetTextOfTotal("合计");          // Text for the total row/column.
            pivotSettings.SetTextOfGrandTotal("总计");    // Text for the grand total row/column.

            // Localize specific subtotal types (example for Sum and Count).
            pivotSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Sum, "小计");
            pivotSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Count, "计数小计");

            // Attach the pivot settings to the previously created global settings.
            globalSettings.PivotSettings = pivotSettings;
        }

        // ------------------------------------------------------------
        // 4. Save the modified workbook (save rule is applied here).
        // ------------------------------------------------------------
        workbook.Save(outputPath);
    }
}