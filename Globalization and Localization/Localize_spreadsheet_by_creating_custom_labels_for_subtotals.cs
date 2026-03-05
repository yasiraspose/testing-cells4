using System;
using Aspose.Cells;

namespace AsposeCellsLocalizationExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing XLSX workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Create an instance of SettableGlobalizationSettings to customize labels
            SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();

            // Set a custom label for the subtotal of the SUM function
            // This label will be used when the Subtotal method creates a total row
            globalization.SetTotalName(ConsolidationFunction.Sum, "Custom Subtotal");

            // Apply the custom globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = globalization;

            // Define the range on which to apply the subtotal.
            // Example: rows 0-4 (A1:B5) where column 0 (A) is the field to group by.
            CellArea area = CellArea.CreateCellArea(0, 0, 4, 1);

            // Apply subtotal:
            //   - group by column index 0 (first column)
            //   - use SUM as the consolidation function
            //   - include column 0 in the subtotal calculation
            //   - replace existing data (true), do not create page breaks (false), and add a total row (true)
            workbook.Worksheets[0].Cells.Subtotal(area, 0, ConsolidationFunction.Sum,
                new int[] { 0 }, true, false, true);

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}