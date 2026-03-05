using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Apply custom globalization settings for label localization
        workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

        // Define the range that will be subtotaled (e.g., A1:B5)
        CellArea area = CellArea.CreateCellArea(0, 0, 4, 1);

        // Apply subtotal:
        // - group by column 0 (first column)
        // - use Sum function
        // - show subtotal and grand total
        cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 0 }, true, true, true);

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }

    // Custom globalization settings to localize "Grand Total" and "Subtotal" labels
    class CustomGlobalizationSettings : GlobalizationSettings
    {
        // Localize the grand total label for any consolidation function
        public override string GetGrandTotalName(ConsolidationFunction functionType)
        {
            // Example: Chinese localization
            return "合计";
        }

        // Localize the subtotal label for different subtotal types
        public override string GetSubTotalName(PivotFieldSubtotalType subTotalType)
        {
            switch (subTotalType)
            {
                case PivotFieldSubtotalType.Sum:
                    return "小计 (求和)";
                case PivotFieldSubtotalType.Count:
                    return "小计 (计数)";
                case PivotFieldSubtotalType.Average:
                    return "小计 (平均)";
                case PivotFieldSubtotalType.Max:
                    return "小计 (最大值)";
                case PivotFieldSubtotalType.Min:
                    return "小计 (最小值)";
                default:
                    return base.GetSubTotalName(subTotalType);
            }
        }
    }
}