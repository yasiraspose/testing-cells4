using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Settings;

namespace AsposeCellsGlobalizationDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // -------------------------------------------------
            // 1. Create and configure SettableGlobalizationSettings
            // -------------------------------------------------
            SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();

            // Customize subtotal label for the SUM function
            globalization.SetTotalName(ConsolidationFunction.Sum, "Custom Sum Total");
            // (Optional) customize other subtotal labels, e.g., AVERAGE
            globalization.SetTotalName(ConsolidationFunction.Average, "Custom Average Total");

            // -------------------------------------------------
            // 2. Create and configure SettableChartGlobalizationSettings
            // -------------------------------------------------
            SettableChartGlobalizationSettings chartGlobals = new SettableChartGlobalizationSettings();

            // Customize chart-related texts
            chartGlobals.SetSeriesName("Custom Series");
            chartGlobals.SetChartTitleName("Custom Pie Chart Title");
            chartGlobals.SetLegendTotalName("Custom Total Legend");
            chartGlobals.SetOtherName("Other Category");
            chartGlobals.SetLegendIncreaseName("Increase");
            chartGlobals.SetLegendDecreaseName("Decrease");

            // Attach the chart globalization settings to the main globalization object
            globalization.ChartSettings = chartGlobals;

            // Apply the globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = globalization;

            // -------------------------------------------------
            // 3. Demonstrate subtotal creation using the customized label
            // -------------------------------------------------
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Assume data exists in A1:B5; create a subtotal on column A (field index 0)
            // The subtotal will use the custom total name defined above.
            CellArea area = CellArea.CreateCellArea(0, 0, 4, 1); // rows 0-4, cols 0-1
            cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 0 }, true, false, true);

            // -------------------------------------------------
            // 4. Locate a pie chart and ensure it reflects the custom globalization
            // -------------------------------------------------
            // For demonstration, create a pie chart if none exists.
            Chart chart = null;
            if (sheet.Charts.Count > 0)
            {
                chart = sheet.Charts[0];
            }
            else
            {
                // Create a simple pie chart using sample data
                int chartIndex = sheet.Charts.Add(ChartType.Pie, 6, 0, 20, 10);
                chart = sheet.Charts[chartIndex];

                // Sample data for the chart
                cells["D1"].PutValue("Category");
                cells["E1"].PutValue("Value");
                cells["D2"].PutValue("A");
                cells["E2"].PutValue(30);
                cells["D3"].PutValue("B");
                cells["E3"].PutValue(45);
                cells["D4"].PutValue("C");
                cells["E4"].PutValue(25);

                // Bind data to the chart
                chart.NSeries.Add("E2:E4", true);
                chart.NSeries.CategoryData = "D2:D4";
            }

            // Set a chart title (the title text itself is not localized, but the title name can be)
            chart.Title.Text = "Demo Pie Chart";

            // -------------------------------------------------
            // 5. Save the modified workbook
            // -------------------------------------------------
            workbook.Save("output.xlsx");
        }
    }
}