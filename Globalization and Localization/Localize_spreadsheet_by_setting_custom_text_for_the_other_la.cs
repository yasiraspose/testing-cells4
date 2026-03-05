using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class LocalizePieChartOtherLabel
{
    static void Main()
    {
        // Load an existing workbook (XLSX format) that contains a pie chart
        Workbook workbook = new Workbook("input.xlsx");

        // Create chart globalization settings and set custom text for the "Other" label
        SettableChartGlobalizationSettings chartSettings = new SettableChartGlobalizationSettings();
        chartSettings.SetOtherName("Custom Other");

        // Create overall globalization settings and assign the chart settings to it
        SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();
        globalization.ChartSettings = chartSettings;

        // Apply the globalization settings to the workbook
        workbook.Settings.GlobalizationSettings = globalization;

        // (Optional) Verify that the workbook contains a pie chart
        Worksheet sheet = workbook.Worksheets[0];
        if (sheet.Charts.Count > 0 && sheet.Charts[0].Type == ChartType.Pie)
        {
            Console.WriteLine("Pie chart detected. Custom 'Other' label will be applied.");
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}