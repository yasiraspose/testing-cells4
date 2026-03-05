using System;
using Aspose.Cells;

namespace AsposeCellsLocalizationDemo
{
    // Custom globalization settings for Russian language
    public class RussianGlobalizationSettings : GlobalizationSettings
    {
        // Localize Boolean values: TRUE -> ИСТИНА, FALSE -> ЛОЖЬ
        public override string GetBooleanValueString(bool value)
        {
            return value ? "ИСТИНА" : "ЛОЖЬ";
        }

        // Localize common Excel error strings
        public override string GetErrorValueString(string err)
        {
            switch (err)
            {
                case "#NAME?":   return "#ИМЯ?";
                case "#DIV/0!":  return "#ДЕЛ/0!";
                case "#REF!":    return "#ССЫЛКА!";
                case "#VALUE!":  return "#ЗНАЧ!";
                case "#N/A":     return "#Н/Д";
                case "#NUM!":    return "#ЧИСЛО!";
                case "#NULL!":   return "#ПУСТО!";
                default:         return base.GetErrorValueString(err);
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file (replace with actual path)
            string inputPath = "input.xlsx";

            // Load the workbook (XLSX format)
            Workbook workbook = new Workbook(inputPath);

            // Apply the custom Russian globalization settings
            workbook.Settings.GlobalizationSettings = new RussianGlobalizationSettings();

            // Example: display localized values of the first row
            // (Assumes the first row contains Boolean and error values)
            for (int col = 0; col < 9; col++)
            {
                Cell cell = workbook.Worksheets[0].Cells[0, col];
                Console.WriteLine($"Cell[0,{col}]: {cell.StringValue}");
            }

            // Save the localized workbook
            string outputPath = "output_localized.xlsx";
            workbook.Save(outputPath);
        }
    }
}