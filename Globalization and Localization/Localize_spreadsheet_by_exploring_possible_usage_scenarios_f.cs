using System;
using Aspose.Cells;

namespace AsposeCellsLocalizationDemo
{
    // Custom globalization settings to localize boolean and error values
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        // Localize boolean values (e.g., Russian)
        public override string GetBooleanValueString(bool bv)
        {
            return bv ? "ИСТИНА" : "ЛОЖЬ";
        }

        // Localize common Excel error strings (e.g., Russian equivalents)
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
            // Path to the source XLSX file (must exist)
            string inputPath = "source.xlsx";

            // Load the workbook (lifecycle: load)
            Workbook wb = new Workbook(inputPath);

            // Apply custom globalization settings (error & boolean localization)
            wb.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // Access the first worksheet and its cells
            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;

            // Insert boolean values
            cells[0, 0].PutValue(true);   // A1
            cells[0, 1].PutValue(false);  // B1

            // Insert various error strings (as raw text)
            string[] errors = new string[]
            {
                "#NAME?", "#DIV/0!", "#REF!", "#VALUE!", "#N/A", "#NUM!", "#NULL!"
            };
            for (int i = 0; i < errors.Length; i++)
            {
                // C1 onward
                cells[0, i + 2].PutValue(errors[i]);
            }

            // Recalculate formulas (not strictly needed here but ensures any formulas are evaluated)
            wb.CalculateFormula();

            // Display localized results in the console
            Console.WriteLine("Localized cell values:");
            for (int col = 0; col < 9; col++)
            {
                Console.WriteLine($"Cell[0,{col}] ({cells[0, col].Name}): {cells[0, col].StringValue}");
            }

            // Save the modified workbook (lifecycle: save)
            string outputPath = "localized_output.xlsx";
            wb.Save(outputPath);

            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}