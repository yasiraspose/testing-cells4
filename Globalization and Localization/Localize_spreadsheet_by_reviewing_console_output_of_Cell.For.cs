using System;
using Aspose.Cells;

namespace AsposeCellsFormulaLocalDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the XLSX file to be examined.
            // You can replace this with any valid path or pass it as a command‑line argument.
            string inputPath = args.Length > 0 ? args[0] : "Sample.xlsx";

            // LoadOptions allow us to control how formulas are handled on load.
            // Setting ParsingFormulaOnOpen to true ensures formulas are parsed immediately,
            // which is required for FormulaLocal to work correctly.
            LoadOptions loadOptions = new LoadOptions
            {
                ParsingFormulaOnOpen = true
            };

            // Load the workbook using the specified options.
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Set the workbook region to a locale different from the default (en‑US).
            // This will affect how FormulaLocal is rendered.
            // Example: German (Germany) – function names like SUM become SUMME.
            workbook.Settings.Region = CountryCode.Germany;

            // Access the first worksheet (you can iterate all worksheets if needed).
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Iterate through all used cells and display both the standard and localized formulas.
            Console.WriteLine("=== Formula Localization Report ===");
            foreach (Cell cell in cells)
            {
                // Skip cells that do not contain a formula.
                if (string.IsNullOrEmpty(cell.Formula))
                    continue;

                // Standard (English) formula.
                string standardFormula = cell.Formula;

                // Locale‑specific formula (German in this example).
                string localizedFormula = cell.FormulaLocal;

                // Alternative way to obtain the localized formula using GetFormula.
                string localizedViaGetFormula = cell.GetFormula(false, true);

                // Output the information.
                Console.WriteLine($"Cell {cell.Name}:");
                Console.WriteLine($"  Standard Formula : {standardFormula}");
                Console.WriteLine($"  Localized Formula: {localizedFormula}");
                Console.WriteLine($"  GetFormula(..., true): {localizedViaGetFormula}");
                Console.WriteLine();
            }

            // Optionally, save the workbook after inspection (e.g., to preserve any changes).
            // Here we simply save a copy with a different name.
            string outputPath = "LocalizedReport.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as '{outputPath}'.");
        }
    }
}