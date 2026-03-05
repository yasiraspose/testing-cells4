using System;
using Aspose.Cells;

namespace AsposeCellsFormulaLocalDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (XLSX format)
            // Replace "input.xlsx" with the path to your source file
            Workbook workbook = new Workbook("input.xlsx");

            // Set the workbook locale – for example German (de-DE)
            // This influences how FormulaLocal is interpreted and displayed
            workbook.Settings.Region = CountryCode.Germany;

            // Access the first worksheet and a target cell (A1)
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];

            // -----------------------------------------------------------------
            // 1. Read the current formula in both standard (English) and local form
            // -----------------------------------------------------------------
            Console.WriteLine("Original formulas:");
            Console.WriteLine($"Standard Formula : {cell.Formula}");
            Console.WriteLine($"Localized Formula: {cell.FormulaLocal}");

            // -----------------------------------------------------------------
            // 2. Set a formula using the localized (German) syntax
            //    In German the SUM function is "SUMME"
            // -----------------------------------------------------------------
            cell.FormulaLocal = "=SUMME(B1:C1)";

            // After setting, both properties reflect the same underlying formula
            Console.WriteLine("\nAfter assigning FormulaLocal:");
            Console.WriteLine($"Standard Formula : {cell.Formula}");
            Console.WriteLine($"Localized Formula: {cell.FormulaLocal}");

            // -----------------------------------------------------------------
            // 3. Optionally calculate the workbook to obtain the result
            // -----------------------------------------------------------------
            workbook.CalculateFormula();

            // Display the calculated value of the cell
            Console.WriteLine($"\nCalculated Value in A1: {cell.Value}");

            // -----------------------------------------------------------------
            // 4. Save the modified workbook
            // -----------------------------------------------------------------
            // Replace "output.xlsx" with the desired output path
            workbook.Save("output.xlsx");
            Console.WriteLine("\nWorkbook saved as 'output.xlsx'.");
        }
    }
}