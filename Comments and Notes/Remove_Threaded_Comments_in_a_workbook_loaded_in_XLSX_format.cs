using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class RemoveThreadedCommentsDemo
    {
        public static void Run()
        {
            // Load the existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets in the workbook and clear comments
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                worksheet.ClearComments();
            }

            // Save the workbook after removing the threaded comments
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            RemoveThreadedCommentsDemo.Run();
        }
    }
}