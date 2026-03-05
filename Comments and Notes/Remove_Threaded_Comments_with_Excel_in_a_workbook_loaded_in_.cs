using System;
using Aspose.Cells;

namespace RemoveThreadedCommentsDemo
{
    class Program
    {
        static void Main()
        {
            // Load the existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Determine the used range of the worksheet
                Cells cells = worksheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxColumn = cells.MaxDataColumn;

                // Scan each cell within the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        // Retrieve threaded comments for the current cell
                        ThreadedCommentCollection threadedComments = worksheet.Comments.GetThreadedComments(row, col);

                        // If there are any threaded comments, clear them
                        if (threadedComments != null && threadedComments.Count > 0)
                        {
                            threadedComments.Clear();
                        }
                    }
                }
            }

            // Save the workbook after removing all threaded comments
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}