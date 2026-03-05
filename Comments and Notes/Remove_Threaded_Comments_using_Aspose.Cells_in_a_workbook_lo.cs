using System;
using Aspose.Cells;

class RemoveThreadedComments
{
    static void Main()
    {
        // Load the workbook (XLSX format)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet worksheet in workbook.Worksheets)
        {
            // Determine the used range of the worksheet
            int maxRow = worksheet.Cells.MaxDataRow;
            int maxCol = worksheet.Cells.MaxDataColumn;

            // Scan every cell within the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
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

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}