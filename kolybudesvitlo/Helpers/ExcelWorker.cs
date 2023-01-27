using Aspose.Cells;

namespace kolybudesvitlo.Helpers
{
    /// <summary>
    /// Class for work with excel table
    /// </summary>
    public static class ExcelWorker
    {
        public static string? FindStreetInExcelTable(string street)
        {
            try
            {
                // upload Excel file
                Workbook wb = new Workbook("Grafik_Rivne.xlsx");

                // get all lists
                WorksheetCollection collection = wb.Worksheets;

                // get a list of jobs using its index
                Worksheet worksheet = collection[0];

                // get number row and column         
                int rows = worksheet.Cells.MaxDataRow;
                int cols = worksheet.Cells.MaxDataColumn;

                return(FindStreet(rows, cols, worksheet, street));
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// walk through the rows in search of the street
        /// </summary>
        /// <param name="rows"></param>
        /// <param name="cols"></param>
        /// <param name="worksheet"></param>
        /// <param name="street"></param>
        /// <returns>null if no result</returns>
        private static string? FindStreet(int rows, int cols, Worksheet worksheet, string street)
        {
            for (int row = 0; row < rows; row++)
            {
                var streetRow = worksheet.Cells[row, 0].Value;
                if (streetRow == null)
                {
                    continue;
                }
                else if (streetRow.ToString().Contains(street))
                {
                    string timeBlackout = "";
                    for (int col = 1; col <= cols; col++)
                    {
                        var timeCol = worksheet.Cells[row, col].Value;
                        timeBlackout += timeCol == null ? "" : timeCol.ToString() + " ";
                    }
                    return timeBlackout;
                }
            }
            return null;
        }
    }
}
