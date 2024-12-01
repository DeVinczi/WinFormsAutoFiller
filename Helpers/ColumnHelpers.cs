using ClosedXML.Excel;

using System.Data;

namespace WinFormsAutoFiller.Helpers
{
    internal static class ColumnHelpers
    {
        public static List<string> FetchColumnData(DataTable table, string targetColumn)
        {
            try
            {
                var results = new List<string>();

                foreach (DataRow row in table.Rows)
                {
                    try
                    {
                        var r = row[targetColumn];
                        // Add the value from the target column to the results list
                        if (row[targetColumn] != DBNull.Value) // Handle nulls if necessary
                        {
                            results.Add(row[targetColumn].ToString());
                        }
                    }
                    catch
                    {
                    }
                }

                return results;
            }
            catch
            {
                throw;
            }
        }

        public static DataTable ReadWorksheetToDataTable(string filePath, string worksheetName)
        {
            try
            {
                DataTable dataTable = new DataTable();

                using (var workbook = new XLWorkbook(filePath))
                {
                    // Access the specific worksheet by name
                    var worksheet = workbook.Worksheet(worksheetName);

                    // Loop through the columns in the first row to add them as DataTable columns
                    bool firstRow = true;
                    foreach (var row in worksheet.RowsUsed())  // Iterate through used rows in the worksheet
                    {
                        if (firstRow)
                        {
                            // Add columns based on the first row (headers)
                            foreach (var cell in row.Cells())
                            {
                                dataTable.Columns.Add(cell.Value.ToString());
                            }
                            firstRow = false;
                            continue;
                        }

                        // Add data rows
                        DataRow dataRow = dataTable.NewRow();
                        int columnIndex = 0;
                        foreach (var cell in row.Cells())
                        {
                            dataRow[columnIndex++] = cell.Value.ToString();  // Add cell value to the DataRow
                        }

                        dataTable.Rows.Add(dataRow);
                    }
                }

                return dataTable;
            }
            catch
            {
                throw;
            }
        }
    }
}
