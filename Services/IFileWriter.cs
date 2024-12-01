using OfficeOpenXml;

using System.Data;

using static System.String;

namespace FormFiller.Services;

public interface IFileWriter
{
    Task<(string, ExcelWorksheet)> ExcelWriterAsync(DataTable dataTable, string filePath);
    Task<(string, ExcelWorksheet)> ExcelWriterAsyncReversed(DataTable dataTable, string filePath);
    Task<bool> DeleteExcelRowAsync(string tempFilePath, int row);
}

public class FileWriter : IFileWriter
{
    public async Task<(string, ExcelWorksheet)> ExcelWriterAsync(DataTable dataTable, string filePath)
    {
        if (dataTable is null)
        {
            return (Empty, null)!;
        }

        var tempFilePath = Path.ChangeExtension(Path.Combine(Path.GetTempPath(), Path.GetTempFileName()), ".xlsx");
        try
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add(filePath[(filePath.LastIndexOf('\\') + 1)..]);
            for (int i = 1; i <= dataTable.Columns.Count; i++)
            {
                worksheet.Cells[1, i].Value = dataTable.Columns[i - 1].ColumnName;
                worksheet.Cells[1, i].Style.Font.Bold = true;
            }

            for (int row = 0; row < dataTable.Rows.Count; row++)
            {
                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    worksheet.Cells[row + 2, col + 1].Value = dataTable.Rows[row][col];
                }
            }

            worksheet.Cells.AutoFitColumns();
            await package.SaveAsAsync(new FileInfo(tempFilePath));

            return (tempFilePath, worksheet);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
            throw;
        }
    }

    public async Task<(string, ExcelWorksheet)> ExcelWriterAsyncReversed(DataTable dataTable, string filePath)
    {
        if (dataTable is null)
        {
            return (string.Empty, null)!;
        }

        var tempFilePath = Path.ChangeExtension(Path.Combine(Path.GetTempPath(), Path.GetTempFileName()), ".xlsx");
        try
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add(filePath[(filePath.LastIndexOf('\\') + 1)..]);

            // Write column headers
            for (int i = 1; i <= dataTable.Columns.Count; i++)
            {
                worksheet.Cells[1, i].Value = dataTable.Columns[i - 1].ColumnName;
                worksheet.Cells[1, i].Style.Font.Bold = true;
            }

            // Write rows in reverse order
            for (int row = 0; row < dataTable.Rows.Count; row++)
            {
                var reversedRow = dataTable.Rows[dataTable.Rows.Count - 1 - row];
                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    worksheet.Cells[row + 2, col + 1].Value = reversedRow[col];
                }
            }

            worksheet.Cells.AutoFitColumns();
            var x = worksheet.Rows;
            await package.SaveAsAsync(new FileInfo(tempFilePath));

            return (tempFilePath, worksheet);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
            throw;
        }
    }

    public async Task<bool> DeleteExcelRowAsync(string tempFilePath, int row)
    {
        try
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using var package = new ExcelPackage(new FileInfo(tempFilePath));

            var worksheet = package.Workbook.Worksheets[0];

            worksheet.DeleteRow(row);

            await package.SaveAsync();
            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
            throw;
        }
    }


}