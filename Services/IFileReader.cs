using ExcelDataReader;

using FormFiller.Constants;
using FormFiller.Models;

using OfficeOpenXml;

using System.Data;
using System.Text;
using System.Text.RegularExpressions;

using WinFormsAutoFiller.Infrastructure;
using WinFormsAutoFiller.Utilis;

namespace FormFiller.Services;

public interface IFileReader
{
    Task<DataTable?> ReadExcelFileAsync(string fileName, Dictionary<string, Regex> columnPatterns, string worksheetName);
    ExcelWorkersRow? ReadExcelRow(string tempFilePath, int startIndex = 1);
    Task<List<string>> GetExcelWorksheetNames(string fileName);
    Result<string, Error> FindWordInWordDocument(string fileName);
}

public class FileReader : IFileReader
{
    public async Task<DataTable?> ReadExcelFileAsync(string fileName, Dictionary<string, Regex> columnPatterns, string worksheetName)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        var encoding = Encoding.GetEncoding("UTF-16");

        await using var stream = File.OpenRead(fileName);
        using var reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration()
        {
            FallbackEncoding = encoding
        });
        var result = reader.AsDataSet(new ExcelDataSetConfiguration
        {
            ConfigureDataTable = (_) => new ExcelDataTableConfiguration
            {
                UseHeaderRow = true
            }
        });

        if (result.Tables.Count == 0) return null;

        var index = result.Tables.IndexOf(worksheetName);
        var tableIndex = index == -1
            ? 0
            : index;

        var sourceTable = result.Tables[tableIndex];
        var newTable = new DataTable();

        if (columnPatterns is null)
        {
            newTable = sourceTable.Copy();
            return newTable;
        }

        foreach (var pattern in columnPatterns)
        {
            var matchingColumn = FindMatchingColumn(sourceTable, pattern.Value);
            if (matchingColumn != null)
            {
                newTable.Columns.Add(pattern.Key, matchingColumn.DataType);
            }
        }

        foreach (DataRow sourceRow in sourceTable.Rows)
        {
            var newRow = newTable.NewRow();

            foreach (var pattern in columnPatterns)
            {
                // Find matching columns for the current pattern
                var matchingColumns = FindMatchingColumns(sourceTable, pattern.Value);

                if (matchingColumns.Count > 0)
                {
                    foreach (var column in matchingColumns)
                    {
                        var columnName = column.ColumnName;

                        // Check if the newTable has the column and sourceRow is not null for it
                        if (newTable.Columns.Contains(pattern.Key) && !sourceRow.IsNull(columnName))
                        {
                            var newValue = sourceRow[columnName];
                            var existingValue = newRow[pattern.Key];

                            if (columnPatterns.ContainsKey(WorkersFormKeys.KwotaOtrzymanegoDofinansowania) &&
                                pattern.Value == columnPatterns[WorkersFormKeys.KwotaOtrzymanegoDofinansowania])
                            {
                                // Special handling for WorkersFormKeys.KwotaOtrzymanegoDofinansowania (numeric merging)
                                if (double.TryParse(newValue.ToString(), out var numericValue))
                                {
                                    if (double.TryParse(existingValue?.ToString(), out var existingNumericValue))
                                    {
                                        newRow[pattern.Key] = existingNumericValue + numericValue;
                                    }
                                    else
                                    {
                                        newRow[pattern.Key] = numericValue;
                                    }
                                }
                                else
                                {
                                    if (string.IsNullOrEmpty(existingValue?.ToString()))
                                    {
                                        newRow[pattern.Key] = newValue;
                                    }
                                }
                            }
                            else
                            {
                                if (string.IsNullOrEmpty(existingValue?.ToString()))
                                {
                                    newRow[pattern.Key] = newValue;
                                }
                            }
                        }
                    }
                }
            }

            newTable.Rows.Add(newRow);
        }

        return newTable;
    }

    public ExcelWorkersRow? ReadExcelRow(string tempFilePath, int startIndex = 1)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using var package = new ExcelPackage(new FileInfo(tempFilePath));

        var worksheet = package.Workbook.Worksheets[0];
        var rowCount = worksheet.Dimension.Rows;
        var colCount = worksheet.Dimension.Columns;
        if (rowCount < 1)
        {
            return null;
        }

        var headers = new List<string>();
        for (int col = startIndex; col <= colCount; col++)
        {
            headers.Add(worksheet.Cells[startIndex, col].Text);
        }

        var excelRow = new ExcelWorkersRow();

        for (int col = startIndex; col <= colCount; col++)
        {
            var header = headers[col - startIndex];
            var cellValue = worksheet.Cells[rowCount, col].Text;
            excelRow.Data[header] = cellValue;
        }

        return excelRow;
    }

    public async Task<List<string>> GetExcelWorksheetNames(string fileName)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        var encoding = Encoding.GetEncoding("UTF-16");

        await using var stream = File.OpenRead(fileName);
        using var reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration()
        {
            FallbackEncoding = encoding
        });
        var result = reader.AsDataSet(new ExcelDataSetConfiguration
        {
            ConfigureDataTable = (_) => new ExcelDataTableConfiguration
            {
                UseHeaderRow = true
            }
        });

        var list = new List<string>();
        if (result.Tables.Count == 0) return null;
        foreach (DataTable resultTable in result.Tables)
        {
            list.Add(resultTable.TableName);
        }

        return list;
    }

    private static DataColumn? FindMatchingColumn(DataTable table, Regex regex)
    {
        return table.Columns.Cast<DataColumn>().FirstOrDefault(col => regex.IsMatch(col.ColumnName));
    }

    private static List<DataColumn> FindMatchingColumns(DataTable table, Regex regex)
    {
        return table.Columns.Cast<DataColumn>().Where(col => regex.IsMatch(col.ColumnName)).ToList();
    }

    static void ProcessTrainingData(DataTable dataTable)
    {
        var data = new List<TrainingData>();

        // Iterate through the rows of the DataTable
        foreach (DataRow row in dataTable.Rows)
        {
            var trainingData = new TrainingData();
            trainingData.Name = row[0].ToString();  // Column 1 is the name

            // Iterate through the remaining columns for training data
            for (int col = 1; col < dataTable.Columns.Count; col++)
            {
                if (int.TryParse(row[col].ToString(), out int trainingCount))
                {
                    trainingData.Trainings.Add(trainingCount);
                }
            }

            data.Add(trainingData);
        }

        // Display the results
        foreach (var entry in data)
        {
            Console.WriteLine($"Name: {entry.Name}");
            for (int i = 0; i < entry.Trainings.Count; i++)
            {
                Console.WriteLine($"  Training {i + 1}: {entry.Trainings[i]}");
            }
        }
    }

    public Result<string, Error> FindWordInWordDocument(string filePath)
    {
        dynamic wordApp = Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application"));

        var doc = wordApp.Documents.Open(filePath);

        var text = doc.Content.Text;

        var match = RegexPatterns.FindOutHoursRegex().Match(text);

        doc.Close();
        wordApp.Quit();

        if (match.Success)
        {
            return match.Groups[1].Value;
        }
        else
        {
            return Errors.WorkHoursAreEmpty;
        }
    }
}