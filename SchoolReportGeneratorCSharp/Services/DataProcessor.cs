using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using ClosedXML.Excel;

namespace SchoolReportGenerator.Services;

/// <summary>
/// Processes Excel data for report card generation.
/// Mirrors Python's DataProcessor class.
/// </summary>
public class DataProcessor : IDisposable
{
    private readonly XLWorkbook _workbook;
    private readonly IXLWorksheet _worksheet;
    public Dictionary<string, string> ColumnMap { get; }

    public DataProcessor(string excelPath, string mappingPath)
    {
        _workbook = new XLWorkbook(excelPath);
        _worksheet = _workbook.Worksheet(1);
        ColumnMap = LoadMapping(mappingPath);
    }

    private Dictionary<string, string> LoadMapping(string mappingPath)
    {
        var json = File.ReadAllText(mappingPath);
        return JsonSerializer.Deserialize<Dictionary<string, string>>(json) 
            ?? new Dictionary<string, string>();
    }

    /// <summary>
    /// Converts column letter (A, B, AA, etc.) to 1-based index.
    /// Mirrors Python's column_to_index function.
    /// </summary>
    public static int ColumnToIndex(string columnName)
    {
        int index = 0;
        for (int i = 0; i < columnName.Length; i++)
        {
            index = index * 26 + (char.ToUpper(columnName[i]) - 'A' + 1);
        }
        return index; // ClosedXML uses 1-based indexing
    }

    /// <summary>
    /// Gets all student rows from the Excel file.
    /// Skips first 4 rows (like Python: df.drop([0, 1, 2, 3])) - starts from row 5.
    /// </summary>
    public IEnumerable<IXLRow> GetStudentRows()
    {
        var lastRow = _worksheet.LastRowUsed()?.RowNumber() ?? 0;
        
        // Skip first 4 rows (rows 1-4), start from row 5
        for (int rowNum = 5; rowNum <= lastRow; rowNum++)
        {
            var row = _worksheet.Row(rowNum);
            // Skip empty rows
            if (!row.IsEmpty())
            {
                yield return row;
            }
        }
    }

    /// <summary>
    /// Process individual student data for report card generation.
    /// Mirrors Python's process_student_data method.
    /// </summary>
    public Dictionary<string, string> ProcessStudentData(IXLRow row, string className)
    {
        var fieldDict = new Dictionary<string, string>();
        var nanValues = new HashSet<string> { "NAN", "NONE", "NA", "" };

        foreach (var kvp in ColumnMap)
        {
            var key = kvp.Key;
            var columnLetter = kvp.Value;
            var columnIndex = ColumnToIndex(columnLetter);
            
            var cellValue = row.Cell(columnIndex).GetString();
            
            // Handle null/empty values (like Python code)
            if (string.IsNullOrEmpty(cellValue) || 
                nanValues.Contains(cellValue.ToUpper().Replace(" ", "")))
            {
                fieldDict[key] = "---";
            }
            else
            {
                fieldDict[key] = cellValue;
            }
        }

        // Add class name (replace underscores with spaces like Python)
        fieldDict["class"] = className.Replace("_", " ");

        return fieldDict;
    }

    public void Dispose()
    {
        _workbook?.Dispose();
    }
}


