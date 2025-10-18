using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AiCalc.ViewModels;

namespace AiCalc.Services;

/// <summary>
/// Provides CSV export and import functionality (Phase 6)
/// </summary>
public static class CsvService
{
    /// <summary>
    /// Export a sheet to CSV format
    /// </summary>
    public static async Task ExportSheetToCsvAsync(SheetViewModel sheet, string filePath)
    {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("File path cannot be empty", nameof(filePath));

        var sb = new StringBuilder();

        // Get dimensions
        int maxRow = sheet.Rows.Count;
        int maxCol = sheet.ColumnCount;

        // Write rows
        for (int row = 0; row < maxRow; row++)
        {
            var rowValues = new string[maxCol];
            
            for (int col = 0; col < maxCol; col++)
            {
                var cell = sheet.GetCell(row, col);
                if (cell != null)
                {
                    // Use display value, escape quotes
                    var value = cell.Value?.DisplayValue ?? "";
                    rowValues[col] = EscapeCsvValue(value);
                }
                else
                {
                    rowValues[col] = "";
                }
            }

            sb.AppendLine(string.Join(",", rowValues));
        }

        await File.WriteAllTextAsync(filePath, sb.ToString(), Encoding.UTF8);
    }

    /// <summary>
    /// Import CSV data into a new sheet
    /// </summary>
    public static async Task<SheetViewModel> ImportCsvToSheetAsync(string filePath, string sheetName, WorkbookViewModel workbook)
    {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("File path cannot be empty", nameof(filePath));
        if (!File.Exists(filePath)) throw new FileNotFoundException("CSV file not found", filePath);

        var csvLines = await File.ReadAllLinesAsync(filePath, Encoding.UTF8);
        
        // Determine dimensions
        int maxColumns = 0;
        var lines = new List<string[]>();
        
        foreach (var line in csvLines)
        {
            if (string.IsNullOrWhiteSpace(line)) continue;
            
            var fields = ParseCsvLine(line);
            lines.Add(fields);
            if (fields.Length > maxColumns)
            {
                maxColumns = fields.Length;
            }
        }

        var rowCount = lines.Count;

        // Create a new sheet with appropriate dimensions
        var sheet = new SheetViewModel(workbook, sheetName, 
            Math.Max(rowCount, 10),  // At least 10 rows
            Math.Max(maxColumns, 10)); // At least 10 columns

        // Populate cells
        for (int row = 0; row < lines.Count; row++)
        {
            var fields = lines[row];
            for (int col = 0; col < fields.Length; col++)
            {
                var cell = sheet.GetCell(row, col);
                if (cell != null && !string.IsNullOrEmpty(fields[col]))
                {
                    cell.RawValue = fields[col];
                }
            }
        }

        return sheet;
    }

    private static string EscapeCsvValue(string value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return "";
        }

        // If contains comma, quote, or newline, wrap in quotes and escape quotes
        if (value.Contains(',') || value.Contains('"') || value.Contains('\n') || value.Contains('\r'))
        {
            return $"\"{value.Replace("\"", "\"\"")}\"";
        }

        return value;
    }

    private static string[] ParseCsvLine(string line)
    {
        var values = new List<string>();
        var currentValue = new StringBuilder();
        bool inQuotes = false;

        for (int i = 0; i < line.Length; i++)
        {
            char c = line[i];

            if (c == '"')
            {
                if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                {
                    // Escaped quote
                    currentValue.Append('"');
                    i++; // Skip next quote
                }
                else
                {
                    // Toggle quote mode
                    inQuotes = !inQuotes;
                }
            }
            else if (c == ',' && !inQuotes)
            {
                // End of value
                values.Add(currentValue.ToString());
                currentValue.Clear();
            }
            else
            {
                currentValue.Append(c);
            }
        }

        // Add last value
        values.Add(currentValue.ToString());

        return values.ToArray();
    }

    /// <summary>
    /// Export entire workbook to multiple CSV files (one per sheet)
    /// </summary>
    public static async Task ExportWorkbookToCsvAsync(WorkbookViewModel workbook, string directoryPath)
    {
        if (workbook == null) throw new ArgumentNullException(nameof(workbook));
        if (string.IsNullOrWhiteSpace(directoryPath)) throw new ArgumentException("Directory path cannot be empty", nameof(directoryPath));

        Directory.CreateDirectory(directoryPath);

        foreach (var sheet in workbook.Sheets)
        {
            var safeSheetName = string.Join("_", sheet.Name.Split(Path.GetInvalidFileNameChars()));
            var filePath = Path.Combine(directoryPath, $"{safeSheetName}.csv");
            await ExportSheetToCsvAsync(sheet, filePath);
        }
    }
}
