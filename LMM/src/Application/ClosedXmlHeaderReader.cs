using ClosedXML.Excel;
using LMM.Domain.Dto;

namespace LMM.Application;

public static class ClosedXmlHeaderReader
{
    public static ExcelHeaderResult ReadHeaders(
        string xlsxPath,
        string worksheetName,
        int headerRowNumber,
        bool trimHeaders = true,
        bool errorOnEmptyHeader = true,
        bool errorOnDuplicateHeader = true)
    {
        if (headerRowNumber < 1)
            throw new ArgumentOutOfRangeException(nameof(headerRowNumber), "Header row must be >= 1.");

        using var wb = new XLWorkbook(xlsxPath);
        var ws = wb.Worksheet(worksheetName);

        var usedRange = ws.RangeUsed();
        if (usedRange == null)
            throw new InvalidOperationException($"Worksheet '{worksheetName}' is empty.");

        var firstUsedCol = usedRange.RangeAddress.FirstAddress.ColumnNumber;
        var lastUsedCol = usedRange.RangeAddress.LastAddress.ColumnNumber;
        var lastUsedRow = usedRange.RangeAddress.LastAddress.RowNumber;

        if (headerRowNumber > lastUsedRow)
            throw new InvalidOperationException(
                $"Header row {headerRowNumber} is below the last used row ({lastUsedRow}) in '{worksheetName}'.");

        var headers = new List<string>();

        for (int col = firstUsedCol; col <= lastUsedCol; col++)
        {
            var cell = ws.Cell(headerRowNumber, col);
            var header = cell.GetString() ?? string.Empty;

            if (trimHeaders)
                header = header.Trim();

            headers.Add(header);
        }

        // Optionally drop trailing empty headers (common when used range is wide)
        // Keep at least 1 column.
        while (headers.Count > 1 && string.IsNullOrWhiteSpace(headers[^1]))
            headers.RemoveAt(headers.Count - 1);

        if (headers.All(h => string.IsNullOrWhiteSpace(h)))
            throw new InvalidOperationException(
                $"Header row {headerRowNumber} in '{worksheetName}' does not contain any header text.");

        if (errorOnEmptyHeader)
        {
            var empties = headers
                .Select((h, idx) => new { h, idx })
                .Where(x => string.IsNullOrWhiteSpace(x.h))
                .Select(x => x.idx + 1) // 1-based index within extracted header list
                .ToList();

            if (empties.Count > 0)
            {
                throw new InvalidOperationException(
                    "The header row contains empty column names at positions: " +
                    string.Join(", ", empties) +
                    ". Please fill them in or remove those columns from the used range.");
            }
        }

        if (errorOnDuplicateHeader)
        {
            var duplicates = headers
                .Where(h => !string.IsNullOrWhiteSpace(h))
                .GroupBy(h => h, StringComparer.Ordinal) // exact match
                .Where(g => g.Count() > 1)
                .Select(g => g.Key)
                .OrderBy(x => x, StringComparer.Ordinal)
                .ToList();

            if (duplicates.Count > 0)
            {
                throw new InvalidOperationException(
                    "Duplicate header names found (exact match): " + string.Join(", ", duplicates) +
                    ". Header names must be unique.");
            }
        }

        return new ExcelHeaderResult(
            WorksheetName: worksheetName,
            HeaderRowNumber: headerRowNumber,
            Headers: headers,
            LastDataRowNumber: lastUsedRow);
    }
}