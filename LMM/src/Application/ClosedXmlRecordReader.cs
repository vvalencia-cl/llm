using ClosedXML.Excel;

namespace LMM.Application;

public static class ClosedXmlRecordReader
{
    public static IEnumerable<Dictionary<string, string>> ReadRecordsFormatted(
        string xlsxPath,
        string worksheetName,
        int headerRowNumber,
        IReadOnlyList<string> headers)
    {
        using var wb = new XLWorkbook(xlsxPath);
        var ws = wb.Worksheet(worksheetName);

        var usedRange = ws.RangeUsed();
        if (usedRange == null)
            yield break;

        var firstUsedCol = usedRange.RangeAddress.FirstAddress.ColumnNumber;
        var lastUsedRow = usedRange.RangeAddress.LastAddress.RowNumber;

        for (int row = headerRowNumber + 1; row <= lastUsedRow; row++)
        {
            var record = new Dictionary<string, string>(StringComparer.Ordinal);

            for (int i = 0; i < headers.Count; i++)
            {
                var header = headers[i];
                if (string.IsNullOrWhiteSpace(header))
                    continue;

                var col = firstUsedCol + i;
                var cell = ws.Cell(row, col);

                // Display/formatted value (what Excel shows)
                var formatted = cell.GetFormattedString();

                // Your rule: blank/null -> ""
                record[header] = formatted ?? string.Empty;
            }

            yield return record;
        }
    }
}