using ClosedXML.Excel;

namespace LMM;

public sealed class ClosedXmlDataSourceService : IDisposable
{
    private readonly XLWorkbook _workbook;
    private bool _disposed;

    public ClosedXmlDataSourceService(string xlsxPath)
    {
        if (string.IsNullOrWhiteSpace(xlsxPath))
            throw new ArgumentException("Excel path is required.", nameof(xlsxPath));

        _workbook = new XLWorkbook(xlsxPath);
    }

    public IReadOnlyList<string> GetWorksheetNames()
    {
        ThrowIfDisposed();

        var names = _workbook.Worksheets.Select(w => w.Name).ToList();
        if (names.Count == 0)
            throw new InvalidOperationException("The Excel file contains no worksheets.");

        return names;
    }

    public ExcelHeaderResult ReadHeaders(
        string worksheetName,
        int headerRowNumber,
        bool trimHeaders = true,
        bool errorOnEmptyHeader = true,
        bool errorOnDuplicateHeader = true)
    {
        ThrowIfDisposed();

        if (headerRowNumber < 1)
            throw new ArgumentOutOfRangeException(nameof(headerRowNumber), "Header row must be >= 1.");

        var ws = GetWorksheetOrThrow(worksheetName);

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
            var header = ws.Cell(headerRowNumber, col).GetString() ?? string.Empty;
            if (trimHeaders) header = header.Trim();
            headers.Add(header);
        }

        // Drop trailing empty headers (helps when used range is wide due to formatting)
        while (headers.Count > 1 && string.IsNullOrWhiteSpace(headers[^1]))
            headers.RemoveAt(headers.Count - 1);

        if (headers.All(string.IsNullOrWhiteSpace))
            throw new InvalidOperationException(
                $"Header row {headerRowNumber} in '{worksheetName}' does not contain any header text.");

        if (errorOnEmptyHeader)
        {
            var emptyPositions = headers
                .Select((h, idx) => new { Header = h, Index = idx })
                .Where(x => string.IsNullOrWhiteSpace(x.Header))
                .Select(x => x.Index + 1) // 1-based position
                .ToList();

            if (emptyPositions.Count > 0)
            {
                throw new InvalidOperationException(
                    "The header row contains empty column names at positions: " +
                    string.Join(", ", emptyPositions) +
                    ". Please fill them in or reduce the used range.");
            }
        }

        if (errorOnDuplicateHeader)
        {
            var duplicates = headers
                .Where(h => !string.IsNullOrWhiteSpace(h))
                .GroupBy(h => h, StringComparer.Ordinal)
                .Where(g => g.Count() > 1)
                .Select(g => g.Key)
                .OrderBy(x => x, StringComparer.Ordinal)
                .ToList();

            if (duplicates.Count > 0)
            {
                throw new InvalidOperationException(
                    "Duplicate header names found (exact match): " +
                    string.Join(", ", duplicates) +
                    ". Header names must be unique.");
            }
        }

        return new ExcelHeaderResult(
            WorksheetName: worksheetName,
            HeaderRowNumber: headerRowNumber,
            Headers: headers,
            LastDataRowNumber: lastUsedRow);
    }

    /// <summary>
    /// Enumerates records using formatted/display values (GetFormattedString()).
    /// Missing/blank cells become "".
    /// </summary>
    public IEnumerable<Dictionary<string, string>> EnumerateRecordsFormatted(
        string worksheetName,
        int headerRowNumber,
        IReadOnlyList<string> headers,
        bool trimFormattedValues = false,
        bool skipFullyEmptyRows = false)
    {
        ThrowIfDisposed();

        var ws = GetWorksheetOrThrow(worksheetName);

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

                var formatted = cell.GetFormattedString() ?? string.Empty;
                if (trimFormattedValues)
                    formatted = formatted.Trim();

                record[header] = formatted;
            }

            if (skipFullyEmptyRows && record.Values.All(string.IsNullOrEmpty))
                continue;

            yield return record;
        }
    }

    public bool TryReadHeaders(
        string worksheetName,
        int headerRowNumber,
        out ExcelHeaderResult? headerResult,
        out string userMessage,
        bool trimHeaders = true,
        bool errorOnEmptyHeader = true,
        bool errorOnDuplicateHeader = true)
    {
        headerResult = null;
        userMessage = "";

        if (_disposed)
        {
            userMessage = "Excel file is not loaded (internal state disposed). Please re-select the Excel file.";
            return false;
        }

        if (string.IsNullOrWhiteSpace(worksheetName))
        {
            userMessage = "Please select a worksheet.";
            return false;
        }

        if (headerRowNumber < 1)
        {
            userMessage = "Header row must be 1 or greater.";
            return false;
        }

        try
        {
            var ws = GetWorksheetOrThrow(worksheetName);

            var usedRange = ws.RangeUsed();
            if (usedRange == null)
            {
                userMessage = $"Worksheet '{worksheetName}' is empty.";
                return false;
            }

            var firstUsedCol = usedRange.RangeAddress.FirstAddress.ColumnNumber;
            var lastUsedCol = usedRange.RangeAddress.LastAddress.ColumnNumber;
            var lastUsedRow = usedRange.RangeAddress.LastAddress.RowNumber;

            if (headerRowNumber > lastUsedRow)
            {
                userMessage =
                    $"Header row {headerRowNumber} is below the last used row ({lastUsedRow}) " +
                    $"in worksheet '{worksheetName}'.";
                return false;
            }

            var headers = new List<string>();
            for (int col = firstUsedCol; col <= lastUsedCol; col++)
            {
                var header = ws.Cell(headerRowNumber, col).GetString() ?? string.Empty;
                if (trimHeaders) header = header.Trim();
                headers.Add(header);
            }

            while (headers.Count > 1 && string.IsNullOrWhiteSpace(headers[^1]))
                headers.RemoveAt(headers.Count - 1);

            if (headers.All(string.IsNullOrWhiteSpace))
            {
                userMessage =
                    $"Header row {headerRowNumber} in worksheet '{worksheetName}' does not contain any header text.";
                return false;
            }

            if (errorOnEmptyHeader)
            {
                var emptyPositions = headers
                    .Select((h, idx) => new { Header = h, Index = idx })
                    .Where(x => string.IsNullOrWhiteSpace(x.Header))
                    .Select(x => x.Index + 1)
                    .ToList();

                if (emptyPositions.Count > 0)
                {
                    userMessage =
                        "The header row contains empty column names at positions: " +
                        string.Join(", ", emptyPositions) +
                        ". Please fill them in or reduce the used range.";
                    return false;
                }
            }

            if (errorOnDuplicateHeader)
            {
                var duplicates = headers
                    .Where(h => !string.IsNullOrWhiteSpace(h))
                    .GroupBy(h => h, StringComparer.Ordinal)
                    .Where(g => g.Count() > 1)
                    .Select(g => g.Key)
                    .OrderBy(x => x, StringComparer.Ordinal)
                    .ToList();

                if (duplicates.Count > 0)
                {
                    userMessage =
                        "Duplicate header names found: " + string.Join(", ", duplicates) + ". " +
                        "Header names must be unique.";
                    return false;
                }
            }

            headerResult = new ExcelHeaderResult(
                WorksheetName: worksheetName,
                HeaderRowNumber: headerRowNumber,
                Headers: headers,
                LastDataRowNumber: lastUsedRow);

            return true;
        }
        catch (Exception ex)
        {
            // Unexpected failure (file issue, invalid workbook, etc.)
            userMessage =
                "Failed to read headers from the Excel file. " +
                "Please ensure the file is a valid .xlsx and is not locked.\n\n" +
                "Details: " + ex.Message;

            return false;
        }
    }

    public IEnumerable<(int RowNumber, Dictionary<string, string> Record)> EnumerateRecordsFormattedWithRowNumber(
        string worksheetName,
        int headerRowNumber,
        IReadOnlyList<string> headers,
        bool trimFormattedValues = false,
        bool skipFullyEmptyRows = false)
    {
        ThrowIfDisposed();

        var ws = GetWorksheetOrThrow(worksheetName);
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

                var formatted = cell.GetFormattedString() ?? string.Empty;
                if (trimFormattedValues)
                    formatted = formatted.Trim();

                record[header] = formatted;
            }

            if (skipFullyEmptyRows && record.Values.All(string.IsNullOrEmpty))
                continue;

            yield return (row, record);
        }
    }


    private IXLWorksheet GetWorksheetOrThrow(string worksheetName)
    {
        if (string.IsNullOrWhiteSpace(worksheetName))
            throw new ArgumentException("Worksheet name is required.", nameof(worksheetName));

        if (!_workbook.Worksheets.TryGetWorksheet(worksheetName, out var ws))
            throw new InvalidOperationException($"Worksheet '{worksheetName}' was not found in the workbook.");

        return ws;
    }

    private void ThrowIfDisposed()
    {
        if (_disposed)
            throw new ObjectDisposedException(nameof(ClosedXmlDataSourceService));
    }

    public void Dispose()
    {
        if (_disposed) return;
        _workbook.Dispose();
        _disposed = true;
    }
}