using ClosedXML.Excel;
using LMM.Domain.Dto;

namespace LMM.Application;

public sealed class ClosedXmlDataSourceService : IDisposable
{
    private readonly XLWorkbook _workbook;
    private bool _disposed;

    public ClosedXmlDataSourceService(string xlsxPath)
    {
        if (string.IsNullOrWhiteSpace(xlsxPath))
            throw new ArgumentException("La ruta del archivo Excel es obligatoria.", nameof(xlsxPath));

        _workbook = new XLWorkbook(xlsxPath);
    }

    public IReadOnlyList<string> GetWorksheetNames()
    {
        ThrowIfDisposed();

        var names = _workbook.Worksheets.Select(w => w.Name).ToList();
        if (names.Count == 0)
            throw new InvalidOperationException("El archivo Excel no contiene hojas de trabajo.");

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
            throw new ArgumentOutOfRangeException(nameof(headerRowNumber), "La fila de encabezado debe ser >= 1.");

        var ws = GetWorksheetOrThrow(worksheetName);

        var usedRange = ws.RangeUsed();
        if (usedRange == null)
            throw new InvalidOperationException($"La hoja de trabajo '{worksheetName}' está vacía.");

        var firstUsedCol = usedRange.RangeAddress.FirstAddress.ColumnNumber;
        var lastUsedCol = usedRange.RangeAddress.LastAddress.ColumnNumber;
        var lastUsedRow = usedRange.RangeAddress.LastAddress.RowNumber;

        if (headerRowNumber > lastUsedRow)
            throw new InvalidOperationException(
                $"La fila de encabezado {headerRowNumber} está por debajo de la última fila utilizada ({lastUsedRow}) en '{worksheetName}'.");

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
                $"La fila de encabezado {headerRowNumber} en '{worksheetName}' no contiene ningún texto de encabezado.");

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
                    "La fila de encabezado contiene nombres de columna vacíos en las posiciones: " +
                    string.Join(", ", emptyPositions) +
                    ". Por favor, rellénelos o reduzca el rango utilizado.");
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
                    "Se encontraron nombres de encabezado duplicados (coincidencia exacta): " +
                    string.Join(", ", duplicates) +
                    ". Los nombres de los encabezados deben ser únicos.");
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
            userMessage = "El archivo Excel no está cargado (estado interno liberado). Por favor, vuelva a seleccionar el archivo Excel.";
            return false;
        }

        if (string.IsNullOrWhiteSpace(worksheetName))
        {
            userMessage = "Por favor, seleccione una hoja de trabajo.";
            return false;
        }

        if (headerRowNumber < 1)
        {
            userMessage = "La fila de encabezado debe ser 1 o superior.";
            return false;
        }

        try
        {
            var ws = GetWorksheetOrThrow(worksheetName);

            var usedRange = ws.RangeUsed();
            if (usedRange == null)
            {
                userMessage = $"La hoja de trabajo '{worksheetName}' está vacía.";
                return false;
            }

            var firstUsedCol = usedRange.RangeAddress.FirstAddress.ColumnNumber;
            var lastUsedCol = usedRange.RangeAddress.LastAddress.ColumnNumber;
            var lastUsedRow = usedRange.RangeAddress.LastAddress.RowNumber;

            if (headerRowNumber > lastUsedRow)
            {
                userMessage =
                    $"La fila de encabezado {headerRowNumber} está por debajo de la última fila utilizada ({lastUsedRow}) " +
                    $"en la hoja de trabajo '{worksheetName}'.";
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
                    $"La fila de encabezado {headerRowNumber} en la hoja de trabajo '{worksheetName}' no contiene ningún texto de encabezado.";
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
                        "La fila de encabezado contiene nombres de columna vacíos en las posiciones: " +
                        string.Join(", ", emptyPositions) +
                        ". Por favor, rellénelos o reduzca el rango utilizado.";
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
                        "Se encontraron nombres de encabezado duplicados: " + string.Join(", ", duplicates) + ". " +
                        "Los nombres de los encabezados deben ser únicos.";
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
                "Error al leer los encabezados del archivo Excel. " +
                "Por favor, asegúrese de que el archivo sea un .xlsx válido y no esté bloqueado.\n\n" +
                "Detalles: " + ex.Message;

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
            throw new ArgumentException("El nombre de la hoja de trabajo es obligatorio.", nameof(worksheetName));

        if (!_workbook.Worksheets.TryGetWorksheet(worksheetName, out var ws))
            throw new InvalidOperationException($"La hoja de trabajo '{worksheetName}' no se encontró en el libro.");

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