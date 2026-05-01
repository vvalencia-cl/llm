using LMM.Application;
using LMM.Domain.Dto;

namespace LMM.UI;

public sealed class MainPresenter
{
    private readonly IMainView _view;
    private ClosedXmlDataSourceService? _excel;
    private ExcelHeaderResult? _headerInfo;
    private List<string> _templateFields = [];
    private Dictionary<string, string> _templateToExcelHeaderMap = new(StringComparer.Ordinal);
    private CancellationTokenSource? _cts;

    private const string OptionalFieldNoneOption = "(Ninguno)";

    public MainPresenter(IMainView view)
    {
        _view = view;
        WireEvents();
        UpdateState();
    }

    private void WireEvents()
    {
        _view.BrowseTemplateClicked += (s, e) => BrowseTemplate();
        _view.BrowseExcelClicked += (s, e) => BrowseExcel();
        _view.BrowseOutputDirClicked += (s, e) => BrowseOutputDir();
        _view.LoadExcelClicked += async (s, e) => await LoadExcelAsync();
        _view.RefreshHeadersClicked += (s, e) => RefreshHeaders();
        _view.ScanTemplateClicked += (s, e) => ScanTemplateAndValidate();
        _view.RunClicked += async (s, e) => await RunMergeAsync();
        _view.CancelClicked += (s, e) => _cts?.Cancel();

        _view.TemplatePathChanged += (s, e) => { _templateFields.Clear(); UpdateState(); };
        _view.ExcelPathChanged += (s, e) => { _templateFields.Clear(); UpdateState(); };
        _view.OutputDirChanged += (s, e) => UpdateState();
        _view.ExcelOptionsChanged += (s, e) => RefreshHeaders();
        _view.FilenameConfigChanged += (s, e) => UpdateFilenamePreview();
    }

    private void UpdateState()
    {
        var state = new AppState(
            ExcelLoaded: _excel != null,
            HeadersReady: _headerInfo?.Headers?.Count > 0,
            TemplateExists: File.Exists(_view.TemplatePath),
            OutputDirExists: Directory.Exists(_view.OutputDir),
            IsProcessing: _cts != null
        );
        _view.ApplyState(state);
    }

    private void BrowseTemplate()
    {
        var path = _view.PickFile("Seleccionar plantilla de Word", "Plantilla de Word (*.docx)|*.docx");
        if (path != null)
        {
            _view.TemplatePath = path;
            _view.AppendLog($"Plantilla: {path}");
        }
    }

    private void BrowseExcel()
    {
        var path = _view.PickFile("Seleccionar archivo de Excel", "Libro de Excel (*.xlsx)|*.xlsx");
        if (path != null)
        {
            _view.ExcelPath = path;
            _view.AppendLog($"Excel: {path}");
        }
    }

    private void BrowseOutputDir()
    {
        var path = _view.PickFolder("Seleccione la carpeta de salida para los PDFs");
        if (path != null)
        {
            _view.OutputDir = path;
            _view.AppendLog($"Salida: {path}");
        }
    }

    private async Task LoadExcelAsync()
    {
        await ExecuteAsync("Analizando Excel...", async () =>
        {
            _excel?.Dispose();
            _excel = null;
            _headerInfo = null;

            var path = _view.ExcelPath;
            if (!File.Exists(path))
            {
                _view.ShowWarning("Excel", "Por favor, seleccione un archivo de Excel válido.");
                return;
            }

            // Simular carga asíncrona real si el servicio lo permitiera, 
            // pero como no queremos tocar el servicio, lo corremos en un Task.
            _excel = await Task.Run(() => new ClosedXmlDataSourceService(path));

            var sheets = _excel.GetWorksheetNames().ToList();
            _view.SetWorksheets(sheets);
            if (sheets.Count > 0)
                _view.SelectedWorksheet = sheets[0];

            _view.AppendLog($"Excel cargado. Hojas: {sheets.Count}");
            RefreshHeaders();
        }, "Error al analizar Excel");
    }

    private void RefreshHeaders()
    {
        _headerInfo = null;
        var worksheet = _view.SelectedWorksheet;

        if (_excel == null || worksheet == null)
        {
            UpdateState();
            return;
        }

        if (!_excel.TryReadHeaders(
                worksheetName: worksheet,
                headerRowNumber: _view.HeaderRow,
                headerResult: out var headerInfo,
                userMessage: out var msg))
        {
            _view.AppendLog("Encabezados: ERROR -> " + msg);
            _view.SetExcelHeaders([]);
            UpdateFilenamePreview();
            UpdateState();
            return;
        }

        _headerInfo = headerInfo!;
        _view.SetExcelHeaders(_headerInfo.Headers);
        
        _templateFields.Clear();
        _templateToExcelHeaderMap = new Dictionary<string, string>(StringComparer.Ordinal);

        _view.AppendLog($"Encabezados cargados. Cantidad: {_headerInfo.Headers.Count}. Fila de encabezado: {_headerInfo.HeaderRowNumber}");
        UpdateFilenamePreview();
        UpdateState();
    }

    private void UpdateFilenamePreview()
    {
        if (_headerInfo == null || _view.FilenameFirstField == null)
        {
            _view.SetFilenamePreview("");
            return;
        }

        var fakeRecord = _headerInfo.Headers
            .Where(h => !string.IsNullOrWhiteSpace(h))
            .ToDictionary(h => h, h => $"<{h}>", StringComparer.Ordinal);

        var second = _view.FilenameSecondField == OptionalFieldNoneOption ? null : _view.FilenameSecondField;
        var third = _view.FilenameThirdField == OptionalFieldNoneOption ? null : _view.FilenameThirdField;

        var previewPath = PdfFilenameBuilder.BuildPdfPath(
            outputDirectory: Directory.Exists(_view.OutputDir) ? _view.OutputDir : Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            record: fakeRecord,
            prefix: _view.FilenamePrefix,
            firstFieldHeader: _view.FilenameFirstField,
            secondFieldHeader: second,
            thirdFieldHeader: third,
            postfix: _view.FilenamePostfix,
            separator: _view.FilenameSeparator,
            emptyFallback: "fila_###");

        _view.SetFilenamePreview(Path.GetFileName(previewPath));
    }

    private void ScanTemplateAndValidate()
    {
        try
        {
            if (!TryScanTemplateAndValidate(out var msg))
            {
                _view.ShowError("Validación de plantilla", msg);
                _view.AppendLog("Validación de plantilla: ERROR -> " + msg.Replace("\n", " "));
                return;
            }

            _view.AppendLog($"Campos de plantilla encontrados: {_templateFields.Count}");
            _view.AppendLog("Campos de plantilla validados contra encabezados de Excel: OK");
        }
        catch (Exception ex)
        {
            _view.ShowError("Error al escanear plantilla", ex.Message);
            _templateFields.Clear();
        }
        finally
        {
            UpdateState();
        }
    }

    private async Task RunMergeAsync()
    {
        if (_excel == null || _headerInfo == null)
        {
            _view.ShowWarning("Ejecutar", "Por favor, cargue primero el Excel y los encabezados.");
            return;
        }

        if (!TryScanTemplateAndValidate(out var validationMsg))
        {
            _view.ShowError("Validación de plantilla", validationMsg);
            _view.AppendLog("Validación de plantilla: ERROR -> " + validationMsg.Replace("\n", " "));
            UpdateState();
            return;
        }

        var config = GetCurrentRunConfig();
        if (!File.Exists(config.TemplatePath))
        {
            _view.ShowWarning("Ejecutar", "Por favor, seleccione una plantilla de Word válida (.docx).");
            return;
        }

        if (!Directory.Exists(config.OutputDir))
        {
            _view.ShowWarning("Ejecutar", "Por favor, seleccione una carpeta de salida válida.");
            return;
        }

        _cts = new CancellationTokenSource();
        UpdateState();
        _view.UpdateProgress(0, 100);
        _view.AppendLog("Iniciando combinación...");

        var totalEstimate = Math.Max(1, _headerInfo.LastDataRowNumber - config.HeaderRow);
        
        var progress = new Progress<MergeProgress>(p =>
        {
            _view.UpdateProgress(p.Done, p.Total);
            if (!string.IsNullOrWhiteSpace(p.Message))
                _view.AppendLog(p.Message);
        });

        try
        {
            await StaTaskRunner.Run(() =>
            {
                using var exporter = new WordPdfExporter();
                int done = 0, ok = 0, fail = 0;

                foreach (var (rowNumber, excelRecord) in _excel.EnumerateRecordsFormattedWithRowNumber(
                             worksheetName: config.WorksheetName,
                             headerRowNumber: config.HeaderRow,
                             headers: _headerInfo.Headers,
                             skipFullyEmptyRows: true))
                {
                    _cts.Token.ThrowIfCancellationRequested();
                    done++;

                    try
                    {
                        var pdfPath = PdfFilenameBuilder.BuildPdfPath(
                            outputDirectory: config.OutputDir,
                            record: excelRecord,
                            prefix: config.Prefix,
                            firstFieldHeader: config.FirstField,
                            secondFieldHeader: config.SecondField,
                            thirdFieldHeader: config.ThirdField,
                            postfix: config.Postfix,
                            separator: config.Separator,
                            emptyFallback: $"row_{rowNumber}");

                        if (File.Exists(pdfPath)) File.Delete(pdfPath);

                        var templateValues = HeaderFieldMapper.BuildTemplateValuesForRecord(
                            templateFields: _templateFields,
                            excelRecord: excelRecord,
                            templateToExcelHeaderMap: _templateToExcelHeaderMap);

                        MailMergeToPdf.MergeAndExportPdfForRecord(
                            record: templateValues,
                            pdfPath: pdfPath,
                            templateDocxPath: config.TemplatePath,
                            wordExporter: exporter);

                        ok++;
                        ((IProgress<MergeProgress>)progress).Report(new MergeProgress(done, totalEstimate, $"Fila {rowNumber}: OK -> {Path.GetFileName(pdfPath)}"));
                    }
                    catch (Exception ex)
                    {
                        fail++;
                        ((IProgress<MergeProgress>)progress).Report(new MergeProgress(done, totalEstimate, $"Fila {rowNumber}: ERROR -> {ex.Message}"));
                    }
                    ((IProgress<MergeProgress>)progress).Report(new MergeProgress(done, totalEstimate, null));
                }

                ((IProgress<MergeProgress>)progress).Report(new MergeProgress(totalEstimate, totalEstimate, $"Completado. OK: {ok}, Errores: {fail}."));
                
                // Alert in UI thread
                Task.Run(() => _view.ShowInfo("Combinación completada", $"Proceso finalizado.\nOK: {ok}\nErrores: {fail}"));
            }, _cts.Token);
        }
        catch (OperationCanceledException)
        {
            _view.AppendLog("Cancelado.");
        }
        catch (Exception ex)
        {
            _view.AppendLog("Error en la ejecución: " + ex.Message);
            _view.ShowError("Error en la ejecución", ex.Message);
        }
        finally
        {
            _cts?.Dispose();
            _cts = null;
            UpdateState();
        }
    }

    private bool TryScanTemplateAndValidate(out string userMessage)
    {
        userMessage = "";
        if (_headerInfo == null)
        {
            userMessage = "Por favor, cargue primero los encabezados de Excel.";
            return false;
        }

        var templatePath = _view.TemplatePath;
        if (!File.Exists(templatePath))
        {
            userMessage = "Por favor, seleccione una plantilla de Word válida (.docx).";
            return false;
        }

        var fields = WordTemplateFieldScanner.GetMergeFieldNamesFromMainBody(templatePath);
        _templateFields = fields.ToList();

        if (_templateFields.Count == 0)
        {
            userMessage = "No se encontraron campos MERGEFIELD en la plantilla de Word.";
            return false;
        }

        try
        {
            _templateToExcelHeaderMap = HeaderFieldMapper.BuildTemplateToExcelHeaderMap(_templateFields, _headerInfo.Headers.ToList());
            var missing = _templateToExcelHeaderMap.Where(kvp => string.IsNullOrEmpty(kvp.Value)).Select(kvp => kvp.Key).OrderBy(x => x).ToList();

            if (missing.Count > 0)
            {
                userMessage = "La plantilla de Word contiene campos de combinación que faltan en la fila de encabezado de Excel:\n\n" +
                              string.Join("\n", missing.Select(m => $" - {m}")) +
                              "\n\nSolución: Agregue estas columnas a la fila de encabezado de Excel.";
                _templateFields.Clear();
                _templateToExcelHeaderMap = [];
                return false;
            }
            return true;
        }
        catch (Exception ex)
        {
            userMessage = ex.Message;
            _templateFields.Clear();
            _templateToExcelHeaderMap = [];
            return false;
        }
    }

    private async Task ExecuteAsync(string logMsg, Func<Task> action, string errorTitle)
    {
        try
        {
            _view.SetBusy(true);
            _view.AppendLog(logMsg);
            await action();
        }
        catch (Exception ex)
        {
            _view.AppendLog($"{logMsg} ERROR -> {ex.Message}");
            _view.ShowError(errorTitle, ex.Message);
        }
        finally
        {
            _view.SetBusy(false);
            UpdateState();
        }
    }

    private RunConfig GetCurrentRunConfig() => new(
        _view.TemplatePath,
        _view.OutputDir,
        _view.SelectedWorksheet!,
        _view.HeaderRow,
        _view.FilenameFirstField!,
        _view.FilenameSecondField == OptionalFieldNoneOption ? null : _view.FilenameSecondField,
        _view.FilenameThirdField == OptionalFieldNoneOption ? null : _view.FilenameThirdField,
        _view.FilenamePrefix,
        _view.FilenamePostfix,
        _view.FilenameSeparator
    );

    private sealed record RunConfig(string TemplatePath, string OutputDir, string WorksheetName, int HeaderRow, 
        string FirstField, string? SecondField, string? ThirdField, string Prefix, string Postfix, string Separator);

    public void DoDispose()
    {
        _cts?.Cancel();
        _cts?.Dispose();
        _excel?.Dispose();
    }
}
