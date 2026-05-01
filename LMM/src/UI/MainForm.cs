using LMM.Application;
using LMM.Domain.Dto;

namespace LMM.UI;

public sealed partial class MainForm : Form
{
    // State
    private ClosedXmlDataSourceService? _excel;
    private ExcelHeaderResult? _headerInfo;
    private List<string> _templateFields = [];
    private Dictionary<string, string> _templateToExcelHeaderMap = new(StringComparer.Ordinal);
    private CancellationTokenSource? _cts;

    private const string OptionalFieldNoneOption = "(Ninguno)";

    public MainForm()
    {
        InitializeComponent();
        WireEvents();
        UpdateUiEnabledState();
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        _cts?.Cancel();
        _cts?.Dispose();
        _excel?.Dispose();
        base.OnFormClosed(e);
    }

    private void WireEvents()
    {
        btnBrowseTemplate.Click += btnBrowseTemplate_Click;
        btnBrowseExcel.Click += btnBrowseExcel_Click;
        btnBrowseOutputDir.Click += btnBrowseOutputDir_Click;
        btnLoadExcel.Click += btnLoadExcel_Click;
        btnRefreshHeaders.Click += btnRefreshHeaders_Click;

        cmbWorksheet.SelectedIndexChanged += cmbWorksheet_SelectedIndexChanged;
        numHeaderRow.ValueChanged += numHeaderRow_ValueChanged;

        cmbFirstField.SelectedIndexChanged += cmbFirstField_SelectedIndexChanged;
        cmbSecondField.SelectedIndexChanged += cmbSecondField_SelectedIndexChanged;
        cmbThirdField.SelectedIndexChanged += cmbThirdField_SelectedIndexChanged;
        txtPrefix.TextChanged += txtPrefix_TextChanged;
        txtPostfix.TextChanged += txtPostfix_TextChanged;
        txtSeparator.TextChanged += txtSeparator_TextChanged;
        txtOutputDir.TextChanged += txtOutputDir_TextChanged;

        btnScanTemplateFields.Click += btnScanTemplateFields_Click;
        btnRun.Click += btnRun_Click;
        btnCancel.Click += btnCancel_Click;

        lstLog.KeyDown += lstLog_KeyDown;
    }

    private void btnBrowseTemplate_Click(object? sender, EventArgs e) => BrowseTemplate();
    private void btnBrowseExcel_Click(object? sender, EventArgs e) => BrowseExcel();
    private void btnBrowseOutputDir_Click(object? sender, EventArgs e) => BrowseOutputDir();

    private async void btnLoadExcel_Click(object? sender, EventArgs e)
    {
        // User feedback: disable relevant controls + show wait cursor + show marquee progress
        var prevCursor = Cursor.Current;
        try
        {
            btnLoadExcel.Enabled = false;
            btnBrowseExcel.Enabled = false;
            txtExcelPath.Enabled = false;

            progressBar.Style = ProgressBarStyle.Marquee;
            progressBar.MarqueeAnimationSpeed = 30;

            Cursor.Current = Cursors.WaitCursor;
            AppendLog("Analizando Excel... (esto puede tardar unos segundos)");

            await LoadExcelAsync();

            AppendLog("Excel analizado.");
        }
        catch (Exception ex)
        {
            AppendLog("Analizar Excel: ERROR -> " + ex.Message);
            MessageBox.Show(ex.Message, "Error al analizar Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            progressBar.MarqueeAnimationSpeed = 0;
            progressBar.Style = ProgressBarStyle.Blocks;

            Cursor.Current = prevCursor;

            // Re-evaluate enabled state based on current app state
            txtExcelPath.Enabled = true;
            btnBrowseExcel.Enabled = true;
            UpdateUiEnabledState();
        }
    }

    private void btnRefreshHeaders_Click(object? sender, EventArgs e) => RefreshHeaders();
    private void cmbWorksheet_SelectedIndexChanged(object? sender, EventArgs e) => RefreshHeaders();
    private void numHeaderRow_ValueChanged(object? sender, EventArgs e) => RefreshHeaders();

    private void cmbFirstField_SelectedIndexChanged(object? sender, EventArgs e) => UpdateFilenamePreview();
    private void cmbSecondField_SelectedIndexChanged(object? sender, EventArgs e) => UpdateFilenamePreview();
    private void cmbThirdField_SelectedIndexChanged(object? sender, EventArgs e) => UpdateFilenamePreview();
    private void txtPrefix_TextChanged(object? sender, EventArgs e) => UpdateFilenamePreview();
    private void txtPostfix_TextChanged(object? sender, EventArgs e) => UpdateFilenamePreview();
    private void txtSeparator_TextChanged(object? sender, EventArgs e) => UpdateFilenamePreview();
    private void txtOutputDir_TextChanged(object? sender, EventArgs e) => UpdateFilenamePreview();

    private void btnScanTemplateFields_Click(object? sender, EventArgs e) => ScanTemplateAndValidate();
    private async void btnRun_Click(object? sender, EventArgs e) => await RunMergeAsync();
    private void btnCancel_Click(object? sender, EventArgs e) => _cts?.Cancel();

    private void lstLog_KeyDown(object? sender, KeyEventArgs e)
    {
        if (e.Control && e.KeyCode == Keys.C)
        {
            CopySelectedLogLinesToClipboard();
            e.Handled = true;
            e.SuppressKeyPress = true;
        }
    }

    private void CopySelectedLogLinesToClipboard()
    {
        if (lstLog.SelectedItems.Count == 0)
            return;

        var lines = lstLog.SelectedItems
            .Cast<object>()
            .Select(x => x?.ToString() ?? string.Empty)
            .ToArray();

        var text = string.Join(Environment.NewLine, lines);

        try
        {
            Clipboard.SetText(text);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "Error al copiar al portapapeles.\n\nDetalles: " + ex.Message,
                "Portapapeles",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }

    private void UpdateUiEnabledState()
    {
        var excelLoaded = _excel != null;
        var headersReady = _headerInfo?.Headers?.Count > 0;

        btnRefreshHeaders.Enabled = excelLoaded;
        cmbWorksheet.Enabled = excelLoaded;
        numHeaderRow.Enabled = excelLoaded;

        cmbFirstField.Enabled = headersReady;
        cmbSecondField.Enabled = headersReady;
        cmbThirdField.Enabled = headersReady;
        txtPrefix.Enabled = headersReady;
        txtPostfix.Enabled = headersReady;
        txtSeparator.Enabled = headersReady;

        // Scan button can remain (optional tool), but is not required for running
        btnScanTemplateFields.Enabled = File.Exists(txtTemplatePath.Text) && headersReady;

        var canRun =
            File.Exists(txtTemplatePath.Text) &&
            excelLoaded &&
            headersReady &&
            Directory.Exists(txtOutputDir.Text) &&
            cmbFirstField.SelectedItem != null &&
            (_cts == null);

        btnRun.Enabled = canRun;
        btnCancel.Enabled = _cts != null;
    }

    private void AppendLog(string line)
    {
        if (InvokeRequired)
        {
            BeginInvoke(new Action<string>(AppendLog), line);
            return;
        }

        lstLog.Items.Add(line);
        lstLog.TopIndex = Math.Max(0, lstLog.Items.Count - 1);
    }

    private void BrowseTemplate()
    {
        using var dlg = new OpenFileDialog
        {
            Filter = "Plantilla de Word (*.docx)|*.docx",
            Title = "Seleccionar plantilla de Word"
        };
        if (dlg.ShowDialog(this) == DialogResult.OK)
        {
            txtTemplatePath.Text = dlg.FileName;
            _templateFields.Clear(); // require re-scan if template changed
            AppendLog($"Plantilla: {dlg.FileName}");
            UpdateUiEnabledState();
        }
    }

    private void BrowseExcel()
    {
        using var dlg = new OpenFileDialog
        {
            Filter = "Libro de Excel (*.xlsx)|*.xlsx",
            Title = "Seleccionar archivo de Excel"
        };
        if (dlg.ShowDialog(this) == DialogResult.OK)
        {
            txtExcelPath.Text = dlg.FileName;
            _templateFields.Clear(); // because mapping depends on headers
            AppendLog($"Excel: {dlg.FileName}");
            UpdateUiEnabledState();
        }
    }

    private void BrowseOutputDir()
    {
        using var dlg = new FolderBrowserDialog
        {
            Description = "Seleccione la carpeta de salida para los PDFs",
            UseDescriptionForTitle = true
        };
        if (dlg.ShowDialog(this) == DialogResult.OK)
        {
            txtOutputDir.Text = dlg.SelectedPath;
            AppendLog($"Salida: {dlg.SelectedPath}");
            UpdateUiEnabledState();
        }
    }

    private async Task LoadExcelAsync()
    {
        try
        {
            _excel?.Dispose();
            _excel = null;
            _headerInfo = null;

            var path = txtExcelPath.Text;
            if (!File.Exists(path))
            {
                MessageBox.Show("Por favor, seleccione un archivo de Excel válido.", "Excel", MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            _excel = new ClosedXmlDataSourceService(path);

            var sheets = _excel.GetWorksheetNames();
            cmbWorksheet.DataSource = sheets.ToList();
            cmbWorksheet.SelectedIndex = 0;

            AppendLog($"Excel cargado. Hojas: {sheets.Count}");
            RefreshHeaders();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Error al cargar Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            UpdateUiEnabledState();
        }

        await Task.CompletedTask;
    }

    private void RefreshHeaders()
    {
        _headerInfo = null;

        if (_excel == null || cmbWorksheet.SelectedItem == null)
        {
            UpdateUiEnabledState();
            return;
        }

        if (!_excel.TryReadHeaders(
                worksheetName: (string)cmbWorksheet.SelectedItem,
                headerRowNumber: (int)numHeaderRow.Value,
                headerResult: out var headerInfo,
                userMessage: out var msg))
        {
            AppendLog("Encabezados: ERROR -> " + msg);
            cmbFirstField.DataSource = null;
            cmbSecondField.DataSource = null;
            cmbThirdField.DataSource = null;
            UpdateFilenamePreview();
            UpdateUiEnabledState();
            return;
        }

        _headerInfo = headerInfo!;

        var headers = _headerInfo.Headers.ToList();
        cmbFirstField.DataSource = headers.ToList();

        var secondOptions = new List<string>(capacity: headers.Count + 1) { OptionalFieldNoneOption };
        secondOptions.AddRange(headers);
        cmbSecondField.DataSource = secondOptions;

        var thirdOptions = new List<string>(capacity: headers.Count + 1) { OptionalFieldNoneOption };
        thirdOptions.AddRange(headers);
        cmbThirdField.DataSource = thirdOptions;

        if (headers.Count > 0)
        {
            cmbFirstField.SelectedIndex = 0;
            cmbSecondField.SelectedIndex = 0; // none by default or maybe headers.Count > 1 ? 2 : 0
            cmbThirdField.SelectedIndex = 0; // none by default
        }

        // Because headers changed, force re-scan/remap
        _templateFields.Clear();
        _templateToExcelHeaderMap = new Dictionary<string, string>(StringComparer.Ordinal);

        AppendLog(
            $"Encabezados cargados. Cantidad: {_headerInfo.Headers.Count}. Fila de encabezado: {_headerInfo.HeaderRowNumber}");
        UpdateFilenamePreview();
        UpdateUiEnabledState();
    }

    private void UpdateFilenamePreview()
    {
        if (_headerInfo == null ||
            cmbFirstField.SelectedItem == null)
        {
            lblFilenamePreview.Text = "";
            return;
        }

        var fakeRecord = _headerInfo.Headers
            .Where(h => !string.IsNullOrWhiteSpace(h))
            .ToDictionary(h => h, h => $"<{h}>", StringComparer.Ordinal);

        var secondHeader = (string?)cmbSecondField.SelectedItem;
        if (string.Equals(secondHeader, OptionalFieldNoneOption, StringComparison.Ordinal))
            secondHeader = null;

        var thirdHeader = (string?)cmbThirdField.SelectedItem;
        if (string.Equals(thirdHeader, OptionalFieldNoneOption, StringComparison.Ordinal))
            thirdHeader = null;

        var previewPath = PdfFilenameBuilder.BuildPdfPath(
            outputDirectory: Directory.Exists(txtOutputDir.Text)
                ? txtOutputDir.Text
                : Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            record: fakeRecord,
            prefix: txtPrefix.Text,
            firstFieldHeader: (string)cmbFirstField.SelectedItem,
            secondFieldHeader: secondHeader,
            thirdFieldHeader: thirdHeader,
            postfix: txtPostfix.Text,
            separator: txtSeparator.Text,
            emptyFallback: "fila_###");

        lblFilenamePreview.Text = Path.GetFileName(previewPath);
    }

    private void ScanTemplateAndValidate()
    {
        try
        {
            if (!TryScanTemplateAndValidate(out var msg))
            {
                MessageBox.Show(msg, "Validación de plantilla", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AppendLog("Validación de plantilla: ERROR -> " + msg.Replace("\n", " "));
                return;
            }

            AppendLog($"Campos de plantilla encontrados: {_templateFields.Count}");
            AppendLog("Campos de plantilla validados contra encabezados de Excel: OK");
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Error al escanear plantilla", MessageBoxButtons.OK, MessageBoxIcon.Error);
            _templateFields.Clear();
        }
        finally
        {
            UpdateUiEnabledState();
        }
    }

    private async Task RunMergeAsync()
    {
        if (_excel == null || _headerInfo == null)
        {
            MessageBox.Show("Por favor, cargue primero el Excel y los encabezados.", "Ejecutar", MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        // Auto-scan + validate (also builds _templateFields and _templateToExcelHeaderMap)
        if (!TryScanTemplateAndValidate(out var validationMsg))
        {
            MessageBox.Show(validationMsg, "Validación de plantilla", MessageBoxButtons.OK, MessageBoxIcon.Error);
            AppendLog("Validación de plantilla: ERROR -> " + validationMsg.Replace("\n", " "));
            UpdateUiEnabledState();
            return;
        }

        var templatePath = txtTemplatePath.Text;
        var outputDir = txtOutputDir.Text;
        var worksheetName = (string)cmbWorksheet.SelectedItem!;
        var headerRow = (int)numHeaderRow.Value;
        var fieldXHeader = (string)cmbFirstField.SelectedItem!;

        var secondHeader = (string?)cmbSecondField.SelectedItem;
        if (string.Equals(secondHeader, OptionalFieldNoneOption, StringComparison.Ordinal))
            secondHeader = null;

        var thirdHeader = (string?)cmbThirdField.SelectedItem;
        if (string.Equals(thirdHeader, OptionalFieldNoneOption, StringComparison.Ordinal))
            thirdHeader = null;

        var prefixText = txtPrefix.Text;
        var postfixText = txtPostfix.Text;
        var separatorText = txtSeparator.Text;

        if (!File.Exists(templatePath))
        {
            MessageBox.Show("Por favor, seleccione una plantilla de Word válida (.docx).", "Ejecutar",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!Directory.Exists(outputDir))
        {
            MessageBox.Show("Por favor, seleccione una carpeta de salida válida.", "Ejecutar", MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        // Disable UI during run
        btnRun.Enabled = false;
        btnCancel.Enabled = true;
        btnLoadExcel.Enabled = false;
        btnScanTemplateFields.Enabled = false;
        btnBrowseExcel.Enabled = false;
        btnBrowseTemplate.Enabled = false;
        btnBrowseOutputDir.Enabled = false;

        progressBar.Value = 0;

        // Total is an estimate; skipFullyEmptyRows may reduce the real count.
        var totalEstimate = Math.Max(1, _headerInfo.LastDataRowNumber - headerRow);
        progressBar.Maximum = totalEstimate;

        _cts = new CancellationTokenSource();
        AppendLog("Iniciando combinación...");

        var progress = new Progress<MergeProgress>(p =>
        {
            progressBar.Maximum = Math.Max(1, p.Total);
            progressBar.Value = Math.Min(progressBar.Maximum, Math.Max(0, p.Done));

            if (!string.IsNullOrWhiteSpace(p.Message))
                AppendLog(p.Message);
        });

        try
        {
            await StaTaskRunner.Run(() =>
            {
                using var exporter = new WordPdfExporter();

                int done = 0, ok = 0, fail = 0;

                foreach (var (rowNumber, excelRecord) in _excel.EnumerateRecordsFormattedWithRowNumber(
                             worksheetName: worksheetName,
                             headerRowNumber: headerRow,
                             headers: _headerInfo.Headers,
                             skipFullyEmptyRows: true))
                {
                    _cts!.Token.ThrowIfCancellationRequested();
                    done++;

                    try
                    {
                        // PDF name uses the ORIGINAL Excel headers user selected
                        var pdfPath = PdfFilenameBuilder.BuildPdfPath(
                            outputDirectory: outputDir,
                            record: excelRecord,
                            prefix: prefixText,
                            firstFieldHeader: fieldXHeader,
                            secondFieldHeader: secondHeader,
                            thirdFieldHeader: thirdHeader,
                            postfix: postfixText,
                            separator: separatorText,
                            emptyFallback: $"row_{rowNumber}");

                        // Overwrite behavior
                        if (File.Exists(pdfPath))
                            File.Delete(pdfPath);

                        // Build values keyed by TEMPLATE fields (supports spaces in Excel headers)
                        var templateValues = HeaderFieldMapper.BuildTemplateValuesForRecord(
                            templateFields: _templateFields,
                            excelRecord: excelRecord,
                            templateToExcelHeaderMap: _templateToExcelHeaderMap);

                        // Merge + Export
                        MailMergeToPdf.MergeAndExportPdfForRecord(
                            record: templateValues,
                            pdfPath: pdfPath,
                            templateDocxPath: templatePath,
                            wordExporter: exporter);

                        ok++;
                        ((IProgress<MergeProgress>)progress).Report(
                            new MergeProgress(done, totalEstimate,
                                $"Fila {rowNumber}: OK -> {Path.GetFileName(pdfPath)}"));
                    }
                    catch (Exception ex)
                    {
                        fail++;
                        ((IProgress<MergeProgress>)progress).Report(
                            new MergeProgress(done, totalEstimate, $"Fila {rowNumber}: ERROR -> {ex.Message}"));
                        // continue
                    }

                    ((IProgress<MergeProgress>)progress).Report(new MergeProgress(done, totalEstimate, null));
                }

                ((IProgress<MergeProgress>)progress).Report(
                    new MergeProgress(totalEstimate, totalEstimate, $"Completado. OK: {ok}, Errores: {fail}."));

                // Show alert after process is completed
                Invoke(() => MessageBox.Show($"Proceso finalizado.\nOK: {ok}\nErrores: {fail}",
                    "Combinación completada", MessageBoxButtons.OK, MessageBoxIcon.Information));
            }, _cts.Token);
        }
        catch (OperationCanceledException)
        {
            AppendLog("Cancelado.");
        }
        catch (Exception ex)
        {
            AppendLog("Error en la ejecución: " + ex.Message);
            MessageBox.Show(ex.Message, "Error en la ejecución", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            _cts?.Dispose();
            _cts = null;

            btnCancel.Enabled = false;
            btnLoadExcel.Enabled = true;
            btnScanTemplateFields.Enabled = true;
            btnBrowseExcel.Enabled = true;
            btnBrowseTemplate.Enabled = true;
            btnBrowseOutputDir.Enabled = true;

            UpdateUiEnabledState();
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

        var templatePath = txtTemplatePath.Text;
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
            // NEW: build mapping that supports spaces in Excel headers
            _templateToExcelHeaderMap = HeaderFieldMapper.BuildTemplateToExcelHeaderMap(
                templateFields: _templateFields,
                excelHeaders: _headerInfo.Headers.ToList());

            var missing = _templateToExcelHeaderMap
                .Where(kvp => string.IsNullOrEmpty(kvp.Value))
                .Select(kvp => kvp.Key)
                .OrderBy(x => x, StringComparer.Ordinal)
                .ToList();

            if (missing.Count > 0)
            {
                userMessage =
                    "La plantilla de Word contiene campos de combinación que faltan en la fila de encabezado de Excel " +
                    "(los espacios y guiones bajos se tratan como equivalentes, pero las mayúsculas/minúsculas deben coincidir):\n\n" +
                    string.Join("\n", missing.Select(m => $" - {m}")) +
                    "\n\nSolución: Agregue estas columnas a la fila de encabezado de Excel (se requiere la ortografía exacta).";
                _templateFields.Clear();
                _templateToExcelHeaderMap = new Dictionary<string, string>(StringComparer.Ordinal);
                return false;
            }
        }
        catch (Exception ex)
        {
            // Ambiguous matches or other mapping issues
            userMessage = ex.Message;
            _templateFields.Clear();
            _templateToExcelHeaderMap = new Dictionary<string, string>(StringComparer.Ordinal);
            return false;
        }

        return true;
    }
}