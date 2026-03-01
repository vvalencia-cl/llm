using LMM.Application;
using LMM.Domain.Dto;

namespace LMM.UI;

public sealed class MainForm : Form
{
    // UI
    private TextBox txtTemplatePath = null!;
    private Button btnBrowseTemplate = null!;

    private TextBox txtExcelPath = null!;
    private Button btnBrowseExcel = null!;
    private Button btnLoadExcel = null!;

    private ComboBox cmbWorksheet = null!;
    private NumericUpDown numHeaderRow = null!;
    private Button btnRefreshHeaders = null!;

    private ComboBox cmbFirstField = null!;
    private ComboBox cmbSecondField = null!;
    private ComboBox cmbThirdField = null!;

    private TextBox txtPrefix = null!;
    private TextBox txtPostfix = null!;
    private TextBox txtSeparator = null!;

    private Label lblFilenamePreview = null!;

    private TextBox txtOutputDir = null!;
    private Button btnBrowseOutputDir = null!;

    private Button btnScanTemplateFields = null!;
    private Button btnRun = null!;
    private Button btnCancel = null!;

    private ProgressBar progressBar = null!;
    private ListBox lstLog = null!;

    // State
    private ClosedXmlDataSourceService? _excel;
    private ExcelHeaderResult? _headerInfo;
    private List<string> _templateFields = new();
    private Dictionary<string, string> _templateToExcelHeaderMap = new(StringComparer.Ordinal);
    private CancellationTokenSource? _cts;

    private const string OptionalFieldNoneOption = "(Ninguno)";

    public MainForm()
    {
        Text = "Combinación de Correspondencia (Plantilla Word + Excel → PDF)";
        Width = 980;
        Height = 720;
        StartPosition = FormStartPosition.CenterScreen;

        // Scale layout based on DPI (helps consistency at 150%)
        AutoScaleMode = AutoScaleMode.Dpi;

        BuildUi();
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

    private void BuildUi()
    {
        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3,
            Padding = new Padding(12),
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 14));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        Controls.Add(root);

        // Top panel (inputs)
        var inputs = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            ColumnCount = 4,
            RowCount = 10,
        };
        inputs.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
        inputs.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        inputs.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
        inputs.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));

        root.Controls.Add(inputs, 0, 0);

        // Row helper
        var r = 0;

        // Template
        txtTemplatePath = new TextBox { PlaceholderText = "Seleccione la plantilla de Word (.docx)..." };
        btnBrowseTemplate = new Button { Text = "Buscar..." };
        AddRow("Plantilla Word", txtTemplatePath, btnBrowseTemplate);

        // Excel
        txtExcelPath = new TextBox { PlaceholderText = "Seleccione el origen de datos Excel (.xlsx)..." };
        btnBrowseExcel = new Button { Text = "Buscar..." };

        AddRow("Archivo Excel", txtExcelPath, btnBrowseExcel, btnLoadExcel);

        btnLoadExcel = new Button { Text = "Cargar Datos" };
        AddRow("Analizar", btnLoadExcel);

        AddDivisoryLine();

        // Worksheet + header row
        cmbWorksheet = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        numHeaderRow = new NumericUpDown { Minimum = 1, Maximum = 100000, Value = 1 };
        btnRefreshHeaders = new Button { Text = "Refrescar columnas de Excel" };

        var excelOptionsPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
        excelOptionsPanel.Name = "Opciones de Excel";
        excelOptionsPanel.Controls.Add(new Label { Text = "Hoja de Excel: ", AutoSize = true });
        excelOptionsPanel.Controls.Add(cmbWorksheet);
        excelOptionsPanel.Controls.Add(new Label { Text = "Nº Fila con Encabezados: ", AutoSize = true });
        excelOptionsPanel.Controls.Add(numHeaderRow);
        excelOptionsPanel.Controls.Add(btnRefreshHeaders);

        AddRow("Opciones de Excel", excelOptionsPanel);

        AddDivisoryLine();

        cmbFirstField = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        cmbSecondField = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        cmbThirdField = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };

        txtPrefix = new TextBox { PlaceholderText = "Prefijo" };
        txtPostfix = new TextBox { PlaceholderText = "Sufijo" };
        txtSeparator = new TextBox { PlaceholderText = "Sep.", Text = "_" };

        var fieldPanel = new TableLayoutPanel { ColumnCount = 6, Dock = DockStyle.Fill, AutoSize = true };
        fieldPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 15)); // prefix
        fieldPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25)); // first
        fieldPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10)); // separator
        fieldPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20)); // second
        fieldPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20)); // third
        fieldPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10)); // postfix

        fieldPanel.Controls.Add(txtPrefix, 0, 0);
        fieldPanel.Controls.Add(cmbFirstField, 1, 0);
        fieldPanel.Controls.Add(txtSeparator, 2, 0);
        fieldPanel.Controls.Add(cmbSecondField, 3, 0);
        fieldPanel.Controls.Add(cmbThirdField, 4, 0);
        fieldPanel.Controls.Add(txtPostfix, 5, 0);

        AddRow("Nombre de archivo", fieldPanel);

        // Filename preview
        lblFilenamePreview = new Label { AutoSize = true, Text = "" };
        AddRow("Nombre de archivo a generar", lblFilenamePreview);

        // Output directory
        txtOutputDir = new TextBox { PlaceholderText = "Seleccione la carpeta de salida..." };
        btnBrowseOutputDir = new Button { Text = "Buscar..." };
        AddRow("Carpeta de salida", txtOutputDir, btnBrowseOutputDir);

        // Actions row
        btnScanTemplateFields = new Button { Text = "Validar campos plantilla" };
        btnRun = new Button { Text = "Ejecutar combinación" };
        btnCancel = new Button { Text = "Cancelar", Enabled = false };

        // Help DPI/layout: buttons size themselves correctly
        btnScanTemplateFields.AutoSize = true;
        btnRun.AutoSize = true;
        btnCancel.AutoSize = true;

        var actionPanel = new FlowLayoutPanel
        {
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            WrapContents = true, // key: at 150% it may need to wrap
            FlowDirection = FlowDirection.LeftToRight,
            Margin = Padding.Empty,
            Padding = Padding.Empty
        };
        actionPanel.Controls.Add(btnScanTemplateFields);
        actionPanel.Controls.Add(btnRun);
        actionPanel.Controls.Add(btnCancel);

        AddRow("Acciones", actionPanel);

        // Separator row
        root.Controls.Add(new Label { Height = 2, BorderStyle = BorderStyle.Fixed3D, Dock = DockStyle.Fill }, 0, 1);

        // Bottom panel (progress + log)
        var bottom = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3
        };
        bottom.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        bottom.RowStyles.Add(new RowStyle(SizeType.Absolute, 8));
        bottom.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        root.Controls.Add(bottom, 0, 2);

        progressBar = new ProgressBar { Dock = DockStyle.Top, Height = 18 };
        bottom.Controls.Add(progressBar, 0, 0);

        bottom.Controls.Add(new Label { Height = 2 }, 0, 1);

        lstLog = new ListBox { Dock = DockStyle.Fill };
        lstLog.SelectionMode = SelectionMode.MultiExtended;
        bottom.Controls.Add(lstLog, 0, 2);

        // Add a context menu for copying selected lines
        var logMenu = new ContextMenuStrip();
        var miCopy = new ToolStripMenuItem("Copiar seleccionados");
        miCopy.Click += (_, __) => CopySelectedLogLinesToClipboard();
        logMenu.Items.Add(miCopy);

        // Optional: enable/disable based on selection
        logMenu.Opening += (_, __) => miCopy.Enabled = lstLog.SelectedItems.Count > 0;

        lstLog.ContextMenuStrip = logMenu;

        // Cosmético: etiquetar los desplegables mediante ToolTips
        var tt = new ToolTip();
        tt.SetToolTip(cmbFirstField,
            "Primer Campo: Columna de Excel utilizada para la primera parte del nombre del PDF");
        tt.SetToolTip(cmbSecondField,
            "Segundo Campo: Columna de Excel utilizada para la segunda parte del nombre del PDF");
        tt.SetToolTip(cmbThirdField,
            "Tercer Campo (opcional): si se elige, se agrega como tercera parte del nombre del PDF");
        return;

        void AddDivisoryLine()
        {
            var divisoryLabel = new Label { Text = "", Height = 1, BackColor = Color.Black, Dock = DockStyle.Fill };
            inputs.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            inputs.Controls.Add(divisoryLabel);
            inputs.SetCellPosition(divisoryLabel, new TableLayoutPanelCellPosition(0, r));
            inputs.SetColumnSpan(divisoryLabel, 3);
            r++;
        }

        void AddRow(string label, Control main, Control? b1 = null, Control? b2 = null)
        {
            inputs.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            inputs.Controls.Add(
                new Label
                {
                    Text = label,
                    AutoSize = true,
                    Anchor = AnchorStyles.Left,
                    Padding = new Padding(0, 6, 0, 0)
                },
                0, r);

            // DPI-friendly layout:
            // - Inputs: Fill
            // - Buttons: AutoSize (avoid clipped height)
            // - FlowLayoutPanel: Dock=Fill so it gets full width (and can wrap); keep AutoSize for height
            if (main is FlowLayoutPanel flp)
            {
                flp.AutoSize = true;
                flp.AutoSizeMode = AutoSizeMode.GrowAndShrink;
                flp.Dock = DockStyle.Fill;
            }
            else if (main is Button)
            {
                main.AutoSize = true;
                main.Anchor = AnchorStyles.Left | AnchorStyles.Right;
            }
            else
            {
                main.Dock = DockStyle.Fill;
            }

            inputs.Controls.Add(main, 1, r);

            if (b1 != null)
            {
                if (b1 is Button)
                {
                    b1.AutoSize = true;
                    b1.Anchor = AnchorStyles.Left | AnchorStyles.Right;
                }
                else
                {
                    b1.Dock = DockStyle.Fill;
                }

                inputs.Controls.Add(b1, 2, r);
            }

            if (b2 != null)
            {
                if (b2 is Button)
                {
                    b2.AutoSize = true;
                    b2.Anchor = AnchorStyles.Left | AnchorStyles.Right;
                }
                else
                {
                    b2.Dock = DockStyle.Fill;
                }

                inputs.Controls.Add(b2, 3, r);
            }

            r++;
        }
    }

    private void WireEvents()
    {
        btnBrowseTemplate.Click += (_, __) => BrowseTemplate();
        btnBrowseExcel.Click += (_, __) => BrowseExcel();
        btnBrowseOutputDir.Click += (_, __) => BrowseOutputDir();

        btnLoadExcel.Click += async (_, __) =>
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
        };

        btnRefreshHeaders.Click += (_, __) => RefreshHeaders();

        cmbWorksheet.SelectedIndexChanged += (_, __) => RefreshHeaders();
        numHeaderRow.ValueChanged += (_, __) => RefreshHeaders();

        cmbFirstField.SelectedIndexChanged += (_, __) => UpdateFilenamePreview();
        cmbSecondField.SelectedIndexChanged += (_, __) => UpdateFilenamePreview();
        cmbThirdField.SelectedIndexChanged += (_, __) => UpdateFilenamePreview();
        txtPrefix.TextChanged += (_, __) => UpdateFilenamePreview();
        txtPostfix.TextChanged += (_, __) => UpdateFilenamePreview();
        txtSeparator.TextChanged += (_, __) => UpdateFilenamePreview();
        txtOutputDir.TextChanged += (_, __) => UpdateFilenamePreview();

        btnScanTemplateFields.Click += (_, __) => ScanTemplateAndValidate();
        btnRun.Click += async (_, __) => await RunMergeAsync();
        btnCancel.Click += (_, __) => _cts?.Cancel();

        // Allow Ctrl+C to copy selected log lines
        lstLog.KeyDown += (_, e) =>
        {
            if (e.Control && e.KeyCode == Keys.C)
            {
                CopySelectedLogLinesToClipboard();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        };
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