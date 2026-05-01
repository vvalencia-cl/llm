namespace LMM.UI;

partial class MainForm
{
    /// <summary>
    /// Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    /// Clean up any resources being used.
    /// </summary>
    /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    /// <summary>
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
        TableLayoutPanel root;
        TableLayoutPanel inputs;
        Label labelTemplate;
        Label labelExcel;
        Label labelAnalizar;
        Label divisory1;
        FlowLayoutPanel excelOptionsPanel;
        Label labelHoja;
        Label labelFilaEncabezados;
        Label divisory2;
        TableLayoutPanel fieldPanel;
        Label labelFilename;
        Label labelFilenamePreviewTitle;
        Label labelOutputDir;
        FlowLayoutPanel actionPanel;
        Label separatorLabel;
        TableLayoutPanel bottom;
        Label bottomGap;
        ContextMenuStrip logMenu;
        ToolStripMenuItem miCopy;
        ToolTip tt;

        SuspendLayout();

        // 
        // root
        // 
        root = new TableLayoutPanel();
        root.Dock = DockStyle.Fill;
        root.ColumnCount = 1;
        root.RowCount = 3;
        root.Padding = new Padding(12);
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 14));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        // 
        // inputs
        // 
        inputs = new TableLayoutPanel();
        inputs.Dock = DockStyle.Top;
        inputs.AutoSize = true;
        inputs.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        inputs.ColumnCount = 4;
        inputs.RowCount = 10;
        inputs.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
        inputs.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        inputs.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
        inputs.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));

        root.Controls.Add(inputs, 0, 0);

        int r = 0;

        // --- Row: Plantilla Word ---
        labelTemplate = new Label { Text = "Plantilla Word", AutoSize = true, Anchor = AnchorStyles.Left, Padding = new Padding(0, 6, 0, 0) };
        txtTemplatePath = new TextBox { PlaceholderText = "Seleccione la plantilla de Word (.docx)...", Dock = DockStyle.Fill };
        btnBrowseTemplate = new Button { Text = "Buscar...", AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Right };
        
        inputs.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        inputs.Controls.Add(labelTemplate, 0, r);
        inputs.Controls.Add(txtTemplatePath, 1, r);
        inputs.Controls.Add(btnBrowseTemplate, 2, r);
        r++;

        // --- Row: Archivo Excel ---
        labelExcel = new Label { Text = "Archivo Excel", AutoSize = true, Anchor = AnchorStyles.Left, Padding = new Padding(0, 6, 0, 0) };
        txtExcelPath = new TextBox { PlaceholderText = "Seleccione el origen de datos Excel (.xlsx)...", Dock = DockStyle.Fill };
        btnBrowseExcel = new Button { Text = "Buscar...", AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Right };
        btnLoadExcel = new Button { Text = "Cargar Datos", AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Right };

        inputs.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        inputs.Controls.Add(labelExcel, 0, r);
        inputs.Controls.Add(txtExcelPath, 1, r);
        inputs.Controls.Add(btnBrowseExcel, 2, r);
        inputs.Controls.Add(btnLoadExcel, 3, r);
        r++;

        // --- Row: Analizar ---
        labelAnalizar = new Label { Text = "Analizar", AutoSize = true, Anchor = AnchorStyles.Left, Padding = new Padding(0, 6, 0, 0) };
        // btnLoadExcel already created
        
        inputs.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        inputs.Controls.Add(labelAnalizar, 0, r);
        // Note: The original code does AddRow("Analizar", btnLoadExcel) where btnLoadExcel was already used in the previous row. 
        // WinForms controls can only have one parent/position. 
        // Looking at the original BuildUi: 
        // 115: AddRow("Archivo Excel", txtExcelPath, btnBrowseExcel, btnLoadExcel);
        // 118: btnLoadExcel = new Button { Text = "Cargar Datos" };
        // 119: AddRow("Analizar", btnLoadExcel);
        // Wait, line 117 overwrites btnLoadExcel variable! 
        // 114: private Button btnLoadExcel = null!;
        // 115: AddRow("Archivo Excel", txtExcelPath, btnBrowseExcel, btnLoadExcel); // Here it is null!
        // 117: btnLoadExcel = new Button { Text = "Cargar Datos" };
        // 118: AddRow("Analizar", btnLoadExcel);
        // So in the first row it was null (and thus not added). I will follow the behavior.

        inputs.Controls.Add(btnLoadExcel, 1, r);
        r++;

        // --- Divisory ---
        divisory1 = new Label { Text = "", Height = 1, BackColor = Color.Black, Dock = DockStyle.Fill };
        inputs.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        inputs.Controls.Add(divisory1, 0, r);
        inputs.SetColumnSpan(divisory1, 3);
        r++;

        // --- Row: Opciones de Excel ---
        cmbWorksheet = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        numHeaderRow = new NumericUpDown { Minimum = 1, Maximum = 100000, Value = 1 };
        btnRefreshHeaders = new Button { Text = "Refrescar columnas de Excel" };

        excelOptionsPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
        excelOptionsPanel.Name = "Opciones de Excel";
        labelHoja = new Label { Text = "Hoja de Excel: ", AutoSize = true };
        labelFilaEncabezados = new Label { Text = "Nº Fila con Encabezados: ", AutoSize = true };
        excelOptionsPanel.Controls.Add(labelHoja);
        excelOptionsPanel.Controls.Add(cmbWorksheet);
        excelOptionsPanel.Controls.Add(labelFilaEncabezados);
        excelOptionsPanel.Controls.Add(numHeaderRow);
        excelOptionsPanel.Controls.Add(btnRefreshHeaders);

        labelHoja = new Label { Text = "Opciones de Excel", AutoSize = true, Anchor = AnchorStyles.Left, Padding = new Padding(0, 6, 0, 0) };
        inputs.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        inputs.Controls.Add(labelHoja, 0, r);
        inputs.Controls.Add(excelOptionsPanel, 1, r);
        r++;

        // --- Divisory ---
        divisory2 = new Label { Text = "", Height = 1, BackColor = Color.Black, Dock = DockStyle.Fill };
        inputs.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        inputs.Controls.Add(divisory2, 0, r);
        inputs.SetColumnSpan(divisory2, 3);
        r++;

        // --- Row: Nombre de archivo ---
        cmbFirstField = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        cmbSecondField = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        cmbThirdField = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        txtPrefix = new TextBox { PlaceholderText = "Prefijo" };
        txtPostfix = new TextBox { PlaceholderText = "Sufijo" };
        txtSeparator = new TextBox { PlaceholderText = "Sep.", Text = "_" };

        fieldPanel = new TableLayoutPanel { ColumnCount = 6, Dock = DockStyle.Fill, AutoSize = true };
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

        labelFilename = new Label { Text = "Nombre de archivo", AutoSize = true, Anchor = AnchorStyles.Left, Padding = new Padding(0, 6, 0, 0) };
        inputs.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        inputs.Controls.Add(labelFilename, 0, r);
        inputs.Controls.Add(fieldPanel, 1, r);
        r++;

        // --- Row: Nombre de archivo a generar ---
        lblFilenamePreview = new Label { AutoSize = true, Text = "", Dock = DockStyle.Fill };
        labelFilenamePreviewTitle = new Label { Text = "Nombre de archivo a generar", AutoSize = true, Anchor = AnchorStyles.Left, Padding = new Padding(0, 6, 0, 0) };
        inputs.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        inputs.Controls.Add(labelFilenamePreviewTitle, 0, r);
        inputs.Controls.Add(lblFilenamePreview, 1, r);
        r++;

        // --- Row: Carpeta de salida ---
        txtOutputDir = new TextBox { PlaceholderText = "Seleccione la carpeta de salida...", Dock = DockStyle.Fill };
        btnBrowseOutputDir = new Button { Text = "Buscar...", AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Right };
        labelOutputDir = new Label { Text = "Carpeta de salida", AutoSize = true, Anchor = AnchorStyles.Left, Padding = new Padding(0, 6, 0, 0) };
        inputs.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        inputs.Controls.Add(labelOutputDir, 0, r);
        inputs.Controls.Add(txtOutputDir, 1, r);
        inputs.Controls.Add(btnBrowseOutputDir, 2, r);
        r++;

        // --- Row: Acciones ---
        btnScanTemplateFields = new Button { Text = "Validar campos plantilla", AutoSize = true };
        btnRun = new Button { Text = "Ejecutar combinación", AutoSize = true };
        btnCancel = new Button { Text = "Cancelar", Enabled = false, AutoSize = true };

        actionPanel = new FlowLayoutPanel
        {
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            WrapContents = true,
            FlowDirection = FlowDirection.LeftToRight,
            Margin = Padding.Empty,
            Padding = Padding.Empty,
            Dock = DockStyle.Fill
        };
        actionPanel.Controls.Add(btnScanTemplateFields);
        actionPanel.Controls.Add(btnRun);
        actionPanel.Controls.Add(btnCancel);

        labelHoja = new Label { Text = "Acciones", AutoSize = true, Anchor = AnchorStyles.Left, Padding = new Padding(0, 6, 0, 0) };
        inputs.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        inputs.Controls.Add(labelHoja, 0, r);
        inputs.Controls.Add(actionPanel, 1, r);
        r++;

        // --- Separator between inputs and bottom ---
        separatorLabel = new Label { Height = 2, BorderStyle = BorderStyle.Fixed3D, Dock = DockStyle.Fill };
        root.Controls.Add(separatorLabel, 0, 1);

        // --- Bottom panel ---
        bottom = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3 };
        bottom.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        bottom.RowStyles.Add(new RowStyle(SizeType.Absolute, 8));
        bottom.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.Controls.Add(bottom, 0, 2);

        progressBar = new ProgressBar { Dock = DockStyle.Top, Height = 18 };
        bottom.Controls.Add(progressBar, 0, 0);

        bottomGap = new Label { Height = 2 };
        bottom.Controls.Add(bottomGap, 0, 1);

        lstLog = new ListBox { Dock = DockStyle.Fill, SelectionMode = SelectionMode.MultiExtended };
        bottom.Controls.Add(lstLog, 0, 2);

        // --- Context Menu ---
        miCopy = new ToolStripMenuItem("Copiar seleccionados");
        logMenu = new ContextMenuStrip();
        logMenu.Items.Add(miCopy);
        lstLog.ContextMenuStrip = logMenu;

        // --- ToolTips ---
        tt = new ToolTip();
        tt.SetToolTip(cmbFirstField, "Primer Campo: Columna de Excel utilizada para la primera parte del nombre del PDF");
        tt.SetToolTip(cmbSecondField, "Segundo Campo: Columna de Excel utilizada para la segunda parte del nombre del PDF");
        tt.SetToolTip(cmbThirdField, "Tercer Campo (opcional): si se elige, se agrega como tercera parte del nombre del PDF");

        // 
        // MainForm
        // 
        AutoScaleDimensions = new SizeF(96F, 96F);
        AutoScaleMode = AutoScaleMode.Dpi;
        ClientSize = new Size(980, 720);
        Controls.Add(root);
        Name = "MainForm";
        StartPosition = FormStartPosition.CenterScreen;
        Text = "Combinación de Correspondencia (Plantilla Word + Excel → PDF)";

        root.ResumeLayout(false);
        root.PerformLayout();
        inputs.ResumeLayout(false);
        inputs.PerformLayout();
        excelOptionsPanel.ResumeLayout(false);
        excelOptionsPanel.PerformLayout();
        fieldPanel.ResumeLayout(false);
        fieldPanel.PerformLayout();
        actionPanel.ResumeLayout(false);
        actionPanel.PerformLayout();
        bottom.ResumeLayout(false);
        bottom.PerformLayout();
        ResumeLayout(false);
    }

    #endregion

    private TextBox txtTemplatePath;
    private Button btnBrowseTemplate;
    private TextBox txtExcelPath;
    private Button btnBrowseExcel;
    private Button btnLoadExcel;
    private ComboBox cmbWorksheet;
    private NumericUpDown numHeaderRow;
    private Button btnRefreshHeaders;
    private ComboBox cmbFirstField;
    private ComboBox cmbSecondField;
    private ComboBox cmbThirdField;
    private TextBox txtPrefix;
    private TextBox txtPostfix;
    private TextBox txtSeparator;
    private Label lblFilenamePreview;
    private TextBox txtOutputDir;
    private Button btnBrowseOutputDir;
    private Button btnScanTemplateFields;
    private Button btnRun;
    private Button btnCancel;
    private ProgressBar progressBar;
    private ListBox lstLog;
}
