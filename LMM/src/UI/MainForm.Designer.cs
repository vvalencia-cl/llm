namespace LMM.UI;

using LMM.UI.Controls;

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
        Label divisory1;
        Label divisory2;
        FlowLayoutPanel actionPanel;
        Label separatorLabel;
        ToolTip tt;

        SuspendLayout();

        // 
        // root
        // 
        root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(12) };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 14));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        // 
        // inputs
        // 
        inputs = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 2, RowCount = 8 };
        inputs.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
        inputs.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        root.Controls.Add(inputs, 0, 0);

        int r = 0;

        // Row: Plantilla Word
        _fpTemplate = new LMM.UI.Controls.FilePickerControl { LabelText = "Plantilla Word", PlaceholderText = "Seleccione la plantilla de Word (.docx)...", Dock = DockStyle.Fill };
        inputs.Controls.Add(_fpTemplate, 0, r);
        inputs.SetColumnSpan(_fpTemplate, 2);
        r++;

        // Row: Archivo Excel
        _fpExcel = new LMM.UI.Controls.FilePickerControl { LabelText = "Archivo Excel", PlaceholderText = "Seleccione el origen de datos Excel (.xlsx)...", Dock = DockStyle.Fill };
        inputs.Controls.Add(_fpExcel, 0, r);
        inputs.SetColumnSpan(_fpExcel, 2);
        r++;

        // Row: Carpeta de Salida
        _fpOutputDir = new LMM.UI.Controls.FilePickerControl { LabelText = "Carpeta de Salida", PlaceholderText = "Seleccione la carpeta donde se guardarán los archivos...", Dock = DockStyle.Fill };
        inputs.Controls.Add(_fpOutputDir, 0, r);
        inputs.SetColumnSpan(_fpOutputDir, 2);
        r++;

        // Row: Borrar contenido
        _chkClearOutputDir = new CheckBox { Text = "Borrar contenido de Carpeta de Salida", AutoSize = true, Checked = false };
        inputs.Controls.Add(_chkClearOutputDir, 1, r);
        r++;

        // Row: Analizar
        _btnLoadExcel = new Button { Text = "Cargar Datos", AutoSize = true };
        inputs.Controls.Add(new Label { Text = "Analizar", AutoSize = true, Anchor = AnchorStyles.Left }, 0, r);
        inputs.Controls.Add(_btnLoadExcel, 1, r);
        r++;

        divisory1 = new Label { Text = "", Height = 1, BackColor = Color.Black, Dock = DockStyle.Fill };
        inputs.Controls.Add(divisory1, 0, r);
        inputs.SetColumnSpan(divisory1, 2);
        r++;

        // Row: Opciones de Excel
        _excelOptions = new LMM.UI.Controls.ExcelOptionsControl { Dock = DockStyle.Fill };
        inputs.Controls.Add(new Label { Text = "Opciones de Excel", AutoSize = true, Anchor = AnchorStyles.Left }, 0, r);
        inputs.Controls.Add(_excelOptions, 1, r);
        r++;

        divisory2 = new Label { Text = "", Height = 1, BackColor = Color.Black, Dock = DockStyle.Fill };
        inputs.Controls.Add(divisory2, 0, r);
        inputs.SetColumnSpan(divisory2, 2);
        r++;

        // Row: Nombre de archivo
        _filenameBuilder = new LMM.UI.Controls.FilenameBuilderControl { Dock = DockStyle.Fill };
        inputs.Controls.Add(new Label { Text = "Nombre de archivo", AutoSize = true, Anchor = AnchorStyles.Left }, 0, r);
        inputs.Controls.Add(_filenameBuilder, 1, r);
        r++;

        // Row: Acciones
        _btnScan = new Button { Text = "Validar campos plantilla", AutoSize = true };
        _btnRun = new Button { Text = "Ejecutar combinación", AutoSize = true };
        _btnCancel = new Button { Text = "Cancelar", Enabled = false, AutoSize = true };
        _btnOpenOutputDir = new Button { Text = "Abrir carpeta de Salida", Enabled = false, AutoSize = true };

        actionPanel = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Dock = DockStyle.Fill };
        actionPanel.Controls.Add(_btnScan);
        actionPanel.Controls.Add(_btnRun);
        actionPanel.Controls.Add(_btnCancel);
        actionPanel.Controls.Add(_btnOpenOutputDir);

        inputs.Controls.Add(new Label { Text = "Acciones", AutoSize = true, Anchor = AnchorStyles.Left }, 0, r);
        inputs.Controls.Add(actionPanel, 1, r);
        r++;

        // Separator
        separatorLabel = new Label { Height = 2, BorderStyle = BorderStyle.Fixed3D, Dock = DockStyle.Fill };
        root.Controls.Add(separatorLabel, 0, 1);

        // Bottom
        _logPanel = new LMM.UI.Controls.LogPanelControl { Dock = DockStyle.Fill };
        root.Controls.Add(_logPanel, 0, 2);

        // ToolTips
        tt = new ToolTip();
        // tt.SetToolTip(_filenameBuilder.CmbFirst, "Primer Campo..."); // Podríamos exponer controles si quisiéramos ToolTips exactos

        // MainForm
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
        actionPanel.ResumeLayout(false);
        actionPanel.PerformLayout();
        ResumeLayout(false);
    }

    #endregion
}
