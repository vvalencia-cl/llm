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
        FlowLayoutPanel actionPanel;
        GroupBox grpInput;
        GroupBox grpOutput;
        GroupBox grpActions;
        TableLayoutPanel layoutInput;
        TableLayoutPanel layoutOutput;

        SuspendLayout();

        // 
        // root
        // 
        root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Padding = new Padding(15) };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        // 
        // inputs container
        // 
        inputs = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 1, RowCount = 3 };
        root.Controls.Add(inputs, 0, 0);

        // --- Grupo: Datos de Origen ---
        grpInput = new GroupBox { Text = "Datos de Origen", Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(10) };
        layoutInput = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 2, RowCount = 4 };
        layoutInput.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
        layoutInput.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        grpInput.Controls.Add(layoutInput);
        inputs.Controls.Add(grpInput, 0, 0);

        // Row: Plantilla Word
        _fpTemplate = new FilePickerControl { PlaceholderText = "Seleccione la plantilla de Word (.docx)...", Dock = DockStyle.Fill };
        layoutInput.Controls.Add(new Label { Text = "Plantilla Word:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
        layoutInput.Controls.Add(_fpTemplate, 1, 0);

        // Row: Archivo Excel
        _fpExcel = new FilePickerControl { PlaceholderText = "Seleccione el origen de datos Excel (.xlsx)...", Dock = DockStyle.Fill };
        layoutInput.Controls.Add(new Label { Text = "Archivo Excel:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 1);
        layoutInput.Controls.Add(_fpExcel, 1, 1);

        // Row: Cargar Datos
        _btnLoadExcel = new Button { Text = "Cargar Datos", AutoSize = true, Padding = new Padding(10, 0, 10, 0) };
        layoutInput.Controls.Add(new Label { Text = "Acción:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 2);
        layoutInput.Controls.Add(_btnLoadExcel, 1, 2);

        // Row: Opciones de Excel
        _excelOptions = new ExcelOptionsControl { Dock = DockStyle.Fill };
        layoutInput.Controls.Add(new Label { Text = "Opciones Excel:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 3);
        layoutInput.Controls.Add(_excelOptions, 1, 3);

        // --- Grupo: Configuración de Salida ---
        grpOutput = new GroupBox { Text = "Configuración de Salida", Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(10), Margin = new Padding(0, 10, 0, 0) };
        layoutOutput = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 2, RowCount = 3 };
        layoutOutput.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
        layoutOutput.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        grpOutput.Controls.Add(layoutOutput);
        inputs.Controls.Add(grpOutput, 0, 1);

        // Row: Carpeta de Salida
        _fpOutputDir = new FilePickerControl { PlaceholderText = "Seleccione la carpeta de destino...", Dock = DockStyle.Fill };
        layoutOutput.Controls.Add(new Label { Text = "Carpeta Salida:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
        layoutOutput.Controls.Add(_fpOutputDir, 1, 0);

        // Row: Borrar contenido
        _chkClearOutputDir = new CheckBox { Text = "Borrar contenido previo", AutoSize = true, Checked = false };
        layoutOutput.Controls.Add(_chkClearOutputDir, 1, 1);

        // Row: Nombre de archivo
        _filenameBuilder = new FilenameBuilderControl { Dock = DockStyle.Fill };
        layoutOutput.Controls.Add(new Label { Text = "Nombre Archivo:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 2);
        layoutOutput.Controls.Add(_filenameBuilder, 1, 2);

        // --- Grupo: Acciones ---
        grpActions = new GroupBox { Text = "Acciones", Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(10), Margin = new Padding(0, 10, 0, 0) };
        actionPanel = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, WrapContents = true };
        grpActions.Controls.Add(actionPanel);
        inputs.Controls.Add(grpActions, 0, 2);

        _btnScan = new Button { Text = "Validar Plantilla", AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
        _btnRun = new Button { Text = "Ejecutar Combinación", AutoSize = true, Padding = new Padding(10, 5, 10, 5), Font = new Font(DefaultFont, FontStyle.Bold) };
        _btnCancel = new Button { Text = "Cancelar", Enabled = false, AutoSize = true, Padding = new Padding(10, 5, 10, 5) };
        _btnOpenOutputDir = new Button { Text = "Ver Resultados", Enabled = false, AutoSize = true, Padding = new Padding(10, 5, 10, 5) };

        actionPanel.Controls.Add(_btnScan);
        actionPanel.Controls.Add(_btnRun);
        actionPanel.Controls.Add(_btnCancel);
        actionPanel.Controls.Add(_btnOpenOutputDir);

        // --- Bottom: Log ---
        _logPanel = new LogPanelControl { Dock = DockStyle.Fill, Margin = new Padding(0, 15, 0, 0) };
        root.Controls.Add(_logPanel, 0, 1);

        // MainForm
        AutoScaleDimensions = new SizeF(96F, 96F);
        AutoScaleMode = AutoScaleMode.Dpi;
        ClientSize = new Size(1000, 750);
        Controls.Add(root);
        Font = new Font("Segoe UI", 9F);
        Name = "MainForm";
        StartPosition = FormStartPosition.CenterScreen;
        Text = "Mail Merge (Plantilla Word + Excel → PDF)";
        AutoScroll = true;

        root.ResumeLayout(false);
        root.PerformLayout();
        inputs.ResumeLayout(false);
        inputs.PerformLayout();
        layoutInput.ResumeLayout(false);
        layoutInput.PerformLayout();
        layoutOutput.ResumeLayout(false);
        layoutOutput.PerformLayout();
        actionPanel.ResumeLayout(false);
        actionPanel.PerformLayout();
        ResumeLayout(false);
    }

    #endregion
}
