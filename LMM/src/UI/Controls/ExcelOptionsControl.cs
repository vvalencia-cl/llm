namespace LMM.UI.Controls;

public partial class ExcelOptionsControl : UserControl
{
    private ComboBox _cmbWorksheet = null!;
    private NumericUpDown _numHeaderRow = null!;
    private Button _btnRefresh = null!;

    public event EventHandler? OptionsChanged;
    public event EventHandler? RefreshClicked;

    public string? SelectedWorksheet
    {
        get => _cmbWorksheet.SelectedItem as string;
        set => _cmbWorksheet.SelectedItem = value;
    }

    public int HeaderRow
    {
        get => (int)_numHeaderRow.Value;
        set => _numHeaderRow.Value = value;
    }

    public ExcelOptionsControl()
    {
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        var panel = new FlowLayoutPanel 
        { 
            Dock = DockStyle.Fill, 
            AutoSize = true, 
            WrapContents = true,
            Padding = new Padding(0)
        };
        
        panel.Controls.Add(new Label { Text = "Hoja:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 0, 5, 0) });
        _cmbWorksheet = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 150 };
        panel.Controls.Add(_cmbWorksheet);

        panel.Controls.Add(new Label { Text = "Fila Encabezados:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(15, 0, 5, 0) });
        _numHeaderRow = new NumericUpDown { Minimum = 1, Maximum = 100000, Value = 1, Width = 60 };
        panel.Controls.Add(_numHeaderRow);

        _btnRefresh = new Button { Text = "Refrescar", AutoSize = true, Margin = new Padding(15, 0, 0, 0) };
        panel.Controls.Add(_btnRefresh);

        _cmbWorksheet.SelectedIndexChanged += (s, e) => OptionsChanged?.Invoke(this, EventArgs.Empty);
        _numHeaderRow.ValueChanged += (s, e) => OptionsChanged?.Invoke(this, EventArgs.Empty);
        _btnRefresh.Click += (s, e) => RefreshClicked?.Invoke(this, EventArgs.Empty);

        this.Controls.Add(panel);
        this.AutoSize = true;
    }

    public void SetWorksheets(IEnumerable<string> worksheets)
    {
        _cmbWorksheet.DataSource = worksheets.ToList();
    }

    public void SetEnabled(bool enabled)
    {
        _cmbWorksheet.Enabled = enabled;
        _numHeaderRow.Enabled = enabled;
        _btnRefresh.Enabled = enabled;
    }
}
