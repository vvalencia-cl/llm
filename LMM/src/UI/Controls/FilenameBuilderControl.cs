namespace LMM.UI.Controls;

public partial class FilenameBuilderControl : UserControl
{
    private ComboBox _cmbFirst = null!;
    private ComboBox _cmbSecond = null!;
    private ComboBox _cmbThird = null!;
    private TextBox _txtPrefix = null!;
    private TextBox _txtPostfix = null!;
    private TextBox _txtSeparator = null!;
    private Label _lblPreview = null!;

    public event EventHandler? ConfigChanged;

    public string Prefix => _txtPrefix.Text;
    public string Postfix => _txtPostfix.Text;
    public string Separator => _txtSeparator.Text;
    public string? FirstField => _cmbFirst.SelectedItem as string;
    public string? SecondField => _cmbSecond.SelectedItem as string;
    public string? ThirdField => _cmbThird.SelectedItem as string;

    public FilenameBuilderControl()
    {
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        var mainPanel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, AutoSize = true };
        
        var fieldPanel = new TableLayoutPanel { ColumnCount = 6, Dock = DockStyle.Fill, AutoSize = true };
        fieldPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 15)); 
        fieldPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25)); 
        fieldPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10)); 
        fieldPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20)); 
        fieldPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20)); 
        fieldPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10));

        _txtPrefix = new TextBox { PlaceholderText = "Prefijo", Dock = DockStyle.Fill };
        _cmbFirst = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill };
        _txtSeparator = new TextBox { PlaceholderText = "Sep.", Text = "_", Dock = DockStyle.Fill };
        _cmbSecond = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill };
        _cmbThird = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill };
        _txtPostfix = new TextBox { PlaceholderText = "Sufijo", Dock = DockStyle.Fill };

        fieldPanel.Controls.Add(_txtPrefix, 0, 0);
        fieldPanel.Controls.Add(_cmbFirst, 1, 0);
        fieldPanel.Controls.Add(_txtSeparator, 2, 0);
        fieldPanel.Controls.Add(_cmbSecond, 3, 0);
        fieldPanel.Controls.Add(_cmbThird, 4, 0);
        fieldPanel.Controls.Add(_txtPostfix, 5, 0);

        _lblPreview = new Label { Text = "", AutoSize = true, Font = new Font(this.Font, FontStyle.Italic) };
        
        mainPanel.Controls.Add(fieldPanel, 0, 0);
        mainPanel.Controls.Add(_lblPreview, 0, 1);

        _txtPrefix.TextChanged += (s, e) => ConfigChanged?.Invoke(this, EventArgs.Empty);
        _txtPostfix.TextChanged += (s, e) => ConfigChanged?.Invoke(this, EventArgs.Empty);
        _txtSeparator.TextChanged += (s, e) => ConfigChanged?.Invoke(this, EventArgs.Empty);
        _cmbFirst.SelectedIndexChanged += (s, e) => ConfigChanged?.Invoke(this, EventArgs.Empty);
        _cmbSecond.SelectedIndexChanged += (s, e) => ConfigChanged?.Invoke(this, EventArgs.Empty);
        _cmbThird.SelectedIndexChanged += (s, e) => ConfigChanged?.Invoke(this, EventArgs.Empty);

        this.Controls.Add(mainPanel);
        this.AutoSize = true;
    }

    public void SetHeaders(IEnumerable<string> headers, string noneOption)
    {
        var list = headers.ToList();
        
        _cmbFirst.DataSource = list.ToList();
        
        var secondList = new List<string> { noneOption };
        secondList.AddRange(list);
        _cmbSecond.DataSource = secondList;

        var thirdList = new List<string> { noneOption };
        thirdList.AddRange(list);
        _cmbThird.DataSource = thirdList;

        if (list.Count > 0)
        {
            _cmbFirst.SelectedIndex = 0;
            _cmbSecond.SelectedIndex = 0;
            _cmbThird.SelectedIndex = 0;
        }
    }

    public void SetPreview(string text)
    {
        _lblPreview.Text = text;
    }

    public void SetEnabled(bool enabled)
    {
        _txtPrefix.Enabled = enabled;
        _txtPostfix.Enabled = enabled;
        _txtSeparator.Enabled = enabled;
        _cmbFirst.Enabled = enabled;
        _cmbSecond.Enabled = enabled;
        _cmbThird.Enabled = enabled;
    }
}
