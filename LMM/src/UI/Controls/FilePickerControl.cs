using System.ComponentModel;

namespace LMM.UI.Controls;

public partial class FilePickerControl : UserControl
{
    private TextBox _textBox = null!;
    private Button _button = null!;

    public event EventHandler? FileSelected;

    [Category("Appearance")]
    public string PlaceholderText
    {
        get => _textBox.PlaceholderText;
        set => _textBox.PlaceholderText = value;
    }

    [Browsable(false)]
    public string SelectedPath
    {
        get => _textBox.Text;
        set => _textBox.Text = value;
    }

    public FilePickerControl()
    {
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        _textBox = new TextBox { Dock = DockStyle.Fill };
        _button = new Button { Dock = DockStyle.Right, Text = "Buscar...", AutoSize = true, Margin = new Padding(5, 0, 0, 0) };

        _button.Click += (s, e) => OnBrowse();
        _textBox.TextChanged += (s, e) => FileSelected?.Invoke(this, EventArgs.Empty);

        this.Controls.Add(_textBox);
        this.Controls.Add(_button);
        this.Height = _textBox.Height;
        this.Padding = new Padding(0);
    }

    protected virtual void OnBrowse()
    {
        BrowseClicked?.Invoke(this, EventArgs.Empty);
    }

    public event EventHandler? BrowseClicked;
}
