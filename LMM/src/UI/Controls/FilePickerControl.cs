using System.ComponentModel;

namespace LMM.UI.Controls;

public partial class FilePickerControl : UserControl
{
    private Label _label = null!;
    private TextBox _textBox = null!;
    private Button _button = null!;

    public event EventHandler? FileSelected;

    [Category("Appearance")]
    public string LabelText
    {
        get => _label.Text;
        set => _label.Text = value;
    }

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
        _label = new Label { Dock = DockStyle.Left, AutoSize = true, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(0, 0, 10, 0) };
        _textBox = new TextBox { Dock = DockStyle.Fill };
        _button = new Button { Dock = DockStyle.Right, Text = "Buscar...", AutoSize = true };

        _button.Click += (s, e) => OnBrowse();
        _textBox.TextChanged += (s, e) => FileSelected?.Invoke(this, EventArgs.Empty);

        this.Controls.Add(_textBox);
        this.Controls.Add(_label);
        this.Controls.Add(_button);
        this.Height = _textBox.Height;
    }

    protected virtual void OnBrowse()
    {
        // Esto se sobrecargará o se usará un delegado si queremos hacerlo genérico, 
        // pero para simplificar, el Presenter dirá qué hacer.
        BrowseClicked?.Invoke(this, EventArgs.Empty);
    }

    public event EventHandler? BrowseClicked;
}
