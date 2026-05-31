namespace LMM.UI.Controls;

public partial class LogPanelControl : UserControl
{
    private ListBox _lstLog = null!;

    public LogPanelControl()
    {
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        _lstLog = new ListBox 
        { 
            Dock = DockStyle.Fill, 
            SelectionMode = SelectionMode.MultiExtended,
            BorderStyle = BorderStyle.FixedSingle,
            Font = new Font("Consolas", 9F) // Monospaced font for logs
        };

        var miCopy = new ToolStripMenuItem("Copiar seleccionados");
        miCopy.Click += (s, e) => CopySelected();
        
        var logMenu = new ContextMenuStrip();
        logMenu.Items.Add(miCopy);
        _lstLog.ContextMenuStrip = logMenu;

        _lstLog.KeyDown += (s, e) =>
        {
            if (e.Control && e.KeyCode == Keys.C)
            {
                CopySelected();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        };

        this.Controls.Add(_lstLog);
    }

    public void AppendLog(string message)
    {
        if (InvokeRequired)
        {
            Invoke(() => AppendLog(message));
            return;
        }
        _lstLog.Items.Add(message);
        _lstLog.TopIndex = Math.Max(0, _lstLog.Items.Count - 1);
    }

    private void CopySelected()
    {
        if (_lstLog.SelectedItems.Count == 0) return;

        var lines = _lstLog.SelectedItems.Cast<object>().Select(x => x?.ToString() ?? string.Empty);
        var text = string.Join(Environment.NewLine, lines);

        try
        {
            Clipboard.SetText(text);
        }
        catch (Exception ex)
        {
            MessageBox.Show("Error al copiar al portapapeles: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
