namespace LMM.UI.Controls;

public partial class LogPanelControl : UserControl
{
    private ProgressBar _progressBar = null!;
    private ListBox _lstLog = null!;

    public LogPanelControl()
    {
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        var mainPanel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2 };
        mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        _progressBar = new ProgressBar { Dock = DockStyle.Top, Height = 18 };
        mainPanel.Controls.Add(_progressBar, 0, 0);

        _lstLog = new ListBox { Dock = DockStyle.Fill, SelectionMode = SelectionMode.MultiExtended };
        mainPanel.Controls.Add(_lstLog, 0, 1);

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

        this.Controls.Add(mainPanel);
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

    public void SetProgress(int value, int maximum)
    {
        if (InvokeRequired)
        {
            Invoke(() => SetProgress(value, maximum));
            return;
        }
        _progressBar.Maximum = Math.Max(1, maximum);
        _progressBar.Value = Math.Min(_progressBar.Maximum, Math.Max(0, value));
    }

    public void SetMarquee(bool active)
    {
        if (InvokeRequired)
        {
            Invoke(() => SetMarquee(active));
            return;
        }
        _progressBar.Style = active ? ProgressBarStyle.Marquee : ProgressBarStyle.Blocks;
        _progressBar.MarqueeAnimationSpeed = active ? 30 : 0;
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
