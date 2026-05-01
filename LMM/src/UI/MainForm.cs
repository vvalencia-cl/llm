using LMM.UI.Controls;

namespace LMM.UI;

public sealed partial class MainForm : Form, IMainView
{
    private readonly MainPresenter _presenter;

    // Controls
    private FilePickerControl _fpTemplate = null!;
    private FilePickerControl _fpExcel = null!;
    private FilePickerControl _fpOutputDir = null!;
    private ExcelOptionsControl _excelOptions = null!;
    private FilenameBuilderControl _filenameBuilder = null!;
    private LogPanelControl _logPanel = null!;
    
    private Button _btnScan = null!;
    private Button _btnRun = null!;
    private Button _btnCancel = null!;
    private Button _btnLoadExcel = null!;

    private const string OptionalFieldNoneOption = "(Ninguno)";

    public MainForm()
    {
        InitializeComponent();
        _presenter = new MainPresenter(this);
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        _presenter.DoDispose();
        base.OnFormClosed(e);
    }

    #region IMainView Implementation

    public string TemplatePath { get => _fpTemplate.SelectedPath; set => _fpTemplate.SelectedPath = value; }
    public string ExcelPath { get => _fpExcel.SelectedPath; set => _fpExcel.SelectedPath = value; }
    public string OutputDir { get => _fpOutputDir.SelectedPath; set => _fpOutputDir.SelectedPath = value; }
    public string? SelectedWorksheet { get => _excelOptions.SelectedWorksheet; set => _excelOptions.SelectedWorksheet = value; }
    public int HeaderRow { get => _excelOptions.HeaderRow; set => _excelOptions.HeaderRow = value; }
    public string FilenamePrefix => _filenameBuilder.Prefix;
    public string FilenamePostfix => _filenameBuilder.Postfix;
    public string FilenameSeparator => _filenameBuilder.Separator;
    public string? FilenameFirstField => _filenameBuilder.FirstField;
    public string? FilenameSecondField => _filenameBuilder.SecondField;
    public string? FilenameThirdField => _filenameBuilder.ThirdField;

    public void SetWorksheets(IEnumerable<string> worksheets) => _excelOptions.SetWorksheets(worksheets);
    public void SetExcelHeaders(IEnumerable<string> headers) => _filenameBuilder.SetHeaders(headers, OptionalFieldNoneOption);
    public void SetFilenamePreview(string preview) => _filenameBuilder.SetPreview(preview);
    public void AppendLog(string message) => _logPanel.AppendLog(message);
    public void UpdateProgress(int value, int total) => _logPanel.SetProgress(value, total);
    public void SetBusy(bool busy)
    {
        _logPanel.SetMarquee(busy);
        this.Cursor = busy ? Cursors.WaitCursor : Cursors.Default;
    }

    public void ApplyState(AppState state)
    {
        if (InvokeRequired) { Invoke(() => ApplyState(state)); return; }

        _excelOptions.SetEnabled(state.CanRefreshHeaders);
        _filenameBuilder.SetEnabled(state.HeadersReady && !state.IsProcessing);
        
        _btnScan.Enabled = state.CanScanTemplate;
        _btnRun.Enabled = state.CanRun;
        _btnCancel.Enabled = state.CanCancel;
        
        _btnLoadExcel.Enabled = !state.IsProcessing;
        _fpExcel.Enabled = !state.IsProcessing;
        _fpTemplate.Enabled = !state.IsProcessing;
        _fpOutputDir.Enabled = !state.IsProcessing;
    }

    public void ShowError(string title, string message) => MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Error);
    public void ShowInfo(string title, string message) => MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
    public void ShowWarning(string title, string message) => MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
    public bool Confirm(string title, string message) => MessageBox.Show(message, title, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;

    public string? PickFile(string title, string filter)
    {
        using var dlg = new OpenFileDialog();
        dlg.Title = title;
        dlg.Filter = filter;
        return dlg.ShowDialog(this) == DialogResult.OK ? dlg.FileName : null;
    }

    public string? PickFolder(string title)
    {
        using var dlg = new FolderBrowserDialog();
        dlg.Description = title;
        dlg.UseDescriptionForTitle = true;
        return dlg.ShowDialog(this) == DialogResult.OK ? dlg.SelectedPath : null;
    }

    public event EventHandler TemplatePathChanged { add => _fpTemplate.FileSelected += value; remove => _fpTemplate.FileSelected -= value; }
    public event EventHandler ExcelPathChanged { add => _fpExcel.FileSelected += value; remove => _fpExcel.FileSelected -= value; }
    public event EventHandler OutputDirChanged { add => _fpOutputDir.FileSelected += value; remove => _fpOutputDir.FileSelected -= value; }
    public event EventHandler ExcelOptionsChanged { add => _excelOptions.OptionsChanged += value; remove => _excelOptions.OptionsChanged -= value; }
    public event EventHandler FilenameConfigChanged { add => _filenameBuilder.ConfigChanged += value; remove => _filenameBuilder.ConfigChanged -= value; }
    public event EventHandler BrowseTemplateClicked { add => _fpTemplate.BrowseClicked += value; remove => _fpTemplate.BrowseClicked -= value; }
    public event EventHandler BrowseExcelClicked { add => _fpExcel.BrowseClicked += value; remove => _fpExcel.BrowseClicked -= value; }
    public event EventHandler BrowseOutputDirClicked { add => _fpOutputDir.BrowseClicked += value; remove => _fpOutputDir.BrowseClicked -= value; }
    public event EventHandler LoadExcelClicked { add => _btnLoadExcel.Click += value; remove => _btnLoadExcel.Click -= value; }
    public event EventHandler RefreshHeadersClicked { add => _excelOptions.RefreshClicked += value; remove => _excelOptions.RefreshClicked -= value; }
    public event EventHandler ScanTemplateClicked { add => _btnScan.Click += value; remove => _btnScan.Click -= value; }
    public event EventHandler RunClicked { add => _btnRun.Click += value; remove => _btnRun.Click -= value; }
    public event EventHandler CancelClicked { add => _btnCancel.Click += value; remove => _btnCancel.Click -= value; }

    #endregion
}