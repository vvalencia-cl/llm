using System.Diagnostics;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using LMM.UI;

namespace LMM.Application.Tests;

public class MainFormTests : IDisposable
{
    private readonly string _testFilesDir;
    private readonly string _excelPath;
    private readonly string _templatePath;
    private readonly string _outputDir;

    public MainFormTests()
    {
        // Try to find project root by looking for LMM.slnx
        string? root = AppDomain.CurrentDomain.BaseDirectory;
        while (root != null && !File.Exists(Path.Combine(root, "LMM.slnx")))
        {
            root = Path.GetDirectoryName(root);
        }
        root ??= Directory.GetCurrentDirectory();

        _testFilesDir = Path.Combine(root, "LMM.Application.Tests", "TestFiles");
        Directory.CreateDirectory(_testFilesDir);

        _excelPath = Path.Combine(_testFilesDir, "data.xlsx");
        _templatePath = Path.Combine(_testFilesDir, "template.docx");
        
        // Output PDFs go to a temp folder to avoid cluttering the repo and being versioned
        _outputDir = Path.Combine(Path.GetTempPath(), $"LMM_Test_Output_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_outputDir);

        // Cleanup old output dir in the repo if it exists from previous versions of the test
        var oldOutputDir = Path.Combine(_testFilesDir, "Output");
        if (Directory.Exists(oldOutputDir))
        {
            try { Directory.Delete(oldOutputDir, true); } catch { }
        }
    }

    public void Dispose()
    {
        // Don't delete _testFilesDir (contains data.xlsx and template.docx) 
        // so user can see/modify them as requested.

        // Clean up the generated PDFs in the temp output directory.
        if (Directory.Exists(_outputDir))
        {
            try
            {
                Directory.Delete(_outputDir, true);
            }
            catch { /* best effort */ }
        }
    }

    private void CreateExcelFile()
    {
        if (File.Exists(_excelPath)) return;

        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Data");
        
        // Headers
        worksheet.Cell(1, 1).Value = "name";
        worksheet.Cell(1, 2).Value = "email";
        worksheet.Cell(1, 3).Value = "message";
        worksheet.Cell(1, 4).Value = "supervisor name";
        worksheet.Cell(1, 5).Value = "RUT";

        // 10 Records
        for (int i = 1; i <= 10; i++)
        {
            int row = i + 1;
            worksheet.Cell(row, 1).Value = $"Employee {i}";
            worksheet.Cell(row, 2).Value = $"employee{i}@example.com";
            worksheet.Cell(row, 3).Value = $"Message for employee {i}";
            worksheet.Cell(row, 4).Value = $"Supervisor {i}";
            worksheet.Cell(row, 5).Value = $"12345678-{i}";
        }
        
        workbook.SaveAs(_excelPath);
    }

    private void CreateWordTemplate()
    {
        if (File.Exists(_templatePath)) return;

        using var doc = WordprocessingDocument.Create(_templatePath, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(
            new Body(
                new Paragraph(
                    new Run(new Text("Dear ")),
                    new SimpleField { Instruction = " MERGEFIELD name " }
                ),
                new Paragraph(
                    new Run(new Text("Your RUT is ")),
                    new SimpleField { Instruction = " MERGEFIELD RUT " }
                ),
                new Paragraph(
                    new Run(new Text("Your supervisor is ")),
                    new SimpleField { Instruction = " MERGEFIELD \"supervisor name\" " }
                ),
                new Paragraph(
                    new Run(new Text("Message: ")),
                    new SimpleField { Instruction = " MERGEFIELD message " }
                ),
                new Paragraph(
                    new Run(new Text("Email: ")),
                    new SimpleField { Instruction = " MERGEFIELD email " }
                )
            )
        );
        doc.Save();
    }

    [Fact]
    public async Task FullMailMergeFlow_WithRealFiles_WorksCorrectly()
    {
        // 1. Arrange
        CreateExcelFile();
        CreateWordTemplate();

        var view = new MockMainView();
        view.TemplatePath = _templatePath;
        view.ExcelPath = _excelPath;
        view.OutputDir = _outputDir;
        view.FilenameFirstField = "name";
        view.FilenameSeparator = "_";

        var presenter = new MainPresenter(view);

        // 2. Act
        
        // Load Excel
        var loadTask = view.WaitForStateChange(s => s.ExcelLoaded);
        view.TriggerLoadExcel();
        await loadTask;

        // Scan Template
        view.TriggerScanTemplate();
        // ScanTemplate is synchronous in Presenter (but might update state)
        Assert.True(view.LastState?.CanRun, "Should be able to run after scan");

        // Run Merge
        var runTask = view.WaitForStateChange(s => s.MergeFinished);
        view.TriggerRun();
        
        // Wait for completion (might take some time due to Word Interop)
        // Set a timeout of 60 seconds
        var completedTask = await Task.WhenAny(runTask, Task.Delay(TimeSpan.FromSeconds(60)));
        if (completedTask != runTask)
        {
            throw new TimeoutException("Mail merge process timed out.");
        }

        // 3. Assert
        var pdfFiles = Directory.GetFiles(_outputDir, "*.pdf");
        Assert.Equal(10, pdfFiles.Length);

        // Check some filenames
        Assert.Contains(pdfFiles, f => Path.GetFileName(f) == "Employee 1.pdf");
        Assert.Contains(pdfFiles, f => Path.GetFileName(f) == "Employee 10.pdf");

        // Verify that Word processes were cleaned up (best effort)
        // Actually, we can't easily check that without counting processes, 
        // but the fact it finished means WordPdfExporter.Dispose was likely called.
    }

    private class MockMainView : IMainView
    {
        public string WindowTitle { get; set; } = "";
        public string TemplatePath { get; set; } = "";
        public string ExcelPath { get; set; } = "";
        public string OutputDir { get; set; } = "";
        public string? SelectedWorksheet { get; set; }
        public int HeaderRow { get; set; } = 1;
        public string FilenamePrefix => "";
        public string FilenamePostfix => "";
        public string FilenameSeparator { get; set; } = "_";
        public string? FilenameFirstField { get; set; }
        public string? FilenameSecondField => "(Ninguno)";
        public string? FilenameThirdField => "(Ninguno)";
        public bool ClearOutputDir { get; set; }

        public AppState? LastState { get; private set; }
        private readonly List<Action<AppState>> _stateCallbacks = new();

        public void SetWorksheets(IEnumerable<string> worksheets) 
        { 
            SelectedWorksheet = worksheets.FirstOrDefault(); 
        }
        public void SetExcelHeaders(IEnumerable<string> headers) { }
        public void SetFilenamePreview(string preview) { }
        public void AppendLog(string message) { Debug.WriteLine(message); }
        public void UpdateProgress(int value, int total) { }
        public void SetBusy(bool busy) { }
        public void ApplyState(AppState state) 
        { 
            LastState = state;
            var callbacks = _stateCallbacks.ToList();
            foreach (var cb in callbacks) cb(state);
        }

        public void ShowError(string title, string message) { Debug.WriteLine($"ERROR: {title} - {message}"); }
        public void ShowInfo(string title, string message) { Debug.WriteLine($"INFO: {title} - {message}"); }
        public void ShowWarning(string title, string message) { Debug.WriteLine($"WARNING: {title} - {message}"); }
        public bool Confirm(string title, string message) => true;
        public string? PickFile(string title, string filter) => null;
        public string? PickFolder(string title) => null;

#pragma warning disable CS0067
        public event EventHandler? TemplatePathChanged;
        public event EventHandler? ExcelPathChanged;
        public event EventHandler? OutputDirChanged;
        public event EventHandler? ExcelOptionsChanged;
        public event EventHandler? FilenameConfigChanged;
        public event EventHandler? BrowseTemplateClicked;
        public event EventHandler? BrowseExcelClicked;
        public event EventHandler? BrowseOutputDirClicked;
        public event EventHandler? LoadExcelClicked;
        public event EventHandler? RefreshHeadersClicked;
        public event EventHandler? ScanTemplateClicked;
        public event EventHandler? RunClicked;
        public event EventHandler? CancelClicked;
        public event EventHandler? OpenOutputDirClicked;
        public event EventHandler? ClearOutputDirChanged;
#pragma warning restore CS0067

        public void TriggerLoadExcel() => LoadExcelClicked?.Invoke(this, EventArgs.Empty);
        public void TriggerScanTemplate() => ScanTemplateClicked?.Invoke(this, EventArgs.Empty);
        public void TriggerRun() => RunClicked?.Invoke(this, EventArgs.Empty);

        public Task WaitForStateChange(Predicate<AppState> condition)
        {
            var tcs = new TaskCompletionSource();
            Action<AppState> handler = null!;
            handler = (state) =>
            {
                if (condition(state))
                {
                    lock (_stateCallbacks) { _stateCallbacks.Remove(handler); }
                    tcs.TrySetResult();
                }
            };
            lock (_stateCallbacks) { _stateCallbacks.Add(handler); }
            
            // Check if already matches
            if (LastState != null && condition(LastState))
            {
                lock (_stateCallbacks) { _stateCallbacks.Remove(handler); }
                tcs.TrySetResult();
            }

            return tcs.Task;
        }
    }
}
