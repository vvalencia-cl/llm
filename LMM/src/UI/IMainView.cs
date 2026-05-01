using LMM.Domain.Dto;

namespace LMM.UI;

public interface IMainView
{
    // Propiedades de datos
    string TemplatePath { get; set; }
    string ExcelPath { get; set; }
    string OutputDir { get; set; }
    
    string? SelectedWorksheet { get; set; }
    int HeaderRow { get; set; }
    
    string FilenamePrefix { get; }
    string FilenamePostfix { get; }
    string FilenameSeparator { get; }
    string? FilenameFirstField { get; }
    string? FilenameSecondField { get; }
    string? FilenameThirdField { get; }

    // Métodos de actualización de UI
    void SetWorksheets(IEnumerable<string> worksheets);
    void SetExcelHeaders(IEnumerable<string> headers);
    void SetFilenamePreview(string preview);
    void AppendLog(string message);
    void UpdateProgress(int value, int total);
    void SetBusy(bool busy);
    void ApplyState(AppState state);
    
    // Diálogos y alertas
    void ShowError(string title, string message);
    void ShowInfo(string title, string message);
    void ShowWarning(string title, string message);
    bool Confirm(string title, string message);
    string? PickFile(string title, string filter);
    string? PickFolder(string title);

    // Eventos
    event EventHandler TemplatePathChanged;
    event EventHandler ExcelPathChanged;
    event EventHandler OutputDirChanged;
    event EventHandler ExcelOptionsChanged;
    event EventHandler FilenameConfigChanged;
    
    event EventHandler BrowseTemplateClicked;
    event EventHandler BrowseExcelClicked;
    event EventHandler BrowseOutputDirClicked;
    event EventHandler LoadExcelClicked;
    event EventHandler RefreshHeadersClicked;
    event EventHandler ScanTemplateClicked;
    event EventHandler RunClicked;
    event EventHandler CancelClicked;
}
