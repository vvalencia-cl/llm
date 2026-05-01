namespace LMM.UI;

/// <summary>
/// Representa el estado de habilitación de los controles en la vista principal.
/// </summary>
public sealed record AppState(
    bool ExcelLoaded,
    bool HeadersReady,
    bool TemplateExists,
    bool OutputDirExists,
    bool IsProcessing
)
{
    public bool CanRefreshHeaders => ExcelLoaded && !IsProcessing;
    public bool CanScanTemplate => TemplateExists && HeadersReady && !IsProcessing;
    public bool CanRun => TemplateExists && ExcelLoaded && HeadersReady && OutputDirExists && !IsProcessing;
    public bool CanCancel => IsProcessing;
}
