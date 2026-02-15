using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace LMM.Application;

public sealed class WordPdfExporter : IDisposable
{
    private Word.Application? _wordApp;
    private bool _disposed;

    public WordPdfExporter()
    {
        _wordApp = new Word.Application
        {
            Visible = false,
            DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
        };
    }

    public void ExportDocxToPdf(string docxPath, string pdfPath)
    {
        if (_disposed) throw new ObjectDisposedException(nameof(WordPdfExporter));
        if (_wordApp == null) throw new InvalidOperationException("Word application is not initialized.");

        if (!File.Exists(docxPath))
            throw new FileNotFoundException("DOCX not found.", docxPath);

        Directory.CreateDirectory(Path.GetDirectoryName(pdfPath)!);

        Word.Document? doc = null;
        try
        {
            object missing = Type.Missing;
            object readOnly = true;
            object isVisible = false;

            doc = _wordApp.Documents.Open(
                FileName: docxPath,
                ReadOnly: readOnly,
                Visible: isVisible,
                ConfirmConversions: false,
                AddToRecentFiles: false,
                NoEncodingDialog: true);

            // Overwrite: Interop will overwrite if the path exists in most cases,
            // but it's safer to delete first at the caller level (as you planned).
            doc.ExportAsFixedFormat(
                OutputFileName: pdfPath,
                ExportFormat: Word.WdExportFormat.wdExportFormatPDF,
                OpenAfterExport: false,
                OptimizeFor: Word.WdExportOptimizeFor.wdExportOptimizeForPrint,
                Range: Word.WdExportRange.wdExportAllDocument,
                Item: Word.WdExportItem.wdExportDocumentContent,
                IncludeDocProps: true,
                KeepIRM: true,
                CreateBookmarks: Word.WdExportCreateBookmarks.wdExportCreateNoBookmarks,
                DocStructureTags: true,
                BitmapMissingFonts: true,
                UseISO19005_1: false);
        }
        finally
        {
            if (doc != null)
            {
                try { doc.Close(SaveChanges: false); } catch { /* best-effort */ }
                ReleaseCom(doc);
            }
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        if (_wordApp != null)
        {
            try { _wordApp.Quit(SaveChanges: false); } catch { /* best-effort */ }
            ReleaseCom(_wordApp);
            _wordApp = null;
        }

        // Helps COM finalize (not strictly required, but reduces zombie WINWORD risk)
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    private static void ReleaseCom(object comObj)
    {
        try
        {
            if (Marshal.IsComObject(comObj))
                Marshal.FinalReleaseComObject(comObj);
        }
        catch
        {
            // best-effort; avoid throwing during cleanup
        }
    }
}