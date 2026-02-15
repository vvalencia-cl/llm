using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace LMM.Application;

public sealed class WordPdfExporter : IDisposable
{
    private Word.Application? wordApp;
    private bool disposed;

    public WordPdfExporter()
    {
        wordApp = new Word.Application
        {
            Visible = false,
            DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
        };
    }

    public void ExportDocxToPdf(string docxPath, string pdfPath)
    {
        if (disposed) throw new ObjectDisposedException(nameof(WordPdfExporter));
        if (wordApp == null) throw new InvalidOperationException("La aplicación Word no está inicializada.");

        if (!File.Exists(docxPath))
            throw new FileNotFoundException("Archivo DOCX no encontrado.", docxPath);

        Directory.CreateDirectory(Path.GetDirectoryName(pdfPath)!);

        Word.Document? doc = null;
        try
        {
            object readOnly = true;
            object isVisible = false;

            doc = wordApp.Documents.Open(
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
        if (disposed) return;
        disposed = true;

        if (wordApp != null)
        {
            try { wordApp.Quit(SaveChanges: false); } catch { /* best-effort */ }
            ReleaseCom(wordApp);
            wordApp = null;
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