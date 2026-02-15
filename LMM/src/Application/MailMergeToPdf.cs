using DocumentFormat.OpenXml.Packaging;

namespace LMM.Application;

public static class MailMergeToPdf
{
    public static void MergeAndExportPdfForRecord(
        Dictionary<string, string> record,
        string pdfPath,
        string templateDocxPath,
        WordPdfExporter wordExporter)
    {
        ArgumentNullException.ThrowIfNull(record);
        if (string.IsNullOrWhiteSpace(pdfPath)) throw new ArgumentException("La ruta del PDF es obligatoria.", nameof(pdfPath));
        if (string.IsNullOrWhiteSpace(templateDocxPath)) throw new ArgumentException("La ruta de la plantilla es obligatoria.", nameof(templateDocxPath));
        ArgumentNullException.ThrowIfNull(wordExporter);

        if (!File.Exists(templateDocxPath))
            throw new FileNotFoundException("Plantilla DOCX no encontrada.", templateDocxPath);

        var tempDocxPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.docx");

        try
        {
            File.Copy(templateDocxPath, tempDocxPath, overwrite: true);

            // Merge using OpenXML
            using (var doc = WordprocessingDocument.Open(tempDocxPath, true))
            {
                // Treat missing keys as "" is already enforced by your merge function,
                // but your record dictionary should already contain all headers anyway.
                OpenXmlMailMerge.ReplaceMergeFieldsInMainBody(doc, record);
            }

            // Export to PDF via Word
            wordExporter.ExportDocxToPdf(tempDocxPath, pdfPath);
        }
        finally
        {
            try
            {
                if (File.Exists(tempDocxPath))
                    File.Delete(tempDocxPath);
            }
            catch
            {
                // best-effort; don't mask the real error
            }
        }
    }
}