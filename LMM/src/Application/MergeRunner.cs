using LMM.Domain.Dto;

namespace LMM.Application;

public static class MergeRunner
{
    public static MergeRunSummary RunMergeContinueOnError(
        ClosedXmlDataSourceService excel,
        string worksheetName,
        int headerRowNumber,
        IReadOnlyList<string> headers,
        Func<Dictionary<string, string>, string> buildPdfPath, // uses FieldX/FieldY + output dir
        Action<int, int> reportProgress, // (done, totalEstimate)
        Action<string> logLine,
        Action<Dictionary<string, string>, string> mergeAndExportPdfForRecord // (record, pdfPath)
    )
    {
        var errors = new List<string>();
        int totalSeen = 0, ok = 0, fail = 0;

        // For progress, you can estimate total rows from headers result:
        // lastDataRow - headerRowNumber, but we may skip empty rows => it's an estimate.
        int totalEstimate = 0;
        try
        {
            // If you kept LastDataRowNumber from headerInfo, use that here.
            // Otherwise just keep 0 and show indeterminate progress in UI.
        }
        catch
        {
            totalEstimate = 0;
        }

        foreach (var (rowNumber, record) in excel.EnumerateRecordsFormattedWithRowNumber(
                     worksheetName,
                     headerRowNumber,
                     headers,
                     skipFullyEmptyRows: true))
        {
            totalSeen++;

            string pdfPath = "";
            try
            {
                pdfPath = buildPdfPath(record);

                // Overwrite behavior (your requirement)
                if (File.Exists(pdfPath))
                    File.Delete(pdfPath);

                mergeAndExportPdfForRecord(record, pdfPath);

                ok++;
                logLine($"Row {rowNumber}: OK -> {Path.GetFileName(pdfPath)}");
            }
            catch (Exception ex)
            {
                fail++;
                var msg = $"Row {rowNumber}: ERROR -> {ex.Message}";
                errors.Add(msg);
                logLine(msg);

                // continue
            }

            reportProgress(ok + fail, totalEstimate);
            // Optional: small pause to keep UI responsive if you're logging heavily
            // Thread.Sleep(1);
        }

        logLine($"Done. Total processed: {totalSeen}. OK: {ok}. Failed: {fail}.");

        return new MergeRunSummary(
            TotalRowsSeen: totalSeen,
            SuccessCount: ok,
            FailureCount: fail,
            ErrorLines: errors);
    }
}