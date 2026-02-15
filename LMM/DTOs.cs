using System;
using System.Collections.Generic;

public sealed record ExcelLoadResult(
    IReadOnlyList<string> WorksheetNames,
    string DefaultWorksheetName);

public sealed record ExcelHeaderResult(
    string WorksheetName,
    int HeaderRowNumber,
    IReadOnlyList<string> Headers,
    int LastDataRowNumber);

public sealed record MergeRunSummary(
    int TotalRowsSeen,
    int SuccessCount,
    int FailureCount,
    IReadOnlyList<string> ErrorLines);

public sealed record MergeProgress(
    int Done,
    int Total,
    string? Message);