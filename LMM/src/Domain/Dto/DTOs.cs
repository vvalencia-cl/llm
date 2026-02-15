namespace LMM.Domain.Dto;

public sealed record ExcelHeaderResult(
    string WorksheetName,
    int HeaderRowNumber,
    IReadOnlyList<string> Headers,
    int LastDataRowNumber);

public sealed record MergeProgress(
    int Done,
    int Total,
    string? Message);