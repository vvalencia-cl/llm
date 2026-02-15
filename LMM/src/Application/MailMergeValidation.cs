namespace LMM.Application;

public static class MailMergeValidation
{
    /// <summary>
    /// Returns a list of template fields that do not have an exact matching Excel header.
    /// Exact match = StringComparer.Ordinal.
    /// </summary>
    public static IReadOnlyList<string> GetMissingHeaders(
        IEnumerable<string> templateFields,
        IEnumerable<string> excelHeaders)
    {
        var headerSet = new HashSet<string>(
            excelHeaders.Where(h => !string.IsNullOrWhiteSpace(h)).Select(h => h.Trim()),
            StringComparer.Ordinal);

        var missing = templateFields
            .Where(f => !headerSet.Contains(f))
            .Distinct(StringComparer.Ordinal)
            .OrderBy(f => f, StringComparer.Ordinal)
            .ToList();

        return missing;
    }
}