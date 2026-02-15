using System.Text.RegularExpressions;

namespace LMM.Application;

public static class HeaderFieldMapper
{

    // Case-sensitive normalization:
    // - Trim
    // - Remove % and :
    // - Treat whitespace, underscore and hyphen as the same separator
    // - Collapse separators to single underscore
    // - Trim leading/trailing underscores (handles Word fields like CARGO_)
    private static string Normalize(string s)
    {
        if (s == null) return "";
        s = s.Trim();

        s = s.Replace("%", "");
        s = s.Replace(":", "");

        // Treat spaces/underscores/hyphens as equivalent separators
        s = Regex.Replace(s, @"[\s_\-]+", "_");

        // Remove accidental leading/trailing underscores (handles Word fields like CARGO_)
        s = s.Trim('_');

        return s;
    }

    // NEW: Join patterns like _01_2026 -> _012026 (Word often does this for month-year)
    private static string JoinMonthYearTokens(string normalized)
    {
        if (string.IsNullOrWhiteSpace(normalized)) return normalized;

        // Join only 2-digit month + 4-digit year tokens
        // Example: SUELDO_01_2026_AJUSTADO => SUELDO_012026_AJUSTADO
        return Regex.Replace(normalized, @"_(\d{2})_(\d{4})(?=_|$)", "_$1$2");
    }

    private static bool ExcelHeaderLikelyGetsMPrefix(string excelHeader)
    {
        if (string.IsNullOrWhiteSpace(excelHeader)) return false;
        var t = excelHeader.TrimStart();

        // If it starts with % or a digit, Word commonly prefixes field names (e.g., M_10_PONDERADO).
        return t.StartsWith("%", StringComparison.Ordinal) ||
               char.IsDigit(t[0]);
    }

    /// <summary>
    /// Builds mapping: templateFieldName -> excelHeaderName.
    /// Prefers exact match; falls back to normalized match.
    /// Also supports Word "M_" prefix for headers that begin with '%' or a digit.
    /// Case-sensitive.
    /// Throws if ambiguous.
    /// </summary>
    public static Dictionary<string, string> BuildTemplateToExcelHeaderMap(
        IReadOnlyList<string> templateFields,
        IReadOnlyList<string> excelHeaders)
    {
        var excelExact = new HashSet<string>(
            excelHeaders.Where(h => !string.IsNullOrWhiteSpace(h)),
            StringComparer.Ordinal);

        var normalizedLookup = new Dictionary<string, List<string>>(StringComparer.Ordinal);

        foreach (var header in excelHeaders.Where(h => !string.IsNullOrWhiteSpace(h)))
        {
            var n = Normalize(header);
            AddNormalized(normalizedLookup, n, header);

            // NEW: add month-year-joined variant (e.g., 01_2026 -> 012026)
            AddNormalized(normalizedLookup, JoinMonthYearTokens(n), header);

            // Existing: Word "M_" prefix variant for headers starting with % or a digit
            if (ExcelHeaderLikelyGetsMPrefix(header))
            {
                AddNormalized(normalizedLookup, "M_" + n, header);
                AddNormalized(normalizedLookup, "M_" + JoinMonthYearTokens(n), header); // NEW
            }
        }

        var map = new Dictionary<string, string>(StringComparer.Ordinal);

        foreach (var field in templateFields)
        {
            if (excelExact.Contains(field))
            {
                map[field] = field;
                continue;
            }

            var nf = Normalize(field);

            // Try normalized key
            if (TryResolve(normalizedLookup, nf, field, out var excelHeader))
            {
                map[field] = excelHeader;
                continue;
            }

            // NEW: try month-year-joined normalized key for template field too
            var nfJoined = JoinMonthYearTokens(nf);
            if (!string.Equals(nfJoined, nf, StringComparison.Ordinal) &&
                TryResolve(normalizedLookup, nfJoined, field, out excelHeader))
            {
                map[field] = excelHeader;
                continue;
            }

            // Existing: if template begins with M_, try without it
            if (nf.StartsWith("M_", StringComparison.Ordinal))
            {
                var withoutM = nf.Substring(2);

                if (TryResolve(normalizedLookup, withoutM, field, out excelHeader))
                {
                    map[field] = excelHeader;
                    continue;
                }

                // NEW: also try joined variant without M_
                var withoutMJoined = JoinMonthYearTokens(withoutM);
                if (!string.Equals(withoutMJoined, withoutM, StringComparison.Ordinal) &&
                    TryResolve(normalizedLookup, withoutMJoined, field, out excelHeader))
                {
                    map[field] = excelHeader;
                    continue;
                }
            }

            map[field] = "";
        }

        return map;
    }

    private static void AddNormalized(Dictionary<string, List<string>> lookup, string key, string originalHeader)
    {
        if (string.IsNullOrWhiteSpace(key)) return;

        if (!lookup.TryGetValue(key, out var list))
        {
            list = new List<string>();
            lookup[key] = list;
        }

        if (!list.Contains(originalHeader, StringComparer.Ordinal))
            list.Add(originalHeader);
    }

    private static bool TryResolve(
        Dictionary<string, List<string>> lookup,
        string normalizedField,
        string originalTemplateField,
        out string excelHeader)
    {
        excelHeader = "";

        if (!lookup.TryGetValue(normalizedField, out var candidates))
            return false;

        if (candidates.Count == 1)
        {
            excelHeader = candidates[0];
            return true;
        }

        // Ambiguous mapping: more than one Excel header normalizes to the same key
        throw new InvalidOperationException(
            $"Ambiguous Excel headers for template field '{originalTemplateField}'. " +
            $"Multiple Excel headers match after normalization ('{normalizedField}'): {string.Join(", ", candidates)}");
    }

    /// <summary>
    /// Builds values keyed by TEMPLATE field names, using the mapping to pull from the Excel record.
    /// Missing values become "".
    /// </summary>
    public static Dictionary<string, string> BuildTemplateValuesForRecord(
        IReadOnlyList<string> templateFields,
        Dictionary<string, string> excelRecord,
        Dictionary<string, string> templateToExcelHeaderMap)
    {
        var values = new Dictionary<string, string>(StringComparer.Ordinal);

        foreach (var field in templateFields)
        {
            var excelHeader = templateToExcelHeaderMap[field];
            if (!string.IsNullOrEmpty(excelHeader) && excelRecord.TryGetValue(excelHeader, out var v) && v != null)
                values[field] = v;
            else
                values[field] = "";
        }

        return values;
    }

}