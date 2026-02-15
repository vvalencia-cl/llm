namespace LMM.Application;

public static class PdfFilenameBuilder
{
    public static string BuildPdfPath(
        string outputDirectory,
        Dictionary<string, string> record,
        string firstFieldHeader,
        string secondFieldHeader,
        string? thirdFieldHeader = null,
        string separator = "_",
        string emptyFallback = "row") // used only if all parts become empty after sanitization
    {
        if (string.IsNullOrWhiteSpace(outputDirectory))
            throw new ArgumentException("El directorio de salida es obligatorio.", nameof(outputDirectory));

        if (!Directory.Exists(outputDirectory))
            throw new DirectoryNotFoundException($"Directorio de salida no encontrado: {outputDirectory}");

        if (record == null)
            throw new ArgumentNullException(nameof(record));

        if (string.IsNullOrWhiteSpace(firstFieldHeader))
            throw new ArgumentException("El encabezado del FirstField es obligatorio.", nameof(firstFieldHeader));

        if (string.IsNullOrWhiteSpace(secondFieldHeader))
            throw new ArgumentException("El encabezado del SecondField es obligatorio.", nameof(secondFieldHeader));

        var firstRaw = record.TryGetValue(firstFieldHeader, out var v1) ? (v1 ?? "") : "";
        var secondRaw = record.TryGetValue(secondFieldHeader, out var v2) ? (v2 ?? "") : "";

        var thirdRaw = "";
        if (!string.IsNullOrWhiteSpace(thirdFieldHeader))
            thirdRaw = record.TryGetValue(thirdFieldHeader, out var v3) ? (v3 ?? "") : "";

        var first = SanitizeFilenamePart(firstRaw);
        var second = SanitizeFilenamePart(secondRaw);
        var third = SanitizeFilenamePart(thirdRaw);

        var parts = new List<string>(capacity: 3);
        if (!string.IsNullOrEmpty(first)) parts.Add(first);
        if (!string.IsNullOrEmpty(second)) parts.Add(second);
        if (!string.IsNullOrEmpty(third)) parts.Add(third);

        var baseName = parts.Count > 0 ? string.Join(separator, parts) : emptyFallback;

        baseName = TrimToMaxBaseNameLength(baseName, maxLength: 180); // keep room for ".pdf" + path

        var fileName = baseName + ".pdf";
        return Path.Combine(outputDirectory, fileName);
    }

    /// <summary>
    /// Makes a string safe for Windows filenames (not paths).
    /// - Removes invalid characters
    /// - Collapses whitespace
    /// - Trims trailing dots/spaces (Windows restriction)
    /// - Avoids reserved device names
    /// </summary>
    public static string SanitizeFilenamePart(string input)
    {
        if (string.IsNullOrWhiteSpace(input))
            return "";

        // Normalize whitespace
        var s = input.Trim();

        // Replace invalid filename chars with space
        var invalid = Path.GetInvalidFileNameChars();
        s = new string(s.Select(ch => invalid.Contains(ch) ? ' ' : ch).ToArray());

        // Collapse consecutive whitespace
        s = CollapseWhitespace(s);

        // Windows: filenames cannot end with dot or space
        s = s.TrimEnd('.', ' ');

        // Avoid reserved device names (CON, PRN, AUX, NUL, COM1.., LPT1..)
        if (IsReservedDeviceName(s))
            s = "_" + s;

        return s;
    }

    private static string CollapseWhitespace(string s)
    {
        var result = new System.Text.StringBuilder(s.Length);
        bool prevWs = false;

        foreach (var ch in s)
        {
            var isWs = char.IsWhiteSpace(ch);
            if (isWs)
            {
                if (!prevWs) result.Append(' ');
                prevWs = true;
            }
            else
            {
                result.Append(ch);
                prevWs = false;
            }
        }

        return result.ToString().Trim();
    }

    private static bool IsReservedDeviceName(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return false;

        var name = s.Trim().TrimEnd('.'); // Windows treats trailing dots specially
        var upper = name.ToUpperInvariant();

        // Exact reserved names
        if (upper is "CON" or "PRN" or "AUX" or "NUL")
            return true;

        // COM1..COM9, LPT1..LPT9
        if (upper.StartsWith("COM", StringComparison.Ordinal) && upper.Length == 4 && char.IsDigit(upper[3]))
            return true;

        if (upper.StartsWith("LPT", StringComparison.Ordinal) && upper.Length == 4 && char.IsDigit(upper[3]))
            return true;

        return false;
    }

    private static string TrimToMaxBaseNameLength(string baseName, int maxLength)
    {
        if (string.IsNullOrEmpty(baseName)) return baseName;
        if (baseName.Length <= maxLength) return baseName;

        return baseName.Substring(0, maxLength).TrimEnd('.', ' ');
    }
}