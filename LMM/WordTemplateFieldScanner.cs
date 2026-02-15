using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

public static class WordTemplateFieldScanner
{
    // Matches: MERGEFIELD FieldName   OR   MERGEFIELD "Field Name"
    private static readonly Regex MergeFieldRegex = new(
        @"\bMERGEFIELD\b\s+(?:(?:""(?<name>[^""]+)"")|(?<name>[^\s\\]+))",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>
    /// Returns distinct MERGEFIELD names found in the main document body.
    /// </summary>
    public static IReadOnlyList<string> GetMergeFieldNamesFromMainBody(string templateDocxPath)
    {
        using var doc = WordprocessingDocument.Open(templateDocxPath, false);
        var body = doc.MainDocumentPart?.Document?.Body
                   ?? throw new InvalidOperationException("Template DOCX has no main document body.");

        var result = new HashSet<string>(StringComparer.Ordinal);

        // 1) fldSimple
        foreach (var sf in body.Descendants<SimpleField>())
        {
            var instr = sf.Instruction?.Value ?? string.Empty;
            var name = ExtractMergeFieldName(instr);
            if (name != null) result.Add(name);
        }

        // 2) Complex fields: scan paragraphs for begin..end ranges and parse instruction text
        foreach (var p in body.Descendants<Paragraph>())
        {
            ScanParagraphForComplexMergeFields(p, result);
        }

        return result.OrderBy(x => x, StringComparer.Ordinal).ToList();
    }

    private static void ScanParagraphForComplexMergeFields(Paragraph paragraph, HashSet<string> output)
    {
        var runs = paragraph.Elements<Run>().ToList();
        if (runs.Count == 0) return;

        bool inField = false;
        bool inInstr = false;
        var instrBuilder = new StringBuilder();

        for (int i = 0; i < runs.Count; i++)
        {
            var run = runs[i];

            var fieldChar = run.GetFirstChild<FieldChar>();
            if (fieldChar != null)
            {
                var t = fieldChar.FieldCharType?.Value; // FieldCharValues?
                if (t == FieldCharValues.Begin)
                {
                    inField = true;
                    inInstr = true;
                    instrBuilder.Clear();
                }
                else if (t == FieldCharValues.Separate)
                {
                    if (inField) inInstr = false;
                }
                else if (t == FieldCharValues.End)
                {
                    if (inField)
                    {
                        var instr = instrBuilder.ToString();
                        var name = ExtractMergeFieldName(instr);
                        if (name != null) output.Add(name);

                        inField = false;
                        inInstr = false;
                        instrBuilder.Clear();
                    }
                }

                continue;
            }

            if (inField && inInstr)
            {
                // Word may store instructions as InstructionText (w:instrText) and/or FieldCode (w:fldCode)
                foreach (var it in run.Elements<FieldCode>())
                    instrBuilder.Append(it.Text);

                foreach (var fc in run.Elements<FieldCode>())
                    instrBuilder.Append(fc.Text);
            }
        }
    }

    private static string? ExtractMergeFieldName(string instruction)
    {
        if (string.IsNullOrWhiteSpace(instruction)) return null;

        var m = MergeFieldRegex.Match(instruction);
        if (!m.Success) return null;

        var name = m.Groups["name"].Value?.Trim();
        return string.IsNullOrWhiteSpace(name) ? null : name;
    }
}