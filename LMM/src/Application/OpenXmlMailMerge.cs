using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace LMM.Application;

public static class OpenXmlMailMerge
{
    // Matches: MERGEFIELD FieldName  OR  MERGEFIELD "Field Name"
    private static readonly Regex MergeFieldRegex = new(
        @"\bMERGEFIELD\b\s+(?:(?:""(?<name>[^""]+)"")|(?<name>[^\s\\]+))",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public static void ReplaceMergeFieldsInMainBody(
        WordprocessingDocument doc,
        IReadOnlyDictionary<string, string?> values)
    {
        if (doc?.MainDocumentPart?.Document?.Body == null)
            throw new ArgumentException("Invalid document: MainDocumentPart/Body is missing.");

        var body = doc.MainDocumentPart.Document.Body;

        ReplaceInSimpleFields(body, values);
        ReplaceInComplexFields(body, values);

        doc.MainDocumentPart.Document.Save();
    }

    private static void ReplaceInSimpleFields(OpenXmlElement root, IReadOnlyDictionary<string, string?> values)
    {
        var simpleFields = root.Descendants<SimpleField>().ToList();
        foreach (var sf in simpleFields)
        {
            var instr = sf.Instruction?.Value ?? "";
            var fieldName = ExtractMergeFieldName(instr);
            if (fieldName == null) continue;

            var replacement = GetValueOrEmpty(values, fieldName);

            sf.RemoveAllChildren<Run>();
            sf.AppendChild(CreateRunLike(null, replacement));
        }
    }

    private static void ReplaceInComplexFields(OpenXmlElement root, IReadOnlyDictionary<string, string?> values)
    {
        var paragraphs = root.Descendants<Paragraph>().ToList();

        foreach (var p in paragraphs)
        {
            var runs = p.Elements<Run>().ToList();
            if (runs.Count == 0) continue;

            bool inField = false;
            bool inInstr = false;

            int fieldBeginRunIndex = -1;
            int fieldSeparateRunIndex = -1;
            int fieldEndRunIndex = -1;

            var instrText = "";

            for (int i = 0; i < runs.Count; i++)
            {
                var run = runs[i];

                var fieldChar = run.GetFirstChild<FieldChar>();
                if (fieldChar != null)
                {
                    var t = fieldChar.FieldCharType?.Value; // FieldCharValues? (nullable)
                    if (t == FieldCharValues.Begin)
                    {
                        inField = true;
                        inInstr = true;

                        fieldBeginRunIndex = i;
                        fieldSeparateRunIndex = -1;
                        fieldEndRunIndex = -1;
                        instrText = "";
                    }
                    else if (t == FieldCharValues.Separate)
                    {
                        if (inField)
                        {
                            inInstr = false;
                            fieldSeparateRunIndex = i;
                        }
                    }
                    else if (t == FieldCharValues.End)
                    {
                        if (inField)
                        {
                            fieldEndRunIndex = i;

                            var fieldName = ExtractMergeFieldName(instrText);
                            if (fieldName != null)
                            {
                                var replacement = GetValueOrEmpty(values, fieldName);

                                ReplaceFieldResultRuns(
                                    paragraph: p,
                                    runsSnapshot: runs,
                                    beginIndex: fieldBeginRunIndex,
                                    separateIndex: fieldSeparateRunIndex,
                                    endIndex: fieldEndRunIndex,
                                    replacementText: replacement);
                            }

                            // reset state
                            inField = false;
                            inInstr = false;

                            fieldBeginRunIndex = -1;
                            fieldSeparateRunIndex = -1;
                            fieldEndRunIndex = -1;
                            instrText = "";
                        }
                    }

                    continue;
                }

                // Instruction text can be split across multiple runs/elements
                if (inField && inInstr)
                {
                    foreach (var it in run.Elements<FieldCode>())
                    {
                        instrText += it.Text ?? "";
                    }
                }
            }
        }
    }

    private static void ReplaceFieldResultRuns(
        Paragraph paragraph,
        List<Run> runsSnapshot,
        int beginIndex,
        int separateIndex,
        int endIndex,
        string replacementText)
    {
        int insertAfterIndex = separateIndex >= 0 ? separateIndex : beginIndex;

        // Remove runs between separate and end (exclusive)
        if (separateIndex >= 0)
        {
            int resultStart = separateIndex + 1;
            int resultEndExclusive = endIndex;

            if (resultStart < resultEndExclusive)
            {
                for (int i = resultStart; i < resultEndExclusive; i++)
                    runsSnapshot[i].Remove();
            }
        }

        Run? styleSourceRun = null;

        if (separateIndex >= 0)
        {
            int resultStart = separateIndex + 1;
            if (resultStart >= 0 && resultStart < endIndex && resultStart < runsSnapshot.Count)
                styleSourceRun = runsSnapshot[resultStart];
        }

        if (styleSourceRun == null && beginIndex >= 0 && beginIndex < runsSnapshot.Count)
            styleSourceRun = runsSnapshot[beginIndex];

        var anchorRun = runsSnapshot[insertAfterIndex];
        var newRun = CreateRunLike(styleSourceRun, replacementText);
        anchorRun.InsertAfterSelf(newRun);
    }

    private static Run CreateRunLike(Run? styleSourceRun, string text)
    {
        var run = new Run();

        var rPr = styleSourceRun?.RunProperties?.CloneNode(true) as RunProperties;
        if (rPr != null)
            run.RunProperties = rPr;

        run.AppendChild(new Text(text ?? "") { Space = SpaceProcessingModeValues.Preserve });
        return run;
    }

    private static string GetValueOrEmpty(IReadOnlyDictionary<string, string?> values, string fieldName)
    {
        return values.TryGetValue(fieldName, out var v) && v != null ? v : "";
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