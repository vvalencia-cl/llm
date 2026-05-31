using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace LMM.Application.Tests;

public class OpenXmlMailMergeTests
{
    [Fact]
    public void ReplaceMergeFieldsInMainBody_HappyPath_Works()
    {
        // Arrange
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(
                new Body(
                    new Paragraph(
                        new Run(new Text("Hello ")),
                        new SimpleField { Instruction = " MERGEFIELD Name " },
                        new Run(new Text("!"))
                    )
                )
            );
            doc.Save();
        }

        stream.Position = 0;
        var values = new Dictionary<string, string?> { { "Name", "World" } };

        // Act
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            OpenXmlMailMerge.ReplaceMergeFieldsInMainBody(doc, values);
        }

        // Assert
        stream.Position = 0;
        using (var doc = WordprocessingDocument.Open(stream, false))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            var text = body.InnerText;
            Assert.Equal("Hello World!", text);
        }
    }

    [Fact]
    public void ReplaceMergeFieldsInMainBody_InvalidDoc_ThrowsArgumentException()
    {
        // Act & Assert
        Assert.Throws<ArgumentException>(() =>
            OpenXmlMailMerge.ReplaceMergeFieldsInMainBody(null!, new Dictionary<string, string?>()));
    }
}