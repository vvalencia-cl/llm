using LMM.Application;

namespace LMM.Application.Tests;

public class PdfFilenameBuilderTests : IDisposable
{
    private readonly string _tempDir;

    public PdfFilenameBuilderTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "LMMTests_" + Guid.NewGuid());
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, true);
    }

    [Fact]
    public void BuildPdfPath_HappyPath_Works()
    {
        // Arrange
        var record = new Dictionary<string, string>
        {
            { "ID", "123" },
            { "Name", "John Doe" }
        };

        // Act
        var result = PdfFilenameBuilder.BuildPdfPath(
            _tempDir,
            record,
            prefix: "REC",
            firstFieldHeader: "ID",
            secondFieldHeader: "Name",
            separator: "-"
        );

        // Assert
        var expected = Path.Combine(_tempDir, "REC-123-John Doe.pdf");
        Assert.Equal(expected, result);
    }

    [Fact]
    public void BuildPdfPath_InvalidOutputDirectory_ThrowsDirectoryNotFoundException()
    {
        // Arrange
        var record = new Dictionary<string, string>();
        var invalidDir = Path.Combine(_tempDir, "NonExistentDir");

        // Act & Assert
        Assert.Throws<DirectoryNotFoundException>(() =>
            PdfFilenameBuilder.BuildPdfPath(invalidDir, record));
    }

    [Fact]
    public void SanitizeFilenamePart_RemovesInvalidChars()
    {
        // Act
        var result = PdfFilenameBuilder.SanitizeFilenamePart("Invalid/Char*Name?");

        // Assert
        Assert.Equal("Invalid Char Name", result);
    }

    [Fact]
    public void SanitizeFilenamePart_HandlesReservedNames()
    {
        // Act
        var result = PdfFilenameBuilder.SanitizeFilenamePart("CON");

        // Assert
        Assert.Equal("_CON", result);
    }
}
