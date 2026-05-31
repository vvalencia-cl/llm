using ClosedXML.Excel;
using LMM.Application;

namespace LMM.Application.Tests;

public class ClosedXmlDataSourceServiceTests : IDisposable
{
    private readonly string _xlsxPath;

    public ClosedXmlDataSourceServiceTests()
    {
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"TestExcel_{Guid.NewGuid()}.xlsx");
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Cell(1, 1).Value = "Header1";
        ws.Cell(1, 2).Value = "Header2";
        ws.Cell(2, 1).Value = "Value1";
        ws.Cell(2, 2).Value = "Value2";
        wb.SaveAs(_xlsxPath);
    }

    public void Dispose()
    {
        if (File.Exists(_xlsxPath))
            File.Delete(_xlsxPath);
    }

    [Fact]
    public void GetWorksheetNames_ReturnsCorrectNames()
    {
        // Act
        using var service = new ClosedXmlDataSourceService(_xlsxPath);
        var names = service.GetWorksheetNames();

        // Assert
        Assert.Single(names);
        Assert.Equal("Sheet1", names[0]);
    }

    [Fact]
    public void ReadHeaders_HappyPath_Works()
    {
        // Act
        using var service = new ClosedXmlDataSourceService(_xlsxPath);
        var result = service.ReadHeaders("Sheet1", 1);

        // Assert
        Assert.Equal(2, result.Headers.Count);
        Assert.Equal("Header1", result.Headers[0]);
        Assert.Equal("Header2", result.Headers[1]);
    }

    [Fact]
    public void ReadHeaders_NonExistentWorksheet_ThrowsInvalidOperationException()
    {
        // Act & Assert
        using var service = new ClosedXmlDataSourceService(_xlsxPath);
        Assert.Throws<InvalidOperationException>(() => service.ReadHeaders("NonExistent", 1));
    }

    [Fact]
    public void EnumerateRecordsFormatted_Works()
    {
        // Act
        using var service = new ClosedXmlDataSourceService(_xlsxPath);
        var headers = new List<string> { "Header1", "Header2" };
        var records = service.EnumerateRecordsFormatted("Sheet1", 1, headers).ToList();

        // Assert
        Assert.Single(records);
        Assert.Equal("Value1", records[0]["Header1"]);
        Assert.Equal("Value2", records[0]["Header2"]);
    }
}
