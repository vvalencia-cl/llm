using LMM.Application;

namespace LMM.Application.Tests;

public class HeaderFieldMapperTests
{
    [Fact]
    public void BuildTemplateToExcelHeaderMap_ExactMatch_Works()
    {
        // Arrange
        var templateFields = new List<string> { "Nombre", "Edad" };
        var excelHeaders = new List<string> { "Nombre", "Edad", "Ciudad" };

        // Act
        var map = HeaderFieldMapper.BuildTemplateToExcelHeaderMap(templateFields, excelHeaders);

        // Assert
        Assert.Equal("Nombre", map["Nombre"]);
        Assert.Equal("Edad", map["Edad"]);
    }

    [Fact]
    public void BuildTemplateToExcelHeaderMap_NormalizedMatch_Works()
    {
        // Arrange
        var templateFields = new List<string> { "Nombre_Completo", "Fecha_Nacimiento" };
        var excelHeaders = new List<string> { "Nombre Completo", "Fecha-Nacimiento" };

        // Act
        var map = HeaderFieldMapper.BuildTemplateToExcelHeaderMap(templateFields, excelHeaders);

        // Assert
        Assert.Equal("Nombre Completo", map["Nombre_Completo"]);
        Assert.Equal("Fecha-Nacimiento", map["Fecha_Nacimiento"]);
    }

    [Fact]
    public void BuildTemplateToExcelHeaderMap_MPrefixMatch_Works()
    {
        // Arrange
        var templateFields = new List<string> { "M_10_Ponderado" };
        var excelHeaders = new List<string> { "10 Ponderado" };

        // Act
        var map = HeaderFieldMapper.BuildTemplateToExcelHeaderMap(templateFields, excelHeaders);

        // Assert
        Assert.Equal("10 Ponderado", map["M_10_Ponderado"]);
    }

    [Fact]
    public void BuildTemplateToExcelHeaderMap_AmbiguousMatch_ThrowsInvalidOperationException()
    {
        // Arrange
        var templateFields = new List<string> { "Test_Field" };
        var excelHeaders = new List<string> { "Test Field", "Test-Field" };

        // Act & Assert
        var ex = Assert.Throws<InvalidOperationException>(() => 
            HeaderFieldMapper.BuildTemplateToExcelHeaderMap(templateFields, excelHeaders));
        Assert.Contains("ambiguos", ex.Message);
    }

    [Fact]
    public void BuildTemplateValuesForRecord_Works()
    {
        // Arrange
        var templateFields = new List<string> { "NameField", "AgeField" };
        var excelRecord = new Dictionary<string, string>
        {
            { "Name", "John Doe" },
            { "Age", "30" }
        };
        var map = new Dictionary<string, string>
        {
            { "NameField", "Name" },
            { "AgeField", "Age" }
        };

        // Act
        var values = HeaderFieldMapper.BuildTemplateValuesForRecord(templateFields, excelRecord, map);

        // Assert
        Assert.Equal("John Doe", values["NameField"]);
        Assert.Equal("30", values["AgeField"]);
    }
}
