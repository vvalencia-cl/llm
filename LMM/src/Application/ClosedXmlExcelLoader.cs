using ClosedXML.Excel;
using LMM.Domain.Dto;

namespace LMM.Application;

public static class ClosedXmlExcelLoader
{
    public static ExcelLoadResult LoadWorksheetNames(string xlsxPath)
    {
        using var wb = new XLWorkbook(xlsxPath);

        var sheetNames = wb.Worksheets.Select(ws => ws.Name).ToList();
        if (sheetNames.Count == 0)
            throw new InvalidOperationException("El archivo Excel no contiene hojas de trabajo.");

        return new ExcelLoadResult(
            WorksheetNames: sheetNames,
            DefaultWorksheetName: sheetNames[0]);
    }
    
}