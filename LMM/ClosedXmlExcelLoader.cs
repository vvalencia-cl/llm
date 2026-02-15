using System;
using System.Linq;
using ClosedXML.Excel;

public static class ClosedXmlExcelLoader
{
    public static ExcelLoadResult LoadWorksheetNames(string xlsxPath)
    {
        using var wb = new XLWorkbook(xlsxPath);

        var sheetNames = wb.Worksheets.Select(ws => ws.Name).ToList();
        if (sheetNames.Count == 0)
            throw new InvalidOperationException("The Excel file contains no worksheets.");

        return new ExcelLoadResult(
            WorksheetNames: sheetNames,
            DefaultWorksheetName: sheetNames[0]);
    }
}