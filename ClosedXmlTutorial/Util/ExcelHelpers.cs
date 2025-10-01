using ClosedXML.Excel;

namespace ClosedXmlTutorial.Util;

public static class ExcelHelpers
{
    public static void SetHeaders(this IXLCell cell, params string[] headers)
    {
        foreach (string text in headers)
        {
            cell.Value = text;
            cell.Style.Font.Bold = true;
            cell.Style.Fill.PatternType = XLFillPatternValues.Solid;
            cell.Style.Fill.BackgroundColor = XLColor.DarkBlue;
            cell.Style.Font.FontColor = XLColor.White;

            cell = cell.CellRight();
        }
    }
}