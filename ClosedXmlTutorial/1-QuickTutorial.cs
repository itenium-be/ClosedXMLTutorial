using System.Drawing;
using ClosedXML.Excel;
using ClosedXmlTutorial.Util;

namespace ClosedXmlTutorial;

public class QuickTutorial
{
    [Test]
    public void BasicUsage()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("MySheet");

        // Setting & getting values
        IXLCell firstCell = sheet.Cell(1, 1);
        firstCell.Value = "will it work?";
        sheet.Cell("A2").FormulaA1 = "CONCATENATE(A1,\" ... Of course it will!\")";
        Assert.That(firstCell.GetString(), Is.EqualTo("will it work?"));

        // Numbers
        var moneyCell = sheet.Cell("A3");
        moneyCell.Style.NumberFormat.Format = "$#,##0.00";
        moneyCell.Value = 1500.25M;

        moneyCell = sheet.Cell("B3");
        moneyCell.Style.NumberFormat.NumberFormatId = (int)XLPredefinedFormat.Number.Precision2WithSeparator;
        moneyCell.Value = 1500.25M;

        // Easily write any Enumerable to a sheet
        // ClosedXml can't list all implemented functions, but they are listed here:
        // https://github.com/closedxml/closedxml/wiki/Evaluating-Formulas#supported-functions
        var data = new[]
        {
            new { FunctionName = "DATE", Description = "Returns the serial number of a particular date" },
            new { FunctionName = "YEAR", Description = "Converts a serial number to a year" },
            new { FunctionName = "CHAR", Description = "Returns the character specified by the code number" },
            new { FunctionName = "FIND", Description = "Finds one text value within another (case-sensitive)" },
        };
        sheet.Cell("A4").InsertTable(data, true);

        // Styling cells
        var someCells = sheet.Cells("A1,A4:B4");
        someCells.Style.Font.Bold = true;
        someCells.Style.Font.SetFontColor(XLColor.Ivory);
        // someCells.Style.Font.FontColor = XLColor.Ivory;
        Assert.That(XLColor.Ivory, Is.EqualTo(XLColor.FromColor(Color.Ivory)));
        // XLColor also has static methods FromArgb, FromHtml, FromKnownColor etc
        someCells.Style.Fill.SetPatternType(XLFillPatternValues.Solid);
        // someCells.Style.Fill.SetBackgroundColor(XLColor.Navy);
        someCells.Style.Fill.BackgroundColor = XLColor.Navy;

        sheet.Columns().AdjustToContents();
        workbook.SaveAs(BinDir.GetPath());
    }
}
