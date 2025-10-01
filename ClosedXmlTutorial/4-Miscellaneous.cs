using ClosedXML.Excel;
using ClosedXmlTutorial.Util;

namespace ClosedXmlTutorial;

public class Miscellaneous
{
    [Test]
    public void ExcelPrinting()
    {
        // More advanced printing options:
        // https://github.com/closedxml/closedxml/wiki#page-setup-print-options

        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Printing");
        sheet.Cell("A1").Value = "Check the print preview (Ctrl+P)";

        sheet.PageSetup.Header.Center.AddText("YourTitle", XLHFOccurrence.AllPages)
            .SetFontSize(24)
            .SetBold()
            .SetUnderline();
        sheet.PageSetup.Header.Right.AddText(XLHFPredefinedText.Date, XLHFOccurrence.AllPages);
        sheet.PageSetup.Header.Left.AddText(XLHFPredefinedText.SheetName, XLHFOccurrence.AllPages);

        sheet.PageSetup.Footer.Right.AddText("Page ", XLHFOccurrence.AllPages);
        sheet.PageSetup.Footer.Right.AddText(XLHFPredefinedText.PageNumber, XLHFOccurrence.AllPages);
        sheet.PageSetup.Footer.Right.AddText(" of ", XLHFOccurrence.AllPages);
        sheet.PageSetup.Footer.Right.AddText(XLHFPredefinedText.NumberOfPages, XLHFOccurrence.AllPages);
        sheet.PageSetup.Footer.Center.AddText(XLHFPredefinedText.SheetName, XLHFOccurrence.AllPages);
        sheet.PageSetup.Footer.Left.AddText(XLHFPredefinedText.Path, XLHFOccurrence.AllPages);
        sheet.PageSetup.Footer.Left.AddText(XLHFPredefinedText.File, XLHFOccurrence.AllPages);

        // See: https://github.com/closedxml/closedxml/wiki/Paper-Size-Lookup-Table
        sheet.PageSetup.PaperSize = XLPaperSize.A4Paper;

        // sheet.PageSetup.Margins

        sheet.PageSetup.SetRowsToRepeatAtTop(1, 2);
        sheet.PageSetup.SetColumnsToRepeatAtLeft(1, 7);

        sheet.SheetView.View = XLSheetViewOptions.PageLayout;

        BinDir.Save(workbook, false);
    }

    [Test]
    public void ConvertingIndexesAndAddresses()
    {
        Assert.That(XLHelper.GetColumnLetterFromNumber(1) + "1", Is.EqualTo("A1"));
        Assert.That(XLHelper.IsValidA1Address("A5"), Is.True);


        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Multiple Ranges");

        ws.Ranges("A1:B2,C3:D4,E5:F6").Style.Fill.BackgroundColor = XLColor.Red;
        ws.Ranges("A5:B6,E1:F2").Style.Fill.BackgroundColor = XLColor.Orange;

        ws.Ranges("A5,E1").Style.Fill.BackgroundColor = XLColor.Black;

        // https://github.com/closedxml/closedxml/wiki/Named-Ranges
        // https://github.com/closedxml/closedxml/wiki/Accessing-Named-Ranges

        BinDir.Save(workbook, false);
    }

    [Test]
    public void SettingWorkbookProperties()
    {
        // https://github.com/closedxml/closedxml/wiki/Workbook-Properties

        using var workbook = new XLWorkbook();
        // XLWorkbookProperties props = workbook.Properties;
        workbook.Properties.Title = "ClosedXML Tutorial Series";
        workbook.Properties.Author = "Wouter Van Schandevijl";
        workbook.Properties.Comments = "";
        workbook.Properties.Keywords = "";
        workbook.Properties.Category = "";

        workbook.Properties.Company = "itenium";
        workbook.CustomProperties.Add("Checked by", "Jan Havlíček");

        workbook.Worksheets.Add("Sheet1");

        BinDir.Save(workbook, false);
    }

    [Test]
    public void AddingComments()
    {
        // https://github.com/closedxml/closedxml/wiki#comments

        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Comments");

        var comment = sheet.Cell("D4").CreateComment();
        comment.Author = "evil corp";
        comment.AddSignature();
        comment.AddText("Bold title:\r\n").SetBold();
        comment.AddText("Unbolded subtext").SetBold(false);

        sheet.Cell("G4").CreateComment().Style.Alignment.SetAutomaticSize();
        sheet.Cell("G4").GetComment().AddText("Things are pretty tight around here");

        sheet.Cell("L4").CreateComment().AddText("Orientation = Vertical");
        sheet.Cell("L4").GetComment().Style
            .Alignment.SetOrientation(XLDrawingTextOrientation.Vertical)
            .Alignment.SetAutomaticSize()
            .Margins.SetAll(0.25)
            .ColorsAndLines.SetFillColor(XLColor.RichCarmine)
            .ColorsAndLines.SetFillTransparency(0.25);

        sheet.Cell("A1").CreateComment().AddText("This is an unusual place for a comment...");
        sheet.Cell("A1").GetComment()
            .Position
            .SetColumn(3)
            .SetColumnOffset(5) // The comment will start in the middle of the column
            .SetRow(10)
            .SetRowOffset(7.5); // The comment will start in the middle of the row

        foreach (var cell in sheet.CellsUsed(XLCellsUsedOptions.Comments, c => c.HasComment))
        {
            cell.GetComment().SetVisible();
        }

        BinDir.Save(workbook, false);
    }

    [Test]
    public void CommentsFromWiki()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Alignment");

        // Automagically adjust the size of the comment to fit the contents
        ws.Cell("A1").CreateComment().Style.Alignment.SetAutomaticSize();
        ws.Cell("A1").GetComment().AddText("Things are pretty tight around here");

        // Default values
        ws.Cell("A3").CreateComment()
            .AddText("Default Alignments:").AddNewLine()
            .AddText("Vertical = Top").AddNewLine()
            .AddText("Horizontal = Left").AddNewLine()
            .AddText("Orientation = Left to Right");

        // Let's change the alignments
        ws.Cell("A8").CreateComment()
            .AddText("Vertical = Bottom").AddNewLine()
            .AddText("Horizontal = Right");
        ws.Cell("A8").GetComment().Style
            .Alignment.SetVertical(XLDrawingVerticalAlignment.Bottom)
            .Alignment.SetHorizontal(XLDrawingHorizontalAlignment.Right);

        // And now the orientation...
        ws.Cell("D3").CreateComment().AddText("Orientation = Bottom to Top");
        ws.Cell("D3").GetComment().Style
            .Alignment.SetOrientation(XLDrawingTextOrientation.BottomToTop)
            .Alignment.SetAutomaticSize();

        ws.Cell("E3").CreateComment().AddText("Orientation = Top to Bottom");
        ws.Cell("E3").GetComment().Style
            .Alignment.SetOrientation(XLDrawingTextOrientation.TopToBottom)
            .Alignment.SetAutomaticSize();

        ws.Cell("F3").CreateComment().AddText("Orientation = Vertical");
        ws.Cell("F3").GetComment().Style
            .Alignment.SetOrientation(XLDrawingTextOrientation.Vertical)
            .Alignment.SetAutomaticSize();

        // Set all comments to visible
        foreach (var cell in ws.CellsUsed(XLCellsUsedOptions.Comments, c => c.HasComment))
        {
            cell.GetComment().SetVisible();
        }

        wb.SaveAs("CommentsAlignment.xlsx");

        BinDir.Save(wb, false);
    }

    [Test]
    public void RichText()
    {
        // https://github.com/closedxml/closedxml/wiki/Using-Rich-Text

        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Rich Text");

        var cell1 = ws.Cell(1, 1).SetValue("The show must go on...");
        cell1.Style.Font.FontColor = XLColor.Blue; // Set the color for the entire cell
        cell1.CreateRichText().Substring(4, 4)
            .SetFontColor(XLColor.Red)
            .SetFontName("Broadway"); // Set the color and font for the word "show"

        var cell = ws.Cell(3, 1);
        cell.CreateRichText()
            .AddText("Hello").SetFontColor(XLColor.Red)
            .AddText(" BIG ").SetFontColor(XLColor.Blue).SetBold()
            .AddText("World").SetFontColor(XLColor.Red);

        BinDir.Save(wb, false);
    }

    [Test]
    public void PasswordProtectionFromManualEditing()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Secret");

        sheet.Cell("D5").Value = "Can't touch this";
        sheet.Cell("D5").Style.Protection.Locked = false;

        sheet.Protection.AllowedElements = XLSheetProtectionElements.SelectLockedCells |
                                            XLSheetProtectionElements.SelectUnlockedCells;
        sheet.Protect("Secret");

        BinDir.Save(workbook, false);
    }
}