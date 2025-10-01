using System.Drawing;
using ClosedXML.Excel;
using ClosedXmlTutorial.Util;
using System.Globalization;

namespace ClosedXmlTutorial;

public class QuickTutorial
{
    [Test]
    public void BasicUsage()
    {
        using var workbook = new XLWorkbook();
        IXLWorksheet sheet = workbook.AddWorksheet("MySheet");

        // Setting & getting values
        IXLCell firstCell = sheet.Cell(1, 1);
        firstCell.Value = "will it work?";
        sheet.Cell("A2").FormulaA1 = "CONCATENATE(A1,\" ... Of course it will!\")";
        Assert.That(firstCell.GetString(), Is.EqualTo("will it work?"));
        Assert.That(firstCell.Value, Is.EqualTo("will it work?"));
        Assert.That(firstCell.Value.GetText(), Is.EqualTo("will it work?"));
        Assert.That(firstCell.GetText(), Is.EqualTo("will it work?"));

        var calculatedCell = sheet.Cell("A2");
        Assert.That(calculatedCell.CachedValue.IsBlank, Is.True);
        Assert.That(calculatedCell.NeedsRecalculation, Is.True);
        sheet.RecalculateAllFormulas();
        Assert.That(calculatedCell.CachedValue.GetText(), Does.Contain("Of course"));

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
        sheet.Cell("A4").Value = "InsertTable";
        sheet.Cell("A5").InsertTable(data, false);

        sheet.Cell("A11").Value = "InsertData (transpose: true)";
        sheet.Cell("A12").InsertData(data, true);

        sheet.Cell("A15").Value = "InsertData (transpose: false)";
        sheet.Cell("A16").InsertData(data, false);

        // Styling cells
        var someCells = sheet.Cells("A1,A4:B4,A11,A15");
        someCells.Style.Font.Bold = true;
        someCells.Style.Font.SetFontColor(XLColor.Ivory);
        // someCells.Style.Font.FontColor = XLColor.Ivory;
        Assert.That(XLColor.Ivory, Is.EqualTo(XLColor.FromColor(Color.Ivory)));
        // XLColor also has static methods FromArgb, FromHtml, FromKnownColor etc
        someCells.Style.Fill.SetPatternType(XLFillPatternValues.Solid);
        // someCells.Style.Fill.SetBackgroundColor(XLColor.Navy);
        someCells.Style.Fill.BackgroundColor = XLColor.Navy;

        // sheet.Columns().AdjustToContents();
        sheet.ColumnsUsed().AdjustToContents();

        workbook.SaveAs(BinDir.GetPath());

        // BinDir.Save(workbook, true);
    }

    [Test]
    public void LoadingAndSaving()
    {
        // ClosedXML crashes if the files does not exist (EPPlus would open or create)
        string path = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "IDoNotExist.xlsx");
        Assert.Throws<FileNotFoundException>(() =>
        {
            XLWorkbook? xlWorkbook = null;
            try
            {
                xlWorkbook = new XLWorkbook(path);
            }
            finally
            {
                xlWorkbook?.Dispose();
            }
        });

        
        using var workbook = new XLWorkbook();

        // Password Protection
        var sheet = workbook.AddWorksheet("Sheet1");
        sheet.Cells("A1").Value = "Password == 123";
        var protection = sheet.Protect("123");
        Assert.That(protection.IsPasswordProtected, Is.True);

        // While this method exists, it's not possible to password protect an entire workbook.
        workbook.Protect("optionalPassword");
        Assert.That(workbook.IsPasswordProtected, Is.True);


        // Load the worksheets from BasicUsage
        sheet.Cells("D1").Value = "ClosedXML doesn't have package.Load()";
        sheet.Cells("D2").Value = "But we can still add sheets from another workbook!";
        using var anotherWorkbook = new XLWorkbook(BinDir.GetPath(nameof(BasicUsage)));
        foreach (var worksheet in anotherWorkbook.Worksheets)
        {
            workbook.AddWorksheet(worksheet);
        }

        workbook.AddWorksheet("Sheet0", 0);
        sheet.TabSelected = true;
        sheet.TabColor = XLColor.Redwood;
        sheet.ActiveCell = sheet.Cell("A5");
        // sheet.Visibility = XLWorksheetVisibility.Hidden;
        int lastCol = sheet.ColumnsUsed().Last().ColumnNumber();
        sheet.Range(1, 1, 1, lastCol).SetAutoFilter();

        // Copy entire sheet
        sheet.CopyTo(workbook, "Copy");

        BinDir.Save(workbook, false);
    }

    [Test]
    public void SelectingCells()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("MySheet");

        // One cell
        IXLCell cellA2 = sheet.Cell("A2");
        var alsoCellA2 = sheet.Cell(2, 1);
        Assert.That(cellA2.Address.ToString(), Is.EqualTo("A2"));
        Assert.That(cellA2.Address.ToString(XLReferenceStyle.R1C1), Is.EqualTo("R2C1"));
        Assert.That(cellA2.Address, Is.EqualTo(alsoCellA2.Address));

        // Get the column from a cell
        Assert.That(cellA2.Address.ColumnNumber, Is.EqualTo(1));
        Assert.That(sheet.Column(1).RangeAddress.ToString(), Is.EqualTo("A:A"));
        Assert.That(sheet.Column("A").RangeAddress.ToString(), Is.EqualTo("A:A"));

        // A range
        IXLRange ranger = sheet.Range("A2:C5");
        var sameRanger = sheet.Range(2, 1, 5, 3);
        Assert.That(ranger.RangeAddress, Is.EqualTo(sameRanger.RangeAddress));

        var twoCells = sheet.Cells("A1,A4"); // Just A1 and A4
        Assert.That(twoCells.Count(), Is.EqualTo(2));
        Assert.That(twoCells.First().Address.ToString(), Is.EqualTo("A1"));
        Assert.That(twoCells.Last().Address.ToString(), Is.EqualTo("A4"));
        var aRow = sheet.Row(1); // A row
        Assert.That(aRow.RangeAddress.ToString(), Is.EqualTo("1:1"));
        var twoColumns = sheet.Range("A:B"); // Two columns
        Assert.That(twoColumns.RangeAddress.ToString(), Is.EqualTo("A:B"));

        // Linq - ClosedXML uses different approach for comments
        var cellsWithComments = sheet.Range("A1:A5").Cells(cell => cell.HasComment);
        var cellsWithComments2 = sheet.Range("A1:A5").CellsUsed(cell => cell.HasComment);

        // Dimensions used
        Assert.That(sheet.LastRowUsed(), Is.Null);
        Assert.That(sheet.LastColumnUsed(), Is.Null);

        ranger.Value = "pushing";
        var usedRange = sheet.RangeUsed();
        Assert.That(usedRange!.RangeAddress, Is.EqualTo(ranger.RangeAddress));

        // Offset: down 5 rows, right 10 columns
        // Offset only exists in EPPlus
        var movedRanger = sheet.Range(
            ranger.FirstCell().CellBelow(5).CellRight(10),
            ranger.LastCell().CellBelow(5).CellRight(10)
        );
        Assert.That(movedRanger.RangeAddress.ToString(), Is.EqualTo("K7:M10"));
        movedRanger.Value = "Moved";

        // Other
        var lastColumn = sheet.Column("M");
        Assert.That(lastColumn.IsEntireColumn, Is.True);
        lastColumn.Style.Font.Bold = true;

        var lastColumnUsed = sheet.CellsUsed(x => x.Style.Font.Bold);
        lastColumnUsed.Value = "lastBold";

        var lastRow = sheet.Row(10);
        Assert.That(lastRow.IsEntireRow, Is.True);
        var lastCellAddress = lastRow.Intersection(lastColumn); // Also has: Intersects
        lastCellAddress.AsRange()!.Value = "bottomRight";
        Assert.That(lastCellAddress.ToString(), Is.EqualTo("M10:M10"));
        Assert.That(lastCellAddress.ToStringFixed(XLReferenceStyle.A1), Is.EqualTo("$M$10:$M$10"));

        Assert.That(ranger.Intersects(movedRanger), Is.False);
        Assert.That(ranger.Intersection(movedRanger).IsValid, Is.False);
        Assert.That(ranger.Union(movedRanger).Count(), Is.EqualTo(24));

        BinDir.Save(workbook, false);
    }

    [Test]
    public void WritingValues()
    {
        Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("MySheet");

        // Format as text
        sheet.Cell("A1").Style.NumberFormat.Format = "@";

        // Numbers
        sheet.Cell("A1").Value = "Numbers";
        Assert.That(sheet.Cell(1, 1).GetString(), Is.EqualTo("Numbers"));
        sheet.Cell("B1").Value = 15.321;
        sheet.Cell("B1").Style.NumberFormat.Format = "#,##0.00";
        Assert.That(sheet.Cell("B1").GetFormattedString(), Is.EqualTo("15.32"));
        Assert.That(sheet.Cell("B1").GetString(), Is.EqualTo("15.321"));
        Assert.That(sheet.Cell("B1").Value, Is.EqualTo(15.321));

        // Percentage
        sheet.Cell("C1").Value = 0.5;
        sheet.Cell("C1").Style.NumberFormat.Format = "0%";
        Assert.That(sheet.Cell("C1").GetString(), Is.EqualTo("0.5"));
        Assert.That(sheet.Cell("C1").GetFormattedString(), Is.EqualTo("50%"));
        Assert.That(sheet.Cell("C1").Value, Is.EqualTo(0.5));

        // Money
        sheet.Cell("A2").Value = "Moneyz";
        sheet.Cells("B2,D2").Value = 15000.23D;
        sheet.Cells("C2,E2").Value = -2000.50D;
        sheet.Range("B2:C2").Style.NumberFormat.Format = "#,##0.00 [$€-813];[Red]-#,##0.00 [$€-813]";
        sheet.Range("D2:E2").Style.NumberFormat.Format = "[$$-409]#,##0";

        // DateTime
        sheet.Cell("A3").Value = "Timey Wimey";
        sheet.Cell("B3").Style.NumberFormat.Format = "yyyy-mm-dd";
        sheet.Cell("B3").FormulaA1 = $"=DATE({DateTime.Now:yyyy,MM,dd})";
        sheet.Cells("C3,D3").Value = DateTime.Now;
        sheet.Cell("C3").Style.NumberFormat.Format = DateTimeFormatInfo.CurrentInfo.FullDateTimePattern;
        sheet.Cell("D3").Style.NumberFormat.Format = "dd/MM/yyyy HH:mm";

        // Hyperlink
        sheet.Cell("C24").SetHyperlink(new XLHyperlink("https://itenium.be"));
        sheet.Cell("C24").Value = "Visit us";
        //sheet.Cell("C24").Style.Font.FontColor = XLColor.Blue;
        //sheet.Cell("C24").Style.Font.Underline = XLFontUnderlineValues.Single;

        // Internal hyperlink
        workbook.Worksheets.Add("Data");
        sheet.Cell("C26").SetHyperlink(new XLHyperlink("'Data'!A1", "(tooltip)"));
        sheet.Cell("C26").Value = "Link to data sheet";

        sheet.Cell("Z1").Clear();

        sheet.ColumnsUsed().AdjustToContents();
        BinDir.Save(workbook, false);
    }

    [Test]
    public void StylingCells()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Styling");

        // Cells with style
        sheet.Cell("A1").Value = "Bold and proud";
        sheet.Cell("A1").Style.Font.FontName = "Stencil";
        sheet.Cell("A1").Style.Font.Bold = true;
        sheet.Cell("A1").Style.Font.FontColor = XLColor.Green;

        // Borders need to be made
        sheet.Range("A1:A2").Style.Border.OutsideBorder = XLBorderStyleValues.Dotted;
        sheet.Range(5, 5, 9, 8).Style.Border.OutsideBorder = XLBorderStyleValues.Dotted;

        // Merge cells
        sheet.Range(5, 5, 9, 8).Merge();

        // More style
        sheet.Cell("D14").Style.Alignment.ShrinkToFit = true;
        sheet.Cell("D14").Style.Font.FontSize = 24;
        sheet.Cell("D14").Value = "Shrinking for fit";

        sheet.Cell("D15").Style.Alignment.WrapText = true;
        sheet.Cell("D15").Value = "A wrap, yummy!";
        sheet.Cell("D16").Value = "No wrap, ouch!";

        // ATTN: These weren't working for me...
        // sheet.Cell("D16").Style.Font.SetFontFamilyNumbering(XLFontFamilyNumberingValues.Script);
        // sheet.Cell("D16").Style.Font.FontFamilyNumbering = XLFontFamilyNumberingValues.Script;

        // Background color
        sheet.Cell("B5").Style.Fill.BackgroundColor = XLColor.Red;

        // Horizontal Alignment
        sheet.Cell("B5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        sheet.Cell("B5").Value = "I'm centered";

        BinDir.Save(workbook, false);
    }

    [Test]
    public void ConditionalFormatting()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("ConditionalFormatting");

        #region Prepare Data
        sheet.Range("B3:E3").Style.Font.Bold = true;
        sheet.Range("B4:E4").Value = 5;
        sheet.Range("B5:E5").Value = 10;
        sheet.Range("B6:E6").Value = 20;
        sheet.Range("B7:E7").Value = 40;
        sheet.Range("B8:E8").Value = 50;
        sheet.Range("B9:E9").Value = 30;
        #endregion

        sheet.Cell("B3").Value = "WhenBetween";
        sheet
            .Range("B4:B9")
            .AddConditionalFormat()
            .WhenBetween(0, 20)
            .Fill.SetBackgroundColor(XLColor.Red);

        sheet.Cell("C3").Value = "ColorScale";
        sheet
            .Range("C4:C9")
            .AddConditionalFormat()
            .ColorScale()
            .LowestValue(XLColor.Red)
            .Midpoint(XLCFContentType.Percent, 50, XLColor.Yellow)
            .HighestValue(XLColor.Green);

        sheet.Cell("D3").Value = "IconSet";
        sheet
            .Range("D4:D9")
            .AddConditionalFormat()
            .IconSet(XLIconSetStyle.ThreeTrafficLights2)
            .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, 0, XLCFContentType.Number)
            .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, 20, XLCFContentType.Number)
            .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, 40, XLCFContentType.Number);

        sheet.Cell("E3").Value = "DataBar";
        sheet
            .Range("E4:E9")
            .AddConditionalFormat()
            .DataBar(XLColor.Red)
            .LowestValue()
            .HighestValue();


        sheet.ColumnsUsed().AdjustToContents();
        BinDir.Save(workbook, false);
    }

    [Test]
    public void FormattingSheetsAndColumns()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Victor");
        sheet.TabColor = XLColor.Ivory;

        // Freeze the top row and left 4 columns
        sheet.SheetView.FreezeRows(1);
        sheet.SheetView.FreezeColumns(4);

        sheet.ShowGridLines = false;
        sheet.ShowRowColHeaders = false;

        // Default selected cell when opening
        sheet.Cell("B6").SetActive();
        sheet.Cell("B6").Select();

        var colE = sheet.Column(5);
        colE.AdjustToContents(); // or colE.Width = value

        // Hide column A
        sheet.Column(1).Hide();

        BinDir.Save(workbook, false);
    }



    [Test]
    public void OutlineSymbolsExample()
    {
        using var workbook = new XLWorkbook();

        CreateSheet(true);
        CreateSheet(false);
        BinDir.Save(workbook, false);

        void CreateSheet(bool showOutlineSymbols)
        {
            var sheet = workbook.Worksheets.Add($"Outline{(showOutlineSymbols ? "True" : "False")}");

            #region Sheet Content
            // Add sample data with headers
            sheet.Cell("A1").Value = "Department";
            sheet.Cell("B1").Value = "Employee";
            sheet.Cell("C1").Value = "Salary";

            // Sales Department
            sheet.Cell("A2").Value = "Sales";
            sheet.Cell("B3").Value = "John Doe";
            sheet.Cell("C3").Value = 50000;
            sheet.Cell("B4").Value = "Jane Smith";
            sheet.Cell("C4").Value = 55000;

            // Marketing Department
            sheet.Cell("A5").Value = "Marketing";
            sheet.Cell("B6").Value = "Bob Johnson";
            sheet.Cell("C6").Value = 48000;
            sheet.Cell("B7").Value = "Alice Brown";
            sheet.Cell("C7").Value = 52000;

            // Create row groups to generate outline symbols
            sheet.Rows("3:4").Group();  // Group Sales employees
            sheet.Rows("6:7").Group();  // Group Marketing employees
            sheet.Rows("2:7").Group();  // Group all departments

            // Create column group
            sheet.Columns("B:C").Group(); // Group employee details

            sheet.Range("A1:C1").Style.Font.Bold = true;
            sheet.Cells("A2,A5").Style.Font.Bold = true;
            #endregion
            
            sheet.ShowOutlineSymbols = showOutlineSymbols;

            sheet.ColumnsUsed().AdjustToContents();
        }
    }
}
