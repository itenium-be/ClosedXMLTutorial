using ClosedXML.Excel;
using ClosedXmlTutorial.Util;

namespace ClosedXmlTutorial;

public class FormulasAndDataValidation
{
    [Test]
    public void BasicFormulas()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Formula");

        // SetHeaders is an extension method - you'll need to update the extension method too
        sheet.Cell("A1").SetHeaders("Product", "Quantity", "Price", "Base total", "Discount", "Total", "Special discount", "Payup");
        sheet.Cell("H5").CreateComment().AddText("Special discount for our most valued customers");

        // Turn filtering on for the headers
        var headerRange = sheet.Range(1, 1, 1, sheet.RangeUsed()!.LastColumn().ColumnNumber());
        headerRange.SetAutoFilter();

        var data = AddThreeRowsDataAndFormat(sheet);

        // Can start formulas with = or not
        sheet.Cell("A5").FormulaA1 = "COUNTA(A2:A4)";
        Assert.That(sheet.Cell("A5").Value, Is.EqualTo(3));
        Assert.That(sheet.Evaluate("COUNTA(A2:A4)"), Is.EqualTo(3));

        // Hide the formula (when the sheet is protected)
        sheet.Cell("A5").Style.Protection.Hidden = true;
        // sheet.Protect("123");

        // Total column
        // EPPlus would make that B3*C3 for D3 but ClosedXML does not
        // sheet.Range("D2:D4").FormulaA1 = "B2*C2";
        // Need to use FormulaR1C1 instead!
        sheet.Range("D2:D4").FormulaR1C1 = "RC[-2]*RC[-1]";
        Assert.That(sheet.Cell("D2").FormulaR1C1, Is.EqualTo("RC[-2]*RC[-1]"));
        Assert.That(sheet.Cell("D4").FormulaR1C1, Is.EqualTo("RC[-2]*RC[-1]"));


        // Total - discount column
        // Calculate formulas before they are available in the sheet
        // (Opening an Excel with Office will do this automatically)
        sheet.Range("F2:F4").FormulaR1C1 = "IF(ISBLANK(RC[-1]),RC[-2],RC[-2]*(1-RC[-1]))";
        Assert.That(sheet.Cell("F2").CachedValue.ToString(), Is.Empty);
        workbook.RecalculateAllFormulas();
        Assert.That(sheet.Cell("F2").CachedValue.ToString(), Is.Not.Empty);

        // Total row
        sheet.Cell("D5").FormulaR1C1 = "SUBTOTAL(9,R[-3]C:R[-1]C)"; // total
        Assert.That(sheet.Cell("D5").FormulaA1, Is.EqualTo("SUBTOTAL(9,D2:D4)"));
        sheet.Cell("F5").FormulaR1C1 = "SUBTOTAL(9,R[-3]C:R[-1]C)"; // total - discount
        Assert.That(sheet.Cell("F5").FormulaA1, Is.EqualTo("SUBTOTAL(9,F2:F4)"));

        workbook.RecalculateAllFormulas();
        sheet.Range("H2:H5").FormulaR1C1 = "RC[-2]*(1-R5C7)"; // R5C7 is G5

        // SUBTOTAL(9 = SUM) // 109 = Sum excluding manually hidden rows
        // AVERAGE (1), COUNT (2), COUNTA (3), MAX (4), MIN (5)
        // PRODUCT (6), STDEV (7), STDEVP (8), SUM (9), VAR (10)

        sheet.Columns().AdjustToContents();

        // Evaluate all dirty formulas when saving
        // workbook.SaveAs(BinDir.GetPath(), new SaveOptions() {EvaluateFormulasBeforeSaving = true});

        BinDir.Save(workbook, false);
    }

    private static Sell[] AddThreeRowsDataAndFormat(IXLWorksheet sheet)
    {
        var data = new SalesGenerator().Generate(3).ToArray();
        // See 3-Import for more about InsertData
        sheet.Cell("A2").InsertData(data);

        // Special discount
        sheet.Cell("G5").Value = 0.2;
        sheet.Cell("G5").Style.NumberFormat.Format = "0%";

        // Formatting is covered in 1-QuickTutorial
        sheet.Range("B2:B5").Style.NumberFormat.Format = "#,##0"; // number
        sheet.Cells("C2:D5,F2:F5,H2:H5").Style.NumberFormat.Format = "[$$-409]#,##0.00"; // money
        sheet.Range("E2:E5").Style.NumberFormat.Format = "0%"; // percentage

        // Border above the totals row
        var lastCell = sheet.RangeUsed()!.RangeAddress.LastAddress;
        sheet.Range(lastCell.RowNumber, 1, lastCell.RowNumber, lastCell.ColumnNumber).Style.Border.TopBorder = XLBorderStyleValues.Double;

        return data;
    }

    [Test]
    public void DataValidation_DropDownComboCell()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Validation");

        sheet.Cell("C6").SetHeaders("DROPDOWN");
        sheet.Column("C").Width = 20;

        var validation = sheet.Cell("C7").CreateDataValidation();
        validation.List("\"Apples,Oranges,Lemons\"");

        validation.ErrorStyle = XLErrorStyle.Stop;
        validation.ErrorTitle = "Invalid Selection";
        validation.ErrorMessage = "We only have those available :(";
        validation.ShowErrorMessage = true;

        validation.InputTitle = "Choose your juice";
        validation.InputMessage = "Apples, oranges or lemons?";
        validation.ShowInputMessage = true;

        validation.IgnoreBlanks = true;

        sheet.Cell("C7").Value = "Pick";
        sheet.Cell("C7").Select();
        BinDir.Save(workbook, false);
    }

    [Test]
    public void DataValidation_FromOtherSheet()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Validation");

        sheet.Cell("C6").SetHeaders("DROPDOWN");
        sheet.Column("C").Width = 20;

        var otherSheet = workbook.Worksheets.Add("OtherSheet");
        otherSheet.Cell("A1").Value = "Kwan";
        otherSheet.Cell("A2").Value = "Nancy";
        otherSheet.Cell("A3").Value = "Tonya";

        var validation = sheet.Cell("C7").CreateDataValidation();
        validation.List("OtherSheet!A1:A4");
        sheet.Cell("C7").Select();

        BinDir.Save(workbook, false);
    }

    [Test]
    public void DataValidation_IntAndDateTime()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("intsAndSuch");

        sheet.Cell("A1").SetHeaders("Integer 1-5", "Date > today", "Time > 13:30:10");

        // Integer validation
        var intValidation = sheet.Cell("A2").CreateDataValidation();
        intValidation.WholeNumber.Between(1, 5);
        // intValidation.Decimal.Between(1, 5); (or for decimals)
        intValidation.InputMessage = "Value between 1 and 5";
        intValidation.ShowInputMessage = true;
        intValidation.ErrorStyle = XLErrorStyle.Warning;
        intValidation.ErrorTitle = "Number out of range";
        sheet.Cell("A2").Select();
        Assert.That(sheet.DataValidations.First().ErrorTitle, Is.EqualTo("Number out of range"));

        // DateTime validation
        var dateTimeValidation = sheet.Cell("B2").CreateDataValidation();
        dateTimeValidation.Date.GreaterThan(DateTime.Now.Date);
        dateTimeValidation.InputMessage = "A date greater than today";
        dateTimeValidation.ShowInputMessage = true;

        // Time validation
        // ATTN: While it was able to save it, I couldn't open it in Excel
        //var timeValidation = sheet.Cell("C2").CreateDataValidation();
        //var timeSpan = new TimeSpan(13, 30, 10);
        //timeValidation.Time.GreaterThan(timeSpan);
        sheet.Cell("C2").Value = "ClosedXML BUG?";

        sheet.ColumnsUsed().AdjustToContents();
        BinDir.Save(workbook, false);
    }
}