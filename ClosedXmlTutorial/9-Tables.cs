using ClosedXML.Excel;

namespace ClosedXmlTutorial;

public class Tables
{
    [Test]
    public void Test()
    {
        // https://docs.closedxml.io/en/latest/features/tables.html
        // https://github.com/closedxml/closedxml/wiki/Inserting-Data
        // https://github.com/closedxml/closedxml/wiki/Inserting-Tables

        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Table");

        var rngData = ws.Range("B3:F6");
        var excelTable = rngData.CreateTable();

        // Add the totals row
        excelTable.ShowTotalsRow = true;
        // Put the average on the field "Income"
        // Notice how we're calling the cell by the column name
        excelTable.Field("Income").TotalsRowFunction = XLTotalsRowFunction.Average;
        // Put a label on the totals cell of the field "DOB"
        excelTable.Field("DOB").TotalsRowLabel = "Average:";
    }

    [Test]
    public void Pivoting()
    {
        // https://github.com/closedxml/closedxml/wiki/Pivot-Table-example
        // https://github.com/closedxml/closedxml/wiki/Transpose-Ranges
    }
}