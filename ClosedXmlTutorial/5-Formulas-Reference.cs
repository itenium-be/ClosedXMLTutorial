using ClosedXML.Excel;
using ClosedXmlTutorial.Util;
using System.Globalization;
using System.Text.RegularExpressions;

namespace ClosedXmlTutorial;

/// <summary>
/// Functions Docs:
/// https://docs.closedxml.io/en/latest/features/functions.html
/// 
/// Assign a formula with either
/// - .FormulaA1 = "A$5"
/// - .FormulaR1C1 = "RC[-2]*RC[-1]"
/// 
/// Note:
/// - Formula may start with '='
/// - Use English function names
/// - Use , as function argument separator
/// </summary>
public class FormulasReference
{
    private const string Fox = "The quick brown fox jumps over the lazy dog";

    [Test]
    public void StringManipulation()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("StringManipulation");

        sheet.Cell("A1").Value = Fox;
        sheet.Cell("A2").Value = "=LEN(A1)";
        sheet.Cell("B2").FormulaA1 = "=LEN(A1)";
        sheet.Assert("B2", Is.EqualTo(Fox.Length));

        sheet.Cell("A3").Value = "=UPPER(A1)";
        sheet.Cell("B3").FormulaA1 = "UPPER(A1)";
        sheet.Assert("B3", Is.EqualTo(Fox.ToUpper()));
        sheet.Cell("C3").FormulaA1 = "PROPER(A1)";
        sheet.Assert("C3", Is.EqualTo("The Quick Brown Fox Jumps Over The Lazy Dog"));

        sheet.Cell("A4").Value = "=LEFT(A1; 3)";
        sheet.Cell("B4").FormulaA1 = "LEFT(A1, 3)";
        sheet.Assert("B4", Is.EqualTo(Fox[..3]));

        sheet.Cell("A5").Value = "=MID(A1; 5; 5)";
        sheet.Cell("B5").FormulaA1 = "MID(A1, 5, 5)";
        sheet.Assert("B5", Is.EqualTo(Fox.Substring(4, 5)));

        sheet.Cell("A6").Value = "=REPLACE(A1; 1; 3; \"A\")";
        sheet.Cell("B6").FormulaA1 = "REPLACE(A1, 1, 3, \"A\")";
        sheet.Assert("B6", Is.EqualTo("A quick brown fox jumps over the lazy dog"));

        sheet.Cell("A7").Value = "=SUBSTITUTE(LOWER(A1); \"the\"; \"a\")";
        sheet.Cell("B7").FormulaA1 = "SUBSTITUTE(LOWER(A1), \"the\", \"a\")";
        sheet.Assert("B7", Is.EqualTo(Regex.Replace(Fox, "the", "a", RegexOptions.IgnoreCase)));

        sheet.Cell("A8").Value = "=REPT(A1; 1; 3; \"A\")";
        sheet.Cell("B8").FormulaA1 = "REPT(\"A\", 3)";
        sheet.Assert("B8", Is.EqualTo("AAA"));

        sheet.Cell("A9").Value = "=CONCATENATE(A1; \" over and\"; \" over again\")";
        sheet.Cell("B9").FormulaA1 = "CONCATENATE(A1, \" over and over again\")";
        sheet.Assert("B9", Is.EqualTo(Fox + " over and over again"));

        sheet.Cell("J9").Value = "=B4 & \" \" & B5";
        sheet.Column("J").Width = 15;
        sheet.Cell("K9").FormulaA1 = "B4 & \" \" & B5";
        sheet.Assert("K9", Is.EqualTo("The quick"));

        sheet.Cell("A10").Value = "=FIND(\"fox\"; A1)";
        sheet.Cell("B10").FormulaA1 = "FIND(\"fox\", A1)";
        sheet.Assert("B10", Is.EqualTo(Fox.IndexOf("fox", StringComparison.InvariantCulture) + 1));

        // FIND returns #VALUE! if not found
        sheet.Cell("C10").FormulaA1 = "FIND(\"FOX\", A1)";
        Assert.That(sheet.Cell("C10").Value.IsError, Is.True);
        Assert.That(sheet.Cell("C10").GetError(), Is.EqualTo(XLError.IncompatibleValue));
        sheet.Cell("D10").Value = "=ISERR(C10)";
        sheet.Column("D").Width = 15;
        sheet.Cell("E10").FormulaA1 = "ISERR(C10)"; // TRUE for any error except #N/A
        sheet.Assert("E10", Is.EqualTo(true));
        sheet.Cell("F10").FormulaA1 = "ISERROR(C10)"; // TRUE for any error including #N/A
        sheet.Assert("F10", Is.EqualTo(true));

        sheet.Cell("A11").Value = "=SEARCH(\"FOX\", A1)";
        sheet.Cell("B11").FormulaA1 = "SEARCH(\"FOX\", A1)";
        sheet.Assert("B11", Is.EqualTo(Fox.IndexOf("fox", StringComparison.InvariantCulture) + 1));

        sheet.Cell("A12").Value = "=T(A1)";
        sheet.Cell("B12").FormulaA1 = "T(A1)";
        sheet.Assert("B12", Is.EqualTo(Fox));

        sheet.Cell("A13").Value = $"=EXACT(A1, \"{Fox}\")";
        sheet.Cell("B13").FormulaA1 = $"EXACT(A1, \"{Fox}\")";
        sheet.Assert("B13", Is.EqualTo(true));

        sheet.Cell("A14").Value = "=ISBLANK(F1)";
        sheet.Cell("B14").FormulaA1 = "ISBLANK(F1)";
        sheet.Assert("B14", Is.EqualTo(true));

        // =CONCAT("A", "B", "C") returns "ABC"
        // TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)
        // =TEXTJOIN("-", FALSE, "A", "", "B"): returns "A--B"

        // CLEAN(text): removes non-printable characters (ASCII < 32)

        sheet.Column(1).Width = 50;

        BinDir.Save(workbook, false);
    }

    [Test]
    public void Math()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Math");

        sheet.Cell("A1").Value = Fox;
        sheet.Cell("B1").Value = 15.32.ToString(CultureInfo.CurrentCulture);
        sheet.Cell("C1").Value = 15.32;
        sheet.Cell("D1").Value = 15.62;
        sheet.Cell("E1").Value = 15.66;
        sheet.Cell("F1").Value = 15.65;
        sheet.Cell("G1").Value = -15.65;
        sheet.Range("B2:D2").Value = "Inactive";
        sheet.Range("E2:G2").Value = "Active";

        sheet.Cell("A3").Value = "VALUE(B1)";
        sheet.Cell("B3").FormulaA1 = "VALUE(B1)";
        sheet.Assert("B3", Is.EqualTo(15.32));

        sheet.Cell("A4").Value = "INT(C1)";
        sheet.Cell("B4").FormulaA1 = "INT(C1)";
        sheet.Assert("B4", Is.EqualTo(15));

        sheet.Cell("A5").Value = "ROUNDDOWN(E1, 1) or TRUNC(E1, 1)";
        sheet.Cell("B5").FormulaA1 = "ROUNDDOWN(E1, 1)";
        sheet.Assert("B5", Is.EqualTo(15.6));
        sheet.Cell("C5").FormulaA1 = "TRUNC(E1, 1)";
        sheet.Assert("C5", Is.EqualTo(15.6));

        sheet.Cell("A6").Value = "FLOOR(C1, 2)";
        sheet.Cell("B6").FormulaA1 = "FLOOR(C1, 2)";
        sheet.Cell("C6").Value = "14 is the nearest multiple of 2 for input 15";
        sheet.Assert("B6", Is.EqualTo(14));

        sheet.Cell("A7").Value = "CEILING(C1, 2)";
        sheet.Cell("B7").FormulaA1 = "CEILING(C1, 2)";
        sheet.Cell("C6").Value = "16 is the nearest multiple of 2 for input 15";
        sheet.Assert("B7", Is.EqualTo(16));

        sheet.Cell("A8").Value = "ROUNDUP(C1, 1)";
        sheet.Cell("B8").FormulaA1 = "ROUNDUP(C1, 1)";
        sheet.Assert("B8", Is.EqualTo(15.4));

        sheet.Cell("A9").Value = "ROUND(E1, 1)";
        sheet.Cell("B9").FormulaA1 = "ROUND(E1, 1)";
        sheet.Assert("B9", Is.EqualTo(15.7));

        sheet.Cell("A10").Value = "ISNUMBER(B1), ISEVEN, ISODD";
        sheet.Cell("B10").FormulaA1 = "ISNUMBER(B1)";
        sheet.Assert("B10", Is.EqualTo(false));
        sheet.Cell("C10").FormulaA1 = "ISEVEN(B1)";
        sheet.Assert("C10", Is.EqualTo(false));
        sheet.Cell("D10").FormulaA1 = "ISODD(B1)";
        sheet.Assert("D10", Is.EqualTo(true));

        sheet.Cell("A11").Value = "MAX(B1:G1), MIN(B1:G1)";
        sheet.Cell("B11").FormulaA1 = "MAX(B1:G1)";
        sheet.Assert("B11", Is.EqualTo(15.66));
        sheet.Cell("C11").FormulaA1 = "MIN(B1:G1)";
        sheet.Assert("C11", Is.EqualTo(-15.65));
        // Also MINA, MAXA:
        // - MIN/MAX ignore text and logical values
        // - MINA/MAXA count TRUE as 1 and FALSE as 0, text as 0

        sheet.Cell("A12").Value = "COUNT(B1:H1)";
        sheet.Cell("B12").FormulaA1 = "COUNT(B1:H1)";
        sheet.Assert("B12", Is.EqualTo(5));

        sheet.Cell("A13").Value = "COUNTA(B1:H1)";
        sheet.Cell("B13").FormulaA1 = "COUNTA(B1:H1)";
        sheet.Assert("B13", Is.EqualTo(6));

        sheet.Cell("A14").Value = "COUNTBLANK(B1:H1)";
        sheet.Cell("B14").FormulaA1 = "COUNTBLANK(B1:H1)";
        sheet.Assert("B14", Is.EqualTo(1));

        sheet.Cell("A14").Value = "COUNTIF(B1:H1, \">15\")";
        sheet.Cell("B14").FormulaA1 = "COUNTIF(B1:H1, \">15\")";
        // ATTN: The assertion says 5 but in Excel it shows 4!!
        sheet.Assert("B14", Is.EqualTo(5));

        sheet.Cell("A15").Value = "COUNTIF(A1:H1, \"*f?x*\")";
        sheet.Cell("B15").FormulaA1 = "COUNTIF(A1:H1, \"*f?x*\")";
        sheet.Assert("B15", Is.EqualTo(1));

        sheet.Cell("A16").Value = "COUNTIFS(B1:G1, \">=15\", B2:G2, \"Active\")";
        sheet.Cell("B16").FormulaA1 = "COUNTIFS(B1:G1, \">=15\", B2:G2, \"Active\")";
        sheet.Assert("B16", Is.EqualTo(2));

        sheet.Cell("A17").Value = "SUMIFS(B1:G1, B2:G2, \"Active\", B1:G1, \">15\")";
        sheet.Cell("B17").FormulaA1 = "SUMIFS(B1:G1, B2:G2, \"Active\", B1:G1, \">15\")";
        Assert.That(sheet.Cell("B17").GetDouble(), Is.EqualTo(31.31).Within(0.001));

        sheet.Cell("A18").Value = "AVERAGEIF(C1:G1, \">15\")";
        sheet.Cell("B18").FormulaA1 = "AVERAGEIF(C1:G1, \">15\")";
        // ATTN: AVERAGEIF is not implemented -- but does work when opening in Excel!
        Assert.That(sheet.Cell("B18").Value.IsError, Is.True);
        Assert.That(sheet.Cell("B18").GetError(), Is.EqualTo(XLError.NameNotRecognized));

        sheet.Cell("A19").Value = "AVERAGEIFS(B1:G1, B2:G2, \"Active\", B1:G1, \">15\")";
        sheet.Cell("B19").FormulaA1 = "AVERAGEIFS(B1:G1, B2:G2, \"Active\", B1:G1, \">15\")";
        // ATTN: AVERAGEIFS is not implemented -- but does work when opening in Excel!
        Assert.That(sheet.Cell("B19").Value.IsError, Is.True);
        Assert.That(sheet.Cell("B19").GetError(), Is.EqualTo(XLError.NameNotRecognized));

        // FIXED(number, [decimals], [no_commas]): rounds as text
        // FLOOR.MATH / CEILING.MATH

        // ABS(number) - abolute value
        // SIGN(number) - returns -1 or 1
        // PRODUCT(range...) - returns arg1 * arg2 * ...
        // POWER(base, exponent) - Or base^exp. Also: SQRT
        // MOD(divident, divisor) - modulo. Also: QUOTIENT
        // RAND() - between 0 and 1
        // RANDBETWEEN(lowest, highest) - both params inclusive
        // LARGE(range, xth) - returns xth largest number; also SMALL()

        // PI, SIN, COS, ASIN, ASINH, TAN, ATAN, ...
        // EXP, LOG, LOG10, LN
        // MEDIAN, STDEV/STDEVA, STDEVP, STDEVPA
        // RANK
        // VAR, VARA, VARP, VARPA

        sheet.Column(1).Width = 50;

        BinDir.Save(workbook, false);
    }

    [Test]
    public void Information()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Information");

        // ISNA: TRUE is #N/A
        // ISTEXT and ISNONTEXT
        // NA(): #N/A
        // TYPE(value): 1=number, 2=text, 4=logical, 16=error, 64=array
        // N(value): converts to number: number->number, text->0, TRUE->1, FALSE->0, error->error
        // ISLOGICAL: TRUE if value is TRUE or FALSE
        // ERROR.TYPE(error): 1=#NULL!, 2=#DIV/0!, 3=#VALUE!, 4=#REF!, 5=#NAME?, 6=#NUM!, 7=#N/A, 8=all other errors

        sheet.Column(1).Width = 50;

        BinDir.Save(workbook, false);
    }

    [Test]
    public void Logical()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Logical");

        // Check if cell is blank (empty or whitespace)
        sheet.Cell("A1").Value = "=IF(OR(ISBLANK(C1), TRIM(C1)=\"\"), 1, 0)";
        sheet.Cell("B1").FormulaA1 = "IF(OR(ISBLANK(C1), TRIM(C1)=\"\"), 1, 0)";
        sheet.Cell("C1").Value = " ";
        sheet.Assert("B1", Is.EqualTo(1));

        sheet.Cell("A2").Value = "=IF(OR(ISBLANK(C2), TRIM(C2)=\"\"), 1, 0)";
        sheet.Cell("B2").FormulaA1 = "IF(OR(ISBLANK(C2), TRIM(C2)=\"\"), 1, 0)";
        sheet.Cell("C2").Value = "Hello";
        sheet.Assert("B2", Is.EqualTo(0));

        // Check if cell is either value
        sheet.Cell("A3").Value = "=IF(OR(C3=\"value\", C3=\"value2\"), \"value1-2\", \"other\")";
        sheet.Cell("B3").FormulaA1 = "IF(OR(C3=\"value\", C3=\"value2\"), \"value1-2\", \"other\")";
        sheet.Cell("C3").Value = "value";
        sheet.Assert("B3", Is.EqualTo("value1-2"));

        sheet.Cell("A4").Value = "=IF(OR(C4=\"value\", C4=\"value2\"), \"value1-2\", \"other\")";
        sheet.Cell("B4").FormulaA1 = "IF(OR(C4=\"value\", C4=\"value2\"), \"value1-2\", \"other\")";
        sheet.Cell("C4").Value = "value2";
        sheet.Assert("B4", Is.EqualTo("value1-2"));

        sheet.Cell("A5").Value = "=IF(OR(C5=\"value\", C5=\"value2\"), \"value1-2\", \"other\")";
        sheet.Cell("B5").FormulaA1 = "IF(OR(C5=\"value\", C5=\"value2\"), \"value1-2\", \"other\")";
        sheet.Cell("C5").Value = "something else";
        sheet.Assert("B5", Is.EqualTo("other"));

        // Basic AND/OR/NOT
        sheet.Cell("A6").Value = "=AND(C6=1, D6=2)";
        sheet.Cell("B6").FormulaA1 = "AND(C6=1, D6=2)";
        sheet.Cell("C6").Value = 1;
        sheet.Cell("D6").Value = 2;
        sheet.Assert("B6", Is.EqualTo(true));

        sheet.Cell("A7").Value = "=OR(C7=2, D7=2)";
        sheet.Cell("B7").FormulaA1 = "OR(C7=2, D7=2)";
        sheet.Cell("C7").Value = 1;
        sheet.Cell("D7").Value = 2;
        sheet.Assert("B7", Is.EqualTo(true));

        sheet.Cell("A8").Value = "=NOT(C8=2)";
        sheet.Cell("B8").FormulaA1 = "NOT(C8=2)";
        sheet.Cell("C8").Value = 1;
        sheet.Assert("B8", Is.EqualTo(true));

        sheet.Column(1).Width = 60;
        sheet.Column(2).Width = 20;
        BinDir.Save(workbook, true);
    }

    [Test]
    public void DateAndTime()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("DateAndTime");

        var dt = new DateTime(2024, 6, 15, 14, 30, 45);
        sheet.Cell("A1").Value = dt;

        sheet.Cell("A3").Value = "YEAR(A1)";
        sheet.Cell("B3").FormulaA1 = "YEAR(A1)";
        sheet.Assert("B3", Is.EqualTo(dt.Year));

        sheet.Cell("A4").Value = "MONTH(A1)";
        sheet.Cell("B4").FormulaA1 = "MONTH(A1)";
        sheet.Assert("B4", Is.EqualTo(dt.Month));

        sheet.Cell("A5").Value = "DAY(A1)";
        sheet.Cell("B5").FormulaA1 = "DAY(A1)";
        sheet.Assert("B5", Is.EqualTo(dt.Day));

        sheet.Cell("A6").Value = "HOUR(A1)";
        sheet.Cell("B6").FormulaA1 = "HOUR(A1)";
        sheet.Assert("B6", Is.EqualTo(dt.Hour));

        sheet.Cell("A7").Value = "MINUTE(A1)";
        sheet.Cell("B7").FormulaA1 = "MINUTE(A1)";
        sheet.Assert("B7", Is.EqualTo(dt.Minute));

        sheet.Cell("A8").Value = "SECOND(A1)";
        sheet.Cell("B8").FormulaA1 = "SECOND(A1)";
        sheet.Assert("B8", Is.EqualTo(dt.Second));

        // DATE(year, month, day)
        sheet.Cell("A9").Value = "DATE(2024, 6, 1)";
        sheet.Cell("B9").FormulaA1 = "DATE(2024, 6, 1)";
        var leDate = sheet.Cell("B9").Value;
        int leExpectedDate = (int)System.Math.Ceiling((new DateTime(2024, 6, 1) - new DateTime(1899, 12, 30)).TotalDays);
        Assert.That(leDate, Is.EqualTo(leExpectedDate));

        // TODAY
        sheet.Cell("A10").Value = "TODAY()";
        sheet.Cell("B10").Style.DateFormat.Format = "yyyy-mm-dd";
        sheet.Cell("B10").FormulaA1 = "TODAY()";
        var today = sheet.Cell("B10").Value;
        int excelToday = (int)System.Math.Ceiling((DateTime.Today - new DateTime(1899, 12, 30)).TotalDays);
        Assert.That(today, Is.EqualTo(excelToday));
        Assert.That(sheet.Cell("B10").GetFormattedString(), Is.EqualTo(DateTime.Today.ToString("yyyy-MM-dd")));

        // NOW
        sheet.Cell("A11").Value = "NOW()";
        sheet.Cell("B11").FormulaA1 = "NOW()";
        double now = sheet.Cell("B11").GetDouble();
        double excelNow = (DateTime.Now - new DateTime(1899, 12, 30)).TotalDays;
        Assert.That(now, Is.EqualTo(excelNow).Within(0.001));

        // TIME
        sheet.Cell("D11").Value = "TIME(14, 30, 45)";
        sheet.Column("D").Width = 20;
        sheet.Cell("E11").FormulaA1 = "TIME(14, 30, 45)";
        sheet.Cell("E11").Style.DateFormat.NumberFormatId = (int)XLPredefinedFormat.DateTime.Hour24MinutesSeconds;

        // DATEVALUE
        sheet.Cell("A12").Value = "2024-06-01";
        sheet.Cell("B12").Value = "DATEVALUE(A12)";
        sheet.Cell("C12").FormulaA1 = "DATEVALUE(A12)";
        sheet.Assert("C12", Is.EqualTo(45444));

        // WEEKDAY
        sheet.Cell("A13").Value = "WEEKDAY(A1)";
        sheet.Cell("B13").FormulaA1 = "WEEKDAY(A1)";
        sheet.Assert("B13", Is.EqualTo((int)dt.DayOfWeek + 1)); // Excel: Sunday=1

        // WEEKNUM & ISOWEEKNUM
        sheet.Cell("A14").Value = "WEEKNUM(C12)";
        sheet.Cell("B14").FormulaA1 = "WEEKNUM(C12)";
        sheet.Assert("B14", Is.EqualTo(22));

        sheet.Cell("D14").Value = "ISOWEEKNUM(C12)";
        sheet.Cell("E14").FormulaA1 = "ISOWEEKNUM(C12)";
        sheet.Assert("E14", Is.EqualTo(22));

        // YEARFRAC
        sheet.Cell("A15").Value = "YEARFRAC(DATE(2024,1,1), DATE(2024,6,1))";
        sheet.Cell("B15").FormulaA1 = "YEARFRAC(DATE(2024,1,1), DATE(2024,6,1))";
        Assert.That(sheet.Cell("B15").GetDouble(), Is.EqualTo(0.416666667).Within(0.0001));

        // TODAY() - 2
        sheet.Cell("A16").Value = "TODAY() - 2";
        sheet.Cell("B16").FormulaA1 = "TODAY() - 2";
        int excelToday2 = (int)System.Math.Ceiling((DateTime.Today - new DateTime(1899, 12, 30)).TotalDays) - 2;
        Assert.That(sheet.Cell("B16").Value, Is.EqualTo(excelToday2));

        // NOW() + "2:00"
        sheet.Cell("A17").Value = "NOW() + TIME(2,0,0)";
        sheet.Cell("B17").FormulaA1 = "NOW() + TIME(2,0,0)";
        double nowPlus2h = (DateTime.Now.AddHours(2) - new DateTime(1899, 12, 30)).TotalDays;
        Assert.That(sheet.Cell("B17").GetDouble(), Is.EqualTo(nowPlus2h).Within(0.001));

        // DAYS360(date1, date2)
        sheet.Cell("A18").Value = "DAYS360(DATE(2024,1,1), DATE(2024,6,1))";
        sheet.Cell("B18").FormulaA1 = "DAYS360(DATE(2024,1,1), DATE(2024,6,1))";
        Assert.That(sheet.Cell("B18").GetDouble(), Is.EqualTo(150));

        // EDATE(date, nrOfMonths)
        sheet.Cell("A20").Value = "EDATE(DATE(2024,6,1), 2)";
        sheet.Cell("B20").FormulaA1 = "EDATE(DATE(2024,6,1), 2)";
        var edate = new DateTime(2024, 6, 1).AddMonths(2);
        int excelEDate = (int)System.Math.Ceiling((edate - new DateTime(1899, 12, 30)).TotalDays);
        Assert.That(sheet.Cell("B20").Value, Is.EqualTo(excelEDate));

        // EOMONTH(date, 0)
        sheet.Cell("A21").Value = "EOMONTH(DATE(2024,6,1), 0)";
        sheet.Cell("B21").FormulaA1 = "EOMONTH(DATE(2024,6,1), 0)";
        var eomonth0 = new DateTime(2024, 6, 30);
        int excelEOMonth0 = (int)System.Math.Ceiling((eomonth0 - new DateTime(1899, 12, 30)).TotalDays);
        Assert.That(sheet.Cell("B21").Value, Is.EqualTo(excelEOMonth0));

        // EOMONTH(date, -2)
        sheet.Cell("A22").Value = "EOMONTH(DATE(2024,6,1), -2)";
        sheet.Cell("B22").FormulaA1 = "EOMONTH(DATE(2024,6,1), -2)";
        var eomonthMinus2 = new DateTime(2024, 4, 30);
        int excelEOMonthMinus2 = (int)System.Math.Ceiling((eomonthMinus2 - new DateTime(1899, 12, 30)).TotalDays);
        Assert.That(sheet.Cell("B22").Value, Is.EqualTo(excelEOMonthMinus2));

        // WORKDAY(date, workDaysToAdd, holidaysRange)
        sheet.Cell("A23").Value = "WORKDAY(DATE(2024,6,1), 10, DATE(2024,6,5))";
        sheet.Cell("B23").FormulaA1 = "WORKDAY(DATE(2024,6,1), 10, DATE(2024,6,5))";
        var end = new DateTime(2024, 6, 17); // 1/6 + 10 days + skip 5/6 (holiday) + skip weekends = 17/6
        int excelWorkday = (int)System.Math.Ceiling((end - new DateTime(1899, 12, 30)).TotalDays);
        Assert.That(sheet.Cell("B23").Value, Is.EqualTo(excelWorkday));

        sheet.Column(1).Width = 50;
        sheet.Column(2).Width = 20;
        BinDir.Save(workbook, false);
    }
}