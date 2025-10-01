ClosedXML Tutorial
==================

Back in 2017 I created a blog series on EPPlus
[Create Excels with C# and EPPlus: A tutorial](https://itenium.be/blog/dotnet/create-xlsx-excel-with-epplus-csharp/)
because it used to be such a great project. But then it went commercial 😀

So time to create a new blog series, now using the MIT licensed ClosedXML!

[ClosedXML/ClosedXML](https://github.com/ClosedXML/ClosedXML) ClosedXML is a .NET library for reading, manipulating and writing Excel 2007+ (.xlsx, .xlsm) files. It aims to provide an intuitive and user-friendly interface to dealing with the underlying OpenXML API. (⭐ 5.2k)

The blog series and this code repository is pretty much a
1 to 1 convertion from [EPPlus](https://github.com/itenium-be/EPPlusTutorial)
to ClosedXML.


## Install

```sh
dotnet add package ClosedXML
```


EPPlus Comparison
-----------------

| EPPlus                           | ClosedXML                     |
|----------------------------------|-------------------------------|
| **Basics**
| new ExcelPackage()               | new XLWorkbook()
| Workbook.Worksheets.Add("str")   | AddWorksheet("str")
|                                  |
| **With Sheet**                   |
| Cells[1, 1]                      | Cell(1, 1)
| Cell["A2"]                       | Cell("A2")
| Cells["A2:C5"]                   | Range("A2:C5")
| Cells["B2,D2"]                   | Cells("B2,D2")
| Dimension                        | LastRowUsed() & LastColumnUsed()
|                                  | Or RangeUsed()
| View.FreezePanes                 | SheetView.Freeze
|                                  | SheetView.FreezeRows & FreezeColumns
| View.ShowGridLines               | ShowGridLines
| View.ShowHeaders                 | ShowRowColHeaders
|                                  |
| **With Cell(s)**
| Address                          | Address.ToString()
| Start.Column                     | Address.ColumnNumber
| Formula                          | FormulaA1
| Text                             | GetString() & GetFormattedString()
| Hyperlink = new Uri()            | SetHyperlink(new XLHyperlink())
| Merge = true                     | Merge() / Unmerge()
|                                  |
| **Styling with Cell.Style**
| Font.Color.SetColor(Color.Ivory)          | Style.Font.SetFontColor(XLColor.Ivory)
|                                           | Style.Font.FontColor = XLColor.Ivory
| Fill.PatternType = ExcelFillStyle.Solid   | Fill.SetPatternType(XLFillPatternValues.Solid)
| Fill.BackgroundColor.SetColor(Color.Navy) | Fill.SetBackgroundColor(XLColor.Navy)
| Border.BorderAround()            | Border.OutsideBorder
| ShrinkToFit = true               | Alignment.ShrinkToFit = true
| WrapText = true                  | Alignment.WrapText = true
| HorizontalAlignment              | Alignment.Horizontal
|                                  |
| **Formulas**
| Formula                          | FormulaA1
| Calculate()                      | RecalculateAllFormulas()
|
| **Misc**                         |
| LoadFromCollection               | InsertTable / InsertData
| Cells.AutoFitColumns()           | ColumnsUsed().AdjustToContents()


Wish List
---------

[ClosedXML.Report](https://github.com/ClosedXML/ClosedXML.Report): Fill a template Excel by using placeholders.

**WebApi**:  

ClosedXML has separate projects for this which provide an extension method.
There is also one for ASP.NET and MVC.

[ClosedXML.Extensions.WebApi](https://github.com/ClosedXML/ClosedXML.Extensions.WebApi)


**Adding a picture**:  

```c#
using var wb = new XLWorkbook();
var ws = wb.AddWorksheet("Sheet1");
var imagePath = @"c:\path\to\your\image.jpg";

var image = ws.AddPicture(imagePath)
    .MoveTo(ws.Cell("B3"))
    .Scale(0.5); // optional: resize picture
      
wb.SaveAs("file.xlsx");
```
