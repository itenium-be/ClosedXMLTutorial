using ClosedXML.Excel;

namespace ClosedXmlTutorial.Util;

public static class ExcelExtensions
{
    public static IXLDataValidation SetDropdownList(this IXLCell cell, IEnumerable<string> options)
    {
        var validation = cell.CreateDataValidation();
        validation.AllowedValues = XLAllowedValues.List;
        validation.InCellDropdown = true;
        validation.IgnoreBlanks = true;
        validation.List($"\"{string.Join(",", options)}\"");
        return validation;
    }

    public static IXLCell SetHyperlink(this IXLCell cell, string url, string text, string? tooltip = null)
    {
        cell.SetValue(text).SetHyperlink(new XLHyperlink(url, tooltip));
        return cell;
    }

    public static IXLStyle SetNumberFormatId(this IXLNumberFormat cellStyleNumberFormat, XLPredefinedFormat.Number format)
    {
        return cellStyleNumberFormat.SetNumberFormatId((int)format);
    }

    public static IXLStyle SetNumberFormatId(this IXLNumberFormat cellStyleNumberFormat, XLPredefinedFormat.DateTime format)
    {
        return cellStyleNumberFormat.SetNumberFormatId((int)format);
    }
}