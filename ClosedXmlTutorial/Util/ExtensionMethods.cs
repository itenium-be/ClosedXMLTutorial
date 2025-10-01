using ClosedXML.Excel;
using NUnit.Framework.Constraints;

namespace ClosedXmlTutorial.Util;

internal static class ExtensionMethods
{
    public static void Assert(this IXLWorksheet sheet, string cellAddress, Constraint constraint)
    {
        var value = sheet.Cell(cellAddress).Value;
        NUnit.Framework.Assert.That(value, constraint, $"For cell {cellAddress}");
    }
}