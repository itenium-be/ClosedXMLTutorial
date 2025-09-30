using System.Diagnostics;
using System.Runtime.CompilerServices;
using ClosedXML.Excel;

namespace ClosedXmlTutorial.Util;

public static class BinDir
{
    /// <summary>
    /// Save Excels from the UnitTests under the bin/excels folder of this project
    /// </summary>
    public static string GetPath(string? fileName = null, int frame = 1, [CallerMemberName] string callerName = "")
    {
        var mth = new StackTrace().GetFrame(frame)?.GetMethod();
        string? cls = mth?.ReflectedType?.Name;

        var dir = new DirectoryInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "excels"));
        Directory.CreateDirectory(dir.FullName);

        var name = cls + (fileName ?? callerName);
        if (!name.Contains('.'))
        {
            name += ".xlsx";
        }
        return Path.Combine(dir.FullName, name);
    }

    public static void Save(XLWorkbook workbook, bool openExcel = false, [CallerMemberName] string callerName = "")
    {
        string path = BinDir.GetPath(null, 2, callerName);
        workbook.SaveAs(path);

        if (openExcel)
        {
            var p = new Process();
            p.StartInfo = new ProcessStartInfo(path)
            {
                FileName = path,
                UseShellExecute = true
            };
            p.Start();
        }
    }
}