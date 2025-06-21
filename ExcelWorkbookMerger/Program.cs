using System;
using System.Runtime.Versioning;
using System.Windows.Forms;
using OfficeOpenXml;

namespace ExcelWorkbookMerger;

internal static class Program
{
    /// <summary>
    ///     The main entry point for the application.
    /// </summary>
    [STAThread]
    [SupportedOSPlatform("windows")]
    private static void Main()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        Application.EnableVisualStyles();
        Application.SetHighDpiMode(HighDpiMode.SystemAware);
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new MainWindow());
    }
}