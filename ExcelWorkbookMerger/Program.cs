using System;
using System.Windows.Forms;
using OfficeOpenXml;

namespace ExcelWorkbookMerger;

internal static class Program
{
    /// <summary>
    ///     The main entry point for the application.
    /// </summary>
    [STAThread]
    private static void Main()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        Application.EnableVisualStyles();
        Application.SetHighDpiMode(HighDpiMode.SystemAware);
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new MainWindow());
    }
}