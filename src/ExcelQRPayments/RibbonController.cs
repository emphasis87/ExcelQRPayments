using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelQRPayments;

[ComVisible(true)]
public class RibbonController : ExcelRibbon
{
    public override string GetCustomUI(string ribbonId)
    {
        using var stream = typeof(RibbonController).Assembly
            .GetManifestResourceStream("ExcelQRPayments.Ribbon.xml");

        if (stream == null)
            return string.Empty;

        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    public string GetVersionLabel(IRibbonControl control)
    {
        var version = typeof(RibbonController).Assembly.GetName().Version;
        return $"v{version?.ToString(3) ?? "0.0.0"}";
    }

    public void OnOpenGitHub(IRibbonControl control)
    {
        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
        {
            FileName = "https://github.com/emphasis87/ExcelQRPayments",
            UseShellExecute = true
        });
    }

    public void OnClearFormulas(IRibbonControl control)
    {
        try
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            var workbook = app.ActiveWorkbook;

            if (workbook == null)
            {
                app.StatusBar = "No active workbook";
                return;
            }

            var clearedCount = 0;

            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                clearedCount += ClearFormulasInSheet(sheet);
            }

            app.StatusBar = clearedCount > 0
                ? $"Cleared {clearedCount} QR Payment formula(s)"
                : "No QR Payment formulas found";
        }
        catch (Exception ex)
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            app.StatusBar = $"Error: {ex.Message}";
        }
    }

    private static int ClearFormulasInSheet(Excel.Worksheet sheet)
    {
        var clearedCount = 0;

        try
        {
            var usedRange = sheet.UsedRange;
            if (usedRange == null) return 0;

            foreach (Excel.Range cell in usedRange)
            {
                try
                {
                    var hasFormula = cell.HasFormula as bool? ?? false;
                    if (!hasFormula) continue;

                    var formula = cell.Formula?.ToString();
                    if (formula?.Contains("InsertPaymentQrCode", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        // Replace formula with its current value
                        cell.Value = cell.Value;
                        clearedCount++;
                    }
                }
                catch
                {
                    // Skip cells that can't be processed
                }
            }
        }
        catch
        {
            // Skip sheets that can't be processed
        }

        return clearedCount;
    }
}
