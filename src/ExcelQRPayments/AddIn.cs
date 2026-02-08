using ExcelDna.Integration;
using ExcelDna.IntelliSense;

namespace ExcelQRPayments;

public class AddIn : IExcelAddIn
{
    public void AutoOpen()
    {
        IntelliSenseServer.Install();
    }

    public void AutoClose()
    {
        IntelliSenseServer.Uninstall();
    }
}
