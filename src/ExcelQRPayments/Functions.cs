using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelQRPayments;

public static class Functions
{
    [ExcelFunction(Description = "Generates a SPAYD string from payment information")]
    public static string GenerateSpayd(
        [ExcelArgument(Description = "IBAN account number")] string iban,
        [ExcelArgument(Description = "Amount to pay")] double amount,
        [ExcelArgument(Description = "Currency (3-letter ISO code, e.g., CZK)")] string currency,
        [ExcelArgument(Description = "Variable symbol (optional)")] object variableSymbol,
        [ExcelArgument(Description = "Specific symbol (optional)")] object specificSymbol,
        [ExcelArgument(Description = "Constant symbol (optional)")] object constantSymbol,
        [ExcelArgument(Description = "Message for recipient (optional, max 60 chars)")] object message,
        [ExcelArgument(Description = "Recipient name (optional)")] object recipientName,
        [ExcelArgument(Description = "BIC/SWIFT code (optional)")] object bic,
        [ExcelArgument(Description = "Due date (optional)")] object dueDate)
    {
        try
        {
            var builder = new SpaydBuilder()
                .WithIban(iban)
                .WithAmount((decimal)amount)
                .WithCurrency(currency);

            if (TryGetString(variableSymbol, out var vs))
                builder.WithVariableSymbol(vs);

            if (TryGetString(specificSymbol, out var ss))
                builder.WithSpecificSymbol(ss);

            if (TryGetString(constantSymbol, out var ks))
                builder.WithConstantSymbol(ks);

            if (TryGetString(message, out var msg))
                builder.WithMessage(msg);

            if (TryGetString(recipientName, out var rn))
                builder.WithRecipientName(rn);

            if (TryGetString(bic, out var bicCode))
                builder.WithBic(bicCode);

            if (TryGetDate(dueDate, out var date))
                builder.WithDueDate(date);

            return builder.Build();
        }
        catch (Exception ex)
        {
            return $"#ERROR: {ex.Message}";
        }
    }

    [ExcelFunction(
        Description = "Generates a payment QR code from SPAYD data and inserts it at the specified cell",
        IsMacroType = true)]
    public static string InsertPaymentQrCode(
        [ExcelArgument(Description = "Target cell reference where the QR code will be placed", AllowReference = true)] object targetCell,
        [ExcelArgument(Description = "IBAN account number")] string iban,
        [ExcelArgument(Description = "Amount to pay")] double amount,
        [ExcelArgument(Description = "Currency (3-letter ISO code, e.g., CZK)")] string currency,
        [ExcelArgument(Description = "Variable symbol (optional)")] object variableSymbol,
        [ExcelArgument(Description = "Specific symbol (optional)")] object specificSymbol,
        [ExcelArgument(Description = "Constant symbol (optional)")] object constantSymbol,
        [ExcelArgument(Description = "Message for recipient (optional, max 60 chars)")] object message,
        [ExcelArgument(Description = "Recipient name (optional)")] object recipientName,
        [ExcelArgument(Description = "BIC/SWIFT code (optional)")] object bic,
        [ExcelArgument(Description = "Due date (optional)")] object dueDate,
        [ExcelArgument(Description = "QR code size in pixels (optional, default 150)")] object size)
    {
        try
        {
            if (!TryGetCellReference(targetCell, out var cellReference))
                return "#ERROR: Invalid cell reference";

            var spayd = GenerateSpayd(iban, amount, currency, variableSymbol, specificSymbol,
                constantSymbol, message, recipientName, bic, dueDate);

            if (spayd.StartsWith("#ERROR"))
                return spayd;

            var qrSize = TryGetDouble(size, out var s) ? (int)s : 150;
            var pixelsPerModule = Math.Max(1, qrSize / 25);

            var tempFile = QrCodeGenerator.SaveToTempFile(spayd, pixelsPerModule);

            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                InsertImageAtReference(cellReference, tempFile, qrSize);
            });

            return spayd;
        }
        catch (Exception ex)
        {
            return $"#ERROR: {ex.Message}";
        }
    }

    private static bool TryGetCellReference(object value, out ExcelReference result)
    {
        if (value is ExcelReference excelRef)
        {
            result = excelRef;
            return true;
        }

        result = null!;
        return false;
    }

    private static void InsertImageAtReference(ExcelReference cellReference, string imagePath, int size)
    {
        try
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;

            var sheetName = (string)XlCall.Excel(XlCall.xlSheetNm, cellReference);
            // Remove the [Book]Sheet format to just Sheet
            if (sheetName.Contains(']'))
                sheetName = sheetName[(sheetName.IndexOf(']') + 1)..];

            var row = cellReference.RowFirst + 1;
            var col = cellReference.ColumnFirst + 1;

            var sheet = (Excel.Worksheet)app.Sheets[sheetName];
            var targetRange = (Excel.Range)sheet.Cells[row, col];

            var existingShapeName = $"PaymentQR_{sheetName}_{row}_{col}";
            TryDeleteShape(sheet, existingShapeName);

            var left = (float)(double)targetRange.Left;
            var top = (float)(double)targetRange.Top;

            var shape = sheet.Shapes.AddPicture(
                Filename: imagePath,
                LinkToFile: Microsoft.Office.Core.MsoTriState.msoFalse,
                SaveWithDocument: Microsoft.Office.Core.MsoTriState.msoTrue,
                Left: left,
                Top: top,
                Width: size,
                Height: size);

            shape.Name = existingShapeName;
            shape.Placement = Excel.XlPlacement.xlMoveAndSize;

            try { File.Delete(imagePath); } catch { }
        }
        catch (Exception ex)
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            app.StatusBar = $"QR Code Error: {ex.Message}";
        }
    }

    private static void TryDeleteShape(Excel.Worksheet sheet, string shapeName)
    {
        try
        {
            sheet.Shapes.Item(shapeName).Delete();
        }
        catch { }
    }

    private static bool TryGetString(object value, out string result)
    {
        if (value is string s && !string.IsNullOrWhiteSpace(s))
        {
            result = s;
            return true;
        }
        if (value is double d)
        {
            result = ((long)d).ToString();
            return true;
        }
        result = string.Empty;
        return false;
    }

    private static bool TryGetDouble(object value, out double result)
    {
        if (value is double d)
        {
            result = d;
            return true;
        }
        result = 0;
        return false;
    }

    private static bool TryGetDate(object value, out DateOnly result)
    {
        if (value is double d)
        {
            var dateTime = DateTime.FromOADate(d);
            result = DateOnly.FromDateTime(dateTime);
            return true;
        }
        result = default;
        return false;
    }
}
