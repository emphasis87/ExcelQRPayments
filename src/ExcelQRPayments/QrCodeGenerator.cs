using System.Drawing;
using System.Drawing.Imaging;
using QRCoder;

namespace ExcelQRPayments;

public static class QrCodeGenerator
{
    public static byte[] GeneratePng(string content, int pixelsPerModule = 10)
    {
        using var qrGenerator = new QRCodeGenerator();
        using var qrCodeData = qrGenerator.CreateQrCode(content, QRCodeGenerator.ECCLevel.M);
        using var qrCode = new PngByteQRCode(qrCodeData);

        return qrCode.GetGraphic(pixelsPerModule);
    }

    public static Bitmap GenerateBitmap(string content, int pixelsPerModule = 10)
    {
        using var qrGenerator = new QRCodeGenerator();
        using var qrCodeData = qrGenerator.CreateQrCode(content, QRCodeGenerator.ECCLevel.M);
        using var qrCode = new QRCode(qrCodeData);

        return qrCode.GetGraphic(pixelsPerModule);
    }

    public static string SaveToTempFile(string content, int pixelsPerModule = 10)
    {
        var pngBytes = GeneratePng(content, pixelsPerModule);
        var tempPath = Path.Combine(Path.GetTempPath(), $"qr_{Guid.NewGuid():N}.png");
        File.WriteAllBytes(tempPath, pngBytes);
        return tempPath;
    }
}
