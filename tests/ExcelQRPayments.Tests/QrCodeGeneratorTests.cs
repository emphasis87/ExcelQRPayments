namespace ExcelQRPayments.Tests;

public class QrCodeGeneratorTests
{
    [Fact]
    public void GeneratePng_WithValidContent_ReturnsPngBytes()
    {
        var content = "SPD*1.0*ACC:CZ6508000000192000145399";

        var result = QrCodeGenerator.GeneratePng(content);

        Assert.NotNull(result);
        Assert.True(result.Length > 0);
        Assert.Equal(0x89, result[0]); // PNG magic number
        Assert.Equal(0x50, result[1]); // 'P'
        Assert.Equal(0x4E, result[2]); // 'N'
        Assert.Equal(0x47, result[3]); // 'G'
    }

    [Fact]
    public void SaveToTempFile_WithValidContent_CreatesFile()
    {
        var content = "SPD*1.0*ACC:CZ6508000000192000145399";

        var filePath = QrCodeGenerator.SaveToTempFile(content);

        try
        {
            Assert.True(File.Exists(filePath));
            Assert.EndsWith(".png", filePath);

            var bytes = File.ReadAllBytes(filePath);
            Assert.True(bytes.Length > 0);
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void GeneratePng_WithDifferentSizes_ProducesDifferentOutputs()
    {
        var content = "SPD*1.0*ACC:CZ6508000000192000145399";

        var small = QrCodeGenerator.GeneratePng(content, pixelsPerModule: 5);
        var large = QrCodeGenerator.GeneratePng(content, pixelsPerModule: 20);

        Assert.True(large.Length > small.Length);
    }
}
