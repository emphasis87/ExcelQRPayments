namespace ExcelQRPayments.Tests;

public class SpaydBuilderTests
{
    [Fact]
    public Task Build_WithOnlyIban_ReturnsMinimalSpayd()
    {
        var spayd = new SpaydBuilder()
            .WithIban("CZ6508000000192000145399")
            .Build();

        return Verify(spayd);
    }

    [Fact]
    public Task Build_WithIbanAndBic_FormatsAccountCorrectly()
    {
        var spayd = new SpaydBuilder()
            .WithIban("CZ6508000000192000145399")
            .WithBic("GIBACZPX")
            .Build();

        return Verify(spayd);
    }

    [Fact]
    public Task Build_WithAmount_FormatsWithTwoDecimals()
    {
        var spayd = new SpaydBuilder()
            .WithIban("CZ6508000000192000145399")
            .WithAmount(1234.5m)
            .Build();

        return Verify(spayd);
    }

    [Fact]
    public Task Build_WithDueDate_FormatsAsYyyymmdd()
    {
        var spayd = new SpaydBuilder()
            .WithIban("CZ6508000000192000145399")
            .WithDueDate(new DateOnly(2024, 12, 31))
            .Build();

        return Verify(spayd);
    }

    [Fact]
    public Task Build_WithSpecialCharacters_EncodesAsteriskAndPercent()
    {
        var spayd = new SpaydBuilder()
            .WithIban("CZ6508000000192000145399")
            .WithMessage("Test*Message%Here")
            .Build();

        return Verify(spayd);
    }

    [Fact]
    public Task Build_WithAllFields_ContainsAllFields()
    {
        var spayd = new SpaydBuilder()
            .WithIban("CZ6508000000192000145399")
            .WithBic("GIBACZPX")
            .WithAmount(999.99m)
            .WithCurrency("CZK")
            .WithVariableSymbol("1234567890")
            .WithSpecificSymbol("0987654321")
            .WithConstantSymbol("0308")
            .WithMessage("Invoice 123")
            .WithRecipientName("Test s.r.o.")
            .WithDueDate(new DateOnly(2024, 6, 15))
            .Build();

        return Verify(spayd);
    }

    [Fact]
    public Task WithIban_NormalizesSpacesAndCase()
    {
        var spayd = new SpaydBuilder()
            .WithIban("cz65 0800 0000 1920 0014 5399")
            .Build();

        return Verify(spayd);
    }

    [Fact]
    public void Build_WithoutIban_ThrowsInvalidOperationException()
    {
        var builder = new SpaydBuilder()
            .WithAmount(100m);

        Assert.Throws<InvalidOperationException>(() => builder.Build());
    }

    [Fact]
    public void WithAmount_NegativeAmount_ThrowsArgumentOutOfRangeException()
    {
        var builder = new SpaydBuilder();

        Assert.Throws<ArgumentOutOfRangeException>(() => builder.WithAmount(-100m));
    }

    [Fact]
    public void WithCurrency_InvalidLength_ThrowsArgumentException()
    {
        var builder = new SpaydBuilder();

        Assert.Throws<ArgumentException>(() => builder.WithCurrency("CZKK"));
    }

    [Fact]
    public void WithMessage_TooLong_ThrowsArgumentException()
    {
        var builder = new SpaydBuilder();
        var longMessage = new string('x', 61);

        Assert.Throws<ArgumentException>(() => builder.WithMessage(longMessage));
    }

    [Fact]
    public void WithVariableSymbol_InvalidFormat_ThrowsArgumentException()
    {
        var builder = new SpaydBuilder();

        Assert.Throws<ArgumentException>(() => builder.WithVariableSymbol("ABC123"));
    }
}
