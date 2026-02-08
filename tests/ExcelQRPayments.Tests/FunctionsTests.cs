using VerifyXunit;

namespace ExcelQRPayments.Tests;

public class FunctionsTests
{
    [Fact]
    public Task GenerateSpayd_WithRequiredFields_ReturnsSpaydString()
    {
        var result = Functions.GenerateSpayd(
            iban: "CZ6508000000192000145399",
            amount: 450.00,
            currency: "CZK",
            variableSymbol: Type.Missing,
            specificSymbol: Type.Missing,
            constantSymbol: Type.Missing,
            message: Type.Missing,
            recipientName: Type.Missing,
            bic: Type.Missing,
            dueDate: Type.Missing);

        return Verifier.Verify(result);
    }

    [Fact]
    public Task GenerateSpayd_WithAllFields_ReturnsSpaydString()
    {
        var result = Functions.GenerateSpayd(
            iban: "CZ6508000000192000145399",
            amount: 1234.56,
            currency: "CZK",
            variableSymbol: "1234567890",
            specificSymbol: "9876543210",
            constantSymbol: "0308",
            message: "Payment for invoice",
            recipientName: "Test Company",
            bic: "GIBACZPX",
            dueDate: Type.Missing);

        return Verifier.Verify(result);
    }

    [Fact]
    public Task GenerateSpayd_WithNumericVariableSymbol_HandlesAsString()
    {
        var result = Functions.GenerateSpayd(
            iban: "CZ6508000000192000145399",
            amount: 100.00,
            currency: "CZK",
            variableSymbol: 12345.0,
            specificSymbol: Type.Missing,
            constantSymbol: Type.Missing,
            message: Type.Missing,
            recipientName: Type.Missing,
            bic: Type.Missing,
            dueDate: Type.Missing);

        return Verifier.Verify(result);
    }
}
