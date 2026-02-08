using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelQRPayments;

public partial class SpaydBuilder
{
    private const string SpaydVersion = "1.0";

    private string? _iban;
    private string? _bic;
    private decimal? _amount;
    private string? _currency;
    private string? _variableSymbol;
    private string? _specificSymbol;
    private string? _constantSymbol;
    private string? _recipientName;
    private string? _message;
    private DateOnly? _dueDate;

    public SpaydBuilder WithIban(string iban)
    {
        _iban = NormalizeIban(iban);
        return this;
    }

    public SpaydBuilder WithBic(string bic)
    {
        _bic = bic?.Trim().ToUpperInvariant();
        return this;
    }

    public SpaydBuilder WithAmount(decimal amount)
    {
        if (amount <= 0)
            throw new ArgumentOutOfRangeException(nameof(amount), "Amount must be positive.");

        _amount = amount;
        return this;
    }

    public SpaydBuilder WithCurrency(string currency)
    {
        if (currency?.Length != 3)
            throw new ArgumentException("Currency must be a 3-letter ISO code.", nameof(currency));

        _currency = currency.ToUpperInvariant();
        return this;
    }

    public SpaydBuilder WithVariableSymbol(string variableSymbol)
    {
        if (!string.IsNullOrWhiteSpace(variableSymbol) && !NumericSymbolRegex().IsMatch(variableSymbol))
            throw new ArgumentException("Variable symbol must contain only digits (max 10).", nameof(variableSymbol));

        _variableSymbol = variableSymbol?.Trim();
        return this;
    }

    public SpaydBuilder WithSpecificSymbol(string specificSymbol)
    {
        if (!string.IsNullOrWhiteSpace(specificSymbol) && !NumericSymbolRegex().IsMatch(specificSymbol))
            throw new ArgumentException("Specific symbol must contain only digits (max 10).", nameof(specificSymbol));

        _specificSymbol = specificSymbol?.Trim();
        return this;
    }

    public SpaydBuilder WithConstantSymbol(string constantSymbol)
    {
        if (!string.IsNullOrWhiteSpace(constantSymbol) && !ConstantSymbolRegex().IsMatch(constantSymbol))
            throw new ArgumentException("Constant symbol must contain only digits (max 4).", nameof(constantSymbol));

        _constantSymbol = constantSymbol?.Trim();
        return this;
    }

    public SpaydBuilder WithRecipientName(string recipientName)
    {
        _recipientName = recipientName?.Trim();
        return this;
    }

    public SpaydBuilder WithMessage(string message)
    {
        if (message?.Length > 60)
            throw new ArgumentException("Message must not exceed 60 characters.", nameof(message));

        _message = message?.Trim();
        return this;
    }

    public SpaydBuilder WithDueDate(DateOnly dueDate)
    {
        _dueDate = dueDate;
        return this;
    }

    public string Build()
    {
        if (string.IsNullOrWhiteSpace(_iban))
            throw new InvalidOperationException("IBAN is required.");

        var sb = new StringBuilder();
        sb.Append("SPD*");
        sb.Append(SpaydVersion);

        AppendField(sb, "ACC", FormatAccount());
        AppendField(sb, "AM", _amount?.ToString("F2", CultureInfo.InvariantCulture));
        AppendField(sb, "CC", _currency);
        AppendField(sb, "DT", _dueDate?.ToString("yyyyMMdd", CultureInfo.InvariantCulture));
        AppendField(sb, "MSG", _message);
        AppendField(sb, "RN", _recipientName);
        AppendField(sb, "X-VS", _variableSymbol);
        AppendField(sb, "X-SS", _specificSymbol);
        AppendField(sb, "X-KS", _constantSymbol);

        return sb.ToString();
    }

    private string FormatAccount()
    {
        if (string.IsNullOrWhiteSpace(_bic))
            return _iban!;

        return $"{_iban}+{_bic}";
    }

    private static void AppendField(StringBuilder sb, string key, string? value)
    {
        if (!string.IsNullOrWhiteSpace(value))
        {
            sb.Append('*');
            sb.Append(key);
            sb.Append(':');
            sb.Append(EncodeValue(value));
        }
    }

    private static string EncodeValue(string value)
    {
        var sb = new StringBuilder();
        foreach (var c in value)
        {
            if (c == '*' || c == '%')
            {
                sb.Append('%');
                sb.Append(((int)c).ToString("X2"));
            }
            else
            {
                sb.Append(c);
            }
        }
        return sb.ToString();
    }

    private static string NormalizeIban(string iban)
    {
        if (string.IsNullOrWhiteSpace(iban))
            throw new ArgumentException("IBAN cannot be empty.", nameof(iban));

        return iban.Replace(" ", "").ToUpperInvariant();
    }

    [GeneratedRegex(@"^\d{1,10}$")]
    private static partial Regex NumericSymbolRegex();

    [GeneratedRegex(@"^\d{1,4}$")]
    private static partial Regex ConstantSymbolRegex();
}
