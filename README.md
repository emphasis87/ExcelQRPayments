# Excel QR Payments

An Excel add-in for generating QR payment codes using the SPAYD (Short Payment Descriptor) format. Commonly used for Czech and Slovak bank payments.

## Features

- Generate SPAYD payment strings from payment details
- Create QR code images and insert them directly into Excel worksheets
- IntelliSense support with parameter descriptions
- Ribbon UI with "Clear Formulas" button to prepare workbooks for sharing

## Installation

1. Download the latest `.xll` file from the [Releases page](https://github.com/emphasis87/ExcelQRPayments/releases/latest):
   - 32-bit Excel: `ExcelQRPayments-AddIn-packed.xll`
   - 64-bit Excel: `ExcelQRPayments-AddIn64-packed.xll`
2. Right-click the downloaded `.xll` file, select **Properties**, check **Unblock**, and click **OK**
3. To install permanently, copy the `.xll` file to a local folder, e.g. `%APPDATA%\Microsoft\AddIns`
4. In Excel, go to **File** > **Options** > **Add-ins**
4. At the bottom, select **Excel Add-ins** and click **Go...**
5. Click **Browse...**, navigate to the `.xll` file, and click **OK**
6. Ensure the add-in is checked in the list and click **OK**

>Alternatively you can quickly try it out by dragging and dropping the `.xll` file into a running Excel window.

## Functions

### GenerateSpayd

Generates a SPAYD string from payment information.

```
=GenerateSpayd(iban, amount, currency, [variableSymbol], [specificSymbol], [constantSymbol], [message], [recipientName], [bic], [dueDate])
```

**Parameters:**
| Parameter | Required | Description |
|-----------|----------|-------------|
| iban | Yes | IBAN account number |
| amount | Yes | Amount to pay |
| currency | Yes | 3-letter ISO currency code (e.g., CZK) |
| variableSymbol | No | Variable symbol (max 10 digits) |
| specificSymbol | No | Specific symbol (max 10 digits) |
| constantSymbol | No | Constant symbol (max 4 digits) |
| message | No | Message for recipient (max 60 chars) |
| recipientName | No | Recipient name |
| bic | No | BIC/SWIFT code |
| dueDate | No | Payment due date |

**Example:**
```
=GenerateSpayd("CZ6508000000192000145399", 1500, "CZK", "1234567890", , , "Invoice 123")
```

**Returns:**
```
SPD*1.0*ACC:CZ6508000000192000145399*AM:1500.00*CC:CZK*MSG:Invoice 123*X-VS:1234567890
```

### InsertPaymentQrCode

Generates a QR code from payment information and inserts it at the specified cell.

```
=InsertPaymentQrCode(targetCell, iban, amount, currency, [variableSymbol], [specificSymbol], [constantSymbol], [message], [recipientName], [bic], [dueDate], [size])
```

**Parameters:**
| Parameter | Required | Description |
|-----------|----------|-------------|
| targetCell | Yes | Cell reference where QR code will be placed (click to select) |
| iban | Yes | IBAN account number |
| amount | Yes | Amount to pay |
| currency | Yes | 3-letter ISO currency code (e.g., CZK) |
| variableSymbol | No | Variable symbol (max 10 digits) |
| specificSymbol | No | Specific symbol (max 10 digits) |
| constantSymbol | No | Constant symbol (max 4 digits) |
| message | No | Message for recipient (max 60 chars) |
| recipientName | No | Recipient name |
| bic | No | BIC/SWIFT code |
| dueDate | No | Payment due date |
| size | No | QR code size in pixels (default: 150) |

**Example:**
```
=InsertPaymentQrCode(E2, "CZ6508000000192000145399", 1500, "CZK", "1234567890", , , "Invoice 123")
```

## Ribbon

The add-in adds a "QR Payments" tab to the Excel ribbon with:

- **Clear Formulas**: Replaces all `InsertPaymentQrCode` formulas with their values, allowing the workbook to be shared with users who don't have the add-in installed. QR code images are preserved.

## Building from source

```bash
dotnet build
```

The `.xll` files will be generated in `src/ExcelQRPayments/bin/Debug/net8.0-windows/publish/`.

## Testing

```bash
dotnet test
```

## Requirements

- .NET 8.0 or later
- Microsoft Excel (32-bit or 64-bit)

## License

MIT
