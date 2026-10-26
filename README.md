# fotoleu Toolbox
A toolbox for creating customer documents and invoices at fotoleu from Excel data
and Word templates. It includes a Swiss QR-bill generator and can combine the
invoice details and QR payment section into one Word document.

## Version history

### V2.x and earlier: fixed worksheet cells

QR inputs were read from fixed cells on a worksheet named `SwissQRCode`.
This tied the add-in to a particular workbook layout.

### V3.x: table-based workbook integration

QR inputs moved to the named Excel table `TabQRCode`, making them independent of
fixed cell locations. Document placeholders and configuration use `TabABBookmarks`.
The V3 series also added SCOR creditor references. These table-based interfaces
remain the basis of V4 and V5.

### V4.x: document workflow and diagnostic improvements

V4 refined the existing Word-document and combined-invoice workflow:

- Added explicit control over whether generated documents are shown in Word.
  The individual document buttons show their result, while intermediate documents
  for a combined invoice are generated without opening visible Word windows.
- Used the configured `ABFilePath` and `QRFilePath` when no explicit output path
  is supplied to the respective generation method.
- Expanded debug output with ribbon-action markers, method names, toolbox and
  template versions, file paths and the document-generation stage.
- Improved diagnostics and exception handling around replacing the QR image in
  Word, and fixed the read-only document-opening issue that interfered with editing.

V4 still generated combined addresses (type **K**) for the creditor and debtor.
It did not introduce the structured-address migration described below.

### V5.x: structured Swiss QR-bills and validated UTF-8 generation

V5 updates QR generation for structured Swiss QR-bill addresses while preserving
the existing Word templates, placeholders, bitmap handoff and invoice assembly:

- **Structured addresses:** Both creditor and debtor use type **S**, with separate
  street, house number, postal code, city and country fields read from `TabQRCode`.
  Street and house number are optional; name, postal code, city and country are
  required. A blank house number is supported and remains an empty QR field.
- **Correct encoding:** QR payloads use UTF-8 without a BOM and error correction
  level **M**, including correct encoding of umlauts. The payload header remains
  `SPC / 0200 / 1`; the add-in version does not change the QR payload version.
- **Validated inputs:** Missing or duplicate keys, Excel formula errors, unreviewed
  migration markers, invalid characters, invalid IBAN/SCOR checksums and invalid
  amounts are rejected before a QR code is generated. Numeric Excel amounts are
  handled independently of the Windows decimal separator.
- **Existing debug-sheet feedback:** `SwissQRCode-Debug` now reports structured-input
  validation and encoding details. Feedback is routed to the active workbook.
  An empty optional `BitmapPath` is explicitly skipped rather than logged as a
  failed copy followed by a success.
- **Regression coverage:** Automated checks cover address fields, references,
  amounts, actual decoded UTF-8 bytes and a QR bitmap with the Swiss cross.
  The supplied V5 Word/PDF test invoices were also decoded and compared with the
  debug output. Banking-app acceptance was additionally reported by the operator;
  this is not a SIX certification or proof of an executed payment.

**V5 requires a workbook migration.** Extend the existing productive workbook
in place using the [structured-address instructions](#structured-addresses-for-swiss-qr-bills).
The old combined-address keys are not used as a fallback. The bundled example
workbook must also be migrated before use with V5.

The toolbox version reported in the debug log is separate from the workbook's
`Version` and template-version values. Existing configuration values such as
`4.0` or `2023-01` do not mean that the V4 add-in is still running.

## Overview
The main logic is in fotoleuToolbox.cs, which contains three main functions:
1. generateAuftragsblatt: Generates a document, loads a word template and replaces placeholders with real values from excel table named "TabABBookmarks"
2. generateQRCodeV2: Generates QR bitmap code, replaces them in the template and stores the newly generated file under the name passed by strFilePath.
3. generateRechnung: Main function, combines the two methods generateAuftragsblatt and generateQRCodeV2 and creates combined final document.

## How to setup in your environment
1. Install the published toolbox to add the "fotoleu Toolbox" ribbon to Excel.
2. Keep the existing `TabABBookmarks` and `TabQRCode` tables in your workbook.
   For a new integration, copy them from `fotoleuToolbox.xlsx`.
3. Link the values in the second column of each table to your workbook data and
   configure the Word-template and output paths.
4. For V5, migrate `TabQRCode` to the structured-address schema below, including
   when starting from the bundled example workbook.
5. To enable debug feedback, keep or copy the `SwissQRCode-Debug` worksheet and
   set cell **D2 to `ON`**.

## Structured addresses for Swiss QR-bills

The updated add-in requires address type **S** for both the creditor and debtor.
Update your existing productive workbook in place; do not replace it with an empty
workbook. The example `fotoleuToolbox.xlsx` still contains the legacy combined-address
fields and must also be migrated before use with this version.

Keep the existing `TabQRCode` table and add the following keys in its `Name` column:

| Creditor (business configuration) | Debtor (selected customer) |
|---|---|
| ContactName | DebitorName |
| ContactStreet | DebitorStreet |
| ContactHouseNumber | DebitorHouseNumber |
| ContactPostalCode | DebitorPostalCode |
| ContactCity | DebitorCity |
| ContactCountry | DebitorCountry |

The existing name and country keys should be retained, not duplicated. Street and
house number may be empty but their keys must exist. Name, postal code, city and
country are required. Country must be an ISO two-letter code, for example `CH`.
Keep postal codes and house numbers as text to preserve leading zeros and suffixes.
The add-in rejects missing keys, duplicate keys, Excel formula errors and unreviewed
values such as `PRUEFEN`, `n/a` and `n/v`. Review any address-extraction proposals
before using them for invoices.

The old `ContactAdressLine1/2` and `DebitorAdressLine1/2` keys may remain for
compatibility, but are not used to construct the new QR payload. There is no fallback
to combined addresses. Keep the existing `Version` key; no particular new version
value is required. The structured keys themselves are validated.

Leave `TabABBookmarks`, Word placeholders, templates and the downstream document
workflow unchanged. Make sure the printed addresses agree with the structured
configuration. Both the QR-only button and combined-invoice button use the same
updated QR generation.

Other input requirements:

- Use a normal CH/LI IBAN, not a QR-IBAN. A nonempty `SCOR` must have a valid ISO
  11649 RF reference and checksum; an empty reference uses `NON`. QRR is not
  exposed by this add-in.
- `Currency` must be `CHF` or `EUR`.
- Link `Amount` to an Excel numeric cell. The value must be 0.01 to 999999999.99,
  with no more than two decimal places. Text amounts use the current Windows
  culture's decimal separator and must not contain grouping separators or currency
  symbols.
- The payload is generated as UTF-8 without a BOM, at error correction level M.
  The payload header remains `0200`; it is not the guideline publication version.
  Current SIX Latin character ranges, including extended Latin letters and the euro
  sign, are supported. Embedded line breaks and unsupported characters are rejected.

### Feedback and deployment

Keep the `SwissQRCode-Debug` worksheet and set **D2 to `ON`** to receive generation
stages, validation errors, payment details and the existing QR-image preview.
Errors use that sheet while debugging is enabled; with debugging disabled or the
sheet absent, errors are shown in a message box. Debug output belongs to the active
workbook, even when multiple workbooks are open.

Build/publish the VSTO add-in using Visual Studio with **Office/SharePoint development
tools**, the .NET Framework 4.7.2 targeting pack and your normal signing configuration.
The signing keys are not included in this repository. Before each production
release, verify generation in Excel against the migrated workbook on the supported
Office setup.

Assembly signing and publishing use separate keys: keep the existing
`AssemblyOriginatorKeyFile` to preserve the assembly identity. ClickOnce manifests
and `setup.exe` use `ManifestCertificateThumbprint`, selected from the current
user's Personal certificate store or imported from `ManifestKeyFile`. The
publishing certificate must be valid and have a private key suitable for code
signing; a successful build alone does not prove that `setup.exe` can be signed.

The private-use publishing certificate selected in this project expires on
26 October 2029. Back it up using `certmgr.msc` > Personal > Certificates >
**fotoleuToolbox publishing 2026** > All Tasks > Export, including its private key
in a password-protected PFX. Keep that backup and password outside the repository.
On another build PC, import the PFX into the current user's Personal store.
The public-only certificate is available locally at
`bin\Publishing\fotoleuToolbox-Publisher.cer`; it cannot sign releases. For this
self-signed certificate, target PCs need to trust that public certificate in
Trusted Root Certification Authorities and Trusted Publishers, subject to local
security policy. Certificate replacement can require renewed trust and
reinstallation of an existing add-in; verify the update on a non-production PC.
Do not change the assembly-signing key when renewing the publishing certificate.

### QR regression checks

The Office-independent tests require the .NET 10 SDK on Windows. ZXing.Net is a
test-only decoder dependency; it is not added to the VSTO add-in.

```powershell
dotnet run --project .\tests\SwissQrCode.Tests.csproj
```

To additionally validate the cached `TabQRCode` inputs from a local workbook
without running macros or changing the workbook:

```powershell
dotnet run --project .\tests\SwissQrCode.Tests.csproj -- --workbook "C:\path\fotoleu 2026.xlsm"
```

# SwissQRCodeExcel
This project was built on existing project SwissQRCodeExcel, see https://github.com/barnstee/SwissQRCodeExcel

QR code generator for Microsoft Excel to be used in the Swiss banking sector. The requirements for the new Swiss QR bill can be read [here](https://www.moneytoday.ch/lexikon/qr-rechnung/).

The QR code generator is an Excel ribbon button that, when clicked, will read the data from the supplied SwissQRBill.xlsx Excel template, generate the QR code and then place it in the correct position in the SwissQRBill.xlsx.
