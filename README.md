# fotoleu Toolbox
A toolbox to create documents used at fotoleu based on Excel and Word. It creates based on word templates and data retrieved from Excel final word documents to be send to customer, like bill. It includes a Swiss QR code creator for bills used within Switzerland.

Up to V2.x the values for the QR code have been read from an excel sheet with name "SwissQRCode" from fixed cells. This was not very convienend.

Starting from V3.x the values for the QR code are read from a table named "TabQRCode". This is more flexible and allows better integration into other excel workbooks.

## Overview
The main logic is in fotoleuToolbox.cs, which contains three main functions:
1. generateAuftragsblatt: Generates a document, loads a word template and replaces placeholders with real values from excel table named "TabABBookmarks"
2. generateQRCodeV2: Generates QR bitmap code, replaces them in the template and stores the newly generated file under the name passed by strFilePath.
3. generateRechnung: Main function, combines the two methods generateAuftragsblatt and generateQRCodeV2 and creates combined final document.

## How to setup in your environment
You can easily setup this toolbox in your environment, be following the steps below.
1. Install the published toolbox; this install the ribbon button "fotoleu Toolbox" in your excel
2. Copy the two tables "TabABBookmarks" and "TabQRCode" form example excel workbook: fotoleuToolbox.xlsx
3. Link/adjust the values in second column of the tables above, with your values (e.g. link them with values out of your excel workbook)
4. In case you like to trace debug messages within excel, copy also the sheet "SwissQRCode-Debug" to you excel; this allows you to enable/disable debug messages
5. Done :-)

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
The signing keys are not included in this repository. A production release still
needs an Excel run against the migrated workbook on the supported Office setup.

Assembly signing and publishing use separate keys: keep the existing
`AssemblyOriginatorKeyFile` to preserve the assembly identity. ClickOnce manifests
and `setup.exe` use `ManifestCertificateThumbprint`, selected from the current
user's Personal certificate store or imported from `ManifestKeyFile`. The
publishing certificate must be valid and
have a private key suitable for code signing; a successful build alone does not
prove that `setup.exe` can be signed.

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
