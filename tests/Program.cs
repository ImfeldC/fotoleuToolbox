using fotoleuToolbox;
using QRCoder;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using ZXing;
using ZXing.Common;
using FormatException = System.FormatException;
using SwissQrCode = QRCoder.PayloadGenerator.SwissQrCode;

internal static class Program
{
    private static int passed;

    private static int Main(string[] args)
    {
        try
        {
            Run("Structured fields, reserved lines, SCOR and version", StructuredPayload);
            Run("Blank optional street and house number", OptionalFields);
            Run("Missing structured keys do not fall back to combined addresses", MissingKeys);
            Run("Empty mandatory fields and migration markers are rejected", InvalidFields);
            Run("Current SIX character repertoire and field lengths", CharactersAndLengths);
            Run("Normal IBAN, QR-IBAN and SCOR validation", AccountsAndReferences);
            Run("Amount bounds, precision and culture-independent numeric Excel values", Amounts);
            Run("UTF-8 bytes, no BOM and ECC M decoded from QR modules", DecodeModules);
            Run("QR bitmap with Swiss cross decodes correctly", DecodeBitmap);

            if (args.Length != 0)
            {
                if (args.Length != 2 || args[0] != "--workbook")
                    throw new ArgumentException("Usage: --workbook <path-to-xlsm>");
                Run("Supplied workbook cached inputs generate a structured UTF-8 QR",
                    () => DecodeInput(new SwissQrCodeInput(ReadWorkbook(args[1]))));
            }

            Console.WriteLine("Passed " + passed + " test groups.");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex);
            return 1;
        }
    }

    private static Dictionary<string, object> ValidValues()
    {
        return new Dictionary<string, object>
        {
            { "ContactName", "Muster GmbH" }, { "ContactStreet", "Musterstrasse" },
            { "ContactHouseNumber", "12a" }, { "ContactPostalCode", "6300" },
            { "ContactCity", "Zug" }, { "ContactCountry", "CH" },
            { "DebitorName", "Example Customer" }, { "DebitorStreet", "Example Road" },
            { "DebitorHouseNumber", "3/1" }, { "DebitorPostalCode", "01234" },
            { "DebitorCity", "Example Town" }, { "DebitorCountry", "DE" },
            { "IBAN", "CH93 0076 2011 6238 5295 7" }, { "Currency", "CHF" },
            { "Amount", 123.45d }, { "SCOR", "RF18539007547034" },
            { "UnstructureMessage", "Invoice 42" }, { "BillInformation", "" }
        };
    }

    private static void StructuredPayload()
    {
        string[] lines = new SwissQrCodeInput(ValidValues()).Payload.ToString().Split(new[] { "\r\n" }, StringSplitOptions.None);
        Equal("SPC", lines[0]);
        Equal("0200", lines[1]);
        Equal("1", lines[2]);
        Equal("S", lines[4]);
        Equal("Musterstrasse", lines[6]);
        Equal("12a", lines[7]);
        Equal("6300", lines[8]);
        Equal("Zug", lines[9]);
        Equal("CH", lines[10]);
        for (int i = 11; i <= 17; i++)
            Equal("", lines[i]);
        Equal("123.45", lines[18]);
        Equal("CHF", lines[19]);
        Equal("S", lines[20]);
        Equal("01234", lines[24]);
        Equal("SCOR", lines[27]);
        Equal("RF18539007547034", lines[28]);
        Equal("EPD", lines[30]);
        Equal(31, lines.Length);
    }

    private static void OptionalFields()
    {
        var values = ValidValues();
        foreach (string key in new[] { "ContactStreet", "ContactHouseNumber", "DebitorStreet", "DebitorHouseNumber" })
            values[key] = null;
        values["SCOR"] = "";
        var input = new SwissQrCodeInput(values);
        string[] lines = input.Payload.ToString().Split(new[] { "\r\n" }, StringSplitOptions.None);
        Equal("", lines[6]);
        Equal("", lines[7]);
        Equal("", lines[22]);
        Equal("", lines[23]);
        Equal("NON", lines[27]);
        Equal("", lines[28]);
        DecodeInput(input);
    }

    private static void MissingKeys()
    {
        foreach (string prefix in new[] { "Contact", "Debitor" })
            foreach (string suffix in new[] { "Name", "Street", "HouseNumber", "PostalCode", "City", "Country" })
            {
                var values = ValidValues();
                string key = prefix + suffix;
                values.Remove(key);
                values[prefix + "AdressLine1"] = "Legacy address 12";
                values[prefix + "AdressLine2"] = "6300 Zug";
                Reject(() => new SwissQrCodeInput(values), key);
            }
    }

    private static void InvalidFields()
    {
        foreach (string prefix in new[] { "Contact", "Debitor" })
            foreach (string suffix in new[] { "Name", "Street", "HouseNumber", "PostalCode", "City", "Country" })
                foreach (string marker in new[] { "PRUEFEN", "pruefen", "PR\u00dcFEN", "n/a", "n/v", "bad\r\nvalue", "bad\tvalue" })
                {
                    var values = ValidValues();
                    string key = prefix + suffix;
                    values[key] = marker;
                    Reject(() => new SwissQrCodeInput(values), key);
                }

        foreach (string suffix in new[] { "Name", "PostalCode", "City", "Country" })
        {
            var values = ValidValues();
            values["Debitor" + suffix] = "   ";
            Reject(() => new SwissQrCodeInput(values), "Debitor" + suffix);
        }
        var error = ValidValues();
        error["DebitorCity"] = new ErrorWrapper(unchecked((int)0x800A07FA));
        Reject(() => new SwissQrCodeInput(error), "DebitorCity");
    }

    private static void CharactersAndLengths()
    {
        var values = ValidValues();
        values["ContactName"] = "M\u00fcller \u010c\u0159 \u0218\u0219\u021a\u021b \u20ac";
        values["DebitorCountry"] = "de";
        Equal("DE", new SwissQrCodeInput(values).Debitor.ToString().Split(new[] { "\r\n" }, StringSplitOptions.None)[6]);
        DecodeInput(new SwissQrCodeInput(values));

        foreach (string value in new[] { "Name\n", "\u0301", "\ud83d\ude00" })
        {
            values = ValidValues();
            values["DebitorName"] = value;
            Reject(() => new SwissQrCodeInput(values));
        }
        foreach (var limit in new[] {
            new KeyValuePair<string, int>("Name", 70), new KeyValuePair<string, int>("Street", 70),
            new KeyValuePair<string, int>("HouseNumber", 16), new KeyValuePair<string, int>("PostalCode", 16),
            new KeyValuePair<string, int>("City", 35) })
        {
            values = ValidValues();
            values["Debitor" + limit.Key] = new string('x', limit.Value);
            new SwissQrCodeInput(values);
            values["Debitor" + limit.Key] = new string('x', limit.Value + 1);
            Reject(() => new SwissQrCodeInput(values));
        }
        foreach (string key in new[] { "UnstructureMessage", "BillInformation", "DebitorHouseNumber" })
        {
            values = ValidValues();
            values[key] = "bad\rvalue";
            Reject(() => new SwissQrCodeInput(values));
        }
        values = ValidValues();
        values["DebitorCountry"] = "XX";
        Reject(() => new SwissQrCodeInput(values));
    }

    private static void AccountsAndReferences()
    {
        foreach (string account in new[] { "CH4431999123000889012", "CH9300762011623852958", "DE89370400440532013000", "CH1" })
        {
            var values = ValidValues();
            values["IBAN"] = account;
            Reject(() => new SwissQrCodeInput(values));
        }
        foreach (string reference in new[] { "RF00539007547034", "INVALID", "RF18", "RF18539007547034\n" })
        {
            var values = ValidValues();
            values["SCOR"] = reference;
            Reject(() => new SwissQrCodeInput(values));
        }
        var valid = ValidValues();
        valid["SCOR"] = "rf18 5390 0754 7034";
        new SwissQrCodeInput(valid);
        valid["Currency"] = "EUR";
        DecodeInput(new SwissQrCodeInput(valid));
    }

    private static void Amounts()
    {
        CultureInfo previous = CultureInfo.CurrentCulture;
        try
        {
            foreach (string culture in new[] { "de-CH", "de-DE", "en-US", "ar-SA" })
            {
                CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo(culture);
                foreach (object amount in new object[] { 0.01m, 999999999.99m, 123.45d, 1.2300m })
                {
                    var values = ValidValues();
                    values["Amount"] = amount;
                    var input = new SwissQrCodeInput(values);
                    Equal(Convert.ToDecimal(amount, CultureInfo.InvariantCulture).ToString("0.00", CultureInfo.InvariantCulture),
                        input.Payload.ToString().Split(new[] { "\r\n" }, StringSplitOptions.None)[18]);
                }
                foreach (object amount in new object[] { 0m, -1m, 1000000000m, 1.001m, double.NaN, double.PositiveInfinity, "CHF 12.00", true, null })
                {
                    var values = ValidValues();
                    values["Amount"] = amount;
                    Reject(() => new SwissQrCodeInput(values), "Amount");
                }
            }
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("de-DE");
            var text = ValidValues();
            text["Amount"] = "12,34";
            Equal(12.34m, new SwissQrCodeInput(text).Amount);
            text["Amount"] = "1.234,00";
            Reject(() => new SwissQrCodeInput(text), "Amount");
        }
        finally
        {
            CultureInfo.CurrentCulture = previous;
        }
    }

    private static void DecodeModules()
    {
        var values = ValidValues();
        values["ContactName"] = "M\u00fcller & S\u00f6hne";
        DecodeInput(new SwissQrCodeInput(values));
    }

    private static void DecodeInput(SwissQrCodeInput input)
    {
        using (QRCodeData qr = input.CreateQrCode())
        {
            int size = qr.ModuleMatrix.Count - 8;
            var matrix = new BitMatrix(size);
            for (int y = 0; y < size; y++)
                for (int x = 0; x < size; x++)
                    matrix[x, y] = qr.ModuleMatrix[y + 4][x + 4];
            DecoderResult result = new ZXing.QrCode.Internal.Decoder().decode(matrix, null);
            Equal("M", result.ECLevel);
            EqualBytes(Encoding.UTF8.GetBytes(input.Payload.ToString()), result.ByteSegments.SelectMany(segment => segment).ToArray());
        }
    }

    private static void DecodeBitmap()
    {
        var values = ValidValues();
        values["ContactName"] = "M\u00fcller & S\u00f6hne";
        var input = new SwissQrCodeInput(values);
        using (QRCodeData data = input.CreateQrCode())
        using (var qr = new QRCode(data))
        using (var icon = new Bitmap(Path.Combine(AppContext.BaseDirectory, "CH-Kreuz_7mm.png")))
        using (Bitmap bitmap = qr.GetGraphic(20, Color.Black, Color.White, icon, 14, 1))
        {
            BitmapData pixels = bitmap.LockBits(new Rectangle(0, 0, bitmap.Width, bitmap.Height),
                ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            byte[] bytes = new byte[pixels.Stride * bitmap.Height];
            try
            {
                Marshal.Copy(pixels.Scan0, bytes, 0, bytes.Length);
            }
            finally
            {
                bitmap.UnlockBits(pixels);
            }
            var luminance = new RGBLuminanceSource(bytes, bitmap.Width, bitmap.Height, RGBLuminanceSource.BitmapFormat.BGR32);
            Result result = new ZXing.QrCode.QRCodeReader().decode(new BinaryBitmap(new HybridBinarizer(luminance)));
            if (result == null)
                throw new Exception("Swiss-cross bitmap could not be decoded.");
            var segments = (IList<byte[]>)result.ResultMetadata[ResultMetadataType.BYTE_SEGMENTS];
            EqualBytes(Encoding.UTF8.GetBytes(input.Payload.ToString()), segments.SelectMany(segment => segment).ToArray());
        }
    }

    private static IDictionary<string, object> ReadWorkbook(string path)
    {
        XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        using (ZipArchive archive = ZipFile.OpenRead(path))
        {
            Func<string, XDocument> readXml = entry =>
            {
                using (Stream stream = archive.GetEntry(entry).Open())
                    return XDocument.Load(stream);
            };
            string[] strings = readXml("xl/sharedStrings.xml").Descendants(spreadsheet + "si")
                .Select(item => string.Concat(item.Descendants(spreadsheet + "t").Select(text => text.Value))).ToArray();
            foreach (ZipArchiveEntry sheet in archive.Entries.Where(entry => entry.FullName.StartsWith("xl/worksheets/sheet") && entry.FullName.EndsWith(".xml")))
            {
                string relPath = "xl/worksheets/_rels/" + Path.GetFileName(sheet.FullName) + ".rels";
                if (archive.GetEntry(relPath) == null)
                    continue;
                foreach (XElement relation in readXml(relPath).Root.Elements().Where(element => ((string)element.Attribute("Type")).EndsWith("/table")))
                {
                    string tablePath = "xl/" + ((string)relation.Attribute("Target")).Replace("../", "");
                    XElement table = readXml(tablePath).Root;
                    if ((string)table.Attribute("name") != "TabQRCode")
                        continue;
                    int[] rows = System.Text.RegularExpressions.Regex.Matches((string)table.Attribute("ref"), @"\d+")
                        .Cast<System.Text.RegularExpressions.Match>().Select(match => int.Parse(match.Value, CultureInfo.InvariantCulture)).ToArray();
                    var result = new Dictionary<string, object>();
                    foreach (XElement row in readXml(sheet.FullName).Descendants(spreadsheet + "row")
                        .Where(row => (int)row.Attribute("r") > rows[0] && (int)row.Attribute("r") <= rows[1]))
                    {
                        XElement[] cells = row.Elements(spreadsheet + "c").ToArray();
                        Func<XElement, object> value = cell =>
                        {
                            if (cell == null)
                                return null;
                            string type = (string)cell.Attribute("t");
                            string text = (string)cell.Element(spreadsheet + "v");
                            if (type == "e")
                                throw new Exception("Workbook input contains a cached Excel error.");
                            if (text == null)
                                return null;
                            if (type == "s")
                                return strings[int.Parse(text, CultureInfo.InvariantCulture)];
                            if (type == "str")
                                return text;
                            return double.Parse(text, CultureInfo.InvariantCulture);
                        };
                        string rowNumber = (string)row.Attribute("r");
                        result.Add((string)value(cells.Single(cell => (string)cell.Attribute("r") == "D" + rowNumber)),
                            value(cells.SingleOrDefault(cell => (string)cell.Attribute("r") == "E" + rowNumber)));
                    }
                    return result;
                }
            }
        }
        throw new Exception("TabQRCode was not found in the workbook.");
    }

    private static void Run(string name, Action test)
    {
        test();
        passed++;
        Console.WriteLine("PASS " + name);
    }

    private static void Reject(Action action, string expectedMessage = null)
    {
        try
        {
            action();
        }
        catch (Exception ex) when (ex is FormatException || ex is SwissQrCode.SwissQrCodeException ||
            ex is SwissQrCode.Contact.SwissQrCodeContactException || ex is SwissQrCode.Iban.SwissQrCodeIbanException ||
            ex is SwissQrCode.Reference.SwissQrCodeReferenceException || ex is SwissQrCode.AdditionalInformation.SwissQrCodeAdditionalInformationException)
        {
            if (expectedMessage != null && !ex.Message.Contains(expectedMessage))
                throw new Exception("Validation did not identify field " + expectedMessage + ".", ex);
            return;
        }
        throw new Exception("Expected invalid input to be rejected.");
    }

    private static void Equal<T>(T expected, T actual)
    {
        if (!EqualityComparer<T>.Default.Equals(expected, actual))
            throw new Exception("Assertion failed: expected " + expected + ", got " + actual + ".");
    }

    private static void EqualBytes(byte[] expected, byte[] actual)
    {
        if (!expected.SequenceEqual(actual))
            throw new Exception("Decoded QR bytes do not match the UTF-8 payload.");
    }
}
