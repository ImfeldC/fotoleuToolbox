using QRCoder;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;

namespace fotoleuToolbox
{
    internal sealed class SwissQrCodeInput
    {
        internal PayloadGenerator.SwissQrCode.Contact Creditor { get; }
        internal PayloadGenerator.SwissQrCode.Contact Debitor { get; }
        internal PayloadGenerator.SwissQrCode.Iban Iban { get; }
        internal PayloadGenerator.SwissQrCode.Currency Currency { get; }
        internal PayloadGenerator.SwissQrCode.AdditionalInformation AdditionalInformation { get; }
        internal PayloadGenerator.SwissQrCode.Reference Reference { get; }
        internal decimal Amount { get; }
        internal PayloadGenerator.SwissQrCode Payload { get; }

        internal SwissQrCodeInput(IDictionary<string, object> values)
        {
            Creditor = ReadContact(values, "Contact");
            Debitor = ReadContact(values, "Debitor");
            Iban = new PayloadGenerator.SwissQrCode.Iban(
                ReadText(values, "IBAN"), PayloadGenerator.SwissQrCode.Iban.IbanType.Iban);

            string currency = ReadText(values, "Currency");
            if (currency == "CHF")
                Currency = PayloadGenerator.SwissQrCode.Currency.CHF;
            else if (currency == "EUR")
                Currency = PayloadGenerator.SwissQrCode.Currency.EUR;
            else
                throw new FormatException("TabQRCode 'Currency' must be CHF or EUR.");

            Amount = ReadAmount(values);
            AdditionalInformation = new PayloadGenerator.SwissQrCode.AdditionalInformation(
                ReadText(values, "UnstructureMessage", true, false),
                ReadText(values, "BillInformation", true, false));

            string reference = ReadText(values, "SCOR", true, false).Replace(" ", "").ToUpperInvariant();
            Reference = reference.Length == 0
                ? new PayloadGenerator.SwissQrCode.Reference(PayloadGenerator.SwissQrCode.Reference.ReferenceType.NON)
                : new PayloadGenerator.SwissQrCode.Reference(
                    PayloadGenerator.SwissQrCode.Reference.ReferenceType.SCOR, reference,
                    PayloadGenerator.SwissQrCode.Reference.ReferenceTextType.CreditorReferenceIso11649);

            Payload = new PayloadGenerator.SwissQrCode(
                Iban, Currency, Creditor, Reference, AdditionalInformation, Debitor, Amount);
        }

        internal QRCodeData CreateQrCode()
        {
            using (var generator = new QRCodeGenerator())
                return generator.CreateQrCode(Payload.ToString(), QRCodeGenerator.ECCLevel.M, forceUtf8: true);
        }

        private static PayloadGenerator.SwissQrCode.Contact ReadContact(
            IDictionary<string, object> values, string prefix)
        {
            string name = ReadText(values, prefix + "Name");
            string postalCode = ReadText(values, prefix + "PostalCode");
            string city = ReadText(values, prefix + "City");
            string country = ReadText(values, prefix + "Country").ToUpperInvariant();
            string street = ReadText(values, prefix + "Street", true);
            string houseNumber = ReadText(values, prefix + "HouseNumber", true);

            try
            {
                return PayloadGenerator.SwissQrCode.Contact.WithStructuredAddress(
                    name, postalCode, city, country, street, houseNumber);
            }
            catch (PayloadGenerator.SwissQrCode.Contact.SwissQrCodeContactException ex)
            {
                throw new FormatException("Invalid TabQRCode address '" + prefix + "': " + ex.Message, ex);
            }
        }

        private static string ReadText(IDictionary<string, object> values, string key,
            bool allowEmpty = false, bool requireKey = true)
        {
            object value;
            if (!values.TryGetValue(key, out value))
            {
                if (!requireKey)
                    return "";
                throw new FormatException("TabQRCode is missing '" + key +
                    "'. Add the structured-address fields before using this add-in.");
            }

            if (value is ErrorWrapper)
                throw new FormatException("TabQRCode '" + key + "' contains an Excel error.");

            string text = Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";
            foreach (char character in text)
                if (char.IsControl(character))
                    throw new FormatException("TabQRCode '" + key + "' must not contain line breaks or control characters.");

            text = text.Trim();
            if ((!allowEmpty && text.Length == 0) ||
                text.Equals("PRUEFEN", StringComparison.OrdinalIgnoreCase) ||
                text.Equals("PR\u00dcFEN", StringComparison.OrdinalIgnoreCase) ||
                text.Equals("n/a", StringComparison.OrdinalIgnoreCase) ||
                text.Equals("n/v", StringComparison.OrdinalIgnoreCase))
                throw new FormatException("TabQRCode '" + key + "' is empty or unreviewed. Correct the source data in Excel.");

            return text;
        }

        private static decimal ReadAmount(IDictionary<string, object> values)
        {
            object value;
            if (!values.TryGetValue("Amount", out value))
                throw new FormatException("TabQRCode is missing 'Amount'.");

            decimal amount;
            if (value is double || value is decimal || value is int)
            {
                try
                {
                    amount = Convert.ToDecimal(value, CultureInfo.InvariantCulture);
                }
                catch (OverflowException ex)
                {
                    throw new FormatException("TabQRCode 'Amount' is not a finite decimal amount.", ex);
                }
            }
            else if (!(value is string) ||
                !decimal.TryParse((string)value, NumberStyles.AllowLeadingWhite |
                    NumberStyles.AllowTrailingWhite | NumberStyles.AllowLeadingSign |
                    NumberStyles.AllowDecimalPoint, CultureInfo.CurrentCulture, out amount))
            {
                throw new FormatException("TabQRCode 'Amount' must be an Excel number or a decimal value " +
                    "in the current Excel/Windows culture, without currency symbols or grouping separators.");
            }

            if (amount < 0.01m || amount > 999999999.99m || decimal.Round(amount, 2) != amount)
                throw new FormatException("TabQRCode 'Amount' must be between 0.01 and 999999999.99 " +
                    "and have no more than two decimal places.");

            return amount;
        }
    }
}
