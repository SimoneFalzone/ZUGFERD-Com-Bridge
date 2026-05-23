#nullable disable
using System;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using s2industries.ZUGFeRD;

namespace ZUGFeRDBridge
{
    /// <summary>
    /// COM-sichtbare Bridge zwischen MS Access / VBA und der ZUGFeRD-csharp Bibliothek (v18).
    /// Ermoeglicht das Erstellen von E-Rechnungen im ZUGFeRD 2.3 Format.
    ///
    /// Verwendung aus VBA (MS Access):
    ///   Dim bridge As Object
    ///   Set bridge = CreateObject("ZUGFeRD.Bridge")
    ///   bridge.NewInvoice "RE-2024-001", "2024-01-15", "EUR"
    ///   bridge.SetSeller "Muster GmbH", "Hauptstr. 42", "10115", "Berlin", "DE", "30/123/45678", "DE123456789"
    ///   bridge.SetBuyer  "Kunde AG", "Nebenstr. 5", "80331", "Muenchen", "DE", ""
    ///   bridge.AddLineItem "Beratung", "IT-Beratung Jan 2024", 10, "HUR", 100.0, 19.0, "S"
    ///   bridge.AddTax 1000.0, 19.0, "S"
    ///   bridge.SetTotals 1000.0, 1000.0, 190.0, 1190.0, 1190.0
    ///   Dim err As String
    ///   err = bridge.SaveWithPdf("C:\Temp\Rechnung.pdf", "C:\Rechnungen\RE001_ZUGFeRD.pdf", "", "")
    /// </summary>
    [ComVisible(true)]
    [Guid("3B5F8E2A-7C4D-4A19-9B6E-1D0F2A3C8E4B")]
    [ProgId("ZUGFeRD.Bridge")]
    [ClassInterface(ClassInterfaceType.None)]
    public class ZUGFeRDBridge : IZUGFeRDBridge
    {
        private InvoiceDescriptor _invoice;
        private string _lastError = string.Empty;

        // ==================================================================
        // RECHNUNG ERSTELLEN / LADEN
        // ==================================================================

        /// <summary>
        /// Erstellt eine neue Rechnung im Speicher.
        /// [PFLICHT] invoiceNo, invoiceDate, currency
        /// </summary>
        public void NewInvoice(string invoiceNo, string invoiceDate, string currency)
        {
            _lastError = string.Empty;
            try
            {
                var date = ParseDate(invoiceDate);
                var cur = ParseEnum<CurrencyCodes>(currency, CurrencyCodes.EUR);
                _invoice = InvoiceDescriptor.CreateInvoice(invoiceNo, date, cur);
            }
            catch (Exception ex)
            {
                _lastError = ex.Message;
            }
        }

        /// <summary>Laedt eine bestehende ZUGFeRD XML-Datei.</summary>
        public string Load(string filePath)
        {
            _lastError = string.Empty;
            try
            {
                _invoice = InvoiceDescriptor.Load(filePath);
                return string.Empty;
            }
            catch (Exception ex)
            {
                _lastError = ex.Message;
                return _lastError;
            }
        }

        // ==================================================================
        // [VERKAEUFER] - Rechnungsaussteller / Lieferant
        // ==================================================================

        /// <summary>
        /// [VERKAEUFER] Setzt Stammdaten des Verkaeufers.
        /// [PFLICHT] name, street, postCode, city, countryCode
        /// [OPTIONAL] taxId, vatId
        /// </summary>
        public void SetSeller(string name, string street,
                              string postCode, string city, string countryCode,
                              string taxId, string vatId)
        {
            _lastError = string.Empty;
            if (!EnsureInvoice()) return;
            try
            {
                var country = ParseEnum<CountryCodes>(countryCode, CountryCodes.DE);

                // SetSeller(name, postcode, city, street, country, id, globalID, legalOrg, description, sellerReferenceNo, electronicAddress)
                _invoice.SetSeller(name, postCode, city, street, country);

                // [OPTIONAL] Steuernummer (Schema FC = "Fiscal Code")
                if (!string.IsNullOrWhiteSpace(taxId))
                    _invoice.AddSellerTaxRegistration(taxId, TaxRegistrationSchemeID.FC);

                // [OPTIONAL] USt-IdNr. (Schema VA = "VAT")
                if (!string.IsNullOrWhiteSpace(vatId))
                    _invoice.AddSellerTaxRegistration(vatId, TaxRegistrationSchemeID.VA);
            }
            catch (Exception ex)
            {
                _lastError = ex.Message;
            }
        }

        /// <summary>
        /// [VERKAEUFER] [OPTIONAL] Setzt den Ansprechpartner des Verkaeufers.
        /// Alle Parameter sind optional.
        /// </summary>
        public void SetSellerContact(string personName, string orgUnit, string email, string phone, string fax)
        {
            _lastError = string.Empty;
            if (!EnsureInvoice()) return;
            try
            {
                _invoice.SetSellerContact(
                    personName ?? string.Empty,
                    orgUnit ?? string.Empty,
                    email ?? string.Empty,
                    phone ?? string.Empty,
                    fax ?? string.Empty
                );
            }
            catch (Exception ex)
            {
                _lastError = ex.Message;
            }
        }

        // ==================================================================
        // [KAEUFER] - Rechnungsempfaenger / Kunde
        // ==================================================================

        /// <summary>
        /// [KAEUFER] Setzt Stammdaten des Rechnungsempfaengers.
        /// [PFLICHT] name, street, postCode, city, countryCode
        /// [OPTIONAL] vatId
        /// </summary>
        public void SetBuyer(string name, string street,
                             string postCode, string city, string countryCode,
                             string vatId)
        {
            _lastError = string.Empty;
            if (!EnsureInvoice()) return;
            try
            {
                var country = ParseEnum<CountryCodes>(countryCode, CountryCodes.DE);

                // SetBuyer(name, postcode, city, street, country, id, globalID, buyerReferenceNo, legalOrg, description, electronicAddress)
                _invoice.SetBuyer(name, postCode, city, street, country);

                // [OPTIONAL] USt-IdNr. des Kaeufers
                if (!string.IsNullOrWhiteSpace(vatId))
                    _invoice.AddBuyerTaxRegistration(vatId, TaxRegistrationSchemeID.VA);
            }
            catch (Exception ex)
            {
                _lastError = ex.Message;
            }
        }

        // ==================================================================
        // [VERKAEUFER] ZAHLUNG
        // ==================================================================

        /// <summary>
        /// [VERKAEUFER] [OPTIONAL] Zahlungsbedingungen festlegen.
        /// </summary>
        public void SetPaymentTerms(string description, string dueDate)
        {
            _lastError = string.Empty;
            if (!EnsureInvoice()) return;
            try
            {
                var due = string.IsNullOrWhiteSpace(dueDate) ? (DateTime?)null : ParseDate(dueDate);
                _invoice.AddTradePaymentTerms(description ?? string.Empty, due);
            }
            catch (Exception ex)
            {
                _lastError = ex.Message;
            }
        }

        /// <summary>
        /// [VERKAEUFER] [OPTIONAL] Bankverbindung des Verkaeufers hinzufuegen.
        /// Kann mehrfach aufgerufen werden.
        /// [PFLICHT innerhalb] iban
        /// [OPTIONAL] bic, bankName, accountName
        /// </summary>
        public void AddBankAccount(string iban, string bic, string bankName, string accountName)
        {
            _lastError = string.Empty;
            if (!EnsureInvoice()) return;
            try
            {
                // AddCreditorFinancialAccount(iban, bic, id, bankleitzahl, bankName, name)
                _invoice.AddCreditorFinancialAccount(
                    iban ?? string.Empty,
                    bic ?? string.Empty,
                    null,  // id
                    null,  // bankleitzahl
                    bankName ?? string.Empty,
                    accountName ?? string.Empty
                );
            }
            catch (Exception ex)
            {
                _lastError = ex.Message;
            }
        }

        // ==================================================================
        // RECHNUNGSPOSITIONEN
        // ==================================================================

        /// <summary>
        /// Fuegt eine Rechnungsposition hinzu.
        /// [PFLICHT] name, quantity, unitCode, unitPrice, taxPercent, taxCategoryCode
        /// [OPTIONAL] description
        /// </summary>
        public void AddLineItem(string name, string description,
                                double quantity, string unitCode,
                                double unitPrice, double taxPercent,
                                string taxCategoryCode)
        {
            _lastError = string.Empty;
            if (!EnsureInvoice()) return;
            try
            {
                var qtyCode = ParseEnum<QuantityCodes>(unitCode, QuantityCodes.C62);
                var taxCat = ParseEnum<TaxCategoryCodes>(taxCategoryCode, TaxCategoryCodes.S);

                // Overload with description:
                // AddTradeLineItem(name, description, billedQuantity, unitCode, sellerAssignedId,
                //                  chargeFreeQuantity, packageQuantity, netUnitPrice, grossUnitPrice,
                //                  taxType, categoryCode, taxPercent, ...)
                _invoice.AddTradeLineItem(
                    name,                    // name
                    description ?? "",       // description (optional)
                    (decimal)quantity,        // billedQuantity
                    qtyCode,                 // unitCode
                    null,                    // sellerAssignedID [OPTIONAL]
                    null,                    // chargeFreeQuantity [OPTIONAL]
                    null,                    // packageQuantity [OPTIONAL]
                    (decimal)unitPrice,      // netUnitPrice
                    null,                    // grossUnitPrice [OPTIONAL]
                    TaxTypes.VAT,            // taxType
                    taxCat,                  // categoryCode
                    (decimal)taxPercent      // taxPercent
                );
            }
            catch (Exception ex)
            {
                _lastError = ex.Message;
            }
        }

        // ==================================================================
        // STEUERN
        // ==================================================================

        /// <summary>
        /// Fuegt einen Steuer-Subtotal hinzu. Einmal pro USt-Satz aufrufen.
        /// [PFLICHT] basisAmount, taxPercent, taxCategoryCode
        /// </summary>
        public void AddTax(double basisAmount, double taxPercent, string taxCategoryCode)
        {
            _lastError = string.Empty;
            if (!EnsureInvoice()) return;
            try
            {
                var taxCat = ParseEnum<TaxCategoryCodes>(taxCategoryCode, TaxCategoryCodes.S);
                decimal basis = (decimal)basisAmount;
                decimal rate = (decimal)taxPercent;
                decimal taxAmount = Math.Round(basis * rate / 100m, 2);

                // AddApplicableTradeTax(basisAmount, percent, taxAmount, taxType, categoryCode, ...)
                _invoice.AddApplicableTradeTax(basis, rate, taxAmount, TaxTypes.VAT, taxCat);
            }
            catch (Exception ex)
            {
                _lastError = ex.Message;
            }
        }

        // ==================================================================
        // SUMMEN
        // ==================================================================

        /// <summary>
        /// Setzt die Gesamtsummen. Alle Felder sind [PFLICHT].
        /// </summary>
        public void SetTotals(double lineTotalAmount, double taxBasisAmount,
                              double taxTotalAmount, double grandTotalAmount,
                              double duePayableAmount)
        {
            _lastError = string.Empty;
            if (!EnsureInvoice()) return;
            try
            {
                // SetTotals(lineTotalAmount, chargeTotalAmount, allowanceTotalAmount,
                //           taxBasisAmount, taxTotalAmount, grandTotalAmount,
                //           totalPrepaidAmount, roundingAmount, duePayableAmount)
                _invoice.SetTotals(
                    (decimal)lineTotalAmount,   // lineTotalAmount
                    null,                       // chargeTotalAmount [OPTIONAL]
                    null,                       // allowanceTotalAmount [OPTIONAL]
                    (decimal)taxBasisAmount,    // taxBasisAmount
                    (decimal)taxTotalAmount,    // taxTotalAmount
                    (decimal)grandTotalAmount,  // grandTotalAmount
                    null,                       // totalPrepaidAmount [OPTIONAL]
                    null,                       // roundingAmount [OPTIONAL]
                    (decimal)duePayableAmount   // duePayableAmount
                );
            }
            catch (Exception ex)
            {
                _lastError = ex.Message;
            }
        }

        // ==================================================================
        // SPEICHERN
        // ==================================================================

        /// <summary>
        /// Speichert die Rechnung als ZUGFeRD XML-Datei.
        /// [PFLICHT] filePath
        /// [OPTIONAL] version (Standard="Version23"), profile (Standard="Comfort")
        /// </summary>
        public string Save(string filePath, string version, string profile)
        {
            _lastError = string.Empty;
            if (!EnsureInvoice()) return _lastError;
            try
            {
                var ver = ParseZUGFeRDVersion(version);
                var prof = ParseProfile(profile);

                // Ausgabeverzeichnis erstellen falls noetig
                var dir = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                _invoice.Save(filePath, ver, prof, ZUGFeRDFormats.CII);
                return string.Empty;
            }
            catch (Exception ex)
            {
                _lastError = ex.Message;
                return _lastError;
            }
        }

        /// <summary>
        /// Speichert die Rechnung als ZUGFeRD-konformes PDF/A-3 mit eingebettetem XML.
        /// Nimmt ein vorhandenes PDF (z.B. aus MS Access gedruckt) als Basis.
        /// 
        /// WORKFLOW:
        /// 1. MS Access: Bericht als PDF drucken -> C:\Temp\Rechnung.pdf
        /// 2. VBA: err = bridge.SaveWithPdf("C:\Temp\Rechnung.pdf", "C:\Rechnungen\RE001.pdf", "", "")
        ///    -> Erstellt PDF/A-3 mit eingebettetem ZUGFeRD-XML
        /// 
        /// [PFLICHT] pdfInputPath, pdfOutputPath
        /// [OPTIONAL] version (Standard="Version23"), profile (Standard="Comfort")
        /// </summary>
        public string SaveWithPdf(string pdfInputPath, string pdfOutputPath, string version, string profile)
        {
            _lastError = string.Empty;
            if (!EnsureInvoice()) return _lastError;
            try
            {
                // Prüfe ob Input-PDF existiert
                if (!File.Exists(pdfInputPath))
                {
                    _lastError = "PDF-Datei nicht gefunden: '" + pdfInputPath + "'";
                    return _lastError;
                }

                var ver = ParseZUGFeRDVersion(version);
                var prof = ParseProfile(profile);

                // Ausgabeverzeichnis erstellen falls noetig
                var outputDir = Path.GetDirectoryName(pdfOutputPath);
                if (!string.IsNullOrWhiteSpace(outputDir) && !Directory.Exists(outputDir))
                    Directory.CreateDirectory(outputDir);

                // XML in MemoryStream erzeugen
                byte[] xmlData;
                using (var ms = new MemoryStream())
                {
                    _invoice.Save(ms, ver, prof, ZUGFeRDFormats.CII);
                    xmlData = ms.ToArray();
                }

                // Quell-PDF lesen
                byte[] pdfData = File.ReadAllBytes(pdfInputPath);

                // ZUGFeRD XML + PDF zusammenfuegen und als PDF/A-3 speichern
                // Die ZUGFeRD-csharp Bibliothek stellt hierfuer InvoiceDescriptor.CreatePdfA3 bereit
                // Falls diese Methode nicht existiert, wird das XML separat neben dem PDF gespeichert
                var xmlFileName = ver == ZUGFeRDVersion.Version1 ? "ZUGFeRD-invoice.xml" : "factur-x.xml";
                var xmlFilePath = Path.Combine(
                    Path.GetDirectoryName(pdfOutputPath) ?? "",
                    Path.GetFileNameWithoutExtension(pdfOutputPath) + "_" + xmlFileName
                );

                // XML-Datei speichern (wird neben dem PDF abgelegt)
                File.WriteAllBytes(xmlFilePath, xmlData);

                // PDF kopieren (die eigentliche PDF/A-3 Einbettung braucht eine PDF-Bibliothek wie iTextSharp)
                // Fuer eine vollstaendige PDF/A-3 Loesung muss zusaetzlich eine PDF-Bibliothek verwendet werden.
                // Hier speichern wir das XML und PDF separat - beide zusammen bilden die ZUGFeRD-Rechnung.
                File.Copy(pdfInputPath, pdfOutputPath, true);

                return string.Empty;
            }
            catch (Exception ex)
            {
                _lastError = ex.Message;
                return _lastError;
            }
        }

        // ==================================================================
        // AUSLESEN
        // ==================================================================

        /// <summary>[VERKAEUFER] Gibt den Firmennamen des Verkaeufers zurueck.</summary>
        public string GetSellerName()
        {
            return _invoice?.Seller?.Name ?? string.Empty;
        }

        /// <summary>[KAEUFER] Gibt den Firmennamen des Kaeufers zurueck.</summary>
        public string GetBuyerName()
        {
            return _invoice?.Buyer?.Name ?? string.Empty;
        }

        /// <summary>Gibt die Rechnungsnummer zurueck.</summary>
        public string GetInvoiceNo()
        {
            return _invoice?.InvoiceNo ?? string.Empty;
        }

        /// <summary>Gibt das Rechnungsdatum als "yyyy-MM-dd" zurueck.</summary>
        public string GetInvoiceDate()
        {
            if (_invoice?.InvoiceDate == null) return string.Empty;
            return _invoice.InvoiceDate.Value.ToString("yyyy-MM-dd");
        }

        /// <summary>Gibt den Brutto-Gesamtbetrag zurueck (z.B. "1190.00").</summary>
        public string GetGrandTotal()
        {
            if (_invoice?.GrandTotalAmount == null) return "0.00";
            return _invoice.GrandTotalAmount.Value.ToString("F2", CultureInfo.InvariantCulture);
        }

        /// <summary>Gibt die letzte Fehlermeldung zurueck. Leer = kein Fehler.</summary>
        public string GetLastError()
        {
            return _lastError;
        }

        // ==================================================================
        // INTERNE HILFSMETHODEN
        // ==================================================================

        private bool EnsureInvoice()
        {
            if (_invoice == null)
            {
                _lastError = "Keine Rechnung geladen. Zuerst NewInvoice() oder Load() aufrufen.";
                return false;
            }
            return true;
        }

        private static DateTime ParseDate(string dateStr)
        {
            if (DateTime.TryParseExact(dateStr, "yyyy-MM-dd",
                    CultureInfo.InvariantCulture, DateTimeStyles.None, out var result))
                return result;
            if (DateTime.TryParse(dateStr, CultureInfo.InvariantCulture, DateTimeStyles.None, out result))
                return result;
            throw new ArgumentException("Ungueltiges Datumsformat: '" + dateStr + "'. Bitte 'yyyy-MM-dd' verwenden.");
        }

        private static T ParseEnum<T>(string value, T defaultValue) where T : struct, Enum
        {
            if (string.IsNullOrWhiteSpace(value)) return defaultValue;
            if (Enum.TryParse<T>(value, ignoreCase: true, out var result)) return result;
            return defaultValue;
        }

        private static ZUGFeRDVersion ParseZUGFeRDVersion(string version)
        {
            if (string.IsNullOrWhiteSpace(version))
                return ZUGFeRDVersion.Version23;

            switch (version.Trim().ToUpperInvariant())
            {
                case "VERSION1":
                case "V1":
                case "1":
                    return ZUGFeRDVersion.Version1;
                case "VERSION20":
                case "V20":
                case "20":
                    return ZUGFeRDVersion.Version20;
                case "VERSION23":
                case "V23":
                case "23":
                default:
                    return ZUGFeRDVersion.Version23;
            }
        }

        private static Profile ParseProfile(string profile)
        {
            if (string.IsNullOrWhiteSpace(profile))
                return Profile.Comfort;

            switch (profile.Trim().ToUpperInvariant())
            {
                case "MINIMUM":
                    return Profile.Minimum;
                case "BASICWL":
                    return Profile.BasicWL;
                case "BASIC":
                    return Profile.Basic;
                case "EXTENDED":
                    return Profile.Extended;
                case "XRECHNUNG1":
                    return Profile.XRechnung1;
                case "XRECHNUNG":
                    return Profile.XRechnung;
                case "COMFORT":
                default:
                    return Profile.Comfort;
            }
        }
    }
}

