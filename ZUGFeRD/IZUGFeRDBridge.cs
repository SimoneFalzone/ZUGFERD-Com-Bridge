using System.Runtime.InteropServices;

namespace ZUGFeRDBridge
{
    /// <summary>
    /// COM-sichtbares Interface fuer die ZUGFeRD Bridge.
    /// Alle Parameter verwenden VBA-kompatible Typen (String, Double).
    /// Datumsangaben werden als String im Format "yyyy-MM-dd" uebergeben.
    /// 
    /// LEGENDE:
    ///   [PFLICHT]    = Pflichtfeld fuer eine gueltige ZUGFeRD-Rechnung
    ///   [OPTIONAL]   = Optionales Feld (leerer String "" wenn nicht benoetigt)
    ///   [VERKAEUFER] = Daten des Rechnungsausstellers / Lieferanten
    ///   [KAEUFER]    = Daten des Rechnungsempfaengers / Kunden
    /// </summary>
    [ComVisible(true)]
    [Guid("7C4A9B2E-3D6F-4E1A-8B5C-2F0D7A3E9C1B")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IZUGFeRDBridge
    {
        // ==================================================================
        // RECHNUNG ERSTELLEN / LADEN
        // ==================================================================

        /// <summary>
        /// Erstellt eine neue leere Rechnung im Speicher.
        /// </summary>
        /// <param name="invoiceNo">[PFLICHT] Rechnungsnummer, z.B. "RE-2024-001"</param>
        /// <param name="invoiceDate">[PFLICHT] Rechnungsdatum als "yyyy-MM-dd"</param>
        /// <param name="currency">[PFLICHT] ISO 4217 Waehrungscode, z.B. "EUR"</param>
        void NewInvoice(string invoiceNo, string invoiceDate, string currency);

        /// <summary>
        /// Laedt eine bestehende ZUGFeRD XML-Datei in den Speicher.
        /// </summary>
        /// <param name="filePath">[PFLICHT] Vollstaendiger Pfad zur XML-Datei</param>
        /// <returns>Leerstring bei Erfolg, Fehlermeldung bei Fehler</returns>
        string Load(string filePath);

        // ==================================================================
        // [VERKAEUFER] - Rechnungsaussteller / Lieferant
        // ==================================================================

        /// <summary>
        /// [VERKAEUFER] Setzt die Stammdaten des Rechnungsausstellers (Lieferant/Verkaeufer).
        /// </summary>
        /// <param name="name">[PFLICHT] Firmenname des Verkaeufers</param>
        /// <param name="street">[PFLICHT] Strasse inkl. Hausnummer, z.B. "Hauptstr. 42"</param>
        /// <param name="postCode">[PFLICHT] Postleitzahl</param>
        /// <param name="city">[PFLICHT] Ort</param>
        /// <param name="countryCode">[PFLICHT] ISO 3166-1 alpha-2 Laendercode, z.B. "DE"</param>
        /// <param name="taxId">[OPTIONAL] Steuernummer, z.B. "30/123/45678" (leer="" wenn nicht benoetigt)</param>
        /// <param name="vatId">[OPTIONAL] USt-IdNr., z.B. "DE123456789" (leer="" wenn nicht benoetigt)</param>
        void SetSeller(string name, string street,
                       string postCode, string city, string countryCode,
                       string taxId, string vatId);

        /// <summary>
        /// [VERKAEUFER] [OPTIONAL] Setzt den Ansprechpartner des Verkaeufers.
        /// Alle Parameter sind optional - leerer String wenn nicht benoetigt.
        /// </summary>
        /// <param name="personName">[OPTIONAL] Name des Ansprechpartners</param>
        /// <param name="orgUnit">[OPTIONAL] Abteilung</param>
        /// <param name="email">[OPTIONAL] E-Mail-Adresse</param>
        /// <param name="phone">[OPTIONAL] Telefonnummer</param>
        /// <param name="fax">[OPTIONAL] Faxnummer</param>
        void SetSellerContact(string personName, string orgUnit, string email, string phone, string fax);

        // ==================================================================
        // [KAEUFER] - Rechnungsempfaenger / Kunde
        // ==================================================================

        /// <summary>
        /// [KAEUFER] Setzt die Stammdaten des Rechnungsempfaengers (Kunde/Kaeufer).
        /// </summary>
        /// <param name="name">[PFLICHT] Firmenname des Kaeufers</param>
        /// <param name="street">[PFLICHT] Strasse inkl. Hausnummer</param>
        /// <param name="postCode">[PFLICHT] Postleitzahl</param>
        /// <param name="city">[PFLICHT] Ort</param>
        /// <param name="countryCode">[PFLICHT] ISO 3166-1 alpha-2, z.B. "DE"</param>
        /// <param name="vatId">[OPTIONAL] USt-IdNr. des Kaeufers (leer="" wenn nicht benoetigt)</param>
        void SetBuyer(string name, string street,
                      string postCode, string city, string countryCode,
                      string vatId);

        // ==================================================================
        // [VERKAEUFER] ZAHLUNG - Bankverbindung und Zahlungsbedingungen
        // ==================================================================

        /// <summary>
        /// [VERKAEUFER] [OPTIONAL] Setzt Zahlungsbedingungen.
        /// </summary>
        /// <param name="description">[OPTIONAL] Freitext, z.B. "Zahlbar innerhalb 30 Tagen netto"</param>
        /// <param name="dueDate">[OPTIONAL] Faelligkeitsdatum als "yyyy-MM-dd" (leer="" = kein Datum)</param>
        void SetPaymentTerms(string description, string dueDate);

        /// <summary>
        /// [VERKAEUFER] [OPTIONAL] Fuegt eine Bankverbindung des Verkaeufers hinzu (SEPA).
        /// Kann mehrfach aufgerufen werden fuer mehrere Konten.
        /// </summary>
        /// <param name="iban">[PFLICHT] IBAN des Verkaeufers</param>
        /// <param name="bic">[OPTIONAL] BIC der Bank</param>
        /// <param name="bankName">[OPTIONAL] Name der Bank</param>
        /// <param name="accountName">[OPTIONAL] Kontoinhaber-Name</param>
        void AddBankAccount(string iban, string bic, string bankName, string accountName);

        // ==================================================================
        // RECHNUNGSPOSITIONEN
        // ==================================================================

        /// <summary>
        /// Fuegt eine Rechnungsposition hinzu.
        /// </summary>
        /// <param name="name">[PFLICHT] Kurzbezeichnung / Artikelname</param>
        /// <param name="description">[OPTIONAL] Ausfuehrliche Beschreibung (leer="" wird ignoriert)</param>
        /// <param name="quantity">[PFLICHT] Menge</param>
        /// <param name="unitCode">[PFLICHT] Einheitencode: "C62"=Stueck, "HUR"=Stunde, "KGM"=kg, "MTR"=Meter</param>
        /// <param name="unitPrice">[PFLICHT] Netto-Einzelpreis</param>
        /// <param name="taxPercent">[PFLICHT] USt-Satz in %, z.B. 19.0 oder 7.0</param>
        /// <param name="taxCategoryCode">[PFLICHT] USt-Kategorie: "S"=Standard, "Z"=Null, "E"=Befreit, "AE"=Reverse Charge</param>
        void AddLineItem(string name, string description,
                         double quantity, string unitCode,
                         double unitPrice, double taxPercent,
                         string taxCategoryCode);

        // ==================================================================
        // STEUERN (einmal pro verwendetem Steuersatz aufrufen)
        // ==================================================================

        /// <summary>
        /// Fuegt einen Steuer-Subtotal hinzu. Einmal pro verwendetem USt-Satz aufrufen.
        /// </summary>
        /// <param name="basisAmount">[PFLICHT] Netto-Bemessungsgrundlage fuer diesen Steuersatz</param>
        /// <param name="taxPercent">[PFLICHT] USt-Satz in %</param>
        /// <param name="taxCategoryCode">[PFLICHT] USt-Kategorie: "S", "Z", "E", "AE"</param>
        void AddTax(double basisAmount, double taxPercent, string taxCategoryCode);

        // ==================================================================
        // SUMMEN
        // ==================================================================

        /// <summary>
        /// Setzt die Gesamtsummen der Rechnung.
        /// </summary>
        /// <param name="lineTotalAmount">[PFLICHT] Summe aller Netto-Positionsbetraege</param>
        /// <param name="taxBasisAmount">[PFLICHT] Netto-Gesamtbetrag (Steuerbasis)</param>
        /// <param name="taxTotalAmount">[PFLICHT] Gesamt-USt-Betrag</param>
        /// <param name="grandTotalAmount">[PFLICHT] Brutto-Gesamtbetrag (Netto + USt)</param>
        /// <param name="duePayableAmount">[PFLICHT] Zu zahlender Betrag</param>
        void SetTotals(double lineTotalAmount, double taxBasisAmount,
                       double taxTotalAmount, double grandTotalAmount,
                       double duePayableAmount);

        // ==================================================================
        // SPEICHERN (XML + optional PDF-Einbettung)
        // ==================================================================

        /// <summary>
        /// Speichert die Rechnung als ZUGFeRD XML-Datei (CII-Format).
        /// </summary>
        /// <param name="filePath">[PFLICHT] Vollstaendiger Pfad, z.B. "C:\Rechnungen\RE001.xml"</param>
        /// <param name="version">[OPTIONAL] "Version23" (Standard/empfohlen), "Version20", "Version1"</param>
        /// <param name="profile">[OPTIONAL] "Comfort" (Standard), "Extended", "Basic", "Minimum", "XRechnung"</param>
        /// <returns>Leerstring bei Erfolg, Fehlermeldung bei Fehler</returns>
        string Save(string filePath, string version, string profile);

        /// <summary>
        /// Speichert die ZUGFeRD XML-Daten zusammen mit einem vorhandenen PDF.
        /// Erzeugt ein PDF/A-3 mit eingebettetem ZUGFeRD-XML.
        /// Das Quell-PDF kann z.B. aus MS Access gedruckt werden.
        /// 
        /// WORKFLOW:
        /// 1. MS Access: Bericht als PDF exportieren -> C:\Temp\Rechnung.pdf
        /// 2. VBA: err = bridge.SaveWithPdf("C:\Temp\Rechnung.pdf", "C:\Rechnungen\RE001.pdf", "", "")
        /// </summary>
        /// <param name="pdfInputPath">[PFLICHT] Pfad zur vorhandenen PDF-Rechnung (z.B. aus MS Access)</param>
        /// <param name="pdfOutputPath">[PFLICHT] Pfad fuer die fertige ZUGFeRD-PDF-Datei</param>
        /// <param name="version">[OPTIONAL] "Version23" (Standard), "Version20"</param>
        /// <param name="profile">[OPTIONAL] "Comfort" (Standard), "Extended", "Basic", "XRechnung"</param>
        /// <returns>Leerstring bei Erfolg, Fehlermeldung bei Fehler</returns>
        string SaveWithPdf(string pdfInputPath, string pdfOutputPath, string version, string profile);

        // ==================================================================
        // AUSLESEN (nach Load oder NewInvoice)
        // ==================================================================

        /// <summary>[VERKAEUFER] Gibt den Firmennamen des Verkaeufers zurueck.</summary>
        string GetSellerName();

        /// <summary>[KAEUFER] Gibt den Firmennamen des Kaeufers zurueck.</summary>
        string GetBuyerName();

        /// <summary>Gibt die Rechnungsnummer zurueck.</summary>
        string GetInvoiceNo();

        /// <summary>Gibt das Rechnungsdatum als "yyyy-MM-dd" zurueck.</summary>
        string GetInvoiceDate();

        /// <summary>Gibt den Brutto-Gesamtbetrag als formatierten String zurueck.</summary>
        string GetGrandTotal();

        /// <summary>Gibt die letzte Fehlermeldung zurueck (leer wenn kein Fehler).</summary>
        string GetLastError();
    }
}

