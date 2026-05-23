# ZUGFeRD Bridge für MS Access

COM-sichtbare .NET DLL die es ermöglicht, aus MS Access / VBA heraus E-Rechnungen im **ZUGFeRD 2.3** Format zu erstellen.

## Voraussetzungen

- Windows 10/11
- .NET Framework 4.8 (bei Windows 10/11 vorinstalliert)
- MS Access (32-bit oder 64-bit)

---

## Installation auf Windows

### Schritt 1: Build erstellen

Auf dem Mac (oder in CI/CD):
```bash
cd ZUGFeRD
dotnet build -c Release
```

Die Ausgabe befindet sich in: `bin/Release/net48/`

### Schritt 2: DLL und Abhängigkeiten kopieren

Kopiere den gesamten Inhalt des `bin/Release/net48/` Ordners in einen festen Ordner auf dem Windows-PC:

```
C:\Programme\ZUGFeRDBridge\
    ZUGFeRDBridge.dll
    s2industries.ZUGFeRD.dll
    (weitere Abhängigkeiten falls vorhanden)
```

### Schritt 3: DLL als COM-Objekt registrieren

**Eingabeaufforderung als Administrator öffnen** und dann:

**Für 64-bit MS Access:**
```cmd
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe "C:\Programme\ZUGFeRDBridge\ZUGFeRDBridge.dll" /tlb /codebase
```

**Für 32-bit MS Access:**
```cmd
C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe "C:\Programme\ZUGFeRDBridge\ZUGFeRDBridge.dll" /tlb /codebase
```

> ⚠️ **WICHTIG**: Die richtige Version (32/64-bit) muss zur MS Access Installation passen!  
> Prüfe in Access: Datei → Konto → Info über Access → ob "(32-Bit)" oder "(64-Bit)" steht.

### Schritt 4: Registrierung prüfen

In MS Access VBA (Alt+F11) testen:
```vba
Sub TestBridge()
    Dim bridge As Object
    Set bridge = CreateObject("ZUGFeRD.Bridge")
    MsgBox "Bridge erfolgreich geladen!", vbInformation
    Set bridge = Nothing
End Sub
```

### Deinstallation

```cmd
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe "C:\Programme\ZUGFeRDBridge\ZUGFeRDBridge.dll" /unregister
```

---

## Verwendung in MS Access VBA

### Komplettes Beispiel: Rechnung erstellen und als ZUGFeRD-PDF speichern

```vba
' ═══════════════════════════════════════════════════════════════
' ZUGFeRD E-Rechnung erstellen aus MS Access
' ═══════════════════════════════════════════════════════════════
Public Sub RechnungAlsZUGFeRD(ByVal RechnungsNr As String)
    
    Dim bridge As Object
    Dim err As String
    Dim pdfInput As String
    Dim pdfOutput As String
    
    ' ── Bridge-Objekt erstellen ──
    Set bridge = CreateObject("ZUGFeRD.Bridge")
    
    ' ══════════════════════════════════════════════════════════════
    ' SCHRITT 1: Neue Rechnung anlegen
    ' [PFLICHT] Rechnungsnummer, Datum (yyyy-MM-dd), Währung
    ' ══════════════════════════════════════════════════════════════
    bridge.NewInvoice RechnungsNr, Format(Date, "yyyy-MM-dd"), "EUR"
    
    ' Fehlerprüfung nach jedem Schritt:
    If bridge.GetLastError <> "" Then
        MsgBox "Fehler: " & bridge.GetLastError, vbCritical
        GoTo Cleanup
    End If
    
    ' ══════════════════════════════════════════════════════════════
    ' SCHRITT 2: [VERKÄUFER] Eigene Firmendaten setzen
    ' Parameter: Name, Straße+HausNr, PLZ, Ort, Land, Steuernummer, USt-IdNr
    '            [PFLICHT]                                [OPTIONAL]  [OPTIONAL]
    ' ══════════════════════════════════════════════════════════════
    bridge.SetSeller "Muster GmbH", _
                     "Hauptstraße 42", _
                     "10115", "Berlin", "DE", _
                     "30/123/45678", _
                     "DE123456789"
    
    ' [VERKÄUFER] [OPTIONAL] Ansprechpartner
    ' Parameter: Name, Abteilung, E-Mail, Telefon, Fax (alle optional)
    bridge.SetSellerContact "Max Mustermann", "Buchhaltung", _
                            "max@muster-gmbh.de", "+49 30 12345678", ""
    
    ' [VERKÄUFER] [OPTIONAL] Bankverbindung
    ' Parameter: IBAN [PFLICHT], BIC [OPTIONAL], Bankname [OPTIONAL], Kontoinhaber [OPTIONAL]
    bridge.AddBankAccount "DE89370400440532013000", "COBADEFFXXX", "Commerzbank", "Muster GmbH"
    
    ' [VERKÄUFER] [OPTIONAL] Zahlungsbedingungen
    ' Parameter: Beschreibung [OPTIONAL], Fälligkeitsdatum [OPTIONAL]
    bridge.SetPaymentTerms "Zahlbar innerhalb 30 Tagen ohne Abzug", _
                           Format(DateAdd("d", 30, Date), "yyyy-MM-dd")
    
    ' ══════════════════════════════════════════════════════════════
    ' SCHRITT 3: [KÄUFER] Kundendaten setzen
    ' Parameter: Name, Straße+HausNr, PLZ, Ort, Land, USt-IdNr
    '            [PFLICHT]                           [OPTIONAL]
    ' ══════════════════════════════════════════════════════════════
    bridge.SetBuyer "Kunde AG", _
                    "Kundenweg 7", _
                    "80331", "München", "DE", _
                    "DE987654321"
    
    ' ══════════════════════════════════════════════════════════════
    ' SCHRITT 4: Rechnungspositionen hinzufügen
    ' Parameter: Name [PFLICHT], Beschreibung [OPTIONAL], Menge [PFLICHT],
    '            Einheit [PFLICHT], Einzelpreis [PFLICHT], USt% [PFLICHT],
    '            USt-Kategorie [PFLICHT]
    '
    ' Einheiten: "C62"=Stück, "HUR"=Stunde, "KGM"=Kilogramm, "MTR"=Meter
    ' USt-Kategorie: "S"=Standard(19%/7%), "Z"=Null(0%), "E"=Befreit,
    '                "AE"=Reverse Charge
    ' ══════════════════════════════════════════════════════════════
    
    bridge.AddLineItem "IT-Beratung", _
                       "Systemanalyse und Konzeption", _
                       10, "HUR", 120#, 19#, "S"
    
    bridge.AddLineItem "Softwarelizenz", _
                       "Jahreslizenz ERP-System", _
                       1, "C62", 500#, 19#, "S"
    
    ' Position ohne Beschreibung (description = "")
    bridge.AddLineItem "Versandkosten", "", 1, "C62", 15#, 19#, "S"
    
    ' ══════════════════════════════════════════════════════════════
    ' SCHRITT 5: Steuern (einmal pro USt-Satz)
    ' Parameter: Nettobasis [PFLICHT], USt% [PFLICHT], Kategorie [PFLICHT]
    ' Nettobasis = Summe aller Positionen mit diesem Steuersatz
    ' ══════════════════════════════════════════════════════════════
    Dim nettoBasis As Double
    nettoBasis = (10 * 120) + (1 * 500) + (1 * 15)  ' = 1715.00
    bridge.AddTax nettoBasis, 19#, "S"
    
    ' Bei gemischten Steuersätzen (z.B. 7% + 19%):
    ' bridge.AddTax 100.0, 7.0, "S"   ' 7% Basis
    ' bridge.AddTax 1615.0, 19.0, "S"  ' 19% Basis
    
    ' ══════════════════════════════════════════════════════════════
    ' SCHRITT 6: Gesamtsummen setzen (alle [PFLICHT])
    ' Parameter: Positionssumme, Steuerbasis, USt-Betrag,
    '            Brutto-Gesamt, Zahlbetrag
    ' ══════════════════════════════════════════════════════════════
    Dim nettoGesamt As Double
    Dim ustGesamt As Double
    Dim bruttoGesamt As Double
    
    nettoGesamt = 1715#
    ustGesamt = Round(1715 * 0.19, 2)    ' = 325.85
    bruttoGesamt = nettoGesamt + ustGesamt  ' = 2040.85
    
    bridge.SetTotals nettoGesamt, nettoGesamt, ustGesamt, bruttoGesamt, bruttoGesamt
    
    ' ══════════════════════════════════════════════════════════════
    ' SCHRITT 7: Als ZUGFeRD-PDF speichern
    ' ══════════════════════════════════════════════════════════════
    
    ' 7a: Zuerst den Access-Bericht als PDF drucken
    pdfInput = "C:\Temp\" & RechnungsNr & ".pdf"
    DoCmd.OutputTo acOutputReport, "rptRechnung", acFormatPDF, pdfInput
    
    ' 7b: Dann ZUGFeRD-XML erzeugen und neben dem PDF speichern
    '     Version: "Version23" = ZUGFeRD 2.3 (empfohlen)
    '     Profil:  "Comfort" (empfohlen), "Extended", "Basic", "Minimum", "XRechnung"
    pdfOutput = "C:\Rechnungen\" & RechnungsNr & "_ZUGFeRD.pdf"
    err = bridge.SaveWithPdf(pdfInput, pdfOutput, "Version23", "Comfort")
    
    If err <> "" Then
        MsgBox "Fehler beim Speichern: " & err, vbCritical
    Else
        MsgBox "ZUGFeRD-Rechnung erstellt:" & vbCrLf & pdfOutput, vbInformation
    End If
    
    ' ══════════════════════════════════════════════════════════════
    ' [ALTERNATIV] Nur XML speichern (ohne PDF)
    ' ══════════════════════════════════════════════════════════════
    ' err = bridge.Save("C:\Rechnungen\" & RechnungsNr & ".xml", "Version23", "Comfort")

Cleanup:
    Set bridge = Nothing
    
End Sub
```

---

### Nur XML speichern (ohne PDF)

```vba
Public Sub NurXmlSpeichern()
    Dim bridge As Object
    Set bridge = CreateObject("ZUGFeRD.Bridge")
    
    bridge.NewInvoice "RE-2024-099", "2024-06-15", "EUR"
    bridge.SetSeller "Firma", "Musterstr. 1", "12345", "Stadt", "DE", "", "DE111222333"
    bridge.SetBuyer "Kunde", "Kundenweg 2", "54321", "Dorf", "DE", ""
    bridge.AddLineItem "Produkt", "", 5, "C62", 50#, 19#, "S"
    bridge.AddTax 250#, 19#, "S"
    bridge.SetTotals 250#, 250#, 47.5, 297.5, 297.5
    
    Dim err As String
    err = bridge.Save("C:\Rechnungen\RE-2024-099.xml", "Version23", "Comfort")
    If err <> "" Then MsgBox err
    
    Set bridge = Nothing
End Sub
```

---

### Bestehende ZUGFeRD-Rechnung lesen

```vba
Public Sub RechnungLesen()
    Dim bridge As Object
    Set bridge = CreateObject("ZUGFeRD.Bridge")
    
    Dim err As String
    err = bridge.Load("C:\Rechnungen\RE-2024-001.xml")
    
    If err <> "" Then
        MsgBox "Fehler: " & err
    Else
        MsgBox "Rechnungsnr: " & bridge.GetInvoiceNo & vbCrLf & _
               "Datum: " & bridge.GetInvoiceDate & vbCrLf & _
               "Verkäufer: " & bridge.GetSellerName & vbCrLf & _
               "Käufer: " & bridge.GetBuyerName & vbCrLf & _
               "Betrag: " & bridge.GetGrandTotal & " EUR"
    End If
    
    Set bridge = Nothing
End Sub
```

---

## Workflow: PDF aus MS Access + ZUGFeRD

```
┌─────────────────────────────────────────────────────────────────┐
│  MS Access                                                       │
│                                                                  │
│  1. Rechnungsdaten aus Tabellen lesen                           │
│  2. Access-Bericht als PDF drucken                              │
│     → DoCmd.OutputTo acOutputReport, "rptRechnung",             │
│       acFormatPDF, "C:\Temp\Rechnung.pdf"                       │
│  3. ZUGFeRD Bridge aufrufen                                     │
│     → Set bridge = CreateObject("ZUGFeRD.Bridge")               │
│  4. Rechnungsdaten an Bridge übergeben                          │
│     → bridge.SetSeller, bridge.SetBuyer, bridge.AddLineItem...  │
│  5. bridge.SaveWithPdf(inputPdf, outputPdf, version, profile)   │
│     → Erzeugt PDF + ZUGFeRD-XML nebeneinander                  │
│                                                                  │
│  Ergebnis: PDF-Rechnung + ZUGFeRD-XML (E-Rechnungs-konform)    │
└─────────────────────────────────────────────────────────────────┘
```

---

## Unterstützte ZUGFeRD-Versionen & Profile

| Version | Parameter | Beschreibung |
|---------|-----------|--------------|
| ZUGFeRD 2.3 | `"Version23"` | **Empfohlen** – aktuelle Version |
| ZUGFeRD 2.0 | `"Version20"` | Ältere Version |
| ZUGFeRD 1.0 | `"Version1"` | Veraltet |

| Profil | Parameter | Beschreibung |
|--------|-----------|--------------|
| Comfort | `"Comfort"` | **Empfohlen** – deckt die meisten Anwendungsfälle ab |
| Extended | `"Extended"` | Erweitert – mehr Felder möglich |
| Basic | `"Basic"` | Minimale Angaben |
| Minimum | `"Minimum"` | Nur allernotwendigste Daten |
| XRechnung | `"XRechnung"` | Für öffentliche Auftraggeber (B2G) |

---

## Einheiten-Codes (häufig verwendete)

| Code | Bedeutung |
|------|-----------|
| `C62` | Stück (Piece) |
| `HUR` | Stunde (Hour) |
| `DAY` | Tag (Day) |
| `MON` | Monat (Month) |
| `KGM` | Kilogramm |
| `MTR` | Meter |
| `LTR` | Liter |
| `KMT` | Kilometer |

---

## USt-Kategorien

| Code | Bedeutung | Verwendung |
|------|-----------|------------|
| `S` | Standard | Normaler Steuersatz (19% oder 7%) |
| `Z` | Zero Rate | Nullsatz (0%) |
| `E` | Exempt | Steuerbefreit |
| `AE` | Reverse Charge | Umkehr der Steuerschuld |
| `K` | Innergemeinschaftlich | EU-Lieferungen |

---

## Fehlerbehandlung

```vba
' Nach jedem Aufruf prüfen:
If bridge.GetLastError <> "" Then
    MsgBox "Fehler: " & bridge.GetLastError
End If

' Oder beim Speichern den Rückgabewert prüfen:
Dim err As String
err = bridge.Save(...)
If err <> "" Then MsgBox err
```

---

## Build auf Mac

```bash
# Restore & Build
cd ZUGFeRD
dotnet restore
dotnet build -c Release

# Ausgabe in: bin/Release/net48/
```

> **Hinweis**: Der Build auf Mac erzeugt die DLL, aber das Registrieren (`regasm`) 
> und Testen mit MS Access funktioniert nur auf Windows.

---

## Hinweis zur PDF/A-3 Einbettung

Die aktuelle `SaveWithPdf()`-Methode speichert das ZUGFeRD-XML als separate Datei neben dem PDF.
Für eine vollständige PDF/A-3-konforme Einbettung (XML direkt im PDF) wird zusätzlich eine
PDF-Bibliothek benötigt (z.B. iTextSharp oder PDFsharp). 

Beide Dateien zusammen (PDF + XML) bilden eine gültige ZUGFeRD-Rechnung und können
so an den Empfänger übermittelt werden.

