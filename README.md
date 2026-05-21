# Proofio Outlook Add-in

Ein Microsoft Outlook VSTO Add-in zur Integration mit [Proofio](https://gutachtenpilot.lovable.app) – dem Gutachten- und Akten-Management-System.

## Features

- 📁 **Mails ablegen** – E-Mails direkt aus Outlook in einer Proofio-Akte ablegen
- 📅 **Termine ablegen** – Kalendertermine einer Akte zuordnen
- ✉️ **Senden & Ablegen** – Ausgehende Mails beim Senden automatisch ablegen
- 🟢 **Grünes Badge** – Abgelegte Mails erhalten automatisch die Outlook-Kategorie „In Proofio abgelegt"
- 🗂️ **Fall anlegen** – Neue Proofio-Fälle direkt aus Outlook erstellen

## Voraussetzungen

| Voraussetzung | Version |
|---|---|
| Windows | 10 / 11 |
| Microsoft Outlook | 2016 / 2019 / 2021 / Microsoft 365 |
| .NET Framework | 4.7.2 |
| Visual Studio Tools for Office Runtime | 4.0 |
| Visual Studio (zum Bauen) | 2019 / 2022 |

## Installation (Entwickler)

```bash
git clone https://github.com/DEIN-USERNAME/proofio-outlook-addin.git
cd proofio-outlook-addin
```

1. `ProofioAddin.sln` in Visual Studio öffnen
2. Konfiguration auf **x86 | Debug** stellen
3. `Strg+F5` – Visual Studio startet Outlook automatisch mit dem Add-in

## Konfiguration

Beim ersten Start in Outlook:

**Proofio-Ribbon → Einstellungen**

| Feld | Wert |
|---|---|
| Bearer-Token | API-Token aus Proofio kopieren |
| API-Basis-URL | `https://gutachtenpilot.lovable.app/api/public/v1` |
| Beim Senden | Nach Wunsch einstellen |

## Zertifikat einrichten (für Distribution)

Nach dem Klonen **kein** `.pfx` im Repo vorhanden. Eigenes Zertifikat hinterlegen:

1. Visual Studio → Rechtsklick Projekt → **Eigenschaften** → Reiter **Signierung**
2. **Aus Datei auswählen...** → eigene `.pfx` wählen
3. Oder: **Testzertifikat erstellen** für interne Verteilung

## Setup erstellen (ClickOnce)

1. Konfiguration auf **x86 | Release** stellen
2. Rechtsklick Projekt → **Veröffentlichen**
3. Ausgabeordner wählen → **Veröffentlichen**
4. `setup.exe` an Nutzer weitergeben

## Projektstruktur

```
ProofioAddin/
├── Api/
│   ├── Models.cs               # Request/Response DTOs
│   └── ProofioApiClient.cs     # HTTP-Client für die Proofio API
├── Ribbons/
│   ├── ExplorerRibbon.cs       # Ribbon im Posteingang
│   ├── MailReadRibbon.cs       # Ribbon beim Lesen einer Mail
│   ├── MailComposeRibbon.cs    # Ribbon beim Verfassen
│   ├── AppointmentRibbon.cs    # Ribbon bei Terminen
│   └── ProofioRibbonManager.cs
├── Services/
│   ├── MailCategoryService.cs  # Grünes Badge nach dem Ablegen
│   ├── MailExtractor.cs        # Extrahiert Mail-Daten für die API
│   ├── AppointmentExtractor.cs # Extrahiert Termin-Daten
│   ├── PendingSendTracker.cs   # Verfolgt "Nach Senden ablegen"
│   ├── TokenStore.cs           # Speichert API-Token sicher
│   └── Logger.cs
├── UI/
│   ├── UITheme.cs              # Design-Tokens, geteilte Controls
│   ├── CasePickerDialog.cs     # Akte auswählen
│   ├── NewCaseDialog.cs        # Neuen Fall anlegen
│   └── TokenDialog.cs          # Einstellungen
└── ThisAddIn.cs                # Add-in Einstiegspunkt
```

## API-Endpunkte

| Endpunkt | Beschreibung |
|---|---|
| `POST /outlook/case` | Neuen Fall anlegen |
| `GET /cases/search?q=` | Akten suchen |
| `POST /outlook/email` | Mail ablegen |
| `POST /outlook/appointment` | Termin ablegen |

## Lizenz

MIT – siehe [LICENSE](LICENSE)
