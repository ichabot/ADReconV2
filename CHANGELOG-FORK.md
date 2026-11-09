# Änderungen in diesem Fork

Basis: [adrecon/ADRecon](https://github.com/adrecon/ADRecon), Commit `d2ba12c` (15.10.2024, v1.27).
Lizenz unverändert AGPL-3.0, siehe `LICENSE.md`.

## v1.27-fork.1 (2026-11)

### Fehlerbehebung
- **Computers: „Ungültiger Aufzählungskontext“ / „Invalid enumeration context“ (ADWS)**
  `Get-ADComputer` hat die Eigenschaft `IPv4Address` abgefragt. Dafür löst das AD-Modul jeden Computer einzeln per DNS auf,
  noch während die ADWS-Abfrage läuft. Bei vielen veralteten Computerobjekten läuft dadurch der Aufzählungskontext ab.
  Die IPv4-Auflösung passiert jetzt erst nach der Abfrage, parallel in den Worker-Threads (wie bisher schon im LDAP-Modus).
  Ausgegeben wird nur noch eine IPv4-Adresse (vorher im LDAP-Modus teilweise IPv6).
- **CSV-Kodierung**: `Export-Csv -Encoding Default` schrieb unter PowerShell 5.1 ANSI, der Excel-Import liest aber UTF-8.
  Umlaute waren dadurch im Report kaputt. CSV wird jetzt immer als UTF-8 mit BOM geschrieben.
- **PowerShell 7 ohne Excel**: Die Suche nach `Microsoft.Office.Interop.Excel.dll` lief ins Leere und erzeugte Folgefehler.

### Neue Parameter
| Parameter | Wirkung |
|---|---|
| `-SkipDNSResolve` | Keine DNS-Auflösung der Computer (Spalte IPv4Address bleibt leer). Schnellster Lauf. |
| `-PreciseLastLogon` | Liest `lastLogon` auf jedem DC (eine Abfrage pro DC) und nimmt den neuesten Wert. Neue Spalten „Last Logon (all DCs)“ und „Logon Age all DCs (days)“. |
| `-CsvDelimiter ';'` | Semikolon als Trennzeichen, damit ein deutsches Excel die CSV per Doppelklick richtig öffnet. Excel-Report und `-GenExcel` erkennen das Trennzeichen automatisch. |

### Neue Spalte für Users und Computers
- **„Inactive (> n days)“**: `True`, wenn die letzte bekannte Anmeldung älter als `-DormantTimeSpan` ist,
  oder wenn sich das Konto nie angemeldet hat und älter als `-DormantTimeSpan` ist.
  Mit `-PreciseLastLogon` wird der genaue Wert aller DCs verwendet, sonst `lastLogonTimestamp` (kann bis ca. 14 Tage hinterherhinken).
  Die Originalspalten „Dormant“ und „Never Logged in“ bleiben unverändert erhalten.

### Geänderte Standardwerte
- Die Standardauswahl (`-Collect Default`) sammelt **keine LAPS-Passwörter und BitLocker-Recovery-Keys** mehr,
  weil diese im Klartext in CSV/Excel landen. Bei Bedarf explizit mit `-Collect LAPS,BitLocker` abfragen.

### Excel ohne installiertes Excel
- Ist kein Excel installiert, aber das Modul **ImportExcel** vorhanden, entsteht `ADRecon-Report-Simple.xlsx`
  (ein Blatt pro CSV, Filter, fixierte Kopfzeile). Ohne Pivots und Diagramme.
  Modul installieren: `Install-Module ImportExcel -Scope CurrentUser`

### Repo
- Semgrep-Workflow entfernt (benötigt ein Token des Original-Projekts und würde täglich fehlschlagen).
- `Invoke-ADAudit.ps1`: Wrapper für KMU-Audits (siehe README).
- README neu geschrieben (Schnellstart, Parameter, Voraussetzungen), Original-README in `ORIGINAL-README.md`.

### Ausgabe
- Konsolen-Banner, About-Blatt und Links im Excel-Report verweisen auf ADReconV2. Logo im Excel-Report standardmäßig aus (`-Logo None`).
  Urheber- und Lizenzhinweise im Quelltext bleiben erhalten (AGPL-3.0).
- `Invoke-ADAudit.ps1` fragt ohne installiertes Excel vor dem Lauf, ob ImportExcel installiert werden soll (`-InstallImportExcel Ask|Yes|No`).

### Nicht getestet
Die Änderungen sind per PowerShell-Parser und C#-Compiler geprüft, aber noch nicht gegen ein echtes AD gelaufen.
