# Änderungen gegenüber dem Original

Dieses Repository ist eine veränderte Fassung von **ADRecon** (Prashant Mahajan und Mitwirkende).

| | |
|---|---|
| Original | [adrecon/ADRecon](https://github.com/adrecon/ADRecon), Commit `d2ba12c` vom 15.10.2024 (v1.27) |
| Fork | [ichabot/ADReconV2](https://github.com/ichabot/ADReconV2), Version `v1.27-fork.1` |
| Geändert | November 2026 (erste Fassung 09.11.2026) |
| Lizenz | GNU Affero General Public License v3.0, unverändert übernommen, siehe [LICENSE.md](LICENSE.md) |

Dieser Hinweis erfüllt die Kennzeichnungspflicht nach Abschnitt 5a der AGPL-3.0.
Die vollständige Änderungshistorie mit Datum je Änderung steht in der Git-Historie dieses Repositorys.

---

## Geänderte und neue Dateien

| Datei | Status |
|---|---|
| `ADRecon.ps1` | **geändert** (Details unten). Urheber- und Lizenzangaben des Originals sind erhalten, im Kopf steht ein Änderungshinweis. |
| `Invoke-ADAudit.ps1` | **neu**: Startskript für AD-Audits in kleinen und mittleren Umgebungen |
| `README.md` | **neu geschrieben** |
| `ORIGINAL-README.md` | README des Originals, unverändert ausgelagert |
| `CHANGELOG-FORK.md` | **neu**: diese Datei |
| `LICENSE.md` | unverändert |
| `.github/FUNDING.yml` | unverändert (Spendenhinweis des Original-Autors) |
| `.github/workflows/semgrep.yml` | **entfernt** (benötigt ein Token des Original-Projekts und würde täglich fehlschlagen) |

---

## Änderungen in ADRecon.ps1

### Fehlerbehebungen
- **Computers, „Ungültiger Aufzählungskontext“ / „Invalid enumeration context“ (ADWS):**
  `Get-ADComputer` fragte `IPv4Address` ab. Dafür löst das AD-Modul jeden Computer per DNS auf, während die ADWS-Abfrage noch läuft.
  Bei vielen veralteten Computerobjekten lief dadurch der Aufzählungskontext ab.
  Die DNS-Auflösung erfolgt jetzt nach der Abfrage, parallel in den Worker-Threads. Ausgegeben wird nur eine IPv4-Adresse.
- **Thread-Sicherheit:** Die Zuordnungstabellen für Gruppen, GPOs und SIDs wurden von mehreren Threads ohne Sperre befüllt.
  Unter PowerShell 7 führte das zu „Operations that change non-concurrent collections…“ und fehlenden Gruppenmitgliedschaften,
  unter Windows PowerShell 5.1 potenziell zu stillen Fehlern. Die Zugriffe sind jetzt gesperrt.
- **PowerShell 7:** Nach dem Kompilieren lief zusätzlich der Zweig für PowerShell 2.0 (`-Language CSharpVersion3`) und brach ab.
  Die Zweige sind jetzt exklusiv.
- **PowerShell 7 ohne Excel:** Die Suche nach `Microsoft.Office.Interop.Excel.dll` erzeugte Folgefehler.
- **CSV-Kodierung:** `Export-Csv -Encoding Default` schrieb unter PowerShell 5.1 ANSI, der Excel-Import las UTF-8.
  CSV wird jetzt als UTF-8 mit BOM geschrieben.
- **Windows PowerShell 5.1:** Mehrfach-Indizes in Methodenaufrufen (Parser-Fehler unter 5.1) ersetzt.

### Neue Parameter
| Parameter | Wirkung |
|---|---|
| `-PreciseLastLogon` | Liest `lastLogon` auf jedem DC (eine Abfrage pro DC) und nimmt den neuesten Wert. Neue Spalten „Last Logon (all DCs)“ und „Logon Age all DCs (days)“ |
| `-SkipDNSResolve` | Keine DNS-Auflösung der Computer (Spalte IPv4Address bleibt leer) |
| `-CsvDelimiter ';'` | Semikolon als Trennzeichen für deutsches Excel. Excel-Report und `-GenExcel` erkennen das Trennzeichen automatisch |
| `-Logo None` | neuer Wert und Standard: kein Logo im Excel-Report |

### Neue Spalte für Users und Computers
- **„Inactive (> n days)“**: `True`, wenn die letzte bekannte Anmeldung älter als `-DormantTimeSpan` ist,
  oder wenn sich das Konto nie angemeldet hat und älter als `-DormantTimeSpan` ist.
  Die Originalspalten „Dormant“ und „Never Logged in“ bleiben erhalten.

### Geänderte Standardwerte
- `-Collect Default` sammelt **keine LAPS-Passwörter und BitLocker-Recovery-Keys** mehr, weil diese im Klartext im Report landen.
  Bei Bedarf explizit mit `-Collect LAPS,BitLocker`.
- Kein Logo im Excel-Report (`-Logo None`).

### Excel-Report ohne installiertes Excel
- Mit dem Modul **ImportExcel** entsteht `AD-Audit-Report.xlsx`: Übersicht mit Prüfpunkten (Prüfen/OK/Info), auffällige Konten,
  privilegierte Gruppen, inaktive Benutzer und Computer, Statistik mit Diagrammen, Rohdaten als Tabellen mit Hervorhebung.
- Mit installiertem Excel bleibt es beim Original-Report.

### Ausgabe
- Startzeile, About-Blatt und Links im Excel-Report verweisen auf ADReconV2.
- Der Kopf des Skripts enthält weiterhin die Autorenangabe des Originals und einen Änderungshinweis.

---

## Invoke-ADAudit.ps1 (neu)
- Voraussetzung PowerShell 7, unter Windows PowerShell 5.1 Abbruch mit Download-Hinweis
- Lädt `ADRecon.ps1` aus diesem Repository per Commit-ID (kein veralteter Zwischenspeicher), ersetzt ältere Versionen, `-Update` erzwingt Neuladen
- Modulauswahl ohne LAPS, BitLocker, Kerberoast und ACLs
- ADWS nur bei nativ PowerShell-7-fähigem AD-Modul, sonst LDAP
- Abfrage vor dem Lauf, ob ImportExcel installiert werden soll (`-InstallImportExcel Ask|Yes|No`)
- Report-Ordner mit eingeschränkten Rechten, Protokoll, `Zusammenfassung.txt`, `Inaktive-Benutzer.csv`, `Privilegierte-Gruppen.csv`
- Privilegierte Gruppen über SID erkannt (sprachunabhängig)

---

## Test
Getestet am 09.11.2026 mit PowerShell 7.6 auf einem Domain Controller (Methode ADWS):
vollständiger Lauf inklusive `-PreciseLastLogon` und `AD-Audit-Report.xlsx`.
Methode LDAP und Rechner außerhalb der Domäne sind nicht getestet.
