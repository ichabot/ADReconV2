# ADReconV2

AD-Audit für kleine und mittlere Umgebungen auf Basis von [ADRecon](https://github.com/adrecon/ADRecon).
Erfasst Benutzer, Gruppen, Gruppenmitgliedschaften, Computer, OUs, GPOs und Kennwortrichtlinien,
markiert inaktive Konten und fasst die wichtigsten Punkte in einer Kurzübersicht zusammen.

> Fork von adrecon/ADRecon v1.27 (Commit `d2ba12c`), Lizenz AGPL-3.0.
> Kein offizielles Projekt, kein Support durch die Original-Autoren. Änderungen: [CHANGELOG-FORK.md](CHANGELOG-FORK.md)

---

## Schnellstart

**PowerShell 7** (`pwsh`) auf einem domänenverbundenen Server oder Client öffnen und ausführen:

```powershell
& ([scriptblock]::Create((irm https://raw.githubusercontent.com/ichabot/ADReconV2/HEAD/Invoke-ADAudit.ps1)))
```

Ist auf dem Rechner kein Excel installiert, fragt das Skript vorab, ob das Modul ImportExcel installiert werden soll (für den Excel-Report).
Der Lauf dauert je nach Umgebung wenige Sekunden bis einige Minuten.
Das Ergebnis liegt danach unter `C:\AD-Audit\Report-<Datum-Uhrzeit>\`.

### Mit Parametern

Parameter werden einfach hinten angehängt:

```powershell
# Inaktiv ab 60 statt 30 Tagen
& ([scriptblock]::Create((irm https://raw.githubusercontent.com/ichabot/ADReconV2/HEAD/Invoke-ADAudit.ps1))) -DormantDays 60

# Von einem Rechner außerhalb der Domäne mit eigenem Konto
& ([scriptblock]::Create((irm https://raw.githubusercontent.com/ichabot/ADReconV2/HEAD/Invoke-ADAudit.ps1))) -DomainController dc01.firma.local -UseCredential

# ADRecon.ps1 neu aus dem Repo laden (nach Updates)
& ([scriptblock]::Create((irm https://raw.githubusercontent.com/ichabot/ADReconV2/HEAD/Invoke-ADAudit.ps1))) -Update
```

### Ohne Internet auf dem Server

1. `Invoke-ADAudit.ps1` und `ADRecon.ps1` aus diesem Repo herunterladen und zusammen in einen Ordner legen (z. B. per USB-Stick).
2. Auf dem Server ausführen:

```powershell
pwsh -ExecutionPolicy Bypass -File .\Invoke-ADAudit.ps1
```

---

## Parameter von Invoke-ADAudit.ps1

| Parameter | Standard | Bedeutung |
|---|---|---|
| `-DormantDays` | `30` | Ab wie vielen Tagen ohne Anmeldung ein Konto als inaktiv gilt |
| `-BaseDir` | `C:\AD-Audit` | Arbeitsordner für ADRecon und die Reports |
| `-DomainController` | – | DC (FQDN), nötig bei `-UseCredential` |
| `-UseCredential` | aus | Fragt ein Domänenkonto ab (für Rechner außerhalb der Domäne) |
| `-PreciseLastLogon` | `$true` | Liest die letzte Anmeldung von jedem DC (genau). `-PreciseLastLogon $false` nutzt nur `lastLogonTimestamp` (bis ca. 14 Tage ungenau) |
| `-SkipDNSResolve` | aus | Keine DNS-Auflösung der Computer, schnellerer Lauf |
| `-Update` | aus | ADRecon.ps1 neu herunterladen |
| `-InstallImportExcel` | `Ask` | Ohne installiertes Excel: vor dem Lauf fragen, ob das Modul ImportExcel installiert wird. `Yes` installiert ohne Rückfrage, `No` erstellt nur CSV |

---

## Ergebnis

```
C:\AD-Audit\Report-20261109-1530\
├── Zusammenfassung.txt          Kurzübersicht (Benutzerzahlen, inaktive Konten, privilegierte Gruppen)
├── Inaktive-Benutzer.csv        aktivierte Benutzer ohne Anmeldung seit n Tagen bzw. nie angemeldet
├── Privilegierte-Gruppen.csv    Mitglieder von Domänen-Admins, Administratoren, Schema-/Organisations-Admins usw.
├── Audit-Protokoll.txt          komplette Konsolenausgabe
├── <Domäne>ADRecon-Report.xlsx  Excel-Report (wenn Excel installiert ist)
├── AD-Audit-Report.xlsx         Audit-Report ohne installiertes Excel (Modul ImportExcel), siehe unten
└── CSV-Files\                   Rohdaten je Modul (Users, Groups, GroupMembers, Computers, ...)
```

Alle CSV-Dateien nutzen Semikolon und UTF-8, sie lassen sich in einem deutschen Excel per Doppelklick öffnen.

### AD-Audit-Report.xlsx

Wird erstellt, wenn auf dem Rechner kein Excel installiert ist (über das Modul ImportExcel):

| Blatt | Inhalt |
|---|---|
| **Übersicht** | Domäne, Funktionsebenen, Prüfpunkte mit Status *Prüfen / OK / Info* und kurzem Hinweis |
| **Auffällige Konten** | ein Eintrag je Befund, z. B. Kennwort läuft nie ab, Kennwort nicht erforderlich, AS-REP Roasting, uneingeschränkte Delegierung |
| **Privilegierte Gruppen** | Mitglieder mit Status (aktiviert, inaktiv, Kennwort läuft nie ab, Kennwortalter) |
| **Inaktive Benutzer / Computer** | aktivierte Konten ohne Anmeldung seit n Tagen |
| **Statistik** | Benutzerstatus und Betriebssysteme mit Diagramm |
| Rohdaten | alle CSV-Dateien als Tabellen mit Filter, kritische Werte rot markiert |

Mit installiertem Excel entsteht stattdessen der Original-Report von ADRecon mit Pivot-Tabellen.

**Wichtige Spalten in `Users.csv` und `Computers.csv`:**

| Spalte | Bedeutung |
|---|---|
| `Inactive (> 30 days)` | `True` = keine Anmeldung seit 30 Tagen, oder nie angemeldet und älter als 30 Tage |
| `Last Logon (all DCs)` | letzte Anmeldung, neuester Wert aller DCs |
| `Never Logged in` | hat sich nie angemeldet |
| `Enabled` | Konto aktiviert |

Privilegierte Gruppen werden über die SID erkannt und funktionieren daher auch mit deutschen Gruppennamen.
Aufgeführt werden direkte Mitglieder; enthaltene Gruppen erscheinen als eigener Eintrag.

---

## Was wird erfasst

Forest, Domäne, Trusts, Sites, Subnetze, Kennwortrichtlinien (inkl. Fine Grained), Domain Controller,
Benutzer, Benutzer-SPNs, Gruppen, Gruppenmitglieder, Gruppenänderungen, OUs, GPOs, GPO-Verknüpfungen, Computer.

**Bewusst nicht erfasst:** LAPS-Passwörter, BitLocker-Recovery-Keys, Kerberoast, ACLs.
Diese können bei Bedarf direkt mit `ADRecon.ps1 -Collect ...` abgefragt werden.

---

## Voraussetzungen

| | |
|---|---|
| **PowerShell** | **PowerShell 7** (`pwsh`). Unter Windows PowerShell 5.1 bricht `Invoke-ADAudit.ps1` mit Hinweis und Download-Link ab |
| **Betriebssystem** | Windows Server oder Windows-Client, auf dem PowerShell 7 läuft, domänenverbunden (sonst `-DomainController` und `-UseCredential`) |
| **AD-Modul** | Läuft das AD-Modul nativ unter PowerShell 7 (RSAT ab Server 2019 / Windows 10 1809), nutzt ADRecon ADWS, sonst automatisch LDAP. RSAT ist nicht zwingend |
| **Domain Controller** | alle DCs mit AD-Webdiensten (ab Server 2008 R2), bei LDAP keine Einschränkung |
| **Rechte** | normaler Domänenbenutzer reicht. Fine Grained Password Policies sind nur mit erweiterten Rechten lesbar |
| **Excel-Report** | Microsoft Excel auf dem Rechner, sonst Modul ImportExcel: `Install-Module ImportExcel -Scope CurrentUser` (fragt das Skript vorab) |
| **Download** | Zugriff auf `raw.githubusercontent.com` (TLS 1.2 wird im Skript aktiviert) |

PowerShell 7 installieren:

- Download (MSI `PowerShell-7.x.x-win-x64.msi`): https://github.com/PowerShell/PowerShell/releases/latest
- Anleitung: https://learn.microsoft.com/powershell/scripting/install/installing-powershell-on-windows
- oder: `winget install --id Microsoft.PowerShell --source winget`

RSAT nachinstallieren (optional):

```powershell
# Server
Install-WindowsFeature RSAT-AD-PowerShell
# Windows 10/11
Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
```

---

## Virenscanner

Sophos, Defender und andere Scanner melden ADRecon häufig als Hacktool, weil es auch von Angreifern zur Aufklärung genutzt wird.
Vor dem Einsatz beim Kunden abstimmen und bei Bedarf eine Ausnahme setzen:

- **AMSI-Erkennung:** Ausnahme auf den Pfad `C:\AD-Audit\ADRecon\ADRecon.ps1`
- **PUA-Erkennung:** Ausnahme mit dem im Event angezeigten Namen

Ausnahmen nach dem Audit wieder entfernen.

---

## Sicherheit

- Der Report enthält eine vollständige Übersicht über Benutzer, Gruppen, Computer und Richtlinien.
- Der Report-Ordner wird beim Anlegen auf Administratoren, SYSTEM und den ausführenden Benutzer beschränkt.
- Report nach dem Abholen vom Server löschen.
- ADRecon liest nur. Es werden keine Änderungen im AD vorgenommen.

---

## ADRecon direkt verwenden

`ADRecon.ps1` lässt sich wie das Original nutzen, mit diesen zusätzlichen Parametern:

```powershell
.\ADRecon.ps1 -DormantTimeSpan 30 -PreciseLastLogon -CsvDelimiter ';'
.\ADRecon.ps1 -Collect Users,Groups,GroupMembers -OutputType CSV -CsvDelimiter ';'
.\ADRecon.ps1 -SkipDNSResolve
```

Alle Parameter des Originals: [ORIGINAL-README.md](ORIGINAL-README.md)

---

## Lizenz und Herkunft

ADRecon © Prashant Mahajan und Mitwirkende, ursprünglich bei Sense of Security, Lizenz [AGPL-3.0](LICENSE.md).
Dieser Fork steht unter derselben Lizenz. Änderungen sind in [CHANGELOG-FORK.md](CHANGELOG-FORK.md) und im Kopf von `ADRecon.ps1` dokumentiert.
