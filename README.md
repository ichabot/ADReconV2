# ADReconV2

AD-Audit für kleine und mittlere Umgebungen auf Basis von [ADRecon](https://github.com/adrecon/ADRecon).

Ein Befehl liest das Active Directory aus und liefert:
- einen **Excel-Report** mit Übersicht, Prüfpunkten und Diagrammen,
- **inaktive Konten** (Standard: 30 Tage), genau über die letzte Anmeldung an allen DCs,
- **Mitglieder privilegierter Gruppen**,
- **auffällige Konten** (Kennwort läuft nie ab, Kennwort nicht erforderlich, AS-REP Roasting, Delegierung usw.),
- alle Rohdaten als CSV.

Das Skript liest nur. Im AD wird nichts geändert.

> Veränderte Fassung von ADRecon v1.27 (Prashant Mahajan und Mitwirkende), Lizenz AGPL-3.0.
> Kein offizielles Projekt der Original-Autoren. Alle Änderungen: [CHANGELOG-FORK.md](CHANGELOG-FORK.md)

---

## Schnellstart

1. **PowerShell 7** (`pwsh`) als Administrator auf einem domänenverbundenen Server oder Client öffnen.
2. Ausführen:

```powershell
& ([scriptblock]::Create((irm https://raw.githubusercontent.com/ichabot/ADReconV2/HEAD/Invoke-ADAudit.ps1)))
```

3. Ist kein Excel installiert, fragt das Skript, ob das Modul **ImportExcel** installiert werden soll. Mit `j` bestätigen, sonst entstehen nur CSV-Dateien.
4. Ergebnis öffnen: `C:\AD-Audit\Report-<Datum-Uhrzeit>\AD-Audit-Report.xlsx`

Der Lauf dauert je nach Umgebung ein bis wenige Minuten.

**Nach einem Update des Repos** einmal mit `-Update` starten, damit die neue `ADRecon.ps1` geladen wird:

```powershell
& ([scriptblock]::Create((irm https://raw.githubusercontent.com/ichabot/ADReconV2/HEAD/Invoke-ADAudit.ps1))) -Update
```

### Weitere Beispiele

```powershell
# Inaktiv ab 60 statt 30 Tagen
& ([scriptblock]::Create((irm https://raw.githubusercontent.com/ichabot/ADReconV2/HEAD/Invoke-ADAudit.ps1))) -DormantDays 60

# Von einem Rechner außerhalb der Domäne mit eigenem Konto
& ([scriptblock]::Create((irm https://raw.githubusercontent.com/ichabot/ADReconV2/HEAD/Invoke-ADAudit.ps1))) -DomainController dc01.firma.local -UseCredential

# Ohne Rückfrage, ImportExcel bei Bedarf installieren
& ([scriptblock]::Create((irm https://raw.githubusercontent.com/ichabot/ADReconV2/HEAD/Invoke-ADAudit.ps1))) -InstallImportExcel Yes
```

### Ohne Internet auf dem Server

1. `Invoke-ADAudit.ps1` und `ADRecon.ps1` aus diesem Repo herunterladen und zusammen in einen Ordner legen (z. B. per USB-Stick).
2. ImportExcel bei Bedarf vorher auf einem Rechner mit Internet holen: `Save-Module ImportExcel -Path <Ordner>` und auf dem Server nach `Documents\PowerShell\Modules` kopieren.
3. In `pwsh` ausführen:

```powershell
pwsh -ExecutionPolicy Bypass -File .\Invoke-ADAudit.ps1
```

---

## Was das Skript macht

1. Prüft, ob PowerShell 7 läuft. Unter Windows PowerShell 5.1 bricht es mit Download-Hinweis ab.
2. Fragt bei fehlendem Excel, ob ImportExcel installiert werden soll.
3. Legt `C:\AD-Audit\Report-<Datum-Uhrzeit>` an, lesbar nur für Administratoren, SYSTEM und den ausführenden Benutzer.
4. Lädt `ADRecon.ps1` aus diesem Repo nach `C:\AD-Audit\ADRecon\` (oder nutzt die vorhandene Fork-Version).
5. Liest das AD aus: über ADWS, wenn das AD-Modul unter PowerShell 7 läuft, sonst über LDAP.
6. Erstellt CSV-Dateien, den Excel-Report, `Zusammenfassung.txt` und ein Protokoll.

---

## Parameter

| Parameter | Standard | Bedeutung |
|---|---|---|
| `-DormantDays` | `30` | Ab wie vielen Tagen ohne Anmeldung ein Konto als inaktiv gilt |
| `-BaseDir` | `C:\AD-Audit` | Arbeitsordner für ADRecon und die Reports |
| `-DomainController` | – | DC (FQDN), nötig bei `-UseCredential` |
| `-UseCredential` | aus | Fragt ein Domänenkonto ab (für Rechner außerhalb der Domäne) |
| `-PreciseLastLogon` | `$true` | Letzte Anmeldung von jedem DC lesen (genau). Mit `$false` nur `lastLogonTimestamp` (bis ca. 14 Tage ungenau) |
| `-SkipDNSResolve` | aus | Keine DNS-Auflösung der Computer, schnellerer Lauf |
| `-Update` | aus | `ADRecon.ps1` neu aus dem Repo laden |
| `-InstallImportExcel` | `Ask` | Ohne Excel: `Ask` fragt, `Yes` installiert ohne Rückfrage, `No` erstellt nur CSV |

---

## Ergebnis

```
C:\AD-Audit\Report-20261109-1530\
├── AD-Audit-Report.xlsx         Excel-Report (ohne installiertes Excel, über ImportExcel)
├── <Domäne>ADRecon-Report.xlsx  Original-Report von ADRecon (nur mit installiertem Excel)
├── Zusammenfassung.txt          Benutzerzahlen, inaktive Konten, privilegierte Gruppen
├── Inaktive-Benutzer.csv        aktivierte Benutzer ohne Anmeldung seit n Tagen bzw. nie angemeldet
├── Privilegierte-Gruppen.csv    Mitglieder von Domänen-Admins, Administratoren, Schema-/Organisations-Admins usw.
├── Audit-Protokoll.txt          komplette Konsolenausgabe
└── CSV-Files\                   Rohdaten je Bereich (Users, Groups, GroupMembers, Computers, ...)
```

Alle CSV-Dateien nutzen Semikolon und UTF-8 und lassen sich in einem deutschen Excel per Doppelklick öffnen.

### AD-Audit-Report.xlsx

| Blatt | Inhalt |
|---|---|
| **Übersicht** | Domäne, Funktionsebenen, rund 30 Prüfpunkte mit Status *Prüfen / OK / Info* und kurzem Hinweis |
| **Auffällige Konten** | ein Eintrag je Befund, als Abarbeitungsliste |
| **Privilegierte Gruppen** | Mitglieder mit Status (aktiviert, inaktiv, Kennwort läuft nie ab, Kennwortalter) |
| **Inaktive Benutzer / Computer** | aktivierte Konten ohne Anmeldung seit n Tagen |
| **Statistik** | Benutzerstatus und Betriebssysteme mit Diagramm |
| Rohdaten | alle CSV-Dateien als Tabellen mit Filter, kritische Werte rot markiert |

**Geprüft wird u. a.:** inaktive und nie angemeldete Konten, Kennwort läuft nie ab, Kennwort älter als die Richtlinie,
Kennwort nicht erforderlich, Kerberos-Vorauthentifizierung aus (AS-REP Roasting), privilegierte Konten mit SPN (Kerberoasting),
uneingeschränkte Delegierung, Betriebssysteme ohne Herstellersupport, Anzahl DCs, SMB1 und SMB-Signierung auf DCs,
Kennwortrichtlinie (Mindestlänge, Komplexität, Kontosperrung) und privilegierte Gruppen.

Privilegierte Gruppen werden über die SID erkannt und funktionieren daher auch mit deutschen Gruppennamen.
Aufgeführt werden direkte Mitglieder, Gruppen als Mitglied erscheinen als `[Gruppe] <Name>`.

**Wichtige Spalten in `Users.csv` und `Computers.csv`:**

| Spalte | Bedeutung |
|---|---|
| `Inactive (> 30 days)` | `True` = keine Anmeldung seit 30 Tagen, oder nie angemeldet und älter als 30 Tage |
| `Last Logon (all DCs)` | letzte Anmeldung, neuester Wert aller DCs |
| `Never Logged in` | hat sich nie angemeldet |
| `Enabled` | Konto aktiviert |

---

## Was wird erfasst

Forest, Domäne, Trusts, Sites, Subnetze, Kennwortrichtlinien (inkl. Fine Grained), Domain Controller,
Benutzer, Benutzer-SPNs, Gruppen, Gruppenmitglieder, Gruppenänderungen, OUs, GPOs, GPO-Verknüpfungen, Computer.

**Bewusst nicht erfasst:** LAPS-Passwörter, BitLocker-Recovery-Keys, Kerberoast, ACLs.
Bei Bedarf direkt mit `ADRecon.ps1 -Collect ...` abfragen.

---

## Voraussetzungen

| | |
|---|---|
| **PowerShell** | **PowerShell 7** (`pwsh`). Unter Windows PowerShell 5.1 bricht das Skript mit Hinweis und Download-Link ab |
| **Rechner** | Windows Server oder Client mit PowerShell 7, domänenverbunden (sonst `-DomainController` und `-UseCredential`) |
| **AD-Modul** | optional. Läuft es nativ unter PowerShell 7 (RSAT ab Server 2019 / Windows 10 1809), wird ADWS genutzt, sonst LDAP |
| **Rechte** | normaler Domänenbenutzer reicht. Fine Grained Password Policies sind nur mit erweiterten Rechten lesbar |
| **Excel-Report** | ohne installiertes Excel das Modul ImportExcel (das Skript fragt vorab) |
| **Internet** | Zugriff auf `raw.githubusercontent.com`, `api.github.com` und die PowerShell Gallery (nur für ImportExcel) |

**PowerShell 7 installieren:**
- Download (MSI `PowerShell-7.x.x-win-x64.msi`): https://github.com/PowerShell/PowerShell/releases/latest
- Anleitung: https://learn.microsoft.com/powershell/scripting/install/installing-powershell-on-windows
- oder: `winget install --id Microsoft.PowerShell --source winget`

---

## Bekannte Probleme

**Virenscanner meldet ADRecon:** Sophos, Defender und andere erkennen ADRecon häufig als Hacktool, weil es auch von Angreifern genutzt wird.
Vor dem Einsatz beim Kunden abstimmen und bei Bedarf eine Ausnahme setzen, nach dem Audit wieder entfernen:
- AMSI-Erkennung: Ausnahme auf den Pfad `C:\AD-Audit\ADRecon\ADRecon.ps1`
- PUA-Erkennung: Ausnahme mit dem im Event angezeigten Namen

**pwsh 7.6.6/7.6.7 stürzt beim Start ab** (`totalWidth ('-1') must be a non-negative value`, `ShowUpdateNotification`):
Fehler in PowerShell selbst, behoben in 7.6.8. Bis zum Update in `cmd` ausführen und alle Konsolen neu öffnen:

```
setx POWERSHELL_UPDATECHECK Off
```

**Neue Version wird nicht geladen:** Mit `-Update` starten. Ohne `-Update` nutzt das Skript eine bereits vorhandene Fork-Version.

---

## Sicherheit

- Der Report enthält eine vollständige Übersicht über Benutzer, Gruppen, Computer und Richtlinien.
- Der Report-Ordner wird beim Anlegen auf Administratoren, SYSTEM und den ausführenden Benutzer beschränkt.
- Report nach dem Abholen vom Server löschen.

---

## ADRecon direkt verwenden

`ADRecon.ps1` lässt sich wie das Original nutzen, mit diesen zusätzlichen Parametern:

```powershell
.\ADRecon.ps1 -DormantTimeSpan 30 -PreciseLastLogon -CsvDelimiter ';'
.\ADRecon.ps1 -Collect Users,Groups,GroupMembers -OutputType CSV -CsvDelimiter ';'
.\ADRecon.ps1 -SkipDNSResolve
# Excel-Report aus vorhandenen CSV-Dateien neu erzeugen
.\ADRecon.ps1 -GenExcel C:\AD-Audit\Report-20261109-1530
```

Alle Parameter des Originals: [ORIGINAL-README.md](ORIGINAL-README.md)

---

## Lizenz und Herkunft

- **Original:** [ADRecon](https://github.com/adrecon/ADRecon) v1.27 von Prashant Mahajan und Mitwirkenden, ursprünglich bei Sense of Security.
- **Lizenz:** GNU Affero General Public License v3.0, siehe [LICENSE.md](LICENSE.md). Dieser Fork steht unter derselben Lizenz.
- **Änderungen:** `ADRecon.ps1` wurde verändert, `Invoke-ADAudit.ps1` ist neu. Was geändert wurde, steht in [CHANGELOG-FORK.md](CHANGELOG-FORK.md)
  und im Kopf von `ADRecon.ps1`, wann es geändert wurde, in der Git-Historie.
- **Quellcode:** vollständig in diesem Repository.
