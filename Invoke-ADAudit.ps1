<#
.SYNOPSIS
    AD-Audit fuer KMU mit ADRecon (Fork ichabot/ADReconV2).

.DESCRIPTION
    Startet ADRecon mit einer reduzierten Modulauswahl (Domaene, Benutzer, Gruppen, Mitgliedschaften,
    OUs, GPOs, Computer), ohne LAPS-Passwoerter, BitLocker-Keys, Kerberoast und ACLs.
    - Inaktivitaet ab -DormantDays (Standard 30 Tage), standardmaessig exakt ueber lastLogon aller DCs
    - CSV mit Semikolon (deutsches Excel)
    - Report-Ordner nur fuer Administratoren, SYSTEM und den ausfuehrenden Benutzer lesbar
    - Protokoll (Transcript) und Kurzzusammenfassung im Report-Ordner

    ADRecon.ps1 wird in dieser Reihenfolge gesucht:
      1. im Ordner dieses Skripts
      2. <BaseDir>\ADRecon\ADRecon.ps1
      3. Download aus github.com/ichabot/ADReconV2 nach <BaseDir>\ADRecon\

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Invoke-ADAudit.ps1

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Invoke-ADAudit.ps1 -DormantDays 60 -BaseDir D:\Audit

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Invoke-ADAudit.ps1 -DomainController dc01.firma.local -UseCredential

.EXAMPLE
    # ADRecon.ps1 neu aus dem Repo laden, auch wenn schon eine Fork-Version vorhanden ist
    powershell -ExecutionPolicy Bypass -File .\Invoke-ADAudit.ps1 -Update

.EXAMPLE
    # ohne Rueckfrage: ImportExcel bei Bedarf installieren (Yes) oder nie (No)
    powershell -ExecutionPolicy Bypass -File .\Invoke-ADAudit.ps1 -InstallImportExcel Yes

.NOTES
    Sophos/Defender erkennen ADRecon haeufig als Hacktool. Ausnahme vorher abstimmen
    (AMSI-Ausnahme auf den Pfad von ADRecon.ps1 bzw. PUA-Ausnahme mit dem gemeldeten Namen).
#>
param(
    [int] $DormantDays = 30,
    [string] $BaseDir = 'C:\AD-Audit',
    [string] $DomainController,
    [switch] $UseCredential,
    [bool] $PreciseLastLogon = $true,
    [switch] $SkipDNSResolve,
    [switch] $Update,
    [ValidateSet('Ask', 'Yes', 'No')]
    [string] $InstallImportExcel = 'Ask'
)

$ErrorActionPreference = 'Stop'
$ForkUrl = 'https://raw.githubusercontent.com/ichabot/ADReconV2/HEAD/ADRecon.ps1'

If ($UseCredential -and -not $DomainController)
{
    throw 'Mit -UseCredential bitte auch -DomainController <FQDN> angeben.'
}

# Excel-Report: ohne installiertes Excel wird das Modul ImportExcel benoetigt (Abfrage vor dem Lauf)
$hasExcel       = Test-Path 'Registry::HKEY_CLASSES_ROOT\Excel.Application'
$hasImportExcel = [bool](Get-Module -ListAvailable -Name ImportExcel)
If (-not $hasExcel -and -not $hasImportExcel -and $InstallImportExcel -ne 'No')
{
    If ($PSVersionTable.PSVersion.Major -lt 5)
    {
        Write-Host '[!] Kein Excel und kein Modul ImportExcel. Installation erst ab PowerShell 5.1 moeglich, es werden nur CSV-Dateien erstellt.'
    }
    Else
    {
        $answer = 'j'
        If ($InstallImportExcel -eq 'Ask')
        {
            Write-Host ''
            Write-Host 'Auf diesem Rechner ist kein Microsoft Excel installiert.'
            Write-Host 'Fuer einen Excel-Report (ein Tabellenblatt je Bereich, mit Filter und fixierter Kopfzeile)'
            Write-Host 'kann das PowerShell-Modul ImportExcel aus der PowerShell Gallery installiert werden.'
            Write-Host 'Es wird nur fuer den aktuellen Benutzer installiert, Excel selbst wird nicht benoetigt.'
            Write-Host 'Ohne das Modul entstehen nur CSV-Dateien.'
            $answer = Read-Host 'ImportExcel jetzt installieren? [j/N]'
        }
        If ($answer -match '^(j|ja|y|yes)$')
        {
            Try
            {
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                If (-not (Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction SilentlyContinue | Where-Object { $_.Version -ge [version]'2.8.5.201' }))
                {
                    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force | Out-Null
                }
                Install-Module -Name ImportExcel -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -Confirm:$false
                Write-Host '[+] ImportExcel installiert.'
            }
            Catch
            {
                Write-Warning "ImportExcel konnte nicht installiert werden: $($_.Exception.Message)"
                Write-Host '[!] Es werden nur CSV-Dateien erstellt.'
            }
        }
        Else
        {
            Write-Host '[*] Ohne ImportExcel, es werden nur CSV-Dateien erstellt.'
        }
        Write-Host ''
    }
    $hasImportExcel = [bool](Get-Module -ListAvailable -Name ImportExcel)
}

$stamp  = Get-Date -Format 'yyyyMMdd-HHmm'
$outDir = Join-Path $BaseDir "Report-$stamp"
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

# Report-Ordner absichern: Vererbung aus, nur Administratoren, SYSTEM und aktueller Benutzer
$me = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
& icacls.exe $outDir /inheritance:r /grant:r "*S-1-5-32-544:(OI)(CI)F" "*S-1-5-18:(OI)(CI)F" "*$($me):(OI)(CI)F" | Out-Null
If ($LASTEXITCODE -ne 0)
{
    Write-Warning "Rechte auf $outDir konnten nicht gesetzt werden (icacls Exitcode $LASTEXITCODE)."
}

Start-Transcript -Path (Join-Path $outDir 'Audit-Protokoll.txt') | Out-Null
Try
{
    # ADRecon.ps1 finden oder laden
    # nur die Fork-Version verwenden (Original kennt -CsvDelimiter, -PreciseLastLogon usw. nicht)
    $tool = $null
    $candidates = @()
    If ($PSScriptRoot) { $candidates += Join-Path $PSScriptRoot 'ADRecon.ps1' }
    $candidates += Join-Path $BaseDir 'ADRecon\ADRecon.ps1'
    If (-not $Update)
    {
        ForEach ($candidate in $candidates)
        {
            If ((Test-Path $candidate) -and (Select-String -Path $candidate -Pattern 'ADReconVersion = "v1.27-fork' -SimpleMatch -Quiet))
            {
                $tool = $candidate
                Break
            }
            ElseIf (Test-Path $candidate)
            {
                Write-Host "[!] $candidate ist nicht die Fork-Version und wird ignoriert."
            }
        }
    }
    If (-not $tool)
    {
        $toolDir = Join-Path $BaseDir 'ADRecon'
        New-Item -ItemType Directory -Path $toolDir -Force | Out-Null
        $tool = Join-Path $toolDir 'ADRecon.ps1'
        Write-Host "[*] Lade ADRecon aus $ForkUrl"
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $ForkUrl -OutFile $tool -UseBasicParsing
    }
    Unblock-File -Path $tool -ErrorAction SilentlyContinue
    $hash = (Get-FileHash -Path $tool -Algorithm SHA256).Hash
    Write-Host "[*] ADRecon: $tool"
    Write-Host "[*] SHA256:  $hash"

    $method = If (Get-Module -ListAvailable -Name ActiveDirectory) { 'ADWS' } Else { 'LDAP' }

    $params = @{
        Method           = $method
        OutputDir        = $outDir
        DormantTimeSpan  = $DormantDays
        OutputType       = $(If ($hasExcel -or $hasImportExcel) { @('CSV', 'Excel') } Else { @('CSV') })
        CsvDelimiter     = ';'
        Collect          = @('Forest','Domain','Trusts','Sites','Subnets','PasswordPolicy','FineGrainedPasswordPolicy',
                             'DomainControllers','Users','UserSPNs','Groups','GroupMembers','GroupChanges','OUs',
                             'GPOs','gPLinks','Computers')
    }
    If ($PreciseLastLogon) { $params.PreciseLastLogon = $true }
    If ($SkipDNSResolve)   { $params.SkipDNSResolve = $true }
    If ($DomainController) { $params.DomainController = $DomainController }
    If ($UseCredential)    { $params.Credential = Get-Credential -Message 'Domaenenkonto (DOMAENE\user)' }

    Write-Host "[*] Starte ADRecon (Methode: $method, inaktiv ab $DormantDays Tagen, lastLogon aller DCs: $PreciseLastLogon)"
    # ADRecon arbeitet mit Warnungen und nicht-abbrechenden Fehlern; 'Stop' wuerde den Lauf beim ersten Problem beenden
    $ErrorActionPreference = 'Continue'
    & $tool @params
    $ErrorActionPreference = 'Stop'

    # Kurzzusammenfassung
    $csvDir = Join-Path $outDir 'CSV-Files'
    $summary = New-Object System.Collections.Generic.List[string]
    $summary.Add("AD-Audit $(Get-Date -Format 'dd.MM.yyyy HH:mm') - Domaene: $env:USERDNSDOMAIN - Inaktiv ab $DormantDays Tagen")
    $summary.Add('')

    $usersCsv = Join-Path $csvDir 'Users.csv'
    If (Test-Path $usersCsv)
    {
        $users     = @(Import-Csv -Path $usersCsv -Delimiter ';')
        $inCol     = "Inactive (> $DormantDays days)"
        $enabled   = @($users | Where-Object { $_.Enabled -eq 'True' })
        $inactive  = @($enabled | Where-Object { $_.$inCol -eq 'True' })
        $never     = @($enabled | Where-Object { $_.'Never Logged in' -eq 'True' })
        $summary.Add("Benutzer gesamt:                    $($users.Count)")
        $summary.Add("  aktiviert:                        $($enabled.Count)")
        $summary.Add("  deaktiviert:                      $($users.Count - $enabled.Count)")
        $summary.Add("  aktiviert und inaktiv:            $($inactive.Count)")
        $summary.Add("  aktiviert und nie angemeldet:     $($never.Count)")
        $summary.Add('')
        $inactive | Select-Object UserName, Name, 'Last Logon Date', 'Last Logon (all DCs)', whenCreated, 'Never Logged in' |
            Export-Csv -Path (Join-Path $outDir 'Inaktive-Benutzer.csv') -NoTypeInformation -Delimiter ';' -Encoding UTF8
    }

    $groupsCsv  = Join-Path $csvDir 'Groups.csv'
    $membersCsv = Join-Path $csvDir 'GroupMembers.csv'
    If ((Test-Path $groupsCsv) -and (Test-Path $membersCsv))
    {
        # privilegierte Gruppen ueber SID/RID (sprachunabhaengig)
        $privRid  = @('-512', '-518', '-519', '-548', '-549', '-550', '-551')
        $privSid  = @('S-1-5-32-544', 'S-1-5-32-548', 'S-1-5-32-549', 'S-1-5-32-550', 'S-1-5-32-551')
        $groups   = @(Import-Csv -Path $groupsCsv -Delimiter ';')
        $privNames = @($groups | Where-Object {
            $sid = [string] $_.SID
            ($privSid -contains $sid) -or ($sid -like 'S-1-5-21-*' -and ($privRid | Where-Object { $sid.EndsWith($_) }))
        } | ForEach-Object {
            $_.Name
            # GroupMembers.csv nutzt den CN aus memberOf, Groups.csv den sAMAccountName
            If ($_.DistinguishedName -match '^CN=(.+?),(OU|CN|DC)=') { $Matches[1] -replace '\\,', ',' }
        } | Sort-Object -Unique)
        $members  = @(Import-Csv -Path $membersCsv -Delimiter ';' | Where-Object { $privNames -contains $_.'Group Name' })
        $summary.Add('Direkte Mitglieder privilegierter Gruppen (Gruppen als Mitglied sind mit aufgefuehrt):')
        ForEach ($g in ($members | Group-Object 'Group Name' | Sort-Object Name))
        {
            # Gruppen als Mitglied haben bei ADRecon keinen UserName ("-"), dann den Namen nehmen
            $names = $g.Group | ForEach-Object {
                If ($_.AccountType -eq 'group' -or $_.'Member UserName' -eq '-') { "[Gruppe] $($_.'Member Name')" } Else { $_.'Member UserName' }
            }
            $summary.Add("  $($g.Name) ($($g.Count)): $(($names | Sort-Object -Unique) -join ', ')")
        }
        $members | Export-Csv -Path (Join-Path $outDir 'Privilegierte-Gruppen.csv') -NoTypeInformation -Delimiter ';' -Encoding UTF8
    }

    $summary | Set-Content -Path (Join-Path $outDir 'Zusammenfassung.txt') -Encoding UTF8
    Write-Host ''
    $summary | ForEach-Object { Write-Host $_ }
    Write-Host ''
    Write-Host "[+] Fertig. Ausgabe: $outDir"
    Write-Host '[!] Der Report enthaelt sensible AD-Daten. Ordner nach dem Abholen vom Server loeschen.'
}
Finally
{
    $ErrorActionPreference = 'Continue'
    Stop-Transcript | Out-Null
}
