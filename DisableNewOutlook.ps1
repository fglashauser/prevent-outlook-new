<#
.SYNOPSIS
  Blockiert das automatische Umschalten auf das neue Outlook in Microsoft 365 
  für alle Benutzer-SIDs unter HKEY_USERS.

.DESCRIPTION
  Das Skript durchsucht HKEY_USERS nach Benutzer-SIDs (S-1-5-21-...) und setzt
  folgende Registry-Werte in jedem Benutzerhive, um das neue Outlook zu deaktivieren:
    - HKCU:\Software\Microsoft\office\16.0\Outlook\Preferences
        "NewOutlookMigrationUserSetting" = 0 (DWORD)
        "UseNewOutlook"                  = 0 (DWORD)

    - HKCU:\Software\Microsoft\office\16.0\Outlook\Options\General
        "DoNewOutlookAutoMigration"              = 0 (DWORD)
        "NewOutlookAutoMigrationRetryIntervals"  = 0 (DWORD)
        "HideNewOutlookToggle"                   = 1 (DWORD)

.PARAMETER WhatIf
  Zeigt an, was ausgeführt werden würde, ohne Änderungen vorzunehmen.

.EXAMPLE
  .\DisableNewOutlook.ps1

.EXAMPLE
  .\DisableNewOutlook.ps1 -WhatIf
  Zeigt nur an, welche Änderungen vorgenommen würden.
#>

param(
    [switch]$WhatIf
)

Write-Host "Starte das Skript zum Deaktivieren des neuen Outlook..." -ForegroundColor Cyan

# Alle möglichen Unter-Keys von HKEY_USERS auslesen
#$sidList = Get-ChildItem -Path "Registry::HKEY_USERS" | Where-Object {
#    $_.PSChildName -match "^S-1-5-"
#}
$sidList = Get-ChildItem -Path "Registry::HKEY_USERS"

if ($sidList.Count -eq 0) {
    Write-Host "Keine gültigen Benutzer-SIDs unter HKEY_USERS gefunden." -ForegroundColor Yellow
    return
}

Write-Host "Gefundene Benutzer-SIDs:" -ForegroundColor Green
$sidList | ForEach-Object {
    Write-Host "  - $($_.PSChildName)"
}

# Die einzelnen Registry-Pfade, in denen wir Einträge vornehmen wollen:
$paths = @(
    "Software\Microsoft\office\16.0\Outlook\Preferences",
    "Software\Microsoft\office\16.0\Outlook\Options\General"
)

foreach($sidKey in $sidList) {
    $sid = $sidKey.PSChildName
    
    Write-Host "`nBearbeite SID: $sid" -ForegroundColor Cyan

    # 1) Die Registry-Pfade anlegen (falls sie noch nicht existieren).
    foreach($path in $paths) {
        $fullPath = "Registry::HKEY_USERS\$sid\$path"
        if ($WhatIf) {
            Write-Host "Would create path $fullPath" -ForegroundColor DarkYellow
        }
        else {
            New-Item -Path $fullPath -Force | Out-Null
        }
    }

    # 2) Jetzt die benötigten Werte setzen:
    #    - Software\Microsoft\office\16.0\Outlook\Preferences
    #        * NewOutlookMigrationUserSetting = 0
    #        * UseNewOutlook                  = 0
    #    - Software\Microsoft\office\16.0\Outlook\Options\General
    #        * DoNewOutlookAutoMigration             = 0
    #        * NewOutlookAutoMigrationRetryIntervals = 0
    #        * HideNewOutlookToggle                  = 1

    $regValues = @(
        @{
            Key   = "Registry::HKEY_USERS\$sid\Software\Microsoft\office\16.0\Outlook\Preferences"
            Name  = "NewOutlookMigrationUserSetting"
            Value = 0
        },
        @{
            Key   = "Registry::HKEY_USERS\$sid\Software\Microsoft\office\16.0\Outlook\Preferences"
            Name  = "UseNewOutlook"
            Value = 0
        },
        @{
            Key   = "Registry::HKEY_USERS\$sid\Software\Microsoft\office\16.0\Outlook\Options\General"
            Name  = "DoNewOutlookAutoMigration"
            Value = 0
        },
        @{
            Key   = "Registry::HKEY_USERS\$sid\Software\Microsoft\office\16.0\Outlook\Options\General"
            Name  = "NewOutlookAutoMigrationRetryIntervals"
            Value = 0
        },
        @{
            Key   = "Registry::HKEY_USERS\$sid\Software\Microsoft\office\16.0\Outlook\Options\General"
            Name  = "HideNewOutlookToggle"
            Value = 1
        }
    )

    foreach($item in $regValues) {
        if ($WhatIf) {
            Write-Host "Would set `"$($item.Name)`" to `"$($item.Value)`" in $($item.Key)" -ForegroundColor DarkYellow
        }
        else {
            Set-ItemProperty -Path $item.Key -Name $item.Name -Type DWord -Value $item.Value
            Write-Host "  -> $($item.Name) wurde auf $($item.Value) gesetzt." -ForegroundColor Green
        }
    }
}

Write-Host "`nSkript wurde beendet. Alle Einträge sind (ggf. im WhatIf-Modus nur simuliert) gesetzt." -ForegroundColor Cyan