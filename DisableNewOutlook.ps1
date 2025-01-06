<#
.SYNOPSIS
  Blockiert das automatische Umschalten auf das neue Outlook in Microsoft 365 
  für alle Benutzer-SIDs unter HKEY_USERS (ausgenommen ".DEFAULT").

.DESCRIPTION
  Das Skript durchsucht HKEY_USERS nach Benutzer-SIDs und setzt
  folgende Registry-Werte in jedem Benutzer-Hive (falls die Pfade existieren),
  um das neue Outlook zu deaktivieren:
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

# Liste aller Keys unter HKEY_USERS abrufen
# Ausschließen von '.DEFAULT'
$sidList = Get-ChildItem -Path "Registry::HKEY_USERS" `
            | Where-Object { $_.PSChildName -ne ".DEFAULT" }

if ($sidList.Count -eq 0) {
    Write-Host "Keine gültigen Benutzer-SIDs unter HKEY_USERS gefunden." -ForegroundColor Yellow
    return
}

Write-Host "Gefundene Benutzer-SIDs (ohne .DEFAULT):" -ForegroundColor Green
$sidList | ForEach-Object {
    Write-Host "  - $($_.PSChildName)"
}

# Registry-Unterpfade, in denen wir Einträge vornehmen wollen
$paths = @(
    "Software\Microsoft\office\16.0\Outlook\Preferences",
    "Software\Microsoft\office\16.0\Outlook\Options\General"
)

foreach($sidKey in $sidList) {

    $sid = $sidKey.PSChildName
    Write-Host "`nBearbeite SID: $sid" -ForegroundColor Cyan
    
    # Diese Sammlung beschreibt, welche Werte wir setzen wollen
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
        # Prüfen, ob dieser Pfad überhaupt existiert
        if (Test-Path $item.Key) {
            # Pfad existiert -> Wir können den Registry-Eintrag setzen
            if ($WhatIf) {
                Write-Host "Would set `"$($item.Name)`" to `"$($item.Value)`" in $($item.Key)" -ForegroundColor DarkYellow
            }
            else {
                Set-ItemProperty -Path $item.Key -Name $item.Name -Type DWord -Value $item.Value
                Write-Host "  -> $($item.Name) wurde auf $($item.Value) gesetzt." -ForegroundColor Green
            }
        }
        else {
            # Pfad existiert nicht -> Überspringen
            Write-Host "  -> Überspringe $($item.Key), da es nicht existiert." -ForegroundColor Yellow
        }
    }
}

Write-Host "`nSkript wurde beendet. Alle Einträge sind (ggf. im WhatIf-Modus nur simuliert) gesetzt." -ForegroundColor Cyan
