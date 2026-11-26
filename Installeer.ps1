<#
  Plant de Studie-overzicht-taken in via Windows Taakplanner (alleen voor jouw account)
  en maakt een snelkoppeling "Studie-overzicht" in het Startmenu.
  Opnieuw uitvoeren is veilig: bestaande taken worden overschreven.
  Verwijderen: .\Installeer.ps1 -Verwijder
#>
param(
    [switch]$Verwijder,
    [string]$TimeEditLink    # je persoonlijke TimeEdit-abonnementslink (.ics); alleen nodig bij de eerste installatie
)
$ErrorActionPreference = 'Stop'

# Persoonlijke instellingen: config.json (staat niet in git)
$configPad = Join-Path $PSScriptRoot 'config.json'
if (-not $Verwijder) {
    if ($TimeEditLink) {
        if ($TimeEditLink -notmatch '^https://cloud\.timeedit\.net/.+\.ics$') {
            throw "Dit lijkt geen TimeEdit-abonnementslink. Hij begint met https://cloud.timeedit.net/ en eindigt op .ics"
        }
        $basis = if (Test-Path $configPad) { $configPad } else { Join-Path $PSScriptRoot 'config.example.json' }
        $c = Get-Content $basis -Raw -Encoding UTF8 | ConvertFrom-Json
        $c.TimeEditIcs = $TimeEditLink
        $c | ConvertTo-Json -Depth 5 | Set-Content $configPad -Encoding UTF8
        Write-Host 'config.json aangemaakt met jouw TimeEdit-link.'
    }
    if (-not (Test-Path $configPad)) {
        throw "Nog geen config.json. Voer de installatie uit met je TimeEdit-link:`n  .\Installeer.ps1 -TimeEditLink `"https://cloud.timeedit.net/.../....ics`""
    }
}
$script = Join-Path $PSScriptRoot 'StudieOverzicht.ps1'
$map    = 'Studie-overzicht'
$namen  = 'Briefing bij aanmelden', 'Weekoverzicht zondag', 'Elk uur vernieuwen', 'Vakknoppen-server'
$lnk    = Join-Path ([Environment]::GetFolderPath('Programs')) 'Studie-overzicht.lnk'
$vakMap = Join-Path ([Environment]::GetFolderPath('Programs')) 'Studie-overzicht vakken'
$proto  = 'HKCU:\Software\Classes\studievak'

if ($Verwijder) {
    foreach ($n in $namen) { Unregister-ScheduledTask -TaskPath "\$map\" -TaskName $n -Confirm:$false -ErrorAction SilentlyContinue }
    Remove-Item $lnk -ErrorAction SilentlyContinue
    Remove-Item $vakMap -Recurse -ErrorAction SilentlyContinue
    Remove-Item $proto -Recurse -ErrorAction SilentlyContinue
    Write-Host 'Taken, snelkoppelingen en vakknoppen verwijderd.'
    return
}

function New-Actie([string]$argumenten) {
    # conhost --headless: draait zonder zwart venster
    New-ScheduledTaskAction -Execute "$env:WINDIR\System32\conhost.exe" `
        -Argument "--headless powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$script`" $argumenten" `
        -WorkingDirectory $PSScriptRoot
}
$user = "$env:USERDOMAIN\$env:USERNAME"
$principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Limited
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 10) -MultipleInstances IgnoreNew

# 1. Briefing bij aanmelden (1 minuut wachten zodat netwerk en Outlook klaar zijn)
$t1 = New-ScheduledTaskTrigger -AtLogOn -User $user
$t1.Delay = 'PT1M'
Register-ScheduledTask -TaskPath "\$map\" -TaskName $namen[0] -Action (New-Actie '-Modus Dagelijks') `
    -Trigger $t1 -Principal $principal -Settings $settings -Force `
    -Description 'Opent het studiedashboard (max. 1x per dag) en zet nieuwe deadlines in Outlook.' | Out-Null

# 2. Weekoverzicht zondag 20:00
$t2 = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At '20:00'
Register-ScheduledTask -TaskPath "\$map\" -TaskName $namen[1] -Action (New-Actie '-Modus Week') `
    -Trigger $t2 -Principal $principal -Settings $settings -Force `
    -Description 'Maakt en opent het weekoverzicht van de komende week.' | Out-Null

# 3. Elk uur stil vernieuwen (dashboard + deadlines in agenda)
$t3 = New-ScheduledTaskTrigger -Once -At (Get-Date).Date.AddHours((Get-Date).Hour + 1) `
    -RepetitionInterval (New-TimeSpan -Hours 1)
Register-ScheduledTask -TaskPath "\$map\" -TaskName $namen[2] -Action (New-Actie '-Modus Vernieuwen') `
    -Trigger $t3 -Principal $principal -Settings $settings -Force `
    -Description 'Houdt het dashboard en de deadlines in Outlook up-to-date.' | Out-Null

# 4. Lokale server voor de vakknoppen (alleen 127.0.0.1), start bij aanmelden en blijft draaien
$serverActie = New-ScheduledTaskAction -Execute "$env:WINDIR\System32\conhost.exe" `
    -Argument "--headless powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$(Join-Path $PSScriptRoot 'Server.ps1')`"" `
    -WorkingDirectory $PSScriptRoot
$serverSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit ([TimeSpan]::Zero) -MultipleInstances IgnoreNew -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)
Register-ScheduledTask -TaskPath "\$map\" -TaskName $namen[3] -Action $serverActie `
    -Trigger (New-ScheduledTaskTrigger -AtLogOn -User $user) -Principal $principal -Settings $serverSettings -Force `
    -Description 'Lokale server (127.0.0.1:48765) voor het dashboard en de vakknoppen.' | Out-Null
Stop-ScheduledTask -TaskPath "\$map\" -TaskName $namen[3] -ErrorAction SilentlyContinue
Start-ScheduledTask -TaskPath "\$map\" -TaskName $namen[3]

# Snelkoppeling in Startmenu: vernieuwen + dashboard openen
$sh = New-Object -ComObject WScript.Shell
$s = $sh.CreateShortcut($lnk)
$s.TargetPath = "$env:WINDIR\System32\conhost.exe"
$s.Arguments = "--headless powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$script`" -Openen"
$s.WorkingDirectory = $PSScriptRoot
$s.IconLocation = "$env:WINDIR\System32\shell32.dll,165"
$s.Description = 'Studie-overzicht vernieuwen en openen'
$s.Save()

# Vakknoppen: link "studievak:<vak>" (alleen voor jouw account) -> OpenVak.ps1
$openVak = Join-Path $PSScriptRoot 'OpenVak.ps1'
New-Item -Path "$proto\shell\open\command" -Force | Out-Null
Set-ItemProperty -Path $proto -Name '(default)' -Value 'URL:Studie-overzicht vak openen'
Set-ItemProperty -Path $proto -Name 'URL Protocol' -Value ''
Set-ItemProperty -Path "$proto\shell\open\command" -Name '(default)' `
    -Value "`"$env:WINDIR\System32\conhost.exe`" --headless powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$openVak`" `"%1`""

# Snelkoppeling per vak in het Startmenu
$vk = Get-Content (Join-Path $PSScriptRoot 'vakken.json') -Raw -Encoding UTF8 | ConvertFrom-Json
New-Item -ItemType Directory -Force $vakMap | Out-Null
Get-ChildItem $vakMap -Filter *.lnk | Remove-Item
foreach ($v in $vk.Vakken) {
    $s = $sh.CreateShortcut((Join-Path $vakMap "$($v.Naam).lnk"))
    $s.TargetPath = "$env:WINDIR\System32\conhost.exe"
    $s.Arguments = "--headless powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$openVak`" $($v.Id)"
    $s.WorkingDirectory = $PSScriptRoot
    $s.IconLocation = "$env:WINDIR\System32\shell32.dll,43"
    $s.Description = "$($v.Naam): cursussen, Blackboard en programma's openen"
    $s.Save()
}

Get-ScheduledTask -TaskPath "\$map\" | ForEach-Object {
    $i = $_ | Get-ScheduledTaskInfo
    '{0,-26} volgende keer: {1}' -f $_.TaskName, $(if ($i.NextRunTime) { $i.NextRunTime } else { 'bij aanmelden' })
}
"Snelkoppeling: $lnk"
"Vakknoppen: $vakMap ($(@($vk.Vakken).Count) vakken)"

# Eerste keer: alles inlezen en het dashboard openen
Write-Host 'Rooster en mail inlezen en dashboard openen...'
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $script -Openen
