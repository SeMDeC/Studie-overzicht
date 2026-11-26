<#
  Studie-overzicht
  ----------------
  Haalt je lessenrooster uit TimeEdit, leest je PXL-mail in Outlook, haalt
  deadlines uit de Blackboard-meldingen, zet die deadlines in je Outlook-agenda
  en maakt overzichtspagina's (dashboard + weekoverzicht).

  Gebruik:
    .\StudieOverzicht.ps1                      # alles vernieuwen (stil)
    .\StudieOverzicht.ps1 -Modus Dagelijks     # vernieuwen + dashboard openen (ochtendbriefing)
    .\StudieOverzicht.ps1 -Modus Week          # weekoverzicht van volgende week maken + openen
    .\StudieOverzicht.ps1 -Openen              # vernieuwen + dashboard openen
    .\StudieOverzicht.ps1 -GeenAgenda          # niets in Outlook-agenda schrijven (test)
#>
param(
    [ValidateSet('Vernieuwen', 'Dagelijks', 'Week')]
    [string]$Modus = 'Vernieuwen',
    [switch]$Openen,
    [switch]$GeenAgenda,
    [switch]$NietOpenen
)

$ErrorActionPreference = 'Stop'
$root    = $PSScriptRoot
if (-not (Test-Path (Join-Path $root 'config.json'))) {
    throw 'config.json ontbreekt. Voer eerst .\Installeer.ps1 -TimeEditLink "<jouw TimeEdit .ics-link>" uit (zie README.md).'
}
$cfg     = Get-Content (Join-Path $root 'config.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$dataDir = Join-Path $root 'data'
New-Item -ItemType Directory -Force $dataDir | Out-Null
$logFile   = Join-Path $dataDir 'log.txt'
$stateFile = Join-Path $dataDir 'deadlines.json'
$nl        = [Globalization.CultureInfo]::GetCultureInfo('nl-BE')
$nu        = Get-Date

function Write-Log([string]$msg) {
    $line = '{0:yyyy-MM-dd HH:mm:ss} [{1}] {2}' -f (Get-Date), $Modus, $msg
    Add-Content -Path $logFile -Value $line -Encoding UTF8
    Write-Host $line
}
function Enc([string]$s) { [System.Net.WebUtility]::HtmlEncode($s) }

# ---------------------------------------------------------------- rooster --
function ConvertFrom-IcsDate([string]$s) {
    if ($s -match 'Z$') {
        return [datetime]::ParseExact($s, "yyyyMMdd'T'HHmmss'Z'", $null,
            [Globalization.DateTimeStyles]'AssumeUniversal,AdjustToUniversal').ToLocalTime()
    }
    if ($s.Length -eq 8) { return [datetime]::ParseExact($s, 'yyyyMMdd', $null) }
    return [datetime]::ParseExact($s.Substring(0, 15), "yyyyMMdd'T'HHmmss", $null)
}

function Get-Rooster {
    $cache = Join-Path $dataDir 'rooster.ics'
    try {
        Invoke-WebRequest -UseBasicParsing -Uri $cfg.TimeEditIcs -OutFile $cache -TimeoutSec 30
    } catch {
        Write-Log "Rooster ophalen mislukt, laatste kopie gebruikt: $($_.Exception.Message)"
    }
    if (-not (Test-Path $cache)) { return @() }

    $raw = [IO.File]::ReadAllText($cache, [Text.Encoding]::UTF8)
    $raw = $raw -replace "\r?\n[ \t]", ''            # gevouwen regels samenvoegen
    $lijst = foreach ($blk in [regex]::Matches($raw, '(?s)BEGIN:VEVENT(.*?)END:VEVENT')) {
        $p = @{}
        foreach ($line in ($blk.Groups[1].Value -split "\r?\n")) {
            if ($line -match '^([A-Z-]+)(;[^:]*)?:(.*)$') {
                $p[$matches[1]] = ($matches[3] -replace '\\n', "`n" -replace '\\([,;\\])', '$1')
            }
        }
        if (-not $p['DTSTART']) { continue }
        $summary = "$($p['SUMMARY'])".Trim()
        $code = ''; $vak = $summary
        if ($summary -match '^(\S*\d\S*)\s+(.+)$') { $code = $matches[1]; $vak = $matches[2] }
        $loc = "$($p['LOCATION'])".Trim()
        if ($loc -match 'Nieuwe lokaalnaam:\s*(.+?)\.?\s+Oude lokaalnaam:\s*(\S+)') {
            $loc = '{0} ({1})' -f $matches[1].TrimEnd('.'), $matches[2]
        }
        $docenten = ("$($p['DESCRIPTION'])" -split "`n" |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ -and $_ -notmatch '^ID\s*\d+' }) -join ', '
        [pscustomobject]@{
            Start    = ConvertFrom-IcsDate $p['DTSTART']
            Einde    = ConvertFrom-IcsDate $p['DTEND']
            Code     = $code
            Vak      = $vak
            Lokaal   = $loc
            Docenten = $docenten
        }
    }
    return @($lijst | Sort-Object Start)
}

# ------------------------------------------------------------------ mail --
$maanden = @{ januari=1; februari=2; maart=3; april=4; mei=5; juni=6; juli=7;
              augustus=8; september=9; oktober=10; november=11; december=12 }

function ConvertFrom-BbDatum([string]$s) {
    # "woensdag 25 november 2026 23:59:00 CET"
    if ($s -match '(\d{1,2})\s+([a-z]+)\s+(\d{4})\s+(\d{1,2}):(\d{2})') {
        $m = $maanden[$matches[2].ToLower()]
        if ($m) { return Get-Date -Year $matches[3] -Month $m -Day $matches[1] -Hour $matches[4] -Minute $matches[5] -Second 0 -Millisecond 0 }
    }
    return $null
}

function Get-VakUitRegel([string]$line) {
    # "41SNE1020 Windows Essentials (2026-27 - GRSNE) (2627-GRSNE-41SNE1020)"
    # "Communication Skills 1 (2026-27 - GRDVO - ...)" (zonder vakcode) mag ook
    if ($line -match '^\s*(.+?)\s+\(\d{4}-\d{2}') { return $matches[1].Trim() }
    return $null
}

# Vakken (voor knoppen en om deadlines per vak te groeperen)
$vakLijst = @()
if (Test-Path (Join-Path $root 'vakken.json')) {
    $vakLijst = @((Get-Content (Join-Path $root 'vakken.json') -Raw -Encoding UTF8 | ConvertFrom-Json).Vakken)
}
function Get-VakNaamVoor($d) {
    $tekst = "$($d.Titel) $($d.Vak)"
    foreach ($v in $vakLijst) {
        if ($tekst -like "*$($v.Code)*" -or $tekst -like "*$($v.Naam)*") { return $v.Naam }
    }
    return 'Overig'
}

function Test-Schoolmail($mail) {
    $wie = "$($mail.SenderName) $($mail.SenderEmailAddress) $($mail.Subject)"
    foreach ($a in $cfg.SchoolAfzenders) { if ($wie -like "*$a*") { return $true } }
    return $false
}

function Get-MailInfo($ns) {
    $inbox = $ns.GetDefaultFolder(6)
    $items = $inbox.Items
    $items.Sort('[ReceivedTime]', $true)
    $sinds = $nu.AddDays(-[int]$cfg.MailDagenTerug)
    $toonSinds = $nu.AddDays(-[int]$cfg.SchoolmailDagenTonen)

    $deadlines  = [ordered]@{}   # sleutel: titel|vak  -> nieuwste melding wint
    $ingediend  = @{}
    $schoolmail = New-Object System.Collections.ArrayList

    foreach ($m in $items) {
        if ($m.Class -ne 43) { continue }                 # alleen mails
        if ($m.ReceivedTime -lt $sinds) { break }
        $subject = "$($m.Subject)"

        if ($subject -eq 'Dagelijkse meldingen') {
            $lines = @(("$($m.Body)" -split "\r?\n") | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            for ($i = 0; $i -lt $lines.Count; $i++) {
                if ($lines[$i] -match '^(?<t>.+?)\s*<[^>]+>\s+moet binnenkort worden ingeleverd:\s*(?<d>.+)$') {
                    $titel = $matches['t'].Trim(); $due = ConvertFrom-BbDatum $matches['d']
                    $vak = $null
                    for ($j = $i + 1; $j -lt [Math]::Min($i + 3, $lines.Count); $j++) {
                        $vak = Get-VakUitRegel $lines[$j]; if ($vak) { break }
                    }
                    if (-not $vak) { $vak = 'Blackboard' }
                    $key = "$titel|$vak"
                    if ($due -and -not $deadlines.Contains($key)) {
                        $deadlines[$key] = [pscustomobject]@{ Sleutel = $key; Titel = $titel; Vak = $vak; Deadline = $due }
                    }
                }
            }
            continue
        }

        if ($subject -eq 'Inzending ontvangen') {
            $lines = @(("$($m.Body)" -split "\r?\n") | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            $vak = $null; $titel = $null
            for ($i = 0; $i -lt $lines.Count; $i++) {
                if (-not $vak) { $v = Get-VakUitRegel $lines[$i]; if ($v) { $vak = $v } }
                if ($lines[$i] -eq 'Beoordeling ingezonden' -and $i + 1 -lt $lines.Count) { $titel = $lines[$i + 1] }
            }
            if ($titel) { $ingediend["$titel|$vak"] = $m.ReceivedTime; $ingediend["$titel|*"] = $m.ReceivedTime }
            continue
        }

        if ($m.ReceivedTime -ge $toonSinds -and (Test-Schoolmail $m)) {
            $tekst = "$subject $($m.Body)".ToLower()
            $signaal = @($cfg.SignaalWoorden | Where-Object { $tekst.Contains($_) })
            $afz = "$($m.SenderName)" -replace '\s+-\s+\S+@\S+$', ''
            [void]$schoolmail.Add([pscustomobject]@{
                Ontvangen = $m.ReceivedTime
                Van       = $afz
                Onderwerp = $subject
                Ongelezen = [bool]$m.UnRead
                Signaal   = ($signaal -join ', ')
            })
        }
    }

    foreach ($d in $deadlines.Values) {
        $t = $null
        if ($ingediend.ContainsKey($d.Sleutel)) { $t = $ingediend[$d.Sleutel] }
        elseif ($ingediend.ContainsKey("$($d.Titel)|*")) { $t = $ingediend["$($d.Titel)|*"] }
        $d | Add-Member -NotePropertyName Ingediend -NotePropertyValue $t
    }
    $grens = $nu.AddDays([int]$cfg.DeadlineMaxDagenVooruit)
    return [pscustomobject]@{
        Deadlines  = @($deadlines.Values | Where-Object { $_.Deadline -le $grens } | Sort-Object Deadline)
        Schoolmail = @($schoolmail)
    }
}

# ------------------------------------------------------- outlook-agenda --
function Sync-DeadlinesNaarAgenda($ns, $deadlines) {
    $state = @{}
    if (Test-Path $stateFile) {
        (Get-Content $stateFile -Raw -Encoding UTF8 | ConvertFrom-Json).PSObject.Properties |
            ForEach-Object { $state[$_.Name] = $_.Value }
    }
    $cat = $cfg.DeadlineCategorie
    $bestaat = $false
    foreach ($c in $ns.Categories) { if ($c.Name -eq $cat) { $bestaat = $true } }
    if (-not $bestaat) { [void]$ns.Categories.Add($cat, 1) }   # 1 = rood

    $agenda = $ns.GetDefaultFolder(9)
    $nieuw = 0; $bijgewerkt = 0
    foreach ($d in $deadlines) {
        $titelTekst = 'Deadline: {0} ({1})' -f $d.Titel, $d.Vak
        if ($d.Ingediend) { $titelTekst = 'Ingediend - ' + $titelTekst }
        $s = $state[$d.Sleutel]
        $appt = $null
        if ($s) {
            if ($s.Status -eq 'verwijderd') { continue }
            try { $appt = $ns.GetItemFromID($s.EntryID) } catch { $appt = $null }
            if (-not $appt -or $appt.Parent.EntryID -eq $ns.GetDefaultFolder(3).EntryID) {
                # door jou verwijderd in Outlook: niet opnieuw aanmaken
                $state[$d.Sleutel] = [pscustomobject]@{ EntryID = $s.EntryID; Status = 'verwijderd'; Deadline = $s.Deadline }
                continue
            }
        } else {
            if ($d.Deadline -lt $nu) { continue }            # oude deadlines niet meer toevoegen
            $appt = $agenda.Items.Add(1)
            $appt.BusyStatus = 0                              # vrij: blokkeert je agenda niet
            $appt.Categories = $cat
            $appt.Body = "Vak: $($d.Vak)`r`nOpdracht: $($d.Titel)`r`nIn te leveren op Blackboard voor $($d.Deadline.ToString('dddd d MMMM yyyy HH:mm', $nl)).`r`n`r`nAutomatisch toegevoegd door Studie-overzicht."
            $nieuw++
        }
        $start = $d.Deadline.AddMinutes(-30)
        $wijzig = ($appt.Subject -ne $titelTekst) -or ($appt.Start -ne $start)
        if ($wijzig -or -not $s) {
            $appt.Subject = $titelTekst
            $appt.Start = $start
            $appt.End = $d.Deadline
            $appt.ReminderSet = (-not $d.Ingediend) -and ($d.Deadline -gt $nu)
            if ($appt.ReminderSet) { $appt.ReminderMinutesBeforeStart = [int]$cfg.HerinneringMinutenVooraf }
            $appt.Save()
            if ($s) { $bijgewerkt++ }
        }
        $state[$d.Sleutel] = [pscustomobject]@{ EntryID = $appt.EntryID; Status = 'actief'; Deadline = $d.Deadline.ToString('s') }
    }
    $obj = New-Object PSObject
    foreach ($k in $state.Keys) { $obj | Add-Member -NotePropertyName $k -NotePropertyValue $state[$k] }
    $obj | ConvertTo-Json -Depth 4 | Set-Content $stateFile -Encoding UTF8
    Write-Log "Agenda: $nieuw deadline(s) toegevoegd, $bijgewerkt bijgewerkt."
}

# ------------------------------------------- items uit je Outlook-agenda --
# Alles in je agenda met de categorie Deadline, PE / toets of Activiteit
# (behalve de deadlines die dit script zelf uit Blackboard aanmaakt).
function Get-AgendaItems($ns) {
    $bbIds = @{}
    if (Test-Path $stateFile) {
        (Get-Content $stateFile -Raw -Encoding UTF8 | ConvertFrom-Json).PSObject.Properties |
            ForEach-Object { $bbIds["$($_.Value.EntryID)"] = 1 }
    }
    $items = $ns.GetDefaultFolder(9).Items
    $items.Sort('[Start]')
    $items.IncludeRecurrences = $true
    $van = $nu.Date.AddDays(-7); $tot = $nu.Date.AddDays(70)
    $f = "[End] >= '" + $van.ToString('g') + "' AND [Start] < '" + $tot.ToString('g') + "'"
    $lijst = New-Object System.Collections.ArrayList
    foreach ($a in $items.Restrict($f)) {
        $cat = "$($a.Categories)"
        $soort = $null
        if ($cat -match 'PE / toets') { $soort = 'toets' }
        elseif ($cat -match $cfg.DeadlineCategorie) { $soort = 'deadline' }
        if (-not $soort -or $bbIds.ContainsKey("$($a.EntryID)")) { continue }
        [void]$lijst.Add([pscustomobject]@{
            Titel   = ("$($a.Subject)" -replace '^Deadline:\s*', '')
            Start   = $a.Start
            Einde   = $a.End
            HeleDag = [bool]$a.AllDayEvent
            Locatie = "$($a.Location)"
            Soort   = $soort
        })
    }
    return @($lijst)
}

# Outlook-deadlines/toetsen omzetten naar hetzelfde formaat als Blackboard-deadlines
function Get-AlleDeadlines($info, $agenda) {
    $uit = @($info.Deadlines)
    $gezien = @{}
    foreach ($x in ($agenda | Where-Object { $_.Soort -ne 'activiteit' -and $_.Einde -ge $nu } | Sort-Object Start)) {
        if ($gezien.ContainsKey($x.Titel)) { continue }      # wekelijkse herhaling: alleen de eerstvolgende
        $gezien[$x.Titel] = 1
        $moment = if ($x.HeleDag -or $x.Soort -eq 'toets') { $x.Start } else { $x.Einde }
        $uit += [pscustomobject]@{
            Titel     = $x.Titel
            Vak       = $(if ($x.Soort -eq 'toets') { 'PE / toets' } elseif ($x.Locatie) { $x.Locatie } else { 'Outlook-agenda' })
            Deadline  = $moment
            Verloopt  = $x.Einde
            HeleDag   = $x.HeleDag
            Ingediend = $null
            Soort     = $x.Soort
        }
    }
    return @($uit | Sort-Object Deadline)
}

function Format-Moment($d) {
    if ($d.HeleDag) {
        $s = $d.Deadline.ToString('ddd d MMM', $nl)
        if ($d.Verloopt -and ($d.Verloopt.Date - $d.Deadline.Date).TotalDays -gt 1) { $s += ' t/m ' + $d.Verloopt.AddDays(-1).ToString('ddd d MMM', $nl) }
        return $s
    }
    return $d.Deadline.ToString('ddd d MMM HH:mm', $nl)
}

# ------------------------------------------------------------------ html --
$css = @'
:root{
  --papier:#F4F6F8;--vlak:#FFFFFF;--inkt:#18212B;--zacht:#5B6676;--lijn:#DCE1E7;
  --accent:#2B59C3;--accent-zacht:#E8EEFB;
  --groen:#1E7A4F;--groen-zacht:#E3F3EA;--amber:#A8650A;--amber-zacht:#FBF0DC;--rood:#B42318;--rood-zacht:#FCE8E6;
  --display:"Segoe UI Variable Display","Segoe UI",system-ui,sans-serif;
  --tekst:"Segoe UI Variable Text","Segoe UI",system-ui,sans-serif;
}
@media (prefers-color-scheme:dark){:root{
  --papier:#12161C;--vlak:#1A2028;--inkt:#E6EAF0;--zacht:#98A2B3;--lijn:#2A323D;
  --accent:#7FA2F5;--accent-zacht:#1E2A44;
  --groen:#5CC98E;--groen-zacht:#16301F;--amber:#F0B25A;--amber-zacht:#33270F;--rood:#F2857B;--rood-zacht:#3A1B18}}
*{box-sizing:border-box}
html{-webkit-text-size-adjust:100%}
body{margin:0;background:var(--papier);color:var(--inkt);font:15px/1.5 var(--tekst);font-variant-numeric:tabular-nums}
.wrap{max-width:1120px;margin:0 auto;padding:40px 32px 56px}
a{color:inherit}
/* kop */
.kop{display:flex;justify-content:space-between;align-items:flex-start;gap:24px;flex-wrap:wrap}
.kop .naam{margin:0 0 6px;font:600 13px/1.3 var(--tekst);color:var(--accent)}
h1{margin:0;font:600 36px/1.1 var(--display);letter-spacing:-.015em}
.samenvatting{margin:10px 0 0;max-width:62ch;color:var(--zacht);font-size:16px}
.samenvatting b{color:var(--inkt);font-weight:600}
.bijgewerkt{color:var(--zacht);font-size:13px;padding-top:4px;white-space:nowrap}
/* secties */
section{margin-top:40px}
h2{margin:0 0 14px;font:600 17px/1.3 var(--display);display:flex;justify-content:space-between;align-items:baseline;gap:12px}
h2 small{font:400 13px var(--tekst);color:var(--zacht)}
.kolommen{display:grid;grid-template-columns:minmax(0,3fr) minmax(0,2fr);gap:48px}
@media (max-width:880px){.kolommen{grid-template-columns:1fr;gap:0}}
.leeg{color:var(--zacht);margin:0}
/* tijdlijn van vandaag */
.dag{background:var(--vlak);border:1px solid var(--lijn);border-radius:14px;padding:20px 22px 16px;margin-top:28px}
.tl{position:relative;height:86px;margin:6px 0 0}
.tl-uren{position:absolute;inset:auto 0 0 0;height:18px}
.tl-uur{position:absolute;bottom:0;transform:translateX(-50%);font-size:12px;color:var(--zacht)}
.tl-raster{position:absolute;top:0;bottom:22px;width:1px;background:var(--lijn)}
.tl-blok{position:absolute;top:4px;height:56px;border-radius:9px;background:var(--accent-zacht);border-left:3px solid var(--accent);padding:6px 9px;overflow:hidden}
.tl-blok b{display:block;font-size:13px;line-height:1.25;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.tl-blok span{display:block;font-size:12px;color:var(--zacht);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.tl-blok{z-index:1}
.tl-blok.voorbij{background:var(--papier);border-left-color:var(--lijn);color:var(--zacht)}
.tl-blok.nu{background:var(--accent);color:#fff;border-left-color:var(--accent)}
.tl-blok.nu span{color:rgba(255,255,255,.85)}
@media (prefers-color-scheme:dark){.tl-blok.nu{color:#0F1726}.tl-blok.nu span{color:rgba(15,23,38,.8)}}
.tl-nu{position:absolute;top:-4px;bottom:18px;width:2px;background:var(--rood);border-radius:2px}
.tl-nu::before{content:"";position:absolute;top:-3px;left:-3px;width:8px;height:8px;border-radius:50%;background:var(--rood)}
.dag-leeg{color:var(--zacht);margin:4px 0 2px}
@media (max-width:700px){
  .tl{height:auto}.tl-uren,.tl-raster,.tl-nu{display:none}
  .tl-blok{position:static;height:auto;margin-bottom:6px;width:auto!important}
}
/* vakknoppen */
.vakken{display:flex;flex-wrap:wrap;gap:8px}
.vakknop{display:inline-flex;align-items:center;gap:8px;padding:9px 14px;border:1px solid var(--lijn);border-radius:999px;background:var(--vlak);text-decoration:none;font-weight:600;font-size:14px;cursor:pointer}
.vakknop:hover{border-color:var(--accent);color:var(--accent)}
.vakknop:focus-visible{outline:2px solid var(--accent);outline-offset:2px}
.vakknop .stip{width:7px;height:7px;border-radius:50%;background:var(--lijn)}
.vakknop.vandaag .stip{background:var(--accent)}
.vakknop .status{font-weight:400;font-size:12px;color:var(--zacht)}
.vakknop.bezig{opacity:.6;cursor:progress}
.vakknop.ok{border-color:var(--groen);color:var(--groen)}.vakknop.ok .stip{background:var(--groen)}
.vakknop.fout{border-color:var(--rood);color:var(--rood)}.vakknop.fout .stip{background:var(--rood)}
.vakken-uitleg{margin:10px 0 0;color:var(--zacht);font-size:13px}
/* deadlines per vak */
.vakgroep+.vakgroep{margin-top:22px}
.vakgroep h3{margin:0;padding-bottom:6px;border-bottom:1px solid var(--lijn);display:flex;justify-content:space-between;align-items:baseline;font:600 14px var(--tekst)}
.vakgroep h3 small{font-weight:400;font-size:12px;color:var(--zacht)}
.taak{display:grid;grid-template-columns:76px minmax(0,1fr) auto;gap:14px;align-items:baseline;padding:10px 0;border-bottom:1px solid var(--lijn)}
.taak .wanneer{font-size:13px;line-height:1.3}
.taak .wanneer small{display:block;color:var(--zacht);font-size:12px}
.taak .wat{line-height:1.35}
.taak .wat small{display:block;color:var(--zacht);font-size:12px;margin-top:2px}
.st{font-size:12px;font-weight:600;padding:2px 8px;border-radius:6px;white-space:nowrap}
.st-rood{background:var(--rood-zacht);color:var(--rood)}.st-amber{background:var(--amber-zacht);color:var(--amber)}
.st-groen{background:var(--groen-zacht);color:var(--groen)}.st-rustig{color:var(--zacht);font-weight:400;padding-right:0}
.taak.klaar .wat,.taak.klaar .wanneer{color:var(--zacht)}
/* mails */
.mail{padding:10px 0;border-bottom:1px solid var(--lijn)}
.mail:first-of-type{padding-top:0}
.mail .ond{display:block;line-height:1.35}
.mail.ongelezen .ond{font-weight:600}
.mail small{color:var(--zacht);font-size:12px}
.mail .let{color:var(--amber);font-weight:600}
/* weekraster */
.week{display:grid;grid-template-columns:repeat(7,minmax(0,1fr));border-top:1px solid var(--lijn)}
.wdag{padding:10px 10px 14px;border-right:1px solid var(--lijn);min-height:110px}
.wdag:last-child{border-right:0}
.wdag h3{margin:0 0 8px;font:600 13px var(--tekst)}
.wdag h3 small{font-weight:400;color:var(--zacht)}
.wdag.vandaag h3{color:var(--accent)}
.wdag .vrij{color:var(--zacht);font-size:12px}
.item{font-size:12px;line-height:1.3;padding:3px 0 3px 8px;border-left:2px solid var(--accent);margin-bottom:6px}
.item b{display:block;font-weight:600}
.item.deadline{border-left-color:var(--rood)}.item.deadline b{color:var(--rood)}
.item.toets{border-left-color:var(--amber)}.item.toets b{color:var(--amber)}
@media (max-width:880px){.week{grid-template-columns:1fr 1fr}.wdag{border-bottom:1px solid var(--lijn)}}
/* uren per vak */
.uren{width:100%;border-collapse:collapse;font-size:14px}
.uren td{padding:9px 0;border-bottom:1px solid var(--lijn);vertical-align:middle}
.uren td.n{text-align:right;color:var(--zacht);width:70px}
.balk{height:6px;border-radius:3px;background:var(--accent);opacity:.75;margin-top:6px}
footer{margin-top:48px;color:var(--zacht);font-size:12px}
@media (max-width:600px){.wrap{padding:24px 18px 40px}h1{font-size:28px}}
@media (prefers-reduced-motion:reduce){*{transition:none!important}}
'@

# Titel zonder herhaling van het vak en zonder "Deze week:" (staat al in de groep/datum)
function Get-KorteTitel([string]$titel) {
    $t = $titel -replace '^Deze week:\s*', ''
    foreach ($v in $vakLijst) {
        $t = $t -replace ('\s*\((' + [regex]::Escape($v.Code) + '\s+)?' + [regex]::Escape($v.Naam) + '\)\s*$'), ''
    }
    return $t.Trim()
}

# Status van een deadline: tekst + kleurklasse
function Get-DeadlineStatus($d) {
    if ($d.Ingediend) { return @('ingediend', 'st-groen') }
    $eind = if ($d.Verloopt) { $d.Verloopt } else { $d.Deadline }
    if ($eind -lt $nu) { return @('verlopen', 'st-rustig') }
    $dagen = [Math]::Max(0, [Math]::Floor(($d.Deadline.Date - $nu.Date).TotalDays))
    $txt = if ($dagen -eq 0) { 'vandaag' } elseif ($dagen -eq 1) { 'morgen' } else { "over $dagen dagen" }
    if ($d.HeleDag -and $d.Deadline -le $nu) { $txt = 'deze week' }
    $cls = if ($dagen -le 2) { 'st-rood' } elseif ($dagen -le 7) { 'st-amber' } else { 'st-rustig' }
    return @($txt, $cls)
}

function Get-TaakHtml($d) {
    $st = Get-DeadlineStatus $d
    $klaar = if ($st[1] -eq 'st-groen' -or $st[0] -eq 'verlopen') { ' klaar' } else { '' }
    $dagTxt = $d.Deadline.ToString('ddd d MMM', $nl)
    $uurTxt = if ($d.HeleDag) {
        if ($d.Verloopt -and ($d.Verloopt.Date - $d.Deadline.Date).TotalDays -gt 1) { 't/m ' + $d.Verloopt.AddDays(-1).ToString('ddd d', $nl) } else { 'hele dag' }
    } else { $d.Deadline.ToString('HH:mm') }
    $bron = if ($d.Soort -eq 'toets') { 'Toets' } elseif ($d.Soort) { 'Uit je agenda' } else { 'Blackboard' }
    "<div class=`"taak$klaar`"><div class=`"wanneer`">$dagTxt<small>$uurTxt</small></div>" +
    "<div class=`"wat`">$(Enc (Get-KorteTitel $d.Titel))<small>$bron</small></div><span class=`"st $($st[1])`">$($st[0])</span></div>"
}

function Get-DeadlinesPerVakHtml($lijst) {
    $volgorde = @($vakLijst | ForEach-Object Naam) + 'Overig'
    $groepen = @($lijst | Group-Object { Get-VakNaamVoor $_ })
    $h = ''
    foreach ($naam in $volgorde) {
        $g = $groepen | Where-Object { $_.Name -eq $naam } | Select-Object -First 1
        if (-not $g) { continue }
        $items = @($g.Group)
        $open = @($items | Where-Object { -not $_.Ingediend -and $(if ($_.Verloopt) { $_.Verloopt } else { $_.Deadline }) -ge $nu }).Count
        $teller = if ($open -eq 1) { '1 open' } elseif ($open) { "$open open" } else { 'niets open' }
        $h += "<div class=`"vakgroep`"><h3>$(Enc $naam)<small>$teller</small></h3>"
        foreach ($d in $items) { $h += Get-TaakHtml $d }
        $h += '</div>'
    }
    return $h
}

# Tijdlijn van vandaag (of 'geen les' + volgende les)
function Get-TijdlijnHtml($rooster) {
    $lessen = @($rooster | Where-Object { $_.Start.Date -eq $nu.Date } | Sort-Object Start)
    $h = "<div class=`"dag`" aria-label=`"Lessen vandaag`">"
    if (-not $lessen) {
        $volgende = $rooster | Where-Object { $_.Start -gt $nu } | Select-Object -First 1
        $h += '<p class="dag-leeg">Vandaag heb je geen les.'
        if ($volgende) { $h += " Je volgende les is $(Enc $volgende.Vak), $($volgende.Start.ToString('dddd d MMMM', $nl)) om $($volgende.Start.ToString('HH:mm'))." }
        return $h + '</p></div>'
    }
    $vanU = [Math]::Min(8, ($lessen | ForEach-Object { $_.Start.Hour } | Measure-Object -Minimum).Minimum)
    $totU = [Math]::Max(18, ($lessen | ForEach-Object { if ($_.Einde.Minute) { $_.Einde.Hour + 1 } else { $_.Einde.Hour } } | Measure-Object -Maximum).Maximum)
    $span = ($totU - $vanU) * 60.0
    $pos = { param($t) [Math]::Round((($t.Hour - $vanU) * 60 + $t.Minute) / $span * 100, 3) }
    $h += '<div class="tl">'
    for ($u = $vanU; $u -le $totU; $u++) {
        $x = [Math]::Round(($u - $vanU) * 60 / $span * 100, 3)
        $h += "<div class=`"tl-raster`" style=`"left:$x%`"></div>"
    }
    $h += '<div class="tl-uren">'
    for ($u = $vanU; $u -le $totU; $u += 2) {
        $x = [Math]::Round(($u - $vanU) * 60 / $span * 100, 3)
        $h += "<span class=`"tl-uur`" style=`"left:$x%`">$($u.ToString('00')):00</span>"
    }
    $h += '</div>'
    foreach ($l in $lessen) {
        $links = & $pos $l.Start; $breed = (& $pos $l.Einde) - $links
        $cls = 'tl-blok'
        if ($l.Start -le $nu -and $l.Einde -gt $nu) { $cls += ' nu' } elseif ($l.Einde -le $nu) { $cls += ' voorbij' }
        $tip = "$($l.Vak), $($l.Start.ToString('HH:mm'))-$($l.Einde.ToString('HH:mm')), $($l.Lokaal), $($l.Docenten)"
        $lokaalKort = Enc ($l.Lokaal -replace '\s*\(.*$', '')
        $tijden = $l.Start.ToString('HH:mm') + '&ndash;' + $l.Einde.ToString('HH:mm')
        $h += "<div class=`"$cls`" style=`"left:calc($links% + 2px);width:calc($breed% - 4px)`" title=`"$(Enc $tip)`">" +
              "<b>$(Enc $l.Vak)</b><span>$tijden, $lokaalKort</span></div>"
    }
    $nuPos = & $pos $nu
    if ($nuPos -ge 0 -and $nuPos -le 100) { $h += "<div class=`"tl-nu`" style=`"left:$nuPos%`" title=`"Nu, $($nu.ToString('HH:mm'))`"></div>" }
    return $h + '</div></div>'
}

function Get-WeekGrid($rooster, $deadlines, $agenda, [datetime]$van) {
    $h = '<div class="week">'
    for ($i = 0; $i -lt 7; $i++) {
        $dag = $van.Date.AddDays($i)
        $cls = if ($dag -eq $nu.Date) { 'wdag vandaag' } else { 'wdag' }
        $h += "<div class=`"$cls`"><h3>$($dag.ToString('ddd', $nl)) <small>$($dag.ToString('d MMM', $nl))</small></h3>"
        $blokken = New-Object System.Collections.ArrayList
        foreach ($d in @($deadlines | Where-Object { $_.Deadline.Date -eq $dag -and -not $_.Ingediend })) {
            [void]$blokken.Add(@($d.Deadline, "<div class=`"item deadline`"><b>Deadline $($d.Deadline.ToString('HH:mm'))</b>$(Enc (Get-KorteTitel $d.Titel))</div>"))
        }
        # meerdaagse items alleen op hun eerste zichtbare dag tonen
        $eersteDag = { param($x) if ($x.Start.Date -lt $van.Date) { $van.Date } else { $x.Start.Date } }
        foreach ($x in @($agenda | Where-Object { ($_.HeleDag -and (& $eersteDag $_) -eq $dag -and $_.Einde.Date -gt $dag) -or (-not $_.HeleDag -and $_.Start.Date -eq $dag) })) {
            $k = if ($x.Soort -eq 'toets') { 'toets' } else { 'deadline' }
            $label = if ($x.Soort -eq 'toets') { 'Toets' } else { 'Deadline' }
            $tijd = if ($x.HeleDag) {
                $laatste = $x.Einde.AddDays(-1)
                if ($laatste.Date -gt $dag) { 't/m ' + $laatste.ToString('ddd', $nl) } else { 'hele dag' }
            } elseif ($x.Soort -eq 'deadline') { $x.Einde.ToString('HH:mm') } else { $x.Start.ToString('HH:mm') }
            $sort = if ($x.HeleDag) { $dag } else { $x.Start }
            [void]$blokken.Add(@($sort, "<div class=`"item $k`"><b>$label $tijd</b>$(Enc (Get-KorteTitel $x.Titel))</div>"))
        }
        foreach ($l in @($rooster | Where-Object { $_.Start.Date -eq $dag })) {
            [void]$blokken.Add(@($l.Start, "<div class=`"item`"><b>$($l.Start.ToString('HH:mm'))&ndash;$($l.Einde.ToString('HH:mm'))</b>$(Enc $l.Vak)</div>"))
        }
        $gesorteerd = @($blokken | ForEach-Object { [pscustomobject]@{ T = $_[0]; Html = $_[1] } } | Sort-Object T)
        foreach ($b in $gesorteerd) { $h += $b.Html }
        if ($blokken.Count -eq 0) { $h += '<div class="vrij">Vrij</div>' }
        $h += '</div>'
    }
    return $h + '</div>'
}

function Get-Pagina([string]$titel, [string]$kop, [string]$samenvatting, [string]$body) {
    @"
<!DOCTYPE html><html lang="nl"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="color-scheme" content="light dark">
<title>$(Enc $titel)</title><style>$css</style></head><body><div class="wrap">
<header class="kop"><div><p class="naam">$(Enc $titel)</p><h1>$(Enc $kop)</h1><p class="samenvatting">$samenvatting</p></div>
<div class="bijgewerkt">Bijgewerkt om $($nu.ToString('HH:mm'))</div></header>
$body
<footer>Rooster uit TimeEdit, deadlines uit je Blackboard-meldingen en je Outlook-agenda.</footer>
</div>
<script>
(function () {
  var SERVER = 'http://127.0.0.1:48765';
  var pagina = /weekoverzicht/.test(location.pathname) ? '/weekoverzicht' : '/';
  if (location.protocol === 'file:') {
    fetch(SERVER + '/', { mode: 'no-cors' }).then(function () { location.replace(SERVER + pagina); }).catch(function () {});
    return;
  }
  setTimeout(function () { location.reload(); }, 15 * 60 * 1000);
  document.addEventListener('click', function (e) {
    var a = e.target.closest && e.target.closest('a.vakknop');
    if (!a) return;
    e.preventDefault();
    if (a.classList.contains('bezig')) return;
    var st = a.querySelector('.status');
    a.classList.add('bezig'); if (st) st.textContent = 'openen...';
    fetch('/open/' + a.getAttribute('data-vak'), { method: 'POST', headers: { 'X-Studie': '1' } })
      .then(function (r) { if (!r.ok) throw new Error(r.status); a.classList.add('ok'); if (st) st.textContent = 'geopend'; })
      .catch(function () { a.classList.add('fout'); if (st) st.textContent = 'lukt niet'; })
      .then(function () { setTimeout(function () { a.classList.remove('bezig', 'ok', 'fout'); if (st) st.textContent = ''; }, 2500); });
  });
})();
</script>
</body></html>
"@
}

function Get-VakknoppenHtml($lessenVandaag) {
    $pad = Join-Path $root 'vakken.json'
    if (-not (Test-Path $pad)) { return '' }
    $vk = Get-Content $pad -Raw -Encoding UTF8 | ConvertFrom-Json
    $codesVandaag = @($lessenVandaag | ForEach-Object Code)
    $h = '<section aria-labelledby="vakken-titel"><h2 id="vakken-titel">Vakken openen<small>cursus, Blackboard en programma''s in een klik</small></h2><div class="vakken">'
    foreach ($v in $vk.Vakken) {
        $wat = @()
        if (@($v.Bladwijzers).Count) { $wat += 'cursus' }
        if ($v.Blackboard) { $wat += 'Blackboard' }
        foreach ($p in $v.Programmas) {
            $n = [IO.Path]::GetFileNameWithoutExtension($p)
            $wat += switch ($n) { 'PacketTracer' { 'Packet Tracer' } 'wt' { 'Terminal' } 'vmware' { 'VMware' } default { $n } }
        }
        $cls = 'vakknop'
        $tip = "Opent: " + ($wat -join ', ')
        if ($codesVandaag -contains $v.Code) { $cls += ' vandaag'; $tip = "Vandaag les. $tip" }
        $h += "<a class=`"$cls`" href=`"studievak:$($v.Id)`" data-vak=`"$($v.Id)`" title=`"$(Enc $tip)`"><span class=`"stip`" aria-hidden=`"true`"></span>$(Enc $v.Naam)<span class=`"status`" aria-live=`"polite`"></span></a>"
    }
    $h += '</div>'
    if ($codesVandaag) { $h += '<p class="vakken-uitleg">Een blauw bolletje betekent dat je vandaag les hebt voor dat vak.</p>' }
    return $h + '</section>'
}

function New-Dashboard($rooster, $info, $agenda) {
    $vandaag = @($rooster | Where-Object { $_.Start.Date -eq $nu.Date })
    $nogTeGaan = @($vandaag | Where-Object { $_.Einde -gt $nu })
    $alle = Get-AlleDeadlines $info $agenda
    $open = @($alle | Where-Object { -not $_.Ingediend -and $(if ($_.Verloopt) { $_.Verloopt } else { $_.Deadline }) -ge $nu })
    $deze7 = @($open | Where-Object { $_.Deadline -lt $nu.Date.AddDays(8) })
    $recent = @($info.Deadlines | Where-Object { $_.Deadline -lt $nu -or $_.Ingediend } | Where-Object { $_.Deadline -ge $nu.AddDays(-7) })
    $ongelezen = @($info.Schoolmail | Where-Object { $_.Ongelezen })

    # samenvatting in gewone zinnen
    $zinnen = @()
    if (-not $vandaag) { $zinnen += 'Vandaag geen les.' }
    elseif (-not $nogTeGaan) { $zinnen += 'Je lessen voor vandaag zitten erop.' }
    elseif ($nogTeGaan.Count -eq 1) { $zinnen += "Nog <b>1 les</b> vandaag." }
    else { $zinnen += "Nog <b>$($nogTeGaan.Count) lessen</b> vandaag." }
    if ($deze7.Count -eq 0) { $zinnen += 'Geen deadlines of toetsen in de komende 7 dagen.' }
    else {
        $eerst = $deze7 | Sort-Object Deadline | Select-Object -First 1
        $wanneer = if ($eerst.HeleDag) { 'vanaf ' + $eerst.Deadline.ToString('dddd', $nl) } else { $eerst.Deadline.ToString('dddd', $nl) + ' om ' + $eerst.Deadline.ToString('HH:mm') }
        $zinnen += "<b>$($deze7.Count)</b> deadlines en toetsen in de komende 7 dagen. Eerst: $(Enc (Get-KorteTitel $eerst.Titel)), $wanneer."
    }
    if ($ongelezen.Count) { $zinnen += "$($ongelezen.Count) ongelezen schoolmail$(if ($ongelezen.Count -gt 1) { 's' })." }

    $b = Get-TijdlijnHtml $rooster
    $b += Get-VakknoppenHtml $vandaag

    $b += '<div class="kolommen"><section aria-labelledby="dl-titel"><h2 id="dl-titel">Deadlines en toetsen<small>per vak</small></h2>'
    if (-not $open -and -not $recent) { $b += '<p class="leeg">Er staan geen deadlines of toetsen open.</p>' }
    else { $b += Get-DeadlinesPerVakHtml ($open + $recent) }
    $b += '</section>'

    $b += "<section aria-labelledby=`"mail-titel`"><h2 id=`"mail-titel`">Schoolmails<small>laatste $($cfg.SchoolmailDagenTonen) dagen</small></h2>"
    $mails = @($info.Schoolmail | Sort-Object @{e = { [bool]$_.Signaal }; Descending = $true }, @{e = 'Ontvangen'; Descending = $true } | Select-Object -First 10)
    if (-not $mails) { $b += '<p class="leeg">Geen nieuwe schoolmails.</p>' }
    foreach ($m in $mails) {
        $cls = if ($m.Ongelezen) { 'mail ongelezen' } else { 'mail' }
        $let = if ($m.Signaal) { "<span class=`"let`" title=`"Bevat: $(Enc $m.Signaal)`">Lijkt belangrijk.</span> " } else { '' }
        $ond = $m.Onderwerp -replace '^\S*\d\S*\s+[^:]{0,60}\(\d{4}-\d{2}[^)]*\):\s*', ''   # vakprefix van Blackboard weglaten
        $b += "<div class=`"$cls`"><span class=`"ond`">$(Enc $ond)</span><small>$let$(Enc $m.Van), $($m.Ontvangen.ToString('ddd d MMM HH:mm', $nl))</small></div>"
    }
    $b += '</section></div>'

    $b += '<section aria-labelledby="week-titel"><h2 id="week-titel">Komende 7 dagen</h2>' + (Get-WeekGrid $rooster $info.Deadlines $agenda $nu.Date) + '</section>'

    $kop = (Get-Culture).TextInfo.ToTitleCase($nu.ToString('dddd', $nl)) + ' ' + $nu.ToString('d MMMM', $nl)
    $pad = Join-Path $root 'dashboard.html'
    [IO.File]::WriteAllText($pad, (Get-Pagina 'Studie-overzicht' $kop ($zinnen -join ' ') $b), (New-Object Text.UTF8Encoding $false))
    return $pad
}

function New-Weekoverzicht($rooster, $info, $agenda) {
    $offset = (8 - [int]$nu.DayOfWeek) % 7
    if ($offset -eq 0 -and $nu.DayOfWeek -ne 'Monday') { $offset = 7 }
    $ma = $nu.Date.AddDays($offset)
    $zo = $ma.AddDays(7)
    $lessen = @($rooster | Where-Object { $_.Start -ge $ma -and $_.Start -lt $zo })
    $alle = Get-AlleDeadlines $info $agenda
    $dls = @($alle | Where-Object { $_.Deadline -ge $ma -and $_.Deadline -lt $zo -and -not $_.Ingediend })
    $openVoor = @($alle | Where-Object { $_.Deadline -ge $nu -and $_.Deadline -lt $ma -and -not $_.Ingediend })
    $uren = [double](($lessen | ForEach-Object { ($_.Einde - $_.Start).TotalHours } | Measure-Object -Sum).Sum)
    $dagenMetLes = @($lessen | ForEach-Object { $_.Start.Date } | Select-Object -Unique).Count

    $zinnen = @()
    if ($lessen) { $zinnen += "<b>$($lessen.Count) lessen</b> op $dagenMetLes dagen, samen $([Math]::Round($uren,1)) uur." }
    else { $zinnen += 'Er staan nog geen lessen in je rooster voor deze week.' }
    if ($dls) { $zinnen += "<b>$($dls.Count)</b> deadlines en toetsen." } else { $zinnen += 'Geen deadlines of toetsen.' }
    if ($openVoor) { $zinnen += "Daarnaast staan er nog $($openVoor.Count) open voor dit weekend." }

    $b = '<section aria-label="Week in een oogopslag" style="margin-top:28px">' + (Get-WeekGrid $rooster $info.Deadlines $agenda $ma) + '</section>'

    $b += '<div class="kolommen"><section aria-labelledby="dl-titel"><h2 id="dl-titel">Deadlines en toetsen<small>per vak</small></h2>'
    if (-not $openVoor -and -not $dls) { $b += '<p class="leeg">Geen deadlines of toetsen deze week.</p>' }
    else { $b += Get-DeadlinesPerVakHtml ($openVoor + $dls) }
    $b += '</section>'

    $b += '<section aria-labelledby="uren-titel"><h2 id="uren-titel">Lesuren per vak</h2>'
    if (-not $lessen) { $b += '<p class="leeg">Nog geen lessen ingepland.</p>' }
    else {
        $groepen = @($lessen | Group-Object Vak | ForEach-Object {
            [pscustomobject]@{ Vak = $_.Name; Uren = [double](($_.Group | ForEach-Object { ($_.Einde - $_.Start).TotalHours } | Measure-Object -Sum).Sum) } } |
            Sort-Object Uren -Descending)
        $max = ($groepen | Measure-Object Uren -Maximum).Maximum
        $b += '<table class="uren">'
        foreach ($g in $groepen) {
            $w = [Math]::Round($g.Uren / $max * 100)
            $b += "<tr><td>$(Enc $g.Vak)<div class=`"balk`" style=`"width:$w%`"></div></td><td class=`"n`">$([Math]::Round($g.Uren,1)) u</td></tr>"
        }
        $b += '</table>'
    }
    $b += '</section></div>'

    $kop = 'Week van ' + $ma.ToString('d MMMM', $nl)
    $pad = Join-Path $root 'weekoverzicht.html'
    [IO.File]::WriteAllText($pad, (Get-Pagina 'Weekoverzicht' $kop ($zinnen -join ' ') $b), (New-Object Text.UTF8Encoding $false))
    return $pad
}

# Opent een pagina via de lokale server (dan werken de vakknoppen); anders als bestand.
function Open-Pagina([string]$bestand, [string]$pad) {
    $tcp = New-Object System.Net.Sockets.TcpClient
    try {
        $ok = $tcp.ConnectAsync('127.0.0.1', 48765).Wait(1500) -and $tcp.Connected
    } catch { $ok = $false } finally { $tcp.Close() }
    if ($ok) { Start-Process "http://127.0.0.1:48765$pad" } else { Start-Process $bestand }
}

# ------------------------------------------------------------------ main --
try {
    $rooster = Get-Rooster
    Write-Log "Rooster: $($rooster.Count) lessen geladen."

    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace('MAPI')
    $info = Get-MailInfo $ns
    Write-Log "Mail: $($info.Deadlines.Count) deadline(s), $($info.Schoolmail.Count) schoolmail(s)."

    if ($cfg.DeadlinesInOutlook -and -not $GeenAgenda) { Sync-DeadlinesNaarAgenda $ns $info.Deadlines }

    $agenda = Get-AgendaItems $ns
    Write-Log "Agenda: $($agenda.Count) eigen deadline(s)/toets(en)/activiteit(en) gevonden."

    $dash = New-Dashboard $rooster $info $agenda
    if ($Modus -eq 'Week') {
        $week = New-Weekoverzicht $rooster $info $agenda
        if (-not $NietOpenen) { Open-Pagina $week '/weekoverzicht' }
    } elseif ($Modus -eq 'Dagelijks' -and -not $NietOpenen) {
        # briefing maximaal 1x per dag openen (ook als je je vaker aanmeldt)
        $vlag = Join-Path $dataDir 'laatste-briefing.txt'
        $vandaag = $nu.ToString('yyyy-MM-dd')
        if (-not (Test-Path $vlag) -or (Get-Content $vlag -Raw).Trim() -ne $vandaag) {
            Open-Pagina $dash '/'
            Set-Content $vlag $vandaag
        }
    } elseif ($Openen -and -not $NietOpenen) {
        Open-Pagina $dash '/'
    }
    Write-Log 'Klaar.'
} catch {
    Write-Log "FOUT: $($_.Exception.Message) (regel $($_.InvocationInfo.ScriptLineNumber))"
    throw
}
