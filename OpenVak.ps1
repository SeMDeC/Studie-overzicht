<#
  Opent alles voor een vak: cursussen uit je Brave-bladwijzermap, de Blackboard-cursus
  en de programma's uit vakken.json.

  Gebruik:
    .\OpenVak.ps1 networks
    .\OpenVak.ps1 studievak:linux        # zo roepen de dashboardknoppen het aan
    .\OpenVak.ps1 -Lijst                  # toon wat elke knop zou openen (opent niets)
#>
param(
    [Parameter(Position = 0)][string]$Vak,
    [switch]$Lijst
)
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$cfg  = Get-Content (Join-Path $root 'vakken.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$log  = Join-Path $root 'data\log.txt'
function Write-Log([string]$m) {
    New-Item -ItemType Directory -Force (Split-Path $log) | Out-Null
    Add-Content $log ('{0:yyyy-MM-dd HH:mm:ss} [OpenVak] {1}' -f (Get-Date), $m) -Encoding UTF8
}

function Get-Bladwijzers {
    # Brave, Chrome of Edge: de eerste browser met een bladwijzermap met de juiste naam wint
    $kandidaten = @(
        'BraveSoftware\Brave-Browser\User Data\Default\Bookmarks',
        'Google\Chrome\User Data\Default\Bookmarks',
        'Microsoft\Edge\User Data\Default\Bookmarks'
    ) | ForEach-Object { Join-Path $env:LOCALAPPDATA $_ } | Where-Object { Test-Path $_ }
    foreach ($bestand in $kandidaten) {
        $lijst = @(Get-BladwijzersUit $bestand)
        if ($lijst.Count) { return $lijst }
    }
    return @()
}

function Get-BladwijzersUit([string]$bestand) {
    $j = Get-Content $bestand -Raw -Encoding UTF8 | ConvertFrom-Json
    $gevonden = New-Object System.Collections.ArrayList
    function Zoek($node, [bool]$binnen) {
        $nuBinnen = $binnen -or ($node.type -eq 'folder' -and $node.name -eq $cfg.BladwijzerMap)
        if ($node.type -eq 'url' -and $binnen) { [void]$gevonden.Add([pscustomobject]@{ Naam = $node.name; Url = $node.url }) }
        foreach ($c in $node.children) { Zoek $c $nuBinnen }
    }
    foreach ($k in 'bookmark_bar', 'other', 'synced') { if ($j.roots.$k) { Zoek $j.roots.$k $false } }
    return @($gevonden)
}

function Get-VakUrls($v, $bladwijzers) {
    $urls = @()
    foreach ($patroon in $v.Bladwijzers) {
        $urls += @($bladwijzers | Where-Object { $_.Naam -like $patroon } | ForEach-Object Url)
    }
    if ($v.Blackboard) { $urls += ($cfg.BlackboardUrl -f $v.Blackboard) }
    return @($urls | Select-Object -Unique)
}

# Pad met joker (bv. "Cisco Packet Tracer*") omzetten naar het echte pad
function Resolve-Programma([string]$p) {
    if ($p -notmatch '[\\/]') { return $p }                     # bv. wt.exe: via PATH
    $hit = Get-Item $p -ErrorAction SilentlyContinue | Sort-Object FullName -Descending | Select-Object -First 1
    if ($hit) { return $hit.FullName }
    return $null
}

function Test-Draait([string]$prog) {
    $naam = [IO.Path]::GetFileNameWithoutExtension($prog)
    if ($naam -eq 'wt') { $naam = 'WindowsTerminal' }
    return [bool](Get-Process -Name $naam -ErrorAction SilentlyContinue)
}

$bladwijzers = Get-Bladwijzers

if ($Lijst) {
    foreach ($v in $cfg.Vakken) {
        "== $($v.Naam)  (studievak:$($v.Id))"
        foreach ($u in (Get-VakUrls $v $bladwijzers)) { '   web : ' + ($u -replace '//[^/@]+@', '//') }   # wachtwoord in URL niet tonen
        foreach ($p in $v.Programmas) { $r = Resolve-Programma $p; '   prog: ' + $(if ($r) { $r } else { "$p  (NIET GEVONDEN)" }) }
    }
    return
}

# "studievak:networks" of "studievak://networks/" -> "networks"
$id = ($Vak -replace '^studievak:(//)?', '').Trim('/ ').ToLower()
$v = $cfg.Vakken | Where-Object { $_.Id -eq $id } | Select-Object -First 1
if (-not $v) { Write-Log "Onbekend vak: '$Vak'"; throw "Onbekend vak: '$Vak'" }

$urls = Get-VakUrls $v $bladwijzers
if ($urls) {
    $browser = if ($cfg.Browser) { Resolve-Programma $cfg.Browser } else { $null }
    if ($browser) { Start-Process -FilePath $browser -ArgumentList (@('--new-window') + $urls) }
    else { foreach ($u in $urls) { Start-Process $u } }          # standaardbrowser
}

foreach ($p in $v.Programmas) {
    $pad = Resolve-Programma $p
    if (-not $pad) { Write-Log "Programma niet gevonden: $p"; continue }
    if (Test-Draait $pad) { continue }
    try { Start-Process -FilePath $pad } catch { Write-Log "Kon $pad niet starten: $($_.Exception.Message)" }
}
Write-Log "$($v.Naam) geopend: $($urls.Count) pagina('s), $(@($v.Programmas).Count) programma('s)."
