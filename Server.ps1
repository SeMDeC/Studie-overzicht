<#
  Kleine lokale webserver voor het Studie-overzicht.
  - Alleen bereikbaar op deze laptop: http://127.0.0.1:48765/
  - GET  /                  -> dashboard.html
  - GET  /weekoverzicht     -> weekoverzicht.html
  - POST /open/<vak>        -> OpenVak.ps1 <vak>  (alleen met header X-Studie: 1,
                               zodat andere websites dit niet kunnen aansturen)
  Wordt gestart door de geplande taak "Vakknoppen-server" bij aanmelden.
#>
$ErrorActionPreference = 'Continue'
$root = $PSScriptRoot
$port = 48765
$log  = Join-Path $root 'data\log.txt'
function Write-Log([string]$m) { Add-Content $log ('{0:yyyy-MM-dd HH:mm:ss} [Server] {1}' -f (Get-Date), $m) -Encoding UTF8 }

$listener = New-Object System.Net.Sockets.TcpListener ([System.Net.IPAddress]::Loopback, $port)
try { $listener.Start() } catch { Write-Log "Poort $port al in gebruik, server draait waarschijnlijk al."; return }
Write-Log "Gestart op http://127.0.0.1:$port/"

function Send-Antwoord($stream, [int]$code, [string]$status, [byte[]]$body, [string]$type = 'text/plain; charset=utf-8') {
    $kop = "HTTP/1.1 $code $status`r`nContent-Type: $type`r`nContent-Length: $($body.Length)`r`nCache-Control: no-store`r`nX-Content-Type-Options: nosniff`r`nConnection: close`r`n`r`n"
    $kb = [Text.Encoding]::ASCII.GetBytes($kop)
    $stream.Write($kb, 0, $kb.Length)
    if ($body.Length) { $stream.Write($body, 0, $body.Length) }
}
function Get-Tekst([string]$s) { [Text.Encoding]::UTF8.GetBytes($s) }

while ($true) {
    $client = $listener.AcceptTcpClient()
    try {
        $stream = $client.GetStream()
        $stream.ReadTimeout = 3000
        $reader = New-Object System.IO.StreamReader($stream, [Text.Encoding]::ASCII, $false, 8192, $true)
        $eerste = $reader.ReadLine()
        if (-not $eerste) { continue }
        $hdr = @{}
        while ($line = $reader.ReadLine()) {
            $i = $line.IndexOf(':')
            if ($i -gt 0) { $hdr[$line.Substring(0, $i).Trim().ToLower()] = $line.Substring($i + 1).Trim() }
        }
        $delen = $eerste.Split(' ')
        $methode = $delen[0]; $pad = ($delen[1] -split '\?')[0]

        # alleen aanvragen naar dit adres (beschermt tegen DNS-rebinding)
        if (@("127.0.0.1:$port", "localhost:$port") -notcontains $hdr['host']) {
            Send-Antwoord $stream 403 'Forbidden' (Get-Tekst 'Niet toegestaan'); continue
        }

        if ($methode -eq 'GET' -and ($pad -eq '/' -or $pad -eq '/dashboard' -or $pad -eq '/weekoverzicht')) {
            $bestand = if ($pad -eq '/weekoverzicht') { 'weekoverzicht.html' } else { 'dashboard.html' }
            $vol = Join-Path $root $bestand
            if (Test-Path $vol) { Send-Antwoord $stream 200 'OK' ([IO.File]::ReadAllBytes($vol)) 'text/html; charset=utf-8' }
            else { Send-Antwoord $stream 404 'Not Found' (Get-Tekst "$bestand bestaat nog niet") }
        }
        elseif ($methode -eq 'POST' -and $pad -match '^/open/([a-z0-9]+)$' -and $hdr['x-studie'] -eq '1') {
            $vak = $matches[1]
            Start-Process -FilePath "$env:WINDIR\System32\conhost.exe" -WindowStyle Hidden `
                -ArgumentList "--headless powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$(Join-Path $root 'OpenVak.ps1')`" $vak"
            Send-Antwoord $stream 200 'OK' (Get-Tekst "ok")
        }
        else {
            Send-Antwoord $stream 404 'Not Found' (Get-Tekst 'Niet gevonden')
        }
    } catch {
        Write-Log "Fout: $($_.Exception.Message)"
    } finally {
        $client.Close()
    }
}
