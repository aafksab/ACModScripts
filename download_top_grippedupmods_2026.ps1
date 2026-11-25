<#
.SYNOPSIS
    Downloads the Top 25 Assetto Corsa mods of 2026 from Gripped Up Mods.
.DESCRIPTION
    Visits each of the 25 mod pages listed on https://grippedupmods.com/top-mods/this-year,
    extracts its numeric download_id, posts to https://grippedupmods.com/dwnld_cf to mint
    a fresh 15-minute Cloudflare R2 presigned URL, and streams each archive to disk.
    Skips already-downloaded files so runs are resumable.
.PARAMETER OutputDir
    Destination directory for downloaded mod archives. Defaults to .\GrippedUpMods_2026.
.PARAMETER WhatIfOnly
    If specified, resolves download_id and mints/prints the R2 URLs without downloading files.
.PARAMETER DelaySeconds
    Pause between requests (default: 1 second) to avoid rate limiting.
.EXAMPLE
    .\download_top_mods_2026.ps1 -WhatIfOnly
.EXAMPLE
    .\download_top_mods_2026.ps1 -OutputDir "D:\AssettoCorsa\Mods\Top2026"
#>

[CmdletBinding()]
param(
    [string]$OutputDir = (Join-Path -Path (Get-Location) -ChildPath "GrippedUpMods_2026"),
    [switch]$WhatIfOnly,
    [int]$DelaySeconds = 1
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Ensure TLS 1.2+ on Windows PowerShell 5.1
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13
}
catch {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
}

$Session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$Session.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"

$Endpoint = "https://grippedupmods.com/dwnld_cf"
$UseCurl  = $null -ne (Get-Command "curl.exe" -ErrorAction SilentlyContinue)

$Mods = @(
    @{ Rank = 1;  Name = "SUPERDRIFT Street BMW Pack";               Type = "Car Pack"; Url = "https://grippedupmods.com/downloads/car-packs/superdrift-street-bmw-pack" }
    @{ Rank = 2;  Name = "Gravy Garage V2";                          Type = "Car Pack"; Url = "https://grippedupmods.com/downloads/car-packs/gravy-garage-v2" }
    @{ Rank = 3;  Name = "Minato Garage V1";                         Type = "Car Pack"; Url = "https://grippedupmods.com/downloads/car-packs/minato-garage-v1" }
    @{ Rank = 4;  Name = "N-Style V1.0";                             Type = "Car Pack"; Url = "https://grippedupmods.com/downloads/car-packs/n-style-v10" }
    @{ Rank = 5;  Name = "BDC Street V6";                            Type = "Car Pack"; Url = "https://grippedupmods.com/downloads/car-packs/bdc-street-v6" }
    @{ Rank = 6;  Name = "NDP Makime Touge";                         Type = "Track";    Url = "https://grippedupmods.com/downloads/tracks/ndp-makime-touge" }
    @{ Rank = 7;  Name = "NoVIbe Street Pack";                       Type = "Car Pack"; Url = "https://grippedupmods.com/downloads/car-packs/novibe-street-pack" }
    @{ Rank = 8;  Name = "Death Wish Garage V2";                     Type = "Car Pack"; Url = "https://grippedupmods.com/downloads/car-packs/death-wish-garage-v2" }
    @{ Rank = 9;  Name = "Matsuri Mayhem Drift Legends";             Type = "Car Pack"; Url = "https://grippedupmods.com/downloads/car-packs/matsuri-mayhem-drift-legends" }
    @{ Rank = 10; Name = "FZC Underground";                          Type = "Track";    Url = "https://grippedupmods.com/downloads/tracks/fzc-underground" }
    @{ Rank = 11; Name = "NDP Kuragari Touge";                       Type = "Track";    Url = "https://grippedupmods.com/downloads/tracks/ndp-kuragari-touge" }
    @{ Rank = 12; Name = "Superdrift Low HP";                        Type = "Car Pack"; Url = "https://grippedupmods.com/downloads/car-packs/superdrift-low-hp" }
    @{ Rank = 13; Name = "VDC Public Car Pack 5.0";                  Type = "Car Pack"; Url = "https://grippedupmods.com/downloads/car-packs/vdc-public-car-pack-50" }
    @{ Rank = 14; Name = "2025 Formula Drift Car Pack V1";           Type = "Car Pack"; Url = "https://grippedupmods.com/downloads/car-packs/2025-formula-drift-car-pack-v1" }
    @{ Rank = 15; Name = "ORIGINAL34 Style Pack 1";                  Type = "Car Pack"; Url = "https://grippedupmods.com/downloads/car-packs/original34-style-pack-1" }
    @{ Rank = 16; Name = "SD061 S14 Kouki Street";                   Type = "Car";      Url = "https://grippedupmods.com/downloads/cars/sd061-s14-kouki-street" }
    @{ Rank = 17; Name = "1997 Toyota Chaser JZX100 Missile HHRC";   Type = "Car";      Url = "https://grippedupmods.com/downloads/cars/1997-toyota-chaser-jzx100-missile-hhrc" }
    @{ Rank = 18; Name = "BDC Street V4";                            Type = "Car Pack"; Url = "https://grippedupmods.com/downloads/car-packs/bdc-street-v4" }
    @{ Rank = 19; Name = "Minato Fukuroi";                           Type = "Track";    Url = "https://grippedupmods.com/downloads/tracks/minato-fukuroi" }
    @{ Rank = 20; Name = "Tykkimaen Moottorirata";                   Type = "Track";    Url = "https://grippedupmods.com/downloads/tracks/tykkimaen-moottorirata" }
    @{ Rank = 21; Name = "Kaanaa Center";                            Type = "Track";    Url = "https://grippedupmods.com/downloads/tracks/kaanaa-center" }
    @{ Rank = 22; Name = "Door Hunterz V1 Street Car Pack";          Type = "Car Pack"; Url = "https://grippedupmods.com/downloads/car-packs/door-hunterz-v1-street-car-pack" }
    @{ Rank = 23; Name = "NForce Drift Series";                      Type = "Car Pack"; Url = "https://grippedupmods.com/downloads/car-packs/nforce-drift-series" }
    @{ Rank = 24; Name = "Adam LZ F22 HGK BBC";                      Type = "Car";      Url = "https://grippedupmods.com/downloads/cars/adam-lz-f22-hgk-bbc" }
    @{ Rank = 25; Name = "Mitsuishi Forest Park";                    Type = "Track";    Url = "https://grippedupmods.com/downloads/tracks/mitsuishi-forest-park" }
)

function Find-DownloadId {
    param([string]$Html)

    $patterns = @(
        # data-download-id="179" or data-download_id="179"
        '(?i)data-download[_-]?id\s*=\s*["'']?(\d+)["'']?',
        # Hidden/form input: name="download_id" value="179" (either attribute order)
        '(?i)<input[^>]*\bname\s*=\s*["'']download_id["''][^>]*\bvalue\s*=\s*["''](\d+)["'']',
        '(?i)<input[^>]*\bvalue\s*=\s*["''](\d+)["''][^>]*\bname\s*=\s*["'']download_id["'']',
        # Inline JS / fetch body: download_id=179 or download_id: 179
        '(?i)\bdownload_id["'']?\s*[:=]\s*["'']?(\d+)["'']?',
        # Download button with data-id="179"
        '(?is)<(?:button|a)[^>]*(?:download|dwnld)[^>]*\bdata-id\s*=\s*["''](\d+)["'']',
        '(?is)<(?:button|a)[^>]*\bdata-id\s*=\s*["''](\d+)["''][^>]*(?:download|dwnld)'
    )

    foreach ($pat in $patterns) {
        $m = [regex]::Match($Html, $pat)
        if ($m.Success) {
            return $m.Groups[1].Value
        }
    }
    return $null
}

function Extract-SignedUrl {
    param([string]$Body)

    if ([string]::IsNullOrWhiteSpace($Body)) { return $null }
    $trimmed = $Body.Trim()

    # 1. Try JSON response (e.g. {"url": "..."} or {"download_url": "..."})
    if ($trimmed.StartsWith("{")) {
        try {
            $obj = $trimmed | ConvertFrom-Json
            foreach ($prop in @("url", "download_url", "downloadUrl", "signed_url", "signedUrl", "link", "file", "data")) {
                if ($obj.PSObject.Properties.Name -contains $prop) {
                    $val = $obj.$prop
                    if ($val -is [string] -and $val -match '^https?://') {
                        return $val
                    }
                }
            }
            # Scan all string properties as fallback
            foreach ($p in $obj.PSObject.Properties) {
                if ($p.Value -is [string] -and $p.Value -match '^https?://') {
                    return $p.Value
                }
            }
        }
        catch { }
    }

    # 2. Plain text or escaped URL inside response body
    $unescaped = $trimmed -replace '\\/', '/'
    $m = [regex]::Match($unescaped, 'https?://[^\s"''<>]+')
    if ($m.Success) {
        return [System.Net.WebUtility]::HtmlDecode($m.Value)
    }

    return $null
}

function Get-FileNameFromSignedUrl {
    param([string]$Url, [string]$FallbackName)

    # Prefer filename from response-content-disposition query param
    $decodedUrl = [Uri]::UnescapeDataString($Url)
    $cdMatch = [regex]::Match($decodedUrl, 'filename\s*=\s*"?([^"&;]+)"?')
    if ($cdMatch.Success) {
        $name = $cdMatch.Groups[1].Value.Trim()
        if (-not [string]::IsNullOrWhiteSpace($name)) {
            return ($name -replace '[<>:"/\\|?*]', '_')
        }
    }

    # Fallback to URL path leaf
    $uri = [Uri]$Url
    $leaf = [System.IO.Path]::GetFileName($uri.AbsolutePath)
    if (-not [string]::IsNullOrWhiteSpace($leaf) -and $leaf -match '\.') {
        return ([Uri]::UnescapeDataString($leaf) -replace '[<>:"/\\|?*]', '_')
    }

    return (($FallbackName -replace '[^\w\-\.]+', '_') + ".zip")
}

if (-not $WhatIfOnly) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

$Downloaded = [System.Collections.Generic.List[string]]::new()
$Skipped    = [System.Collections.Generic.List[string]]::new()
$Failed     = [System.Collections.Generic.List[string]]::new()

Write-Host "Processing $($Mods.Count) mods from Gripped Up Mods (Top Mods 2026)..." -ForegroundColor Cyan
if (-not $WhatIfOnly) {
    Write-Host "Output directory : $OutputDir" -ForegroundColor Gray
    Write-Host "Downloader       : $(if ($UseCurl) { 'curl.exe (streaming)' } else { 'Invoke-WebRequest' })`n" -ForegroundColor Gray
}

foreach ($mod in $Mods) {
    $label = "[#$($mod.Rank)] $($mod.Name)"
    Write-Host "$label" -ForegroundColor White

    try {
        # Step 1: GET mod page (also populates PHPSESSID in $Session)
        $pageResp = Invoke-WebRequest -UseBasicParsing -Uri $mod.Url -WebSession $Session
        $downloadId = Find-DownloadId -Html $pageResp.Content

        if (-not $downloadId) {
            Write-Warning "  Could not find download_id on $($mod.Url)"
            $Failed.Add("$label | Missing download_id | $($mod.Url)")
            continue
        }

        Write-Host "  download_id : $downloadId" -ForegroundColor DarkGray

        # Step 2: POST /dwnld_cf to mint fresh presigned R2 URL
        $postHeaders = @{
            "accept"          = "*/*"
            "accept-language" = "en-US,en;q=0.9"
            "origin"          = "https://grippedupmods.com"
            "referer"         = $mod.Url
        }

        $apiResp = Invoke-WebRequest -UseBasicParsing `
            -Uri $Endpoint `
            -Method "POST" `
            -WebSession $Session `
            -Headers $postHeaders `
            -ContentType "application/x-www-form-urlencoded" `
            -Body "download_id=$downloadId"

        $signedUrl = Extract-SignedUrl -Body $apiResp.Content

        if (-not $signedUrl) {
            $preview = if ($apiResp.Content.Length -gt 160) { $apiResp.Content.Substring(0, 160) } else { $apiResp.Content }
            Write-Warning "  /dwnld_cf did not return a URL. Response: $preview"
            $Failed.Add("$label | No URL from /dwnld_cf (id=$downloadId) | $($mod.Url)")
            continue
        }

        $fileName = Get-FileNameFromSignedUrl -Url $signedUrl -FallbackName $mod.Name
        Write-Host "  File        : $fileName" -ForegroundColor DarkGray

        if ($WhatIfOnly) {
            Write-Host "  Signed URL  : $signedUrl" -ForegroundColor DarkCyan
            continue
        }

        $destPath = Join-Path -Path $OutputDir -ChildPath $fileName
        if (Test-Path -LiteralPath $destPath) {
            Write-Host "  Already exists, skipping." -ForegroundColor Yellow
            $Skipped.Add($fileName)
            continue
        }

        # Step 3: Stream archive directly to disk (fresh 15-min R2 signature)
        $partPath = "$destPath.part"
        if (Test-Path -LiteralPath $partPath) {
            Remove-Item -LiteralPath $partPath -Force
        }

        Write-Host "  Downloading..." -ForegroundColor Green
        if ($UseCurl) {
            & curl.exe -L -f --progress-bar -o $partPath $signedUrl
            if ($LASTEXITCODE -ne 0) {
                throw "curl.exe exited with code $LASTEXITCODE"
            }
        }
        else {
            $prevProgress = $ProgressPreference
            $ProgressPreference = "SilentlyContinue"
            try {
                Invoke-WebRequest -UseBasicParsing -Uri $signedUrl -OutFile $partPath
            }
            finally {
                $ProgressPreference = $prevProgress
            }
        }

        Move-Item -LiteralPath $partPath -Destination $destPath -Force
        $sizeMB = [Math]::Round((Get-Item -LiteralPath $destPath).Length / 1MB, 2)
        Write-Host "  Saved ($sizeMB MB)" -ForegroundColor Green
        $Downloaded.Add($fileName)
    }
    catch {
        Write-Warning "  Failed: $($_.Exception.Message)"
        $Failed.Add("$label | Error: $($_.Exception.Message) | $($mod.Url)")
    }

    if ($DelaySeconds -gt 0) {
        Start-Sleep -Seconds $DelaySeconds
    }
}

Write-Host "`n--- Summary ---" -ForegroundColor Cyan
Write-Host "Downloaded : $($Downloaded.Count)" -ForegroundColor Green
Write-Host "Skipped    : $($Skipped.Count)" -ForegroundColor Yellow
Write-Host "Failed     : $($Failed.Count)" -ForegroundColor Red

if ($Failed.Count -gt 0 -and -not $WhatIfOnly) {
    $failFile = Join-Path -Path $OutputDir -ChildPath "failed_downloads.txt"
    $Failed | Set-Content -LiteralPath $failFile -Encoding UTF8
    Write-Host "Saved failed items to: $failFile" -ForegroundColor Yellow
}