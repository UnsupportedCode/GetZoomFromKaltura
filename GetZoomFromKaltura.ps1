# GetZoomFromKaltura.ps1
# Downloads MP4 via ffmpeg using manifestUrl (from "Copy Debug Info" JSON). Prompts before overwriting an existing MP4.
# Then, downloads VTT Subtitles.
# Finally, converts combined VTT subtitles file to SRT format.

# Demo URL: https://kaltura.uga.edu/media/t/1_q8h1i5tv

# Edit $ffmpeg path if necessary.
$ffmpeg = "$env:UserProfile\Desktop\SmartStart\ffmpeg-8.0-full_build\bin\ffmpeg.exe"

function Resolve-Ffmpeg {
    param(
        [string]$Candidate
    )

    # Helper: expand ~ and environment vars, make relative paths absolute
    function Expand-PathLike($p) {
        if (-not $p) { return $null }
        $expanded = [Environment]::ExpandEnvironmentVariables($p)
        if ($expanded.StartsWith('~')) { $expanded = $expanded -replace '^~', $env:USERPROFILE }
        if (-not [System.IO.Path]::IsPathRooted($expanded)) { $expanded = Join-Path (Get-Location) $expanded }
        return $expanded
    }

    # 1) If candidate is provided and points to an existing file, use it
    if ($Candidate -and -not [string]::IsNullOrWhiteSpace($Candidate)) {
        try {
            $candidatePath = Expand-PathLike $Candidate
            if ($candidatePath -and (Test-Path $candidatePath)) {
                $fi = Get-Item $candidatePath -ErrorAction Stop
                if (-not $fi.PSIsContainer) { return $fi.FullName }
            }
        } catch { }
    }

    # 2) Look for ffmpeg* folders adjacent to the script file (same directory as this .ps1)
    try {
        $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
        if ($scriptDir) {
            $entries = Get-ChildItem -LiteralPath $scriptDir -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -imatch '^ffmpeg' }
            foreach ($d in $entries) {
                # common candidate locations inside the folder
                $candidates = @(
                    (Join-Path $d.FullName 'bin\ffmpeg.exe'),
                    (Join-Path $d.FullName 'ffmpeg.exe'),
                    (Join-Path $d.FullName 'bin\ffmpeg'),      # unix-like
                    (Join-Path $d.FullName 'ffmpeg')          # unix-like
                )
                foreach ($c in $candidates) {
                    if (Test-Path $c) {
                        try { $fi = Get-Item $c -ErrorAction Stop; if (-not $fi.PSIsContainer) { return $fi.FullName } } catch {}
                    }
                }
            }
        }
    } catch { }

    # 3) Try to find 'ffmpeg' on PATH via Get-Command
    try {
        $cmd = Get-Command ffmpeg -ErrorAction SilentlyContinue
        if ($cmd) { return $cmd.Source }
    } catch { }

    # 4) Try Get-Command on the original candidate name (handles "ffmpeg.exe" or "ffmpeg")
    if ($Candidate) {
        try {
            $cmd2 = Get-Command $Candidate -ErrorAction SilentlyContinue
            if ($cmd2) { return $cmd2.Source }
        } catch { }
    }

    return $null
}

# Intro text
Write-Host "GetZoomFromKaltura: A small utility to download meeting video & subtitles from Kaltura."
Write-Host " "

# Resolve ffmpeg executable
$resolvedFfmpeg = Resolve-Ffmpeg -Candidate $ffmpeg

if (-not $resolvedFfmpeg) {
    Write-Host ""
    Write-Host "ERROR: ffmpeg executable not found."
    Write-Host ""
    Write-Host "What you can do:"
    Write-Host "- Option A: Edit the script and set the variable $ffmpeg to the full path of your ffmpeg.exe, for example:"
    Write-Host "    `\$ffmpeg = \"C:\\\\tools\\\\ffmpeg\\\\bin\\\\ffmpeg.exe\""
    Write-Host "- Option B: Place an ffmpeg distribution folder next to this script whose name starts with 'ffmpeg' (for example 'ffmpeg-8.0-full_build'). The script will search that folder for bin\\ffmpeg.exe or ffmpeg.exe."
    Write-Host "- Option C: Install ffmpeg and add it to your system PATH so the command 'ffmpeg' is available from PowerShell."
    Write-Host "- Option D: Download a static build from https://ffmpeg.org/ and extract it; then either point \$ffmpeg at the extracted ffmpeg.exe or add its folder to PATH."
    Write-Host ""
    Write-Host "After installing or setting \$ffmpeg, rerun this script."
    Write-Host ""
    exit 1
}

# Use the resolved path from now on
$ffmpeg = $resolvedFfmpeg
Write-Host ("Using ffmpeg: {0}" -f $ffmpeg)
Write-Host " "
Write-Host "Please right-click the video, select 'Copy Debug Info' (JSON) and press Enter."
Read-Host "Press Enter when ready..."
$clipboard = Get-Clipboard
if ([string]::IsNullOrWhiteSpace($clipboard)) { Write-Host "Clipboard empty. Exiting."; exit 1 }

try { $json = $clipboard | ConvertFrom-Json } catch { $json = $null; Write-Host "Clipboard JSON parse failed. Exiting."; exit 1 }

# Build headers (User-Agent + optional decoded Referer)
$ua = if ($json.userAgent) { $json.userAgent } else { "Mozilla/5.0 (Windows NT 10.0; Win64; x64)" }
$refDecoded = $null
if ($json.referrer) {
    try {
        $decoded = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($json.referrer))
        if (-not [string]::IsNullOrWhiteSpace($decoded)) { $refDecoded = $decoded } else { $refDecoded = $json.referrer }
    } catch { $refDecoded = $json.referrer }
}
$wrHeaders = @{ "User-Agent" = $ua }
if ($refDecoded) { $wrHeaders["Referer"] = $refDecoded }
$ffmpegHeaders = "User-Agent: $ua`r`n"
if ($refDecoded) { $ffmpegHeaders += "Referer: $refDecoded`r`n" }

# Input manifest and output base name
$input = if ($json.manifestUrl) { $json.manifestUrl } else { $null }
$defaultBase = if ($json.entryId) { $json.entryId } else { "kaltura_video" }
$baseName = Read-Host "Enter output filename base (default: $defaultBase)"
if ([string]::IsNullOrWhiteSpace($baseName)) { $baseName = $defaultBase }

$folder = Get-Location
$mp4 = Join-Path $folder ("/$baseName/$baseName.mp4")

# Step: Video download with existing-file prompt
$downloadVideo = $true
if ($input) {
    if (Test-Path $mp4) {
        Write-Host ""
        Write-Host ("A file named {0} already exists ({1} bytes)." -f $mp4, (Get-Item $mp4).Length)
        $resp = Read-Host "Re-download and overwrite this file? (Y/N) [default N]"
        if ([string]::IsNullOrWhiteSpace($resp)) { $resp = "N" }
        $resp = $resp.Trim().ToUpperInvariant()
        if ($resp -eq "Y" -or $resp -eq "YES") {
            Write-Host "User chose to re-download and overwrite existing file."
            Remove-Item -Path $mp4 -Force -ErrorAction SilentlyContinue
            $downloadVideo = $true
        } else {
            Write-Host "Skipping video download; keeping existing file."
            $downloadVideo = $false
        }
    } else {
        $downloadVideo = $true
    }

    if ($downloadVideo) {
        Write-Host "`nDownloading MP4 (ffmpeg) to: $mp4"
        Write-Host "URL: "
        Write-Host "$input"
        Write-Host ""
        & $ffmpeg -hide_banner -y -protocol_whitelist "file,http,https,tcp,tls" -headers $ffmpegHeaders -i "$input" -map 0:v -map 0:a -c copy "$mp4"
        if (Test-Path $mp4) {
            Write-Host ("Download finished: {0} ({1} bytes)" -f $mp4, (Get-Item $mp4).Length)
        } else {
            Write-Host "ffmpeg did not produce the expected MP4. Check ffmpeg path, network, and the manifestUrl."
        }
    } else {
        Write-Host "`nUsing existing MP4: $mp4"
    }
} else {
    Write-Host "manifestUrl not found in JSON. Skipping video download."
}

Write-Host "`nDone for $baseName"

# DOWNLOAD SUBTITLES

# Segment base URL (default is the known example -- after which it MUST be made dynamic)
$defaultSegmentBase = "https://cfvod.kaltura.com/api_v3/index.php/service/caption_captionasset/action/serveWebVTT/captionAssetId/1_ov25lmoc/segmentDuration/300/ks/djJ8MTcyNzQxMXxTnTHvSjrPM6YKcX8WUl0nLnS1mDYpwp2-AvYgwypxOJtxeiIUW5P_x2pOxjtefMY0yYsVXAd5ODHpIaXYu0Cj3HZPIkAd9j4ytcYxgiJ65uXVO0gJoVi4Q-BFWQ06XpKkzDUflGTmSE96_I-aElhyFtBd7fvkKQNbh5sTcpMo9UKg6pfWaXPqtkNCH6NirpB6JoUupMZTJywSWJ_Je6fCpIipDauMOm6h2d2rnyIGxWMkH7C_C3W3v2v3i7kxyAUEBjC0SMvCu5LmVZ3W2jzS/version/11/segmentIndex/"
$segBase = $defaultSegmentBase
# Discover subtitles playlist (making it dynamic from manifestUrl)
$segBase = $null
try {
    if (-not $input) { throw "No manifestUrl available to discover subtitles." }
    Write-Host "Fetching master manifest: $input"
    $master = Invoke-WebRequest -Uri $input -Headers $wrHeaders -ErrorAction Stop
    if ($master.Content -is [byte[]]) {
        $masterText = [System.Text.Encoding]::UTF8.GetString($master.Content)
    } else {
        $masterText = $master.Content
    }

    # Find the SUBTITLES URI from an EXT-X-MEDIA line with TYPE=SUBTITLES
    $subsUri = $null
    foreach ($line in $masterText -split "`n") {
        if ($line -match 'EXT-X-MEDIA:.*TYPE=SUBTITLES') {
            # Extract URI="..."; handle both double and single quotes
            if ($line -match 'URI=(?:"([^"]+)"|''([^'']+)'')') {
                $subsUri = if ($matches[1]) { $matches[1] } else { $matches[2] }
                break
            }
        }
    }

    if (-not $subsUri) { throw "No SUBTITLES URI found in master manifest." }

    # Resolve subsUri to absolute URL if relative
    if ($subsUri -notmatch '^[a-zA-Z][a-zA-Z0-9+\-.]*:') {
        # master.RequestUri gives the final absolute URI used; fall back to $input
        $baseUri = $master.BaseResponse.ResponseUri.AbsoluteUri
        $subsUri = ([System.Uri]::new([System.Uri]::new($baseUri), $subsUri)).AbsoluteUri
    }

    Write-Host "Fetching subtitles playlist: $subsUri"
    $subsResp = Invoke-WebRequest -Uri $subsUri -Headers $wrHeaders -ErrorAction Stop
    if ($subsResp.Content -is [byte[]]) {
        try { $subsText = [System.Text.Encoding]::UTF8.GetString($subsResp.Content) }
        catch { $subsText = [System.Text.Encoding]::Default.GetString($subsResp.Content) }
    } else {
        $subsText = $subsResp.Content
    }

    # Determine segment base: find first segment line and remove trailing 'segmentIndex/1.vtt' (or similar)
    $firstSegLine = ($subsText -split "`n" | Where-Object { $_ -and ($_ -notmatch '^#') } | Select-Object -First 1)
    if (-not $firstSegLine) { throw "No segment lines found in subtitles playlist." }

    # If the subs playlist lists absolute URLs, use the directory of those URLs
    if ($firstSegLine -match '^[a-zA-Z][a-zA-Z0-9+\-.]*:') {
        $segUri = ([System.Uri]::new([System.Uri]::new($subsResp.BaseResponse.ResponseUri.AbsoluteUri), $firstSegLine)).AbsoluteUri
        # strip the filename portion (keep trailing slash)
        $segBase = ($segUri -replace '/[^/]+$','') + '/'
    } else {
        # relative path like "segmentIndex/1.vtt" -> base is the subsUri directory + that path up to 'segmentIndex/'
        $subsBase = ($subsResp.BaseResponse.ResponseUri.AbsoluteUri -replace '/[^/]*$','') + '/'
        # if firstSegLine contains a slash, remove the tail after the last slash to produce base that ends with '/'
        $segBase = ([System.Uri]::new([System.Uri]::new($subsBase), $firstSegLine)).AbsoluteUri -replace '/[^/]+$','' 
        if (-not $segBase.EndsWith('/')) { $segBase += '/' }
    }

    Write-Host "Determined segments base: $segBase"
} catch {
    Write-Host "Could not determine subtitle segment base dynamically: $($_.Exception.Message)"
    Write-Host "Falling back to default hardcoded segment base."
    $segBase = $defaultSegmentBase
}


# Ensure trailing slash
if (-not $segBase.EndsWith('/')) { $segBase += '/' }

$vtt = Join-Path $folder ("$baseName/$baseName.vtt")
$srt = Join-Path $folder ("$baseName/$baseName.srt")

# Create temp folder for segments
$tmp = Join-Path $env:TEMP ("kaltura_subs_compact_" + [Guid]::NewGuid().ToString())
New-Item -ItemType Directory -Path $tmp | Out-Null

# Helper: join base URL and relative path into an absolute URL
function Join-Url($base, $path) {
    if ($path -match '^[a-zA-Z][a-zA-Z0-9+\-.]*:') { return $path }
    return ([System.Uri]::new([System.Uri]::new($base), $path)).AbsoluteUri
}

# Determine numeric segment range from the subtitles playlist (if available)
$start = 1
$end = 1
try {
    # Attempt to re-use $subsText if present; otherwise fetch the subs playlist again
    if (-not $subsText) {
        $tmpResp = Invoke-WebRequest -Uri $subsUri -Headers $wrHeaders -ErrorAction Stop
        if ($tmpResp.Content -is [byte[]]) {
            try { $subsText = [System.Text.Encoding]::UTF8.GetString($tmpResp.Content) }
            catch { $subsText = [System.Text.Encoding]::Default.GetString($tmpResp.Content) }
        } else {
            $subsText = $tmpResp.Content
        }
    }

    # Collect segment path lines (non-comment, non-empty)
    $segLines = $subsText -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -and ($_ -notlike '#*') }

    # Extract numeric indices from lines like "segmentIndex/1.vtt" or "1.vtt"
    $indices = @()
    foreach ($l in $segLines) {
        if ($l -match '([0-9]+)\.vtt') { $indices += [int]$matches[1] }
        elseif ($l -match '/([0-9]+)$') { $indices += [int]$matches[1] }
    }

    if ($indices.Count -gt 0) {
        $indices = $indices | Sort-Object
        $start = $indices[0]
        $end = $indices[-1]
    } else {
        # If we couldn't parse numbers, attempt a safe scan by trying sequential downloads until a miss
        $start = 1
        $end = 200
    }

    Write-Host "Subtitle segment index range: $start .. $end"
} catch {
    Write-Host "Failed to determine numeric segment range: $($_.Exception.Message)"
    Write-Host "Using conservative default range 1..200"
    $start = 1
    $end = 200
}

Write-Host "`nDownloading segments $start..$end from: $segBase"
$downloaded = 0
for ($i = $start; $i -le $end; $i++) {
    $url = $segBase + "$i.vtt"
    $out = Join-Path $tmp ("seg{0}.vtt" -f $i)
    Write-Host ("[{0}/{1}] GET {2}" -f ($i - $start + 1), ($end - $start + 1), $url)
    try {
        Invoke-WebRequest -Uri $url -OutFile $out -Headers $wrHeaders -UseBasicParsing -ErrorAction Stop
        if (Test-Path $out) {
            $size = (Get-Item $out).Length
        } else {
            $size = 0
        }
        if ($size -gt 10) {
            Write-Host ("Downloaded {0} bytes" -f $size)
            $downloaded++
        } else {
            Write-Host "Segment downloaded but too small or empty; removing and stopping."
            Remove-Item $out -Force -ErrorAction SilentlyContinue
            break
        }
    } catch {
        Write-Host ("Failed to download segment {0}: {1}" -f $i, $_.Exception.Message)
        break
    }
}

if ($downloaded -eq 0) {
    Write-Host "`nNo segments were downloaded. Cleaning up and exiting."
    Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
    exit 1
}

# Concatenate segments into a single VTT (strip duplicate WEBVTT headers),
# Sorting the files by numeric segment index so order is correct.
$final = ""
foreach ($file in Get-ChildItem -Path $tmp -Filter "seg*.vtt" | Sort-Object {[int]($_.BaseName -replace '\D','')}) {
    $c = Get-Content $file.FullName -Raw
    if ($final.Length -eq 0) { $final += $c } else { $final += "`n" + ($c -replace '^WEBVTT\s*','') }
}
Set-Content -Path $vtt -Value $final -Encoding UTF8

Write-Host ("Concatenated {0} segments into {1} ({2} bytes)" -f $downloaded, $vtt, (Get-Item $vtt).Length)



# CONVERT VTT -> SRT (ffmpeg)
# VTT-to-SRT PowerShell converter (no ffmpeg)
# Input: Path to VTT file; Output: same base name .srt
#param(
#    [string]$vttPath = $vtt
#)

if (-not (Test-Path $vtt)) { Write-Host "VTT not found: $vtt"; exit 1 }

$srtPath = [System.IO.Path]::ChangeExtension($vtt, ".srt")
$content = Get-Content $vtt -Raw

# Remove BOM if present
if ($content.StartsWith([char]0xFEFF)) { $content = $content.Substring(1) }

# Normalize line endings to `n, trim leading/trailing whitespace
$content = ($content -replace "`r`n", "`n" -replace "`r", "`n").Trim("`n")

# Remove any "WEBVTT" header and any header metadata lines up to first blank line
$content = $content -replace '^\s*WEBVTT[^\n]*\n',''
# Also remove header blocks like "NOTE" and "Region" lines that start a block
$content = $content -replace '(?m)^(NOTE|REGION|STYLE)[\s\S]*?(\n{2,}|\z)',''

# Split into blocks by two or more newlines (these are VTT cue blocks)
$blocks = $content -split "`n{2,}"

# Helper to detect a VTT time line; returns cleaned timestamp or $null
function Parse-VttTimeLine($line) {
    # VTT times look like "00:00:09.562 --> 00:00:10.392" optionally with settings after the arrow
    if ($line -match '^\s*([0-9]{1,2}:[0-9]{2}:[0-9]{2}\.[0-9]{3}|[0-9]{1,2}:[0-9]{2}\.[0-9]{3})\s*-->\s*([0-9]{1,2}:[0-9]{2}:[0-9]{2}\.[0-9]{3}|[0-9]{1,2}:[0-9]{2}\.[0-9]{3})') {
        $start = $matches[1]; $end = $matches[2]
        # Ensure both timestamps have hours (VTT can omit hours); normalize to HH:MM:SS.mmm
        function Norm($t) {
            if ($t -match '^[0-9]{1}:[0-9]{2}\.[0-9]{3}$' -or $t -match '^[0-9]{1,2}:[0-9]{2}\.[0-9]{3}$') {
                # mm:ss.mmm -> 00:mm:ss.mmm
                return ("00:" + $t)
            }
            if ($t -match '^[0-9]{1,2}:[0-9]{2}:[0-9]{2}\.[0-9]{3}$') { return $t }
            return $t
        }
        $s = Norm($start); $e = Norm($end)
        # Replace dot milliseconds with comma for SRT
        $s = $s -replace '\.([0-9]{3})$', ',$1'
        $e = $e -replace '\.([0-9]{3})$', ',$1'
        return "$s --> $e"
    }
    return $null
}

$srtLines = New-Object System.Collections.Generic.List[string]
$index = 0
foreach ($block in $blocks) {
    $lines = ($block -split "`n") | ForEach-Object { $_.TrimEnd() } | Where-Object { $_ -ne "" }
    if ($lines.Count -eq 0) { continue }
    # If first line is an ID (non-timestamp), skip it or include as speaker label later
    $first = $lines[0]
    $timeline = $null
    if ($first -match '-->') {
        $timeline = Parse-VttTimeLine($first)
        $textStart = 0
    } else {
        # maybe second line contains time
        if ($lines.Count -gt 1 -and $lines[1] -match '-->') {
            $timeline = Parse-VttTimeLine($lines[1])
            $textStart = 2
        } else {
            # Not a cue block with time; skip
            continue
        }
    }
    if (-not $timeline) { continue }
    $index++
    $srtLines.Add($index.ToString())
    # join text lines (from textStart..end) and strip cue settings and HTML-like tags
    $text = ($lines[$textStart..($lines.Count - 1)] -join " `n")
    # Remove inline cue settings like "<c>" tags, or speaker markers e.g. "Rob Martin:" — we keep speaker text
    $text = $text -replace '<[^>]+>',''
    # Ensure empty lines inside text are preserved as \n
    $srtLines.Add($timeline)
    $srtLines.Add($text)
    $srtLines.Add("")  # blank separator
}

if ($srtLines.Count -eq 0) {
    Write-Host "No cues found in VTT. Inspect first 200 characters of file:"
    $preview = (Get-Content $vtt -Raw).Substring(0,[Math]::Min(200,(Get-Content $vtt -Raw).Length))
    Write-Host $preview
    exit 1
}

# Save SRT
$srtLines | Out-File -FilePath $srt -Encoding UTF8
Write-Host "Created SRT: $srt (cues: $index)"
