# Enriches the catalog with clean official product images pulled from each URL.
# Prefers og:image (color-correct primary variant); falls back to JSON-LD / twitter:image.
# If a site blocks us (e.g. Madewell 403) or has no usable image, keeps the curated
# deck screenshot (which is already the color the owner selected).
#
# Pipeline order:  1) tools\parse-deck.ps1   2) tools\fetch-images.ps1
#
# Params let you re-run for just failures or a single brand without refetching everything.
param(
    [switch]$OnlyMissing,      # only fetch items that don't already have a web image
    [string]$Brand,            # limit to one brand (e.g. -Brand Aritzia)
    [int]$DelayMs = 350        # politeness delay between requests
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$dataDir = Join-Path $root 'data'
$catalogPath = Join-Path $dataDir 'catalog.json'
if (-not (Test-Path $catalogPath)) { throw "Run tools\parse-deck.ps1 first (missing catalog.json)." }

$webDir = Join-Path $dataDir 'images_web'
if (-not (Test-Path $webDir)) { New-Item -ItemType Directory -Path $webDir | Out-Null }

$catalog = Get-Content $catalogPath -Raw | ConvertFrom-Json

# Persist catalog to json + csv (called periodically so progress is never lost)
function Save-Catalog {
    $catalog | ConvertTo-Json -Depth 6 | Set-Content -Path $catalogPath -Encoding UTF8
    $csvPath = Join-Path $dataDir 'catalog.csv'
    $catalog | ForEach-Object {
        [pscustomobject]@{
            id          = $_.id
            brand       = $_.brand
            category    = $_.category
            weather     = $_.weather
            occasions   = ($_.occasions -join '; ')
            image       = $_.image
            imageSource = $_.imageSource
            url         = $_.url
        }
    } | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
}

$headers = @{
    'User-Agent'                = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36'
    'Accept'                    = 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8'
    'Accept-Language'           = 'en-US,en;q=0.9'
    'Upgrade-Insecure-Requests' = '1'
    'sec-ch-ua'                 = '"Chromium";v="125", "Not.A/Brand";v="24"'
    'sec-ch-ua-mobile'          = '?0'
    'sec-ch-ua-platform'        = '"Windows"'
    'Sec-Fetch-Dest'            = 'document'
    'Sec-Fetch-Mode'            = 'navigate'
    'Sec-Fetch-Site'            = 'none'
    'Sec-Fetch-User'            = '?1'
}

# Image-download headers: deliberately DO NOT offer avif, so format-auto CDNs
# (e.g. Aritzia/Cloudinary) return jpg/webp that browsers and the app can display.
$imgHeaders = @{
    'User-Agent'      = $headers['User-Agent']
    'Accept'          = 'image/webp,image/png,image/jpeg,image/*;q=0.8,*/*;q=0.5'
    'Accept-Language' = 'en-US,en;q=0.9'
}

# Normalize an image URL for maximum format compatibility (force jpg on Cloudinary/Quince)
function Get-NormalizedImageUrl([string]$u) {
    if ($u -match '/image/upload/') {
        # Cloudinary (Aritzia): force jpg
        if ($u -match '(^.*?/image/upload/)([^/]*?)(/.*$)') {
            $prefix = $Matches[1]; $trans = $Matches[2]; $rest = $Matches[3]
            if ($trans -match 'f_auto') { $trans = $trans -replace 'f_auto', 'f_jpg' }
            elseif ($trans -match 'f_[a-z0-9]+') { $trans = $trans -replace 'f_[a-z0-9]+', 'f_jpg' }
            elseif ($trans) { $trans = 'f_jpg,' + $trans }
            else { $trans = 'f_jpg' }
            $u = $prefix + $trans + $rest
        }
    }
    # Quince (Contentful): force jpg format param
    $u = $u -replace '(?i)fm=web[p]?', 'fm=jpg'
    return $u
}

# Return the best image URL from page HTML, preferring the color-correct primary (og:image)
function Get-BestImageUrl([string]$html) {
    $og = New-Object System.Collections.ArrayList
    $ld = New-Object System.Collections.ArrayList
    $tw = New-Object System.Collections.ArrayList

    foreach ($m in [regex]::Matches($html, '<meta[^>]+property=["'']og:image(?::secure_url)?["''][^>]+content=["'']([^"'']+)["'']')) { [void]$og.Add($m.Groups[1].Value) }
    foreach ($m in [regex]::Matches($html, '<meta[^>]+content=["'']([^"'']+)["''][^>]+property=["'']og:image["'']')) { [void]$og.Add($m.Groups[1].Value) }
    foreach ($m in [regex]::Matches($html, '<meta[^>]+name=["'']twitter:image["''][^>]+content=["'']([^"'']+)["'']')) { [void]$tw.Add($m.Groups[1].Value) }
    foreach ($m in [regex]::Matches($html, '(?s)<script[^>]+type=["'']application/ld\+json["''][^>]*>(.*?)</script>')) {
        foreach ($im in [regex]::Matches($m.Groups[1].Value, '"image"\s*:\s*("(?<s>[^"]+)"|\[\s*"(?<a>[^"]+)")')) {
            if ($im.Groups['s'].Success) { [void]$ld.Add($im.Groups['s'].Value) }
            elseif ($im.Groups['a'].Success) { [void]$ld.Add($im.Groups['a'].Value) }
        }
    }

    foreach ($list in @($og, $ld, $tw)) {
        foreach ($u in $list) {
            if ($u) {
                $clean = [System.Net.WebUtility]::HtmlDecode($u).Trim()
                $clean = $clean -replace '\\/', '/'
                if ($clean -match '^//') { $clean = 'https:' + $clean }
                if ($clean -match '^http://') { $clean = $clean -replace '^http://', 'https://' }
                if ($clean -match '^https?://') { return $clean }
            }
        }
    }
    return $null
}

# Detect image type from magic bytes; returns extension or $null if not an image
function Get-ImageExt([byte[]]$b) {
    if ($b.Length -lt 12) { return $null }
    if ($b[0] -eq 0xFF -and $b[1] -eq 0xD8 -and $b[2] -eq 0xFF) { return '.jpg' }
    if ($b[0] -eq 0x89 -and $b[1] -eq 0x50 -and $b[2] -eq 0x4E -and $b[3] -eq 0x47) { return '.png' }
    if ($b[0] -eq 0x47 -and $b[1] -eq 0x49 -and $b[2] -eq 0x46) { return '.gif' }
    if ($b[0] -eq 0x52 -and $b[1] -eq 0x49 -and $b[2] -eq 0x46 -and $b[8] -eq 0x57 -and $b[9] -eq 0x45) { return '.webp' }
    # ISO-BMFF 'ftyp' box (avif/heic) at bytes 4-7
    if ($b[4] -eq 0x66 -and $b[5] -eq 0x74 -and $b[6] -eq 0x79 -and $b[7] -eq 0x70) {
        $brand = ( -join ($b[8..11] | ForEach-Object { [char]$_ }))
        if ($brand -match 'avif') { return '.avif' }
        if ($brand -match 'hei|mif1') { return '.heic' }
    }
    return $null
}

function Invoke-WithRetry([scriptblock]$Action, [int]$Tries = 3) {
    for ($i = 1; $i -le $Tries; $i++) {
        try { return & $Action }
        catch {
            if ($i -eq $Tries) { throw }
            Start-Sleep -Milliseconds (250 * $i)
        }
    }
}

$results = @{}     # id -> status string
$total = 0; $done = 0

foreach ($item in $catalog) {
    if ($Brand -and $item.brand -ne $Brand) { continue }
    if ($OnlyMissing -and $item.imageSource -eq 'web') { continue }
    $total++

    $status = 'screenshot'
    $webImage = $null
    $origin = $null
    try {
        $page = Invoke-WithRetry { Invoke-WebRequest -Uri $item.url -Headers $headers -TimeoutSec 30 -MaximumRedirection 5 -UseBasicParsing }
        $imgUrl = Get-BestImageUrl $page.Content
        if ($imgUrl) {
            $imgUrl = Get-NormalizedImageUrl $imgUrl
            $tmpFile = Join-Path $env:TEMP "img_dl_$($item.id)"
            Invoke-WithRetry { Invoke-WebRequest -Uri $imgUrl -Headers $imgHeaders -TimeoutSec 30 -MaximumRedirection 5 -UseBasicParsing -OutFile $tmpFile }
            $bytes = [System.IO.File]::ReadAllBytes($tmpFile)
            $ext = Get-ImageExt $bytes
            if ($ext -and $bytes.Length -gt 3000) {
                # clear any previous web file for this id (ext may differ)
                Get-ChildItem (Join-Path $webDir "$($item.id).*") -ErrorAction SilentlyContinue | Remove-Item -Force
                $dest = Join-Path $webDir "$($item.id)$ext"
                [System.IO.File]::WriteAllBytes($dest, $bytes)
                $webImage = "images_web/$($item.id)$ext"
                $origin = $imgUrl
                $status = 'web'
            }
            else { $status = 'bad-image' }
            Remove-Item $tmpFile -ErrorAction SilentlyContinue
        }
        else { $status = 'no-meta' }
    }
    catch {
        if ($_.Exception.Message -match '403') { $status = 'blocked-403' }
        elseif ($_.Exception.Message -match '404') { $status = 'notfound-404' }
        else { $status = 'error' }
    }

    if ($status -eq 'web') {
        $item | Add-Member -NotePropertyName imageWeb -NotePropertyValue $webImage -Force
        $item | Add-Member -NotePropertyName imageOrigin -NotePropertyValue $origin -Force
        $item | Add-Member -NotePropertyName imageSource -NotePropertyValue 'web' -Force
        $item.image = $webImage
    }
    else {
        # keep curated screenshot
        if (-not $item.imageSource) { $item | Add-Member -NotePropertyName imageSource -NotePropertyValue 'screenshot' -Force }
    }

    $results[$item.id] = $status
    $done++
    Write-Host ("[{0,3}/{1}] #{2,-3} {3,-16} {4}" -f $done, $total, $item.id, $item.brand, $status)
    if ($done % 15 -eq 0) { Save-Catalog }   # checkpoint progress
    Start-Sleep -Milliseconds $DelayMs
}

# --- save catalog ----------------------------------------------------------

Save-Catalog

# --- comparison review sheet (web image vs curated screenshot) -------------

$sb = New-Object System.Text.StringBuilder
[void]$sb.Append(@'
<!doctype html><html><head><meta charset="utf-8"><title>Catalog images</title>
<style>
 body{font-family:system-ui,Segoe UI,Arial,sans-serif;margin:24px;background:#fafafa;color:#222}
 h1{font-size:20px} .grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(230px,1fr));gap:14px}
 .card{background:#fff;border:1px solid #e2e2e2;border-radius:10px;padding:10px}
 .imgs{display:flex;gap:8px} .imgs figure{margin:0;flex:1;text-align:center}
 .imgs img{width:100%;height:210px;object-fit:contain;background:#f4f4f4;border-radius:6px;border:1px solid #e6e6e6}
 .imgs figcaption{font-size:11px;color:#888;margin-top:3px}
 .meta{font-size:12px;margin-top:8px} .meta b{color:#111}
 .src-web{color:#178a4c;font-weight:600} .src-screenshot{color:#b26b00;font-weight:600}
 .meta a{color:#2a6;word-break:break-all}
 .filter{margin:12px 0}
</style></head><body>
<h1>Catalog images &mdash; official vs. screenshot</h1>
<p>Left = clean image fetched from the product page (color from the URL). Right = your deck screenshot.
The app uses whichever is marked <span class="src-web">web</span>; items marked
<span class="src-screenshot">screenshot</span> kept your curated image.</p>
<div class="grid">
'@)
foreach ($it in $catalog) {
    if ($Brand -and $it.brand -ne $Brand) { continue }
    $web = if ($it.imageWeb) { $it.imageWeb } else { '' }
    $shot = "images/$($it.id).png"
    $srcClass = if ($it.imageSource -eq 'web') { 'src-web' } else { 'src-screenshot' }
    [void]$sb.Append("<div class='card'><div class='imgs'>")
    if ($web) { [void]$sb.Append("<figure><img src='$web' loading='lazy'><figcaption>official</figcaption></figure>") }
    [void]$sb.Append("<figure><img src='$shot' loading='lazy'><figcaption>screenshot</figcaption></figure>")
    [void]$sb.Append("</div><div class='meta'><b>#$($it.id) &middot; $($it.brand)</b> &middot; <span class='$srcClass'>$($it.imageSource)</span><br>$($it.category) &middot; $($it.weather)<br><a href='$($it.url)' target='_blank'>open product</a></div></div>")
}
[void]$sb.Append("</div></body></html>")
Set-Content -Path (Join-Path $dataDir 'review.html') -Value $sb.ToString() -Encoding UTF8

# --- summary ---------------------------------------------------------------

Write-Host "`n===== Summary ====="
$results.Values | Group-Object | Sort-Object Count -Descending | ForEach-Object { "  {0,-14} {1}" -f $_.Name, $_.Count }
$webCount = @($catalog | Where-Object { $_.imageSource -eq 'web' }).Count
Write-Host "`nItems using official web image: $webCount / $($catalog.Count)"
Write-Host "`nBy brand (web vs screenshot):"
$catalog | Group-Object brand | Sort-Object Name | ForEach-Object {
    $w = @($_.Group | Where-Object { $_.imageSource -eq 'web' }).Count
    "  {0,-18} web={1,-3} screenshot={2}" -f $_.Name, $w, ($_.Count - $w)
}
Write-Host "`nReview: data\review.html"
