# Finalizes images into a single canonical data/images/ folder:
#   - web items  -> use the fetched official image (images_web/<id>.<ext>)
#   - blocked items (Madewell/Zara/UO/Free People) -> keep their curated screenshot
# Then removes images_web/ and the images_all/ review dump, and rewrites review.html
# to show the single final image per product.
#
# Pipeline order: parse-deck.ps1 -> fetch-images.ps1 -> finalize-images.ps1

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$dataDir = Join-Path $root 'data'
$catalogPath = Join-Path $dataDir 'catalog.json'
if (-not (Test-Path $catalogPath)) { throw "Missing catalog.json. Run parse-deck.ps1 then fetch-images.ps1 first." }

$imagesDir = Join-Path $dataDir 'images'
$webDir = Join-Path $dataDir 'images_web'
$allDir = Join-Path $dataDir 'images_all'

# Guard: finalize consumes images_web/. If it's gone, this was already run — refuse,
# so we don't repoint web items at non-existent .png files. Re-run fetch-images.ps1 first.
if (-not (Test-Path $webDir)) {
    throw "data\images_web not found. Finalize already ran (or fetch didn't). Re-run tools\fetch-images.ps1 before finalizing."
}

$catalog = Get-Content $catalogPath -Raw | ConvertFrom-Json

$movedWeb = 0; $keptShot = 0; $missing = 0

foreach ($item in $catalog) {
    $id = $item.id
    if ($item.imageSource -eq 'web') {
        $src = Get-ChildItem (Join-Path $webDir "$id.*") -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($src) {
            # remove any stale images/<id>.* then place the official image there
            Get-ChildItem (Join-Path $imagesDir "$id.*") -ErrorAction SilentlyContinue | Remove-Item -Force
            $destName = "$id$($src.Extension)"
            Copy-Item $src.FullName (Join-Path $imagesDir $destName) -Force
            $item.image = "images/$destName"
            $movedWeb++
        }
        else {
            $missing++
            Write-Host "WARN: web image missing for #$id ($($item.brand)) — keeping screenshot"
            $item.imageSource = 'screenshot'
            $item.image = "images/$id.png"
        }
    }
    else {
        # keep curated screenshot
        if (Test-Path (Join-Path $imagesDir "$id.png")) {
            $item.image = "images/$id.png"
            $keptShot++
        }
        else {
            $missing++
            Write-Host "WARN: screenshot missing for #$id ($($item.brand))"
        }
    }
    # drop now-redundant field
    if ($item.PSObject.Properties['imageWeb']) { $item.PSObject.Properties.Remove('imageWeb') }
}

# Remove any screenshots that are no longer referenced (web items replaced them)
$keep = @{}
foreach ($item in $catalog) { $keep[[System.IO.Path]::GetFileName($item.image)] = $true }
Get-ChildItem (Join-Path $imagesDir '*') -ErrorAction SilentlyContinue | ForEach-Object {
    if (-not $keep.ContainsKey($_.Name)) { Remove-Item $_.FullName -Force }
}

# Drop the intermediate folders
foreach ($d in @($webDir, $allDir)) { if (Test-Path $d) { Remove-Item $d -Recurse -Force } }

# --- save catalog ----------------------------------------------------------

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

# --- final review sheet (single image per product) -------------------------

$sb = New-Object System.Text.StringBuilder
[void]$sb.Append(@'
<!doctype html><html><head><meta charset="utf-8"><title>Catalog</title>
<style>
 body{font-family:system-ui,Segoe UI,Arial,sans-serif;margin:24px;background:#fafafa;color:#222}
 h1{font-size:20px}
 .grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(190px,1fr));gap:14px}
 .card{background:#fff;border:1px solid #e2e2e2;border-radius:10px;padding:10px}
 .card img{width:100%;height:240px;object-fit:contain;background:#f4f4f4;border-radius:6px;border:1px solid #e6e6e6}
 .meta{font-size:12px;margin-top:8px} .meta b{color:#111}
 .src-web{color:#178a4c;font-weight:600} .src-screenshot{color:#b26b00;font-weight:600}
 .meta a{color:#2a6;word-break:break-all}
</style></head><body>
<h1>Catalog</h1>
<p>Final image used by the app for each product.</p>
<div class="grid">
'@)
foreach ($it in $catalog) {
    $srcClass = if ($it.imageSource -eq 'web') { 'src-web' } else { 'src-screenshot' }
    [void]$sb.Append("<div class='card'><img src='$($it.image)' loading='lazy'><div class='meta'><b>#$($it.id) &middot; $($it.brand)</b> &middot; <span class='$srcClass'>$($it.imageSource)</span><br>$($it.category) &middot; $($it.weather)<br><a href='$($it.url)' target='_blank'>open product</a></div></div>")
}
[void]$sb.Append("</div></body></html>")
Set-Content -Path (Join-Path $dataDir 'review.html') -Value $sb.ToString() -Encoding UTF8

# --- summary ---------------------------------------------------------------

$imgCount = (Get-ChildItem (Join-Path $imagesDir '*') -ErrorAction SilentlyContinue | Measure-Object).Count
$webFinal = @($catalog | Where-Object { $_.imageSource -eq 'web' }).Count
$shotFinal = @($catalog | Where-Object { $_.imageSource -eq 'screenshot' }).Count
Write-Host "`nOfficial web images placed: $webFinal"
Write-Host "Curated screenshots kept:   $shotFinal"
if ($missing) { Write-Host "Missing (check warnings):   $missing" }
Write-Host "Files now in data\images:   $imgCount  (expected $($catalog.Count))"
Write-Host "Removed data\images_web and data\images_all"
