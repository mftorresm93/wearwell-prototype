# Parses PRYECTICO!!!!.pptx into a structured catalog (data/catalog.json + data/catalog.csv)
# Walks slides in order, tracking current weather section and garment category,
# then emits one record per product URL with brand + category + weather + occasion tags.

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$pptx = Join-Path $root 'PRYECTICO!!!!.pptx'
if (-not (Test-Path $pptx)) { throw "Deck not found: $pptx" }

# Copy (deck may be locked by PowerPoint) and extract
$copy = Join-Path $env:TEMP 'pptx_copy_parse.pptx'
Copy-Item $pptx $copy -Force
$tmp = Join-Path $env:TEMP 'pptx_extract_parse'
if (Test-Path $tmp) { Remove-Item $tmp -Recurse -Force }
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory($copy, $tmp)

$slideFiles = Get-ChildItem "$tmp\ppt\slides\*.xml" |
    Sort-Object { [int]($_.BaseName -replace '\D', '') }

# --- helpers ---------------------------------------------------------------

function Get-Brand([string]$url) {
    try { $h = ([uri]$url).Host } catch { return 'Unknown' }
    $h = $h -replace '^www\.', ''
    $map = @{
        'aritzia.com'          = 'Aritzia'
        'quince.com'           = 'Quince'
        'everlane.com'         = 'Everlane'
        'madewell.com'         = 'Madewell'
        'zara.com'             = 'Zara'
        'thereformation.com'   = 'Reformation'
        'maje.com'             = 'Maje'
        'us.maje.com'          = 'Maje'
        'loefflerrandall.com'  = 'Loeffler Randall'
        'veronicabeard.com'    = 'Veronica Beard'
        'urbanoutfitters.com'  = 'Urban Outfitters'
        'champssports.com'     = 'Champs Sports'
        'freepeople.com'       = 'Free People'
        'verafiedny.com'       = 'Verafied NY'
        'songmontofficial.com' = 'Songmont'
    }
    foreach ($k in $map.Keys) { if ($h -eq $k -or $h.EndsWith(".$k")) { return $map[$k] } }
    # fallback: second-level domain, title-cased
    $parts = $h.Split('.')
    if ($parts.Count -ge 2) { $sld = $parts[$parts.Count - 2] } else { $sld = $h }
    return (Get-Culture).TextInfo.ToTitleCase($sld)
}

# Map a title-slide phrase to a normalized garment category (handles Spanish)
function Get-Category([string]$text) {
    $t = $text.ToLower()
    $rules = [ordered]@{
        'going out top' = 'Going-out tops'
        'camisa'        = 'Shirts'
        'blazer'        = 'Blazers'
        'belt'          = 'Belts'
        'sueter'        = 'Sweaters'
        'sweater'       = 'Sweaters'
        'chaqueta'      = 'Jackets'
        'jacket'        = 'Jackets'
        'vest'          = 'Vests'
        'skirt'         = 'Skirts'
        'work dress'    = 'Dresses'
        'dress'         = 'Dresses'
        'work clothes'  = 'Work clothes'
        'pant'          = 'Pants'
        'jean'          = 'Jeans'
        't-shirt'       = 'T-shirts'
        'sneaker'       = 'Sneakers'
        'bootie'        = 'Booties'
        'shoe'          = 'Shoes'
        'bag'           = 'Bags'
    }
    foreach ($k in $rules.Keys) { if ($t.Contains($k)) { return $rules[$k] } }
    return $null
}

$occasionMap = @{ 'WORK' = 'Work'; 'CASUAL' = 'Casual'; 'GOING OUT' = 'Going out' }

# Return center-relevant geometry (EMU) for a shape/picture node, or $null
function Get-OffExt($node) {
    $xfrm = $node.SelectSingleNode(".//*[local-name()='xfrm']")
    if (-not $xfrm) { return $null }
    $off = $xfrm.SelectSingleNode("*[local-name()='off']")
    $ext = $xfrm.SelectSingleNode("*[local-name()='ext']")
    if (-not $off -or -not $ext) { return $null }
    [pscustomobject]@{
        x  = [double]$off.x;  y  = [double]$off.y
        cx = [double]$ext.cx; cy = [double]$ext.cy
    }
}

# --- image output dirs -----------------------------------------------------

$dataDir = Join-Path $root 'data'
if (-not (Test-Path $dataDir)) { New-Item -ItemType Directory -Path $dataDir | Out-Null }
$imagesDir = Join-Path $dataDir 'images'       # chosen image per product: <id>.png
$allDir = Join-Path $dataDir 'images_all'      # every slide image (for manual review/fix)
foreach ($d in @($imagesDir, $allDir)) {
    if (Test-Path $d) { Remove-Item $d -Recurse -Force }
    New-Item -ItemType Directory -Path $d | Out-Null
}

# --- walk slides -----------------------------------------------------------

$weather = $null
$category = $null
$records = New-Object System.Collections.ArrayList
$reviewSlides = New-Object System.Collections.ArrayList
$idCounter = 1

foreach ($file in $slideFiles) {
    $xml = [xml](Get-Content $file.FullName)
    $runs = @($xml.SelectNodes("//*[local-name()='t']") | ForEach-Object { $_.'#text' } |
        Where-Object { $_ -and $_.Trim().Length -gt 0 } | ForEach-Object { $_.Trim() })
    if ($runs.Count -eq 0) { continue }

    $joined = ($runs -join ' ')

    # URLs by paragraph (robust against a URL split across multiple <t> runs)
    $urls = New-Object System.Collections.ArrayList
    foreach ($p in $xml.SelectNodes("//*[local-name()='p']")) {
        $ptxt = (($p.SelectNodes(".//*[local-name()='t']") | ForEach-Object { $_.'#text' }) -join '').Trim()
        if ($ptxt -match '^https?://') { [void]$urls.Add($ptxt) }
    }
    $urls = @($urls)

    # Section divider (weather)
    if ($urls.Count -eq 0) {
        switch -Regex ($joined.ToUpper()) {
            '^ALL YEAR$'     { $weather = 'All year'; continue }
            '^SUMMER$'       { $weather = 'Hot';      continue }
            '^COLD WEATHER$' { $weather = 'Cold';     continue }
        }
        # Title slide -> category (also refine weather from Spanish/English hints)
        $cat = Get-Category $joined
        if ($cat) {
            $category = $cat
            if ($joined -match '(?i)verano') { $weather = 'Hot' }
            elseif ($joined -match '(?i)cold weather') { $weather = 'Cold' }
            elseif ($joined -match '(?i)all year') { if (-not $weather) { $weather = 'All year' } }
        }
        continue
    }

    # Content slide: gather occasion tags present on this slide
    $occ = New-Object System.Collections.ArrayList
    foreach ($tag in $occasionMap.Keys) {
        if ($runs -contains $tag) { [void]$occ.Add($occasionMap[$tag]) }
    }
    $occArr = @($occ | Sort-Object -Unique)

    $slideNo = [int]($file.BaseName -replace '\D', '')

    # Map relationship ids -> embedded media files for this slide
    $relsPath = Join-Path $file.DirectoryName "_rels\$($file.Name).rels"
    $rIdToMedia = @{}
    if (Test-Path $relsPath) {
        $relsXml = [xml](Get-Content $relsPath)
        foreach ($rel in $relsXml.Relationships.Relationship) {
            if ($rel.Type -like '*/image') {
                $rIdToMedia[$rel.Id] = Join-Path $tmp ('ppt\' + ($rel.Target -replace '^\.\./', ''))
            }
        }
    }

    # Collect pictures (center geometry + copy to images_all for review)
    $pics = New-Object System.Collections.ArrayList
    $picIdx = 0
    foreach ($pic in $xml.SelectNodes("//*[local-name()='pic']")) {
        $g = Get-OffExt $pic
        $blip = $pic.SelectSingleNode(".//*[local-name()='blip']")
        if (-not $blip) { continue }
        $embed = ($blip.Attributes | Where-Object { $_.LocalName -eq 'embed' } | Select-Object -First 1).Value
        $media = $rIdToMedia[$embed]
        if (-not $media -or -not (Test-Path $media)) { continue }
        $picIdx++
        $allName = "slide${slideNo}_$picIdx.png"
        Copy-Item $media (Join-Path $allDir $allName) -Force
        if ($g) { $pcx = $g.x + $g.cx / 2; $pcy = $g.y + $g.cy / 2 } else { $pcx = $null; $pcy = $null }
        [void]$pics.Add([pscustomobject]@{
            cx     = $pcx
            cy     = $pcy
            media  = $media
            allRel = "images_all/$allName"
            used   = $false
        })
    }

    # Anchor each URL (distribute multiple URLs vertically within their text box)
    $anchors = @{}
    foreach ($sp in $xml.SelectNodes("//*[local-name()='sp']")) {
        $spUrls = New-Object System.Collections.ArrayList
        foreach ($p in $sp.SelectNodes(".//*[local-name()='p']")) {
            $ptxt = (($p.SelectNodes(".//*[local-name()='t']") | ForEach-Object { $_.'#text' }) -join '').Trim()
            if ($ptxt -match '^https?://') { [void]$spUrls.Add($ptxt) }
        }
        if ($spUrls.Count -eq 0) { continue }
        $g = Get-OffExt $sp
        for ($k = 0; $k -lt $spUrls.Count; $k++) {
            if ($g) {
                $anchors[$spUrls[$k]] = [pscustomobject]@{
                    x = $g.x + $g.cx / 2
                    y = $g.y + $g.cy * (($k + 0.5) / $spUrls.Count)
                }
            }
        }
    }

    # Greedy nearest-unique assignment of pictures to URLs
    $urlImage = @{}
    $pairs = New-Object System.Collections.ArrayList
    foreach ($u in $urls) {
        $a = $anchors[$u]
        if ($a -and $null -ne $a.x) {
            foreach ($pic in $pics) {
                if ($null -eq $pic.cx) { continue }
                $dx = $a.x - $pic.cx; $dy = $a.y - $pic.cy
                [void]$pairs.Add([pscustomobject]@{ url = $u; pic = $pic; d = [math]::Sqrt($dx * $dx + $dy * $dy) })
            }
        }
    }
    foreach ($pair in ($pairs | Sort-Object d)) {
        if ($urlImage.ContainsKey($pair.url) -or $pair.pic.used) { continue }
        $urlImage[$pair.url] = $pair.pic
        $pair.pic.used = $true
    }
    # Fallback for any URL still unmatched: take remaining pictures in order
    foreach ($u in $urls) {
        if ($urlImage.ContainsKey($u)) { continue }
        $free = $pics | Where-Object { -not $_.used } | Select-Object -First 1
        if ($free) { $urlImage[$u] = $free; $free.used = $true }
    }

    $reviewProducts = New-Object System.Collections.ArrayList
    foreach ($u in $urls) {
        $imgRel = $null
        if ($urlImage.ContainsKey($u)) {
            $destName = "$idCounter.png"
            Copy-Item $urlImage[$u].media (Join-Path $imagesDir $destName) -Force
            $imgRel = "images/$destName"
        }
        $rec = [ordered]@{
            id        = $idCounter
            brand     = (Get-Brand $u)
            category  = $category
            weather   = $weather
            occasions = $occArr
            image     = $imgRel
            url       = $u
        }
        [void]$records.Add([pscustomobject]$rec)
        [void]$reviewProducts.Add([pscustomobject]@{ id = $idCounter; url = $u; brand = $rec.brand; image = $imgRel })
        $idCounter++
    }

    [void]$reviewSlides.Add([pscustomobject]@{
        slide    = $slideNo
        category = $category
        weather  = $weather
        images   = @($pics | ForEach-Object { $_.allRel })
        products = @($reviewProducts)
    })
}

# --- output ----------------------------------------------------------------

$jsonPath = Join-Path $dataDir 'catalog.json'
$records | ConvertTo-Json -Depth 5 | Set-Content -Path $jsonPath -Encoding UTF8

$csvPath = Join-Path $dataDir 'catalog.csv'
$records | ForEach-Object {
    [pscustomobject]@{
        id        = $_.id
        brand     = $_.brand
        category  = $_.category
        weather   = $_.weather
        occasions = ($_.occasions -join '; ')
        image     = $_.image
        url       = $_.url
    }
} | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

# Controlled vocabulary for the app UI (Beach included; no items tagged yet)
$taxonomy = [ordered]@{
    weathers   = @('All year', 'Cold', 'Hot')
    occasions  = @('Work', 'Casual', 'Going out', 'Beach')
    categories = @($records | ForEach-Object { $_.category } | Sort-Object -Unique)
    brands     = @($records | ForEach-Object { $_.brand } | Sort-Object -Unique)
}
$taxPath = Join-Path $dataDir 'taxonomy.json'
$taxonomy | ConvertTo-Json -Depth 5 | Set-Content -Path $taxPath -Encoding UTF8

# --- review contact sheet --------------------------------------------------

$sb = New-Object System.Text.StringBuilder
[void]$sb.Append(@'
<!doctype html><html><head><meta charset="utf-8"><title>Catalog image review</title>
<style>
 body{font-family:system-ui,Segoe UI,Arial,sans-serif;margin:24px;background:#fafafa;color:#222}
 h1{font-size:20px} .slide{background:#fff;border:1px solid #e2e2e2;border-radius:10px;padding:16px;margin:18px 0}
 .slide h2{font-size:14px;margin:0 0 10px;color:#555;font-weight:600}
 .prod{display:flex;gap:12px;align-items:center;padding:8px 0;border-top:1px dashed #eee}
 .prod img{width:90px;height:120px;object-fit:cover;border-radius:6px;border:1px solid #ddd;background:#f0f0f0}
 .prod .meta{font-size:13px} .prod .meta b{color:#111} .prod a{color:#2a6;word-break:break-all;font-size:12px}
 .cands{display:flex;flex-wrap:wrap;gap:6px;margin-top:10px}
 .cands img{width:70px;height:92px;object-fit:cover;border-radius:5px;border:1px solid #ccc}
 .cands .lbl{font-size:11px;color:#999;width:100%}
</style></head><body>
<h1>Catalog image review</h1>
<p>Each product shows the auto-matched photo. If one looks wrong, pick the correct file from the slide's images below and rename it to the product id in <code>data/images/</code>.</p>
'@)
foreach ($s in $reviewSlides) {
    [void]$sb.Append("<div class='slide'><h2>Slide $($s.slide) &middot; $($s.category) &middot; $($s.weather)</h2>")
    foreach ($p in $s.products) {
        $img = if ($p.image) { $p.image } else { '' }
        [void]$sb.Append("<div class='prod'><img src='$img' loading='lazy'><div class='meta'><b>#$($p.id) &middot; $($p.brand)</b><br><a href='$($p.url)' target='_blank'>$($p.url)</a></div></div>")
    }
    if ($s.images.Count -gt 0) {
        [void]$sb.Append("<div class='cands'><span class='lbl'>All $($s.images.Count) image(s) on this slide:</span>")
        foreach ($im in $s.images) { [void]$sb.Append("<img src='$im' loading='lazy' title='$im'>") }
        [void]$sb.Append("</div>")
    }
    [void]$sb.Append("</div>")
}
[void]$sb.Append("</body></html>")
$reviewPath = Join-Path $dataDir 'review.html'
$sb.ToString() | Set-Content -Path $reviewPath -Encoding UTF8

# --- summary ---------------------------------------------------------------

$withImg = @($records | Where-Object { $_.image }).Count
Write-Host "Total items: $($records.Count)  |  with image: $withImg  |  without: $($records.Count - $withImg)"
Write-Host "`nBy weather:"
$records | Group-Object weather | Sort-Object Name | ForEach-Object { "  {0,-10} {1}" -f $_.Name, $_.Count }
Write-Host "`nBy category:"
$records | Group-Object category | Sort-Object Name | ForEach-Object { "  {0,-16} {1}" -f $_.Name, $_.Count }
Write-Host "`nBy brand:"
$records | Group-Object brand | Sort-Object Count -Descending | ForEach-Object { "  {0,-18} {1}" -f $_.Name, $_.Count }
Write-Host "`nItems with no category: $(@($records | Where-Object { -not $_.category }).Count)"
Write-Host "Items with no occasions: $(@($records | Where-Object { $_.occasions.Count -eq 0 }).Count)"
Write-Host "`nWrote:"
Write-Host "  $jsonPath"
Write-Host "  $csvPath"
Write-Host "  $taxPath"
Write-Host "  $reviewPath  (open in a browser to verify image matches)"
Write-Host "  $imagesDir\  ($withImg chosen images)"
Write-Host "  $allDir\  (all extracted images for manual fixes)"
