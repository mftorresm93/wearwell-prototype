# Organizes the catalog for the outfit app:
#  - fixes mis-categorized items (the old "Work clothes" category -> real garment types)
#  - normalizes weather naming (Summer -> Hot)
#  - adds a `slot` field = the item's role in an outfit (Top/Bottom/Dress/Layer/Shoes/Bag/Accessory)
#  - sorts by slot -> category -> weather -> brand
#  - refreshes taxonomy.json and rebuilds review.html grouped by slot -> category
# Operates on existing data only — does NOT touch images or refetch anything.
#
# Pipeline order: parse-deck.ps1 -> fetch-images.ps1 -> finalize-images.ps1 -> organize-catalog.ps1

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$dataDir = Join-Path $root 'data'
$catalogPath = Join-Path $dataDir 'catalog.json'
if (-not (Test-Path $catalogPath)) { throw "Missing catalog.json." }

$catalog = Get-Content $catalogPath -Raw | ConvertFrom-Json

# Weather buckets and display order
$weatherOrder = @('All year', 'Cold', 'Hot')

# Slot = the item's role when assembling an outfit
$slotOrder = @('Top', 'Layer', 'Bottom', 'Dress', 'Shoes', 'Bag', 'Accessory')
$slotOfCategory = @{
    'Shirts'         = 'Top';      'T-shirts' = 'Top'; 'Going-out tops' = 'Top'; 'Sweaters' = 'Top'
    'Blazers'        = 'Layer';    'Jackets'  = 'Layer'; 'Vests' = 'Layer'
    'Jeans'          = 'Bottom';   'Pants'    = 'Bottom'; 'Skirts' = 'Bottom'
    'Dresses'        = 'Dress'
    'Shoes'          = 'Shoes';    'Sneakers' = 'Shoes'; 'Booties' = 'Shoes'; 'Sandals' = 'Shoes'
    'Bags'           = 'Bag'
    'Belts'          = 'Accessory'
}

# Category display order within a slot
$categoryOrder = @(
    'Shirts', 'T-shirts', 'Going-out tops', 'Sweaters',
    'Blazers', 'Jackets', 'Vests',
    'Jeans', 'Pants', 'Skirts',
    'Dresses',
    'Shoes', 'Sneakers', 'Booties', 'Sandals',
    'Bags',
    'Belts'
)

# Fix the old "Work clothes" bucket: those are real garments tagged for Work (occasion).
$categoryFixById = @{
    150 = 'T-shirts'   # cashmere tee
    151 = 'Sweaters'   # organic cotton sweater tee
    152 = 'Vests'      # yak crewneck sweater vest
    153 = 'Sweaters'   # cashmere short-sleeve polo
    154 = 'Dresses'    # jersey knotted midi dress
    155 = 'Dresses'    # cashmere sleeveless midi sweater dress
}

function Rank($value, $order) {
    $i = [array]::IndexOf($order, $value)
    if ($i -lt 0) { return [int]::MaxValue }
    return $i
}

foreach ($item in $catalog) {
    # normalize weather naming
    if ($item.weather -eq 'Summer') { $item.weather = 'Hot' }
    # fix mis-categorized items
    if ($categoryFixById.ContainsKey([int]$item.id)) { $item.category = $categoryFixById[[int]$item.id] }
    # assign outfit slot
    $slot = $slotOfCategory[$item.category]
    if (-not $slot) { $slot = 'Other' }
    if ($item.PSObject.Properties.Name -contains 'slot') { $item.slot = $slot }
    else { $item | Add-Member -NotePropertyName slot -NotePropertyValue $slot }
}

# Sort: slot -> category -> weather -> brand -> id
$sorted = $catalog | Sort-Object `
    @{ Expression = { Rank $_.slot $slotOrder } }, `
    @{ Expression = { Rank $_.category $categoryOrder } }, `
    @{ Expression = { $_.category } }, `
    @{ Expression = { Rank $_.weather $weatherOrder } }, `
    @{ Expression = { $_.brand } }, `
    @{ Expression = { [int]$_.id } }

# --- save catalog (clean, ordered fields) ----------------------------------

$out = foreach ($it in $sorted) {
    $o = [ordered]@{
        id          = $it.id
        brand       = $it.brand
        slot        = $it.slot
        category    = $it.category
        weather     = $it.weather
        occasions   = $it.occasions
        image       = $it.image
        imageSource = $it.imageSource
        url         = $it.url
    }
    if ($it.PSObject.Properties.Name -contains 'price') { $o.price = $it.price }
    if ($it.PSObject.Properties.Name -contains 'currency') { $o.currency = $it.currency }
    if ($it.PSObject.Properties.Name -contains 'imageOrigin') { $o.imageOrigin = $it.imageOrigin }
    [pscustomobject]$o
}
$out | ConvertTo-Json -Depth 6 | Set-Content -Path $catalogPath -Encoding UTF8

$csvPath = Join-Path $dataDir 'catalog.csv'
$out | ForEach-Object {
    [pscustomobject]@{
        id          = $_.id
        brand       = $_.brand
        slot        = $_.slot
        category    = $_.category
        weather     = $_.weather
        occasions   = ($_.occasions -join '; ')
        price       = $_.price
        image       = $_.image
        imageSource = $_.imageSource
        url         = $_.url
    }
} | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

# --- taxonomy --------------------------------------------------------------

$presentCats = @($categoryOrder | Where-Object { $out.category -contains $_ })
$extraCats = @($out | ForEach-Object { $_.category } | Sort-Object -Unique | Where-Object { $categoryOrder -notcontains $_ })
$taxonomy = [ordered]@{
    weathers   = $weatherOrder
    occasions  = @('Work', 'Casual', 'Going out', 'Beach')
    slots      = $slotOrder
    categories = @($presentCats + $extraCats)
    brands     = @($out | ForEach-Object { $_.brand } | Sort-Object -Unique)
}
$taxonomy | ConvertTo-Json -Depth 5 | Set-Content -Path (Join-Path $dataDir 'taxonomy.json') -Encoding UTF8

# --- grouped review sheet (slot -> category) -------------------------------

$sb = New-Object System.Text.StringBuilder
[void]$sb.Append(@'
<!doctype html><html><head><meta charset="utf-8"><title>Catalog</title>
<style>
 body{font-family:system-ui,Segoe UI,Arial,sans-serif;margin:24px;background:#fafafa;color:#222}
 h1{font-size:22px} h2{font-size:20px;margin:30px 0 4px;border-bottom:3px solid #222;padding-bottom:4px}
 h3{font-size:14px;color:#666;margin:16px 0 8px;text-transform:uppercase;letter-spacing:.04em}
 .grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:14px}
 .card{background:#fff;border:1px solid #e2e2e2;border-radius:10px;padding:10px}
 .card img{width:100%;height:230px;object-fit:contain;background:#f4f4f4;border-radius:6px;border:1px solid #e6e6e6}
 .meta{font-size:12px;margin-top:8px} .meta b{color:#111}
 .badge{display:inline-block;font-size:11px;padding:1px 7px;border-radius:10px;background:#eee;color:#444;margin-left:4px}
 .w-Allyear{background:#e7f0ff;color:#2456b8} .w-Cold{background:#e6f4f7;color:#1f7a8c} .w-Hot{background:#fdeee0;color:#b26b00}
 .src-web{color:#178a4c;font-weight:600} .src-screenshot{color:#b26b00;font-weight:600}
 .meta a{color:#2a6;word-break:break-all}
 .toc{font-size:13px;color:#444;margin:8px 0 4px} .toc a{color:#2a6;margin-right:12px;text-decoration:none;font-weight:600}
</style></head><body>
<h1>Catalog &middot; by outfit slot</h1>
'@)

# table of contents (by slot)
[void]$sb.Append("<div class='toc'>")
foreach ($slot in $slotOrder) {
    $n = @($out | Where-Object { $_.slot -eq $slot }).Count
    if ($n -eq 0) { continue }
    [void]$sb.Append("<a href='#$slot'>$slot ($n)</a>")
}
[void]$sb.Append("</div>")

foreach ($slot in $slotOrder) {
    $slotItems = @($out | Where-Object { $_.slot -eq $slot })
    if ($slotItems.Count -eq 0) { continue }
    [void]$sb.Append("<h2 id='$slot'>$slot &middot; $($slotItems.Count)</h2>")
    $catsHere = @($categoryOrder | Where-Object { $slotItems.category -contains $_ })
    foreach ($cat in $catsHere) {
        $catItems = @($slotItems | Where-Object { $_.category -eq $cat })
        if ($catItems.Count -eq 0) { continue }
        [void]$sb.Append("<h3>$cat &middot; $($catItems.Count)</h3><div class='grid'>")
        foreach ($it in $catItems) {
            $srcClass = if ($it.imageSource -eq 'web') { 'src-web' } else { 'src-screenshot' }
            $wClass = 'w-' + ($it.weather -replace '\s', '')
            $occ = ($it.occasions -join ', ')
            [void]$sb.Append("<div class='card'><img src='$($it.image)' loading='lazy'><div class='meta'><b>#$($it.id) &middot; $($it.brand)</b> <span class='badge $wClass'>$($it.weather)</span><br>$occ<br><span class='$srcClass'>$($it.imageSource)</span> &middot; <a href='$($it.url)' target='_blank'>open product</a></div></div>")
        }
        [void]$sb.Append("</div>")
    }
}
[void]$sb.Append("</body></html>")
Set-Content -Path (Join-Path $dataDir 'review.html') -Value $sb.ToString() -Encoding UTF8

# --- summary ---------------------------------------------------------------

Write-Host "Organized $($out.Count) items by slot -> category -> weather -> brand.`n"
Write-Host "Outfit slots:"
foreach ($slot in $slotOrder) {
    $slotItems = @($out | Where-Object { $_.slot -eq $slot })
    if ($slotItems.Count -eq 0) { continue }
    $cats = ($slotItems | Group-Object category | Sort-Object { Rank $_.Name $categoryOrder } | ForEach-Object { "$($_.Name)=$($_.Count)" }) -join ', '
    "  {0,-9} {1,3}   ({2})" -f $slot, $slotItems.Count, $cats
}
Write-Host "`nUpdated: catalog.json, catalog.csv, taxonomy.json, review.html"
