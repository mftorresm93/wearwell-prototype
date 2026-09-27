<#
  add-item.ps1
  Add one product to the catalog by pasting its URL.

  It fetches the product page, derives the brand from the domain, guesses the
  category (you can override), pulls the official product image and the price,
  assigns the next free ID, and appends a clean record to data/catalog.json.
  Nothing existing is touched.

  EXAMPLES:
    # Simplest - let it guess the category, then review the tags after:
    .\tools\add-item.ps1 -Url "https://www.aritzia.com/us/en/product/.../123.html"

    # Set the tags up front:
    .\tools\add-item.ps1 -Url "https://..." -Category Sweaters -Occasions Work,Casual -Weather "All year"

    # See what it WOULD add without writing anything:
    .\tools\add-item.ps1 -Url "https://..." -DryRun

  After adding, fix any tag with edit-item.ps1 and review the catalog with review-catalog.ps1.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Url,
    [ValidateSet('Shirts', 'T-shirts', 'Going-out tops', 'Sweaters', 'Blazers', 'Jackets', 'Vests',
        'Jeans', 'Pants', 'Skirts', 'Dresses', 'Shoes', 'Sneakers', 'Booties', 'Sandals', 'Bags', 'Belts')]
    [string]$Category,
    [string[]]$Occasions,
    [ValidateSet('All year', 'Cold', 'Hot')]
    [string]$Weather = 'All year',
    [switch]$NoImage,
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$dataDir = Join-Path $root 'data'
$catalogPath = Join-Path $dataDir 'catalog.json'
$imagesDir = Join-Path $dataDir 'images'
if (-not (Test-Path $catalogPath)) { throw "Missing catalog.json." }
if (-not (Test-Path $imagesDir)) { New-Item -ItemType Directory -Path $imagesDir | Out-Null }

$headers = @{
    'User-Agent'      = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0 Safari/537.36'
    'Accept'          = 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
    'Accept-Language' = 'en-US,en;q=0.9'
}

# Category -> outfit slot (must match organize-catalog.ps1)
$slotOfCategory = @{
    'Shirts' = 'Top'; 'T-shirts' = 'Top'; 'Going-out tops' = 'Top'; 'Sweaters' = 'Top'
    'Blazers' = 'Layer'; 'Jackets' = 'Layer'; 'Vests' = 'Top'
    'Jeans' = 'Bottom'; 'Pants' = 'Bottom'; 'Skirts' = 'Bottom'
    'Dresses' = 'Dress'
    'Shoes' = 'Shoes'; 'Sneakers' = 'Shoes'; 'Booties' = 'Shoes'; 'Sandals' = 'Shoes'
    'Bags' = 'Bag'
    'Belts' = 'Accessory'
}

function Get-Brand([string]$url) {
    try { $h = ([uri]$url).Host } catch { return 'Unknown' }
    $h = $h -replace '^www\.', ''
    $map = @{
        'aritzia.com' = 'Aritzia'; 'quince.com' = 'Quince'; 'everlane.com' = 'Everlane'
        'madewell.com' = 'Madewell'; 'zara.com' = 'Zara'; 'thereformation.com' = 'Reformation'
        'maje.com' = 'Maje'; 'us.maje.com' = 'Maje'; 'loefflerrandall.com' = 'Loeffler Randall'
        'veronicabeard.com' = 'Veronica Beard'; 'urbanoutfitters.com' = 'Urban Outfitters'
        'champssports.com' = 'Champs Sports'; 'freepeople.com' = 'Free People'
        'verafiedny.com' = 'Verafied NY'; 'songmontofficial.com' = 'Songmont'
        'jwpei.com' = 'JW PEI'
    }
    foreach ($k in $map.Keys) { if ($h -eq $k -or $h.EndsWith(".$k")) { return $map[$k] } }
    $parts = $h.Split('.')
    $sld = if ($parts.Count -ge 2) { $parts[$parts.Count - 2] } else { $h }
    return (Get-Culture).TextInfo.ToTitleCase($sld)
}

# Guess a category from free text (URL path + page title)
function Guess-Category([string]$text) {
    $t = $text.ToLower()
    $rules = [ordered]@{
        'going out top' = 'Going-out tops'; 'going-out top' = 'Going-out tops'
        'blazer' = 'Blazers'; 'belt' = 'Belts'; 'cardigan' = 'Sweaters'; 'sweater' = 'Sweaters'
        'knit' = 'Sweaters'; 'jacket' = 'Jackets'; 'coat' = 'Jackets'; 'vest' = 'Vests'
        'skirt' = 'Skirts'; 'dress' = 'Dresses'; 'trouser' = 'Pants'; 'pant' = 'Pants'
        'jean' = 'Jeans'; 'denim' = 'Jeans'; 't-shirt' = 'T-shirts'; 'tee' = 'T-shirts'
        'sneaker' = 'Sneakers'; 'bootie' = 'Booties'; 'boot' = 'Booties'
        'tote' = 'Bags'; 'purse' = 'Bags'; 'handbag' = 'Bags'; 'bag' = 'Bags'
        'sandal' = 'Sandals'; 'loafer' = 'Shoes'; 'heel' = 'Shoes'; 'flat' = 'Shoes'; 'shoe' = 'Shoes'
        'blouse' = 'Shirts'; 'shirt' = 'Shirts'; 'top' = 'Shirts'
    }
    foreach ($k in $rules.Keys) { if ($t -like "*$k*") { return $rules[$k] } }
    return $null
}

# Pull a price + currency from the page HTML (JSON-LD, then meta tags, then Shopify cents)
function Get-PriceFromHtml([string]$Html) {
    foreach ($m in [regex]::Matches($Html, '(?s)<script[^>]*type="application/ld\+json"[^>]*>(.*?)</script>')) {
        try { $obj = $m.Groups[1].Value.Trim() | ConvertFrom-Json } catch { continue }
        $nodes = @()
        if ($obj -is [System.Array]) { $nodes += $obj } else { $nodes += $obj }
        if ($obj.PSObject.Properties.Name -contains '@graph') { $nodes += $obj.'@graph' }
        foreach ($node in $nodes) {
            if (-not $node -or -not $node.offers) { continue }
            $offerList = @()
            if ($node.offers -is [System.Array]) { $offerList += $node.offers } else { $offerList += $node.offers }
            foreach ($of in $offerList) {
                $p = $of.price
                if ($null -eq $p -and $of.priceSpecification) { $p = $of.priceSpecification.price }
                if ($null -ne $p -and "$p" -match '[0-9]') {
                    $cur = $of.priceCurrency
                    if (-not $cur -and $of.priceSpecification) { $cur = $of.priceSpecification.priceCurrency }
                    return [pscustomobject]@{ Price = "$p"; Currency = "$cur" }
                }
            }
        }
    }
    $og = [regex]::Match($Html, '(?i)(?:og:price:amount|product:price:amount)"\s+content="([0-9]+(?:\.[0-9]+)?)"')
    if ($og.Success) {
        $cur = [regex]::Match($Html, '(?i)(?:og:price:currency|product:price:currency)"\s+content="([A-Za-z]{3})"').Groups[1].Value
        return [pscustomobject]@{ Price = $og.Groups[1].Value; Currency = $cur.ToUpper() }
    }
    $sh = [regex]::Match($Html, '"price"\s*:\s*([0-9]{3,})\s*,\s*"price_min"')
    if ($sh.Success) {
        return [pscustomobject]@{ Price = ([double]$sh.Groups[1].Value / 100).ToString([System.Globalization.CultureInfo]::InvariantCulture); Currency = '' }
    }
    return $null
}
function Normalize-Price([string]$Raw) {
    $n = 0.0
    if (-not [double]::TryParse(($Raw -replace '[^0-9.]', ''), [ref]$n)) { return $null }
    if ($Raw -notmatch '\.' -and $n -ge 1000) { $n = $n / 100 }
    return [math]::Round($n, 2)
}

# Best product image URL (og:image both orders, JSON-LD image, twitter:image)
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
                $clean = [System.Net.WebUtility]::HtmlDecode($u).Trim() -replace '\\/', '/'
                if ($clean -match '^//') { $clean = 'https:' + $clean }
                if ($clean -match '^http://') { $clean = $clean -replace '^http://', 'https://' }
                if ($clean -match '^https?://') { return $clean }
            }
        }
    }
    return $null
}

# Detect image type from magic bytes; returns extension or $null
function Get-ImageExt([byte[]]$b) {
    if ($b.Length -lt 12) { return $null }
    if ($b[0] -eq 0xFF -and $b[1] -eq 0xD8 -and $b[2] -eq 0xFF) { return '.jpg' }
    if ($b[0] -eq 0x89 -and $b[1] -eq 0x50 -and $b[2] -eq 0x4E -and $b[3] -eq 0x47) { return '.png' }
    if ($b[0] -eq 0x47 -and $b[1] -eq 0x49 -and $b[2] -eq 0x46) { return '.gif' }
    if ($b[0] -eq 0x52 -and $b[1] -eq 0x49 -and $b[2] -eq 0x46 -and $b[8] -eq 0x57 -and $b[9] -eq 0x45) { return '.webp' }
    return $null
}

# --- fetch the page --------------------------------------------------------
$brand = Get-Brand $Url
Write-Host "Fetching $Url ..." -ForegroundColor Cyan
$blocked = $false
try {
    $resp = Invoke-WebRequest -Uri $Url -UseBasicParsing -TimeoutSec 25 -Headers $headers
    $html = $resp.Content
}
catch {
    Write-Warning "Could not fetch the page ($($_.Exception.Message)). I'll still add the item, but without an auto image/price."
    $html = ''
    $blocked = $true
}

# --- category --------------------------------------------------------------
if (-not $Category) {
    $title = ''
    if ($html) { $title = [regex]::Match($html, '(?i)<meta[^>]+property=["'']og:title["''][^>]+content=["'']([^"'']+)["'']').Groups[1].Value }
    $Category = Guess-Category ("$Url $title")
    if (-not $Category) { throw "Could not guess the category. Re-run with -Category, e.g. -Category Sweaters." }
    Write-Host "Guessed category: $Category  (override with -Category if wrong)" -ForegroundColor Yellow
}
$slot = $slotOfCategory[$Category]

# --- occasions default -----------------------------------------------------
if (-not $Occasions -or $Occasions.Count -eq 0) {
    $Occasions = @('Casual')
    Write-Host "No -Occasions given; defaulting to Casual. Adjust later with edit-item.ps1." -ForegroundColor Yellow
}

# --- price -----------------------------------------------------------------
$price = $null; $currency = ''
if ($html) {
    $found = Get-PriceFromHtml $html
    if ($found) { $price = Normalize-Price $found.Price; $currency = $found.Currency }
}

# --- next id ---------------------------------------------------------------
$catalog = Get-Content $catalogPath -Raw | ConvertFrom-Json
$nextId = (($catalog | ForEach-Object { [int]$_.id } | Measure-Object -Maximum).Maximum) + 1

# --- image -----------------------------------------------------------------
$imageField = ''
$imgUrl = if ($html) { Get-BestImageUrl $html } else { $null }
if (-not $NoImage -and $imgUrl) {
    if ($DryRun) {
        $imageField = "images/$nextId.<jpg|png|webp>"
        Write-Host "Would download image from: $imgUrl" -ForegroundColor DarkGray
    }
    else {
        try {
            $wc = New-Object System.Net.WebClient
            $wc.Headers['User-Agent'] = $headers['User-Agent']
            $bytes = $wc.DownloadData($imgUrl)
            $ext = Get-ImageExt $bytes
            if (-not $ext) {
                if ($imgUrl -match '(?i)\.(png|webp|jpe?g)(\?|$)') { $ext = '.' + ($Matches[1].ToLower() -replace 'jpeg', 'jpg') } else { $ext = '.jpg' }
            }
            $imagePath = Join-Path $imagesDir "$nextId$ext"
            [System.IO.File]::WriteAllBytes($imagePath, $bytes)
            $imageField = "images/$nextId$ext"
        }
        catch { Write-Warning "Image download failed: $($_.Exception.Message)" }
    }
}

# --- build the record ------------------------------------------------------
$record = [ordered]@{
    id          = $nextId
    brand       = $brand
    slot        = $slot
    category    = $Category
    weather     = $Weather
    occasions   = $Occasions
    image       = $imageField
    imageSource = 'web'
    url         = $Url
}
if ($null -ne $price) { $record.price = $price }
if ($currency) { $record.currency = $currency }

Write-Host ''
Write-Host 'New item:' -ForegroundColor Green
[pscustomobject]$record | Format-List
if (-not $imageField) { Write-Warning "No image was saved. Add one to data/images/$nextId.<jpg|png|webp> so it shows in the app." }

if ($DryRun) { Write-Host 'DryRun: nothing was written.' -ForegroundColor Yellow; return }

# --- append + save ---------------------------------------------------------
$catalog += [pscustomobject]$record
$catalog | ConvertTo-Json -Depth 12 | Set-Content -Path $catalogPath -Encoding UTF8

Write-Host ''
Write-Host "Added #$nextId ($brand $Category) to the catalog." -ForegroundColor Cyan
Write-Host "Review/adjust tags with:  .\tools\edit-item.ps1 -Id $nextId -Show"
