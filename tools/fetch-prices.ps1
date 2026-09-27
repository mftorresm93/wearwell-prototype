<#
  fetch-prices.ps1
  Reads data/catalog.json, fetches each product URL, extracts the price + currency
  from the page's JSON-LD structured data, normalizes cents, and writes them back.
  Anything it cannot fetch (blocked/missing) is listed for manual entry.

  Usage:
    .\tools\fetch-prices.ps1            # fill in only items missing a price
    .\tools\fetch-prices.ps1 -Force     # re-fetch every item
    .\tools\fetch-prices.ps1 -DelayMs 800
#>
[CmdletBinding()]
param(
  [switch]$Force,
  [int]$DelayMs = 600
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$dataDir = Join-Path $root 'data'
$catalogPath = Join-Path $dataDir 'catalog.json'

$catalog = Get-Content -Raw $catalogPath | ConvertFrom-Json

$headers = @{
  'User-Agent'      = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0 Safari/537.36'
  'Accept'          = 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
  'Accept-Language' = 'en-US,en;q=0.9'
}

# Pulls a numeric price + currency out of any JSON-LD Product/offers blocks in the HTML.
function Get-PriceFromHtml {
  param([string]$Html)

  $matches = [regex]::Matches($Html, '(?s)<script[^>]*type="application/ld\+json"[^>]*>(.*?)</script>')
  foreach ($m in $matches) {
    $json = $m.Groups[1].Value.Trim()
    try { $obj = $json | ConvertFrom-Json } catch { continue }

    # JSON-LD may be a single object, an array, or wrapped in @graph.
    $nodes = @()
    if ($obj -is [System.Array]) { $nodes += $obj } else { $nodes += $obj }
    if ($obj.PSObject.Properties.Name -contains '@graph') { $nodes += $obj.'@graph' }

    foreach ($node in $nodes) {
      if (-not $node) { continue }
      $offers = $node.offers
      if (-not $offers) { continue }
      $offerList = @()
      if ($offers -is [System.Array]) { $offerList += $offers } else { $offerList += $offers }
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

  # Fallback 1: Open Graph / product meta tags (Shopify + many others expose a clean amount).
  $og = [regex]::Match($Html, '(?i)(?:og:price:amount|product:price:amount)"\s+content="([0-9]+(?:\.[0-9]+)?)"')
  if ($og.Success) {
    $cur = [regex]::Match($Html, '(?i)(?:og:price:currency|product:price:currency)"\s+content="([A-Za-z]{3})"').Groups[1].Value
    return [pscustomobject]@{ Price = $og.Groups[1].Value; Currency = $cur.ToUpper() }
  }

  # Fallback 2: Shopify product JSON stores the price in cents, e.g. "price":8800,"price_min":8800.
  $sh = [regex]::Match($Html, '"price"\s*:\s*([0-9]{3,})\s*,\s*"price_min"')
  if ($sh.Success) {
    return [pscustomobject]@{ Price = ([double]$sh.Groups[1].Value / 100).ToString([System.Globalization.CultureInfo]::InvariantCulture); Currency = '' }
  }

  return $null
}

# Turns a raw JSON-LD price string into a clean number, fixing cents-encoded integers.
function Normalize-Price {
  param([string]$Raw)
  $n = 0.0
  if (-not [double]::TryParse(($Raw -replace '[^0-9.]', ''), [ref]$n)) { return $null }
  # Integers with no decimal that are suspiciously large are usually minor units (cents).
  if ($Raw -notmatch '\.' -and $n -ge 1000) { $n = $n / 100 }
  return [math]::Round($n, 2)
}

$updated = 0
$review = @()
$total = ($catalog | Where-Object { $_.url }).Count
$i = 0

foreach ($item in $catalog) {
  if (-not $item.url) { continue }
  $i++

  $hasPrice = ($item.PSObject.Properties.Name -contains 'price') -and ($null -ne $item.price) -and ("$($item.price)" -ne '')
  if ($hasPrice -and -not $Force) { continue }

  Write-Host ("[{0}/{1}] #{2} {3} ..." -f $i, $total, $item.id, $item.brand) -NoNewline
  try {
    $resp = Invoke-WebRequest -Uri $item.url -UseBasicParsing -TimeoutSec 25 -Headers $headers
    $found = Get-PriceFromHtml -Html $resp.Content
    if ($found) {
      $price = Normalize-Price -Raw $found.Price
      if ($null -ne $price) {
        $item | Add-Member -NotePropertyName price -NotePropertyValue $price -Force
        if ($found.Currency) { $item | Add-Member -NotePropertyName currency -NotePropertyValue $found.Currency -Force }
        $updated++
        Write-Host (" {0} {1}" -f $found.Currency, $price) -ForegroundColor Green
      } else {
        Write-Host " no usable number" -ForegroundColor Yellow
        $review += [pscustomobject]@{ id = $item.id; brand = $item.brand; reason = "unparsable: $($found.Price)"; url = $item.url }
      }
    } else {
      Write-Host " no price in page" -ForegroundColor Yellow
      $review += [pscustomobject]@{ id = $item.id; brand = $item.brand; reason = 'no JSON-LD price'; url = $item.url }
    }
  } catch {
    $msg = $_.Exception.Message
    Write-Host (" ERROR: {0}" -f $msg) -ForegroundColor Red
    $review += [pscustomobject]@{ id = $item.id; brand = $item.brand; reason = $msg; url = $item.url }
  }

  Start-Sleep -Milliseconds $DelayMs
}

$catalog | ConvertTo-Json -Depth 12 | Set-Content -Path $catalogPath -Encoding UTF8

Write-Host ""
Write-Host ("Done. Prices set/updated: {0}. Needs manual review: {1}." -f $updated, $review.Count) -ForegroundColor Cyan
if ($review.Count) {
  $reviewPath = Join-Path $dataDir 'price-review.csv'
  $review | Export-Csv -Path $reviewPath -NoTypeInformation -Encoding UTF8
  Write-Host ("Review list written to: {0}" -f $reviewPath)
  $review | Format-Table -AutoSize
}
