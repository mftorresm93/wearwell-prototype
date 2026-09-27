# Test robust product-image extraction across brands (og:image, JSON-LD, twitter:image)
$ErrorActionPreference = 'Continue'

$urls = @(
    'https://www.aritzia.com/us/en/product/bare-merino-wool-aspect-polo-sweater/127835.html?color=37623',
    'https://www.quince.com/women/cashmere/cashmere-crewneck-sweater?color=heather-grey&gender=women',
    'https://www.everlane.com/products/womens-organic-cotton-box-cut-tee-white',
    'https://www.madewell.com/p/womens/clothing/tees/the-perfect-crewneck-tee-in-allday-cotton/NY429/?ccode=KI1567',
    'https://www.zara.com/us/en/leather-slingback-shoes-p11524810.html?v1=545422395'
)

$headers = @{
    'User-Agent'      = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36'
    'Accept'          = 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8'
    'Accept-Language' = 'en-US,en;q=0.9'
}

function Extract-Images([string]$html) {
    $found = [ordered]@{}
    # og:image (both attribute orders)
    foreach ($m in [regex]::Matches($html, '<meta[^>]+property=["'']og:image(?::secure_url)?["''][^>]+content=["'']([^"'']+)["'']')) {
        $found["og:$($m.Groups[1].Value)"] = $m.Groups[1].Value
    }
    foreach ($m in [regex]::Matches($html, '<meta[^>]+content=["'']([^"'']+)["''][^>]+property=["'']og:image["'']')) {
        $found["og2:$($m.Groups[1].Value)"] = $m.Groups[1].Value
    }
    # twitter:image
    foreach ($m in [regex]::Matches($html, '<meta[^>]+name=["'']twitter:image["''][^>]+content=["'']([^"'']+)["'']')) {
        $found["tw:$($m.Groups[1].Value)"] = $m.Groups[1].Value
    }
    # JSON-LD Product.image
    foreach ($m in [regex]::Matches($html, '(?s)<script[^>]+type=["'']application/ld\+json["''][^>]*>(.*?)</script>')) {
        $imgs = [regex]::Matches($m.Groups[1].Value, '"image"\s*:\s*("(?<s>[^"]+)"|\[(?<a>[^\]]+)\])')
        foreach ($im in $imgs) {
            if ($im.Groups['s'].Success) { $found["ld:$($im.Groups['s'].Value)"] = $im.Groups['s'].Value }
            elseif ($im.Groups['a'].Success) {
                foreach ($u in [regex]::Matches($im.Groups['a'].Value, '"([^"]+)"')) { $found["lda:$($u.Groups[1].Value)"] = $u.Groups[1].Value }
            }
        }
    }
    return $found.Values | Select-Object -Unique
}

foreach ($u in $urls) {
    Write-Host "`n===== $u"
    try {
        $r = Invoke-WebRequest -Uri $u -Headers $headers -TimeoutSec 30 -MaximumRedirection 5 -UseBasicParsing
        $imgs = @(Extract-Images $r.Content)
        Write-Host "  status=$($r.StatusCode)  found $($imgs.Count) image url(s):"
        $imgs | Select-Object -First 6 | ForEach-Object { Write-Host "    $_" }
    }
    catch {
        Write-Host "  ERROR: $($_.Exception.Message)"
    }
}
