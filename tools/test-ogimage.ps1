# Test: can we read og:image (official product photo) from each product URL?
$ErrorActionPreference = 'Continue'
$urls = @(
    'https://www.aritzia.com/us/en/product/bare-merino-wool-aspect-polo-sweater/127835.html?color=37623',
    'https://www.quince.com/women/cashmere/cashmere-crewneck-sweater?color=heather-grey&gender=women',
    'https://www.everlane.com/products/womens-organic-cotton-box-cut-tee-white',
    'https://www.madewell.com/p/womens/clothing/tees/the-perfect-crewneck-tee-in-allday-cotton/NY429/?ccode=KI1567'
)
$ua = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36'
foreach ($u in $urls) {
    Write-Host "`n--- $u"
    try {
        $r = Invoke-WebRequest -Uri $u -UserAgent $ua -TimeoutSec 25 -MaximumRedirection 5 -UseBasicParsing
        $html = $r.Content
        $og = [regex]::Match($html, '<meta[^>]+property=["'']og:image["''][^>]+content=["'']([^"'']+)["'']')
        if (-not $og.Success) {
            $og = [regex]::Match($html, '<meta[^>]+content=["'']([^"'']+)["''][^>]+property=["'']og:image["'']')
        }
        if ($og.Success) { Write-Host "  status=$($r.StatusCode)  og:image = $($og.Groups[1].Value)" }
        else { Write-Host "  status=$($r.StatusCode)  og:image NOT FOUND" }
    }
    catch {
        Write-Host "  ERROR: $($_.Exception.Message)"
    }
}
