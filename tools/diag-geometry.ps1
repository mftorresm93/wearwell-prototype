# Diagnostic: dump geometry of pics and text boxes for a given slide
param([int]$Slide = 18)
$ErrorActionPreference = 'Stop'
$tmp = Join-Path $env:TEMP 'pptx_extract_parse'
$path = Join-Path $tmp "ppt\slides\slide$Slide.xml"
$xml = [xml](Get-Content $path)

function OffExt($node) {
    $xfrm = $node.SelectSingleNode(".//*[local-name()='xfrm']")
    if (-not $xfrm) { return $null }
    $off = $xfrm.SelectSingleNode("*[local-name()='off']")
    $ext = $xfrm.SelectSingleNode("*[local-name()='ext']")
    [pscustomobject]@{
        x  = [int64]$off.x; y = [int64]$off.y
        cx = [int64]$ext.cx; cy = [int64]$ext.cy
    }
}

Write-Host "=== PICTURES (slide$Slide) ==="
$i = 0
foreach ($pic in $xml.SelectNodes("//*[local-name()='pic']")) {
    $i++
    $g = OffExt $pic
    $blip = $pic.SelectSingleNode(".//*[local-name()='blip']")
    $embed = $blip.Attributes['r:embed'].Value
    if ($g) {
        "  pic{0}  x={1,-9} y={2,-9} cx={3,-9} cy={4,-9} area={5,-12} rId={6}" -f `
            $i, $g.x, $g.y, $g.cx, $g.cy, ($g.cx * $g.cy), $embed
    }
    else { "  pic{0}  (no xfrm) rId={1}" -f $i, $embed }
}

Write-Host "`n=== TEXT BOXES (sp with URLs) ==="
foreach ($sp in $xml.SelectNodes("//*[local-name()='sp']")) {
    $txt = ($sp.SelectNodes(".//*[local-name()='t']") | ForEach-Object { $_.'#text' }) -join ' | '
    if ($txt -match 'http') {
        $g = OffExt $sp
        if ($g) { "  box x={0,-9} y={1,-9} cx={2,-9} cy={3}" -f $g.x, $g.y, $g.cx, $g.cy }
        else { "  box (no xfrm)" }
        "    text: $txt"
    }
}
