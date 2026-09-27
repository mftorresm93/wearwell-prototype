$c = Get-Content (Join-Path $PSScriptRoot '..\data\catalog.json') -Raw | ConvertFrom-Json
Write-Host 'imageSource counts:'
$c | Group-Object imageSource | ForEach-Object { '  {0,-12} {1}' -f $_.Name, $_.Count }
Write-Host "`nSample Madewell items:"
$c | Where-Object { $_.brand -eq 'Madewell' } | Select-Object -First 3 id, imageSource, image | Format-Table -AutoSize | Out-String | Write-Host
$img = Join-Path $PSScriptRoot '..\data\images'
Write-Host 'Extension breakdown in images/:'
Get-ChildItem (Join-Path $img '*') | Group-Object Extension | ForEach-Object { '  {0} {1}' -f $_.Name, $_.Count }
Write-Host "`nCheck a few Madewell files exist:"
$c | Where-Object { $_.brand -eq 'Madewell' } | Select-Object -First 3 | ForEach-Object {
    $p = Join-Path $img (Split-Path $_.image -Leaf)
    '  #{0} {1} exists={2}' -f $_.id, $_.image, (Test-Path $p)
}
