# Diagnostic: compare picture count vs URL count per content slide
$ErrorActionPreference = 'Stop'
$tmp = Join-Path $env:TEMP 'pptx_extract_parse'
$slides = Get-ChildItem "$tmp\ppt\slides\*.xml" | Sort-Object { [int]($_.BaseName -replace '\D', '') }
$mismatch = 0; $contentSlides = 0
foreach ($f in $slides) {
    $xml = [xml](Get-Content $f.FullName)
    $runs = @($xml.SelectNodes("//*[local-name()='t']") | ForEach-Object { $_.'#text'.Trim() } | Where-Object { $_ })
    $urls = @($runs | Where-Object { $_ -match '^https?://' })
    if ($urls.Count -eq 0) { continue }
    $contentSlides++
    $pics = @($xml.SelectNodes("//*[local-name()='pic']")).Count
    if ($pics -ne $urls.Count) {
        $mismatch++
        "{0,-12} urls={1} pics={2}" -f $f.BaseName, $urls.Count, $pics
    }
}
Write-Host "`nContent slides: $contentSlides  |  slides where pics != urls: $mismatch"
