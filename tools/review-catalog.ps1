<#
  review-catalog.ps1
  Review the catalog as a table so you can spot items that need better tags.

  EXAMPLES:
    .\tools\review-catalog.ps1                       # everything, grouped by slot
    .\tools\review-catalog.ps1 -Slot Top             # only tops
    .\tools\review-catalog.ps1 -Brand Everlane       # only one brand
    .\tools\review-catalog.ps1 -NeedsAttention       # only items missing price/image/occasions
#>
[CmdletBinding()]
param(
    [string]$Slot,
    [string]$Brand,
    [string]$Category,
    [switch]$NeedsAttention
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$dataDir = Join-Path $root 'data'
$catalog = Get-Content (Join-Path $dataDir 'catalog.json') -Raw | ConvertFrom-Json

$rows = $catalog | ForEach-Object {
    $hasPrice = ($_.PSObject.Properties.Name -contains 'price') -and ($null -ne $_.price) -and ("$($_.price)" -ne '')
    $hasImage = $_.image -and $_.image.Trim()
    $hasOcc = $_.occasions -and @($_.occasions).Count -gt 0
    $flags = @()
    if (-not $hasPrice) { $flags += 'no-price' }
    if (-not $hasImage) { $flags += 'no-image' }
    if (-not $hasOcc) { $flags += 'no-occasions' }
    [pscustomobject]@{
        id        = [int]$_.id
        slot      = $_.slot
        category  = $_.category
        brand     = $_.brand
        weather   = $_.weather
        occasions = (@($_.occasions) -join ',')
        price     = if ($hasPrice) { '$' + $_.price } else { '' }
        needs     = ($flags -join ' ')
    }
}

if ($Slot) { $rows = $rows | Where-Object { $_.slot -eq $Slot } }
if ($Brand) { $rows = $rows | Where-Object { $_.brand -eq $Brand } }
if ($Category) { $rows = $rows | Where-Object { $_.category -eq $Category } }
if ($NeedsAttention) { $rows = $rows | Where-Object { $_.needs } }

$rows = $rows | Sort-Object slot, category, brand, id
$rows | Format-Table -AutoSize

Write-Host ''
Write-Host ("Showing {0} item(s). Needing attention: {1}." -f `
        @($rows).Count, @($rows | Where-Object { $_.needs }).Count) -ForegroundColor Cyan
Write-Host "Fix tags with:  .\tools\edit-item.ps1 -Id <id> -Occasions Work,Casual -Weather 'All year'"
