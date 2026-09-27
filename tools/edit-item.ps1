<#
  edit-item.ps1
  Review or fix the tags on a single catalog item by ID.

  EXAMPLES:
    .\tools\edit-item.ps1 -Id 74 -Show                         # print the item
    .\tools\edit-item.ps1 -Id 74 -Occasions Work,Casual        # set occasions
    .\tools\edit-item.ps1 -Id 74 -Category Pants               # fix category (slot auto-updates)
    .\tools\edit-item.ps1 -Id 74 -Weather Cold -Price 88       # set weather + price
    .\tools\edit-item.ps1 -Id 74 -Delete                       # remove the item (and its image)
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][int]$Id,
    [string]$Brand,
    [ValidateSet('Shirts', 'T-shirts', 'Going-out tops', 'Sweaters', 'Blazers', 'Jackets', 'Vests',
        'Jeans', 'Pants', 'Skirts', 'Dresses', 'Shoes', 'Sneakers', 'Booties', 'Sandals', 'Bags', 'Belts')]
    [string]$Category,
    [string[]]$Occasions,
    [ValidateSet('All year', 'Cold', 'Hot')]
    [string]$Weather,
    [double]$Price,
    [string]$Url,
    [switch]$Delete,
    [switch]$Show
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$dataDir = Join-Path $root 'data'
$catalogPath = Join-Path $dataDir 'catalog.json'
$imagesDir = Join-Path $dataDir 'images'

$slotOfCategory = @{
    'Shirts' = 'Top'; 'T-shirts' = 'Top'; 'Going-out tops' = 'Top'; 'Sweaters' = 'Top'
    'Blazers' = 'Layer'; 'Jackets' = 'Layer'; 'Vests' = 'Top'
    'Jeans' = 'Bottom'; 'Pants' = 'Bottom'; 'Skirts' = 'Bottom'
    'Dresses' = 'Dress'
    'Shoes' = 'Shoes'; 'Sneakers' = 'Shoes'; 'Booties' = 'Shoes'; 'Sandals' = 'Shoes'
    'Bags' = 'Bag'
    'Belts' = 'Accessory'
}

$catalog = Get-Content $catalogPath -Raw | ConvertFrom-Json
$item = $catalog | Where-Object { [int]$_.id -eq $Id } | Select-Object -First 1
if (-not $item) { throw "No item with id #$Id." }

if ($Show) { $item | Format-List; return }

if ($Delete) {
    $catalog = @($catalog | Where-Object { [int]$_.id -ne $Id })
    $catalog | ConvertTo-Json -Depth 12 | Set-Content -Path $catalogPath -Encoding UTF8
    Get-ChildItem (Join-Path $imagesDir "$Id.*") -ErrorAction SilentlyContinue | Remove-Item -Force
    Write-Host "Deleted #$Id and its image." -ForegroundColor Cyan
    return
}

$changed = @()
if ($PSBoundParameters.ContainsKey('Brand')) { $item.brand = $Brand; $changed += 'brand' }
if ($PSBoundParameters.ContainsKey('Category')) {
    $item.category = $Category
    $item.slot = $slotOfCategory[$Category]
    $changed += 'category', 'slot'
}
if ($PSBoundParameters.ContainsKey('Occasions')) { $item.occasions = $Occasions; $changed += 'occasions' }
if ($PSBoundParameters.ContainsKey('Weather')) { $item.weather = $Weather; $changed += 'weather' }
if ($PSBoundParameters.ContainsKey('Url')) { $item.url = $Url; $changed += 'url' }
if ($PSBoundParameters.ContainsKey('Price')) {
    if ($item.PSObject.Properties.Name -contains 'price') { $item.price = $Price }
    else { $item | Add-Member -NotePropertyName price -NotePropertyValue $Price -Force }
    $changed += 'price'
}

if ($changed.Count -eq 0) { Write-Host "Nothing to change. Use -Show to view, or pass a field to set." -ForegroundColor Yellow; return }

$catalog | ConvertTo-Json -Depth 12 | Set-Content -Path $catalogPath -Encoding UTF8
Write-Host ("Updated #{0}: {1}" -f $Id, ($changed -join ', ')) -ForegroundColor Cyan
$item | Format-List
