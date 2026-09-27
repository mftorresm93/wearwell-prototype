# Generates a single "styled model" photo of a full outfit using the OpenAI image API.
#
# It takes a set of catalog item IDs, builds a descriptive prompt from their
# brand/slot/category metadata, and (by default) sends the items' actual product
# images as visual references to gpt-image-1 so the result resembles YOUR pieces.
# The generated PNG is saved to data/generated/.
#
# Pipeline fit:  the outfit rules engine / board picks item IDs  ->  this script
# turns that outfit into a model photo.
#
# COST: this is the only script that spends money. Each image costs a few cents,
# billed to the OpenAI API balance tied to your key. Use -DryRun to preview the
# prompt and reference list for free (no API call, no charge).
#
# SETUP (one time):
#   1) Put money on your API balance at platform.openai.com  (separate from ChatGPT Plus)
#   2) In this terminal set your key (paste it here, NOT into chat):
#        $env:OPENAI_API_KEY = "sk-..."
#
# EXAMPLES:
#   # Free preview of the prompt for the demo outfit:
#   tools\generate-outfit-image.ps1 -DryRun
#
#   # Generate the demo Work outfit (top 2, jeans 58, loafers 88, bag 98):
#   tools\generate-outfit-image.ps1 -Items 2,58,88,98 -Occasion Work
#
#   # Prompt-only (ignore product photos, cheaper/looser):
#   tools\generate-outfit-image.ps1 -Items 2,58,88,98 -NoReference

param(
    [int[]]$Items = @(2, 58, 88, 98),                       # catalog IDs that make up the outfit
    [string]$Occasion = 'Work',                             # styling context for the prompt
    [string]$Weather = 'All year',                          # weather context for the prompt
    [string]$Background = 'a bright, minimal home interior with natural light',
    [string]$Notes = '',                                    # free-text extra styling details for the prompt
    [ValidateSet('1024x1024', '1024x1536', '1536x1024', 'auto')]
    [string]$Size = '1024x1536',                            # portrait suits a full-body model
    [ValidateSet('low', 'medium', 'high', 'auto')]
    [string]$Quality = 'high',
    [ValidateSet('low', 'high')]
    [string]$InputFidelity = 'high',                        # 'high' makes gpt-image-1 preserve reference details (bags, prints, faces)
    [string]$Model = 'gpt-image-1',
    [string]$OutFile,                                       # optional explicit output path
    [switch]$NoReference,                                   # generate from text prompt only
    [switch]$DryRun                                         # build + show the prompt, make NO API call
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$dataDir = Join-Path $root 'data'
$catalogPath = Join-Path $dataDir 'catalog.json'
if (-not (Test-Path $catalogPath)) { throw "Missing catalog.json. Run tools\parse-deck.ps1 first." }

$catalog = Get-Content $catalogPath -Raw | ConvertFrom-Json

# --- Resolve the requested items (preserve the order they were passed in) ---
$byId = @{}
foreach ($it in $catalog) { $byId[[int]$it.id] = $it }

$chosen = @()
foreach ($id in $Items) {
    if ($byId.ContainsKey([int]$id)) { $chosen += $byId[[int]$id] }
    else { Write-Warning "No catalog item with id #$id - skipping." }
}
if ($chosen.Count -eq 0) { throw "None of the requested item IDs were found in the catalog." }

# --- Turn a catalog category into a natural noun for the prompt ---
function Get-ItemNoun($item) {
    switch ($item.category) {
        'Shirts' { 'top' }
        'T-shirts' { 't-shirt' }
        'Going-out tops' { 'going-out top' }
        'Sweaters' { 'sweater' }
        'Blazers' { 'blazer' }
        'Jackets' { 'jacket' }
        'Vests' { 'vest' }
        'Jeans' { 'jeans' }
        'Pants' { 'trousers' }
        'Skirts' { 'skirt' }
        'Dresses' { 'dress' }
        'Shoes' { 'shoes' }
        'Sneakers' { 'sneakers' }
        'Booties' { 'ankle boots' }
        'Bags' { 'handbag' }
        'Belts' { 'belt' }
        default { [string]$item.category }
    }
}
$pluralNouns = @('jeans', 'trousers', 'shoes', 'sneakers', 'ankle boots')

# --- Collect reference images (in outfit order) for the chosen items ---
$imagesDir = Join-Path $dataDir 'images'
$refs = @()
if (-not $NoReference) {
    foreach ($it in $chosen) {
        $f = Get-ChildItem (Join-Path $imagesDir "$($it.id).*") -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($f) {
            $mime = switch ($f.Extension.ToLower()) {
                '.jpg' { 'image/jpeg' }
                '.jpeg' { 'image/jpeg' }
                '.png' { 'image/png' }
                '.webp' { 'image/webp' }
                default { $null }
            }
            if ($mime) {
                $refs += [pscustomobject]@{
                    Path = $f.FullName; Name = $f.Name; Mime = $mime
                    Slot = [string]$it.slot; Noun = (Get-ItemNoun $it); Brand = [string]$it.brand
                }
            }
            else { Write-Warning "Skipping unsupported reference image type for #$($it.id): $($f.Name)" }
        }
        else { Write-Warning "No product image found for #$($it.id) - it won't be used as a reference." }
    }
}

# --- Describe each piece for the prompt (correct article + slot role) ---
$descParts = @()
foreach ($it in $chosen) {
    $noun = Get-ItemNoun $it
    $article = if ($pluralNouns -contains $noun) { '' } else { 'a ' }
    $descParts += "$($it.slot.ToLower()): $article$($it.brand) $noun"
}
$pieceList = ($descParts -join '; ')

# --- Build the prompt ---
$accClause = ''
if (-not $NoReference -and $refs.Count -gt 0) {
    $refList = @()
    for ($i = 0; $i -lt $refs.Count; $i++) {
        $r = $refs[$i]
        $refList += "image $($i + 1) is the $($r.Brand) $($r.Noun)"
    }
    $refMap = $refList -join '; '
    $fidelityClause = "`nYou are given $($refs.Count) reference product images. In order, $refMap. Treat every reference image as the absolute ground truth for that item's color, shade, pattern, fabric, texture, hardware, cut, length, and proportions. Reproduce each item exactly as shown. Do NOT recolor, restyle, embellish, simplify, upgrade, shorten, lengthen, resize, or swap any item, and do NOT invent or add anything that is not shown in a reference image."

    $accNotes = @()
    if ($chosen | Where-Object { $_.category -eq 'Jeans' }) { $accNotes += "The jeans must match their reference image exactly: the same denim wash and color depth, the same fading and whiskering pattern, the same rise, and the same leg cut and full length - do NOT make them skinnier, wider, cropped, distressed, or a lighter or darker wash than shown." }
    if ($chosen | Where-Object { $_.category -eq 'Bags' }) { $accNotes += "The handbag must be clearly visible (held in the hand or worn on the shoulder) and must match its reference image exactly in color, shape, size, silhouette, and hardware; keep it in sharp focus and do not shrink, restyle, or simplify it." }
    if ($chosen | Where-Object { $_.category -eq 'Belts' }) { $accNotes += "The belt must be worn at the waist and match its reference image exactly in color, width, and buckle style." }
    if ($accNotes.Count) { $accClause = "`n" + ($accNotes -join ' ') }
}
else { $fidelityClause = '' }
$notesClause = if ([string]::IsNullOrWhiteSpace($Notes)) { '' } else { " Additional details (only to clarify, never to override the reference images): $Notes." }

$prompt = @"
Full-body editorial fashion photograph of a single female model wearing exactly one complete outfit for a $Occasion setting in $Weather weather.$fidelityClause
The outfit is made up of EXACTLY these items and nothing else - $pieceList.
Do not add, invent, or substitute any clothing that is not in that list: no extra jacket, coat, blazer, cardigan, scarf, hat, belt, sunglasses, or jewelry unless it is listed above.
Keep every garment's exact color, shade, and pattern as shown in its reference image; do not recolor anything and do not change the shoe style or shoe color.$accClause
The model has a clear, natural face with no sunglasses and no bangs or fringe; hair is kept away from the forehead.
Show the full body from head to toe, including the shoes, in a natural relaxed standing pose.
Setting: $Background. Soft natural lighting, realistic proportions, photorealistic, high resolution, clean and cohesive styling.$notesClause
"@.Trim()

# --- Decide output path ---
$genDir = Join-Path $dataDir 'generated'
if (-not (Test-Path $genDir)) { New-Item -ItemType Directory -Path $genDir | Out-Null }
if (-not $OutFile) {
    $stamp = ($Items -join '-')
    $OutFile = Join-Path $genDir "outfit-$stamp.png"
}

# --- Show what we're about to do ---
Write-Host ''
Write-Host 'Outfit:' -ForegroundColor Cyan
$chosen | ForEach-Object { Write-Host ("  #{0,-3} {1,-8} {2,-14} {3}" -f $_.id, $_.slot, $_.category, $_.brand) }
Write-Host ''
Write-Host 'Prompt:' -ForegroundColor Cyan
Write-Host $prompt
Write-Host ''
if ($NoReference) { Write-Host 'Reference images: (none - prompt only)' -ForegroundColor Cyan }
else { Write-Host "Reference images: $($refs.Count)" -ForegroundColor Cyan; $refs | ForEach-Object { Write-Host "  $($_.Name)" } }
Write-Host ''

if ($DryRun) {
    Write-Host 'DryRun: no API call made, nothing was charged.' -ForegroundColor Yellow
    Write-Host "Would save to: $OutFile"
    return
}

# --- Require the API key only when actually calling out ---
$apiKey = $env:OPENAI_API_KEY
if ([string]::IsNullOrWhiteSpace($apiKey)) {
    throw "OPENAI_API_KEY is not set. In this terminal run:  `$env:OPENAI_API_KEY = `"sk-...`"  (paste your key here, not into chat), then re-run."
}

Write-Host 'Calling OpenAI image API...' -ForegroundColor Green

$b64 = $null
if ($refs.Count -gt 0) {
    # Reference-guided: multipart POST to /v1/images/edits with image[] parts.
    Add-Type -AssemblyName System.Net.Http
    $client = [System.Net.Http.HttpClient]::new()
    $client.Timeout = [TimeSpan]::FromMinutes(5)
    $client.DefaultRequestHeaders.Authorization = [System.Net.Http.Headers.AuthenticationHeaderValue]::new('Bearer', $apiKey)
    try {
        $form = [System.Net.Http.MultipartFormDataContent]::new()
        $form.Add([System.Net.Http.StringContent]::new($Model), 'model')
        $form.Add([System.Net.Http.StringContent]::new($prompt), 'prompt')
        $form.Add([System.Net.Http.StringContent]::new($Size), 'size')
        $form.Add([System.Net.Http.StringContent]::new($Quality), 'quality')
        $form.Add([System.Net.Http.StringContent]::new($InputFidelity), 'input_fidelity')
        $form.Add([System.Net.Http.StringContent]::new('1'), 'n')
        foreach ($r in $refs) {
            $bytes = [System.IO.File]::ReadAllBytes($r.Path)
            $bc = [System.Net.Http.ByteArrayContent]::new($bytes)
            $bc.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::new($r.Mime)
            $form.Add($bc, 'image[]', $r.Name)
        }
        $resp = $client.PostAsync('https://api.openai.com/v1/images/edits', $form).GetAwaiter().GetResult()
        $respBody = $resp.Content.ReadAsStringAsync().GetAwaiter().GetResult()
        if (-not $resp.IsSuccessStatusCode) { throw "OpenAI API error ($([int]$resp.StatusCode)): $respBody" }
        $b64 = ($respBody | ConvertFrom-Json).data[0].b64_json
    }
    finally { $client.Dispose() }
}
else {
    # Prompt-only: JSON POST to /v1/images/generations.
    $body = @{ model = $Model; prompt = $prompt; size = $Size; quality = $Quality; n = 1 } | ConvertTo-Json
    $resp = Invoke-RestMethod -Uri 'https://api.openai.com/v1/images/generations' -Method Post `
        -Headers @{ Authorization = "Bearer $apiKey" } -ContentType 'application/json' -Body $body
    $b64 = $resp.data[0].b64_json
}

if ([string]::IsNullOrWhiteSpace($b64)) { throw "No image returned by the API." }

[System.IO.File]::WriteAllBytes($OutFile, [Convert]::FromBase64String($b64))
Write-Host ''
Write-Host "Saved: $OutFile" -ForegroundColor Green
Write-Host "Open it in VS Code, or add it to a board with:  <img src='generated/$(Split-Path $OutFile -Leaf)'>"
