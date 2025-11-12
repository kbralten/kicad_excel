<#
Query KiCad Excel Bridge API endpoints and print results.

Usage:
  .\query_kicad_api.ps1 [command] [id]

Commands:
  info               - GET /v1/ (API validation)
  categories         - GET /v1/categories.json
  parts <partId>     - GET /v1/parts/{partId}.json
  parts-for <catId>  - GET /v1/parts/category/{catId}.json
  all                - Call `info` and `categories` and show first category parts

Examples:
  .\query_kicad_api.ps1 info
  .\query_kicad_api.ps1 categories
  .\query_kicad_api.ps1 parts P12345
  .\query_kicad_api.ps1 parts-for resistors
  .\query_kicad_api.ps1 all

Note: Server default base url: http://localhost:8088/kicad-api/v1/
#>

param(
    [string]$Command = "info",
    [string]$Id
)

$base = 'http://localhost:8088/kicad-api/v1'

function Get-Json {
    param([string]$url)
    try {
        $resp = Invoke-RestMethod -Uri $url -Method Get -UseBasicParsing -ErrorAction Stop
        $json = $resp | ConvertTo-Json -Depth 10
        Write-Host "GET $url" -ForegroundColor Cyan
        Write-Host $json
    }
    catch {
        Write-Host "Error calling ${url}:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
}

switch ($Command.ToLower()) {
    'info' {
        Get-Json "$base/"
        break
    }
    'categories' {
        Get-Json "$base/categories.json"
        break
    }
    'parts' {
        if (-not $Id) { Write-Host "parts requires a <partId>" -ForegroundColor Yellow; break }
        Get-Json "$base/parts/$Id.json"
        break
    }
    'parts-for' {
        if (-not $Id) { Write-Host "parts-for requires a <categoryId>" -ForegroundColor Yellow; break }
        Get-Json "$base/parts/category/$Id.json"
        break
    }
    'all' {
        Get-Json "$base/"
        Start-Sleep -Milliseconds 200
        Get-Json "$base/categories.json"
        # attempt to query first category's parts (best-effort)
        try {
            $cats = Invoke-RestMethod -Uri "$base/categories.json" -Method Get -UseBasicParsing -ErrorAction Stop
            if ($cats -and $cats[0]) {
                $firstId = $cats[0].Id
                if ($firstId) { Get-Json "$base/parts/category/$firstId.json" }
            }
        } catch { }
        break
    }
    Default {
        Write-Host "Unknown command: $Command" -ForegroundColor Yellow
        break
    }
}
