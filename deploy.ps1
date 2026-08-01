#Requires -Version 5.1
<#
.SYNOPSIS
  Build, version, and ship the Finanzas Familia frontend.

.DESCRIPTION
  Two modes:
    - Default:  prompts for a version tag, runs `clasp deploy -d <tag>`,
                resolves the new deployment id via `clasp list-deployments`,
                rewrites both GAS URLs in netlify/index.html, ships to Netlify prod.
    - Redeploy: skips clasp + regex; just re-pushes netlify/index.html as-is.

  Run from the repo root in pwsh.

.PARAMETER Redeploy
  Skip the GAS deploy + URL rewrite. Just runs `netlify deploy --prod`.

.PARAMETER Help
  Print this help and exit. Equivalent to `Get-Help .\deploy.ps1 -Full`.

.EXAMPLE
  .\deploy.ps1
  Prompts for a version tag, deploys GAS, rewrites the URLs, deploys to Netlify prod.

.EXAMPLE
  .\deploy.ps1 -Redeploy
  Re-pushes the existing netlify/index.html without touching the GAS deployment.

.EXAMPLE
  Get-Help .\deploy.ps1 -Full
  Shows the full help text via PowerShell.
#>

[CmdletBinding()]
param(
    [switch]$Redeploy,
    [switch]$Help
)

if ($Help) {
    Get-Help $MyInvocation.MyCommand.Path -Full | Out-String | Write-Host
    return
}

$ErrorActionPreference = 'Stop'

$repoRoot  = Split-Path -Parent $MyInvocation.MyCommand.Path
$indexPath = Join-Path $repoRoot 'netlify\index.html'

if (-not $Redeploy) {
    Write-Host '>> clasp list-deployments --json (current latest)'
    $currentJson = & clasp list-deployments --json
    if ($LASTEXITCODE -ne 0) { throw "clasp list-deployments failed ($LASTEXITCODE)" }

    $current = $currentJson | ConvertFrom-Json
    if ($null -eq $current) { $current = @() }
    if ($current -isnot [System.Array]) { $current = @($current) }

    $latest = $current | Sort-Object -Property versionNumber -Descending | Select-Object -First 1
    if ($latest) {
        Write-Host ">> current description: $($latest.description)"
        Write-Host ">> current id:          $($latest.deploymentId)"
    } else {
        Write-Host '>> no deployments yet'
    }

    $version = Read-Host 'Version tag for the deployment'
    if ([string]::IsNullOrWhiteSpace($version)) { throw 'Version is required.' }

Write-Host ">> clasp deploy -d $version"
& clasp deploy -d "$version" @args
if ($LASTEXITCODE -ne 0) { throw "clasp deploy failed ($LASTEXITCODE)" }

Write-Host '>> clasp list-deployments'
$deploymentsJson = & clasp list-deployments --json
if ($LASTEXITCODE -ne 0) { throw "clasp list-deployments failed ($LASTEXITCODE)" }

# clasp --json returns an array; tolerate either array or single-object shapes.
$deployments = $deploymentsJson | ConvertFrom-Json
if ($null -eq $deployments) { $deployments = @() }
if ($deployments -isnot [System.Array]) { $deployments = @($deployments) }

$match = $deployments | Where-Object { $_.deploymentId -and $_.description -eq $version } | Select-Object -First 1
if (-not $match) { throw "No deployment found with description '$version'." }

$newId = $match.deploymentId
Write-Host ">> deployment id: $newId"

$index = Get-Content -Raw -Path $indexPath
$urlPattern = 'https://script\.google\.com/macros/s/[^/]+/exec'
$updated = [regex]::Replace($index, $urlPattern, "https://script.google.com/macros/s/$newId/exec")

$versionPattern = '<p class="version">[^<]*</p>'
$updated = [regex]::Replace($updated, $versionPattern, "<p class=""version"">$version</p>")

if ($updated -eq $index) { Write-Warning 'No URL replaced — pattern may have shifted.' }

Set-Content -Path $indexPath -Value $updated -NoNewline
}

Write-Host '>> netlify deploy --prod'
& netlify deploy --dir=.\netlify --no-build --prod
if ($LASTEXITCODE -ne 0) { throw "netlify deploy failed ($LASTEXITCODE)" }

if ($Redeploy) {
    Write-Host 'Done. Redeploy only — netlify/index.html unchanged.'
} else {
    Write-Host "Done. Version='$version' id=$newId"
}
