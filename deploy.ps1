#Requires -Version 5.1
<#
.SYNOPSIS
  Deploy GAS script, swap the deployment ID in netlify/index.html, push to Netlify.

.DESCRIPTION
  Prompts for a version string, runs `clasp deploy -d <version>`, resolves the
  new deployment id via `clasp list-deployments`, rewrites both Google Apps
  Script URLs in netlify/index.html, and ships the site to Netlify prod.

  Run from the repo root in pwsh.
#>

$ErrorActionPreference = 'Stop'

$repoRoot  = Split-Path -Parent $MyInvocation.MyCommand.Path
$indexPath = Join-Path $repoRoot 'netlify\index.html'

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

if ($updated -eq $index) { Write-Warning 'No URL replaced — pattern may have shifted.' }

Set-Content -Path $indexPath -Value $updated -NoNewline
Write-Host '>> netlify deploy --prod'
& netlify deploy --dir=.\netlify --no-build --prod
if ($LASTEXITCODE -ne 0) { throw "netlify deploy failed ($LASTEXITCODE)" }

Write-Host "Done. Version='$version' id=$newId"
