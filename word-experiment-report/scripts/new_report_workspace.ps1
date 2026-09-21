param(
    [Parameter(Mandatory=$true)]
    [string]$ProjectRoot,
    [string]$TemplatePath,
    [switch]$CopyMathModelReference
)

$ErrorActionPreference = 'Stop'
$skillRoot = Split-Path -Parent $PSScriptRoot
$dirs = @('input','references','code','screenshots','working','output','qa')
foreach ($name in $dirs) {
    New-Item -ItemType Directory -Force -Path (Join-Path $ProjectRoot $name) | Out-Null
}
Set-Content -LiteralPath (Join-Path $ProjectRoot 'README.md') -Value @"
# Word report project

Use input/ for the guide, references/ for templates, code/ for rerunnable source, screenshots/ for evidence, working/ for drafts, output/ for final DOCX/PDF, and qa/ for render checks.
"@ -Encoding UTF8

if ($TemplatePath) {
    if (-not (Test-Path -LiteralPath $TemplatePath -PathType Leaf)) { throw "Template not found: $TemplatePath" }
    Copy-Item -LiteralPath $TemplatePath -Destination (Join-Path $ProjectRoot 'references') -Force
}
if ($CopyMathModelReference) {
    $mathTemplate = Join-Path $skillRoot 'workspace-template\references\数学建模国赛标准论文Word模板.doc'
    Copy-Item -LiteralPath $mathTemplate -Destination (Join-Path $ProjectRoot 'references') -Force
}
Write-Output "Created report workspace: $ProjectRoot"


