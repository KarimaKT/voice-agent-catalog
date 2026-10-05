[CmdletBinding()]
param(
    [string]$PacPath = "pac",
    [string]$OutputPath
)

$ErrorActionPreference = "Stop"

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$sourcePath = Join-Path $repositoryRoot "copilot-studio\sample-agent\solution-source"
$packagesPath = Join-Path $repositoryRoot "copilot-studio\packages"
$targetPath = if ($OutputPath) {
    $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
} else {
    Join-Path $packagesPath "pat-manager-handoff-solution.zip"
}
$workingPath = Join-Path ([System.IO.Path]::GetTempPath()) "voice-agent-catalog-package-$([guid]::NewGuid().ToString('N'))"
$buildPath = Join-Path $workingPath "build"
$unpackPath = Join-Path $workingPath "unpacked"
$forbiddenPatterns = @(
    "@[A-Za-z0-9.-]+\.onmicrosoft\.com",
    "[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}",
    "https://[^/]+\.crm[0-9]*\.dynamics\.com",
    "/subscriptions/[0-9a-fA-F-]{36}",
    "C:\\Users\\",
    "DataverseEndpoint",
    "EnvironmentId",
    "AccountInfo",
    "AgentManagementEndpoint",
    "connectionReferenceLogicalName",
    "shared_office365",
    "[?&](sig|code|sv|se|sp)="
)

if (-not (Test-Path -LiteralPath $sourcePath -PathType Container)) {
    throw "Sample agent source was not found at $sourcePath."
}

function Assert-PublishableContent {
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [Parameter(Mandatory)]
        [string]$Description
    )

    $matches = Get-ChildItem -LiteralPath $Path -Recurse -File |
        Where-Object { $_.Extension -in @(".json", ".txt", ".xml", ".yml", ".yaml") } |
        Select-String -Pattern $forbiddenPatterns -AllMatches
    if ($matches) {
        $locations = $matches | ForEach-Object { "$($_.Path):$($_.LineNumber)" }
        throw "$Description contains values that must not be published:`n$($locations -join "`n")"
    }
}

Assert-PublishableContent -Path $sourcePath -Description "Sample agent source"
New-Item -ItemType Directory -Path $buildPath, $unpackPath | Out-Null

try {
    & $PacPath copilot pack `
        --publisher-prefix vac `
        --project-dir $sourcePath `
        --solution-name VoiceAgentCatalogPatSample `
        --output-path $buildPath
    if ($LASTEXITCODE -ne 0) {
        throw "pac copilot pack failed with exit code $LASTEXITCODE."
    }

    $generatedPath = Join-Path $buildPath "VoiceAgentCatalogPatSample.zip"
    if (-not (Test-Path -LiteralPath $generatedPath -PathType Leaf)) {
        throw "Power Platform CLI did not create the expected solution ZIP."
    }

    & $PacPath solution unpack `
        --zipfile $generatedPath `
        --folder $unpackPath `
        --packagetype Unmanaged `
        --errorlevel Error
    if ($LASTEXITCODE -ne 0) {
        throw "The generated solution ZIP could not be unpacked."
    }

    Assert-PublishableContent -Path $unpackPath -Description "Generated solution"

    $targetDirectory = Split-Path -Parent $targetPath
    New-Item -ItemType Directory -Force -Path $targetDirectory | Out-Null
    Copy-Item -LiteralPath $generatedPath -Destination $targetPath -Force
    Write-Output "Created and validated $targetPath"
} finally {
    if (Test-Path -LiteralPath $workingPath) {
        Remove-Item -LiteralPath $workingPath -Recurse -Force
    }
}
