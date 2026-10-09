param([string] $Destination = ".azure\release-stage")

$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot
$target = [IO.Path]::GetFullPath((Join-Path $root $Destination))
$ignored = & git -C $root check-ignore -- $target
if ($LASTEXITCODE -ne 0 -or !$ignored) { throw "The release stage must be Git-ignored." }
if (Test-Path $target) { throw "The release stage already exists. Use a new explicitly named stage." }
New-Item -ItemType Directory -Path $target | Out-Null
foreach ($directory in @("src", "tests", "catalog", "env", "appPackage")) {
  Copy-Item (Join-Path $root $directory) (Join-Path $target $directory) -Recurse
}
foreach ($file in @(
  "package.json", "package-lock.json", "tsconfig.json", "tsconfig.node.json",
  "tsconfig.app.json", "tsup.config.js", "vite.config.js", "index.html", "auth-callback.html",
  "web.config", ".webappignore", "m365agents.yml", "aad.manifest.json",
  "aad.api.manifest.json"
)) {
  Copy-Item (Join-Path $root $file) $target
}
[pscustomobject]@{ Stage = $target; DependenciesCopied = $false; EnvironmentFilesRemainIgnored = $true }
