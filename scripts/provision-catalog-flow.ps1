param(
  [Parameter(Mandatory)] [guid] $EnvironmentId,
  [Parameter(Mandatory)] [uri] $SiteUrl,
  [Parameter(Mandatory)] [guid] $ListId,
  [Parameter(Mandatory)] [string] $ConnectionName,
  [Parameter(Mandatory)] [uri] $CatalogUrl,
  [Parameter(Mandatory)] [string] $FieldMapPath,
  [string] $MetadataConnectionName,
  [string] $TemplatePath = (Join-Path $PSScriptRoot "..\catalog\catalog-flow.template.json"),
  [string] $StatePath = ".azure\catalog-flow.json"
)

$ErrorActionPreference = "Stop"
if ($SiteUrl.Scheme -ne "https" -or $CatalogUrl.Scheme -ne "https") {
  throw "The SharePoint site and catalog URLs must use HTTPS."
}
$stateFile = [IO.Path]::GetFullPath($StatePath)
$ignored = & git check-ignore -- $stateFile
if ($LASTEXITCODE -ne 0 -or !$ignored) {
  throw "The flow state file must be inside a Git-ignored directory."
}
$fields = Get-Content $FieldMapPath -Raw | ConvertFrom-Json -AsHashtable
$template = Get-Content $templatePath -Raw
if ($template.Contains("__DATAVERSE_CONNECTION__") -and !$MetadataConnectionName) {
  throw "The registration template requires an authenticated maker-owned Dataverse connection."
}
$replacements = @{
  "__SITE_URL__" = $SiteUrl.AbsoluteUri.TrimEnd("/")
  "__LIST_ID__" = $ListId.ToString()
  "__CONNECTION_NAME__" = $ConnectionName
  "__CATALOG_URL__" = $CatalogUrl.AbsoluteUri
  "__ENVIRONMENT_ID__" = $EnvironmentId.ToString()
  "__DATAVERSE_CONNECTION__" = $MetadataConnectionName
}
foreach ($field in @("Description", "EnvironmentId", "SchemaName", "Enabled", "CompletionPhrase", "EndsConversation", "WelcomeMessage")) {
  $name = $fields[$field]
  if (!$name -or $name -notmatch '^[A-Za-z_][A-Za-z0-9_]*$') {
    throw "The field map has no valid internal name for $field."
  }
  $replacements["__FIELD_$($field.ToUpperInvariant())__"] = $name
}
foreach ($key in $replacements.Keys) {
  $escaped = ConvertTo-Json ([string]$replacements[$key]) -Compress
  $template = $template.Replace($key, $escaped.Substring(1, $escaped.Length - 2))
}
if ($template -match "__[A-Z_]+__") { throw "The flow template has unresolved placeholders." }
$definition = $template | ConvertFrom-Json
$token = & az account get-access-token --resource "https://service.flow.microsoft.com/" --query accessToken -o tsv
if ($LASTEXITCODE -ne 0 -or !$token) { throw "Power Automate authentication failed." }
$headers = @{ Authorization = "Bearer $token" }
$base = "https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/$EnvironmentId/flows"

if (Test-Path $stateFile) {
  $state = Get-Content $stateFile -Raw | ConvertFrom-Json
  if ($state.environmentId -ne $EnvironmentId.ToString()) { throw "Existing flow state belongs to a different environment." }
  $flowId = $state.flowId
  $null = Invoke-RestMethod "$base/$flowId`?api-version=2016-11-01" -Headers $headers -TimeoutSec 30
  $flow = Invoke-RestMethod "$base/$flowId`?api-version=2016-11-01" -Method Patch -Headers $headers -ContentType "application/json" -Body $template -TimeoutSec 60
} else {
  $flow = Invoke-RestMethod "$base`?api-version=2016-11-01" -Method Post -Headers $headers -ContentType "application/json" -Body $template -TimeoutSec 60
  $flowId = $flow.name
}
if (!$flowId) { throw "Power Automate did not return a flow ID." }
New-Item -ItemType Directory -Force (Split-Path $stateFile) | Out-Null
@{ environmentId = $EnvironmentId.ToString(); flowId = $flowId } |
  ConvertTo-Json | Set-Content $stateFile -Encoding utf8NoBOM
if (!$definition.properties.definition.triggers.manual) {
  [pscustomobject]@{ FlowId = $flowId; StateFile = $stateFile; Trigger = "SharePoint item created" }
  return
}
$callback = Invoke-RestMethod "$base/$flowId/triggers/manual/listCallbackUrl?api-version=2016-11-01" -Method Post -Headers $headers -ContentType "application/json" -Body "{}" -TimeoutSec 30
$callbackUrl = $callback.value
if (!$callbackUrl) { $callbackUrl = $callback.response.value }
if (!$callbackUrl) { throw "The flow callback URL is unavailable." }
@{ environmentId = $EnvironmentId.ToString(); flowId = $flowId; callbackUrl = $callbackUrl } |
  ConvertTo-Json | Set-Content $stateFile -Encoding utf8NoBOM
[pscustomobject]@{ FlowId = $flowId; StateFile = $stateFile; SecretSaved = $true }
