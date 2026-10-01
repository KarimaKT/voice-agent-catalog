@maxLength(20)
@minLength(4)
@description('Base name used for the POC resources')
param resourceBaseName string

@description('App Service plan SKU')
param webAppSku string = 'B1'

@description('Microsoft Entra tenant that contains the agent')
param tenantId string

@description('Client ID for the browser SPA')
param clientId string

@description('Copilot Studio environment ID')
param copilotEnvironmentId string

@description('Copilot Studio agent schema name')
param copilotSchemaName string

@secure()
@description('Signed Power Automate HTTP trigger URL for the maker-owned catalog flow')
param catalogFlowUrl string = ''

param location string = resourceGroup().location
param speechLocale string = 'en-US'
param speechVoiceName string = 'en-US-AvaMultilingualNeural'
param avatarCharacter string = 'lisa'
param avatarStyle string = 'casual-sitting'
@allowed([
  'false'
  'true'
])
param demoMode string = 'false'

var tags = {
  application: 'voice-agent-catalog'
  environment: 'poc'
}
var planName = '${resourceBaseName}-plan'
var webAppName = '${resourceBaseName}-app'
var speechName = '${resourceBaseName}-speech'
var keyVaultName = take('${resourceBaseName}-kv', 24)
var logAnalyticsName = '${resourceBaseName}-logs'
var appInsightsName = '${resourceBaseName}-insights'
var speechSecretName = 'speech-key'
var catalogFlowSecretName = 'catalog-flow-url'
var catalogFlowEnabled = !empty(catalogFlowUrl) && catalogFlowUrl != 'disabled'

resource serverfarm 'Microsoft.Web/serverfarms@2022-09-01' = {
  name: planName
  location: location
  kind: 'linux'
  tags: tags
  sku: {
    name: webAppSku
  }
  properties: {
    reserved: true
  }
}

resource speech 'Microsoft.CognitiveServices/accounts@2023-05-01' = {
  name: speechName
  location: location
  kind: 'SpeechServices'
  tags: tags
  sku: {
    name: 'S0'
  }
  properties: {
    customSubDomainName: speechName
    publicNetworkAccess: 'Enabled'
    networkAcls: {
      defaultAction: 'Allow'
    }
  }
}

resource keyVault 'Microsoft.KeyVault/vaults@2023-07-01' = {
  name: keyVaultName
  location: location
  tags: tags
  properties: {
    tenantId: tenantId
    sku: {
      family: 'A'
      name: 'standard'
    }
    enableRbacAuthorization: false
    enableSoftDelete: true
    softDeleteRetentionInDays: 7
    enablePurgeProtection: true
    publicNetworkAccess: 'Enabled'
    accessPolicies: [
      {
        tenantId: tenantId
        objectId: webApp.identity.principalId
        permissions: {
          secrets: [
            'get'
          ]
        }
      }
    ]
  }
}

resource speechKeySecret 'Microsoft.KeyVault/vaults/secrets@2023-07-01' = {
  parent: keyVault
  name: speechSecretName
  properties: {
    value: speech.listKeys().key1
  }
}

resource catalogFlowSecret 'Microsoft.KeyVault/vaults/secrets@2023-07-01' = if (catalogFlowEnabled) {
  parent: keyVault
  name: catalogFlowSecretName
  properties: {
    value: catalogFlowUrl
  }
}

resource logAnalytics 'Microsoft.OperationalInsights/workspaces@2022-10-01' = {
  name: logAnalyticsName
  location: location
  tags: tags
  properties: {
    sku: {
      name: 'PerGB2018'
    }
    retentionInDays: 30
  }
}

resource applicationInsights 'Microsoft.Insights/components@2020-02-02' = {
  name: appInsightsName
  location: location
  kind: 'web'
  tags: tags
  properties: {
    Application_Type: 'web'
    WorkspaceResourceId: logAnalytics.id
  }
}

resource webApp 'Microsoft.Web/sites@2022-09-01' = {
  name: webAppName
  location: location
  kind: 'app,linux'
  tags: union(tags, { 'azd-service-name': 'web' })
  identity: {
    type: 'SystemAssigned'
  }
  properties: {
    serverFarmId: serverfarm.id
    httpsOnly: true
    siteConfig: {
      linuxFxVersion: 'NODE|22-lts'
      alwaysOn: webAppSku != 'F1'
      appCommandLine: 'npm start'
      healthCheckPath: '/api/health'
      http20Enabled: true
      minTlsVersion: '1.2'
      ftpsState: 'Disabled'
      appSettings: [
        {
          name: 'WEBSITE_RUN_FROM_PACKAGE'
          value: '1'
        }
        {
          name: 'NODE_ENV'
          value: 'production'
        }
        {
          name: 'TENANT_ID'
          value: tenantId
        }
        {
          name: 'AAD_APP_CLIENT_ID'
          value: clientId
        }
        {
          name: 'COPILOT_ENVIRONMENT_ID'
          value: copilotEnvironmentId
        }
        {
          name: 'COPILOT_SCHEMA_NAME'
          value: copilotSchemaName
        }
        {
          name: 'DEMO_MODE'
          value: demoMode
        }
        {
          name: 'CATALOG_FLOW_URL'
          value: catalogFlowEnabled ? '@Microsoft.KeyVault(VaultName=${keyVaultName};SecretName=${catalogFlowSecretName})' : ''
        }
        {
          name: 'SPEECH_REGION'
          value: location
        }
        {
          name: 'SPEECH_ENDPOINT'
          value: speech.properties.endpoint
        }
        {
          name: 'SPEECH_KEY'
          value: '@Microsoft.KeyVault(VaultName=${keyVaultName};SecretName=${speechSecretName})'
        }
        {
          name: 'SPEECH_LOCALE'
          value: speechLocale
        }
        {
          name: 'SPEECH_VOICE_NAME'
          value: speechVoiceName
        }
        {
          name: 'AVATAR_CHARACTER'
          value: avatarCharacter
        }
        {
          name: 'AVATAR_STYLE'
          value: avatarStyle
        }
        {
          name: 'APPLICATIONINSIGHTS_CONNECTION_STRING'
          value: applicationInsights.properties.ConnectionString
        }
      ]
    }
  }
}

output AZURE_APP_SERVICE_RESOURCE_ID string = webApp.id
output TAB_DOMAIN string = webApp.properties.defaultHostName
output TAB_ENDPOINT string = 'https://${webApp.properties.defaultHostName}'
output SPEECH_RESOURCE_NAME string = speech.name
output KEY_VAULT_NAME string = keyVault.name
output APPLICATION_INSIGHTS_NAME string = applicationInsights.name
