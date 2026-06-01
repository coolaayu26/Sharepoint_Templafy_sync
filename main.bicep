// Bicep template for SharePoint → Templafy Sync Function App
// Deploy: az deployment group create -g <rg> -f main.bicep -p @params.json

@description('Base name prefix for all resources (e.g. "sptfsync")')
param baseName string = 'sptfsync'

@description('Azure region for all resources')
param location string = resourceGroup().location

@description('AAD App Registration tenant ID (for Graph API)')
@secure()
param graphTenantId string

@secure()
param graphClientId string

@secure()
param graphClientSecret string

param sharepointSiteId string
param sharepointDriveId string
param sharepointRootFolder string = '/'

param templafyTenantId string

@secure()
param templafyApiKey string

param templafySpaceId string

param dryRun bool = false
param maxFileSizeMb int = 100

// ── Storage Account (state blob + Function runtime) ──────────────────────────

resource storageAccount 'Microsoft.Storage/storageAccounts@2023-01-01' = {
  name: '${baseName}sa'
  location: location
  sku: { name: 'Standard_LRS' }
  kind: 'StorageV2'
  properties: {
    minimumTlsVersion: 'TLS1_2'
    allowBlobPublicAccess: false
    supportsHttpsTrafficOnly: true
  }
}

resource syncStateContainer 'Microsoft.Storage/storageAccounts/blobServices/containers@2023-01-01' = {
  name: '${storageAccount.name}/default/sync-state'
  properties: {
    publicAccess: 'None'
  }
}

// ── Key Vault ─────────────────────────────────────────────────────────────────

resource keyVault 'Microsoft.KeyVault/vaults@2023-07-01' = {
  name: '${baseName}-kv'
  location: location
  properties: {
    sku: { family: 'A', name: 'standard' }
    tenantId: subscription().tenantId
    enableSoftDelete: true
    softDeleteRetentionInDays: 7
    enableRbacAuthorization: true
  }
}

// Secrets
resource secretGraphClientSecret 'Microsoft.KeyVault/vaults/secrets@2023-07-01' = {
  parent: keyVault
  name: 'graph-client-secret'
  properties: { value: graphClientSecret }
}

resource secretTemplafyApiKey 'Microsoft.KeyVault/vaults/secrets@2023-07-01' = {
  parent: keyVault
  name: 'templafy-api-key'
  properties: { value: templafyApiKey }
}

// ── App Service Plan (Consumption) ────────────────────────────────────────────

resource appServicePlan 'Microsoft.Web/serverfarms@2023-01-01' = {
  name: '${baseName}-plan'
  location: location
  sku: { name: 'Y1', tier: 'Dynamic' }
  kind: 'functionapp'
  properties: { reserved: true }  // Linux
}

// ── Application Insights ──────────────────────────────────────────────────────

resource appInsights 'Microsoft.Insights/components@2020-02-02' = {
  name: '${baseName}-ai'
  location: location
  kind: 'web'
  properties: { Application_Type: 'web' }
}

// ── Function App ──────────────────────────────────────────────────────────────

resource functionApp 'Microsoft.Web/sites@2023-01-01' = {
  name: '${baseName}-func'
  location: location
  kind: 'functionapp,linux'
  identity: { type: 'SystemAssigned' }
  properties: {
    serverFarmId: appServicePlan.id
    siteConfig: {
      pythonVersion: '3.11'
      appSettings: [
        { name: 'AzureWebJobsStorage', value: 'DefaultEndpointsProtocol=https;AccountName=${storageAccount.name};AccountKey=${storageAccount.listKeys().keys[0].value};EndpointSuffix=core.windows.net' }
        { name: 'WEBSITE_CONTENTAZUREFILECONNECTIONSTRING', value: 'DefaultEndpointsProtocol=https;AccountName=${storageAccount.name};AccountKey=${storageAccount.listKeys().keys[0].value};EndpointSuffix=core.windows.net' }
        { name: 'WEBSITE_CONTENTSHARE', value: '${baseName}func' }
        { name: 'FUNCTIONS_EXTENSION_VERSION', value: '~4' }
        { name: 'FUNCTIONS_WORKER_RUNTIME', value: 'python' }
        { name: 'APPINSIGHTS_INSTRUMENTATIONKEY', value: appInsights.properties.InstrumentationKey }
        // Graph / SharePoint
        { name: 'GRAPH_TENANT_ID',        value: graphTenantId }
        { name: 'GRAPH_CLIENT_ID',         value: graphClientId }
        { name: 'GRAPH_CLIENT_SECRET',     value: '@Microsoft.KeyVault(SecretUri=${secretGraphClientSecret.properties.secretUri})' }
        { name: 'SHAREPOINT_SITE_ID',      value: sharepointSiteId }
        { name: 'SHAREPOINT_DRIVE_ID',     value: sharepointDriveId }
        { name: 'SHAREPOINT_ROOT_FOLDER',  value: sharepointRootFolder }
        // Templafy
        { name: 'TEMPLAFY_TENANT_ID',      value: templafyTenantId }
        { name: 'TEMPLAFY_API_KEY',        value: '@Microsoft.KeyVault(SecretUri=${secretTemplafyApiKey.properties.secretUri})' }
        { name: 'TEMPLAFY_SPACE_ID',       value: templafySpaceId }
        // State store
        { name: 'STATE_STORAGE_CONNECTION_STRING', value: 'DefaultEndpointsProtocol=https;AccountName=${storageAccount.name};AccountKey=${storageAccount.listKeys().keys[0].value};EndpointSuffix=core.windows.net' }
        { name: 'STATE_CONTAINER_NAME',    value: 'sync-state' }
        { name: 'STATE_BLOB_NAME',         value: 'sharepoint-templafy-state.json' }
        // Behaviour
        { name: 'DRY_RUN',               value: string(dryRun) }
        { name: 'MAX_FILE_SIZE_MB',       value: string(maxFileSizeMb) }
      ]
    }
  }
}

// Grant Function App managed identity access to Key Vault secrets
resource kvSecretUserRole 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  name: guid(keyVault.id, functionApp.id, 'KeyVaultSecretsUser')
  scope: keyVault
  properties: {
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', '4633458b-17de-408a-b874-0445c86b69e6') // Key Vault Secrets User
    principalId: functionApp.identity.principalId
    principalType: 'ServicePrincipal'
  }
}

output functionAppName string = functionApp.name
output keyVaultName string = keyVault.name
output storageAccountName string = storageAccount.name
