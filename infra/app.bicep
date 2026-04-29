// RG-scoped resources for spo-link-inventory.
// Invoked by main.bicep — not meant to run standalone.

param location string
param projectName string
param spoTenantHost string
param entraTenantId string
param entraClientId string
param adminGroupId string
param legacySpHosts string
param authMode string
@secure()
param clientCertPemBase64 string

// ---- Storage Account ----------------------------------------------------

// Storage account names: 3-24 chars, lowercase alphanumeric only.
var storageAccountName = take('st${replace(projectName, '-', '')}${uniqueString(resourceGroup().id)}', 24)

resource storageAccount 'Microsoft.Storage/storageAccounts@2023-05-01' = {
  name: storageAccountName
  location: location
  sku: { name: 'Standard_LRS' }
  kind: 'StorageV2'
  properties: {
    minimumTlsVersion: 'TLS1_2'
    allowBlobPublicAccess: false
    supportsHttpsTrafficOnly: true
  }
  tags: { purpose: 'link-inventory' }
}

// ---- Application Insights -----------------------------------------------

resource logWorkspace 'Microsoft.OperationalInsights/workspaces@2023-09-01' = {
  name: 'log-${projectName}'
  location: location
  properties: {
    sku: { name: 'PerGB2018' }
    retentionInDays: 30
  }
}

resource appInsights 'Microsoft.Insights/components@2020-02-02' = {
  name: 'appi-${projectName}'
  location: location
  kind: 'web'
  properties: {
    Application_Type: 'web'
    WorkspaceResourceId: logWorkspace.id
  }
}

// ---- Function App (Linux Consumption) -----------------------------------

resource hostingPlan 'Microsoft.Web/serverfarms@2023-12-01' = {
  name: 'plan-${projectName}'
  location: location
  sku: { name: 'Y1', tier: 'Dynamic' }
  kind: 'linux'
  properties: { reserved: true }
}

var storageConnectionString = 'DefaultEndpointsProtocol=https;AccountName=${storageAccount.name};AccountKey=${storageAccount.listKeys().keys[0].value};EndpointSuffix=${environment().suffixes.storage}'

resource functionApp 'Microsoft.Web/sites@2023-12-01' = {
  name: 'func-${projectName}'
  location: location
  kind: 'functionapp,linux'
  identity: { type: 'SystemAssigned' }
  properties: {
    serverFarmId: hostingPlan.id
    siteConfig: {
      linuxFxVersion: 'NODE|20'
      cors: {
        allowedOrigins: [
          'https://${spoTenantHost}'
          'https://localhost:4321'
        ]
        supportCredentials: false
      }
      appSettings: concat(
        [
          { name: 'AzureWebJobsStorage', value: storageConnectionString }
          { name: 'TABLE_CONNECTION_STRING', value: storageConnectionString }
          { name: 'FUNCTIONS_EXTENSION_VERSION', value: '~4' }
          { name: 'FUNCTIONS_WORKER_RUNTIME', value: 'node' }
          { name: 'WEBSITE_RUN_FROM_PACKAGE', value: '1' }
          { name: 'APPLICATIONINSIGHTS_CONNECTION_STRING', value: appInsights.properties.ConnectionString }
          { name: 'SPO_TENANT_HOST', value: spoTenantHost }
          { name: 'LINK_INVENTORY_TENANT_ID', value: entraTenantId }
          { name: 'LINK_INVENTORY_CLIENT_ID', value: entraClientId }
          { name: 'LINK_INVENTORY_ADMIN_GROUP_ID', value: adminGroupId }
          { name: 'LINK_INVENTORY_AUTH_MODE', value: authMode }
          { name: 'LEGACY_SP_HOSTS', value: legacySpHosts }
        ],
        // Cert mode: ship the combined PEM as an app setting. For
        // production, replace this with a Key Vault reference of the
        // form @Microsoft.KeyVault(SecretUri=...) — Bicep can't enforce
        // that here, so it's documented in docs/deploy.md.
        authMode == 'cert' ? [
          { name: 'LINK_INVENTORY_CLIENT_CERT_PEM_BASE64', value: clientCertPemBase64 }
        ] : []
      )
      ftpsState: 'Disabled'
      minTlsVersion: '1.2'
      http20Enabled: true
    }
    httpsOnly: true
  }
  tags: { purpose: 'link-inventory' }
}

// ---- Role assignments: MI on Storage ------------------------------------

// Built-in role definition GUIDs
var storageBlobDataContributor = 'ba92f5b4-2d11-453d-a403-e96b0029c9fe'
var storageQueueDataContributor = '974c5e8b-45b9-4653-ba55-5f855dd0fb88'
var storageTableDataContributor = '0a9a7e1f-b9d0-4cc4-a60d-0319b160aaa3'

resource roleBlobDataContributor 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  scope: storageAccount
  name: guid(storageAccount.id, functionApp.id, storageBlobDataContributor)
  properties: {
    principalId: functionApp.identity.principalId
    principalType: 'ServicePrincipal'
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', storageBlobDataContributor)
  }
}

resource roleQueueDataContributor 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  scope: storageAccount
  name: guid(storageAccount.id, functionApp.id, storageQueueDataContributor)
  properties: {
    principalId: functionApp.identity.principalId
    principalType: 'ServicePrincipal'
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', storageQueueDataContributor)
  }
}

resource roleTableDataContributor 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  scope: storageAccount
  name: guid(storageAccount.id, functionApp.id, storageTableDataContributor)
  properties: {
    principalId: functionApp.identity.principalId
    principalType: 'ServicePrincipal'
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', storageTableDataContributor)
  }
}

// ---- Outputs ------------------------------------------------------------

output functionAppName string = functionApp.name
output functionAppHostname string = functionApp.properties.defaultHostName
output miPrincipalId string = functionApp.identity.principalId
output storageAccountName string = storageAccount.name
