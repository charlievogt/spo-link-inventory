// spo-link-inventory — sub-level deployment
//
// Creates: resource group, storage account, function app + system-assigned MI,
// application insights, log analytics workspace, role assignments, app settings.
//
// Run with:
//   az deployment sub create \
//     --location <region> \
//     --template-file infra/main.bicep \
//     --parameters infra/main.parameters.json
//
// Prerequisites: resource providers must be registered on the subscription
// (see scripts/register-providers.sh — one-shot for new subscriptions).
//
// Outputs include the function app's MI principal ID, which feeds into
// scripts/add-federated-credential.sh on the SP-tenant side.

targetScope = 'subscription'

@description('Azure region for all resources.')
param location string = 'centralus'

@description('Short project identifier — used as the suffix in resource names. Lowercase alphanumeric, 3-15 chars.')
@minLength(3)
@maxLength(15)
param projectName string

@description('SharePoint tenant hostname (no protocol, no trailing slash). Example: contoso.sharepoint.com')
param spoTenantHost string

@description('GUID of the Entra tenant where the SharePoint app registration lives (typically the SharePoint tenant).')
param entraTenantId string

@description('Application (client) ID of the Entra app registration used for OBO. Created by scripts/setup-entra.sh.')
param entraClientId string

@description('Object ID of the Entra group whose members are admins. Created by scripts/setup-entra.sh.')
param adminGroupId string

@description('Comma-separated list of legacy/on-prem SharePoint hostnames classified as "onprem" (optional).')
param legacySpHosts string = ''

@description('Auth mode for the function app: "cert" (default, any-tenant) or "federation" (same-tenant only).')
@allowed([
  'cert'
  'federation'
])
param authMode string = 'cert'

@description('Combined PEM (cert + private key, base64-encoded). Required when authMode=cert. Pass via --parameters @cert.json — do NOT commit to source.')
@secure()
param clientCertPemBase64 string = ''

@description('Resource group name. Defaults to rg-<projectName>.')
param resourceGroupName string = 'rg-${projectName}'

resource rg 'Microsoft.Resources/resourceGroups@2024-03-01' = {
  name: resourceGroupName
  location: location
  tags: {
    purpose: 'link-inventory'
    project: projectName
  }
}

module resources 'app.bicep' = {
  name: 'app-${projectName}'
  scope: rg
  params: {
    location: location
    projectName: projectName
    spoTenantHost: spoTenantHost
    entraTenantId: entraTenantId
    entraClientId: entraClientId
    adminGroupId: adminGroupId
    legacySpHosts: legacySpHosts
    authMode: authMode
    clientCertPemBase64: clientCertPemBase64
  }
}

@description('Function App hostname for SPFx web part property "functionUrl" (prefix with https://).')
output functionAppHostname string = resources.outputs.functionAppHostname

@description('Function App name — pass to "func azure functionapp publish".')
output functionAppName string = resources.outputs.functionAppName

@description('Managed Identity principal (object) ID — feeds into scripts/add-federated-credential.sh.')
output miPrincipalId string = resources.outputs.miPrincipalId

@description('Azure subscription tenant ID — issuer for the federated credential.')
output azureTenantId string = subscription().tenantId

@description('Storage account name (for diagnostics).')
output storageAccountName string = resources.outputs.storageAccountName
