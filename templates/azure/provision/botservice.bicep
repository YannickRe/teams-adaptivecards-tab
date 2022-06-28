@secure()
param provisionParameters object
param botEndpoint string

var resourceBaseName = provisionParameters.resourceBaseName
var botAadAppClientId = provisionParameters['botAadAppClientId'] // Read AAD app client id for Azure Bot Service from parameters
var m365ClientId = provisionParameters['m365ClientId']
var m365ClientSecret = provisionParameters['m365ClientSecret']
var m365TenantId = provisionParameters['m365TenantId']
var botServiceName = contains(provisionParameters, 'botServiceName') ? provisionParameters['botServiceName'] : '${resourceBaseName}' // Try to read name for Azure Bot Service from parameters
var botServiceSku = contains(provisionParameters, 'botServiceSku') ? provisionParameters['botServiceSku'] : 'F0' // Try to read SKU for Azure Bot Service from parameters
var botDisplayName = contains(provisionParameters, 'botDisplayName') ? provisionParameters['botDisplayName'] : '${resourceBaseName}' // Try to read display name for Azure Bot Service from parameters

// Register your web service as a bot with the Bot Framework
resource botService 'Microsoft.BotService/botServices@2021-03-01' = {
  kind: 'azurebot'
  location: 'global'
  name: botServiceName
  properties: {
    displayName: botDisplayName
    endpoint: uri(botEndpoint, '/api/messages')
    msaAppId: botAadAppClientId
  }
  sku: {
    name: botServiceSku // You can follow https://aka.ms/teamsfx-bicep-add-param-tutorial to add botServiceSku property to provisionParameters to override the default value "F0".
  }
}

// Connect the bot service to Microsoft Teams
resource botServiceMsTeamsChannel 'Microsoft.BotService/botServices/channels@2021-03-01' = {
  parent: botService
  location: 'global'
  name: 'MsTeamsChannel'
  properties: {
    channelName: 'MsTeamsChannel'
  }
}

// Add OAUTH2 Connection Settings
resource botServiceOAuth2 'Microsoft.BotService/botServices/connections@2018-07-12' = {
  parent: botService
  location: 'global'
  name: 'AADConnection'
  properties: {
    scopes: 'openid profile'
    serviceProviderId: '30dd229c-58e3-4a48-bdfd-91ec48eb906c'
    clientId: m365ClientId
    clientSecret: m365ClientSecret
    parameters: [
        {
            key: 'clientId'
            value: m365ClientId
        }
        {
            key: 'clientSecret'
            value: m365ClientSecret
        }
        {
            key: 'tenantId'
            value: m365TenantId
        }
    ]
  }
}