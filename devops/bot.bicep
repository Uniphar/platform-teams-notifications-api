param botName string
param teamsBotAppId string
param tenantId string = tenant().tenantId
param botDisplayName string = 'Teams Notifications'
param environment string
param endpoint string
// Data residency
@allowed([
  'westeurope'
  'global'
  'westus'
  'centralindia'
])
param location string = 'westeurope'
param logAnalyticsWorkspaceId string

resource botService 'Microsoft.BotService/botServices@2023-09-15-preview' = {
  name: botName
  kind: 'azurebot'
  location: location
  sku: {
    name: 'S1'
  }
  properties: {
    displayName: 'Bot for ${botDisplayName} ${environment}'
    msaAppId: teamsBotAppId
    msaAppType: 'SingleTenant'
    msaAppTenantId: tenantId
    endpoint: endpoint
  }
}

resource teamsChannel 'Microsoft.BotService/botServices/channels@2023-09-15-preview' = {
  parent: botService
  name: 'MsTeamsChannel'
  location: location
  properties: {
    channelName: 'MsTeamsChannel'
    properties: {
      enableCalling: false
      incomingCallRoute: 'graphPma'
      callingWebhook: null
      isEnabled: true
      deploymentEnvironment: 'CommercialDeployment'
      acceptedTerms: true
    }
  }
}

resource botDiagnostics 'Microsoft.Insights/diagnosticSettings@2021-05-01-preview' = {
  name: '${botName}-diagnostics'
  scope: botService
  properties: {
    workspaceId: logAnalyticsWorkspaceId
    logs: [
      {
        categoryGroup: 'allLogs'
        enabled: true
      }
    ]
    metrics: [
      {
        category: 'AllMetrics'
        enabled: true
      }
    ]
  }
}
