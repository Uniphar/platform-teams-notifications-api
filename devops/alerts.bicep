param environment string
param logAnalyticsWorkspaceId string

@allowed([
  'westeurope'
  'global'
  'westus'
  'centralindia'
])
param location string = 'westeurope'

resource platformEngineeringApplicationsLow 'Microsoft.Insights/actionGroups@2024-10-01-preview' existing = {
  name: 'platform-engineering-applications-low'
  scope: resourceGroup('observability')
}

#disable-next-line no-unused-existing-resources
resource platformEngineeringApplicationsHigh 'Microsoft.Insights/actionGroups@2024-10-01-preview' existing = {
  name: 'platform-engineering-applications-high'
  scope: resourceGroup('observability')
}

module exceptionDetectedAlert 'devops.alerts.scheduledqueryrules.bicep' = {
  name: 'PlatformTeamsNotificationApi-ExceptionDetectedAlert'
  params: {
    location: location
    alertName: 'PlatformTeamsNotificationApi-ExceptionDetectedAlert'
    environment: environment
    logAnalyticsWorkspaceId: logAnalyticsWorkspaceId
    query: '''AppExceptions
              | where AppRoleName == 'platform-teams-notification-api'
                and ExceptionType != "System.OperationCanceledException"
                and ExceptionType != "System.Threading.Tasks.TaskCanceledException"
           '''
    ActionGroupIds: [
      environment == 'prod' ? platformEngineeringApplicationsHigh.id : platformEngineeringApplicationsLow.id
    ]
  }
}
