param environment string
param logAnalyticsWorkspaceId string

param location string = resourceGroup().location
resource platformEngineeringApplicationsLow 'Microsoft.Insights/actionGroups@2024-10-01-preview' existing = {
  name: 'platform-engineering-applications-low'
  scope: resourceGroup('observability')
}

#disable-next-line no-unused-existing-resources
resource platformEngineeringApplicationsHigh 'Microsoft.Insights/actionGroups@2024-10-01-preview' existing = {
  name: 'platform-engineering-applications-high'
  scope: resourceGroup('observability')
}

module podNotesControllerExceptionDetectedAlert 'devops.alerts.scheduledqueryrules.bicep' = {
  name: 'PlatformTeamsNotificationApi-ExceptionDetectedAlert'
  params: {
    location: location
    alertName: 'PlatformTeamsNotificationApi-ExceptionDetectedAlert'
    environment: environment
    logAnalyticsWorkspaceId: logAnalyticsWorkspaceId
    query: '''AppExceptions
              | where AppRoleName == 'platform-teams-notification-api'
                  and ProblemId !contains "System.OperationCanceledException"
                  and ProblemId !contains "System.Threading.Tasks.TaskCanceledException"
           '''
    ActionGroupIds: [
      environment == 'prod' ? platformEngineeringApplicationsHigh.id : platformEngineeringApplicationsLow.id
      // for now low is more than enough, we can always add high later if needed
      //platformEngineeringApplicationsHigh.id
    ]
  }
}
