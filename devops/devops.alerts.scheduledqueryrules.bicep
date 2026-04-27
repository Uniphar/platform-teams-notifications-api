param location string = resourceGroup().location
param alertName string
param environment string
param logAnalyticsWorkspaceId string
param query string
param ActionGroupIds string[]
param alertSeverity int = 3

resource alert 'Microsoft.Insights/scheduledQueryRules@2021-08-01' = {
  name: '${alertName}-${environment}'
  location: location
  properties: {
    severity: alertSeverity
    enabled: true
    scopes: [
      logAnalyticsWorkspaceId
    ]
    evaluationFrequency: environment == 'prod' ? 'PT5M' : 'PT60M'
    windowSize: environment == 'prod' ? 'PT5M' : 'PT60M'
    criteria: {
      allOf: [
        {
          query: query
          operator: 'GreaterThan'
          timeAggregation: 'Count'
          threshold: 0
          failingPeriods: {
            numberOfEvaluationPeriods: 1
            minFailingPeriodsToAlert: 1
          }
        }
      ]
    }
    autoMitigate: false
    actions: {
      actionGroups: ActionGroupIds
      customProperties: {}
    }
  }
}
