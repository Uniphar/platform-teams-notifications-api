function Initialize-PlatformTeamsNotificationApi {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [parameter(Mandatory = $true, Position = 0)]
        [ValidateSet('dev', 'test', 'prod')]
        [string] $Environment,

        [parameter(Mandatory = $true)]
        [string] $SubscriptionId,

        [parameter(Mandatory = $true)]
        [string] $ResourceGroupName,

        [parameter(Mandatory = $true)]
        [string] $LogAnalyticsWorkspaceId
    )

    $alertsTemplate = Join-Path $PSScriptRoot -ChildPath ".\alerts.bicep"
    $deploymentName = "deploy-$(Get-Date -Format 'yyyyMMddHHmmss')-teams-notification-api-alerts"

    Select-AzSubscription -SubscriptionId $SubscriptionId

    New-AzResourceGroupDeployment `
        -Mode Incremental `
        -Name $deploymentName `
        -ResourceGroupName $ResourceGroupName `
        -TemplateFile $alertsTemplate `
        -environment $Environment `
        -logAnalyticsWorkspaceId $LogAnalyticsWorkspaceId `
        -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
}