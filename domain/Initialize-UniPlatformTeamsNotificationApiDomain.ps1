function Initialize-UniPlatformTeamsNotificationApiDomain {
    <#
    .SYNOPSIS
    Deploys the PlatformTeamsNotificationApi domain.

    .DESCRIPTION
    Deploys the PlatformTeamsNotificationApi domain, including Azure Monitor scheduled query rule alerts
    for exception detection.

    .PARAMETER Environment
    The environment to deploy the domain to.

    .EXAMPLE
    Initialize-UniPlatformTeamsNotificationApi -Environment dev
    Initializes the PPlatformTeamsNotificationApi domain in the dev environment.

    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [parameter(Mandatory = $true, Position = 0)]
        [ValidateSet('dev', 'test', 'prod')]
        [string] $Environment
    )

    $domainTemplateFile = Join-Path $PSScriptRoot -ChildPath ".\domain.bicep"

    $devopsDomainRgName = Resolve-UniResourceName 'resource-group' $p_devopsDomain -Environment $Environment
    $logAnalyticsWorkspace = Resolve-UniMainLogAnalytics -Environment $Environment
    $deploymentName = Resolve-DeploymentName -Suffix '-PlatformTeamsNotificationApi'

    New-AzResourceGroupDeployment -Mode Incremental `
        -Name $deploymentName `
        -ResourceGroupName $devopsDomainRgName `
        -logAnalyticsWorkspaceId $logAnalyticsWorkspace.ResourceId `
        -environment $Environment `
        -TemplateFile $domainTemplateFile
   
}
