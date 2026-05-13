function Initialize-PlatformTeamsNotificationApi {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [parameter(Mandatory = $true, Position = 0)]
        [ValidateSet('dev')]
        [string] $Environment,
        
        [parameter(Mandatory = $true)]
        [string] $SubscriptionId,
        
        [parameter(Mandatory = $true)]
        [string] $ServicePrincipalName,
        
        [parameter(Mandatory = $true)]
        [string] $ResourceGroupName,
        
        [parameter(Mandatory = $true)]
        [string] $LogAnalyticsWorkspaceId
    )
    $botTemplate = Join-Path $PSScriptRoot -ChildPath ".\bot.bicep"
    
    Select-AzSubscription -SubscriptionId $SubscriptionId
    $envSA = Get-AzADServicePrincipal -DisplayName $ServicePrincipalName
    if (!$envSA) {
        throw "Service principal '$ServicePrincipalName' not found"
    }
    
    # Minimal application permissions required for:
    # - Installing a Teams app into any team, channel, or chat
    # - Reading all messages (teams/channel/chats)
    # - Adding users to teams/channels/chats
    # - Creating teams
    # - Reading/writing files in channels

    $botPermissionsNeeded = @(
        # Group control to create groups, update membership, and manage private/shared channels.
        "Group.ReadWrite.All"
        # Create new Microsoft Teams
        "Team.Create"
        # Read and update team settings
        "TeamSettings.ReadWrite.All"
        # Not everything is fully covered by Group.ReadWrite.All
        "TeamMember.ReadWrite.All"
        # Add/remove users from channels, including private/shared channels
        "ChannelMember.ReadWrite.All"
        # Read and update channels
        "ChannelSettings.ReadWrite.All"
        # Read *all* channel messages across the entire tenant (standard + private + shared)
        "ChannelMessage.Read.All"
        # Send messages as the bot/app into any channel
        "ChannelMessage.Send"
        # Read/write all chat messages (1:1 + group chats) and send messages
        "Chat.ReadWrite.All"
        # Add/remove members from chats (1:1 or multiparty)
        "ChatMember.ReadWrite.All"
        # Needed to get the teams app ids from the app catalog, which we have installed already
        "AppCatalog.ReadWrite.All"
        # read all installed Teams apps in the tenant
        "TeamsAppInstallation.Read.All"
        # install/uninstall a Teams app for a USER and read installed apps
        "TeamsAppInstallation.ReadWriteAndConsentSelfForUser.All"
        # install/uninstall a Teams app for a CHAT and read installed apps
        "TeamsAppInstallation.ReadWriteAndConsentSelfForChat.All"
        # to install a Teams app for a TEAM and read installed apps
        "TeamsAppInstallation.ReadWriteAndConsentSelfForTeam.All"
        # Read and write files in Teams channels (SharePoint-backed)
        "Files.ReadWrite.All"
        # Read user profiles – needed for resolving user IDs to add to teams/chats/channels
        "User.Read.All"
    )
    $deploymentName = "deploy-$(Get-Date -Format 'yyyyMMddHHmmss')-teams-notification-api-bot"
    $token = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com/" -AsSecureString
    Connect-MgGraph -AccessToken $token.Token -NoWelcome
    # we need a bot service which is not the workload identity, but a separate app registration (so we can get its secret for local debugging)
    # so to do this we will create a debug bot in the dev environment and give it the same permissions as the workload identity
    if ($Environment -eq 'dev') {
        $teamsBotNameDebug = "platform-teams-notification-api-debug-bot"
        $teamsBotEntraIdAppDebug = Get-AzADApplication -DisplayName $teamsBotNameDebug
        if (!$teamsBotEntraIdAppDebug) { 
            $teamsBotEntraIdAppDebug = New-AzADApplication -DisplayName $teamsBotNameDebug -SigninAudience AzureADMyOrg 
        }
        $teamsBotEntraIdServicePrincipalDebug = Get-AzADServicePrincipal -ApplicationId $teamsBotEntraIdAppDebug.AppId

        if (!$teamsBotEntraIdServicePrincipalDebug) {
            $teamsBotEntraIdServicePrincipalDebug = New-AzADServicePrincipal -ApplicationId $teamsBotEntraIdAppDebug.AppId 
        }
        
        $debugDeploymentConfig = @{
            Mode                    = 'Incremental'
            Name                    = $deploymentName
            ResourceGroupName       = $ResourceGroupName
            TemplateFile            = $botTemplate
            endpoint                = 'https://XXXXX.devtunnels.ms/platform-teams-notification-api/api/messages'
            environment             = 'debug'
            botName                 = $teamsBotNameDebug
            teamsBotAppId           = $teamsBotEntraIdAppDebug.AppId
            logAnalyticsWorkspaceId = $LogAnalyticsWorkspaceId
            Verbose                 = ($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
        }
        # creates the resources for the bot only!
        # endpoint is just an example, you will need to change it all the time
        New-AzResourceGroupDeployment @debugDeploymentConfig

     
        # for debug purposes, give the same creds as the workload
        Grant-GraphPermissions -ServicePrincipalDisplayName $teamsBotNameDebug -Permissions $botPermissionsNeeded
    }
    # $deploymentConfig = @{
    #     Mode                    = 'Incremental'
    #     Name                    = $deploymentName
    #     ResourceGroupName       = $ResourceGroupName
    #     TemplateFile            = $botTemplate
    #     endpoint                = "https://api.$Environment.uniphar.ie/platform-teams-notification-api/api/messages"
    #     environment             = $Environment
    #     botName                 = "platform-teams-notification-api-$Environment-bot"
    #     teamsBotAppId           = $envSA.AppId
    #     logAnalyticsWorkspaceId = $LogAnalyticsWorkspaceId
    #     Verbose                 = ($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)
    # }
    # # deploy the bot service, using the workload identity of the k8s cluster
    # New-AzResourceGroupDeployment @deploymentConfig

    # # platform-teams-notifications-api needs permissions to graph stuff, since we use workload identity we can use this
    # # in the future add revoke existing if needed and use a custom workload identity for the bot
    # Grant-GraphPermissions -ServicePrincipalDisplayName $ServicePrincipalName -Permissions $botPermissionsNeeded
}

function Grant-GraphPermissions {
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        [string] $ServicePrincipalDisplayName,
        
        [parameter(Mandatory = $true)]
        [string[]] $Permissions
    )
    
    # Get the service principal
    $sp = Get-AzADServicePrincipal -DisplayName $ServicePrincipalDisplayName
    if (!$sp) {
        throw "Service principal '$ServicePrincipalDisplayName' not found"
    }
    
    # Get Microsoft Graph service principal
    $graphSp = Get-AzADServicePrincipal -Filter "displayName eq 'Microsoft Graph'"
    
    # Get the app roles from Microsoft Graph
    $appRoles = $graphSp.AppRole | Where-Object { $Permissions -contains $_.Value }
    
    foreach ($appRole in $appRoles) {
        # Check if permission already exists
        $existingAssignment = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -ErrorAction SilentlyContinue | 
        Where-Object { $_.AppRoleId -eq $appRole.Id -and $_.ResourceId -eq $graphSp.Id }
        
        if (!$existingAssignment) {
            Write-Verbose "Granting permission: $($appRole.Value)"
            New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id `
                -PrincipalId $sp.Id `
                -ResourceId $graphSp.Id `
                -AppRoleId $appRole.Id
        }
        else {
            Write-Verbose "Permission already exists: $($appRole.Value)"
        }
    }
}