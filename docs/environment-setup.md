# Environment Setup Guide

This guide describes the manual steps to provision the infrastructure for the Platform Teams Notification API. In normal circumstances the infra team handles this. This document is intended for situations where that is not the case, or for contributors who need to understand how the environment is structured.

> For local development setup (dev tunnels, debug bot, appsettings, etc.) see the [README](../README.md#local-setup).

---

## Prerequisites

- [Azure CLI](https://learn.microsoft.com/en-us/cli/azure/install-azure-cli) installed and authenticated (`az login`)
- Contributor access on the target resource group
- Privileged Role Administrator or Global Administrator in Entra ID (for granting Graph API admin consent)
- An existing AKS cluster with a workload identity service principal already configured

---

## 1. Bot Service

The API communicates with Teams via an Azure Bot Service. The bot's identity is the **workload identity service principal** that the AKS pod already uses — no separate app registration is needed for non-debug environments.

### 1a. Deploy the Bot Service

The `devops/bot.bicep` template provisions the bot service, the Teams channel, diagnostic settings, and alert rules in a single deployment. Run the deployment with the workload identity's App ID:

```bash
ENVIRONMENT=dev          # dev | test | prod
RESOURCE_GROUP=<resource-group-name>
SUBSCRIPTION_ID=<subscription-id>
LOG_ANALYTICS_ID=<log-analytics-workspace-resource-id>

# Retrieve the App ID of the workload identity service principal
BOT_APP_ID=$(az ad sp list \
  --display-name "<workload-identity-service-principal-name>" \
  --query "[0].appId" -o tsv)

az account set --subscription "$SUBSCRIPTION_ID"

az deployment group create \
  --resource-group "$RESOURCE_GROUP" \
  --name "deploy-$(date +%Y%m%d%H%M%S)-teams-bot" \
  --template-file devops/bot.bicep \
  --parameters \
    botName="platform-teams-notification-api-${ENVIRONMENT}-bot" \
    teamsBotAppId="$BOT_APP_ID" \
    environment="$ENVIRONMENT" \
    endpoint="https://api.${ENVIRONMENT}.uniphar.ie/platform-teams-notification-api/api/messages" \
    logAnalyticsWorkspaceId="$LOG_ANALYTICS_ID"
```

> **Note:** The `bot.bicep` template assumes action groups named `platform-engineering-applications-low` and `platform-engineering-applications-high` exist in a resource group called `observability` within the same subscription. These must be in place before deploying.

---

## 2. Microsoft Graph API Permissions

The workload identity service principal needs the following **application** (daemon/non-interactive) permissions on Microsoft Graph. These require admin consent and cannot be delegated.

| Permission                                                | Purpose                                                           |
| --------------------------------------------------------- | ----------------------------------------------------------------- |
| `Group.ReadWrite.All`                                     | Create groups, update membership, manage private/shared channels  |
| `Team.Create`                                             | Create new Microsoft Teams                                        |
| `TeamSettings.ReadWrite.All`                              | Read and update team settings                                     |
| `TeamMember.ReadWrite.All`                                | Manage team membership (supplements `Group.ReadWrite.All`)        |
| `ChannelMember.ReadWrite.All`                             | Add/remove users from channels, including private/shared channels |
| `ChannelSettings.ReadWrite.All`                           | Read and update channel settings                                  |
| `ChannelMessage.Read.All`                                 | Read all channel messages across the tenant                       |
| `ChannelMessage.Send`                                     | Send messages as the bot into any channel                         |
| `Chat.ReadWrite.All`                                      | Read/write all chat messages (1:1 and group) and send messages    |
| `ChatMember.ReadWrite.All`                                | Add/remove members from chats                                     |
| `AppCatalog.ReadWrite.All`                                | Read app IDs from the Teams app catalog                           |
| `TeamsAppInstallation.Read.All`                           | Read all installed Teams apps in the tenant                       |
| `TeamsAppInstallation.ReadWriteAndConsentSelfForUser.All` | Install/uninstall the app for a user                              |
| `TeamsAppInstallation.ReadWriteAndConsentSelfForChat.All` | Install/uninstall the app for a chat                              |
| `TeamsAppInstallation.ReadWriteAndConsentSelfForTeam.All` | Install/uninstall the app for a team                              |
| `Files.ReadWrite.All`                                     | Read/write files in Teams channels (SharePoint-backed)            |
| `User.Read.All`                                           | Resolve user IDs when adding members to teams/chats/channels      |

### Grant the Permissions

The following PowerShell snippet uses the `Az` and `Microsoft.Graph` modules (which are already present in the environment) to assign the roles. Run it once per environment after the service principal is in place.

```powershell
# Connect — uses the same Az session already established
$token = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com/" -AsSecureString
Connect-MgGraph -AccessToken $token.Token -NoWelcome

$servicePrincipalName = "<workload-identity-service-principal-name>"

$permissions = @(
    "Group.ReadWrite.All"
    "Team.Create"
    "TeamSettings.ReadWrite.All"
    "TeamMember.ReadWrite.All"
    "ChannelMember.ReadWrite.All"
    "ChannelSettings.ReadWrite.All"
    "ChannelMessage.Read.All"
    "ChannelMessage.Send"
    "Chat.ReadWrite.All"
    "ChatMember.ReadWrite.All"
    "AppCatalog.ReadWrite.All"
    "TeamsAppInstallation.Read.All"
    "TeamsAppInstallation.ReadWriteAndConsentSelfForUser.All"
    "TeamsAppInstallation.ReadWriteAndConsentSelfForChat.All"
    "TeamsAppInstallation.ReadWriteAndConsentSelfForTeam.All"
    "Files.ReadWrite.All"
    "User.Read.All"
)

$sp      = Get-AzADServicePrincipal -DisplayName $servicePrincipalName
$graphSp = Get-AzADServicePrincipal -Filter "displayName eq 'Microsoft Graph'"
$roles   = $graphSp.AppRole | Where-Object { $permissions -contains $_.Value }

foreach ($role in $roles) {
    $exists = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id |
              Where-Object { $_.AppRoleId -eq $role.Id -and $_.ResourceId -eq $graphSp.Id }
    if (-not $exists) {
        Write-Host "Granting: $($role.Value)"
        New-MgServicePrincipalAppRoleAssignment `
            -ServicePrincipalId $sp.Id `
            -PrincipalId $sp.Id `
            -ResourceId $graphSp.Id `
            -AppRoleId $role.Id
    }
}
```

---

## 3. Debug Bot (Local Development Only)

For local development a separate app registration is used because the workload identity uses federation and cannot provide a client secret. See the [Local Setup section in the README](../README.md#local-setup) for the full walkthrough.

The steps at a high level are:

1. Create a dedicated app registration for the debug bot:

   ```bash
   az ad app create \
     --display-name "platform-teams-notification-api-debug-bot" \
     --sign-in-audience AzureADMyOrg
   ```

2. Create a service principal for it:

   ```bash
   DEBUG_BOT_APP_ID=$(az ad app list \
     --display-name "platform-teams-notification-api-debug-bot" \
     --query "[0].appId" -o tsv)

   az ad sp create --id "$DEBUG_BOT_APP_ID"
   ```

3. Deploy the bot service pointing at your dev tunnel endpoint (see the README for tunnel setup):

   ```bash
   az deployment group create \
     --resource-group "$RESOURCE_GROUP" \
     --name "deploy-$(date +%Y%m%d%H%M%S)-teams-debug-bot" \
     --template-file devops/bot.bicep \
     --parameters \
       botName="platform-teams-notification-api-debug-bot" \
       teamsBotAppId="$DEBUG_BOT_APP_ID" \
       environment="debug" \
       endpoint="https://<your-tunnel>.devtunnels.ms/platform-teams-notification-api/api/messages" \
       logAnalyticsWorkspaceId="$LOG_ANALYTICS_ID"
   ```

4. Grant the debug bot the same Graph permissions as the workload identity (run the same PowerShell snippet from [step 2](#2-microsoft-graph-api-permissions), substituting `platform-teams-notification-api-debug-bot` as the service principal name).

5. Create a client secret for local use and add it to `appsettings.local.json` (see the [README](../README.md#local-setup)).

---

## 4. Teams App Manifest

Once the bot service is running you need to register the Teams application. See the [Initial create of the app section in the README](../README.md#initial-create-of-the-app) for full instructions.

In brief:

1. Go to [https://dev.teams.microsoft.com/apps](https://dev.teams.microsoft.com/apps) and create a new app from the manifest template in `src/Teams.Notifications.Api/appManifest/<ENV>`.
2. Replace the `manifest.id`, `bots.botId`, and `webApplicationInfo.id` with the App ID of the bot service you deployed in step 1 (this is the **Microsoft App ID** shown under **Bot Services → Configuration**).
3. Submit the app for approval.
4. Approve the pending app at [https://admin.teams.microsoft.com/policies/manage-apps](https://admin.teams.microsoft.com/policies/manage-apps) (requires Teams admin or PIM elevation). Filter by **Custom app** or search by name. Propagation can take a few hours; using Teams in the browser tends to be faster.

---

## 5. Alert Rules Only

If the bot service is already deployed and you only need to update/redeploy the alert rules, you can use the `alerts.bicep` template directly (or run `Initialize-PlatformTeamsNotificationApi.ps1`):

```bash
az deployment group create \
  --resource-group "$RESOURCE_GROUP" \
  --name "deploy-$(date +%Y%m%d%H%M%S)-teams-alerts" \
  --template-file devops/alerts.bicep \
  --parameters \
    environment="$ENVIRONMENT" \
    logAnalyticsWorkspaceId="$LOG_ANALYTICS_ID"
```
