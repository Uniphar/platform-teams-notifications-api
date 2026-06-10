# Platform Teams Notifications API

This is the first iteration of the Notifications API. In this iteration, the API will handle a single API call to:

1. Create an app.
2. Place it in the appropriate Teams channel.
3. Set up a card in the Teams channel with two buttons:
   - One for reprocessing (dummy functionality).
   - One to view a file.

4. Retrieve the file from a blob storage location.
5. Display the file in Teams when the button is clicked.

## Environment Setup

If you need to provision the Azure infrastructure manually (bot service, Graph API permissions, Teams app manifest, alert rules), see the [Environment Setup Guide](docs/environment-setup.md).

## local setup

Local setup will use the debug bot to communicate making it possible to locally debug.
We run the rest of the application in Azure AKS with WorkLoad identities, these are created with `azwi` and their clientId's are stored in a Service account, we use the service acount in the pod to authenticate, more information can be found here: `https://learn.microsoft.com/en-us/azure/aks/workload-identity-overview`

Since this uses federation we cannot use it locally, for this we will use a bot service setup with a client secret, we also have to have an api endpoint, to make it work locally:

1. Create a dev tunnel:

   ```bash
   devtunnel user login # only once every 24h or so
   devtunnel host -p 3978 --allow-anonymous

   ```

2. On the Azure Bot (for local/debug: devops-debug-bot), select **Settings**, then **Configuration**, and update the **Messaging endpoint** to `{tunnel-url}/api/messages` eg: `https://kw238403-3978.eun1.devtunnels.ms/platform-teams-notification-api/api/messages`
3. Change the secret of the appsettings.local, you can create this by hand, the client-id and tenant is already setup but might need to be changed if this is a new application
4. Run the application
5. Add the bot to teams, select **Settings**, then **Channels**, and click on the link **Open in Teams** or let the bot install it for you

## Initial create of the app

devops-azure will create the bot services automatically, but to be able to use the app you have to go to:
`https://dev.teams.microsoft.com/apps` and create a new app, use the templates from `src\Teams.Notifications.Api\appManifest\ENV`
You will have to change the following:

1. Manifest.id (take that from the manifest that is created when you create the app)
2. The bots.botId/webApplicationInfo.id and resource, leave the `api://botid-` part for the resource, just do a guid replace, the ID should be the AppId of the app, this is equal to the `Microsoft App ID` in the `Configuration` of your bot that is under `Bot services`
   1. The required Graph API permissions must be granted manually — see the [Environment Setup Guide](docs/environment-setup.md#2-microsoft-graph-api-permissions) for the full list and the PowerShell snippet to assign them.
3. You might want to change the names, logo's etc, we use Uniphar since that is what we use it for

The pending apps you can find in `https://admin.teams.microsoft.com/policies/manage-apps` with pim you can approve these, to view, choose app type= custom app or search for the name, it will take a few hours to propagate, if you open teams in the browser it tends to be there quicker.

### Where to find stuff

`https://api.dev.uniphar.ie/platform-teams-notification-api/swagger` for the swagger page (change dev to the right env)

### Easy formatting

If you move the pre-commit file from the root, to `.git/hooks/pre-commit` it will automatically format your Adaptive card templates, otherwise you will either have to run the task `Run formatter` in vscode or the build will fail!
