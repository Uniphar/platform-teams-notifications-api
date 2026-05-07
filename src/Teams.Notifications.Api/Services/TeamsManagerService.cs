namespace Teams.Notifications.Api.Services;

public class TeamsManagerService(GraphServiceClient graphClient, IConfiguration config)
{
    private readonly string _clientId = config["AZURE_CLIENT_ID"] ?? throw new ArgumentNullException(nameof(config), "Missing AZURE_CLIENT_ID");


    public async Task<string> GetTeamsAppIdAsync(CancellationToken token)
    {
        var apps = await graphClient
            .AppCatalogs
            .TeamsApps
            .GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Filter = $"appDefinitions/any(a:a/authorization/clientAppId eq '{_clientId}')";
                    requestConfiguration.QueryParameters.Expand = ["appDefinitions"];
                },
                token);

        var teamsApp = apps?.Value?.FirstOrDefault();
        return teamsApp?.Id ?? throw new InvalidOperationException($"Teams app with client ID {_clientId} not found in app catalog");
    }

    public async Task CheckOrInstallBotIsInTeam(string teamId, CancellationToken token)
    {
        var foundAppInstallsResponse = await graphClient
            .Teams[teamId]
            .InstalledApps
            .GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Expand = ["teamsAppDefinition"];
                    requestConfiguration.QueryParameters.Filter = $"teamsAppDefinition/authorization/clientAppId eq '{_clientId}'";
                },
                token);
        if (foundAppInstallsResponse?.Value?.Count == 0)
        {
            var teamsAppId = await GetTeamsAppIdAsync(token);
            var requestBody = new UserScopeTeamsAppInstallation
            {
                ConsentedPermissionSet = new()
                {
                    // same as the actual manifest permissions
                    ResourceSpecificPermissions = new()
                    {
                        new()
                        {
                            PermissionValue = "Channel.Create.Group",
                            PermissionType = TeamsAppResourceSpecificPermissionType.Application
                        },
                        new()
                        {
                            PermissionValue = "Channel.Delete.Group",
                            PermissionType = TeamsAppResourceSpecificPermissionType.Application
                        },
                        new()
                        {
                            PermissionValue = "ChannelMessage.Read.Group",
                            PermissionType = TeamsAppResourceSpecificPermissionType.Application
                        },
                        new()
                        {
                            PermissionValue = "ChannelMessage.Send.Group",
                            PermissionType = TeamsAppResourceSpecificPermissionType.Application
                        },
                        new()
                        {
                            PermissionValue = "ChannelSettings.Read.Group",
                            PermissionType = TeamsAppResourceSpecificPermissionType.Application
                        },
                        new()
                        {
                            PermissionValue = "ChannelSettings.ReadWrite.Group",
                            PermissionType = TeamsAppResourceSpecificPermissionType.Application
                        },
                        new()
                        {
                            PermissionValue = "Member.Read.Group",
                            PermissionType = TeamsAppResourceSpecificPermissionType.Application
                        },
                        new()
                        {
                            PermissionValue = "Owner.Read.Group",
                            PermissionType = TeamsAppResourceSpecificPermissionType.Application
                        },
                        new()
                        {
                            PermissionValue = "TeamMember.Read.Group",
                            PermissionType = TeamsAppResourceSpecificPermissionType.Application
                        },
                        new()
                        {
                            PermissionValue = "TeamsActivity.Send.Group",
                            PermissionType = TeamsAppResourceSpecificPermissionType.Application
                        }
                    }
                },
                AdditionalData = new Dictionary<string, object>
                {
                    ["teamsApp@odata.bind"] = $"https://graph.microsoft.com/beta/appCatalogs/teamsApps/{teamsAppId}"
                }
            };


            await graphClient
                .Teams[teamId]
                .InstalledApps
                .PostAsync(requestBody, cancellationToken: token);
        }
    }

    public async Task<string> GetTeamIdAsync(string teamName, CancellationToken token)
    {
        var groups = await graphClient.Teams.GetAsync(request =>
            {
                request.QueryParameters.Filter = $"displayName eq '{teamName}'";
                request.QueryParameters.Select = ["id"];
            },
            token);

        if (groups is not { Value: [{ Id: var teamId }] }) throw new InvalidOperationException($"Team with name {teamName} does not exist");
        return teamId ?? throw new InvalidOperationException($"Team with name {teamName} does not exist");
    }

    public async Task<string> GetChannelIdAsync(string teamId, string channelName, CancellationToken token)
    {
        var channels = await graphClient
            .Teams[teamId]
            .Channels
            .GetAsync(request =>
                {
                    request.QueryParameters.Filter = $"displayName eq '{channelName}'";
                    request.QueryParameters.Select = ["id"];
                },
                token);

        if (channels is not { Value: [{ Id: var channelId }] }) throw new InvalidOperationException($"Channel with name {channelName} does not exist");
        return channelId ?? throw new InvalidOperationException($"Channel with name {channelName} does not exist");
    }

    public async Task<string> GetGroupNameUniqueName(string groupId, CancellationToken token)
    {
        // teamId and groupId is the same, but if you look up group from a team it won't work!
        var group = await graphClient
            .Groups[groupId]
            .GetAsync(cancellationToken: token);
        return group?.UniqueName ?? throw new InvalidOperationException($"No group found for team {groupId}");
    }

    public async Task<string> GetTeamName(string teamId, CancellationToken token)
    {
        var team = await graphClient
            .Teams[teamId]
            .GetAsync(cancellationToken: token);
        return team?.DisplayName ?? throw new InvalidOperationException($"No DisplayName found for team {teamId}");
    }

    public async Task<string> GetUserAadObjectIdAsync(string userPrincipalName, CancellationToken token)
    {
        var user = await graphClient
            .Users[userPrincipalName]
            .GetAsync(request => { request.QueryParameters.Select = ["id"]; }, token);

        return user?.Id ?? throw new InvalidOperationException($"User with principal name {userPrincipalName} not found");
    }

    public async Task<string?> GetOrInstallChatAppIdAsync(string aadObjectId, CancellationToken token)
    {
        // check if app is installed for the user
        var installedChatResource = await graphClient
            .Users[aadObjectId]
            .Teamwork
            .InstalledApps
            .GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Expand = ["teamsAppDefinition"];
                    requestConfiguration.QueryParameters.Filter = $"teamsAppDefinition/authorization/clientAppId eq '{_clientId}'";
                },
                token);
        var id = installedChatResource?.Value?.FirstOrDefault()?.Id;
        if (!string.IsNullOrWhiteSpace(id)) return id;

        var teamsAppId = await GetTeamsAppIdAsync(token);
        var requestBody = new UserScopeTeamsAppInstallation
        {
            ConsentedPermissionSet = new()
            {
                ResourceSpecificPermissions = new()
                {
                    new()
                    {
                        PermissionValue = "TeamsActivity.Send.User",
                        PermissionType = TeamsAppResourceSpecificPermissionType.Application
                    }
                }
            },
            AdditionalData = new Dictionary<string, object>
            {
                ["teamsApp@odata.bind"] = $"https://graph.microsoft.com/beta/appCatalogs/teamsApps/{teamsAppId}"
            }
        };

        // create
        await graphClient.Users[aadObjectId].Teamwork.InstalledApps.PostAsync(requestBody, cancellationToken: token);
        // get the resource that we created
        installedChatResource = await graphClient
            .Users[aadObjectId]
            .Teamwork
            .InstalledApps
            .GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Expand = ["teamsAppDefinition"];
                    requestConfiguration.QueryParameters.Filter = $"teamsAppDefinition/authorization/clientAppId eq '{_clientId}'";
                },
                token);
        return installedChatResource?.Value?.FirstOrDefault()?.Id;
    }

    public async Task<string?> GetChatIdAsync(string installedAppId, string aadObjectId, CancellationToken token)
    {
        if (string.IsNullOrWhiteSpace(installedAppId)) return null;
        var chat = await graphClient.Users[aadObjectId].Teamwork.InstalledApps[installedAppId].Chat.GetAsync(cancellationToken: token);
        return chat?.Id;
    }

    public async Task<ChatMessage?> GetChatMessageByUniqueId(string chatId, string userAadObjectId, string jsonFileName, string uniqueId, CancellationToken token)
    {
        var messagesResponse = await graphClient
            .Chats[chatId]
            .Messages
            .GetAsync(
                requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Top = 50; // default is 20, but chats can be very active so we increase it a bit
                },
                token);
        // no need to do anything if there is no message
        var responses = messagesResponse?.Value;
        if (responses == null) return null;
        var foundMessage = responses.Select(s => s.GetCardThatHas(jsonFileName, uniqueId)).FirstOrDefault(x => x != null);
        if (foundMessage != null) return foundMessage;
        var chatPageCount = 0;
        // go max 4 times top* 5 will be the number of messages
        while (foundMessage == null && messagesResponse?.OdataNextLink != null && chatPageCount < 4)
        {
            chatPageCount++;
            var configuration = new RequestInformation
            {
                HttpMethod = Method.GET,
                URI = new(messagesResponse.OdataNextLink)
            };

            messagesResponse = await graphClient.RequestAdapter.SendAsync(configuration, _ => new ChatMessageCollectionResponse(), cancellationToken: token);
            if (messagesResponse?.Value == null) throw new NullReferenceException("Messages should not be null if there is a next page");
            foundMessage = messagesResponse.Value.Select(s => s.GetCardThatHas(jsonFileName, uniqueId)).FirstOrDefault(x => x != null);
        }

        return foundMessage;
    }

    public async Task<string?> GetMessageIdByUniqueId(string teamId, string channelId, string jsonFileName, string uniqueId, CancellationToken token) => (await GetMessageByUniqueId(teamId, channelId, jsonFileName, uniqueId, token))?.Id;

    public async Task<ChatMessage?> GetMessageByUniqueId(string teamId, string channelId, string jsonFileName, string uniqueId, CancellationToken token)
    {
        // can't filter, so we will just go 400 messages back MAX
        var messagesResponse = await graphClient
            .Teams[teamId]
            .Channels[channelId]
            .Messages
            .GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Top = 50; // default is 20, but chats can be very active so we increase it a bit
                },
                token);
        var responses = messagesResponse
            ?.Value
            ?.Where(x => x.DeletedDateTime == null &&
                         x.From?.Application != null &&
                         x.From.Application.Id == _clientId
            )
            .ToList();
        // no need to do anything if there is no message
        if (responses == null) return null;
        var foundMessage = responses.Select(s => s.GetCardThatHas(jsonFileName, uniqueId)).FirstOrDefault(x => x != null);
        if (foundMessage != null) return foundMessage;
        var channelPageCount = 0;
        // go max 4 times top* 5 will be the number of messages 
        while (messagesResponse?.OdataNextLink != null && channelPageCount < 4)
        {
            channelPageCount++;
            var configuration = new RequestInformation
            {
                HttpMethod = Method.GET,
                URI = new(messagesResponse.OdataNextLink)
            };

            messagesResponse = await graphClient.RequestAdapter.SendAsync(configuration, _ => new ChatMessageCollectionResponse(), cancellationToken: token);
            if (messagesResponse?.Value == null) throw new NullReferenceException("Messages should not be null if there is a next page");
            foundMessage = messagesResponse.Value.Select(s => s.GetCardThatHas(jsonFileName, uniqueId)).FirstOrDefault(x => x != null);
        }

        return foundMessage;
    }

    public async Task<ChatMessage?> GetMessageById(string teamId, string channelId, string messageId, CancellationToken token)
    {
        try
        {
            var chatMessage = await graphClient
                .Teams[teamId]
                .Channels[channelId]
                .Messages[messageId]
                .GetAsync(cancellationToken: token);

            return chatMessage;
        }
        catch (Exception)
        {
            return null;
        }
    }

    public async Task<(bool Succes, string Url)> UploadFile(string teamId, string channelId, string fileLocation, Stream fileStream, CancellationToken token)
    {
        var filesFolder = await graphClient.Teams[teamId].Channels[channelId].FilesFolder.GetAsync(cancellationToken: token);
        if (filesFolder == null) throw new InvalidOperationException("No files folder found for the channel");

        var driveId = filesFolder.ParentReference?.DriveId;
        if (driveId == null) throw new InvalidOperationException("No drive found for the channel");

        var item = graphClient.Drives[driveId].Items["root"];
        // same as the list, we need to make sure you don't just drop it in the sharepoint site folder
        var content = item.ItemWithPath(fileLocation).Content;
        await content.PutAsync(fileStream, cancellationToken: token);
        var itemFound = await item.ItemWithPath(fileLocation).GetAsync(cancellationToken: token);
        // add web=1 to open in web view, this will make it possible to edit it in browser
        return itemFound is { WebUrl: not null } ? (true, itemFound.WebUrl + "?web=1") : (false, string.Empty);
    }

    public async Task<(bool Succes, string Url)> GetFileUrl(string teamId, string channelId, string fileLocation, CancellationToken token)
    {
        var filesFolder = await graphClient.Teams[teamId].Channels[channelId].FilesFolder.GetAsync(cancellationToken: token);
        var driveId = filesFolder?.ParentReference?.DriveId;
        if (driveId == null) throw new InvalidOperationException("No drive found for the channel");
        var item = await GetDriveItem(driveId, fileLocation, token);
        // add web=1 to open in web view, this will make it possible to edit it in browser
        return item is { WebUrl: not null } ? (true, item.WebUrl + "?web=1") : (false, string.Empty);
    }

    public async Task<string> GetFileNameAsync(string teamId, string channelId, string fileLocation, CancellationToken token)
    {
        var filesFolder = await graphClient.Teams[teamId].Channels[channelId].FilesFolder.GetAsync(cancellationToken: token);
        var driveId = filesFolder?.ParentReference?.DriveId;
        if (driveId == null) throw new InvalidOperationException("No drive found for the channel");
        var item = await GetDriveItem(driveId, fileLocation, token);
        return item?.Name ?? throw new InvalidOperationException("Name not found");
    }

    private async Task<DriveItem?> GetDriveItem(string driveId, string fileUrl, CancellationToken cancellationToken = default)
    {
        var path = Path.GetDirectoryName(fileUrl);
        var rootRequest = graphClient.Drives[driveId].Root;
        var children = rootRequest.ItemWithPath(path).Children;
        var driveItems = (await children.GetAsync(cancellationToken: cancellationToken))?.Value;
        return driveItems?.FirstOrDefault(x => x.Name == Path.GetFileName(fileUrl));
    }
}