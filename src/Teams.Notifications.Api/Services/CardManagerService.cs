using AdaptiveCard = AdaptiveCards.AdaptiveCard;

namespace Teams.Notifications.Api.Services;

public sealed class CardManagerService(IChannelAdapter adapter, ITeamsManagerService teamsManagerService, IConfiguration config, ILogger<CardManagerService> logger, ICustomEventTelemetryClient telemetry) : ICardManagerService
{
    private readonly string _clientId = config["AZURE_CLIENT_ID"] ?? throw new ArgumentNullException(nameof(config), "Missing AZURE_CLIENT_ID");
    private readonly string _tenantId = config["AZURE_TENANT_ID"] ?? throw new ArgumentNullException(nameof(config), "Missing AZURE_TENANT_ID");

    public async Task DeleteCardAsync(string jsonFileName, string uniqueId, string teamName, string channelName, CancellationToken token)
    {
        var stopwatch = Stopwatch.StartNew();
        var teamId = await teamsManagerService.GetTeamIdAsync(teamName, token);
        await teamsManagerService.CheckOrInstallBotIsInTeam(teamId, token);
        var channelId = await teamsManagerService.GetChannelIdAsync(teamId, channelName, token);
        var conversationReference = GetConversationReference(channelId);
        var id = await teamsManagerService.GetMessageIdByUniqueId(teamId, channelId, jsonFileName, uniqueId, token);
        // check that we found the item to delete
        if (string.IsNullOrWhiteSpace(id)) throw new InvalidOperationException($"Card with unique ID '{uniqueId}' not found in team '{teamName}', channel '{channelName}'");
        conversationReference.ActivityId = id;
        // delete the item
        await adapter.ContinueConversationAsync(AgentClaims.CreateIdentity(_clientId),
            conversationReference,
            async (turnContext, cancellationToken) =>
            {
                await adapter.DeleteActivityAsync(turnContext, conversationReference, cancellationToken);
                telemetry.TrackEvent("ChannelDeleteMessage",
                    new()
                    {
                        ["Team"] = teamName,
                        ["Channel"] = channelName,
                        ["Id"] = id,
                        ["Duration"] = stopwatch.ElapsedMilliseconds
                    });
            },
            token);
    }

    public async Task<string?> GetCardAsync(string jsonFileName, string uniqueId, string teamName, string channelName, CancellationToken token)
    {
        var stopwatch = Stopwatch.StartNew();
        var teamId = await teamsManagerService.GetTeamIdAsync(teamName, token);
        await teamsManagerService.CheckOrInstallBotIsInTeam(teamId, token);
        var channelId = await teamsManagerService.GetChannelIdAsync(teamId, channelName, token);
        var chatMessage = await teamsManagerService.GetMessageByUniqueId(teamId, channelId, jsonFileName, uniqueId, token);
        telemetry.TrackEvent("ChannelGetCard",
            new()
            {
                ["Team"] = teamName,
                ["Channel"] = channelName,
                ["JsonFileName"] = jsonFileName,
                ["UniqueId"] = uniqueId,
                ["Duration"] = stopwatch.ElapsedMilliseconds
            });
        // check that we found the item to delete
        return chatMessage?.GetAdaptiveCardFromChatMessage();
    }

    public async Task CreateMessageToUserAsync(string message, string user, CancellationToken token)
    {
        var stopwatch = Stopwatch.StartNew();
        var userAadObjectId = await teamsManagerService.GetUserAadObjectIdAsync(user, token);
        var installedAppId = await teamsManagerService.GetOrInstallChatAppIdAsync(userAadObjectId, token);
        if (string.IsNullOrWhiteSpace(installedAppId)) throw new InvalidOperationException($"Unable to install or retrieve chat app for user '{user}'");
        var chatId = await teamsManagerService.GetChatIdAsync(installedAppId, userAadObjectId, token);
        if (string.IsNullOrWhiteSpace(chatId)) throw new InvalidOperationException($"Unable to retrieve chat for user '{user}'");
        var conversationReference = GetConversationReference(chatId);
        await adapter.ContinueConversationAsync(AgentClaims.CreateIdentity(_clientId),
            conversationReference,
            async (turnContext, cancellationToken) =>
            {
                // item is new
                var newResult = await turnContext.SendActivityAsync(MessageFactory.Text(message), cancellationToken);
                telemetry.TrackEvent("ChatNewMessage",
                    new()
                    {
                        ["MessageId"] = newResult.Id,
                        ["Duration"] = stopwatch.ElapsedMilliseconds
                    });
            },
            token);
    }

    public async Task CreateOrUpdateAsync<T>(string jsonFileName, T model, string user, CancellationToken token) where T : BaseTemplateModel
    {
        var stopwatch = Stopwatch.StartNew();
        var userAadObjectId = await teamsManagerService.GetUserAadObjectIdAsync(user, token);

        var installedAppId = await teamsManagerService.GetOrInstallChatAppIdAsync(userAadObjectId, token);


        if (string.IsNullOrWhiteSpace(installedAppId)) throw new InvalidOperationException($"Unable to install or retrieve chat app for user '{user}'");

        var chatId = await teamsManagerService.GetChatIdAsync(installedAppId, userAadObjectId, token);

        if (string.IsNullOrWhiteSpace(chatId)) throw new InvalidOperationException($"Unable to retrieve chat for user '{user}'");
        var chatMessage = await teamsManagerService.GetChatMessageByUniqueId(chatId, userAadObjectId, jsonFileName, model.UniqueId, token);

        var activity = new Activity
        {
            Type = "message",
            Attachments = new List<Attachment>
            {
                new()
                {
                    ContentType = AdaptiveCard.ContentType,
                    Content = await CreateCardFromTemplateAsync(jsonFileName, null, model, teamsManagerService, token: token)
                }
            }
        };

        var conversationReference = GetConversationReference(chatId);
        var idFromOldMessage = chatMessage?.Id;
        // found an existing card so update id
        if (!string.IsNullOrWhiteSpace(idFromOldMessage))
        {
            activity.Id = idFromOldMessage;
            conversationReference.ActivityId = idFromOldMessage;
        }


        await adapter.ContinueConversationAsync(AgentClaims.CreateIdentity(_clientId),
            conversationReference,
            async (turnContext, cancellationToken) =>
            {
                if (string.IsNullOrWhiteSpace(idFromOldMessage))
                {
                    // item is new
                    var newResult = await turnContext.SendActivityAsync(activity, cancellationToken);
                    telemetry.TrackEvent("ChatNewMessage",
                        new()
                        {
                            ["MessageId"] = newResult.Id,
                            ["Duration"] = stopwatch.ElapsedMilliseconds
                        });
                    return;
                }

                // item needs update
                var updateResult = await turnContext.UpdateActivityAsync(activity, cancellationToken);
                telemetry.TrackEvent("ChatUpdateMessage",
                    new()
                    {
                        ["MessageId"] = updateResult.Id,
                        ["Duration"] = stopwatch.ElapsedMilliseconds
                    });
            },
            token);
    }


    public async Task CreateOrUpdateAsync<T>(string jsonFileName, IFormFile? file, T model, string teamName, string channelName, CancellationToken token) where T : BaseTemplateModel
    {
        var stopwatch = Stopwatch.StartNew();
        var teamId = await teamsManagerService.GetTeamIdAsync(teamName, token);
        await teamsManagerService.CheckOrInstallBotIsInTeam(teamId, token);
        var channelId = await teamsManagerService.GetChannelIdAsync(teamId, channelName, token);

        var activity = new Activity
        {
            Type = "message",
            Attachments = new List<Attachment>
            {
                new()
                {
                    ContentType = AdaptiveCard.ContentType,
                    Content = await CreateCardFromTemplateAsync(jsonFileName, file, model, teamsManagerService, teamId, channelId, channelName, token)
                }
            }
        };
        var conversationReference = GetConversationReference(channelId);
        var idFromOldMessage = await teamsManagerService.GetMessageIdByUniqueId(teamId, channelId, jsonFileName, model.UniqueId, token);
        // found an existing card so update id
        if (!string.IsNullOrWhiteSpace(idFromOldMessage))
        {
            activity.Id = idFromOldMessage;
            conversationReference.ActivityId = idFromOldMessage;
        }

        await adapter.ContinueConversationAsync(AgentClaims.CreateIdentity(_clientId),
            conversationReference,
            async (turnContext, cancellationToken) =>
            {
                if (string.IsNullOrWhiteSpace(idFromOldMessage))
                {
                    // item is new
                    var newResult = await turnContext.SendActivityAsync(activity, cancellationToken);
                    telemetry.TrackEvent("ChannelNewMessage",
                        new()
                        {
                            ["Team"] = teamName,
                            ["Channel"] = channelName,
                            ["MessageId"] = newResult.Id,
                            ["Duration"] = stopwatch.ElapsedMilliseconds
                        });
                    return;
                }

                // item needs update
                var updateResult = await turnContext.UpdateActivityAsync(activity, cancellationToken);
                telemetry.TrackEvent("ChannelUpdateMessage",
                    new()
                    {
                        ["Team"] = teamName,
                        ["Channel"] = channelName,
                        ["MessageId"] = updateResult.Id,
                        ["Duration"] = stopwatch.ElapsedMilliseconds
                    });
            },
            token);
    }

    public async Task RemoveActionsFromCardAsync(string teamId, string channelId, string messageId, string[] actionsToRemove, CancellationToken token)
    {
        var stopwatch = Stopwatch.StartNew();
        // to get the card, this doesn't get provided by the action invoke
        var chatMessage = await teamsManagerService.GetMessageById(teamId, channelId, messageId, token);
        var cardJson = chatMessage?.GetAdaptiveCardFromChatMessage();
        if (string.IsNullOrWhiteSpace(cardJson))
        {
            telemetry.TrackEvent("NoAdaptiveCardFound",
                new()
                {
                    ["Team"] = teamId,
                    ["Channel"] = channelId,
                    ["MessageId"] = messageId
                });
            throw new InvalidOperationException("Card not found in team");
        }

        var card = AdaptiveCard.FromJson(cardJson).Card;
        // remove all actions that match the verbs
        foreach (var actionVerb in actionsToRemove)
        foreach (var adaptiveAction in card.Actions.Where(a => a is AdaptiveExecuteAction exe && exe.Verb == actionVerb).ToList())
            card.Actions.Remove(adaptiveAction);

        var activity = new Activity
        {
            Type = "message",
            Id = messageId,
            Attachments = new List<Attachment>
            {
                new()
                {
                    ContentType = AdaptiveCard.ContentType,
                    Content = card.ToJson()
                }
            }
        };

        var conversationReference = GetConversationReference(channelId);
        conversationReference.ActivityId = messageId;

        await adapter.ContinueConversationAsync(AgentClaims.CreateIdentity(_clientId),
            conversationReference,
            async (turnContext, cancellationToken) =>
            {
                var updateResult = await turnContext.UpdateActivityAsync(activity, cancellationToken);
                telemetry.TrackEvent("ChannelUpdateCardRemoveActions",
                    new()
                    {
                        ["Team"] = teamId,
                        ["Channel"] = channelId,
                        ["MessageId"] = updateResult.Id,
                        ["ActionsRemoved"] = string.Join(",", actionsToRemove),
                        ["Duration"] = stopwatch.ElapsedMilliseconds
                    });
            },
            token);
    }

    public async Task<string> CreateCardFromTemplateAsync<T>(string jsonFileName, IFormFile? formFile, T model, ITeamsManagerService teamsManagerService, string? teamId = null, string? channelId = null, string? channelName = null, CancellationToken token = default) where T : BaseTemplateModel
    {
        var text = await File.ReadAllTextAsync($"./Templates/{jsonFileName}", token);
        var props = text.GetMustachePropertiesFromString();
        var fileUrl = string.Empty;
        var fileLocation = string.Empty;
        var fileName = string.Empty;
        if (!string.IsNullOrEmpty(teamId) && !string.IsNullOrEmpty(channelId))
        {
            if (props.HasFileTemplate() && formFile != null)
            {
                fileName = formFile.FileName;
                fileLocation = channelName + "/error/" + formFile.FileName;
                try
                {
                    await using var stream = formFile.OpenReadStream();
                    await teamsManagerService.UploadFile(teamId, channelId, fileLocation, stream, token);
                    fileUrl = await teamsManagerService.GetFileUrl(teamId, channelId, fileLocation, token);
                }
                catch (Exception ex)
                {
                    logger.LogError(ex, "Error uploading file to Teams, continuing");
                }
            }
        }

        // replace all props with the values

        foreach (var (propertyName, type) in props)
        {
            text = text.FindPropAndReplace(model,
                new()
                {
                    Property = propertyName,
                    Type = type,
                    File = new()
                    {
                        Url = fileUrl,
                        Location = fileLocation,
                        Name = fileName
                    }
                });
        }

        var item = AdaptiveCard.FromJson(text).Card;
        if (item == null) throw new ArgumentNullException(nameof(jsonFileName));
        // some solution to be able to track a unique id across the channel
        item.Body.Add(new AdaptiveTextBlock(jsonFileName)
        {
            Color = AdaptiveTextColor.Accent,
            Size = AdaptiveTextSize.Small,
            Id = model.UniqueId,
            IsSubtle = true,
            IsVisible = false,
            Wrap = true
        });
        return item.ToJson();
    }

    private ConversationReference GetConversationReference(string channelId) =>
        new()
        {
            ChannelId = Channels.Msteams,
            ServiceUrl = $"https://smba.trafficmanager.net/emea/{_tenantId}",
            Conversation = new(id: channelId),
            ActivityId = channelId,
            Agent = new(agenticAppId: _clientId)
        };
}