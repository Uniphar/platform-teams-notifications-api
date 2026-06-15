using AdaptiveCard = AdaptiveCards.AdaptiveCard;

namespace Teams.Notifications.Api.Services;

public sealed class CardManagerService(IChannelAdapter adapter, ITeamsManagerService teamsManagerService, ICosmosMessageStore cosmosMessageStore, IConfiguration config, ILogger<CardManagerService> logger, ICustomEventTelemetryClient telemetry) : ICardManagerService
{
    private readonly string _clientId = config["AZURE_CLIENT_ID"] ?? throw new ArgumentNullException(nameof(config), "Missing AZURE_CLIENT_ID");
    private readonly string _tenantId = config["AZURE_TENANT_ID"] ?? throw new ArgumentNullException(nameof(config), "Missing AZURE_TENANT_ID");

    public async Task DeleteCardAsync(string jsonFileName, string uniqueId, string teamName, string channelName, CancellationToken token)
    {
        try
        {
            var stopwatch = Stopwatch.StartNew();
            var teamId = await teamsManagerService.GetTeamIdAsync(teamName, token);
            await teamsManagerService.CheckOrInstallBotIsInTeam(teamId, token);
            var channelId = await teamsManagerService.GetChannelIdAsync(teamId, channelName, token);
            var conversationReference = GetConversationReference(channelId);
            var stored = await cosmosMessageStore.FindByChannelAsync(jsonFileName, uniqueId, token);

            if (stored is null)
            {
                var errorMsg = $"Card with unique ID '{uniqueId}' not found in team '{teamName}', channel '{channelName}'";
                logger.LogWarning(errorMsg);
                throw new InvalidOperationException(errorMsg);
            }

            conversationReference.ActivityId = stored.Id;
            await adapter.ContinueConversationAsync(AgentClaims.CreateIdentity(_clientId),
                conversationReference,
                async (turnContext, cancellationToken) =>
                {
                    try
                    {
                        await adapter.DeleteActivityAsync(turnContext, conversationReference, cancellationToken);
                        await cosmosMessageStore.DeleteAsync(stored.Id, cancellationToken);

                        telemetry.TrackEvent("ChannelDeleteMessage",
                            new()
                            {
                                ["Team"] = teamName,
                                ["Channel"] = channelName,
                                ["Id"] = stored.Id,
                                ["UniqueId"] = uniqueId,
                                ["Duration"] = stopwatch.ElapsedMilliseconds
                            });
                    }
                    catch (Exception ex)
                    {
                        logger.LogError(ex, "Error deleting message '{MessageId}' from channel '{Channel}'", stored.Id, channelName);
                        throw;
                    }
                },
                token);
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Error deleting card with unique ID '{UniqueId}' from team '{Team}', channel '{Channel}'", uniqueId, teamName, channelName);
            throw;
        }
    }

    public async Task<string?> GetCardAsync(string jsonFileName, string uniqueId, string teamName, string channelName, CancellationToken token)
    {
        try
        {
            var stopwatch = Stopwatch.StartNew();
            var teamId = await teamsManagerService.GetTeamIdAsync(teamName, token);
            await teamsManagerService.CheckOrInstallBotIsInTeam(teamId, token);
            var channelId = await teamsManagerService.GetChannelIdAsync(teamId, channelName, token);
            var stored = await cosmosMessageStore.FindByChannelAsync(jsonFileName, uniqueId, token);

            telemetry.TrackEvent("ChannelGetCard",
                new()
                {
                    ["Team"] = teamName,
                    ["Channel"] = channelName,
                    ["JsonFileName"] = jsonFileName,
                    ["UniqueId"] = uniqueId,
                    ["Found"] = (stored != null).ToString(),
                    ["Duration"] = stopwatch.ElapsedMilliseconds
                });

            return stored?.CardJson;
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Error retrieving card '{UniqueId}' from team '{Team}', channel '{Channel}'", uniqueId, teamName, channelName);
            throw;
        }
    }

    public async Task CreateMessageToUserAsync(string message, string user, CancellationToken token)
    {
        try
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
                    try
                    {
                        var newResult = await turnContext.SendActivityAsync(MessageFactory.Text(message), cancellationToken);

                        telemetry.TrackEvent("ChatNewMessage",
                            new()
                            {
                                ["MessageId"] = newResult.Id,
                                ["Duration"] = stopwatch.ElapsedMilliseconds
                            });
                    }
                    catch (Exception ex)
                    {
                        logger.LogError(ex, "Error sending text message to user '{User}'", user);
                        throw;
                    }
                },
                token);
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Error creating message for user '{User}'", user);
            throw;
        }
    }

    public async Task CreateOrUpdateAsync<T>(string jsonFileName, T model, string user, CancellationToken token) where T : BaseTemplateModel
    {
        try
        {
            var stopwatch = Stopwatch.StartNew();
            var userAadObjectId = await teamsManagerService.GetUserAadObjectIdAsync(user, token);
            var installedAppId = await teamsManagerService.GetOrInstallChatAppIdAsync(userAadObjectId, token);

            if (string.IsNullOrWhiteSpace(installedAppId)) throw new InvalidOperationException($"Unable to install or retrieve chat app for user '{user}'");

            var chatId = await teamsManagerService.GetChatIdAsync(installedAppId, userAadObjectId, token);
            if (string.IsNullOrWhiteSpace(chatId)) throw new InvalidOperationException($"Unable to retrieve chat for user '{user}'");

            var stored = await cosmosMessageStore.FindByChatAsync(jsonFileName, model.UniqueId, token);
            var cardJson = await CreateCardFromTemplateAsync(jsonFileName, null, model, token: token);

            await CreateOrUpdateChatCardAsync(jsonFileName, model, cardJson, stored, chatId, stopwatch, token);
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Error creating or updating card for user '{User}'", user);
            throw;
        }
    }

    public async Task CreateOrUpdateAsync<T>(string jsonFileName, IFormFile? file, T model, string teamName, string channelName, CancellationToken token) where T : BaseTemplateModel
    {
        try
        {
            var stopwatch = Stopwatch.StartNew();
            var teamId = await teamsManagerService.GetTeamIdAsync(teamName, token);
            await teamsManagerService.CheckOrInstallBotIsInTeam(teamId, token);
            var channelId = await teamsManagerService.GetChannelIdAsync(teamId, channelName, token);

            var stored = await cosmosMessageStore.FindByChannelAsync(jsonFileName, model.UniqueId, token);
            var cardJson = await CreateCardFromTemplateAsync(jsonFileName, file, model, teamId, channelId, channelName, token);

            await CreateOrUpdateChannelCardAsync(jsonFileName, model, cardJson, stored, teamId, channelId, stopwatch, token);
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Error creating or updating card for team '{Team}' channel '{Channel}'", teamName, channelName);
            throw;
        }
    }

    public async Task RemoveActionsFromCardAsync(string teamId, string channelId, string messageId, string[] actionsToRemove, CancellationToken token)
    {
        try
        {
            var stopwatch = Stopwatch.StartNew();
            var stored = await cosmosMessageStore.FindByChannelMessageIdAsync(messageId, token);
            var cardJson = stored?.CardJson;

            if (string.IsNullOrWhiteSpace(cardJson))
            {
                var errorMsg = $"Card not found in team '{teamId}', channel '{channelId}', message '{messageId}'";
                logger.LogWarning(errorMsg);
                throw new InvalidOperationException(errorMsg);
            }

            var card = AdaptiveCard.FromJson(cardJson).Card;
            if (card == null)
            {
                var errorMsg = $"Failed to parse adaptive card for message '{messageId}'";
                logger.LogError(errorMsg);
                throw new InvalidOperationException(errorMsg);
            }

            // Remove all actions that match the verbs
            var actionsRemoved = 0;
            foreach (var actionVerb in actionsToRemove)
            {
                var actionsToRemoveList = card
                    .Actions
                    .Where(a => a is AdaptiveExecuteAction exe && exe.Verb == actionVerb)
                    .ToList();

                foreach (var adaptiveAction in actionsToRemoveList)
                {
                    card.Actions.Remove(adaptiveAction);
                    actionsRemoved++;
                }
            }

            var updatedCardJson = card.ToJson();
            var activity = new Activity
            {
                Type = "message",
                Id = messageId,
                Attachments = new List<Attachment>
                {
                    new()
                    {
                        ContentType = AdaptiveCard.ContentType,
                        Content = updatedCardJson
                    }
                }
            };

            var conversationReference = GetConversationReference(channelId);
            conversationReference.ActivityId = messageId;

            await adapter.ContinueConversationAsync(AgentClaims.CreateIdentity(_clientId),
                conversationReference,
                async (turnContext, cancellationToken) =>
                {
                    try
                    {
                        var updateResult = await turnContext.UpdateActivityAsync(activity, cancellationToken);
                        var updatedStored = stored! with
                        {
                            CardJson = updatedCardJson,
                            UpdatedAt = DateTimeOffset.UtcNow
                        };
                        await cosmosMessageStore.UpsertAsync(updatedStored, cancellationToken);

                        telemetry.TrackEvent("ChannelUpdateCardRemoveActions",
                            new()
                            {
                                ["Team"] = teamId,
                                ["Channel"] = channelId,
                                ["MessageId"] = updateResult.Id,
                                ["ActionsRemoved"] = string.Join(",", actionsToRemove),
                                ["ActionsRemovedCount"] = actionsRemoved.ToString(),
                                ["Duration"] = stopwatch.ElapsedMilliseconds
                            });
                    }
                    catch (Exception ex)
                    {
                        logger.LogError(ex, "Error updating card with removed actions for message '{MessageId}'", messageId);
                        throw;
                    }
                },
                token);
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Error removing actions from card in team '{Team}', channel '{Channel}', message '{Message}'", teamId, channelId, messageId);
            throw;
        }
    }

    private async Task CreateOrUpdateChatCardAsync<T>(string fileName, T model, string card, StoredMessage? stored, string chatId, Stopwatch stopwatch, CancellationToken token) where T : BaseTemplateModel
    {
        var activity = new Activity
        {
            Type = "message",
            Attachments = new List<Attachment>
            {
                new()
                {
                    ContentType = AdaptiveCard.ContentType,
                    Content = card
                }
            }
        };

        var conversationReference = GetConversationReference(chatId);
        var idFromOldMessage = stored?.Id;

        if (!string.IsNullOrWhiteSpace(idFromOldMessage))
        {
            activity.Id = idFromOldMessage;
            conversationReference.ActivityId = idFromOldMessage;
        }

        await adapter.ContinueConversationAsync(
            AgentClaims.CreateIdentity(_clientId),
            conversationReference,
            async (turnContext, cancellationToken) =>
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(idFromOldMessage))
                    {
                        // Create new message
                        var newResult = await turnContext.SendActivityAsync(activity, cancellationToken);
                        var persistedMessageId = ResolveMessageId(newResult.Id, idFromOldMessage, "send", fileName);
                        await UpsertChatStoredMessageAsync(persistedMessageId, chatId, fileName, model.UniqueId, card, null, cancellationToken);

                        telemetry.TrackEvent("ChatNewMessage",
                            new()
                            {
                                ["MessageId"] = persistedMessageId,
                                ["Duration"] = stopwatch.ElapsedMilliseconds
                            });
                    }
                    else
                    {
                        // Update existing message
                        var updateResult = await turnContext.UpdateActivityAsync(activity, cancellationToken);
                        var persistedMessageId = ResolveMessageId(updateResult.Id, idFromOldMessage, "update", fileName);
                        await UpsertChatStoredMessageAsync(persistedMessageId, chatId, fileName, model.UniqueId, card, stored, cancellationToken);

                        telemetry.TrackEvent("ChatUpdateMessage",
                            new()
                            {
                                ["MessageId"] = persistedMessageId,
                                ["Duration"] = stopwatch.ElapsedMilliseconds
                            });
                    }
                }
                catch (Exception ex)
                {
                    logger.LogError(ex, "Error sending or updating/adding chat message for {FileName}", fileName);
                    throw;
                }
            },
            token);
    }

    private async Task CreateOrUpdateChannelCardAsync<T>(string fileName, T model, string card, StoredMessage? stored, string teamId, string channelId, Stopwatch stopwatch, CancellationToken token) where T : BaseTemplateModel
    {
        var activity = new Activity
        {
            Type = "message",
            Attachments = new List<Attachment>
            {
                new()
                {
                    ContentType = AdaptiveCard.ContentType,
                    Content = card
                }
            }
        };

        var conversationReference = GetConversationReference(channelId);
        var idFromOldMessage = stored?.Id;

        if (!string.IsNullOrWhiteSpace(idFromOldMessage))
        {
            activity.Id = idFromOldMessage;
            conversationReference.ActivityId = idFromOldMessage;
        }

        await adapter.ContinueConversationAsync(
            AgentClaims.CreateIdentity(_clientId),
            conversationReference,
            async (turnContext, cancellationToken) =>
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(idFromOldMessage))
                    {
                        // Create new message
                        var newResult = await turnContext.SendActivityAsync(activity, cancellationToken);
                        var persistedMessageId = ResolveMessageId(newResult.Id, idFromOldMessage, "send", fileName);
                        await UpsertChannelStoredMessageAsync(persistedMessageId, teamId, channelId, fileName, model.UniqueId, card, null, cancellationToken);

                        telemetry.TrackEvent("ChannelNewMessage",
                            new()
                            {
                                ["TeamId"] = teamId,
                                ["ChannelId"] = channelId,
                                ["MessageId"] = persistedMessageId,
                                ["Duration"] = stopwatch.ElapsedMilliseconds
                            });
                    }
                    else
                    {
                        // Update existing message
                        var updateResult = await turnContext.UpdateActivityAsync(activity, cancellationToken);
                        var persistedMessageId = ResolveMessageId(updateResult.Id, idFromOldMessage, "update", fileName);
                        await UpsertChannelStoredMessageAsync(persistedMessageId, teamId, channelId, fileName, model.UniqueId, card, stored, cancellationToken);

                        telemetry.TrackEvent("ChannelUpdateMessage",
                            new()
                            {
                                ["TeamId"] = teamId,
                                ["ChannelId"] = channelId,
                                ["MessageId"] = persistedMessageId,
                                ["Duration"] = stopwatch.ElapsedMilliseconds
                            });
                    }
                }
                catch (Exception ex)
                {
                    logger.LogError(ex, "Error sending or updating channel message for {FileName} in team '{TeamId}' channel '{ChannelId}'", fileName, teamId, channelId);
                    throw;
                }
            },
            token);
    }

    public async Task<string> CreateCardFromTemplateAsync<T>(string jsonFileName, IFormFile? formFile, T model, string? teamId = null, string? channelId = null, string? channelName = null, CancellationToken token = default) where T : BaseTemplateModel
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
                    var result = await teamsManagerService.UploadFile(teamId, channelId, fileLocation, stream, token);
                    if (!result.Success)
                    {
                        telemetry.TrackEvent("FileUploadFailed",
                            new()
                            {
                                ["Team"] = teamId,
                                ["Channel"] = channelId,
                                ["FileName"] = fileName
                            });
                    }

                    // set file url
                    fileUrl = result.Url;
                }
                catch (Exception ex)
                {
                    logger.LogError(ex, "Error uploading file to Teams, continuing");
                }
            }
        }

        // replace all props with the values
        //if file prop is empty, it will just leave the value out
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

    private Task UpsertChatStoredMessageAsync(string messageId, string chatId, string jsonFileName, string uniqueId, string cardJson, StoredMessage? existing, CancellationToken token)
    {
        var now = DateTimeOffset.UtcNow;
        var doc = new StoredMessage
        {
            Id = messageId,
            UniqueId = uniqueId,
            JsonFileName = jsonFileName,
            ChatId = chatId,
            CardJson = cardJson,
            CreatedAt = existing?.CreatedAt ?? now,
            UpdatedAt = now
        };
        return cosmosMessageStore.UpsertAsync(doc, token);
    }

    private static string ResolveMessageId(string? candidateMessageId, string? fallbackMessageId, string operation, string fileName)
    {
        var resolved = !string.IsNullOrWhiteSpace(candidateMessageId) ? candidateMessageId : fallbackMessageId;
        if (!string.IsNullOrWhiteSpace(resolved)) return resolved;

        throw new InvalidOperationException($"Teams did not return a valid message id during '{operation}' for '{fileName}'.");
    }

    private Task UpsertChannelStoredMessageAsync(string messageId, string teamId, string channelId, string jsonFileName, string uniqueId, string cardJson, StoredMessage? existing, CancellationToken token)
    {
        var now = DateTimeOffset.UtcNow;
        var doc = new StoredMessage
        {
            Id = messageId,
            UniqueId = uniqueId,
            JsonFileName = jsonFileName,
            TeamId = teamId,
            ChannelId = channelId,
            CardJson = cardJson,
            CreatedAt = existing?.CreatedAt ?? now,
            UpdatedAt = now
        };
        return cosmosMessageStore.UpsertAsync(doc, token);
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