namespace Teams.Notifications.Api.Agents.CardHandler;

internal static class LogicAppActionHandler
{
    internal static async Task<AdaptiveCardInvokeResponse> HandleProcessVerbLogicAppAsync(this ITurnContext turnContext,
        object data,
        ICustomEventTelemetryClient telemetry,
        ILogger logger,
        ITeamsManagerService teamsManagerService,
        IFrontgateApiService frontgateApiService,
        ICardManagerService cardManagerService,
        ITeamsCardEventPublisher teamsCardEventPublisher,
        CancellationToken cancellationToken
    )
    {
        using (telemetry.WithProperties([new("ActionExecute", "LogicAppErrorProcessActionModel")]))
        {
            try
            {
                var model = ProtocolJsonSerializer.ToObject<LogicAppErrorProcessActionModel>(data);
                var teamsChannelData = turnContext.Activity.GetChannelData<TeamsChannelData>();

                var teamDetails = await TeamsInfo.GetTeamDetailsAsync(turnContext, cancellationToken: cancellationToken);
                var channels = await TeamsInfo.GetTeamChannelsAsync(turnContext, cancellationToken: cancellationToken);
                var channel = channels.FirstOrDefault(x => x.Id == teamsChannelData.Channel.Id);

                if (channel?.Name == null)
                {
                    var errorMsg = "Something went wrong reprocessing the file: channel name is null or missing";
                    logger.LogError(errorMsg);
                    telemetry.TrackEvent("LogAppProcessFile_NoChannelName");
                    throw new InvalidOperationException(errorMsg);
                }

                var teamId = teamDetails.AadGroupId;
                var channelName = channel.Name;

                if (string.IsNullOrWhiteSpace(teamId) || string.IsNullOrWhiteSpace(channelName))
                {
                    var errorMsg = "Team or channelName is missing from the context.";
                    logger.LogError(errorMsg);
                    throw new InvalidOperationException(errorMsg);
                }

                var channelId = await teamsManagerService.GetChannelIdAsync(teamId, channelName, cancellationToken);
                var fileName = await teamsManagerService.GetFileNameAsync(teamId, channelId, model.PostFileLocation ?? string.Empty, cancellationToken);
                var groupUniqueName = await teamsManagerService.GetGroupNameUniqueName(teamId, cancellationToken);
                var teamName = await teamsManagerService.GetTeamName(teamId, cancellationToken);

                // in a conversation, the ReplyToId is the message id, instead of the normal id
                // see https://learn.microsoft.com/en-us/microsoftteams/platform/task-modules-and-cards/cards/cards-actions?tabs=csharp#example-of-incoming-invoke-message
                var messageId = turnContext.Activity.ReplyToId;

                var fileInfo = new LogicAppFrontgateFileInformation
                {
                    file_name = fileName,
                    storage_reference = groupUniqueName,
                    initial_display_name = teamName,
                    storage_folder = $"/{channelName}/error/"
                };

                // Upload the file to the external API
                using var uploadResponse = await frontgateApiService.UploadFileAsync(model.PostOriginalBlobUri ?? string.Empty, fileInfo, cancellationToken);

                if (uploadResponse.IsSuccessStatusCode)
                {
                    await cardManagerService.RemoveActionsFromCardAsync(teamId, channelId, messageId, ["Process"], cancellationToken);

                    // Publish resolved event via Service Bus — best-effort, do not fail the card action if SB is unavailable
                    if (!string.IsNullOrWhiteSpace(model.PostUniqueId))
                    {
                        try
                        {
                            await teamsCardEventPublisher.PublishCardResolvedAsync(model.PostUniqueId, cancellationToken);
                            telemetry.TrackEvent("TeamsCardResolvedPublished", new() { ["UniqueId"] = model.PostUniqueId });
                        }
                        catch (Exception ex)
                        {
                            logger.LogError(ex, "Failed to publish TeamsCardResolved for UniqueId {UniqueId}", model.PostUniqueId);
                            telemetry.TrackEvent("TeamsCardResolvedPublishFailed", new() { ["UniqueId"] = model.PostUniqueId, ["Error"] = ex.Message });
                        }
                    }

                    telemetry.TrackEvent("ReprocessFileSuccess",
                        new()
                        {
                            ["Team"] = teamName,
                            ["Channel"] = channelName,
                            ["FileName"] = fileName,
                            ["MessageId"] = messageId
                        });

                    return AdaptiveCardInvokeResponseFactory.Message(model.PostSuccessMessage ?? "Success");
                }

                var errorMessage = "Something went wrong sending file";
                try
                {
                    var responseContent = await uploadResponse.Content.ReadAsStringAsync(cancellationToken);
                    if (!string.IsNullOrWhiteSpace(responseContent)) errorMessage = $"Failed to send file: {responseContent}";
                }
                catch (Exception ex)
                {
                    logger.LogError(ex, "Failed to read error message from upload response");
                }

                var uploadFailureMsg = $"File upload failed with status code {uploadResponse.StatusCode}: {errorMessage}";
                logger.LogError(uploadFailureMsg);
                telemetry.TrackEvent("LogAppProcessFileUploadFailed",
                    new()
                    {
                        ["Team"] = teamName,
                        ["Channel"] = channelName,
                        ["FileName"] = fileName,
                        ["MessageId"] = messageId,
                        ["StatusCode"] = uploadResponse.StatusCode,
                        ["ErrorMessage"] = errorMessage
                    });
                throw new InvalidOperationException(uploadFailureMsg);
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error processing LogicApp file action");
                telemetry.TrackEvent("LogAppProcessFileError",
                    new()
                    {
                        ["Error"] = ex.Message,
                        ["ExceptionType"] = ex.GetType().Name
                    });
                throw;
            }
        }
    }
}
