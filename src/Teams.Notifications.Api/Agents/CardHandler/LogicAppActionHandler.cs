namespace Teams.Notifications.Api.Agents.CardHandler;

internal static class LogicAppActionHandler
{
    internal static async Task<AdaptiveCardInvokeResponse> HandleProcessVerbLogicAppAsync(this ITurnContext turnContext,
        object data,
        ICustomEventTelemetryClient telemetry,
        ILogger logger,
        TeamsManagerService teamsManagerService,
        IFrontgateApiService frontgateApiService,
        ICardManagerService cardManagerService,
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
                    telemetry.TrackEvent("LogAppProcessFile_NoChannelName");
                    return AdaptiveCardInvokeResponseFactory.BadRequest("Something went wrong reprocessing the file");
                }

                var teamId = teamDetails.AadGroupId;
                var channelName = channel.Name;

                if (string.IsNullOrWhiteSpace(teamId) || string.IsNullOrWhiteSpace(channelName)) throw new InvalidOperationException("Team or channelName is missing from the context.");

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
                var uploadResponse = await frontgateApiService.UploadFileAsync(model.PostOriginalBlobUri ?? string.Empty, fileInfo, cancellationToken);

                if (uploadResponse.IsSuccessStatusCode)
                {
                    await cardManagerService.RemoveActionsFromCardAsync(teamId, channelId, messageId, ["Process"], cancellationToken);

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

                var messageToUser = "Something went wrong sending file";
                try
                {
                    var errorMessage = await uploadResponse.Content.ReadAsStringAsync(cancellationToken);
                    if (!string.IsNullOrWhiteSpace(errorMessage)) messageToUser = $"Failed to send file: {errorMessage}";
                }
                catch (Exception ex)
                {
                    //Do nothing, we just sent the user the error message
                    logger.LogWarning(ex, "Failed to read error message from upload response");
                }

                telemetry.TrackEvent("LogAppProcessFileUploadFailed",
                    new()
                    {
                        ["Team"] = teamName,
                        ["Channel"] = channelName,
                        ["FileName"] = fileName,
                        ["MessageId"] = messageId,
                        ["StatusCode"] = uploadResponse.StatusCode
                    });
                return AdaptiveCardInvokeResponseFactory.BadRequest(messageToUser);
            }
            catch (Exception ex)
            {
                telemetry.TrackEvent("LogAppProcessFileError",
                    new()
                    {
                        ["Error"] = ex.Message
                    });
                throw;
            }
        }
    }
}