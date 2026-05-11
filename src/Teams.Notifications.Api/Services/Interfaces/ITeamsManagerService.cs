namespace Teams.Notifications.Api.Services.Interfaces;

public interface ITeamsManagerService
{
    Task CheckOrInstallBotIsInTeam(string teamId, CancellationToken token);
    Task<string> GetTeamIdAsync(string teamName, CancellationToken token);
    Task<string> GetChannelIdAsync(string teamId, string channelName, CancellationToken token);
    Task<string> GetGroupNameUniqueName(string groupId, CancellationToken token);
    Task<string> GetTeamName(string teamId, CancellationToken token);
    Task<string> GetUserAadObjectIdAsync(string userPrincipalName, CancellationToken token);
    Task<string?> GetOrInstallChatAppIdAsync(string aadObjectId, CancellationToken token);
    Task<string?> GetChatIdAsync(string installedAppId, string aadObjectId, CancellationToken token);
    Task<ChatMessage?> GetChatMessageByUniqueId(string chatId, string jsonFileName, string uniqueId, CancellationToken token);
    Task<string?> GetMessageIdByUniqueId(string teamId, string channelId, string jsonFileName, string uniqueId, CancellationToken token);
    Task<ChatMessage?> GetMessageByUniqueId(string teamId, string channelId, string jsonFileName, string uniqueId, CancellationToken token);
    Task<ChatMessage?> GetMessageById(string teamId, string channelId, string messageId, CancellationToken token);
    Task<(bool Success, string Url)> UploadFile(string teamId, string channelId, string fileLocation, Stream fileStream, CancellationToken token);
    Task<string> GetFileNameAsync(string teamId, string channelId, string fileLocation, CancellationToken token);
}