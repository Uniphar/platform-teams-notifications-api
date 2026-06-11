namespace Teams.Notifications.Api.Services.Interfaces;

public interface ICosmosMessageStore
{
    Task<StoredMessage?> FindByChatAsync(string chatId, string jsonFileName, string uniqueId, CancellationToken token);
    Task<StoredMessage?> FindByChannelAsync(string teamId, string channelId, string jsonFileName, string uniqueId, CancellationToken token);
    Task<StoredMessage?> FindByChannelMessageIdAsync(string teamId, string channelId, string messageId, CancellationToken token);
    Task UpsertAsync(StoredMessage message, CancellationToken token);
    Task DeleteAsync(string messageId, CancellationToken token);
    Task EnsureContainerIsProvisioned();
}