namespace Teams.Notifications.Api.Services.Interfaces;

public interface ICosmosMessageStore
{
    Task<StoredMessage?> FindByChatAsync(string jsonFileName, string uniqueId, CancellationToken token);
    Task<StoredMessage?> FindByChannelAsync(string jsonFileName, string uniqueId, CancellationToken token);
    Task<StoredMessage?> FindByChannelMessageIdAsync(string messageId, CancellationToken token);
    Task UpsertAsync(StoredMessage message, CancellationToken token);
    Task DeleteAsync(string uniqueId, CancellationToken token);
    Task EnsureContainerIsProvisioned();
}