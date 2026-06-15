namespace Teams.Notifications.Api.Services;

public sealed class CosmosMessageStore(CosmosClient client, IOptions<CosmosOptions> options) : ICosmosMessageStore
{
    private readonly Container _container = client.GetContainer(options.Value.DatabaseName, options.Value.ContainerName);


    public Task<StoredMessage?> FindMessageByUniqueId(string uniqueId, CancellationToken token) => ReadByUniqueIdAsync(uniqueId, token);

    public async Task<StoredMessage?> FindByChannelMessageIdAsync(string messageId, CancellationToken token)
    {
        var query = new QueryDefinition("SELECT TOP 1 * FROM c WHERE c.messageId = @messageId")
            .WithParameter("@messageId", messageId);

        var iterator = _container.GetItemQueryIterator<StoredMessage>(query, requestOptions: new() { MaxItemCount = 1 });
        if (!iterator.HasMoreResults) return null;
        var page = await iterator.ReadNextAsync(token);
        return page.FirstOrDefault();
    }

    public async Task UpsertAsync(StoredMessage message, CancellationToken token)
    {
        ValidateForUpsert(message);
        await _container.UpsertItemAsync(message, new(message.Id), cancellationToken: token);
    }

    public async Task DeleteAsync(string uniqueId, CancellationToken token)
    {
        try
        {
            await _container.DeleteItemAsync<StoredMessage>(uniqueId, new(uniqueId), cancellationToken: token);
        }
        catch (CosmosException ex) when (ex.StatusCode == HttpStatusCode.NotFound)
        {
            // already gone, nothing to do
        }
    }

    public async Task EnsureContainerIsProvisioned()
    {
        var dbResponse = await client.CreateDatabaseIfNotExistsAsync(options.Value.DatabaseName);

        var containerProperties = new ContainerProperties(options.Value.ContainerName, "/id")
        {
            DefaultTimeToLive = options.Value.RetentionDays * 86400
        };

        await dbResponse.Database.CreateContainerIfNotExistsAsync(containerProperties);
    }

    private async Task<StoredMessage?> ReadByUniqueIdAsync(string uniqueId, CancellationToken token)
    {
        try
        {
            var response = await _container.ReadItemAsync<StoredMessage>(uniqueId, new(uniqueId), cancellationToken: token);
            return response.Resource;
        }
        catch (CosmosException ex) when (ex.StatusCode == HttpStatusCode.NotFound)
        {
            return null;
        }
    }

    private static void ValidateForUpsert(StoredMessage message)
    {
        if (string.IsNullOrWhiteSpace(message.Id)) throw new InvalidOperationException("Cannot upsert StoredMessage with an empty id.");

        if (message.Id.Any(c => c is '/' or '\\' or '?' or '#' || char.IsControl(c)))
        {
            throw new InvalidOperationException(
                $"Cannot upsert StoredMessage with id '{message.Id}'. Cosmos id cannot contain '/', '\\', '?', '#', or control characters.");
        }
    }
}