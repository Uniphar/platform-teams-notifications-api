namespace Teams.Notifications.Api.Services;

public sealed class CosmosMessageStore(CosmosClient client, IOptions<CosmosOptions> options) : ICosmosMessageStore
{

    private readonly Container _container = client.GetContainer(options.Value.DatabaseName, options.Value.ContainerName);

    public async Task<StoredMessage?> FindByChatAsync(string chatId, string jsonFileName, string uniqueId, CancellationToken token) => await QuerySingleAsync(jsonFileName, uniqueId, token);

    public async Task<StoredMessage?> FindByChannelAsync(string teamId, string channelId, string jsonFileName, string uniqueId, CancellationToken token) => await QuerySingleAsync(jsonFileName, uniqueId, token);

    public async Task<StoredMessage?> FindByChannelMessageIdAsync(string teamId, string channelId, string messageId, CancellationToken token)
    {
        try
        {
            var response = await _container.ReadItemAsync<StoredMessage>(messageId, PartitionKey.None, cancellationToken: token);
            return response.Resource;
        }
        catch (CosmosException ex) when (ex.StatusCode == HttpStatusCode.NotFound)
        {
            return null;
        }
    }

    public async Task UpsertAsync(StoredMessage message, CancellationToken token)
    {
        ValidateForUpsert(message);
        await _container.UpsertItemAsync(message, PartitionKey.None, cancellationToken: token);
    }

    public async Task DeleteAsync(string messageId, CancellationToken token)
    {
        try
        {
            await _container.DeleteItemAsync<StoredMessage>(messageId, PartitionKey.None, cancellationToken: token);
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

    private async Task<StoredMessage?> QuerySingleAsync(string jsonFileName, string uniqueId, CancellationToken token)
    {
        var query = new QueryDefinition("SELECT TOP 1 * FROM c WHERE c.uniqueId = @uniqueId AND c.jsonFileName = @jsonFileName ORDER BY c._ts DESC")
            .WithParameter("@uniqueId", uniqueId)
            .WithParameter("@jsonFileName", jsonFileName);

        var iterator = _container.GetItemQueryIterator<StoredMessage>(query, requestOptions: new() { MaxItemCount = 1 });
        if (!iterator.HasMoreResults) return null;
        var page = await iterator.ReadNextAsync(token);
        return page.FirstOrDefault();
    }

    private static void ValidateForUpsert(StoredMessage message)
    {
        if (string.IsNullOrWhiteSpace(message.Id))
        {
            throw new InvalidOperationException("Cannot upsert StoredMessage with an empty id.");
        }

        if (message.Id.Any(c => c is '/' or '\\' or '?' or '#' || char.IsControl(c)))
        {
            throw new InvalidOperationException(
                $"Cannot upsert StoredMessage with id '{message.Id}'. Cosmos id cannot contain '/', '\\', '?', '#', or control characters.");
        }
    }
}