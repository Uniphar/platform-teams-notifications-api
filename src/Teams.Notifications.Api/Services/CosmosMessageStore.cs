namespace Teams.Notifications.Api.Services;

public sealed class CosmosMessageStore(CosmosClient client, IOptions<CosmosOptions> options) : ICosmosMessageStore
{
    private readonly Container _container = client.GetContainer(options.Value.DatabaseName, options.Value.ContainerName);

    public Task<StoredMessage?> FindByChatAsync(string chatId, string jsonFileName, string uniqueId, CancellationToken token) => QuerySingleAsync(StoredMessage.ChatPartition(chatId), jsonFileName, uniqueId, token);

    public Task<StoredMessage?> FindByChannelAsync(string teamId, string channelId, string jsonFileName, string uniqueId, CancellationToken token) => QuerySingleAsync(StoredMessage.ChannelPartition(teamId, channelId), jsonFileName, uniqueId, token);

    public async Task<StoredMessage?> FindByChannelMessageIdAsync(string teamId, string channelId, string messageId, CancellationToken token)
    {
        try
        {
            var response = await _container.ReadItemAsync<StoredMessage>(messageId, new(StoredMessage.ChannelPartition(teamId, channelId)), cancellationToken: token);
            return response.Resource;
        }
        catch (CosmosException ex) when (ex.StatusCode == HttpStatusCode.NotFound)
        {
            return null;
        }
    }

    public async Task UpsertAsync(StoredMessage message, CancellationToken token)
    {
        await _container.UpsertItemAsync(message, new(message.PartitionKey), cancellationToken: token);
    }

    public async Task DeleteAsync(string partitionKey, string messageId, CancellationToken token)
    {
        try
        {
            await _container.DeleteItemAsync<StoredMessage>(messageId, new(partitionKey), cancellationToken: token);
        }
        catch (CosmosException ex) when (ex.StatusCode == HttpStatusCode.NotFound)
        {
            // already gone, nothing to do
        }
    }

    public async Task EnsureContainerIsProvisioned()
    {
        var dbResponse = await client.CreateDatabaseIfNotExistsAsync(options.Value.DatabaseName);

        var containerProperties = new ContainerProperties(options.Value.ContainerName, "/pk")
        {
            DefaultTimeToLive = options.Value.RetentionDays * 86400
        };

        await dbResponse.Database.CreateContainerIfNotExistsAsync(containerProperties);

        var containerResponse = await _container.ReadContainerAsync();
        var partitionKeyPath = containerResponse.Resource.PartitionKeyPath;
        if (!string.Equals(partitionKeyPath, "/pk", StringComparison.Ordinal))
        {
            throw new InvalidOperationException(
                $"Cosmos container '{options.Value.ContainerName}' in database '{options.Value.DatabaseName}' has partition key path '{partitionKeyPath}'. Expected '/pk'.");
        }
    }

    private async Task<StoredMessage?> QuerySingleAsync(string partitionKey, string jsonFileName, string uniqueId, CancellationToken token)
    {
        var query = new QueryDefinition("SELECT TOP 1 * FROM c WHERE c.uniqueId = @uniqueId AND c.jsonFileName = @jsonFileName ORDER BY c._ts DESC")
            .WithParameter("@uniqueId", uniqueId)
            .WithParameter("@jsonFileName", jsonFileName);

        var iterator = _container.GetItemQueryIterator<StoredMessage>(query, requestOptions: new() { PartitionKey = new(partitionKey), MaxItemCount = 1 });
        if (!iterator.HasMoreResults) return null;
        var page = await iterator.ReadNextAsync(token);
        return page.FirstOrDefault();
    }
}