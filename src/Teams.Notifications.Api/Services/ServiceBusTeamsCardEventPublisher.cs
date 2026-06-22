namespace Teams.Notifications.Api.Services;

public sealed class ServiceBusTeamsCardEventPublisher : ITeamsCardEventPublisher, IAsyncDisposable
{
    internal const string TopicName = "platform-teams-notification-card-events";

    private readonly ServiceBusSender _sender;
    private readonly string _tenantId;

    public ServiceBusTeamsCardEventPublisher(IConfiguration configuration)
    {
        var ns = configuration["FullyQualifiedNamespaceServiceBus"] ?? throw new InvalidOperationException("FullyQualifiedNamespaceServiceBus is required");
        _tenantId = configuration["AZURE_TENANT_ID"] ?? throw new InvalidOperationException("AZURE_TENANT_ID is required");

        var client = new ServiceBusClient(ns, new DefaultAzureCredential());
        _sender = client.CreateSender(TopicName);
    }

    public ValueTask DisposeAsync() => _sender.DisposeAsync();

    public Task PublishCardCreatedAsync(string uniqueId, string teamsDeepLink, string cardType, bool hasActionableFile, CancellationToken token) =>
        SendAsync(new TeamsCardCreatedCommand
            {
                UniqueId = uniqueId,
                TeamsDeepLink = teamsDeepLink,
                CardType = cardType,
                HasActionableFile = hasActionableFile,
                TenantId = _tenantId
            },
            "TeamsCardCreated",
            token);

    public Task PublishCardUpdatedAsync(string uniqueId, string teamsDeepLink, string cardType, bool hasActionableFile, CancellationToken token) =>
        SendAsync(new TeamsCardUpdatedCommand
            {
                UniqueId = uniqueId,
                TeamsDeepLink = teamsDeepLink,
                CardType = cardType,
                HasActionableFile = hasActionableFile,
                TenantId = _tenantId
            },
            "TeamsCardUpdated",
            token);

    public Task PublishCardDeletedAsync(string uniqueId, CancellationToken token) => SendAsync(new TeamsCardDeletedCommand { UniqueId = uniqueId, TenantId = _tenantId }, "TeamsCardDeleted", token);

    public Task PublishCardResolvedAsync(string uniqueId, CancellationToken token) => SendAsync(new TeamsCardResolvedCommand { UniqueId = uniqueId, TenantId = _tenantId }, "TeamsCardResolved", token);

    private Task SendAsync<T>(T command, string subject, CancellationToken token)
    {
        var message = new ServiceBusMessage(JsonSerializer.SerializeToUtf8Bytes(command))
        {
            Subject = subject,
            ContentType = "application/json",
            MessageId = Guid.NewGuid().ToString("N")
        };
        return _sender.SendMessageAsync(message, token);
    }
}