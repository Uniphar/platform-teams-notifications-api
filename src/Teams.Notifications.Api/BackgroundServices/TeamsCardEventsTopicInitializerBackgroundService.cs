namespace Teams.Notifications.Api.BackgroundServices;

/// <summary>
/// Ensures the Teams card events topic exists before the API starts handling requests.
/// </summary>
public sealed class TeamsCardEventsTopicInitializerBackgroundService(
    IConfiguration configuration,
    ILogger<TeamsCardEventsTopicInitializerBackgroundService> logger) : IHostedService
{
    public async Task StartAsync(CancellationToken cancellationToken)
    {
        var serviceBusNamespace = configuration["FullyQualifiedNamespaceServiceBus"]
                                  ?? throw new InvalidOperationException("FullyQualifiedNamespaceServiceBus is required");

        var adminClient = new ServiceBusAdministrationClient(serviceBusNamespace, new DefaultAzureCredential());

        if (!await adminClient.TopicExistsAsync(ServiceBusTeamsCardEventPublisher.TopicName, cancellationToken))
        {
            try
            {
                await adminClient.CreateTopicAsync(ServiceBusTeamsCardEventPublisher.TopicName, cancellationToken);
            }
            catch (ServiceBusException ex) when (ex.Reason == ServiceBusFailureReason.MessagingEntityAlreadyExists)
            {
                // Safe race: another instance created the topic.
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Failed creating Service Bus topic {Topic}", ServiceBusTeamsCardEventPublisher.TopicName);
                throw;
            }
        }
    }

    public Task StopAsync(CancellationToken cancellationToken) => Task.CompletedTask;
}
