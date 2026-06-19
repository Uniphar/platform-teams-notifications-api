using System.Data;
using System.Text.Json;
using Azure.Messaging.ServiceBus;
using Azure.Messaging.ServiceBus.Administration;
using Teams.Notifications.Api.Tests.TeamsClient;
using IntegrationSuiteErrorRequest = Teams.Notifications.Api.Tests.TeamsClient.IntegrationSuiteErrorModel;
using LogicAppErrorRequest = Teams.Notifications.Api.Tests.TeamsClient.LogicAppErrorModel;

namespace Teams.Notifications.Api.Tests.IntegrationTests;

[TestClass]
[TestCategory("Integration")]
public sealed class TeamsNotificationApiIntegrationTests
{
    private const string CardEventsTopicName = "platform-teams-notification-card-events";
    private const string DevTeamName = "DAWN - Integrations Errors Dev";
    private const string TestTeamName = "DAWN - Integrations Errors Test";
    private const string LogicAppChannelName = "Logic App Errors";
    private const string IntegrationSuiteChannelName = "Integration Suite Errors";
    private static string _subscriptionName = string.Empty;

    private static CancellationToken _cancellationToken;
    private static TeamsNotificationApi _client = null!;
    private static string _channelTeamName = string.Empty;
    private static string _serviceBusNamespace = string.Empty;

    [ClassCleanup]
    public static async Task ClassCleanup()
    {
        await DeleteSubscriptionIfPresentAsync(_subscriptionName);
    }

    [ClassInitialize]
    public static async Task ClassInitialize(TestContext context)
    {
        _cancellationToken = context.CancellationToken;
        var runEnv = (context.Properties["Environment"]?.ToString() ?? "dev").Trim().ToLowerInvariant();
        var externalTenantId = context.Properties["AZURE_ENTRA_EXTERNAL_TENANT_ID"]?.ToString() ?? throw new NoNullAllowedException("Needs to be initialised for integration test");

        var apiProjectPath = Path.GetFullPath("../../../../Teams.Notifications.Api");
        var env = runEnv == "local" ? "dev" : runEnv;
        var config = new ConfigurationBuilder()
            .SetBasePath(apiProjectPath)
            .AddJsonFile("appsettings.json", false, false)
            .AddJsonFile($"appsettings.{runEnv}.json", true, false)
            .AddAzureKeyVault(new($"https://uni-devops-app-{env}-kv.vault.azure.net/"), new DefaultAzureCredential())
            .Build();

        _serviceBusNamespace = ResolveServiceBusNamespace(config);
        var tokenCredentials = new ClientSecretCredential(externalTenantId,
            config["integration-test-platform-teams-notification-api-client-id"] ?? throw new NoNullAllowedException("integration-test-platform-teams-notification-api-client-id missing in configuration"),
            config["integration-test-platform-teams-notification-api-client-secret"] ?? throw new NoNullAllowedException("integration-test-platform-teams-notification-api-client-secret missing in configuration"));
        var customHandler = new HttpClientAuthorizationDelegatingHandler(tokenCredentials, config)
        {
            InnerHandler = new HttpClientHandler()
        };
        var httpClient = new HttpClient(customHandler)
        {
            // if prod no .prod in the URL
            BaseAddress = new($"https://api.{env}.uniphar.ie/".Replace(".prod", ""))
        };

        _client = new(httpClient);
        _channelTeamName = env == "test" ? TestTeamName : DevTeamName;
        var subscriptionName = $"teams-tests-sub-{Guid.NewGuid().ToString("N")[..5]}";
        _subscriptionName = await CreateTemporaryCardEventsSubscriptionAsync(subscriptionName, _cancellationToken);
    }

    [TestMethod]
    public async Task IntegrationSuiteError_CanBeAddedAndRemoved_EndToEnd()
    {
        var uniqueId = $"int-test-{Guid.NewGuid():N}";

        var model = new IntegrationSuiteErrorRequest
        {
            UniqueId = uniqueId,
            FlowName = "Teams notification API integration test",
            TimeStamp = DateTime.UtcNow.ToString("O"),
            FlowMessageId = Guid.NewGuid().ToString("N"),
            Source = "IntegrationTest",
            Destination = "Teams",
            Status = "Failed",
            Intermediary = "Automation",
            PurchaseOrderByCustomer = "PO-INT-001",
            Error = "Created by TeamsNotificationApi integration test"
        };

        try
        {
            await _client.IntegrationSuiteErrorPOSTAsync(_channelTeamName, IntegrationSuiteChannelName, model, _cancellationToken);
            await AssertCardEventPublishedAsync(_subscriptionName, uniqueId, "TeamsCardCreated", _cancellationToken);
            await _client.IntegrationSuiteErrorGetsAnObject(uniqueId, _channelTeamName, IntegrationSuiteChannelName, cancellationToken: _cancellationToken);

            var card = await _client.IntegrationSuiteErrorGETAsync(uniqueId, _channelTeamName, IntegrationSuiteChannelName, _cancellationToken);
            Assert.IsFalse(string.IsNullOrWhiteSpace(card));
            StringAssert.Contains(card, uniqueId);
        }
        finally
        {
            await DeleteIfPresentAsync(() => _client.IntegrationSuiteErrorDELETEAsync(uniqueId, _channelTeamName, IntegrationSuiteChannelName, CancellationToken.None));
        }

        await _client.IntegrationSuiteErrorGetsEmpty(uniqueId, _channelTeamName, IntegrationSuiteChannelName, cancellationToken: _cancellationToken);
    }

    [TestMethod]
    public async Task LogicAppError_CanBeAddedAndRemoved_EndToEnd()
    {
        var uniqueId = $"int-test-{Guid.NewGuid():N}";
        var model = new LogicAppErrorRequest
        {
            UniqueId = uniqueId,
            LogicAppFlow = "Teams notification API integration test",
            TimeStamp = DateTime.UtcNow.ToString("O"),
            OriginalBlobUri = "https://storage.example.test/blob.csv",
            ObjectType = "IntegrationTest",
            ErrorMessage = "Created by TeamsNotificationApi integration test"
        };

        try
        {
            await _client.LogicAppErrorPOSTAsync(_channelTeamName, LogicAppChannelName, model, _cancellationToken);
            await AssertCardEventPublishedAsync(_subscriptionName, uniqueId, "TeamsCardCreated", _cancellationToken);
            await _client.LogicAppErrorGetsAnObject(uniqueId, _channelTeamName, LogicAppChannelName, cancellationToken: _cancellationToken);

            var card = await _client.LogicAppErrorGETAsync(uniqueId, _channelTeamName, LogicAppChannelName, _cancellationToken);
            Assert.IsFalse(string.IsNullOrWhiteSpace(card));
            StringAssert.Contains(card, uniqueId);
        }
        finally
        {
            await DeleteIfPresentAsync(() => _client.LogicAppErrorDELETEAsync(uniqueId, _channelTeamName, LogicAppChannelName, CancellationToken.None));
        }

        await _client.LogicAppErrorGetsEmpty(uniqueId, _channelTeamName, LogicAppChannelName, cancellationToken: _cancellationToken);
    }


    private static async Task DeleteIfPresentAsync(Func<Task> deleteAction)
    {
        try
        {
            await deleteAction();
        }
        catch (ApiException ex) when (ex.StatusCode is 400 or 403 or 404) { }
    }

    private static async Task<string> CreateTemporaryCardEventsSubscriptionAsync(string subscriptionName, CancellationToken cancellationToken)
    {
        EnsureServiceBusConfiguration();

        var adminClient = new ServiceBusAdministrationClient(_serviceBusNamespace, new DefaultAzureCredential());
        var options = new CreateSubscriptionOptions(CardEventsTopicName, subscriptionName)
        {
            AutoDeleteOnIdle = TimeSpan.FromHours(1),
            DefaultMessageTimeToLive = TimeSpan.FromDays(1),
            MaxDeliveryCount = 5,
            LockDuration = TimeSpan.FromMinutes(5),
            EnableBatchedOperations = true,
            UserMetadata = "teams-notifications-api-integration-test"
        };

        try
        {
            await adminClient.CreateSubscriptionAsync(options, cancellationToken);
        }
        catch (ServiceBusException ex) when (ex.Reason == ServiceBusFailureReason.MessagingEntityNotFound)
        {
            Assert.Inconclusive($"Service Bus topic '{CardEventsTopicName}' was not found in namespace '{_serviceBusNamespace}'. Configure the topic for this environment to enable e2e message assertions.");
        }

        return subscriptionName;
    }

    private static async Task DeleteSubscriptionIfPresentAsync(string subscriptionName)
    {
        if (string.IsNullOrWhiteSpace(_serviceBusNamespace)) return;
        var adminClient = new ServiceBusAdministrationClient(_serviceBusNamespace, new DefaultAzureCredential());
        try
        {
            await adminClient.DeleteSubscriptionAsync(CardEventsTopicName, subscriptionName);
        }
        catch (ServiceBusException ex) when (ex.Reason == ServiceBusFailureReason.MessagingEntityNotFound) { }
    }

    private static async Task AssertCardEventPublishedAsync(string subscriptionName, string expectedUniqueId, string expectedSubject, CancellationToken cancellationToken)
    {
        EnsureServiceBusConfiguration();
        var client = new ServiceBusClient(_serviceBusNamespace, new DefaultAzureCredential());
        var receiver = client.CreateReceiver(CardEventsTopicName,
            subscriptionName,
            new()
            {
                ReceiveMode = ServiceBusReceiveMode.PeekLock
            });

        var timeoutAt = DateTimeOffset.UtcNow.AddMinutes(2);
        while (DateTimeOffset.UtcNow < timeoutAt)
        {
            var message = await receiver.ReceiveMessageAsync(TimeSpan.FromSeconds(10), cancellationToken);
            if (message is null) continue;

            if (!string.Equals(message.Subject, expectedSubject, StringComparison.Ordinal)) continue;

            using var body = JsonDocument.Parse(message.Body);
            if (!body.RootElement.TryGetProperty("uniqueId", out var idProperty)) continue;
            if (string.Equals(idProperty.GetString(), expectedUniqueId, StringComparison.Ordinal)) return;
        }

        Assert.Fail($"Expected Service Bus message '{expectedSubject}' for uniqueId '{expectedUniqueId}' was not received.");
    }

    private static void EnsureServiceBusConfiguration()
    {
        if (string.IsNullOrWhiteSpace(_serviceBusNamespace)) Assert.Inconclusive("Service Bus namespace is not configured for integration tests. Configure 'FullyQualifiedNamespaceServiceBus' in appsettings/KeyVault.");
    }

    private static string ResolveServiceBusNamespace(IConfiguration config) => config["FullyQualifiedNamespaceServiceBus"] ?? string.Empty;
}