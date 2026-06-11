using System.Data;
using Teams.Notifications.Api.Tests.TeamsClient;
using IntegrationSuiteErrorRequest = Teams.Notifications.Api.Tests.TeamsClient.IntegrationSuiteErrorModel;
using LogicAppErrorRequest = Teams.Notifications.Api.Tests.TeamsClient.LogicAppErrorModel;

namespace Teams.Notifications.Api.Tests.IntegrationTests;

[TestClass]
[TestCategory("Integration")]
public sealed class TeamsNotificationApiIntegrationTests
{
    private const string DevTeamName = "DAWN - Integrations Errors Dev";
    private const string TestTeamName = "DAWN - Integrations Errors Test";
    private const string LogicAppChannelName = "Logic App Errors";
    private const string IntegrationSuiteChannelName = "Integration Suite Errors";

    private static CancellationToken _cancellationToken;
    private static TeamsNotificationApi? _client;
    private static string _channelTeamName = string.Empty;

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
            await _client!.IntegrationSuiteErrorPOSTAsync(_channelTeamName, IntegrationSuiteChannelName, model, _cancellationToken);
            await _client!.IntegrationSuiteErrorGetsAnObject(uniqueId, _channelTeamName, IntegrationSuiteChannelName, cancellationToken: _cancellationToken);

            var card = await _client.IntegrationSuiteErrorGETAsync(uniqueId, _channelTeamName, IntegrationSuiteChannelName, _cancellationToken);
            Assert.IsFalse(string.IsNullOrWhiteSpace(card));
            StringAssert.Contains(card, uniqueId);
        }
        finally
        {
            await DeleteIfPresentAsync(() => _client!.IntegrationSuiteErrorDELETEAsync(uniqueId, _channelTeamName, IntegrationSuiteChannelName, CancellationToken.None));
        }

        await _client!.IntegrationSuiteErrorGetsEmpty(uniqueId, _channelTeamName, IntegrationSuiteChannelName, cancellationToken: _cancellationToken);
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
            await _client!.LogicAppErrorPOSTAsync(_channelTeamName, LogicAppChannelName, model, _cancellationToken);
            await _client!.LogicAppErrorGetsAnObject(uniqueId, _channelTeamName, LogicAppChannelName, cancellationToken: _cancellationToken);

            var card = await _client.LogicAppErrorGETAsync(uniqueId, _channelTeamName, LogicAppChannelName, _cancellationToken);
            Assert.IsFalse(string.IsNullOrWhiteSpace(card));
            StringAssert.Contains(card, uniqueId);
        }
        finally
        {
            await DeleteIfPresentAsync(() => _client!.LogicAppErrorDELETEAsync(uniqueId, _channelTeamName, LogicAppChannelName, CancellationToken.None));
        }

        await _client!.LogicAppErrorGetsEmpty(uniqueId, _channelTeamName, LogicAppChannelName, cancellationToken: _cancellationToken);
    }


    private static async Task DeleteIfPresentAsync(Func<Task> deleteAction)
    {
        try
        {
            await deleteAction();
        }
        catch (ApiException ex) when (ex.StatusCode is 400 or 403 or 404) { }
    }
}