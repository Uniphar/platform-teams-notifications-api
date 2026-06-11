using Teams.Notifications.Api.Tests.TeamsClient;
using IntegrationSuiteErrorRequest = Teams.Notifications.Api.Tests.TeamsClient.IntegrationSuiteErrorModel;
using LogicAppErrorRequest = Teams.Notifications.Api.Tests.TeamsClient.LogicAppErrorModel;

namespace Teams.Notifications.Api.Tests.IntegrationTests;

[TestClass]
[TestCategory("Integration")]
[TestCategory("Smoke")]
public sealed class TeamsNotificationApiIntegrationTests
{
    private const string DevTeamName = "DAWN - Integrations Errors Dev";
    private const string TestTeamName = "DAWN - Integrations Errors Test";
    private const string LogicAppChannelName = "Logic App Errors";
    private const string IntegrationSuiteChannelName = "Integration Suite Errors";

    private static CancellationToken _cancellationToken;
    private static TeamsNotificationApi? _client;
    private static string _channelTeamName = string.Empty;
    private static string _configuredEnvironment = string.Empty;
    private static string? _skipReason;

    [ClassInitialize]
    public static async Task ClassInitialize(TestContext context)
    {
        _cancellationToken = context.CancellationToken;
        _configuredEnvironment = (context.Properties["Environment"]?.ToString() ?? "dev").Trim().ToLowerInvariant();

        if (_configuredEnvironment == "prod")
        {
            _skipReason = "These end-to-end card mutation tests are intended for dev and test only.";
            return;
        }

        try
        {
            var keyVaultCredential = CreateKeyVaultCredential();
            var apiCredential = CreateApiCredential(context);
            var configuration = await BuildConfigurationAsync(_configuredEnvironment, keyVaultCredential);
            var handler = new HttpClientAuthorizationDelegatingHandler(apiCredential, configuration)
            {
                InnerHandler = new HttpClientHandler()
            };

            var httpClient = new HttpClient(handler)
            {
                BaseAddress = new(GetApiBaseUrl(_configuredEnvironment))
            };

            _client = new(httpClient);
            _channelTeamName = _configuredEnvironment == "test" ? TestTeamName : DevTeamName;
        }
        catch (CredentialUnavailableException ex)
        {
            _skipReason = $"Azure credentials unavailable for end-to-end tests: {ex.Message}";
        }
        catch (AuthenticationFailedException ex)
        {
            _skipReason = $"Azure authentication failed for end-to-end tests: {ex.Message}";
        }
    }

    [TestMethod]
    public async Task IntegrationSuiteError_CanBeAddedAndRemoved_EndToEnd()
    {
        SkipIfNeeded();

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
        SkipIfNeeded();

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

    private static async Task<IConfiguration> BuildConfigurationAsync(string environment, TokenCredential credential)
    {
        var apiProjectPath = Path.GetFullPath("../../../../Teams.Notifications.Api");
        var builder = new ConfigurationBuilder()
            .SetBasePath(apiProjectPath)
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: false)
            .AddJsonFile($"appsettings.{environment}.json", optional: true, reloadOnChange: false);

        if (environment == "local")
        {
            builder.AddJsonFile("appsettings.local.json", optional: false, reloadOnChange: false);
        }
        else
        {
            builder.AddAzureKeyVault(new($"https://uni-devops-app-{environment}-kv.vault.azure.net/"), credential);
        }

        return await Task.FromResult(builder.Build());
    }

    private static TokenCredential CreateKeyVaultCredential() => new DefaultAzureCredential();

    private static TokenCredential CreateApiCredential(TestContext context)
    {
        if (_configuredEnvironment == "local")
        {
            var tenantId = context.Properties["TenantId"]?.ToString() ?? throw new InvalidOperationException("TenantId is required for local test runs.");
            var clientId = context.Properties["ClientId"]?.ToString() ?? throw new InvalidOperationException("ClientId is required for local test runs.");
            var clientSecret = context.Properties["ClientSecret"]?.ToString() ?? throw new InvalidOperationException("ClientSecret is required for local test runs.");
            return new ClientSecretCredential(tenantId, clientId, clientSecret);
        }

        var externalTenantId = context.Properties["AZURE_ENTRA_EXTERNAL_TENANT_ID"]?.ToString();
        return string.IsNullOrWhiteSpace(externalTenantId)
            ? new DefaultAzureCredential()
            : new DefaultAzureCredential(new DefaultAzureCredentialOptions { TenantId = externalTenantId });
    }

    private static string GetApiBaseUrl(string environment) => environment switch
    {
        "prod" => "https://api.uniphar.ie/",
        _ => $"https://api.{environment}.uniphar.ie/"
    };

    private static async Task DeleteIfPresentAsync(Func<Task> deleteAction)
    {
        try
        {
            await deleteAction();
        }
        catch (ApiException ex) when (ex.StatusCode == 400 || ex.StatusCode == 403 || ex.StatusCode == 404)
        {
        }
    }

    private static void SkipIfNeeded()
    {
        if (!string.IsNullOrWhiteSpace(_skipReason))
        {
            Assert.Inconclusive(_skipReason);
        }
    }
}