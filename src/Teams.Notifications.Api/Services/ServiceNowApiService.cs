namespace Teams.Notifications.Api.Services;

public sealed class ServiceNowApiService(IHttpClientFactory factory, IConfiguration configuration) : IServiceNowApiService
{
    private readonly TokenCredential _credential = new DefaultAzureCredential(new DefaultAzureCredentialOptions
    {
        TenantId = Environment.GetEnvironmentVariable("AZURE_ENTRA_EXTERNAL_TENANT_ID")
    });

    private readonly HttpClient _httpClient = factory.CreateClient("service-now-api");

    private readonly string _serviceNowApiScope = $"api://service-now-api/{configuration["service-now-api-client-id"]}/.default";

    public async Task ResolveIncidentAsync(string uniqueId, CancellationToken cancellationToken)
    {
        var token = await _credential.GetTokenAsync(new([_serviceNowApiScope]), cancellationToken);
        _httpClient.DefaultRequestHeaders.Authorization = new("Bearer", token.Token);
        var encodedId = Uri.EscapeDataString(uniqueId);
        var response = await _httpClient.PostAsync($"/service-now-api/LogicAppError/resolve?uniqueId={encodedId}", null, cancellationToken);
        response.EnsureSuccessStatusCode();
    }
}