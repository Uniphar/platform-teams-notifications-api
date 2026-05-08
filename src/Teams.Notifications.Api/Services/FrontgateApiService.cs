namespace Teams.Notifications.Api.Services;

public class FrontgateApiService(IHttpClientFactory factory, IConfiguration configuration) : IFrontgateApiService
{
    private readonly HttpClient _httpClient = factory.CreateClient("frontgate-api");

    private readonly TokenCredential credential = new DefaultAzureCredential(new DefaultAzureCredentialOptions
    {
        TenantId = Environment.GetEnvironmentVariable("AZURE_ENTRA_EXTERNAL_TENANT_ID")
    });

    private readonly string frontgateApiScope = $"api://frontgate-api/{configuration["frontgate-api-client-id"]}/.default";


    public async Task<HttpResponseMessage> UploadFileAsync(string originalBlobUrl, LogicAppFrontgateFileInformation fileInfo, CancellationToken cancellationTokentoken)
    {
        var token = await credential.GetTokenAsync(new([frontgateApiScope]), cancellationTokentoken);
        _httpClient.DefaultRequestHeaders.Authorization = new("Bearer", token.Token);
        return await _httpClient.PostAsJsonAsync("/frontgate/Reprocess/reprocess-file-logic-app?originalBlobUrl=" + originalBlobUrl, fileInfo, cancellationTokentoken);
    }
}