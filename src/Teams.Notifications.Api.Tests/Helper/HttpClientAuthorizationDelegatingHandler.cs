namespace Teams.Notifications.Api.Tests.Helper;

internal sealed class HttpClientAuthorizationDelegatingHandler : DelegatingHandler
{
    private readonly TokenCredential _credential;
    private readonly string _scope;

    public HttpClientAuthorizationDelegatingHandler(TokenCredential credential, IConfiguration configuration)
    {
        _credential = credential;
        var apiClientId = configuration["platform-teams-notification-api-client-id"]
            ?? throw new InvalidOperationException("platform-teams-notification-api-client-id is required");
        _scope = $"api://platform-teams-notification-api/{apiClientId}/.default";
    }

    protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
    {
        var token = await _credential.GetTokenAsync(new([_scope]), cancellationToken);
        request.Headers.Authorization = new("Bearer", token.Token);
        return await base.SendAsync(request, cancellationToken);
    }
}