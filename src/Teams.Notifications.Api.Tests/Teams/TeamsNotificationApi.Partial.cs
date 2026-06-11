using Teams.Notifications.Api.Tests.TeamsClient;

namespace Teams.Notifications.Api.Tests.TeamsClient;

public partial class TeamsNotificationApi
{
    partial void Initialize()
    {
        BaseUrl = _httpClient.BaseAddress?.ToString() ?? string.Empty;
        _httpClient.Timeout = TimeSpan.FromMinutes(5);
    }
}