using Teams.Notifications.Api.Tests.TeamsClient;

namespace Teams.Notifications.Api.Tests.Helper;

internal static class TeamsNotificationApiExtensions
{
    public static Task IntegrationSuiteErrorGetsAnObject(this TeamsNotificationApi client, string uniqueId, string teamName, string channelName, TimeSpan? timeout = default, CancellationToken cancellationToken = default)
    {
        return Condition.WaitUntil(async () =>
            {
                try
                {
                    var result = await client.IntegrationSuiteErrorGETAsync(uniqueId, teamName, channelName, cancellationToken);
                    return !string.IsNullOrWhiteSpace(result);
                }
                catch (ApiException ex) when (ex.StatusCode == 404)
                {
                    return false;
                }
            },
            timeout);
    }

    public static Task IntegrationSuiteErrorGetsEmpty(this TeamsNotificationApi client, string uniqueId, string teamName, string channelName, TimeSpan? timeout = default, CancellationToken cancellationToken = default)
    {
        return Condition.WaitUntil(async () =>
            {
                try
                {
                    var result = await client.IntegrationSuiteErrorGETAsync(uniqueId, teamName, channelName, cancellationToken);
                    return string.IsNullOrWhiteSpace(result);
                }
                catch (ApiException ex) when (ex.StatusCode == 404)
                {
                    return true;
                }
            },
            timeout);
    }

    public static Task LogicAppErrorGetsAnObject(this TeamsNotificationApi client, string uniqueId, string teamName, string channelName, TimeSpan? timeout = default, CancellationToken cancellationToken = default)
    {
        return Condition.WaitUntil(async () =>
            {
                try
                {
                    var result = await client.LogicAppErrorGETAsync(uniqueId, teamName, channelName, cancellationToken);
                    return !string.IsNullOrWhiteSpace(result);
                }
                catch (ApiException ex) when (ex.StatusCode == 404)
                {
                    return false;
                }
            },
            timeout);
    }

    public static Task LogicAppErrorGetsEmpty(this TeamsNotificationApi client, string uniqueId, string teamName, string channelName, TimeSpan? timeout = default, CancellationToken cancellationToken = default)
    {
        return Condition.WaitUntil(async () =>
            {
                try
                {
                    var result = await client.LogicAppErrorGETAsync(uniqueId, teamName, channelName, cancellationToken);
                    return string.IsNullOrWhiteSpace(result);
                }
                catch (ApiException ex) when (ex.StatusCode == 404)
                {
                    return true;
                }
            },
            timeout);
    }
}