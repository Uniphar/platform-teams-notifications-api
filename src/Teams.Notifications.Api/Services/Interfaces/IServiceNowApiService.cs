namespace Teams.Notifications.Api.Services.Interfaces;

public interface IServiceNowApiService
{
    /// <summary>Resolves the ServiceNow incident associated with <paramref name="uniqueId" />.</summary>
    Task ResolveIncidentAsync(string uniqueId, CancellationToken cancellationToken);
}
