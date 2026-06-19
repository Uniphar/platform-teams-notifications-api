namespace Teams.Notifications.Api.Services.Interfaces;

public interface ITeamsCardEventPublisher
{
    Task PublishCardCreatedAsync(string uniqueId, string teamsDeepLink, string cardType, bool hasActionableFile, CancellationToken token);
    Task PublishCardUpdatedAsync(string uniqueId, string teamsDeepLink, string cardType, bool hasActionableFile, CancellationToken token);
    Task PublishCardDeletedAsync(string uniqueId, CancellationToken token);
    Task PublishCardResolvedAsync(string uniqueId, CancellationToken token);
}
