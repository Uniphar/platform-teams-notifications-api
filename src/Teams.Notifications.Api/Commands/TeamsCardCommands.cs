namespace Teams.Notifications.Api.Commands;

/// <summary>Published when a Teams channel card is created for the first time.</summary>
public sealed record TeamsCardCreatedCommand
{
    public required string UniqueId { get; init; }
    public required string TeamsDeepLink { get; init; }
    public required string CardType { get; init; }
    public required bool HasActionableFile { get; init; }
    public required string TenantId { get; init; }
}

/// <summary>Published when an existing Teams channel card is updated (e.g. a file was attached).</summary>
public sealed record TeamsCardUpdatedCommand
{
    public required string UniqueId { get; init; }
    public required string TeamsDeepLink { get; init; }
    public required string CardType { get; init; }
    public required bool HasActionableFile { get; init; }
    public required string TenantId { get; init; }
}

/// <summary>Published when a Teams channel card is deleted.</summary>
public sealed record TeamsCardDeletedCommand
{
    public required string UniqueId { get; init; }
    public required string TenantId { get; init; }
}

/// <summary>Published when a user completes a card action (e.g. reprocessed a file).</summary>
public sealed record TeamsCardResolvedCommand
{
    public required string UniqueId { get; init; }
    public required string TenantId { get; init; }
}
