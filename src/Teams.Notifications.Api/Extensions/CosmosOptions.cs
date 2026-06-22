namespace Teams.Notifications.Api.Extensions;

public sealed class CosmosOptions
{
    public const string SectionName = "CosmosStore";

    public required string ConnectionString { get; set; }
    public string DatabaseName { get; set; } = "teams-notifications";
    public string ContainerName { get; set; } = "messages";
    public int RetentionDays { get; set; } = 30;
}