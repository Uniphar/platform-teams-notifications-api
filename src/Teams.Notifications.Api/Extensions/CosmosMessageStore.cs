namespace Teams.Notifications.Api.Extensions;

public sealed class CosmosMessageStore
{
    public const string SectionName = "CosmosMessageStore";

    public string Endpoint { get; set; } = string.Empty;
    public string DatabaseName { get; set; } = "teams-notifications";
    public string ContainerName { get; set; } = "messages";
    public int RetentionDays { get; set; } = 30;
}