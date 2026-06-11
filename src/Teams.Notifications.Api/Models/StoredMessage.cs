namespace Teams.Notifications.Api.Models;

public sealed record StoredMessage
{
    /// <summary>Graph message id. Document id within the partition.</summary>
    [JsonPropertyName("id")]
    public string Id { get; init; } = string.Empty;

    /// <summary>Partition key: <c>chat:{chatId}</c> or <c>channel:{teamId}:{channelId}</c>.</summary>
    [JsonPropertyName("pk")]
    public required string PartitionKey { get; init; }

    [JsonPropertyName("uniqueId")]
    public required string UniqueId { get; init; }

    [JsonPropertyName("jsonFileName")]
    public required string JsonFileName { get; init; }

    [JsonPropertyName("chatId")]
    public string? ChatId { get; init; }

    [JsonPropertyName("teamId")]
    public string? TeamId { get; init; }

    [JsonPropertyName("channelId")]
    public string? ChannelId { get; init; }

    [JsonPropertyName("cardJson")]
    public required string CardJson { get; init; }

    [JsonPropertyName("createdAt")]
    public DateTimeOffset CreatedAt { get; init; }

    [JsonPropertyName("updatedAt")]
    public DateTimeOffset UpdatedAt { get; init; }

    public static string ChatPartition(string chatId) => $"chat:{chatId}";
    public static string ChannelPartition(string teamId, string channelId) => $"channel:{teamId}:{channelId}";
}