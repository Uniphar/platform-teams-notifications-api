namespace Teams.Notifications.Api.Models;

public sealed class StoredMessage
{
    /// <summary>Graph message id. Document id within the partition.</summary>
    [JsonPropertyName("id")]
    public string Id { get; set; } = string.Empty;

    /// <summary>Partition key: <c>chat:{chatId}</c> or <c>channel:{teamId}:{channelId}</c>.</summary>
    [JsonPropertyName("pk")]
    public string PartitionKey { get; set; } = string.Empty;

    [JsonPropertyName("uniqueId")]
    public string UniqueId { get; set; } = string.Empty;

    [JsonPropertyName("jsonFileName")]
    public string JsonFileName { get; set; } = string.Empty;

    [JsonPropertyName("chatId")]
    public string? ChatId { get; set; }

    [JsonPropertyName("teamId")]
    public string? TeamId { get; set; }

    [JsonPropertyName("channelId")]
    public string? ChannelId { get; set; }

    [JsonPropertyName("cardJson")]
    public string CardJson { get; set; } = string.Empty;

    [JsonPropertyName("createdAt")]
    public DateTimeOffset CreatedAt { get; set; }

    [JsonPropertyName("updatedAt")]
    public DateTimeOffset UpdatedAt { get; set; }

    public static string ChatPartition(string chatId) => $"chat:{chatId}";
    public static string ChannelPartition(string teamId, string channelId) => $"channel:{teamId}:{channelId}";
}