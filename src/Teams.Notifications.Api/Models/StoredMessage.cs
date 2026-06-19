namespace Teams.Notifications.Api.Models;

public sealed record StoredMessage
{
    /// <summary>Cosmos document id (unique notification id).</summary>
    [JsonPropertyName("id")]
    public required string Id { get; init; }

    /// <summary>Graph message id returned by Teams after sending.</summary>
    [JsonPropertyName("messageId")]
    public string? MessageId { get; init; }

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
}