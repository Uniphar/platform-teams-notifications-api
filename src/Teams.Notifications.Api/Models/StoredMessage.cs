using Newtonsoft.Json;

namespace Teams.Notifications.Api.Models;

public sealed class StoredMessage
{
    /// <summary>Graph message id. Document id within the partition.</summary>
    [JsonProperty("id")]
    public string Id { get; set; } = string.Empty;

    /// <summary>Partition key: <c>chat:{chatId}</c> or <c>channel:{teamId}:{channelId}</c>.</summary>
    [JsonProperty("pk")]
    public string PartitionKey { get; set; } = string.Empty;

    [JsonProperty("uniqueId")]
    public string UniqueId { get; set; } = string.Empty;

    [JsonProperty("jsonFileName")]
    public string JsonFileName { get; set; } = string.Empty;

    [JsonProperty("chatId", NullValueHandling = NullValueHandling.Ignore)]
    public string? ChatId { get; set; }

    [JsonProperty("teamId", NullValueHandling = NullValueHandling.Ignore)]
    public string? TeamId { get; set; }

    [JsonProperty("channelId", NullValueHandling = NullValueHandling.Ignore)]
    public string? ChannelId { get; set; }

    [JsonProperty("cardJson")]
    public string CardJson { get; set; } = string.Empty;

    [JsonProperty("createdAt")]
    public DateTimeOffset CreatedAt { get; set; }

    [JsonProperty("updatedAt")]
    public DateTimeOffset UpdatedAt { get; set; }

    public static string ChatPartition(string chatId) => $"chat:{chatId}";
    public static string ChannelPartition(string teamId, string channelId) => $"channel:{teamId}:{channelId}";
}