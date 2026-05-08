namespace Teams.Notifications.AdaptiveCardGen;

internal static class PropertyHelper
{
    private static readonly Regex MustacheRegex = new("{{(?<name>.*?):(?<type>.*?)}}",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    /// <summary>
    ///     Gives back the props, eg { "Title": "Bla"} will return "Title"
    /// </summary>
    /// <param name="json">compliant json</param>
    /// <returns>List of the props</returns>
    public static List<PropWithMustache> ExtractPropertiesFromJson(this string json)
    {
        using var doc = JsonDocument.Parse(json);

        return (from property in doc.RootElement.EnumerateObject()
            let value = property.Value.GetString()
            where !string.IsNullOrWhiteSpace(value)
            select new PropWithMustache { Property = property.Name, MustacheProperties = value.GetMustachePropertiesFromString().FirstOrDefault() }).ToList();
    }

    /// <summary>
    ///     Very simple regex to go from {{name:type}} to a list of properties with their types
    /// </summary>
    /// <param name="content"></param>
    /// <returns>Distinct list of all properties in the string</returns>
    public static Dictionary<string, string> GetMustachePropertiesFromString(this string content)
    {
        // 
        var matches = MustacheRegex.Matches(content);
        var properties = matches
            .Select(x => new { name = x.Groups["name"].Value, type = x.Groups["type"].Value })
            .DistinctByProps(x => x.name);
        return properties.ToDictionary(m => m.name, m => m.type);
    }
}

public record PropWithMustache
{
    public string? Property { get; set; }
    public KeyValuePair<string, string>? MustacheProperties { get; set; }
}