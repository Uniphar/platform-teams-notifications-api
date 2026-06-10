namespace Teams.Notifications.Api.Services;

public record PropHelperItem
{
    public required PropHelperItemFile File;
    public required string Property;
    public required string Type;
}

public record PropHelperItemFile
{
    public required string Location;
    public required string Name;
    public required string Url;
}

public static class PropertyHelper
{
    private static readonly Regex MustacheRegex = new("{{(?<name>.*?):(?<type>.*?)}}",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

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

    private static IEnumerable<TSource> DistinctByProps<TSource, TKey>(this IEnumerable<TSource> source, Func<TSource, TKey> keySelector)
    {
        HashSet<TKey> seenKeys = [];
        foreach (var element in source)
        {
            if (seenKeys.Add(keySelector(element))) yield return element;
        }
    }

    /// <summary>
    ///     checks if the list has any file template, which is either FileUrl or FileName or FileLocation
    /// </summary>
    /// <param name="nameAndType"></param>
    /// <returns></returns>
    public static bool HasFileTemplate(this Dictionary<string, string> nameAndType)
    {
        return nameAndType.Any(x => x is { Value: "file" or "file?", Key: "FileUrl" or "FileName" or "FileLocation" });
    }

    private static string? ToJsonString(this string? value)
    {
        if (value == null || string.IsNullOrWhiteSpace(value)) return value;
        return JsonEncodedText.Encode(value).Value;
    }

    public static string FindPropAndReplace<T>(this string jsonString, T model, PropHelperItem item)
    {
        var toReplace = "{{" + item.Property + ":" + item.Type + "}}";
        switch (item.Type)
        {
            // optional string, will remove the block if empty
            case "string?":
                var valueString = model.TryGetStringPropertyValue(item.Property).ToJsonString();
                if (!string.IsNullOrEmpty(valueString)) return jsonString.Replace(toReplace, valueString);
                // Parse JSON and remove objects from arrays where the property value matches the placeholder
                var rootString = JsonNode.Parse(jsonString);
                rootString = RemoveObjectsWithPlaceholder(rootString, toReplace);
                return rootString?.ToJsonString(new() { WriteIndented = false }) ?? jsonString;

            // required string
            case "string":
                return jsonString.Replace(toReplace, model.TryGetStringPropertyValue(item.Property).ToJsonString());
            case "int":
                return jsonString.Replace(toReplace, model.TryGetIntPropertyValue(item.Property)?.ToString() ?? string.Empty);
            case "file":
            case "file?":
                return jsonString.ReplaceForFile(toReplace, item.File);
            default:
                return jsonString;
        }
    }

    // Recursively remove objects from arrays where the property value matches the placeholder
    private static JsonNode? RemoveObjectsWithPlaceholder(JsonNode? node, string toReplace)
    {
        switch (node)
        {
            case JsonArray array:
            {
                for (var i = array.Count - 1; i >= 0; i--)
                {
                    if (array[i] is JsonObject obj && ObjectContainsPlaceholder(obj, toReplace))
                        array.RemoveAt(i);
                    else
                        RemoveObjectsWithPlaceholder(array[i], toReplace);
                }

                break;
            }
            case JsonObject obj:
            {
                foreach (var prop in obj) RemoveObjectsWithPlaceholder(prop.Value, toReplace);
                break;
            }
        }

        return node;
    }

    // Helper: Recursively checks if any value in the object matches
    private static bool ObjectContainsPlaceholder(JsonObject obj, string toReplace)
    {
        foreach (var prop in obj)
        {
            switch (prop.Value)
            {
                case JsonValue value when value.ToString().Contains(toReplace):
                case JsonObject childObj when ObjectContainsPlaceholder(childObj, toReplace):
                    return true;
                case JsonArray arr:
                {
                    foreach (var item in arr)
                    {
                        switch (item)
                        {
                            case JsonObject arrObj when ObjectContainsPlaceholder(arrObj, toReplace):
                            case JsonValue arrVal when arrVal.ToString().Contains(toReplace):
                                return true;
                        }
                    }

                    break;
                }
            }
        }

        return false;
    }

    private static string ReplaceForFile(this string content, string toReplace, PropHelperItemFile file)
    {
        // if we don't have a file, we need to remove it anyway
        if (string.IsNullOrWhiteSpace(file.Url) || string.IsNullOrWhiteSpace(file.Location) || string.IsNullOrWhiteSpace(file.Name))
        {
            var root = JsonNode.Parse(content);
            root = RemoveObjectsWithPlaceholder(root, toReplace);
            return root?.ToJsonString(new() { WriteIndented = false }) ?? content;
        }

        var toReplaceWith = toReplace switch
        {
            "{{FileUrl:file}}" or "{{FileUrl:file?}}" => file.Url,
            "{{FileLocation:file}}" or "{{FileLocation:file?}}" => file.Location,
            "{{FileName:file}}" or "{{FileName:file?}}" => file.Name,
            _ => string.Empty
        };
        content = content.Replace(toReplace, toReplaceWith);
        return content;
    }

    private static int? TryGetIntPropertyValue<T>(this T model, string propertyName)
    {
        var property = typeof(T).GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
        if (property == null || property.PropertyType != typeof(int)) return null;
        return (int?)(property.GetValue(model) ?? null);
    }

    private static string? TryGetStringPropertyValue<T>(this T model, string propertyName)
    {
        var property = typeof(T).GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
        if (property == null || property.PropertyType != typeof(string)) return null;

        return property.GetValue(model) as string;
    }

    public static IFormFile? GetFileValue<T>(this T model)
    {
        var property = typeof(T).GetProperty("File", BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
        if (property == null || property.PropertyType != typeof(IFormFile)) return null;

        return property.GetValue(model) as IFormFile;
    }
}