using System.Text.RegularExpressions;

namespace Teams.Notifications.Formatter.Util;

internal static class PropertyHelper
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
    ///     checks if the types are valid, atm int, string or file
    /// </summary>
    /// <param name="nameAndType"> types you want to check</param>
    /// <returns>True if no mismatches were found</returns>
    public static Tuple<bool, Dictionary<string, string>> IsValidTypes(this Dictionary<string, string> nameAndType)
    {
        //name is key, type is value, due to dict
        var wrongItems = nameAndType
            .Where(x => x.Value is not
                ("int" or "string" or "string?" or "file" or "file?")
            )
            .ToDictionary(x => x.Key, x => x.Value);

        return new(!wrongItems.Any(), wrongItems);
    }

    /// <summary>
    ///     Files are uniquely named, this checks that
    /// </summary>
    /// <param name="nameAndType">Full list of props</param>
    /// <returns>true if the files props are correct</returns>
    public static Tuple<bool, Dictionary<string, string>> IsValidFile(this Dictionary<string, string> nameAndType)
    {
        var wrongItems = nameAndType.Where(x => x is { Value: "file" or "file?", Key: not ("FileUrl" or "FileName" or "FileLocation") }).ToDictionary(x => x.Key, x => x.Value);
        return new(!wrongItems.Any(), wrongItems);
    }
}