namespace Teams.Notifications.AdaptiveCardGen;

[Generator]
public class AdaptiveCardTemplateGenerator : IIncrementalGenerator

{
    public void Initialize(IncrementalGeneratorInitializationContext context)
    {
        // defined in the Teams.Notifications.Api.csproj as additional item's
        var templateAndContent = context
            .AdditionalTextsProvider
            .Where(static file => file.Path.EndsWith(".json", StringComparison.Ordinal) && file.Path.Contains("Templates"))
            .Select(static (file, _) => file);

        // get the content of each item, when you call this method
        context.RegisterSourceOutput(templateAndContent,
            (spc, item) => { CreateFiles(item.Path, item.GetText()!.ToString(), spc); });
    }

    private static void CreateFiles(string path, string content, SourceProductionContext spc)
    {
        var fileName = Path.GetFileNameWithoutExtension(path);
        foreach (var action in GetExecuteActions(content))
        {
            var props = action.DataJson.ExtractPropertiesFromJson();
            if (!props.Any()) continue;
            var actionModelName = $"{fileName}{action.Verb}ActionModel";
            var actionSource = GenerateActionModel(actionModelName, props);
            spc.AddSource($"{actionModelName}.g.cs", SourceText.From(actionSource, Encoding.UTF8));
        }

        var modelProperties = content.GetMustachePropertiesFromString();
        var modelName = $"{fileName}Model";
        var controllerName = $"{fileName}Controller";

        var modelSource = GenerateModel(modelName, modelProperties);
        spc.AddSource($"{modelName}.g.cs", SourceText.From(modelSource, Encoding.UTF8));

        var controllerSource = GenerateController(fileName, spc);
        spc.AddSource($"{controllerName}.g.cs", SourceText.From(controllerSource, Encoding.UTF8));
    }

    private static string GetTypeFromActionModelMustache(KeyValuePair<string, string>? argMustacheProperties)
    {
        var prop = argMustacheProperties?.Value;
        switch (prop)
        {
            case "file":
                return "required string";
            case "file?":
                return "string?";
        }

        // seems double but intellisense doesn't like it otherwise
        if (prop == null || string.IsNullOrWhiteSpace(prop)) return "string?";
        return prop;
    }

    private static string GenerateActionModel(string actionModelName, List<PropWithMustache> props)
    {
        var propertiesOfTheModel = string.Join("\n",
            props.OrderBy(x => x.Property).Select(p => $"        public {MakeRequiredIfNeeded(GetTypeFromActionModelMustache(p.MustacheProperties))} {p.Property} {{ get; set; }}"));
        return
            $$"""
              #nullable enable
              namespace Teams.Notifications.Api.Action.Models;
              public class {{actionModelName}}
              {
              {{propertiesOfTheModel}}
              }
              #nullable disable
              """;
    }

    private static string MakeRequiredIfNeeded(string input)
    {
        return input switch
        {
            "string" => "required string",
            "int" => "required int",
            _ => input
        };
    }

    private static string GenerateModel(string modelName, Dictionary<string, string> props)
    {
        // filter out all the Files, this is a separate controller
        // filter out UniqueId since it is already defined in BaseTemplateModel
        props = props
            .Where(x => x.Value != "file")
            .Where(x => x.Value != "file?")
            .Where(x => !x.Key.Equals("UniqueId", StringComparison.OrdinalIgnoreCase))
            .ToDictionary(x => x.Key, x => x.Value);


        // key is the prop name, value the type, since keys are distinct by nature in Dictionaries
        var propertiesOfTheModel = string.Join("\n",
            props
                .OrderBy(x => x.Value)
                .Select(p => $"        public {MakeRequiredIfNeeded(p.Value)} {p.Key} {{ get; set; }}"));

        return
            $$"""
              #nullable enable
              namespace Teams.Notifications.Api.Models;
              public class {{modelName}} : BaseTemplateModel
              {
              {{propertiesOfTheModel}}
              }
              #nullable disable
              """;
    }

    private static string GenerateController(string name, SourceProductionContext spc)
    {
        var text = ReadResource("CardTemplateController.csgen", spc)
            .Replace("{{name}}", name);
        return text;
    }

    private static string ReadResource(string name, SourceProductionContext spc)
    {
        var assembly = Assembly.GetExecutingAssembly();
        var resourcePath = assembly
            .GetManifestResourceNames()
            .Single(str => str.EndsWith(name, StringComparison.Ordinal));
        using var stream = assembly.GetManifestResourceStream(resourcePath);
        if (stream == null)
        {
            spc.ReportDiagnostic(Diagnostic.Create(
                new(
                    "ACG001",
                    "AdaptiveCard generation file could not be found",
                    "Name: {0}",
                    "AdaptiveCardGen",
                    DiagnosticSeverity.Error,
                    true),
                Location.None,
                name));
            return string.Empty;
        }

        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    private static IEnumerable<ExecuteActionTemplate> GetExecuteActions(string content)
    {
        using var doc = JsonDocument.Parse(content);
        if (!doc.RootElement.TryGetProperty("actions", out var actions) || actions.ValueKind != JsonValueKind.Array) yield break;

        foreach (var action in actions.EnumerateArray())
        {
            if (!action.TryGetProperty("type", out var typeElement) || !string.Equals(typeElement.GetString(), "Action.Execute", StringComparison.Ordinal)) continue;

            if (!action.TryGetProperty("verb", out var verbElement)) continue;

            var verb = verbElement.GetString();
            if (string.IsNullOrWhiteSpace(verb)) continue;

            if (!action.TryGetProperty("data", out var dataElement) || dataElement.ValueKind != JsonValueKind.Object) continue;

            yield return new(verb, Regex.Replace(dataElement.GetRawText(), @"\r\n?|\n", string.Empty));
        }
    }

    private sealed record ExecuteActionTemplate(string Verb, string DataJson);
}