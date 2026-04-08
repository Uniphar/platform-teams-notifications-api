using AdaptiveCards;
using Teams.Notifications.Formatter.Util;

namespace Teams.Notifications.Formatter;

internal sealed class FormatCommand : Command<FormatCommand.Settings>
{
    protected override int Execute(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        var differ = new FilesDiffer(Directory.GetCurrentDirectory());
        differ.AddAllUnderPath("./../Teams.Notifications.Api/Templates", "*.json", FormatFile);

        if (!settings.Check) return !differ.Apply() ? throw new("Failed to apply changes to config files.") : 0;

        if (differ.Check("Formatting", "File differs after formatting")) return 0;

        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLineInterpolated($"[red]One of more config files were not formatted before commiting.[/] Run `dotnet run format` in the `{typeof(FormatCommand).Assembly.GetName().Name}` project directory.");
        GitHubActions.Error("Formatting", $"One of more config files were not formatted before commiting. Run `dotnet run format` in the `{typeof(FormatCommand).Assembly.GetName().Name}` project directory, and commit the updated config files.");
        throw new("Failed to apply changes to config files.");
    }

    private static void FormatFile(string sourcePath, Stream formattedFile)
    {
        var text = File.ReadAllText(sourcePath);
        var props = text.GetMustachePropertiesFromString();
        var (validTypes, wrongItemsTypes) = props.IsValidTypes();
        if (!validTypes)
        {
            var file = Path.GetFileName(sourcePath);
            AnsiConsole.MarkupLineInterpolated($"[bold red]The following file has incompatible properties[/] [bold white]{file}[/] ");
            var table = new Table();
            table.AddColumn(new("[green]Template[/]"));
            table.AddColumn(new(new Markup("[yellow]Type[/]")));
            table.AddColumn(new("[blue]Property name[/]"));
            wrongItemsTypes.ToList().ForEach(x => table.AddRow("[bold green]{{" + x.Key + ":" + x.Value + "}}[/]", $"[yellow]{x.Value}[/]", $"[blue]{x.Key}[/]"));
            AnsiConsole.Write(table);
            GitHubActions.Error("Formatting", $"One of the files has incompatible properties, check the following file: {file} for property: {string.Join(",", wrongItemsTypes.Keys)}, unrecognised type(s) {string.Join(",", wrongItemsTypes.Values)}");
            throw new InvalidDataException($"Unrecognised types {string.Join(",", wrongItemsTypes.Values)}");
        }

        var (validFile, wrongItemsFile) = props.IsValidFile();
        if (!validFile)
        {
            var file = Path.GetFileName(sourcePath);
            AnsiConsole.MarkupLineInterpolated($"[bold red]The following file has a file-url or file-name but not the File as property name[/] [bold white]{file}[/]");
            AnsiConsole.MarkupLine("Only [bold white]{{FileName:file}}[/] or/and [bold white]{{FileUrl:file}}[/] or/and [bold white]{{FileLocation:file}}[/] , which will create a IFormFile File entry to upload to");

            GitHubActions.Error("Formatting", $"One of the files has incompatible properties, check the following file: {file} for property: {string.Join(",", wrongItemsFile.Keys)}, unrecognised type(s) {string.Join(",", wrongItemsFile.Values)}");
            throw new InvalidDataException($"Unrecognised types {string.Join(",", wrongItemsFile.Values)}");
        }

        var item = AdaptiveCard.FromJson(text).Card;
        var formatted = item.ToJson() ?? string.Empty;

        var sw = new StreamWriter(formattedFile);
        sw.Write(formatted);
        sw.Flush(); //otherwise you are risking empty stream
    }

    public sealed class Settings : CommandSettings
    {
        [CommandOption("--check")]
        public bool Check { get; init; }
    }
}