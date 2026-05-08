namespace Teams.Notifications.Formatter.Util;

internal static class GitHubActions
{
    public static bool IsCI { get; } = Environment.GetEnvironmentVariable("CI") is not null;

    public static void Error(string title, string message, string? file = null, SourceRange start = default, SourceRange end = default)
    {
        if (!IsCI) return;

        ArgumentException.ThrowIfNullOrWhiteSpace(title);
        ArgumentException.ThrowIfNullOrWhiteSpace(message);

        var filePart = !string.IsNullOrWhiteSpace(file)
            ? $"file={Git.GetRepoRelativePath(file)},"
            : null;

        AnsiConsole.WriteLine($"::error {filePart}{start}{end}title={title}::{message}");
    }
}

internal readonly struct SourceRange
{
    public int Line { get; }
    public int Column { get; }

    public SourceRange(int line) : this(line, 1) { }

    public SourceRange(int line, int col)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(line);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(col);
        Line = line;
        Column = col;
    }

    public static implicit operator SourceRange(int line) => new(line);
    public static implicit operator SourceRange(long line) => new((int)line);
    public static implicit operator SourceRange(int? line) => line.HasValue ? line.Value : default;
    public static implicit operator SourceRange(long? line) => line.HasValue ? line.Value : default;
    public static implicit operator SourceRange((int line, int col) range) => new(range.line, range.col);

    public override string? ToString() =>
        this switch
        {
            { Line: > 0, Column: > 0 } => $"line={Line},col={Column},",
            { Line: > 0 } => $"line={Line},",
            _ => null
        };
}

file static class Git
{
    public static string GetRepoRoot(string path = ".")
    {
        for (var currentPath = Path.GetFullPath(path); !string.IsNullOrWhiteSpace(currentPath); currentPath = Path.GetDirectoryName(currentPath))
        {
            if (Directory.Exists(Path.Combine(currentPath, ".git"))) return currentPath;
        }

        throw new InvalidOperationException("Not a path in a git repo");
    }

    public static string GetRepoRelativePath(string path)
    {
        var gitRepoRoot = GetRepoRoot(Path.GetDirectoryName(path)!);
        return Path.GetRelativePath(gitRepoRoot, path);
    }
}