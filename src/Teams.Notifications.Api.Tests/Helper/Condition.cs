namespace Teams.Notifications.Api.Tests.Helper;

internal static class Condition
{
    public static async Task WaitUntil(Func<Task<bool>> condition, TimeSpan? timeout, [CallerArgumentExpression(nameof(condition))] string? conditionText = default)
    {
        var delay = TimeSpan.FromSeconds(5);
        timeout ??= TimeSpan.FromMinutes(2);

        var stopwatch = Stopwatch.StartNew();
        while (stopwatch.Elapsed < timeout && !await condition()) await Task.Delay(delay);

        if (stopwatch.Elapsed < timeout) return;

        if (conditionText is null) throw new TimeoutException("Condition failed.");

        var normalizedCondition = conditionText.Trim();
        normalizedCondition = normalizedCondition.StartsWith("async () =>", StringComparison.Ordinal)
            ? normalizedCondition[11..].Trim()
            : normalizedCondition;
        normalizedCondition = normalizedCondition.StartsWith("await", StringComparison.Ordinal)
            ? normalizedCondition[5..].Trim()
            : normalizedCondition;

        throw new TimeoutException($"Condition failed: {normalizedCondition}");
    }
}