namespace Teams.Notifications.AdaptiveCardGen;

public static class IEnumerableExtensions
{
    public static IEnumerable<TSource> DistinctByProps<TSource, TKey>(this IEnumerable<TSource> source, Func<TSource, TKey> keySelector)
    {
        HashSet<TKey> seenKeys = [];
        foreach (var element in source)
        {
            if (seenKeys.Add(keySelector(element))) yield return element;
        }
    }
}