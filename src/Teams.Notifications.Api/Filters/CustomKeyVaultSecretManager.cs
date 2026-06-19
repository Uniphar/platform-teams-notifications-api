namespace Teams.Notifications.Api.Filters;

public sealed class CustomKeyVaultSecretManager : KeyVaultSecretManager
{
    private readonly HashSet<string> _secretNames;
    private readonly Dictionary<string, string> _keyOverrides;

    public CustomKeyVaultSecretManager(params string[] secretNames)
        : this(secretNames, keyOverrides: null)
    {
    }

    public CustomKeyVaultSecretManager(
        IEnumerable<string> secretNames,
        IReadOnlyDictionary<string, string>? keyOverrides)
    {
        if (secretNames is null)
            throw new ArgumentNullException(nameof(secretNames));

        _secretNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var secretName in secretNames)
        {
            if (string.IsNullOrWhiteSpace(secretName))
                throw new ArgumentException("Secret names cannot contain null or empty values.", nameof(secretNames));

            _secretNames.Add(secretName);
        }

        _keyOverrides = keyOverrides is null
            ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            : new Dictionary<string, string>(keyOverrides, StringComparer.OrdinalIgnoreCase);

        foreach (var secretName in _keyOverrides.Keys)
        {
            if (string.IsNullOrWhiteSpace(secretName))
                throw new ArgumentException("Secret override names cannot contain null or empty values.", nameof(keyOverrides));

            _secretNames.Add(secretName);
        }

        if (_secretNames.Count == 0)
            throw new ArgumentException("At least one secret name must be provided.", nameof(secretNames));
    }

    public override bool Load(SecretProperties secret)
    {
        return _secretNames.Any(name =>
            secret.Name.Equals(name, StringComparison.OrdinalIgnoreCase)
            || secret.Name.StartsWith($"{name}-", StringComparison.OrdinalIgnoreCase));
    }

    public override string GetKey(KeyVaultSecret secret)
    {
        if (_keyOverrides.TryGetValue(secret.Name, out var key))
            return key;

        return secret.Name.Replace("--", ConfigurationPath.KeyDelimiter, StringComparison.Ordinal);
    }
}