// Polyfill required for C# 9+ records/init-only setters when targeting netstandard2.0.
namespace System.Runtime.CompilerServices
{
    internal static class IsExternalInit { }
}
