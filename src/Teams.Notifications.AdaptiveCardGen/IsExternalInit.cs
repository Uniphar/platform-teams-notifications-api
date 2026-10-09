// Polyfill required for C# 9+ records/init-only setters when targeting netstandard2.0.

#pragma warning disable IDE0130
namespace System.Runtime.CompilerServices;
#pragma warning restore IDE0130

// ReSharper disable once UnusedType.Global
internal static class IsExternalInit;