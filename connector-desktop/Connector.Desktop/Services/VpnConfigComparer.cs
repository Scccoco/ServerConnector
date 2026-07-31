using System.IO;

namespace Connector.Desktop.Services;

public static class VpnConfigComparer
{
    public static string Normalize(string value)
    {
        return (value ?? string.Empty)
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n')
            .Trim();
    }

    public static bool FileMatches(string path, string expected)
    {
        try
        {
            return File.Exists(path) &&
                   string.Equals(
                       Normalize(File.ReadAllText(path)),
                       Normalize(expected),
                       StringComparison.Ordinal);
        }
        catch
        {
            return false;
        }
    }
}
