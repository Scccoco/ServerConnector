using Connector.Desktop.Services;
using Xunit;

namespace Connector.Desktop.Tests;

public sealed class VpnConfigComparerTests
{
    [Fact]
    public void Normalize_IgnoresLineEndingAndOuterWhitespaceDifferences()
    {
        const string windows = "\r\n[Interface]\r\nAddress = 10.77.123.2/32\r\n";
        const string unix = "[Interface]\nAddress = 10.77.123.2/32\n\n";

        Assert.Equal(
            VpnConfigComparer.Normalize(windows),
            VpnConfigComparer.Normalize(unix));
    }

    [Fact]
    public void FileMatches_ReturnsFalseForDifferentAllowedIps()
    {
        var path = Path.GetTempFileName();
        try
        {
            File.WriteAllText(path, "AllowedIPs = 10.77.123.0/24");

            Assert.False(
                VpnConfigComparer.FileMatches(
                    path,
                    "AllowedIPs = 10.77.123.0/24, 62.113.36.107/32"));
        }
        finally
        {
            File.Delete(path);
        }
    }
}
