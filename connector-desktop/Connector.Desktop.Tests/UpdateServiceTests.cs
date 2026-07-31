using System.Net;
using System.Security.Cryptography;
using System.Text;
using Connector.Desktop.Services;
using Xunit;

namespace Connector.Desktop.Tests;

public sealed class UpdateServiceTests
{
    [Fact]
    public async Task DownloadInstallerAsync_VerifiesValidSha256()
    {
        var payload = Encoding.UTF8.GetBytes("verified-msi-payload");
        var manifest = BuildManifest(payload, "99.0.1");
        var service = new UpdateService(new HttpClient(new StaticResponseHandler(payload)));
        string? path = null;

        try
        {
            path = await service.DownloadInstallerAsync(manifest, CancellationToken.None);
            Assert.Equal(payload, await File.ReadAllBytesAsync(path));
        }
        finally
        {
            if (path is not null && File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public async Task DownloadInstallerAsync_RejectsAndDeletesInvalidSha256()
    {
        var payload = Encoding.UTF8.GetBytes("tampered-msi-payload");
        var manifest = BuildManifest(payload, "99.0.2");
        manifest.Sha256 = new string('0', 64);
        var service = new UpdateService(new HttpClient(new StaticResponseHandler(payload)));
        var expectedPath = Path.Combine(Path.GetTempPath(), "StructuraConnectorUpdates", "StructuraConnector_99.0.2.msi");

        await Assert.ThrowsAsync<InvalidDataException>(
            () => service.DownloadInstallerAsync(manifest, CancellationToken.None));

        Assert.False(File.Exists(expectedPath));
    }

    [Fact]
    public async Task DownloadInstallerAsync_RequiresHttps()
    {
        var payload = Encoding.UTF8.GetBytes("verified-msi-payload");
        var manifest = BuildManifest(payload, "99.0.3");
        manifest.MsiUrl = "http://example.test/Connector.Desktop.Setup.msi";
        var service = new UpdateService(new HttpClient(new StaticResponseHandler(payload)));

        await Assert.ThrowsAsync<InvalidOperationException>(
            () => service.DownloadInstallerAsync(manifest, CancellationToken.None));
    }

    [Fact]
    public async Task TryGetUpdateAsync_RequiresHttpsManifest()
    {
        var payload = Encoding.UTF8.GetBytes("{}");
        var service = new UpdateService(new HttpClient(new StaticResponseHandler(payload)));

        var manifest = await service.TryGetUpdateAsync(
            "http://example.test/updates/latest.json",
            CancellationToken.None);

        Assert.Null(manifest);
    }

    private static UpdateManifest BuildManifest(byte[] payload, string version)
    {
        return new UpdateManifest
        {
            Version = version,
            MsiUrl = "https://example.test/Connector.Desktop.Setup.msi",
            Sha256 = Convert.ToHexString(SHA256.HashData(payload)).ToLowerInvariant()
        };
    }

    private sealed class StaticResponseHandler(byte[] payload) : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(
            HttpRequestMessage request,
            CancellationToken cancellationToken)
        {
            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new ByteArrayContent(payload)
            });
        }
    }
}
