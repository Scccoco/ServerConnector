using System.Net;
using System.Net.Sockets;
using Connector.Desktop.Services;
using Xunit;

namespace Connector.Desktop.Tests;

public sealed class TcpConnectivityProbeTests
{
    [Fact]
    public async Task CanConnectAsync_ReturnsTrue_WhenListenerAcceptsConnection()
    {
        using var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        var port = ((IPEndPoint)listener.LocalEndpoint).Port;
        var acceptTask = listener.AcceptTcpClientAsync();
        var probe = new TcpConnectivityProbe();

        var connected = await probe.CanConnectAsync(
            IPAddress.Loopback.ToString(),
            port,
            TimeSpan.FromSeconds(2));

        Assert.True(connected);
        using var accepted = await acceptTask;
    }

    [Fact]
    public async Task CanConnectAsync_ReturnsFalse_WhenPortIsClosed()
    {
        int port;
        using (var listener = new TcpListener(IPAddress.Loopback, 0))
        {
            listener.Start();
            port = ((IPEndPoint)listener.LocalEndpoint).Port;
        }

        var probe = new TcpConnectivityProbe();
        var connected = await probe.CanConnectAsync(
            IPAddress.Loopback.ToString(),
            port,
            TimeSpan.FromSeconds(2));

        Assert.False(connected);
    }

    [Theory]
    [InlineData("", 445)]
    [InlineData("localhost", 0)]
    [InlineData("localhost", 65536)]
    public async Task CanConnectAsync_ReturnsFalse_ForInvalidTarget(string host, int port)
    {
        var probe = new TcpConnectivityProbe();

        var connected = await probe.CanConnectAsync(host, port, TimeSpan.FromSeconds(1));

        Assert.False(connected);
    }
}
