using System.Net.Sockets;

namespace Connector.Desktop.Services;

public sealed class TcpConnectivityProbe
{
    public async Task<bool> CanConnectAsync(
        string host,
        int port,
        TimeSpan timeout,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(host) || port is < 1 or > 65535 || timeout <= TimeSpan.Zero)
        {
            return false;
        }

        using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeoutCts.CancelAfter(timeout);
        using var client = new TcpClient();

        try
        {
            await client.ConnectAsync(host.Trim(), port, timeoutCts.Token);
            return client.Connected;
        }
        catch (OperationCanceledException)
        {
            return false;
        }
        catch (SocketException)
        {
            return false;
        }
    }
}
