using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.ServiceProcess;
using System.Text;

namespace Connector.Desktop.Services;

// Brings up the firm AmneziaWG VPN tunnel on the user's PC using the bundled AmneziaWG client
// (tools/awg/{amneziawg.exe, awg.exe, wintun.dll}). The connector runs per-user without admin,
// so installing the tunnel service needs a one-time UAC elevation; afterwards the service is
// SYSTEM auto-start and survives reboots (no further prompts).
//
// Everything is opt-in / gated by the server (cfg vpn.enabled + a config delivered via bootstrap).
// If no VPN config is delivered, none of this runs and the connector behaves exactly as before.
public sealed class VpnProvisioningService
{
    private readonly string _awgExe;
    private readonly string _confDir;

    public string LogFilePath { get; }

    public VpnProvisioningService()
    {
        var baseDir = AppContext.BaseDirectory;
        _awgExe = Path.Combine(baseDir, "tools", "awg", "amneziawg.exe");

        var root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ConnectorAgentDesktop");
        _confDir = Path.Combine(root, "vpn");
        Directory.CreateDirectory(_confDir);
        LogFilePath = Path.Combine(root, "vpn.log");
    }

    public bool BundledClientPresent => File.Exists(_awgExe);

    private static string ServiceName(string tunnel) => "AmneziaWGTunnel$" + tunnel;
    private string ConfigPath(string tunnel) => Path.Combine(_confDir, tunnel + ".conf");

    public bool IsTunnelReady(string configContent, string tunnel)
    {
        return BundledClientPresent &&
               IsTunnelInstalled(tunnel) &&
               IsTunnelRunning(tunnel) &&
               VpnConfigComparer.FileMatches(ConfigPath(tunnel), configContent);
    }

    // In-process service queries (no sc.exe spawn) — safe to call from the UI thread.
    public bool IsTunnelInstalled(string tunnel)
    {
        try
        {
            using var sc = new ServiceController(ServiceName(tunnel));
            _ = sc.Status; // throws InvalidOperationException if the service does not exist
            return true;
        }
        catch (InvalidOperationException)
        {
            return false;
        }
        catch
        {
            return false;
        }
    }

    public bool IsTunnelRunning(string tunnel)
    {
        try
        {
            using var sc = new ServiceController(ServiceName(tunnel));
            return sc.Status == ServiceControllerStatus.Running;   // StartPending is NOT "up" (it may crash right after)
        }
        catch
        {
            return false;
        }
    }

    // Install (or reinstall) the tunnel from the given config and start it. One-time UAC.
    public VpnResult Enable(string configContent, string tunnel)
    {
        if (!BundledClientPresent)
        {
            return VpnResult.Fail("Встроенный клиент AmneziaWG не найден в составе коннектора. Переустановите коннектор.");
        }
        if (string.IsNullOrWhiteSpace(configContent))
        {
            return VpnResult.Fail("Сервер не передал конфигурацию VPN.");
        }

        var confPath = ConfigPath(tunnel);
        var normalizedConfig = VpnConfigComparer.Normalize(configContent) + Environment.NewLine;
        var installed = IsTunnelInstalled(tunnel);
        var configIsCurrent = VpnConfigComparer.FileMatches(confPath, normalizedConfig);

        // The normal startup path must be silent. A tunnel service is automatic and
        // survives reboots; if it is already running with the current server config,
        // there is nothing to reinstall and therefore no reason to show UAC again.
        if (installed && configIsCurrent && IsTunnelRunning(tunnel))
        {
            AppendLog("tunnel already running with current config");
            return VpnResult.Success("VPN уже включён и использует актуальную конфигурацию.");
        }

        if (installed && configIsCurrent)
        {
            var startScript =
                "$ErrorActionPreference='Stop'; " +
                "Start-Service -Name " + PowerShellLiteral(ServiceName(tunnel));
            var startCode = RunElevatedPowerShell(startScript, out var startError);
            if (startCode == ElevationCancelled)
            {
                return VpnResult.Fail("Запуск VPN отменён (не подтверждён запрос прав администратора).");
            }
            if (startCode == 0 && WaitForStableTunnel(tunnel))
            {
                AppendLog("existing tunnel service started");
                return VpnResult.Success("VPN включён.");
            }

            AppendLog("existing tunnel start failed: " + startError);
            return VpnResult.Fail(
                "Не удалось запустить установленный VPN-туннель. " +
                "Повторите подключение и проверьте лог VPN.");
        }

        if (installed)
        {
            // Keep the currently working config untouched until the user approves
            // the one elevated reconfiguration operation. The elevated script swaps
            // the file and reinstalls the service under a single UAC confirmation.
            var pendingPath = confPath + ".pending";
            try
            {
                WriteProtectedConfig(pendingPath, normalizedConfig);
                var reconfigureCode = RunElevatedReconfigure(
                    tunnel,
                    confPath,
                    pendingPath,
                    out var reconfigureError);
                if (reconfigureCode == ElevationCancelled)
                {
                    return VpnResult.Fail(
                        "Обновление VPN отменено (не подтверждён запрос прав администратора).");
                }
                if (reconfigureCode == 0 && WaitForStableTunnel(tunnel) &&
                    VpnConfigComparer.FileMatches(confPath, normalizedConfig))
                {
                    AppendLog("tunnel reconfigured and running");
                    return VpnResult.Success("VPN включён с актуальной конфигурацией.");
                }

                AppendLog("tunnel reconfigure failed: " + reconfigureError);
                return VpnResult.Fail(
                    "Не удалось обновить VPN-туннель. Проверьте лог VPN и повторите подключение.");
            }
            catch (Exception ex)
            {
                AppendLog("reconfigure preparation failed: " + ex.Message);
                return VpnResult.Fail("Не удалось обновить конфигурацию VPN: " + ex.Message);
            }
            finally
            {
                TryDeleteConf(pendingPath);
            }
        }

        try
        {
            // Tunnel name must match the file name (AmneziaWG derives it from the .conf filename).
            WriteProtectedConfig(confPath, normalizedConfig);
        }
        catch (Exception ex)
        {
            AppendLog("write conf failed: " + ex.Message);
            return VpnResult.Fail("Не удалось сохранить конфигурацию VPN: " + ex.Message);
        }

        AppendLog($"installtunnelservice {tunnel} from {confPath}");
        var code = RunElevated("/installtunnelservice \"" + confPath + "\"", out var elevErr);
        if (code == ElevationCancelled)
        {
            return VpnResult.Fail("Установка VPN отменена (не подтверждён запрос прав администратора). Нажмите ещё раз и подтвердите.");
        }

        // /installtunnelservice runs elevated with UseShellExecute, so we can't read its stdout;
        // verify by service state instead. The .conf MUST remain on disk — the tunnel service
        // re-reads it on every (re)start; deleting it makes the tunnel fail with
        // "Unable to load configuration ... cannot find the file".
        var running = WaitForStableTunnel(tunnel);
        if (running)
        {
            AppendLog("tunnel running and stable");
            return VpnResult.Success("VPN-доступ к общей папке включён.");
        }

        var hint = string.IsNullOrWhiteSpace(elevErr) ? "" : " " + elevErr;
        AppendLog("tunnel did not stay RUNNING." + hint);
        return VpnResult.Fail("VPN установлен, но туннель не удержался. Проверьте лог VPN (кнопка «Открыть лог»)." + hint);
    }

    public VpnResult Disable(string tunnel)
    {
        var confPath = ConfigPath(tunnel);
        if (!IsTunnelInstalled(tunnel))
        {
            TryDeleteConf(confPath);
            return VpnResult.Success("VPN уже отключён.");
        }
        var code = RunElevated("/uninstalltunnelservice " + tunnel, out _);
        if (code == ElevationCancelled)
        {
            return VpnResult.Fail("Отключение VPN отменено (не подтверждён запрос прав администратора).");
        }
        WaitUntil(() => !IsTunnelInstalled(tunnel), TimeSpan.FromSeconds(8));
        TryDeleteConf(confPath);
        AppendLog("tunnel uninstalled");
        return VpnResult.Success("VPN-доступ отключён.");
    }

    private void TryDeleteConf(string confPath)
    {
        try
        {
            if (File.Exists(confPath))
            {
                File.Delete(confPath);
            }
        }
        catch (Exception ex)
        {
            AppendLog("conf cleanup skipped: " + ex.Message);
        }
    }

    private void WriteProtectedConfig(string path, string configContent)
    {
        File.WriteAllText(path, configContent, new UTF8Encoding(false));
        // The tunnel service re-reads this file on every start, so it must stay
        // on disk; lock it down to SYSTEM, Administrators, and the owner.
        HardenConfAcl(path);
    }

    // The .conf holds a private key and must stay on disk (the tunnel service re-reads it), so lock it
    // down: remove inheritance, grant only SYSTEM + Administrators + the current user. Uses icacls
    // (no extra NuGet); the current user owns the file in their own LocalAppData, so no admin needed.
    private void HardenConfAcl(string path)
    {
        try
        {
            RunQuiet("icacls.exe", "\"" + path + "\" /inheritance:r");
            RunQuiet("icacls.exe", "\"" + path + "\" /grant:r \"*S-1-5-18:F\" \"*S-1-5-32-544:F\" \"" + Environment.UserName + ":F\"");
        }
        catch (Exception ex)
        {
            AppendLog("conf ACL harden skipped: " + ex.Message);
        }
    }

    private static void RunQuiet(string file, string arguments)
    {
        var psi = new ProcessStartInfo(file, arguments)
        {
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
        };
        using var p = Process.Start(psi);
        p?.WaitForExit();
    }

    public void AppendLog(string message)
    {
        try
        {
            File.AppendAllText(LogFilePath, $"[{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss}] {message}{Environment.NewLine}");
        }
        catch
        {
            // logging must never break VPN control
        }
    }

    private const int ElevationCancelled = -1223;

    private int RunElevated(string arguments, out string error)
    {
        error = string.Empty;
        try
        {
            var psi = new ProcessStartInfo(_awgExe, arguments)
            {
                UseShellExecute = true,   // required for Verb=runas (UAC)
                Verb = "runas",
                WindowStyle = ProcessWindowStyle.Hidden,
            };
            using var p = Process.Start(psi);
            if (p is null)
            {
                error = "не удалось запустить процесс AmneziaWG";
                return -1;
            }
            p.WaitForExit();
            return p.ExitCode;
        }
        catch (Win32Exception ex) when (ex.NativeErrorCode == 1223)
        {
            // ERROR_CANCELLED — user declined the UAC prompt
            return ElevationCancelled;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return -1;
        }
    }

    private int RunElevatedReconfigure(
        string tunnel,
        string confPath,
        string pendingPath,
        out string error)
    {
        var serviceName = ServiceName(tunnel);
        var backupPath = confPath + ".rollback";
        var script = $$"""
            $ErrorActionPreference = 'Stop'
            $awg = {{PowerShellLiteral(_awgExe)}}
            $tunnel = {{PowerShellLiteral(tunnel)}}
            $serviceName = {{PowerShellLiteral(serviceName)}}
            $confPath = {{PowerShellLiteral(confPath)}}
            $pendingPath = {{PowerShellLiteral(pendingPath)}}
            $backupPath = {{PowerShellLiteral(backupPath)}}
            if (Test-Path -LiteralPath $confPath) {
                Copy-Item -LiteralPath $confPath -Destination $backupPath -Force
            }
            try {
                & $awg /uninstalltunnelservice $tunnel
                $deadline = [DateTime]::UtcNow.AddSeconds(15)
                while ((Get-Service -Name $serviceName -ErrorAction SilentlyContinue) -and
                       [DateTime]::UtcNow -lt $deadline) {
                    Start-Sleep -Milliseconds 300
                }
                if (Get-Service -Name $serviceName -ErrorAction SilentlyContinue) {
                    throw 'Previous tunnel service is still deleting.'
                }
                Copy-Item -LiteralPath $pendingPath -Destination $confPath -Force
                & $awg /installtunnelservice $confPath
                if ($LASTEXITCODE -ne 0) {
                    throw "AmneziaWG install failed with exit code $LASTEXITCODE."
                }
            }
            catch {
                if (Test-Path -LiteralPath $backupPath) {
                    Copy-Item -LiteralPath $backupPath -Destination $confPath -Force
                    & $awg /installtunnelservice $confPath
                }
                throw
            }
            finally {
                Remove-Item -LiteralPath $backupPath -Force -ErrorAction SilentlyContinue
            }
            """;
        return RunElevatedPowerShell(script, out error);
    }

    private int RunElevatedPowerShell(string script, out string error)
    {
        error = string.Empty;
        try
        {
            var encoded = Convert.ToBase64String(Encoding.Unicode.GetBytes(script));
            var psi = new ProcessStartInfo(
                "powershell.exe",
                "-NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand " + encoded)
            {
                UseShellExecute = true,
                Verb = "runas",
                WindowStyle = ProcessWindowStyle.Hidden,
            };
            using var process = Process.Start(psi);
            if (process is null)
            {
                error = "не удалось запустить повышенный процесс";
                return -1;
            }
            process.WaitForExit();
            return process.ExitCode;
        }
        catch (Win32Exception ex) when (ex.NativeErrorCode == 1223)
        {
            return ElevationCancelled;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return -1;
        }
    }

    private bool WaitForStableTunnel(string tunnel)
    {
        var running = WaitUntil(() => IsTunnelRunning(tunnel), TimeSpan.FromSeconds(15));
        if (running)
        {
            Thread.Sleep(2500);
            running = IsTunnelRunning(tunnel);
        }
        return running;
    }

    private static string PowerShellLiteral(string value)
    {
        return "'" + value.Replace("'", "''", StringComparison.Ordinal) + "'";
    }

    private static bool WaitUntil(Func<bool> condition, TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        while (DateTime.UtcNow < deadline)
        {
            if (condition())
            {
                return true;
            }
            Thread.Sleep(400);
        }
        return condition();
    }
}

public sealed class VpnResult
{
    public bool IsSuccess { get; init; }
    public string Message { get; init; } = "";

    public static VpnResult Success(string message) => new() { IsSuccess = true, Message = message };
    public static VpnResult Fail(string message) => new() { IsSuccess = false, Message = message };
}
