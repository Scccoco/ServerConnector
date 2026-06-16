using System.Diagnostics;
using System.Windows.Input;
using Connector.Desktop.Mvvm;

namespace Connector.Desktop.Features.Structura;

// ViewModel for the Structura "Сервисы" feature module. Holds the six Structura access values (Speckle +
// Nextcloud URLs / logins / encrypted passwords), surfaces a one-line access status, opens each service in the
// browser, and hands decrypted credentials to the View's dialog. The values are pushed in by the shell via
// Initialize(...) — the VM never touches AppSettings/SettingsService directly (decrypt is an injected callback).
public sealed class StructuraViewModel : ObservableObject
{
    private Func<string, string> _decrypt = s => s;

    // Shell-owned glue, wired after construction by the shell (so the module can be built arg-less in the catalog):
    // Log -> AppendLog (best-effort journal); Decrypt -> SettingsService.DecryptToken (turns the *Cipher into plaintext).
    public Action<string>? Log { get; set; }
    public Func<string, string> Decrypt { get => _decrypt; set => _decrypt = value ?? (s => s); }

    // The View sets this to open the credentials dialog: (title, domain, login, decryptedPassword).
    public Action<string, string, string, string>? ShowAccessHandler { get; set; }

    public StructuraViewModel(Action<string>? log = null, Func<string, string>? decrypt = null)
    {
        Log = log;
        if (decrypt is not null) _decrypt = decrypt;

        OpenSpeckleCommand = new RelayCommand(OpenSpeckle);
        OpenNextcloudCommand = new RelayCommand(OpenNextcloud);
        ShowSpeckleAccessCommand = new RelayCommand(ShowSpeckleAccess);
        ShowNextcloudAccessCommand = new RelayCommand(ShowNextcloudAccess);
    }

    private string _speckleUrl = "";
    public string SpeckleUrl { get => _speckleUrl; set { if (SetProperty(ref _speckleUrl, value)) RefreshStatus(); } }

    private string _speckleLogin = "";
    public string SpeckleLogin { get => _speckleLogin; set { if (SetProperty(ref _speckleLogin, value)) RefreshStatus(); } }

    private string _specklePasswordCipher = "";
    public string SpecklePasswordCipher { get => _specklePasswordCipher; set => SetProperty(ref _specklePasswordCipher, value); }

    private string _nextcloudUrl = "";
    public string NextcloudUrl { get => _nextcloudUrl; set { if (SetProperty(ref _nextcloudUrl, value)) RefreshStatus(); } }

    private string _nextcloudLogin = "";
    public string NextcloudLogin { get => _nextcloudLogin; set { if (SetProperty(ref _nextcloudLogin, value)) RefreshStatus(); } }

    private string _nextcloudPasswordCipher = "";
    public string NextcloudPasswordCipher { get => _nextcloudPasswordCipher; set => SetProperty(ref _nextcloudPasswordCipher, value); }

    private string _accessStatusLine = "Structura: доступы еще не получены с сервера";
    public string AccessStatusLine { get => _accessStatusLine; private set => SetProperty(ref _accessStatusLine, value); }

    public ICommand OpenSpeckleCommand { get; }
    public ICommand OpenNextcloudCommand { get; }
    public ICommand ShowSpeckleAccessCommand { get; }
    public ICommand ShowNextcloudAccessCommand { get; }

    // Bulk-set the six values from the shell's settings snapshot (suppresses per-setter status churn).
    public void Initialize(
        string speckleUrl, string speckleLogin, string specklePasswordCipher,
        string nextcloudUrl, string nextcloudLogin, string nextcloudPasswordCipher)
    {
        _speckleUrl = speckleUrl ?? "";
        _speckleLogin = speckleLogin ?? "";
        _specklePasswordCipher = specklePasswordCipher ?? "";
        _nextcloudUrl = nextcloudUrl ?? "";
        _nextcloudLogin = nextcloudLogin ?? "";
        _nextcloudPasswordCipher = nextcloudPasswordCipher ?? "";
        OnPropertyChanged(nameof(SpeckleUrl));
        OnPropertyChanged(nameof(SpeckleLogin));
        OnPropertyChanged(nameof(SpecklePasswordCipher));
        OnPropertyChanged(nameof(NextcloudUrl));
        OnPropertyChanged(nameof(NextcloudLogin));
        OnPropertyChanged(nameof(NextcloudPasswordCipher));
        RefreshStatus();
    }

    public void RefreshStatus()
    {
        var hasSpeckle = !string.IsNullOrWhiteSpace(SpeckleUrl);
        var hasNextcloud = !string.IsNullOrWhiteSpace(NextcloudUrl);
        var hasAnyLogin =
            !string.IsNullOrWhiteSpace(SpeckleLogin) ||
            !string.IsNullOrWhiteSpace(NextcloudLogin);

        AccessStatusLine = hasSpeckle || hasNextcloud
            ? hasAnyLogin
                ? "Structura: доступы получены с сервера и привязаны к текущему токену"
                : "Structura: ссылки получены, логины еще не назначены в админке"
            : "Structura: доступы еще не получены с сервера";
    }

    private void OpenSpeckle() => OpenStructuraWebPage(SpeckleUrl, "Speckle");

    private void OpenNextcloud() => OpenStructuraWebPage(NextcloudUrl, "Nextcloud");

    private void ShowSpeckleAccess() => ShowAccess(
        "Сервер моделей - Speckle", SpeckleUrl, SpeckleLogin, SpecklePasswordCipher);

    private void ShowNextcloudAccess() => ShowAccess(
        "Среда взаимодействия - Nextcloud", NextcloudUrl, NextcloudLogin, NextcloudPasswordCipher);

    private void ShowAccess(string title, string domain, string login, string passwordCipherBase64) =>
        ShowAccessHandler?.Invoke(title, domain, login, _decrypt(passwordCipherBase64));

    private void OpenStructuraWebPage(string url, string label)
    {
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri))
        {
            throw new InvalidOperationException("Некорректная ссылка " + label + ".");
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = uri.ToString(),
            UseShellExecute = true
        });
        Log?.Invoke("Открыта страница " + label + ": " + uri);
    }
}
