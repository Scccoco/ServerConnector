using System.Diagnostics;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace Connector.Desktop.Services;

// Patches Tekla Structures on this PC to extend native IFC export (62 entities + persist + property-set entities).
// Unlike ModelSharingProvisioningService (which IL-patches one DLL in memory), this service REPLACES five PREBUILT
// patched binaries plus one XSD edit, from a per-Tekla-build "patch set" staged on disk (delivered out-of-band).
//
//   bin\IFCExport4.dll
//   bin\Model.dll
//   bin\CommonObjects.dll
//   bin\Features\PropertyPaneFeature.dll
//   bin\Features\PropertySetsUIFeature.dll
//   Environments\common\inp\IfcPropertySetConfigurations.xsd   (text edit: +custom entities, idempotent)
//
// The patched DLLs are byte-exact per Tekla build (2024.0 SP5 vs 2025.0 SP7 are DIFFERENT sets). Applying a set to a
// mismatched build would corrupt the binary, so apply is HARD-GATED: it verifies the live TeklaStructures.exe build
// matches the staged manifest, and each staged file's SHA-256 matches the manifest, BEFORE touching bin. Idempotent
// (keeps a pristine .ifc-orig backup per file, detects a Tekla service-pack replacement), with one-click Rollback.
public sealed class IfcExportPatchService
{
    private const string MarkerDll = "IFCExport4.dll";      // anchor DLL that identifies a Tekla bin
    private const string BackupSuffix = ".ifc-orig";        // pristine backup (distinct from any manual *.bak)
    private const string StateFileName = ".structura-ifc.json";
    private const string ManifestFileName = "manifest.json";
    private const string XsdRelFromRoot = @"Environments\common\inp\IfcPropertySetConfigurations.xsd";
    private const string XsdMarker = "Structura-IFC-pset-entities";
    private const int FileReplaceMaxAttempts = 3;

    public string LogFilePath { get; }

    public IfcExportPatchService()
    {
        var root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ConnectorAgentDesktop");
        Directory.CreateDirectory(root);
        LogFilePath = Path.Combine(root, "ifc-patcher.log");
    }

    public bool IsTeklaRunning()
    {
        try { return Process.GetProcessesByName("TeklaStructures").Length > 0; }
        catch { return false; }
    }

    // Find the Tekla bin folder (contains IFCExport4.dll). Order: configured -> derive from Extensions path -> scan.
    public string ResolveTeklaBin(string? configuredBin, string? extensionsLocalPath)
    {
        if (!string.IsNullOrWhiteSpace(configuredBin) && File.Exists(Path.Combine(configuredBin, MarkerDll)))
            return configuredBin.Trim();

        if (!string.IsNullOrWhiteSpace(extensionsLocalPath))
        {
            var idx = extensionsLocalPath.IndexOf(@"\Environments", StringComparison.OrdinalIgnoreCase);
            if (idx > 0)
            {
                var candidate = Path.Combine(extensionsLocalPath.Substring(0, idx), "bin");
                if (File.Exists(Path.Combine(candidate, MarkerDll))) return candidate;
            }
        }

        try
        {
            const string teklaRoot = @"C:\TeklaStructures";
            if (Directory.Exists(teklaRoot))
            {
                var match = Directory.GetDirectories(teklaRoot)
                    .Select(d => Path.Combine(d, "bin"))
                    .Where(b => File.Exists(Path.Combine(b, MarkerDll)))
                    .OrderByDescending(b => b, StringComparer.OrdinalIgnoreCase)
                    .FirstOrDefault();
                if (!string.IsNullOrWhiteSpace(match)) return match;
            }
        }
        catch { /* fall through */ }

        return @"C:\TeklaStructures\2025.0\bin";
    }

    // Exact Tekla build from the (never-patched) TeklaStructures.exe, e.g. "2025.0.56843.0". Empty if not found.
    public string DetectBuild(string teklaBin)
    {
        try
        {
            var exe = Path.Combine(teklaBin, "TeklaStructures.exe");
            if (!File.Exists(exe)) return "";
            var fvi = FileVersionInfo.GetVersionInfo(exe);
            return fvi.FileVersion?.Trim() ?? "";
        }
        catch { return ""; }
    }

    public string ResolveXsdPath(string teklaBin)
    {
        var teklaRoot = Path.GetDirectoryName(teklaBin.TrimEnd('\\', '/')) ?? teklaBin;
        return Path.Combine(teklaRoot, XsdRelFromRoot);
    }

    public IfcPatchStatus GetStatus(string teklaBin)
    {
        var status = new IfcPatchStatus { TeklaBin = teklaBin, DetectedBuild = DetectBuild(teklaBin) };
        var anchor = Path.Combine(teklaBin, MarkerDll);
        status.TeklaFound = File.Exists(anchor);
        if (!status.TeklaFound) return status;

        var state = ReadState(Path.Combine(teklaBin, StateFileName));
        if (state is null) return status;

        try
        {
            var allMatch = state.Files.Count > 0 && state.Files.All(f =>
            {
                var live = Path.Combine(teklaBin, f.TargetRelpath);
                return File.Exists(live) && string.Equals(ComputeSha(File.ReadAllBytes(live)), f.PatchedSha, StringComparison.OrdinalIgnoreCase);
            });
            status.Applied = allMatch;
            status.NeedsReapply = !allMatch; // a Tekla service pack likely replaced one of our files
            status.SetVersion = state.SetVersion;
            status.AppliedBuild = state.TeklaBuild;
            status.AppliedUtc = state.AppliedUtc;
        }
        catch { /* treat unreadable state as not applied */ }

        return status;
    }

    public IfcPatchResult Apply(IfcPatchRequest request)
    {
        var teklaBin = request.TeklaBin.Trim();
        var stagingDir = request.StagingDir.Trim();

        if (IsTeklaRunning())
            return IfcPatchResult.Fail("Сейчас запущена Tekla Structures. Закройте Tekla и повторите установку патча.");
        if (!File.Exists(Path.Combine(teklaBin, MarkerDll)))
            return IfcPatchResult.Fail("Не найдена папка bin Tekla (нет " + MarkerDll + ") по пути: " + teklaBin + ".");

        IfcPatchManifest manifest;
        try { manifest = LoadManifest(stagingDir); }
        catch (Exception ex) { return IfcPatchResult.Fail("Не удалось прочитать набор патча (manifest.json).", ex.Message); }

        // --- HARD GATE 1: Tekla build must match the staged set (byte offsets are build-specific) ---
        var build = DetectBuild(teklaBin);
        if (string.IsNullOrEmpty(build) || !BuildMatches(build, manifest.TeklaBuild))
        {
            return IfcPatchResult.Fail(
                $"Сборка Tekla ({(string.IsNullOrEmpty(build) ? "не определена" : build)}) не совпадает с набором патча ({manifest.TeklaBuild}). " +
                "Этот набор собран под другую версию Tekla — установка отменена во избежание повреждения файлов.");
        }

        // --- HARD GATE 2: every staged file's SHA-256 must match the manifest (integrity / not corrupt) ---
        foreach (var f in manifest.Files)
        {
            var staged = Path.Combine(stagingDir, f.TargetRelpath);
            if (!File.Exists(staged))
                return IfcPatchResult.Fail("В наборе патча отсутствует файл: " + f.TargetRelpath + ".");
            var sha = ComputeSha(File.ReadAllBytes(staged));
            if (!string.Equals(sha, f.Sha256, StringComparison.OrdinalIgnoreCase))
                return IfcPatchResult.Fail("Контрольная сумма файла набора не совпадает (повреждён?): " + f.TargetRelpath + ".",
                    $"expected {f.Sha256}, got {sha}");
        }

        var newState = new IfcPatchState { TeklaBuild = build, SetVersion = manifest.SetVersion, AppliedUtc = DateTimeOffset.UtcNow };
        var replaced = new List<string>(); // live paths already replaced -> for auto-rollback on failure

        try
        {
            AppendLog($"Старт установки IFC-патча: bin='{teklaBin}', build='{build}', set='{manifest.SetVersion}'.");

            foreach (var f in manifest.Files)
            {
                var live = Path.Combine(teklaBin, f.TargetRelpath);
                var backup = live + BackupSuffix;
                if (!File.Exists(live))
                    throw new IOException("Не найден заменяемый файл Tekla: " + live);

                var pristine = CapturePristine(live, backup, f.TargetRelpath, ReadState(Path.Combine(teklaBin, StateFileName)));
                var patched = File.ReadAllBytes(Path.Combine(stagingDir, f.TargetRelpath));
                ReplaceFile(live, patched);
                replaced.Add(live);
                newState.Files.Add(new IfcPatchFileState
                {
                    TargetRelpath = f.TargetRelpath,
                    PristineSha = ComputeSha(pristine),
                    PatchedSha = ComputeSha(patched)
                });
                AppendLog("Заменён: " + f.TargetRelpath);
            }

            // XSD config edit (idempotent, makes its own .bak)
            if (manifest.Xsd is { Entities.Count: > 0 })
            {
                var xsdPath = ResolveXsdPath(teklaBin);
                var added = ApplyXsd(xsdPath, manifest.Xsd.Entities);
                newState.XsdApplied = true;
                newState.XsdPath = xsdPath;
                AppendLog($"XSD обновлён ({xsdPath}): +{added} сущностей.");
            }

            WriteState(Path.Combine(teklaBin, StateFileName), newState);

            // --- VERIFY: every live file now equals the staged patched SHA (+ XSD marker present) ---
            foreach (var f in manifest.Files)
            {
                var live = Path.Combine(teklaBin, f.TargetRelpath);
                if (!string.Equals(ComputeSha(File.ReadAllBytes(live)), f.Sha256, StringComparison.OrdinalIgnoreCase))
                    throw new IOException("Проверка после замены не прошла: " + f.TargetRelpath);
            }
            if (newState.XsdApplied && !File.ReadAllText(newState.XsdPath).Contains(XsdMarker))
                throw new IOException("Проверка XSD не прошла (маркер не найден).");

            AppendLog("IFC-патч установлен успешно.");
            return IfcPatchResult.Success("Tekla на этом компьютере пропатчена (IFC-экспорт расширен). " +
                "Перезапустите Tekla, чтобы изменения вступили в силу.");
        }
        catch (Exception ex)
        {
            AppendLog("Ошибка установки, выполняю авто-откат: " + ex.Message);
            AutoRollback(teklaBin, replaced); // restore the files we already touched
            var hint = ex is UnauthorizedAccessException
                ? "Нет прав на запись в папку Tekla — запустите с правами администратора (UAC)."
                : "Закройте Tekla Structures и программы, открывшие папку bin, и повторите.";
            return IfcPatchResult.Fail("Не удалось установить IFC-патч (выполнен откат изменений). " + hint, ex.Message);
        }
    }

    public IfcPatchResult Rollback(string teklaBin)
    {
        teklaBin = teklaBin.Trim();
        if (IsTeklaRunning())
            return IfcPatchResult.Fail("Сейчас запущена Tekla Structures. Закройте Tekla и повторите откат.");

        var statePath = Path.Combine(teklaBin, StateFileName);
        var state = ReadState(statePath);
        if (state is null || state.Files.Count == 0)
            return IfcPatchResult.Fail("Нет данных об установленном патче — откатывать нечего.");

        try
        {
            foreach (var f in state.Files)
            {
                var live = Path.Combine(teklaBin, f.TargetRelpath);
                var backup = live + BackupSuffix;
                if (File.Exists(backup))
                {
                    ReplaceFile(live, File.ReadAllBytes(backup));
                    AppendLog("Восстановлен из эталона: " + f.TargetRelpath);
                }
            }
            if (state.XsdApplied && !string.IsNullOrEmpty(state.XsdPath))
                RollbackXsd(state.XsdPath);

            TryDelete(statePath);
            AppendLog("Откат IFC-патча выполнен.");
            return IfcPatchResult.Success("Файлы Tekla восстановлены к исходному состоянию. Перезапустите Tekla.");
        }
        catch (Exception ex)
        {
            AppendLog("Ошибка отката: " + ex.Message);
            return IfcPatchResult.Fail("Не удалось полностью откатить патч. Закройте Tekla и повторите.", ex.Message);
        }
    }

    // ---- internals ----

    private void AutoRollback(string teklaBin, List<string> replacedLivePaths)
    {
        foreach (var live in replacedLivePaths)
        {
            try
            {
                var backup = live + BackupSuffix;
                if (File.Exists(backup)) ReplaceFile(live, File.ReadAllBytes(backup));
            }
            catch { /* best-effort */ }
        }
    }

    // Capture the pristine (un-patched) bytes for a file, re-capturing if a Tekla update replaced our patch.
    private byte[] CapturePristine(string live, string backup, string relpath, IfcPatchState? prevState)
    {
        if (!File.Exists(backup))
        {
            File.Copy(live, backup);            // first apply on a stock bin: live IS the pristine
            return File.ReadAllBytes(backup);
        }
        var liveSha = ComputeSha(File.ReadAllBytes(live));
        var prev = prevState?.Files.FirstOrDefault(x => string.Equals(x.TargetRelpath, relpath, StringComparison.OrdinalIgnoreCase));
        if (prev is not null && string.Equals(liveSha, prev.PatchedSha, StringComparison.OrdinalIgnoreCase))
            return File.ReadAllBytes(backup);   // live is exactly our patch -> backup is its pristine
        if (prev is not null && !string.IsNullOrEmpty(prev.PristineSha) &&
            !string.Equals(liveSha, prev.PristineSha, StringComparison.OrdinalIgnoreCase))
        {
            File.Copy(live, backup, overwrite: true); // SP update replaced the file -> re-capture new pristine
            return File.ReadAllBytes(backup);
        }
        return File.ReadAllBytes(backup);
    }

    private static bool BuildMatches(string detected, string manifestBuild)
    {
        if (string.IsNullOrWhiteSpace(manifestBuild)) return false;
        var d = detected.Trim();
        var m = manifestBuild.Trim();
        return d.StartsWith(m, StringComparison.OrdinalIgnoreCase) || m.StartsWith(d, StringComparison.OrdinalIgnoreCase);
    }

    // Port of tools/scripts/pset-entities-xsd.ps1: insert missing entities after IfcWindow, idempotent, .bak. Returns count added.
    private int ApplyXsd(string xsdPath, IReadOnlyDictionary<string, string> entities)
    {
        if (!File.Exists(xsdPath)) throw new FileNotFoundException("Не найден XSD: " + xsdPath);
        var raw = File.ReadAllText(xsdPath);
        if (raw.Contains(XsdMarker)) return 0; // already applied

        var nl = raw.Contains("\r\n") ? "\r\n" : "\n";
        var lines = raw.Replace("\r\n", "\n").Split('\n').ToList();

        var missing = entities.Where(kv => !raw.Contains("value=\"" + kv.Key + "\"")).ToList();
        if (missing.Count == 0) return 0;

        var iWin = lines.FindIndex(l => l.Contains("value=\"IfcWindow\""));
        if (iWin < 0) throw new InvalidOperationException("В XSD не найден enumeration IfcWindow — структура изменилась.");
        var iClose = iWin;
        while (iClose < lines.Count && !lines[iClose].Contains("</xs:enumeration>")) iClose++;
        if (iClose >= lines.Count) throw new InvalidOperationException("Не найден закрывающий тег enumeration для IfcWindow.");

        var indent = new string(lines[iWin].TakeWhile(c => c is ' ' or '\t').ToArray());
        var block = new List<string> { $"{indent}<!-- {XsdMarker}: custom IFC entities (mirror of «Объект IFC» dropdown) -->" };
        foreach (var (name, dom) in missing)
        {
            block.Add($"{indent}<xs:enumeration value=\"{name}\">");
            block.Add($"{indent}\t<xs:annotation>");
            block.Add($"{indent}\t\t<xs:documentation>{dom}</xs:documentation>");
            block.Add($"{indent}\t</xs:annotation>");
            block.Add($"{indent}</xs:enumeration>");
        }

        var bak = xsdPath + ".bak";
        if (!File.Exists(bak)) File.Copy(xsdPath, bak);

        var outLines = new List<string>();
        outLines.AddRange(lines.Take(iClose + 1));
        outLines.AddRange(block);
        outLines.AddRange(lines.Skip(iClose + 1));
        File.WriteAllText(xsdPath, string.Join(nl, outLines), new UTF8Encoding(false));
        return missing.Count;
    }

    private void RollbackXsd(string xsdPath)
    {
        var bak = xsdPath + ".bak";
        if (File.Exists(bak)) { File.Copy(bak, xsdPath, overwrite: true); AppendLog("XSD восстановлен из .bak."); }
    }

    private IfcPatchManifest LoadManifest(string stagingDir)
    {
        var path = Path.Combine(stagingDir, ManifestFileName);
        var m = JsonSerializer.Deserialize<IfcPatchManifest>(File.ReadAllText(path),
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
        if (m is null || m.Files.Count == 0) throw new InvalidOperationException("Пустой или некорректный manifest.json.");
        return m;
    }

    private static void ReplaceFile(string targetPath, byte[] content)
    {
        for (var attempt = 1; attempt <= FileReplaceMaxAttempts; attempt++)
        {
            var tempPath = targetPath + ".structura-ifc-" + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(targetPath)!);
                File.WriteAllBytes(tempPath, content);
                if (File.Exists(targetPath))
                {
                    var info = new FileInfo(targetPath);
                    if ((info.Attributes & FileAttributes.ReadOnly) != 0) info.Attributes &= ~FileAttributes.ReadOnly;
                    File.Copy(tempPath, targetPath, overwrite: true);
                    File.Delete(tempPath);
                }
                else
                {
                    File.Move(tempPath, targetPath);
                }
                return;
            }
            catch (Exception ex) when ((ex is IOException or UnauthorizedAccessException) && attempt < FileReplaceMaxAttempts)
            {
                TryDelete(tempPath);
                Thread.Sleep(TimeSpan.FromMilliseconds(300 * attempt));
            }
            catch
            {
                TryDelete(tempPath);
                throw;
            }
        }
    }

    private static void TryDelete(string path)
    {
        try { if (File.Exists(path)) File.Delete(path); } catch { /* best-effort */ }
    }

    public void AppendLog(string message)
    {
        try { File.AppendAllText(LogFilePath, $"[{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss}] {message}{Environment.NewLine}"); }
        catch { /* logging must never break patching */ }
    }

    private static string ComputeSha(byte[] data)
    {
        using var sha = SHA256.Create();
        return Convert.ToHexString(sha.ComputeHash(data));
    }

    private static IfcPatchState? ReadState(string statePath)
    {
        try
        {
            return File.Exists(statePath)
                ? JsonSerializer.Deserialize<IfcPatchState>(File.ReadAllText(statePath))
                : null;
        }
        catch { return null; }
    }

    private static void WriteState(string statePath, IfcPatchState state)
    {
        try { File.WriteAllText(statePath, JsonSerializer.Serialize(state, new JsonSerializerOptions { WriteIndented = true })); }
        catch { /* status detection degrades gracefully without the sidecar */ }
    }
}

// ---- manifest (shipped with each per-build patch set) ----
public sealed class IfcPatchManifest
{
    public string TeklaBuild { get; set; } = "";   // e.g. "2025.0.56843" (prefix-matched against TeklaStructures.exe)
    public string SetVersion { get; set; } = "";   // e.g. "2025.0-SP7-1"
    public List<IfcPatchFileEntry> Files { get; set; } = new();
    public IfcXsdSpec? Xsd { get; set; }
}

public sealed class IfcPatchFileEntry
{
    public string TargetRelpath { get; set; } = ""; // relative to Tekla bin, e.g. "Features\\PropertyPaneFeature.dll"
    public string Sha256 { get; set; } = "";
}

public sealed class IfcXsdSpec
{
    public Dictionary<string, string> Entities { get; set; } = new(); // entity name -> IFC domain (ARCH/MEP/INFRA/STRU)
}

// ---- request / result / status / per-bin state ----
public sealed class IfcPatchRequest
{
    public string TeklaBin { get; set; } = "";
    public string StagingDir { get; set; } = ""; // folder with manifest.json + the patched files mirrored by TargetRelpath
}

public sealed class IfcPatchResult
{
    public bool IsSuccess { get; init; }
    public string Message { get; init; } = "";
    public string TechnicalDetails { get; init; } = "";
    public static IfcPatchResult Success(string message) => new() { IsSuccess = true, Message = message };
    public static IfcPatchResult Fail(string message, string technicalDetails = "") =>
        new() { IsSuccess = false, Message = message, TechnicalDetails = technicalDetails };
}

public sealed class IfcPatchStatus
{
    public string TeklaBin { get; set; } = "";
    public bool TeklaFound { get; set; }
    public string DetectedBuild { get; set; } = "";
    public bool Applied { get; set; }
    public bool NeedsReapply { get; set; }
    public string SetVersion { get; set; } = "";
    public string AppliedBuild { get; set; } = "";
    public DateTimeOffset? AppliedUtc { get; set; }
}

internal sealed class IfcPatchState
{
    public string TeklaBuild { get; set; } = "";
    public string SetVersion { get; set; } = "";
    public List<IfcPatchFileState> Files { get; set; } = new();
    public bool XsdApplied { get; set; }
    public string XsdPath { get; set; } = "";
    public DateTimeOffset? AppliedUtc { get; set; }
}

internal sealed class IfcPatchFileState
{
    public string TargetRelpath { get; set; } = "";
    public string PristineSha { get; set; } = "";
    public string PatchedSha { get; set; } = "";
}
