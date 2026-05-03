using System.Diagnostics;
using System.IO.Compression;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32;

const string AppName = "KeywordOCR v3";
const string AppProcessName = "KeywordOcr";
const string Marker = "KOCRPAYLOADv1";

try
{
    var noRun = args.Any(a => string.Equals(a, "--no-run", StringComparison.OrdinalIgnoreCase)
        || string.Equals(a, "/NoRun", StringComparison.OrdinalIgnoreCase)
        || string.Equals(a, "/quiet", StringComparison.OrdinalIgnoreCase));

    var installDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "KeywordOcr");

    var tempPayload = Path.Combine(Path.GetTempPath(), $"KeywordOcr_payload_{Guid.NewGuid():N}.zip");
    ExtractAppendedPayload(tempPayload);

    StopInstalledApp(installDir);

    Directory.CreateDirectory(installDir);
    ZipFile.ExtractToDirectory(tempPayload, installDir, overwriteFiles: true);
    File.Delete(tempPayload);

    var exePath = Path.Combine(installDir, "KeywordOcr.exe");
    if (!File.Exists(exePath))
        throw new FileNotFoundException("KeywordOcr.exe was not installed.", exePath);

    CreateShortcuts(exePath, installDir);
    RegisterUninstall(exePath, installDir);

    Console.WriteLine($"{AppName} installed to {installDir}");

    if (!noRun)
    {
        Process.Start(new ProcessStartInfo
        {
            FileName = exePath,
            WorkingDirectory = installDir,
            UseShellExecute = true
        });
    }

    return 0;
}
catch (Exception ex)
{
    Console.Error.WriteLine(ex.Message);
    Console.Error.WriteLine(ex);
    Console.Error.WriteLine("Press Enter to close.");
    Console.ReadLine();
    return 1;
}

static void ExtractAppendedPayload(string destinationPath)
{
    var exePath = Environment.ProcessPath ?? Assembly.GetExecutingAssembly().Location;
    var markerBytes = Encoding.ASCII.GetBytes(Marker);

    using var input = File.OpenRead(exePath);
    if (input.Length < markerBytes.Length + sizeof(long))
        throw new InvalidDataException("Installer payload marker was not found.");

    input.Seek(-markerBytes.Length, SeekOrigin.End);
    var markerBuffer = new byte[markerBytes.Length];
    ReadExactly(input, markerBuffer);
    if (!markerBuffer.SequenceEqual(markerBytes))
        throw new InvalidDataException("Installer payload marker was not found.");

    input.Seek(-(markerBytes.Length + sizeof(long)), SeekOrigin.End);
    Span<byte> lengthBuffer = stackalloc byte[sizeof(long)];
    input.ReadExactly(lengthBuffer);
    var payloadLength = BitConverter.ToInt64(lengthBuffer);
    var payloadOffset = input.Length - markerBytes.Length - sizeof(long) - payloadLength;
    if (payloadOffset < 0 || payloadLength <= 0)
        throw new InvalidDataException("Installer payload length is invalid.");

    input.Seek(payloadOffset, SeekOrigin.Begin);
    using var output = File.Create(destinationPath);
    CopyBytes(input, output, payloadLength);
}

static void StopInstalledApp(string installDir)
{
    foreach (var process in Process.GetProcessesByName(AppProcessName))
    {
        try
        {
            var path = process.MainModule?.FileName;
            if (!string.IsNullOrWhiteSpace(path)
                && path.StartsWith(installDir, StringComparison.OrdinalIgnoreCase))
            {
                process.Kill(entireProcessTree: true);
                process.WaitForExit(5000);
            }
        }
        catch
        {
        }
    }
}

static void CreateShortcuts(string exePath, string installDir)
{
    var desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
    CreateShortcut(Path.Combine(desktop, $"{AppName}.lnk"), exePath, installDir);

    var programs = Environment.GetFolderPath(Environment.SpecialFolder.Programs);
    var startMenuDir = Path.Combine(programs, AppName);
    Directory.CreateDirectory(startMenuDir);
    CreateShortcut(Path.Combine(startMenuDir, $"{AppName}.lnk"), exePath, installDir);
}

static void CreateShortcut(string shortcutPath, string exePath, string installDir)
{
    var shellType = Type.GetTypeFromProgID("WScript.Shell")
        ?? throw new InvalidOperationException("WScript.Shell is not available.");
    dynamic shell = Activator.CreateInstance(shellType)!;
    dynamic shortcut = shell.CreateShortcut(shortcutPath);
    shortcut.TargetPath = exePath;
    shortcut.WorkingDirectory = installDir;
    shortcut.IconLocation = $"{exePath},0";
    shortcut.Description = AppName;
    shortcut.Save();
    Marshal.FinalReleaseComObject(shortcut);
    Marshal.FinalReleaseComObject(shell);
}

static void RegisterUninstall(string exePath, string installDir)
{
    var uninstallScript = Path.Combine(installDir, "uninstall.ps1");
    using var key = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\KeywordOcr")
        ?? throw new InvalidOperationException("Unable to create uninstall registry key.");

    key.SetValue("DisplayName", AppName);
    key.SetValue("DisplayVersion", "3.0");
    key.SetValue("Publisher", "rkghrud-dev");
    key.SetValue("InstallLocation", installDir);
    key.SetValue("DisplayIcon", $"{exePath},0");
    key.SetValue("UninstallString", $"powershell.exe -NoProfile -ExecutionPolicy Bypass -File \"{uninstallScript}\"");
    key.SetValue("QuietUninstallString", $"powershell.exe -NoProfile -ExecutionPolicy Bypass -File \"{uninstallScript}\" -Quiet");
    key.SetValue("NoModify", 1, RegistryValueKind.DWord);
    key.SetValue("NoRepair", 1, RegistryValueKind.DWord);

    var sizeKb = Directory.EnumerateFiles(installDir, "*", SearchOption.AllDirectories)
        .Select(path => new FileInfo(path).Length)
        .Sum() / 1024;
    key.SetValue("EstimatedSize", (int)Math.Min(sizeKb, int.MaxValue), RegistryValueKind.DWord);
}

static void ReadExactly(Stream stream, byte[] buffer)
{
    var offset = 0;
    while (offset < buffer.Length)
    {
        var read = stream.Read(buffer, offset, buffer.Length - offset);
        if (read == 0)
            throw new EndOfStreamException();
        offset += read;
    }
}

static void CopyBytes(Stream input, Stream output, long bytesToCopy)
{
    var buffer = new byte[1024 * 1024];
    var remaining = bytesToCopy;
    while (remaining > 0)
    {
        var read = input.Read(buffer, 0, (int)Math.Min(buffer.Length, remaining));
        if (read == 0)
            throw new EndOfStreamException();
        output.Write(buffer, 0, read);
        remaining -= read;
    }
}
