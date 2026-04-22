using System;
using System.IO;

namespace KeywordOcr.App.Services;

internal static class DesktopKeyStore
{
    public static string DirectoryPath =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Desktop", "key");

    public static string GetPath(string fileName) => Path.Combine(DirectoryPath, fileName);
}
