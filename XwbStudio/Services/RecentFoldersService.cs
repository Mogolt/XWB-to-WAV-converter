using System.IO;
using System.Text.Json;

namespace XwbStudio.Services;

/// <summary>Persists the most recently used input folders next to the executable.</summary>
public sealed class RecentFoldersService
{
    private const string FileName = "recent_folders.json";
    private const int MaxRecent = 5;

    private static string StorePath => Path.Combine(AppContext.BaseDirectory, FileName);

    public List<string> Folders { get; } = [];

    public RecentFoldersService()
    {
        try
        {
            if (File.Exists(StorePath))
            {
                var loaded = JsonSerializer.Deserialize<List<string>>(File.ReadAllText(StorePath));
                if (loaded is not null)
                    Folders.AddRange(loaded.Take(MaxRecent));
            }
        }
        catch
        {
            // Corrupt or unreadable history is not fatal.
        }
    }

    public void Add(string path)
    {
        if (string.IsNullOrWhiteSpace(path) || !Directory.Exists(path))
            return;

        Folders.Remove(path);
        Folders.Insert(0, path);
        if (Folders.Count > MaxRecent)
            Folders.RemoveRange(MaxRecent, Folders.Count - MaxRecent);

        try
        {
            File.WriteAllText(StorePath, JsonSerializer.Serialize(Folders));
        }
        catch
        {
            // Read-only install dir etc. — history just won't persist.
        }
    }
}
