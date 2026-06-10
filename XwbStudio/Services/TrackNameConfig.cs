using System.IO;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace XwbStudio.Services;

/// <summary>
/// Loads the optional config.json that maps hex track names to friendly names,
/// keyed by bank file name: { "track_names": { "bio4bgm": { "00000000": "main_menu" } } }.
/// </summary>
public static class TrackNameConfig
{
    public const string FileName = "config.json";

    public static Dictionary<string, Dictionary<string, string>> Load(string path)
    {
        var result = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);

        var root = JsonNode.Parse(File.ReadAllText(path));
        if (root?["track_names"] is not JsonObject banks)
            return result;

        foreach (var (bankName, names) in banks)
        {
            if (names is not JsonObject map)
                continue;
            var bank = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var (hex, friendly) in map)
            {
                if (friendly is JsonValue v && v.TryGetValue<string>(out var s))
                    bank[hex] = s;
            }
            result[bankName] = bank;
        }

        return result;
    }

    public static void WriteTemplate(string path)
    {
        var template = new JsonObject
        {
            ["_readme"] = "Optional: map hex track names to friendly names. Example below.",
            ["_example"] = new JsonObject
            {
                ["bio4bgm"] = new JsonObject
                {
                    ["00000000"] = "main_menu_theme",
                    ["00000001"] = "village_ambience",
                },
            },
            ["track_names"] = new JsonObject(),
        };

        File.WriteAllText(path, template.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));
    }
}
