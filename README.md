# XWB Studio

A modern GUI for extracting, injecting, and converting XWB (XACT wave bank)
audio. Originally built to make RE4 (2005 PC) audio modding painless, but it
works on XWB files from any XACT-based game.

The reason this exists: older XWB tools are either command-line only, crash on
larger banks, or both. This one is a single Windows EXE — double-click, pick a
folder, done.

**Version 2.0** is a ground-up rewrite in C#/WPF (previously Python/tkinter):
same features and workflow, much smoother UI — animated tab transitions,
eased progress, hardware-accelerated rendering — and a cleaner codebase
underneath. The old Python version lives on in [`legacy/`](legacy/xwb_extractor.py).

## Features

- **Extract** — point it at a folder of `.xwb` files and dump every track to
  WAV. Supports PCM, XMA, ADPCM, WMA codecs. Optional `config.json` to map
  hex track indices to friendly names per bank.
- **Individual Extraction** — browse a single XWB, see every track with
  codec / rate / size, preview with the built-in player, extract only the
  tracks you want.
- **Inject** — load an XWB, select a track, preview, pick a replacement WAV,
  hit **Replace & Rebuild**. Writes the modified XWB in place or to a new
  folder.
- **Convert** — bundle a list of WAVs (or a whole folder — drag & drop works)
  into a new `.xwb` wave bank with a custom bank name.
- **Recent folders** shortcut for fast re-extraction across sessions.

## Supported games

Built and tested against **Resident Evil 4 (2005 PC / Ultimate HD Edition)**.
Works on XWB files from other XACT-based games too — the format is standard
(little- and big-endian banks, format versions 1–46, compact banks, and banks
embedded in container files). The occasional game-specific oddity may surface;
file an issue if you hit one.

## Installation

**Windows 10 / 11, x64.**

1. Download `XwbStudio.exe` from the
   [latest release](https://github.com/Mogolt/XWB-to-WAV-converter/releases/latest).
2. Run it. Single self-contained EXE — no install, no dependencies, no .NET
   required on the target machine.

Portable by design. Put it anywhere and launch from there.

## Usage

### Extract everything from a folder of XWBs

1. Open the **Extract** tab.
2. Set *XWB Folder (input)* to the folder containing your `.xwb` files.
3. Set *Output Folder*.
4. *(Optional)* Load a `config.json` that maps hex track indices to names —
   your output files get friendly names like `main_menu_theme.wav` instead of
   `00000000.wav`.
5. Hit **Extract All**.

`config.json` format:

```json
{
  "track_names": {
    "bio4bgm": {
      "00000000": "main_menu_theme",
      "00000001": "village_ambience"
    }
  }
}
```

### Extract specific tracks only

1. **Extract** tab → tick **Individual Extraction**.
2. Browse to a single `.xwb`. The right pane lists every track with codec /
   sample rate / duration / size.
3. Select the ones you want, preview with the play button, then extract.

### Replace a track in an XWB (inject)

1. **Inject** tab → browse and load the XWB you want to modify.
2. Pick a track, preview it so you know you've got the right one.
3. Browse a replacement `.wav`.
4. **Replace & Rebuild** — overwrites in place, or writes to a new folder if
   you'd rather keep the original.

### Build a new XWB from WAVs

1. **Convert** tab → add WAVs individually or point at a folder.
2. Set the output `.xwb` path and a bank name.
3. **Convert to XWB**.

## Building from source

Requires the [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0).

```powershell
dotnet build XwbStudio/XwbStudio.csproj            # debug build
dotnet run --project CoreTests/CoreTests.csproj    # round-trip tests
dotnet publish XwbStudio/XwbStudio.csproj -c Release -r win-x64 `
    --self-contained true -p:PublishSingleFile=true   # portable EXE
```

Project layout:

```
XwbStudio/
├── Core/          binary format logic — parser, extractor, builder, injector
├── Services/      audio preview (winmm), recent folders, track-name config
├── ViewModels/    MVVM layer — all app behavior
├── Views/         one XAML view per tab
└── Themes/        dark theme + animated control styles
CoreTests/         build → parse → extract → inject round-trip test
legacy/            the original Python/tkinter version (1.x)
```

## Not distributing game files

XWB Studio contains **no** game audio. You must extract your own `.xwb` files
from your own legally-owned copy of the game.

## License

[GPL-2.0-or-later](LICENSE). Copyright © 2026 Mogolt.

## Related projects

- **[ACB Tool](https://github.com/Mogolt/acb-tool)** — sibling project for
  the CRIWARE ADX2 (`.acb`/`.awb`) banks used in the 2016 PS4 / 2019 Switch
  ports of RE4. Same UX, different container family.

## Author

**Mogolt**.
