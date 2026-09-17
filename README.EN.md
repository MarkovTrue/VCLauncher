[Русский](README.md) | [English](README.EN.md)

# <img src="Assets/Icons/HeaderIcon.png" alt="Logo" width="28" align="absmiddle"/>&nbsp; VCLauncher

[![Release](https://img.shields.io/github/v/release/MarkovTrue/VCLauncher?label=Release&color=%238a2be2&logo=github&logoColor=white)](https://github.com/MarkovTrue/VCLauncher/releases) [![Downloads](https://img.shields.io/github/downloads/MarkovTrue/VCLauncher/total?label=Downloads&color=%230078D4)](https://github.com/MarkovTrue/VCLauncher/releases)

A handy GUI launcher for [Video-compare](https://github.com/pixop/video-compare). It helps you visually compare two versions of a movie and pick the best one for your collection. It finds the time offset between releases on its own.


![Preview](Assets/Preview.en.png)

![Preview](Assets/Compare.png)

*Native overlay: Blu-ray release on the left, Open Matte 16:9 hybrid on the right*

### Features

- Two overlay modes: native or cropped to a common height
- Quick detection of the time offset between files
- The comparison window shrinks if the video doesn't fit the screen
- Quick switching between comparison modes: direct or vertical
- Drag-and-drop support
- The `Video-compare` launch command can be edited
- Cheat sheet with the main `Video-compare` hotkeys
- Cached resolutions and offsets, files and settings are remembered
- Light and dark themes, English and Russian
- The `Video-compare` console is hidden, errors are shown in a dialog

### VCLauncher uses (already bundled in the release)

- [Video-compare](https://github.com/pixop/video-compare) - the comparison engine itself
- [FFmpeg](https://github.com/FFmpeg/FFmpeg) - for extracting audio and video streams
- `Sync` - a CLI utility for offset detection, [algorithm description](SYNC.EN.md)

### Note

⚠️ Your antivirus may falsely flag the exe files. This is a known quirk of compiled AutoIt scripts and PyInstaller builds. The launcher source code is open, you can review it and compile it yourself.
