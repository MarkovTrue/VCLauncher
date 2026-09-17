[Русский](README.md) | [English](README.en.md)

# <img src="Assets/Icons/HeaderIcon.png" alt="Logo" width="28" align="absmiddle"/>&nbsp; VCLauncher

[![GitHub Release](https://img.shields.io/github/release/MarkovTrue/VCLauncher)](https://github.com/MarkovTrue/VCLauncher/releases) [![Downloads](https://img.shields.io/github/downloads/MarkovTrue/VCLauncher/total?label=downloads&color=blue)](https://github.com/MarkovTrue/VCLauncher/releases)

A handy GUI launcher for [Video Compare](https://github.com/pixop/video-compare). It helps you visually compare two versions of a movie and decide which one fits your collection better. You can also quickly detect a time offset and use it to transfer audio tracks between releases.


![Preview](Assets/Preview.en.png)

![Preview](Assets/Compare.png)

### Features

- ✂️ Automatic height-crop calculation for comparing videos with different aspect ratios
- ↔️ Quick detection of the time offset between video files
- Auto-scaling the window when the video exceeds the screen
- Quick switching between comparison modes
- Drag-and-drop support
- Manual editing of the `Video-compare` launch command
- Cheat sheet with some commonly used `Video-compare` hotkeys
- Caching and autosave
- Theme and localization support
- Hiding the `Video-compare` console

### VCLauncher uses (already bundled in the release)

- [Video-compare](https://github.com/pixop/video-compare) - the comparison engine itself
- [FFmpeg](https://github.com/FFmpeg/FFmpeg) - for extracting audio/video streams
- `Sync.exe` - proprietary CLI utility for time shift search

### Note

⚠️ Your antivirus may falsely flag the exe files. This is a known quirk of compiled AutoIt scripts and PyInstaller builds. The launcher source code is open, so if in doubt you can review it and compile it yourself.
