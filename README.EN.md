[Русский](README.md) | [English](README.EN.md)

# <img src="Preview/HeaderIcon.png" width="28" height="34" align="absmiddle" alt=""> VCLauncher

[![Release](https://img.shields.io/github/v/release/MarkovTrue/VCLauncher?label=Release&color=%238a2be2&logo=data%3Aimage%2Fsvg%2Bxml%3Bbase64%2CPHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNCAyNCIgZmlsbD0ibm9uZSIgc3Ryb2tlPSJ3aGl0ZSIgc3Ryb2tlLXdpZHRoPSIyIiBzdHJva2UtbGluZWNhcD0icm91bmQiIHN0cm9rZS1saW5lam9pbj0icm91bmQiPjxwYXRoIGQ9Ik0xMSAyMS43M2EyIDIgMCAwIDAgMiAwbDctNEEyIDIgMCAwIDAgMjEgMTZWOGEyIDIgMCAwIDAtMS0xLjczbC03LTRhMiAyIDAgMCAwLTIgMGwtNyA0QTIgMiAwIDAgMCAzIDh2OGEyIDIgMCAwIDAgMSAxLjczeiIvPjxwYXRoIGQ9Ik0xMiAyMlYxMiIvPjxwb2x5bGluZSBwb2ludHM9IjMuMjkgNyAxMiAxMiAyMC43MSA3Ii8%2BPC9zdmc%2B)](https://github.com/MarkovTrue/VCLauncher/releases) [![Downloads](https://img.shields.io/github/downloads/MarkovTrue/VCLauncher/total?label=Downloads&color=%230078D4&logo=data%3Aimage%2Fsvg%2Bxml%3Bbase64%2CPHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNCAyNCIgZmlsbD0ibm9uZSIgc3Ryb2tlPSJ3aGl0ZSIgc3Ryb2tlLXdpZHRoPSIyIiBzdHJva2UtbGluZWNhcD0icm91bmQiIHN0cm9rZS1saW5lam9pbj0icm91bmQiPjxwYXRoIGQ9Ik0yMSAxNXY0YTIgMiAwIDAgMS0yIDJINWEyIDIgMCAwIDEtMi0ydi00Ii8%2BPHBvbHlsaW5lIHBvaW50cz0iNyAxMCAxMiAxNSAxNyAxMCIvPjxsaW5lIHgxPSIxMiIgeDI9IjEyIiB5MT0iMTUiIHkyPSIzIi8%2BPC9zdmc%2B)](https://github.com/MarkovTrue/VCLauncher/releases)

A handy GUI launcher for [Video-compare](https://github.com/pixop/video-compare). It helps you visually compare two versions of a movie and pick the best one for your collection. It finds the time offset between releases on its own.


![Preview](Preview/Preview.en.png)

<p align="center">
  <img src="Preview/Compare.png" alt="Compare"><br>
  <sub>Native overlay: Blu-ray release on the left, Open Matte 16:9 hybrid on the right</sub>
</p>

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

### VCLauncher uses (bundled in the release)

- [`Video-compare`](https://github.com/pixop/video-compare) – comparison: overlay, modes, hotkeys
- [`FFmpeg`](https://github.com/FFmpeg/FFmpeg) – video resolution and frames for offset detection
- [`Sync`](SYNC.EN.md) – time offset detection between files

### Note

⚠️ Your antivirus may falsely flag the exe files. This is a known quirk of compiled AutoIt scripts and PyInstaller builds. The launcher source code is open, you can review it and compile it yourself.
