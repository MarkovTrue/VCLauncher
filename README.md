[Русский](README.md) | [English](README.EN.md)

# <img src="Preview/HeaderIcon.png" width="30" height="36" align="absmiddle" alt=""> VCLauncher

[![Release](https://img.shields.io/github/v/release/MarkovTrue/VCLauncher?label=Release&color=%238a2be2&logo=data%3Aimage%2Fsvg%2Bxml%3Bbase64%2CPHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNCAyNCIgZmlsbD0ibm9uZSIgc3Ryb2tlPSJ3aGl0ZSIgc3Ryb2tlLXdpZHRoPSIyIiBzdHJva2UtbGluZWNhcD0icm91bmQiIHN0cm9rZS1saW5lam9pbj0icm91bmQiPjxwYXRoIGQ9Ik0xMSAyMS43M2EyIDIgMCAwIDAgMiAwbDctNEEyIDIgMCAwIDAgMjEgMTZWOGEyIDIgMCAwIDAtMS0xLjczbC03LTRhMiAyIDAgMCAwLTIgMGwtNyA0QTIgMiAwIDAgMCAzIDh2OGEyIDIgMCAwIDAgMSAxLjczeiIvPjxwYXRoIGQ9Ik0xMiAyMlYxMiIvPjxwb2x5bGluZSBwb2ludHM9IjMuMjkgNyAxMiAxMiAyMC43MSA3Ii8%2BPC9zdmc%2B)](https://github.com/MarkovTrue/VCLauncher/releases) [![Downloads](https://img.shields.io/github/downloads/MarkovTrue/VCLauncher/total?label=Downloads&color=%230078D4&logo=data%3Aimage%2Fsvg%2Bxml%3Bbase64%2CPHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNCAyNCIgZmlsbD0ibm9uZSIgc3Ryb2tlPSJ3aGl0ZSIgc3Ryb2tlLXdpZHRoPSIyIiBzdHJva2UtbGluZWNhcD0icm91bmQiIHN0cm9rZS1saW5lam9pbj0icm91bmQiPjxwYXRoIGQ9Ik0yMSAxNXY0YTIgMiAwIDAgMS0yIDJINWEyIDIgMCAwIDEtMi0ydi00Ii8%2BPHBvbHlsaW5lIHBvaW50cz0iNyAxMCAxMiAxNSAxNyAxMCIvPjxsaW5lIHgxPSIxMiIgeDI9IjEyIiB5MT0iMTUiIHkyPSIzIi8%2BPC9zdmc%2B)](https://github.com/MarkovTrue/VCLauncher/releases)

Удобный GUI-лаунчер для [Video-compare](https://github.com/pixop/video-compare). Помогает наглядно сравнить две версии фильма и выбрать лучшую для коллекции. Сам находит смещение между релизами.


![Preview](Preview/Preview.png)

![Preview](Preview/Compare.png)

*Нативное наложение: слева Blu-ray релиз, справа Open Matte 16:9 гибрид*

### Возможности

- Два способа наложения: нативное или с подрезкой до общей высоты
- Быстрый поиск временной задержки между файлами
- Окно сравнения уменьшится, если видео выходит за пределы экрана
- Быстрое переключение между режимами сравнения: прямое или вертикальное
- Поддержка перетаскивания drag-and-drop
- Команду запуска `Video-compare` можно отредактировать
- Шпаргалка по основным горячим клавишам `Video-compare`
- Кеш разрешений и смещений, файлы и настройки запоминаются
- Светлая и тёмная тема, русский и английский язык
- Консоль `Video-compare` скрыта, ошибки показываются в окне

### VCLauncher использует (уже включено в релиз)

- [Video-compare](https://github.com/pixop/video-compare) - сам движок сравнения
- [FFmpeg](https://github.com/FFmpeg/FFmpeg) - для извлечения аудио и видео потоков
- `Sync` - консольная утилита для поиска смещения, [описание алгоритма](SYNC.md)

### Примечание

⚠️ Антивирус может ложно срабатывать на exe-файлы. Это известная особенность скомпилированных AutoIt скриптов и PyInstaller сборок. Исходный код лаунчера открыт, можно изучить его и скомпилировать самостоятельно.
