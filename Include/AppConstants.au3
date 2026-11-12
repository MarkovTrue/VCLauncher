#include-once

; ============================================================
; AppConstants.au3 - идентичность приложения и сетевые ссылки
; ============================================================
; ВНИМАНИЕ: версии до 1.11 включительно качают этот файл raw
; из ветки main и ищут regex строку $gc_sAppVersion. Пока они
; в ходу, формат строки менять нельзя. С 1.12 проверка идёт
; по последнему релизу GitHub, см. $gc_sVerCheckUrl.

; Версия exe, менять здесь вместе с $gc_sAppVersion. Aut2Exe берёт pragma и из подключаемого файла
#pragma compile(ProductVersion, 1.12)
#pragma compile(FileVersion, 1.12)

Global Const $gc_sAppTitle   = "VCLauncher"            ; имя продукта
Global Const $gc_sAppVersion = "1.12"                  ; версия приложения
Global Const $gc_sAppName    = $gc_sAppTitle & " " & $gc_sAppVersion

; Последний релиз GitHub - источник проверки обновлений. Версия в поле tag_name.
Global Const $gc_sVerCheckUrl = "https://api.github.com/repos/MarkovTrue/VCLauncher/releases/latest"
Global Const $gc_sReleasesUrl = "https://github.com/MarkovTrue/VCLauncher/releases"
