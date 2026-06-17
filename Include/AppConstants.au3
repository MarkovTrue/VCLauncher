#include-once

; ============================================================
; AppConstants.au3 - идентичность приложения и сетевые ссылки
; ============================================================
; ВНИМАНИЕ: этот файл качается raw из ветки main для проверки
; обновлений. Строка $gc_sAppVersion парсится regex, поэтому
; её формат менять нельзя.

Global Const $gc_sAppTitle   = "VCLauncher"            ; имя продукта
Global Const $gc_sAppVersion = "1.08"                  ; версия приложения
Global Const $gc_sAppName    = $gc_sAppTitle & " " & $gc_sAppVersion

; Сырой URL этого же файла в ветке main - источник проверки обновлений.
Global Const $gc_sVerCheckUrl = "https://raw.githubusercontent.com/MarkovTrue/VCLauncher/main/Include/AppConstants.au3"
Global Const $gc_sReleasesUrl = "https://github.com/MarkovTrue/VCLauncher/releases"
