# outlook-lib

Автоматически обновляемая библиотека для Outlook VBA.

## Установка

1. Открой Outlook → ALT+F11 (VBA Editor).
2. Импортируй файлы из папки `src/`:
   - GitHubUpdater.bas
   - OutlookCoreModule.bas
   - OutlookHelper.cls
3. Вставь содержимое `ThisOutlookSession.txt` в `ThisOutlookSession` вручную.
4. Сохрани проект и перезапусти Outlook.

## Обновление

При каждом запуске Outlook выполняется проверка на GitHub и, если версия новее — обновление модулей.

## Состав

- GitHubUpdater.bas — загрузчик (не обновляется)
- OutlookCoreModule.bas — логика библиотеки
- OutlookHelper.cls — вспомогательный класс
- VERSION.txt — текущая версия
