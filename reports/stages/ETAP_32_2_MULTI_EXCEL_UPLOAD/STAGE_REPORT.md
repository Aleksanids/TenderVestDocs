# ETAP 32.2 — загрузка двух Excel-таблиц листов 4 и 5

status: `ok_with_warnings`

## Цель

Добавить возможность выбрать две Excel-таблицы за один импорт: отдельный файл с листом 4 / реестром закупок и отдельный файл с листом 5 / реестром контрактов. После загрузки приложение должно работать как раньше: общий реестр, карточки строк, нормализованные JSON/CSV и дальнейшие действия.

## Что изменено

- `src/tendervestdocs/services/import_service.py` — добавлены `preview_registries(...)` и `import_registries(...)`, общий импорт нескольких файлов в один `normalized_registry_v1`, поля `source_files` и `copied_source_files` в результатах/отчетах.
- `src/tendervestdocs/ui/local_app.py` — `/api/import/preview` и `/api/import` принимают один или два upload-файла; добавлен state-level multi-file путь.
- `src/tendervestdocs/ui/static/index.html` — input загрузки поддерживает `multiple`, пользовательская подпись обновлена под два Excel.
- `src/tendervestdocs/ui/static/app.js` — добавлен `selectedFiles`, ограничение до двух файлов, отправка обоих файлов в `FormData`.

## Проверено локально

- `python -m py_compile src\tendervestdocs\services\import_service.py src\tendervestdocs\ui\local_app.py` — passed.
- `python -m pytest tests\test_import_service.py tests\test_ui_shell.py --basetemp=runtime\pytest_tmp_multi_excel_upload` — `93 passed`.
- `python -m compileall -q src tests` — passed.

## Что не запускалось

- live-download;
- Playwright / Selenium;
- browser discovery;
- EXE / build / PyInstaller;
- full pytest;
- network smoke;
- OCR;
- ZIP rebuild.

## Ограничения

- GitHub repo сейчас является компактным workflow/reporting контуром и не содержит полный source checkout приложения. В этот PR синхронизированы только измененные source-файлы и отчет этапа.
- Изменённые локальные тесты сохранены в рабочем root `D:\Codex\TenderVestDocs`, но не добавлены в этот compact GitHub repo, чтобы не запускать неполный pytest tree в CI.
- Multi-file mapping для нестандартных файлов не расширялся: новый путь рассчитан на распознанные Excel листов 4 и 5.

## Безопасность

- Исходные пользовательские Excel / CSV / DOCX / PDF / ZIP / HTML не изменялись.
- Пользовательские документы, скачанные документы, runtime/cache/temp/downloads и тяжелые архивы не добавлялись.
- Cookies / tokens / passwords / browser profiles / `.env` / credentials не читались.
- Сеть, ЕИС, ЭТП, browser automation, Selenium и Playwright не запускались.
- EXE / build / PyInstaller не создавались.
