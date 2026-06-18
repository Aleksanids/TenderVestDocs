# ETAP 32.2 — проверки

## Локальные команды

```powershell
$env:PYTHONUTF8='1'; python -m py_compile src\tendervestdocs\services\import_service.py src\tendervestdocs\ui\local_app.py
$env:PYTHONUTF8='1'; python -m pytest tests\test_import_service.py tests\test_ui_shell.py --basetemp=runtime\pytest_tmp_multi_excel_upload
$env:PYTHONUTF8='1'; python -m compileall -q src tests
```

## Результаты

- `py_compile` — passed.
- targeted pytest — `93 passed`.
- `compileall src tests` — passed.

## Не запускалось

- live-download;
- Playwright / Selenium;
- browser discovery;
- EXE / build / PyInstaller;
- full pytest;
- network smoke;
- OCR;
- ZIP rebuild.
