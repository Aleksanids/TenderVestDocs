# TESTS — ETAP 32.2 Multi Excel Upload

## Passed

```text
python -X utf8 -m py_compile src\tendervestdocs\services\sheet6_check_data_export.py src\tendervestdocs\services\import_service.py src\tendervestdocs\ui\local_app.py
```

```text
python -X utf8 -m pytest tests\test_sheet6_check_data_export.py tests\test_ui_ux_contract.py -q --basetemp=D:\Codex\tmp_pytest_hotfix_32_2
35 passed
```

```text
python -X utf8 -m pytest tests\test_import_service.py tests\test_ui_shell.py -q --basetemp=D:\Codex\tmp_pytest_multi_excel
93 passed
```

```text
python -X utf8 -m pytest tests -k "excel or import or registry or workbook or contract" -q --basetemp=D:\Codex\tmp_pytest_multi_excel_regression
139 passed, 363 deselected
```

```text
python -X utf8 -m compileall src tests
passed
```

```text
python -X utf8 -m pytest tests -q --basetemp=D:\Codex\tmp_pytest_full
499 passed, 3 skipped
```

## Functional Smoke

- TEST fixtures two-file import: `769` procurement + `100` contract.
- Real Downloads Excel copies: `578` procurement + `47` contract.
- User-role LocalAppState scenario: ok, rows `625`, cards/actions ok.

## Not Run

- Browser automation / Playwright / Selenium.
- Live network / ЕИС / ЭТП.
- OCR.
- EXE / PyInstaller build.
