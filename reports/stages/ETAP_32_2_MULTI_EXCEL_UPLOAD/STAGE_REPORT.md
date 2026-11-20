# ETAP 32.2 — Multi Excel Upload

Статус: `ok`

## Итог

Загрузка двух Excel-реестров проверена и доведена до зеленого локального gate.

Подтверждено:

- один Excel листа 4 импортируется как `procurement`;
- один Excel листа 5 импортируется как `contract`;
- два отдельных Excel импортируются в единый `normalized_registry_v1`;
- `source_dataset_type`, `source_file`, `source_sheet`, `source_row_number` сохраняются;
- контракты не смешиваются с закупками;
- UI/API слой передает два файла через multi-file flow;
- пользовательский сценарий через `LocalAppState` проходит без browser automation.

## Исправления после validation

- `sheet6_check_data_export.py`: восстановлены helper-функции связи строк и сборка card-row для export-data, закрыт regression `_find_related_row` / `_find_related_contract_row` / `_export_data_table_row`.
- `local_app.py`: статус export-data в UI больше не называет карточные строки контрольными узлами.
- Локальные UI regression tests синхронизированы под multi-file wording в полном checkout.

## Реальные проверки

TEST fixtures:

- лист 4: `769` procurement / `0` contract;
- лист 5: `0` procurement / `100` contract;
- два файла: `769` procurement / `100` contract;
- combined workbook: `769` procurement / `100` contract.

Реальные Excel из Downloads, на temp-копиях:

- лист 4 `(1)`: `578` procurement / `0` contract;
- лист 5 `(2)`: `0` procurement / `47` contract;
- два файла: `578` procurement / `47` contract, total `625`.

User-role smoke:

- create project: ok;
- preview: `combined_workbook`, mapping modal не нужен;
- import: ok, rows `625`;
- карточка первой строки и действия доступны;
- top action скачивания доступен.

## Тесты

```text
python -X utf8 -m py_compile src\tendervestdocs\services\sheet6_check_data_export.py src\tendervestdocs\services\import_service.py src\tendervestdocs\ui\local_app.py
passed

python -X utf8 -m pytest tests\test_sheet6_check_data_export.py tests\test_ui_ux_contract.py -q --basetemp=D:\Codex\tmp_pytest_hotfix_32_2
35 passed

python -X utf8 -m pytest tests\test_import_service.py tests\test_ui_shell.py -q --basetemp=D:\Codex\tmp_pytest_multi_excel
93 passed

python -X utf8 -m pytest tests -k "excel or import or registry or workbook or contract" -q --basetemp=D:\Codex\tmp_pytest_multi_excel_regression
139 passed, 363 deselected

python -X utf8 -m compileall src tests
passed

python -X utf8 -m pytest tests -q --basetemp=D:\Codex\tmp_pytest_full
499 passed, 3 skipped
```

## Ограничения

- Browser automation / Playwright / Selenium не запускались.
- Live network / ЕИС / ЭТП не запускались.
- EXE / PyInstaller build не запускался.
- Пользовательские Excel не добавлялись в PR.
- ZIP не создавался.
