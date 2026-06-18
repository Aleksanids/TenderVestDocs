"""Small local HTTP app for the TenderVestDocs user-ready workflow."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from email import policy
from email.parser import BytesParser
import argparse
import json
import os
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
import re
import shutil
from typing import Any, Callable, Mapping, Sequence
from urllib.parse import parse_qs, urlparse

from tendervestdocs.core.exceptions import ProjectServiceError
from tendervestdocs.domain import normalize_procurement_rows
from tendervestdocs.domain.procurement import ProcurementRecord
from tendervestdocs.services import (
    AiReadyZipBuilderService,
    AiPackageMetadataService,
    CONTROLLED_DOWNLOAD_MANIFEST_NAME,
    ControlledDownloadService,
    DocumentInventoryService,
    EIS_44_SUPPLIER_RESULTS_FIXTURE,
    ImportPreviewResult,
    ImportService,
    ProjectService,
    ProcurementContractLinkService,
    ProtocolDataExtractionService,
    ProtocolInventoryService,
    ProtocolUiIntegrationService,
    NormativeUiIntegrationService,
    Sheet6CheckDataExportService,
    RegistryEnrichmentService,
    TextExtractionService,
    parse_supplier_results_html_file,
)
from tendervestdocs.services.import_service import (
    DATASET_CONTRACT,
    DATASET_PROCUREMENT,
    MAPPING_REQUIRED_ACTION,
    MAPPING_REQUIRED_STATUS,
    PROFILE_UNKNOWN_REQUIRES_MAPPING,
    _header_index_for_mapping,
    mapping_targets_json,
    validate_column_mapping,
)
from tendervestdocs.ui.shell import build_card_sections, build_empty_ui_shell_payload, build_ui_shell_payload


PROJECT_ROOT = Path(__file__).resolve().parents[3]
STATIC_DIR = Path(__file__).resolve().parent / "static"
PRIMARY_USER_PROJECTS_ROOT = Path(
    os.environ.get("TVD_USER_PROJECTS_ROOT", r"D:\!Личные файлы\Работа\AnalyticsProjects")
)
FALLBACK_USER_PROJECTS_ROOT = Path(
    os.environ.get("TVD_FALLBACK_PROJECTS_ROOT", r"D:\Codex\TenderVestDocs_USER_PROJECTS")
)
DEFAULT_PILOT_ROOT = FALLBACK_USER_PROJECTS_ROOT
DEFAULT_PROJECT_NAME = "Новый проект TenderVestDocs"
DOWNLOAD_FILTER_LIMIT = 30
SAFE_NAME_RE = re.compile(r'[\s\\/:*?"<>|\x00-\x1f]+')
FIELD_MAPPING_FIELDS = (
    ("procurement_number", "Номер"),
    ("law_type", "Закон"),
    ("source_url", "Источник / ссылка"),
    ("source_raw", "Источник / текст"),
    ("priority", "Приоритет"),
    ("nmck", "НМЦК"),
    ("okpd2", "ОКПД2"),
    ("ktru", "КТРУ"),
    ("customer_name", "Заказчик"),
    ("region", "Регион"),
    ("subject", "Предмет"),
)
FolderPickerProvider = Callable[[], Path | None]
FOLDER_PICKER_PROVIDER: FolderPickerProvider | None = None


def select_project_folder() -> Path | None:
    """Open a local Windows folder picker and return the selected project root."""

    if FOLDER_PICKER_PROVIDER is not None:
        selected = FOLDER_PICKER_PROVIDER()
        return selected.resolve() if selected is not None else None
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception as exc:  # pragma: no cover - depends on local desktop runtime
        raise RuntimeError("Диалог выбора папки недоступен в текущей среде.") from exc

    root = tk.Tk()
    root.withdraw()
    try:
        root.attributes("-topmost", True)
    except tk.TclError:
        pass
    try:
        selected = filedialog.askdirectory(
            parent=root,
            title="Выберите папку проекта TenderVestDocs",
            mustexist=True,
        )
    finally:
        root.destroy()
    if not selected:
        return None
    return Path(selected).resolve()


@dataclass
class LocalAppState:
    pilot_root: Path
    allow_network_downloads: bool = False
    allow_network_enrichment: bool = False
    project_service: ProjectService = field(default_factory=ProjectService)
    import_service: ImportService | None = None
    current_project_root: Path | None = None
    current_project_name: str = "Проект не выбран"
    current_records: list[ProcurementRecord] = field(default_factory=list)
    current_raw_rows: list[dict[str, Any]] = field(default_factory=list)
    payload: dict[str, Any] = field(default_factory=build_empty_ui_shell_payload)
    last_preview: ImportPreviewResult | None = None
    last_upload_path: Path | None = None
    last_upload_paths: tuple[Path, ...] = field(default_factory=tuple)
    last_field_mapping: dict[str, str] = field(default_factory=dict)
    last_import_session_id: str = ""
    last_import_result: Mapping[str, Any] | None = None
    last_download_manifest: Path | None = None
    last_document_inventory_summary: Mapping[str, Any] | None = None
    last_text_extraction_summary: Mapping[str, Any] | None = None
    last_ai_package_metadata_summary: Mapping[str, Any] | None = None
    last_ai_ready_zip_summary: Mapping[str, Any] | None = None
    ai_ready_zip_stage_runtime_root: Path | None = None
    last_enrichment_manifest: Path | None = None
    startup_project_warning: str = ""
    enriched_row_ids: set[str] = field(default_factory=set)
    enrichment_results_by_row_id: dict[str, dict[str, Any]] = field(default_factory=dict)
    normative_results_by_row_id: dict[str, dict[str, Any]] = field(default_factory=dict)
    last_normative_ui_integration_summary: Mapping[str, Any] | None = None
    last_sheet6_export_summary: Mapping[str, Any] | None = None

    def __post_init__(self) -> None:
        self.pilot_root = self.pilot_root.resolve()
        self.pilot_root.mkdir(parents=True, exist_ok=True)
        for folder in ("input_copies", "projects", "screenshots", "download_outputs", "browser_temp", "reports_tmp"):
            (self.pilot_root / folder).mkdir(parents=True, exist_ok=True)
        if self.import_service is None:
            self.import_service = ImportService(project_service=self.project_service)
        self._restore_last_project_on_start()
        self._set_project_context_on_payload()

    @property
    def upload_root(self) -> Path:
        path = self.pilot_root / "browser_temp" / "uploads"
        path.mkdir(parents=True, exist_ok=True)
        return path

    @property
    def pilot_projects_root(self) -> Path:
        path = self.pilot_root / "projects"
        path.mkdir(parents=True, exist_ok=True)
        return path

    @property
    def default_user_projects_root(self) -> Path:
        if PRIMARY_USER_PROJECTS_ROOT.exists() and PRIMARY_USER_PROJECTS_ROOT.is_dir():
            path = PRIMARY_USER_PROJECTS_ROOT / "TenderVestDocs_Local_App" / "projects"
            path.mkdir(parents=True, exist_ok=True)
            return path.resolve()
        FALLBACK_USER_PROJECTS_ROOT.mkdir(parents=True, exist_ok=True)
        return FALLBACK_USER_PROJECTS_ROOT.resolve()

    def create_project(
        self,
        project_name: str = DEFAULT_PROJECT_NAME,
        *,
        user_root: Path | None = None,
    ) -> dict[str, Any]:
        had_visible_registry = bool(self.current_records)
        previous_payload = self.payload
        base_root = (user_root.resolve() if user_root else self.default_user_projects_root)
        project = self.project_service.create_project(base_root, project_name)
        return self._finish_project_creation(project, previous_payload, had_visible_registry)

    def create_project_workspace(
        self,
        project_name: str = DEFAULT_PROJECT_NAME,
        *,
        parent_root: Path,
    ) -> dict[str, Any]:
        had_visible_registry = bool(self.current_records)
        previous_payload = self.payload
        project = self.project_service.create_project_workspace(parent_root.resolve(), project_name)
        return self._finish_project_creation(project, previous_payload, had_visible_registry)

    def _finish_project_creation(
        self,
        project: Any,
        previous_payload: dict[str, Any],
        had_visible_registry: bool,
    ) -> dict[str, Any]:
        self.project_service.remember_last_project(project.project_root)
        self.current_project_root = project.project_root
        self.current_project_name = project.project_name
        if had_visible_registry and self.last_upload_path is not None:
            try:
                saved = self._save_current_registry_to_project()
                saved["message"] = "Проект создан, ранее загруженный реестр сохранен в проект."
                return saved
            except Exception as exc:
                self.payload = previous_payload
                self.payload["stage"] = (
                    "Проект создан. Реестр сохранен на экране, но нормализованный реестр не удалось "
                    f"сохранить автоматически: {exc}"
                )
        else:
            self.current_records = []
            self.current_raw_rows = []
            self.enriched_row_ids = set()
            self.enrichment_results_by_row_id = {}
            self.normative_results_by_row_id = {}
            self.last_enrichment_manifest = None
            self.last_normative_ui_integration_summary = None
            self.payload = build_empty_ui_shell_payload()
            self.payload["stage"] = "Проект создан. Выберите табличный файл."
        self._set_project_context_on_payload()
        self._apply_download_project_guard()
        return {
            "status": "ok",
            "project_root": str(project.project_root),
            "project_name": project.project_name,
            "default_project_root": str(self.default_user_projects_root),
            "payload": self.payload,
        }

    def open_project(self, project_root: Path | None = None) -> dict[str, Any]:
        selected_root = project_root or self.current_project_root
        if selected_root is None:
            restored = self.project_service.restore_last_project()
            if restored.status != "ok" or restored.project is None:
                message = restored.message or "Пользовательский проект не найден. Создайте проект или укажите путь."
                self.payload["stage"] = message
                self._set_project_context_on_payload()
                return {
                    "status": "blocked",
                    "message": message,
                    "default_project_root": str(self.default_user_projects_root),
                    "last_project_root": str(restored.last_project_root or ""),
                    "payload": self.payload,
                }
            project = restored.project
        else:
            project = self.project_service.open_project(selected_root)
            self.project_service.remember_last_project(project.project_root)
        self.current_project_root = project.project_root
        self.current_project_name = project.project_name
        self.current_records = []
        self.current_raw_rows = []
        self.enriched_row_ids = set()
        self.enrichment_results_by_row_id = {}
        self.normative_results_by_row_id = {}
        self.last_enrichment_manifest = None
        self.last_normative_ui_integration_summary = None
        self.payload = build_empty_ui_shell_payload()
        self.payload["stage"] = "Проект открыт. Выберите табличный файл."
        self._set_project_context_on_payload()
        self._apply_download_project_guard()
        return {
            "status": "ok",
            "project_root": str(project.project_root),
            "project_name": project.project_name,
            "default_project_root": str(self.default_user_projects_root),
            "payload": self.payload,
        }

    def preview_import(self, upload_path: Path, *, sheet_name: str | None = None) -> dict[str, Any]:
        assert self.import_service is not None
        preview = self.import_service.preview_registry(upload_path, sheet_name=sheet_name)
        dataset_type = _dataset_type_from_preview(preview)
        self.last_upload_path = upload_path
        self.last_upload_paths = (upload_path,)
        self.last_import_session_id = _import_session_id(upload_path)
        self.last_preview = preview
        self.payload["import_preview"] = preview.to_json_dict()
        self.payload["field_mapping_fields"] = mapping_targets_json(dataset_type)
        if preview.manual_mapping_required:
            self.payload["import_mapping"] = self._build_mapping_payload(
                preview,
                upload_path,
                dataset_type=dataset_type,
            )
            self.payload["mapping_modal_open"] = True
        else:
            self.payload.pop("import_mapping", None)
            self.payload["mapping_modal_open"] = False
        self.payload["stage"] = (
            "Файл распознан как: Лист 4 — реестр закупок товаров."
            if preview.file_profile == "Стандартный лист 4"
            else "Файл выбран. Проверьте сопоставление колонок."
        )
        self._set_project_context_on_payload()
        return {
            "status": "ok",
            "preview": preview.to_json_dict(),
            "field_mapping_fields": mapping_targets_json(dataset_type),
            "import_mapping": self.payload.get("import_mapping"),
            "payload": self.payload,
        }

    def preview_imports(self, upload_paths: Sequence[Path]) -> dict[str, Any]:
        paths = tuple(Path(path) for path in upload_paths)
        if len(paths) == 1:
            return self.preview_import(paths[0])
        assert self.import_service is not None
        preview = self.import_service.preview_registries(paths)
        dataset_type = _dataset_type_from_preview(preview)
        self.last_upload_path = paths[0] if paths else None
        self.last_upload_paths = paths
        self.last_import_session_id = _import_session_id_for_paths(paths)
        self.last_preview = preview
        self.payload["import_preview"] = preview.to_json_dict()
        self.payload["field_mapping_fields"] = mapping_targets_json(dataset_type)
        if preview.manual_mapping_required:
            self.payload["import_mapping"] = self._build_mapping_payload(
                preview,
                paths[0],
                dataset_type=dataset_type,
            )
            self.payload["mapping_modal_open"] = True
        else:
            self.payload.pop("import_mapping", None)
            self.payload["mapping_modal_open"] = False
        self.payload["stage"] = (
            "Файлы распознаны как реестр закупок и реестр контрактов."
            if preview.profile_code == "combined_workbook" and not preview.manual_mapping_required
            else "Файлы выбраны. Проверьте сопоставление колонок."
        )
        self._set_project_context_on_payload()
        return {
            "status": "ok",
            "preview": preview.to_json_dict(),
            "field_mapping_fields": mapping_targets_json(dataset_type),
            "import_mapping": self.payload.get("import_mapping"),
            "payload": self.payload,
        }

    def import_uploaded(
        self,
        upload_path: Path,
        *,
        sheet_name: str | None = None,
        field_mapping: Mapping[str, str] | None = None,
    ) -> dict[str, Any]:
        if self.current_project_root is None:
            return self._import_without_project(
                upload_path,
                sheet_name=sheet_name,
                field_mapping=field_mapping,
            )
        assert self.import_service is not None
        preview = self.import_service.preview_registry(upload_path, sheet_name=sheet_name)
        self.last_upload_path = upload_path
        self.last_upload_paths = (upload_path,)
        self.last_import_session_id = _import_session_id(upload_path)
        self.last_preview = preview
        clean_mapping = _clean_mapping(field_mapping)
        self.last_field_mapping = dict(clean_mapping)
        dataset_type = _dataset_type_from_mapping(
            clean_mapping,
            fallback=_dataset_type_from_preview(preview),
        )
        if preview.manual_mapping_required and not clean_mapping:
            reusable = self.import_service.find_reusable_column_mapping(
                self.current_project_root,
                dataset_type=dataset_type,
                source_columns=preview.columns,
            )
            if reusable:
                clean_mapping = reusable
                dataset_type = _dataset_type_from_mapping(reusable, fallback=dataset_type)
            else:
                return self._mapping_required_response(preview, upload_path)

        if clean_mapping:
            validation_errors = validate_column_mapping(dataset_type, clean_mapping)
            if validation_errors:
                return self._mapping_required_response(
                    preview,
                    upload_path,
                    field_mapping=clean_mapping,
                    validation_errors=validation_errors,
                )
            self.import_service.save_column_mapping(
                self.current_project_root,
                dataset_type=dataset_type,
                field_mapping=clean_mapping,
                source_columns=preview.columns,
                sample_values=preview.sample_values,
                source_file=upload_path,
                sheet_name=preview.sheet_name,
                header_row=preview.header_row,
            )

        result = self.import_service.import_registry(
            self.current_project_root,
            upload_path,
            sheet_name=sheet_name,
            field_mapping=clean_mapping,
            source_dataset_type=dataset_type,
        )
        records = normalize_procurement_rows(result.rows)
        self.current_records = records
        self.current_raw_rows = _rows_to_json_dicts(result.rows)
        self.enriched_row_ids = set()
        self.enrichment_results_by_row_id = {}
        self.normative_results_by_row_id = {}
        self.last_enrichment_manifest = None
        self.last_normative_ui_integration_summary = None
        payload = build_ui_shell_payload(
            list(result.rows),
            project_name=self.current_project_name,
            project_root=str(self.current_project_root),
            stage="Реестр импортирован. Выберите строку для проверки карточки и скачивания.",
        )
        payload["import_result"] = _import_result_summary(result)
        payload["last_loaded_file_name"] = upload_path.name
        payload["last_imported_at"] = _now_label()
        self.payload = payload
        self._set_project_context_on_payload()
        self._merge_contract_link_cards_from_existing_outputs()
        self._apply_download_project_guard()
        self.last_import_result = payload["import_result"]
        return {
            "status": result.status,
            "import_result": payload["import_result"],
            "payload": payload,
        }

    def import_uploaded_files(
        self,
        upload_paths: Sequence[Path],
        *,
        field_mapping: Mapping[str, str] | None = None,
    ) -> dict[str, Any]:
        paths = tuple(Path(path) for path in upload_paths)
        if len(paths) == 1:
            return self.import_uploaded(paths[0], field_mapping=field_mapping)
        if self.current_project_root is None:
            return {
                "status": "project_required",
                "message": "Для импорта двух Excel сначала создайте или откройте проект.",
                "payload": self.payload,
            }
        assert self.import_service is not None
        preview = self.import_service.preview_registries(paths)
        self.last_upload_path = paths[0] if paths else None
        self.last_upload_paths = paths
        self.last_import_session_id = _import_session_id_for_paths(paths)
        self.last_preview = preview
        clean_mapping = _clean_mapping(field_mapping)
        self.last_field_mapping = dict(clean_mapping)
        if preview.manual_mapping_required:
            return self._mapping_required_response(preview, paths[0])

        result = self.import_service.import_registries(
            self.current_project_root,
            paths,
        )
        records = normalize_procurement_rows(result.rows)
        self.current_records = records
        self.current_raw_rows = _rows_to_json_dicts(result.rows)
        self.enriched_row_ids = set()
        self.enrichment_results_by_row_id = {}
        self.normative_results_by_row_id = {}
        self.last_enrichment_manifest = None
        self.last_normative_ui_integration_summary = None
        payload = build_ui_shell_payload(
            list(result.rows),
            project_name=self.current_project_name,
            project_root=str(self.current_project_root),
            stage="Реестры импортированы. Выберите строку для проверки карточки и скачивания.",
        )
        payload["import_result"] = _import_result_summary(result)
        payload["last_loaded_file_name"] = _file_list_label(paths)
        payload["last_imported_at"] = _now_label()
        self.payload = payload
        self._set_project_context_on_payload()
        self._merge_contract_link_cards_from_existing_outputs()
        self._apply_download_project_guard()
        self.last_import_result = payload["import_result"]
        return {
            "status": result.status,
            "import_result": payload["import_result"],
            "payload": payload,
        }

    def _import_without_project(
        self,
        upload_path: Path,
        *,
        sheet_name: str | None = None,
        field_mapping: Mapping[str, str] | None = None,
    ) -> dict[str, Any]:
        """Build a visible registry payload without creating a user project or writing artifacts."""

        assert self.import_service is not None
        preview = self.import_service.preview_registry(upload_path, sheet_name=sheet_name)
        self.last_upload_path = upload_path
        self.last_upload_paths = (upload_path,)
        self.last_import_session_id = _import_session_id(upload_path)
        self.last_preview = preview
        clean_mapping = _clean_mapping(field_mapping)
        self.last_field_mapping = dict(clean_mapping)
        dataset_type = _dataset_type_from_mapping(
            clean_mapping,
            fallback=_dataset_type_from_preview(preview),
        )
        if preview.manual_mapping_required and not clean_mapping:
            return self._mapping_required_response(preview, upload_path)
        if clean_mapping:
            validation_errors = validate_column_mapping(dataset_type, clean_mapping)
            if validation_errors:
                return self._mapping_required_response(
                    preview,
                    upload_path,
                    field_mapping=clean_mapping,
                    validation_errors=validation_errors,
                )

        normalized_rows = self._read_rows_for_temporary_view(
            upload_path,
            sheet_name=sheet_name,
            field_mapping=clean_mapping,
            dataset_type=dataset_type,
        )
        self.current_records = normalize_procurement_rows(normalized_rows)
        self.current_raw_rows = _rows_to_json_dicts(normalized_rows)
        self.enriched_row_ids = set()
        self.enrichment_results_by_row_id = {}
        self.normative_results_by_row_id = {}
        self.last_enrichment_manifest = None
        self.last_normative_ui_integration_summary = None
        payload = build_ui_shell_payload(
            self.current_raw_rows,
            project_name="Проект не выбран",
            project_root="",
            stage="Временная работа без проекта. Реестр загружен для просмотра.",
        )
        payload["import_result"] = {
            "status": "temporary_view",
            "source_file": str(upload_path),
            "copied_source_file": "",
            "normalized_registry_json_path": "",
            "normalized_registry_csv_path": "",
            "normalization_report_json_path": "",
            "normalization_report_md_path": "",
            "sheet_name": preview.sheet_name,
            "file_profile": preview.profile_code,
            "file_profile_label": preview.file_profile,
            "sheet_names": [
                str(profile.get("sheet_name") or "")
                for profile in preview.sheet_profiles
            ] or ([preview.sheet_name] if preview.sheet_name else []),
            "dataset_counts": _count_records_by_dataset(self.current_records),
            "sheet_profiles": [],
            "header_row": preview.header_row,
            "rows_count": len(self.current_records),
            "physical_data_rows_count": preview.physical_rows_count,
            "skipped_empty_rows_count": preview.skipped_rows_count,
            "columns_detected": dict(preview.columns_detected),
            "missing_required_columns": list(preview.missing_required_columns),
            "warnings": [
                {
                    "code": "temporary_view_without_project",
                    "message": "Реестр показан без проекта; файлы нормализованного реестра не записывались.",
                }
            ],
            "errors": [],
        }
        payload["last_loaded_file_name"] = upload_path.name
        payload["last_imported_at"] = _now_label()
        payload["project_status"] = "Проект не создан"
        payload["registry_status"] = "Загружен"
        payload["project_warning"] = "Для скачивания документов создайте или откройте проект."
        self.payload = payload
        self.last_import_result = payload["import_result"]
        self._set_project_context_on_payload()
        self._merge_contract_link_cards_from_existing_outputs()
        self._apply_download_project_guard()
        return {
            "status": "ok",
            "message": "Реестр загружен для временного просмотра без проекта.",
            "import_result": payload["import_result"],
            "payload": self.payload,
        }

    def _read_rows_for_temporary_view(
        self,
        upload_path: Path,
        *,
        sheet_name: str | None,
        field_mapping: Mapping[str, str],
        dataset_type: str,
    ) -> list[Any]:
        assert self.import_service is not None
        selected_tables, _profile_code, _profile_label = self.import_service._select_import_tables(
            self.import_service._read_tabular_tables(upload_path, sheet_name=sheet_name)
        )
        apply_mapping = field_mapping if len(selected_tables) == 1 else {}
        normalized_rows: list[Any] = []
        for selected in selected_tables:
            header_index = selected.header_index
            columns_by_index = dict(selected.columns_by_index)
            columns_detected = dict(selected.columns_detected)
            effective_dataset_type = selected.source_dataset_type
            if selected.profile_code == PROFILE_UNKNOWN_REQUIRES_MAPPING and apply_mapping:
                effective_dataset_type = dataset_type
            if apply_mapping:
                header_index = _header_index_for_mapping(
                    selected.tabular.rows,
                    apply_mapping,
                    default=header_index,
                )
                columns_by_index, columns_detected, _mapping_warnings = self.import_service._apply_field_mapping(
                    rows=selected.tabular.rows,
                    header_index=header_index,
                    columns_by_index=columns_by_index,
                    columns_detected=columns_detected,
                    field_mapping=apply_mapping,
                )
            table_rows, _skipped = self.import_service._build_rows(
                rows=selected.tabular.rows,
                header_index=header_index,
                columns_by_index=columns_by_index,
                source_type=selected.tabular.source_type,
                source_dataset_type=effective_dataset_type,
                source_name=upload_path.name,
                source_sheet=selected.tabular.sheet_name or "",
                row_id_offset=len(normalized_rows),
            )
            normalized_rows.extend(table_rows)
        return normalized_rows

    def _save_current_registry_to_project(self) -> dict[str, Any]:
        if self.current_project_root is None or self.last_upload_path is None:
            self._set_project_context_on_payload()
            return {
                "status": "ok",
                "project_root": str(self.current_project_root or ""),
                "project_name": self.current_project_name,
                "payload": self.payload,
            }
        assert self.import_service is not None
        dataset_type = _dataset_type_from_mapping(
            self.last_field_mapping,
            fallback=_dataset_type_from_preview(self.last_preview),
        )
        result = self.import_service.import_registry(
            self.current_project_root,
            self.last_upload_path,
            field_mapping=self.last_field_mapping,
            source_dataset_type=dataset_type,
        )
        records = normalize_procurement_rows(result.rows)
        self.current_records = records
        self.current_raw_rows = _rows_to_json_dicts(result.rows)
        payload = build_ui_shell_payload(
            list(result.rows),
            project_name=self.current_project_name,
            project_root=str(self.current_project_root),
            stage="Проект создан. Ранее загруженный реестр сохранен в проект.",
            sources_enriched=bool(self.enriched_row_ids),
            enriched_row_ids=set(self.enriched_row_ids),
            source_refreshed_at=self.payload.get("source_refreshed_at", ""),
            enrichment_results=self.enrichment_results_by_row_id,
            normative_checked=self._normative_checked_for_export(),
        )
        payload["import_result"] = _import_result_summary(result)
        payload["last_loaded_file_name"] = self.last_upload_path.name
        payload["last_imported_at"] = _now_label()
        self.payload = payload
        self.last_import_result = payload["import_result"]
        self._set_project_context_on_payload()
        self._apply_download_project_guard()
        return {
            "status": result.status,
            "project_root": str(self.current_project_root),
            "project_name": self.current_project_name,
            "default_project_root": str(self.default_user_projects_root),
            "import_result": payload["import_result"],
            "payload": self.payload,
        }

    def save_current_mapping(
        self,
        field_mapping: Mapping[str, str],
        *,
        dataset_type: str = DATASET_PROCUREMENT,
    ) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Сначала создайте или откройте проект.",
                "payload": self.payload,
            }
        if self.last_preview is None or self.last_upload_path is None:
            return {
                "status": "blocked",
                "message": "Сначала выберите и проверьте табличный файл.",
                "payload": self.payload,
            }
        assert self.import_service is not None
        clean_mapping = _clean_mapping(field_mapping)
        validation_errors = validate_column_mapping(dataset_type, clean_mapping)
        if validation_errors:
            return self._mapping_required_response(
                self.last_preview,
                self.last_upload_path,
                field_mapping=clean_mapping,
                validation_errors=validation_errors,
            )
        saved = self.import_service.save_column_mapping(
            self.current_project_root,
            dataset_type=dataset_type,
            field_mapping=clean_mapping,
            source_columns=self.last_preview.columns,
            sample_values=self.last_preview.sample_values,
            source_file=self.last_upload_path,
            sheet_name=self.last_preview.sheet_name,
            header_row=self.last_preview.header_row,
        )
        self.payload["stage"] = "Сопоставление колонок сохранено. Можно повторить импорт."
        self.payload["import_mapping"] = self._build_mapping_payload(
            self.last_preview,
            self.last_upload_path,
            field_mapping=clean_mapping,
            dataset_type=dataset_type,
        )
        self.payload["mapping_modal_open"] = True
        self._set_project_context_on_payload()
        return {
            "status": "ok",
            "message": "Сопоставление колонок сохранено.",
            "mapping": saved,
            "payload": self.payload,
        }

    def reset_column_mappings(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Сначала создайте или откройте проект.",
                "payload": self.payload,
            }
        assert self.import_service is not None
        settings_path = self.import_service.clear_column_mappings(self.current_project_root)
        self.payload["stage"] = "Сохраненные сопоставления сброшены."
        self._set_project_context_on_payload()
        return {
            "status": "ok",
            "message": "Сохраненные сопоставления сброшены.",
            "settings_path": str(settings_path),
            "payload": self.payload,
        }

    def open_link(self, row_id: str) -> dict[str, Any]:
        raw_row = next((row for row in self.current_raw_rows if _text(row.get("row_id")) == row_id), None)
        if raw_row is not None:
            source_url = (
                _text(raw_row.get("contract_url"))
                or _text(raw_row.get("source_url"))
                or _text(raw_row.get("normalized_source_url"))
                or _text(raw_row.get("generated_eis_url"))
            )
            if source_url:
                return {"status": "ok", "message": "Ссылка готова к открытию.", "url": source_url}
        row = _record_by_row_id(self.current_records, row_id)
        if row is None:
            return {
                "status": "blocked",
                "message": "В выбранной строке нет публичной ссылки.",
                "url": "",
            }
        row_data = row.to_json_dict() if hasattr(row, "to_json_dict") else {}
        source_url = (
            _text(row_data.get("contract_url"))
            or _text(row.source.source_url)
            or _text(row_data.get("source_url"))
        )
        if not source_url:
            return {
                "status": "blocked",
                "message": "В выбранной строке нет публичной ссылки.",
                "url": "",
            }
        return {"status": "ok", "message": "Ссылка готова к открытию.", "url": source_url}

    def download_rows(self, row_ids: list[str], *, scope_label: str) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "error",
                "error_code": "project_required",
                "project_required": True,
                "message": "Для скачивания документов создайте или откройте проект.",
                "payload": self.payload,
            }
        selected_records = [
            record for record in self.current_records if not row_ids or record.row_id in set(row_ids)
        ]
        if not selected_records:
            return {"status": "blocked", "message": "Нет строк для скачивания."}
        skipped_by_limit = max(0, len(selected_records) - DOWNLOAD_FILTER_LIMIT)
        selected_records = selected_records[:DOWNLOAD_FILTER_LIMIT]
        service = ControlledDownloadService(
            self.current_project_root,
            project_service=self.project_service,
            max_rows=DOWNLOAD_FILTER_LIMIT,
        )
        result = service.download_records(
            selected_records,
            scope_label=scope_label,
            allow_network=self.allow_network_downloads,
        )
        self.last_download_manifest = result.manifest_path
        self._merge_download_result_into_payload(result)
        self._apply_download_project_guard()
        documents_found_count = sum(_safe_row_documents_found_count(row) for row in result.rows)
        documents_downloaded_count = _safe_result_downloaded_files_count(result)
        documents_skipped_count = max(0, documents_found_count - documents_downloaded_count)
        downloaded_files = _downloaded_files_payload(result.rows)
        issues = _download_issues_payload(result.rows)
        return {
            "status": result.status,
            "downloaded_rows_count": result.downloaded_rows_count,
            "downloaded_files_count": documents_downloaded_count,
            "documents_found_count": documents_found_count,
            "documents_downloaded_count": documents_downloaded_count,
            "documents_skipped_count": documents_skipped_count,
            "downloaded_files": downloaded_files,
            "saved_files": downloaded_files,
            "warnings": issues["warnings"],
            "errors": issues["errors"],
            "selected_rows_count": len(selected_records),
            "skipped_by_limit": skipped_by_limit,
            "manifest_path": str(result.manifest_path),
            "payload": self.payload,
        }

    def preview_protocol_results(self, row_ids: list[str], *, scope_label: str) -> dict[str, Any]:
        fixture = PROJECT_ROOT / "TEST" / "fixtures" / EIS_44_SUPPLIER_RESULTS_FIXTURE
        if not fixture.is_file():
            return {
                "status": "error",
                "error_code": "preview_fixture_missing",
                "message": "Локальный HTML-источник результатов протокола не найден.",
                "project_required": False,
                "payload": self.payload,
            }

        session_id = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        temp_root = (self.pilot_root / "protocol_results_preview" / session_id).resolve()
        temp_root.mkdir(parents=True, exist_ok=True)
        temp_file = temp_root / fixture.name
        created_count = 0
        deleted_count = 0
        cleanup_warnings: list[str] = []
        parsed: dict[str, list[dict[str, Any]]] = {
            "supplier_result_evidence": [],
            "embedded_contract_evidence": [],
        }
        status = "ok"
        message = "Результаты протокола получены во временном режиме."
        try:
            shutil.copy2(fixture, temp_file)
            created_count = 1
            parsed = parse_supplier_results_html_file(
                temp_file,
                evidence_storage_status="temp_download_parsed_deleted",
            )
        except Exception as exc:
            status = "error"
            message = f"Временный разбор результатов протокола не выполнен: {exc}"
            cleanup_warnings.append(f"temp_download_parse_failed:{type(exc).__name__}")
        finally:
            if temp_file.exists():
                try:
                    temp_file.unlink()
                    deleted_count += 1
                except OSError as exc:
                    cleanup_warnings.append(f"temp_file_cleanup_failed:{exc}")
            try:
                temp_root.rmdir()
            except OSError:
                if temp_root.exists():
                    cleanup_warnings.append("temp_session_folder_not_empty_after_cleanup")

        parsed_summaries = self._build_protocol_preview_summaries(parsed)
        summaries = self._preview_summaries_for_requested_rows(parsed_summaries, row_ids)
        if summaries:
            self._merge_protocol_ui_cards(summaries)
        self.payload["protocol_results_preview"] = {
            "status": status,
            "storage_mode": "temp_parse",
            "scope_label": scope_label,
            "requested_row_ids": list(row_ids),
            "project_required": False,
            "temp_files_created_count": created_count,
            "temp_files_deleted_count": deleted_count,
            "runtime_temp_root": str(temp_root.parent),
            "temp_session_path": str(temp_root),
            "temp_session_exists_after_cleanup": temp_root.exists(),
            "protocols_parsed_count": len(summaries),
            "persistent_documents_created": False,
            "document_inventory_updated": False,
            "exports_created": False,
            "ai_ready_artifacts_created": False,
            "live_network_used": False,
            "warnings": cleanup_warnings,
        }
        self.payload["stage"] = (
            "Результаты протокола получены во временном режиме; файлы удалены после разбора."
            if status == "ok"
            else message
        )
        self._set_project_context_on_payload()
        return {
            "status": status,
            "message": message,
            "storage_mode": "temp_parse",
            "scope_label": scope_label,
            "requested_row_ids": list(row_ids),
            "selected_rows_count": len(row_ids),
            "project_required": False,
            "temp_files_created_count": created_count,
            "temp_files_deleted_count": deleted_count,
            "temp_session_path": str(temp_root),
            "temp_session_exists_after_cleanup": temp_root.exists(),
            "protocols_parsed_count": len(summaries),
            "protocol_ui_summaries": summaries,
            "parsed_protocol_ui_summaries": parsed_summaries,
            "supplier_result_evidence": parsed.get("supplier_result_evidence", []),
            "embedded_contract_evidence": parsed.get("embedded_contract_evidence", []),
            "persistent_documents_created": False,
            "document_inventory_updated": False,
            "exports_created": False,
            "ai_ready_artifacts_created": False,
            "live_network_used": False,
            "warnings": cleanup_warnings,
            "payload": self.payload,
        }

    def refresh_document_inventory(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Для инвентаризации документов создайте или откройте проект.",
                "payload": self.payload,
            }
        records = self._current_rows_for_refresh()
        service = DocumentInventoryService(
            self.current_project_root,
            project_service=self.project_service,
        )
        result = service.refresh_inventory(records if records else None)
        summary = dict(result.get("summary") or {})
        self.last_document_inventory_summary = summary
        self.payload["document_inventory"] = {
            "status": result.get("status"),
            "summary": summary,
            "inventory_root": result.get("inventory_root"),
            "paths": result.get("paths"),
        }
        self.payload["stage"] = "Инвентаризация документов обновлена."
        self._set_project_context_on_payload()
        return {
            "status": result.get("status"),
            "message": "Инвентаризация документов обновлена.",
            "summary": summary,
            "inventory_root": result.get("inventory_root"),
            "paths": result.get("paths"),
            "payload": self.payload,
        }

    def get_document_inventory_summary(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "payload": self.payload,
            }
        service = DocumentInventoryService(
            self.current_project_root,
            project_service=self.project_service,
        )
        summary = service.load_summary()
        return {
            "status": summary.get("status", "ok"),
            "summary": summary,
            "payload": self.payload,
        }

    def get_document_inventory_for_row(self, row_id: str) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "row_id": row_id,
                "items": [],
                "payload": self.payload,
            }
        service = DocumentInventoryService(
            self.current_project_root,
            project_service=self.project_service,
        )
        result = service.load_row_inventory(row_id)
        result["payload"] = self.payload
        return result

    def refresh_text_extraction(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Для извлечения текста создайте или откройте проект.",
                "payload": self.payload,
            }
        service = TextExtractionService(
            self.current_project_root,
            project_service=self.project_service,
        )
        result = service.refresh_text_extraction()
        summary = dict(result.get("summary") or {})
        self.last_text_extraction_summary = summary
        self.payload["text_extraction"] = {
            "status": result.get("status"),
            "summary": summary,
            "paths": result.get("paths"),
        }
        self.payload["stage"] = "Извлечение текста и фрагментов обновлено."
        self._set_project_context_on_payload()
        return {
            "status": result.get("status"),
            "message": "Извлечение текста и фрагментов обновлено.",
            "summary": summary,
            "paths": result.get("paths"),
            "payload": self.payload,
        }

    def get_text_extraction_summary(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "payload": self.payload,
            }
        service = TextExtractionService(
            self.current_project_root,
            project_service=self.project_service,
        )
        summary = service.load_summary()
        return {
            "status": summary.get("status", "ok"),
            "summary": summary,
            "payload": self.payload,
        }

    def get_extracted_text_for_document(self, document_id: str) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "document_id": document_id,
                "payload": self.payload,
            }
        service = TextExtractionService(
            self.current_project_root,
            project_service=self.project_service,
        )
        result = service.load_document_text(document_id)
        result["payload"] = self.payload
        return result

    def get_fragments_for_row(self, row_id: str) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "row_id": row_id,
                "fragments": [],
                "payload": self.payload,
            }
        service = TextExtractionService(
            self.current_project_root,
            project_service=self.project_service,
        )
        result = service.load_fragments_for_row(row_id)
        result["payload"] = self.payload
        return result

    def get_fragments_for_document(self, document_id: str) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "document_id": document_id,
                "fragments": [],
                "payload": self.payload,
            }
        service = TextExtractionService(
            self.current_project_root,
            project_service=self.project_service,
        )
        result = service.load_fragments_for_document(document_id)
        result["payload"] = self.payload
        return result

    def refresh_protocol_inventory(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Для инвентаризации протоколов создайте или откройте проект.",
                "payload": self.payload,
            }
        service = ProtocolInventoryService(
            self.current_project_root,
            project_service=self.project_service,
        )
        result = service.build_protocol_inventory()
        summary = dict(result.get("summary") or {})
        self.payload["protocol_inventory"] = {
            "status": result.get("status"),
            "summary": summary,
            "paths": result.get("paths"),
        }
        self.payload["stage"] = "Инвентаризация протоколов ETAP 25 обновлена."
        self._set_project_context_on_payload()
        return {
            "status": result.get("status"),
            "message": "Инвентаризация протоколов ETAP 25 обновлена.",
            "summary": summary,
            "paths": result.get("paths"),
            "payload": self.payload,
        }

    def get_protocol_inventory(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "items": [],
                "payload": self.payload,
            }
        service = ProtocolInventoryService(
            self.current_project_root,
            project_service=self.project_service,
        )
        result = service.get_protocol_inventory()
        result["payload"] = self.payload
        return result

    def get_protocol_inventory_summary(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "payload": self.payload,
            }
        service = ProtocolInventoryService(
            self.current_project_root,
            project_service=self.project_service,
        )
        summary = service.get_protocol_summary()
        return {
            "status": summary.get("status", "ok"),
            "summary": summary,
            "payload": self.payload,
        }

    def check_protocol_source_coverage(
        self,
        row_ids: list[str],
        *,
        allow_network: bool = False,
        local_evidence_paths: list[str] | None = None,
    ) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "payload": self.payload,
            }
        rows_for_check = self._current_rows_for_refresh()
        by_id = {_text(row.get("row_id")): row for row in rows_for_check if _text(row.get("row_id"))}
        selected = [by_id[row_id] for row_id in row_ids if row_id in by_id]
        if not selected:
            selected = [str(row_id) for row_id in row_ids if str(row_id)]
        if not selected:
            selected = rows_for_check[:5]
        service = ProtocolInventoryService(
            self.current_project_root,
            project_service=self.project_service,
        )
        result = service.check_protocol_source_coverage(
            selected,
            allow_network=allow_network,
            local_evidence_paths=[Path(path) for path in (local_evidence_paths or [])],
        )
        self.payload["protocol_source_coverage"] = {
            "status": result.get("status"),
            "summary": result.get("summary"),
        }
        self.payload["stage"] = "Проверка покрытия источников протоколов ETAP 25 выполнена."
        self._set_project_context_on_payload()
        result["payload"] = self.payload
        return result

    def refresh_protocol_data(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Для извлечения данных протоколов создайте или откройте проект.",
                "payload": self.payload,
            }
        service = ProtocolDataExtractionService(
            self.current_project_root,
            project_service=self.project_service,
        )
        result = service.extract_protocol_data()
        summary = dict(result.get("summary") or {})
        self.payload["protocol_data"] = {
            "status": result.get("status"),
            "summary": summary,
            "paths": result.get("paths"),
        }
        self.payload["stage"] = "Извлечение данных протоколов ETAP 26 обновлено."
        self._set_project_context_on_payload()
        return {
            "status": result.get("status"),
            "message": "Извлечение данных протоколов ETAP 26 обновлено.",
            "summary": summary,
            "paths": result.get("paths"),
            "payload": self.payload,
        }

    def get_protocol_data(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "items": [],
                "payload": self.payload,
            }
        service = ProtocolDataExtractionService(
            self.current_project_root,
            project_service=self.project_service,
        )
        result = service.get_protocol_data()
        result["payload"] = self.payload
        return result

    def get_protocol_data_summary(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "payload": self.payload,
            }
        service = ProtocolDataExtractionService(
            self.current_project_root,
            project_service=self.project_service,
        )
        summary = service.get_protocol_extraction_summary()
        return {
            "status": summary.get("status", "ok"),
            "summary": summary,
            "payload": self.payload,
        }

    def get_protocol_participants(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "participants": [],
                "payload": self.payload,
            }
        service = ProtocolDataExtractionService(
            self.current_project_root,
            project_service=self.project_service,
        )
        result = service.get_protocol_participants()
        result["payload"] = self.payload
        return result

    def get_protocol_price_offers(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "price_offers": [],
                "payload": self.payload,
            }
        service = ProtocolDataExtractionService(
            self.current_project_root,
            project_service=self.project_service,
        )
        result = service.get_protocol_price_offers()
        result["payload"] = self.payload
        return result

    def get_protocol_results(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "results": [],
                "payload": self.payload,
            }
        service = ProtocolDataExtractionService(
            self.current_project_root,
            project_service=self.project_service,
        )
        result = service.get_protocol_results()
        result["payload"] = self.payload
        return result

    def refresh_protocol_ui_summary(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Для сводки протоколов создайте или откройте проект.",
                "payload": self.payload,
            }
        service = ProtocolUiIntegrationService(
            self.current_project_root,
            project_service=self.project_service,
        )
        result = service.build_protocol_ui_integration()
        summary = dict(result.get("summary") or {})
        self.payload["protocol_ui_summary"] = {
            "status": result.get("status"),
            "summary": summary,
            "paths": result.get("paths"),
        }
        self.payload["stage"] = "Сводка протоколов ETAP 27 обновлена."
        self._merge_protocol_ui_cards(result.get("protocol_ui_summaries"))
        self._set_project_context_on_payload()
        return {
            "status": result.get("status"),
            "message": "Сводка протоколов ETAP 27 обновлена.",
            "summary": summary,
            "paths": result.get("paths"),
            "payload": self.payload,
            "protocol_ui_summaries": result.get("protocol_ui_summaries", []),
        }

    def get_protocol_ui_summary(self, row_id: str | None = None) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "protocol_ui_summary": {},
                "payload": self.payload,
            }
        service = ProtocolUiIntegrationService(
            self.current_project_root,
            project_service=self.project_service,
        )
        result = service.get_protocol_ui_summary(row_id)
        if not row_id:
            self._merge_protocol_ui_cards(result.get("protocol_ui_summaries"))
        result["payload"] = self.payload
        return result

    def get_protocol_export(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "payload": self.payload,
            }
        service = ProtocolUiIntegrationService(
            self.current_project_root,
            project_service=self.project_service,
        )
        result = service.get_protocol_export()
        result["payload"] = self.payload
        return result

    def get_protocol_ai_ready_summary(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "payload": self.payload,
            }
        service = ProtocolUiIntegrationService(
            self.current_project_root,
            project_service=self.project_service,
        )
        result = service.get_protocol_ai_ready_summary()
        result["payload"] = self.payload
        return result

    def run_normative_check(self, row_ids: list[str], *, scope_label: str) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Для нормативной проверки создайте или откройте проект.",
                "payload": self.payload,
            }
        rows_for_check = self._current_rows_for_refresh()
        if not rows_for_check:
            return {
                "status": "blocked",
                "message": "Сначала загрузите табличный реестр.",
                "payload": self.payload,
            }
        available_ids = [_text(row.get("row_id")) for row in rows_for_check if _text(row.get("row_id"))]
        requested_ids = {str(item) for item in row_ids if str(item)}
        selected_ids = (set(available_ids) if not requested_ids else set(available_ids) & requested_ids)
        if not selected_ids:
            return {
                "status": "blocked",
                "message": "Нет строк для нормативной проверки.",
                "payload": self.payload,
            }

        service = NormativeUiIntegrationService(
            self.current_project_root,
            project_service=self.project_service,
        )
        result = service.build_normative_ui_integration(
            rows_for_check,
            row_ids=sorted(selected_ids),
            scope_label=scope_label,
        )
        self.last_normative_ui_integration_summary = result.get("summary", {})
        self.payload["normative_ui_export"] = {
            "status": result.get("status"),
            "summary": result.get("summary"),
            "paths": result.get("paths"),
            "sheet6_readiness": result.get("sheet6_readiness"),
            "ai_ready_zip_rebuilt": False,
        }
        self.payload["stage"] = f"Нормативная проверка выполнена: {len(selected_ids)} строк."
        self._merge_normative_ui_rows(result.get("ui_rows"))
        self._update_export_data_action_state()
        self._set_project_context_on_payload()
        self._apply_download_project_guard()
        return {
            "status": result.get("status"),
            "message": f"Нормативная проверка выполнена: {len(selected_ids)} строк.",
            "summary": result.get("summary"),
            "paths": result.get("paths"),
            "selected_rows_count": len(selected_ids),
            "total_rows_count": len(available_ids),
            "payload": self.payload,
            "ui_rows": result.get("ui_rows", []),
            "sheet6_readiness": result.get("sheet6_readiness"),
            "live_network_used": False,
            "documents_downloaded": False,
            "source_excel_changed": False,
            "ai_ready_zip_rebuilt": False,
        }

    def get_normative_export(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "payload": self.payload,
            }
        service = NormativeUiIntegrationService(
            self.current_project_root,
            project_service=self.project_service,
        )
        result = service.get_normative_export()
        result["payload"] = self.payload
        return result

    def get_normative_ai_ready_summary(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "payload": self.payload,
            }
        service = NormativeUiIntegrationService(
            self.current_project_root,
            project_service=self.project_service,
        )
        result = service.get_normative_ai_ready_summary()
        result["payload"] = self.payload
        return result

    def get_normative_sheet6_readiness(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "payload": self.payload,
                "excel_generated": False,
            }
        service = NormativeUiIntegrationService(
            self.current_project_root,
            project_service=self.project_service,
        )
        result = service.get_sheet6_readiness()
        result["payload"] = self.payload
        return result

    def run_sheet6_export(self, row_ids: list[str], *, scope_label: str) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Для экспорта данных создайте или откройте проект.",
                "payload": self.payload,
            }
        readiness = self._export_data_readiness_for_ui()
        if readiness["status"] not in {"export_data_ready", "export_data_ready_with_warnings"}:
            self._update_export_data_action_state()
            return {
                "status": readiness["status"],
                "export_data_status": readiness["status"],
                "message": readiness["message"],
                "payload": self.payload,
                "can_export_data": False,
            }
        rows_for_export = self._current_rows_for_refresh()
        available_ids = [_text(row.get("row_id")) for row in rows_for_export if _text(row.get("row_id"))]
        requested_ids = {str(item) for item in row_ids if str(item)}
        selected_ids = sorted(set(available_ids) if not requested_ids else set(available_ids) & requested_ids)

        service = Sheet6CheckDataExportService(
            self.current_project_root,
            project_service=self.project_service,
        )
        result = service.build_sheet6_export(
            rows_for_export,
            row_ids=selected_ids,
            scope_label=scope_label,
        )
        summary = result.get("summary") if isinstance(result.get("summary"), Mapping) else {}
        self.last_sheet6_export_summary = summary
        self.payload["sheet6_export"] = {
            "status": result.get("status"),
            "export_data_status": result.get("export_data_status") or result.get("status"),
            "summary": summary,
            "paths": result.get("paths"),
            "export_sheet_name": result.get("export_sheet_name"),
            "export_columns": result.get("export_columns"),
            "technical_fields": result.get("technical_fields"),
        }
        rows_count = int(summary.get("control_rows_count") or 0)
        warnings_count = int(summary.get("warnings_count") or 0)
        missing_count = int(summary.get("missing_inputs_count") or 0)
        result_paths = result.get("paths") if isinstance(result.get("paths"), Mapping) else {}
        xlsx_path = _text(result_paths.get("xlsx"))
        self.payload["stage"] = (
            f"Экспорт данных создан: {rows_count} строк; "
            f"warnings {warnings_count}; missing inputs {missing_count}. Файл: {xlsx_path}"
        )
        self._update_export_data_action_state()
        self._set_project_context_on_payload()
        self._apply_download_project_guard()
        return {
            "status": result.get("status"),
            "export_data_status": result.get("export_data_status") or result.get("status"),
            "message": self.payload["stage"],
            "summary": summary,
            "paths": result.get("paths"),
            "payload": self.payload,
            "export_rows": result.get("items", []),
            "export_sheet_name": result.get("export_sheet_name"),
            "export_columns": result.get("export_columns"),
            "technical_fields": result.get("technical_fields"),
            "xlsx_path": xlsx_path,
            "control_rows_count": rows_count,
            "warnings_count": warnings_count,
            "missing_inputs_count": missing_count,
            "live_network_used": False,
            "documents_downloaded": False,
            "source_excel_changed": False,
        }

    def get_sheet6_export(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "payload": self.payload,
            }
        service = Sheet6CheckDataExportService(
            self.current_project_root,
            project_service=self.project_service,
        )
        result = service.get_sheet6_export()
        result["payload"] = self.payload
        return result

    def get_contract_link_indexes(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "payload": self.payload,
                "future_matching_not_run": True,
            }
        service = ProcurementContractLinkService(project_service=self.project_service)
        result = service.get_link_indexes(self.current_project_root)
        result["payload"] = self.payload
        return result

    def get_contract_link_pre_candidates(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "payload": self.payload,
                "future_matching_not_run": True,
            }
        service = ProcurementContractLinkService(project_service=self.project_service)
        result = service.get_link_pre_candidates(self.current_project_root)
        result["payload"] = self.payload
        return result

    def get_contract_link_summary(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "payload": self.payload,
                "future_matching_not_run": True,
            }
        service = ProcurementContractLinkService(project_service=self.project_service)
        result = service.get_link_model_summary(self.current_project_root)
        result["payload"] = self.payload
        return result

    def get_contract_link_matches(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "payload": self.payload,
            }
        service = ProcurementContractLinkService(project_service=self.project_service)
        result = service.get_match_results(self.current_project_root)
        result["payload"] = self.payload
        return result

    def get_contract_link_match_summary(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "payload": self.payload,
            }
        service = ProcurementContractLinkService(project_service=self.project_service)
        result = service.get_match_summary(self.current_project_root)
        result["payload"] = self.payload
        return result

    def get_contract_link_manual_review_matches(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "payload": self.payload,
            }
        service = ProcurementContractLinkService(project_service=self.project_service)
        result = service.get_manual_review_matches(self.current_project_root)
        result["payload"] = self.payload
        return result

    def get_contract_link_conflict_matches(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "payload": self.payload,
            }
        service = ProcurementContractLinkService(project_service=self.project_service)
        result = service.get_conflict_matches(self.current_project_root)
        result["payload"] = self.payload
        return result

    def refresh_contract_link_ui_export(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Для ETAP 30 создайте или откройте проект.",
                "payload": self.payload,
            }
        service = ProcurementContractLinkService(project_service=self.project_service)
        result = service.build_contract_link_ui_exports(self.current_project_root)
        self.payload["contract_link_ui_export"] = {
            "status": result.get("status"),
            "summary": result.get("summary"),
            "paths": result.get("paths"),
        }
        self.payload["stage"] = "Связь с контрактом ETAP 30 обновлена."
        self._merge_contract_link_cards(result.get("ui_rows"))
        self._set_project_context_on_payload()
        return {
            "status": result.get("status"),
            "message": "Связь с контрактом ETAP 30 обновлена.",
            "summary": result.get("summary"),
            "paths": result.get("paths"),
            "payload": self.payload,
            "ui_rows": result.get("ui_rows", []),
        }

    def get_contract_link_ui_export(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "payload": self.payload,
            }
        service = ProcurementContractLinkService(project_service=self.project_service)
        result = service.get_contract_link_ui_export(self.current_project_root)
        self._merge_contract_link_cards(result.get("items") or result.get("ui_rows"))
        result["payload"] = self.payload
        return result

    def get_contract_link_coverage_summary(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "payload": self.payload,
            }
        service = ProcurementContractLinkService(project_service=self.project_service)
        result = service.get_contract_link_coverage_summary(self.current_project_root)
        result["payload"] = self.payload
        return result

    def get_contract_link_ai_ready_summary(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "payload": self.payload,
            }
        service = ProcurementContractLinkService(project_service=self.project_service)
        result = service.get_contract_link_ai_ready_summary(self.current_project_root)
        result["payload"] = self.payload
        return result

    def generate_ai_package_metadata(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Для metadata-only AI package создайте или откройте проект.",
                "payload": self.payload,
            }
        service = AiPackageMetadataService(
            self.current_project_root,
            project_service=self.project_service,
        )
        result = service.generate_metadata_package()
        summary = dict(result.summary)
        self.last_ai_package_metadata_summary = summary
        self.payload["ai_package_metadata"] = {
            "status": result.status,
            "summary": summary,
            "paths": result.paths.to_json_dict(),
            "validation": result.validation,
        }
        self.payload["stage"] = "Metadata-only AI package ETAP 22 создан."
        self._set_project_context_on_payload()
        return {
            "status": result.status,
            "message": "Metadata-only AI package ETAP 22 создан.",
            "summary": summary,
            "manifest": result.manifest,
            "validation": result.validation,
            "package_root": str(result.package_root),
            "files": [str(path) for path in result.files],
            "paths": result.paths.to_json_dict(),
            "payload": self.payload,
        }

    def get_ai_package_metadata_summary(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "payload": self.payload,
            }
        service = AiPackageMetadataService(
            self.current_project_root,
            project_service=self.project_service,
        )
        summary = service.load_summary()
        return {
            "status": summary.get("status", "ok"),
            "summary": summary,
            "payload": self.payload,
        }

    def validate_ai_package_metadata(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "payload": self.payload,
            }
        service = AiPackageMetadataService(
            self.current_project_root,
            project_service=self.project_service,
        )
        validation = service.validate_existing_package()
        return {
            "status": validation.get("status", "failed"),
            "validation": validation,
            "payload": self.payload,
        }

    def list_ai_package_metadata_files(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "files": [],
                "payload": self.payload,
            }
        service = AiPackageMetadataService(
            self.current_project_root,
            project_service=self.project_service,
        )
        result = service.list_package_files()
        result["payload"] = self.payload
        return result

    def generate_ai_ready_zip(self, metadata_package_path: str = "") -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Для AI-ready ZIP создайте или откройте проект.",
                "payload": self.payload,
            }
        package_path = Path(metadata_package_path) if metadata_package_path else None
        service = AiReadyZipBuilderService(
            project_service=self.project_service,
            stage_runtime_root=self.ai_ready_zip_stage_runtime_root,
        )
        result = service.build_zip_from_metadata_package(self.current_project_root, package_path)
        summary = dict(result.summary)
        self.last_ai_ready_zip_summary = summary
        self.payload["ai_ready_zip"] = {
            "status": result.status,
            "summary": summary,
            "validation": result.validation,
            "upload_queue": result.upload_queue,
            "manual_upload_required": True,
            "auto_upload": False,
            "api_upload": False,
            "cloud_upload": False,
        }
        self.payload["stage"] = "AI-ready ZIP ETAP 23 создан для ручной загрузки."
        self._set_project_context_on_payload()
        return {
            "status": result.status,
            "message": "AI-ready ZIP ETAP 23 создан для ручной загрузки.",
            "zip_path": str(result.zip_path),
            "summary": summary,
            "validation": result.validation,
            "upload_queue": result.upload_queue,
            "source_mutation_proof": result.source_mutation_proof,
            "payload": self.payload,
        }

    def validate_ai_ready_zip(self, zip_path: str = "") -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "payload": self.payload,
            }
        service = AiReadyZipBuilderService(
            project_service=self.project_service,
            stage_runtime_root=self.ai_ready_zip_stage_runtime_root,
        )
        if not zip_path:
            summary = service.get_ai_ready_zip_summary(self.current_project_root)
            items = summary.get("items") or []
            zip_path = _text(items[0].get("zip_path")) if items else ""
        if not zip_path:
            return {
                "status": "not_created",
                "message": "AI-ready ZIP еще не создан.",
                "payload": self.payload,
            }
        validation = service.validate_zip(Path(zip_path))
        return {
            "status": validation.get("status", "failed"),
            "validation": validation,
            "payload": self.payload,
        }

    def get_upload_queue(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "payload": self.payload,
            }
        service = AiReadyZipBuilderService(
            project_service=self.project_service,
            stage_runtime_root=self.ai_ready_zip_stage_runtime_root,
        )
        result = service.get_upload_queue(self.current_project_root)
        result["payload"] = self.payload
        return result

    def get_ai_ready_zip_summary(self) -> dict[str, Any]:
        if self.current_project_root is None:
            return {
                "status": "blocked",
                "message": "Проект не открыт.",
                "payload": self.payload,
            }
        service = AiReadyZipBuilderService(
            project_service=self.project_service,
            stage_runtime_root=self.ai_ready_zip_stage_runtime_root,
        )
        summary = service.get_ai_ready_zip_summary(self.current_project_root)
        return {
            "status": summary.get("status", "ok"),
            "summary": summary,
            "payload": self.payload,
        }

    def refresh_sources(self, row_ids: list[str], *, scope_label: str) -> dict[str, Any]:
        rows_for_refresh = self._current_rows_for_refresh()
        if not rows_for_refresh:
            return {
                "status": "blocked",
                "message": "Сначала загрузите табличный реестр.",
                "payload": self.payload,
            }
        available_ids = [_text(row.get("row_id")) for row in rows_for_refresh if _text(row.get("row_id"))]
        requested_ids = {str(item) for item in row_ids if str(item)}
        selected_ids = (set(available_ids) if not requested_ids else set(available_ids) & requested_ids)
        if not selected_ids:
            return {
                "status": "blocked",
                "message": "Нет строк для обновления сведений.",
                "payload": self.payload,
            }

        service = RegistryEnrichmentService(project_service=self.project_service)
        batch = service.enrich_rows(
            rows_for_refresh,
            row_ids=selected_ids,
            project_root=self.current_project_root,
            scope_label=scope_label,
            allow_network=self.allow_network_enrichment,
        )
        for row_result in batch.rows:
            self.enrichment_results_by_row_id[row_result.row_id] = row_result.to_json_dict()
        self.last_enrichment_manifest = batch.manifest_path
        self.enriched_row_ids.update(selected_ids)
        refreshed_at = batch.finished_at or _now_label()
        is_filter_scope = scope_label in {"ui_current_filter", "current_filter", "filtered"}
        scope_text = "по текущему фильтру" if is_filter_scope else "по всей таблице"
        payload = build_ui_shell_payload(
            rows_for_refresh,
            project_name=self.current_project_name,
            project_root=str(self.current_project_root or ""),
            stage=f"Сведения обновлены {scope_text}: {len(selected_ids)} строк.",
            sources_enriched=bool(self.enriched_row_ids),
            enriched_row_ids=set(self.enriched_row_ids),
            source_refreshed_at=refreshed_at,
            enrichment_results=self.enrichment_results_by_row_id,
            normative_checked=self._normative_checked_for_export(),
        )
        if self.last_import_result is not None:
            payload["import_result"] = self.last_import_result
        payload["last_loaded_file_name"] = self.last_upload_path.name if self.last_upload_path else ""
        payload["last_imported_at"] = self.payload.get("last_imported_at", "")
        payload["source_refreshed_at"] = refreshed_at
        payload["source_refresh_scope"] = scope_label
        payload["source_refresh_status"] = (
            "Сведения обновлены"
            if len(self.enriched_row_ids) >= len(available_ids)
            else "Сведения обновлены частично"
        )
        payload["source_refresh"] = {
            "status": batch.status,
            "scope_label": scope_label,
            "refreshed_rows_count": len(selected_ids),
            "enriched_rows_count": len(self.enriched_row_ids),
            "total_rows_count": len(available_ids),
            "refreshed_at": refreshed_at,
            "status_counts": batch.to_json_dict().get("status_counts", {}),
            "manifest_path": str(batch.manifest_path or ""),
            "documents_downloaded": False,
        }
        self.payload = payload
        self._set_project_context_on_payload()
        self._apply_download_project_guard()
        return {
            "status": batch.status,
            "message": f"Сведения обновлены {scope_text}: {len(selected_ids)} строк.",
            "scope_label": scope_label,
            "refreshed_rows_count": len(selected_ids),
            "enriched_rows_count": len(self.enriched_row_ids),
            "total_rows_count": len(available_ids),
            "manifest_path": str(batch.manifest_path or ""),
            "status_counts": batch.to_json_dict().get("status_counts", {}),
            "payload": self.payload,
        }

    def _current_rows_for_refresh(self) -> list[dict[str, Any]]:
        if self.current_raw_rows:
            return [dict(row) for row in self.current_raw_rows]
        return _rows_to_json_dicts(self.current_records)

    def _normative_checked_for_export(self) -> bool:
        if self.last_normative_ui_integration_summary:
            return True
        payload = self.payload.get("normative_ui_export")
        return isinstance(payload, Mapping) and _text(payload.get("status")) in {"ok", "ok_with_warnings"}

    def _export_data_readiness_for_ui(self) -> dict[str, str]:
        rows_for_export = self._current_rows_for_refresh()
        if not rows_for_export:
            return {
                "status": "registry_not_imported",
                "message": "Сначала загрузите табличный реестр.",
            }
        if not self.enriched_row_ids:
            return {
                "status": "registry_enrichment_not_started",
                "message": "Экспорт данных доступен после обновления данных и нормативной проверки.",
            }
        if not self._normative_checked_for_export():
            return {
                "status": "normative_check_not_started",
                "message": "Экспорт данных доступен после обновления данных и нормативной проверки.",
            }
        if len(self.enriched_row_ids) < len(rows_for_export):
            return {
                "status": "export_data_ready_with_warnings",
                "message": "Экспорт данных готов с предупреждениями: обновлена только часть строк.",
            }
        return {
            "status": "export_data_ready",
            "message": "Экспорт данных готов.",
        }

    def _update_export_data_action_state(self) -> None:
        readiness = self._export_data_readiness_for_ui()
        status = readiness["status"]
        self.payload["export_data_status"] = status
        for action in self.payload.get("top_actions", []):
            if not isinstance(action, dict) or action.get("action") != "sheet6_export":
                continue
            action["label"] = "Экспорт данных"
            action["enabled"] = status in {"export_data_ready", "export_data_ready_with_warnings"}
            action["disabled_reason"] = "" if action["enabled"] else readiness["message"]

    def _apply_download_project_guard(self) -> None:
        has_project = self.current_project_root is not None
        for row in self.payload.get("rows", []):
            if not isinstance(row, dict):
                continue
            documents = row.get("documents")
            if not isinstance(documents, dict):
                continue
            can_download = bool(documents.get("download_available"))
            row_enriched = row.get("source_refresh_status") == "Сведения обновлены"
            disabled_reason = ""
            if can_download and not row_enriched:
                disabled_reason = "Сначала обновите данные по строке или по фильтру."
                documents["disabled_reason"] = disabled_reason
                documents["text"] = disabled_reason
            elif can_download and not has_project:
                disabled_reason = "Для скачивания документов создайте или откройте проект."
                documents["disabled_reason"] = disabled_reason
                documents["text"] = disabled_reason
            for action in row.get("actions", []):
                if not isinstance(action, dict) or action.get("action") != "download_row":
                    continue
                if can_download:
                    action["enabled"] = bool(has_project and row_enriched)
                    action["disabled_reason"] = "" if has_project and row_enriched else disabled_reason
                    action["text"] = (
                        "Скачать документы контракта"
                        if row.get("source_dataset_type") == "contract"
                        else "Скачать документы"
                    )
            row["card_sections"] = build_card_sections(row)
        selected = self.payload.get("selected_row")
        if isinstance(selected, dict):
            row_id = selected.get("row_id")
            replacement = next(
                (
                    row
                    for row in self.payload.get("rows", [])
                    if isinstance(row, dict) and row.get("row_id") == row_id
                ),
                None,
            )
            if isinstance(replacement, dict):
                self.payload["selected_row"] = replacement
                self.payload["card_sections"] = replacement.get("card_sections", [])

    def _build_protocol_preview_summaries(self, parsed: Mapping[str, Any]) -> list[dict[str, Any]]:
        supplier_items = [
            dict(item)
            for item in parsed.get("supplier_result_evidence", [])
            if isinstance(item, Mapping)
        ]
        contract_items = [
            dict(item)
            for item in parsed.get("embedded_contract_evidence", [])
            if isinstance(item, Mapping)
        ]
        by_row_id: dict[str, dict[str, dict[str, Any]]] = {}
        for item in supplier_items:
            row_id = _text(item.get("row_id")) or f"row-supplier-result-{_text(item.get('procurement_number'))}"
            by_row_id.setdefault(row_id, {})["supplier"] = item
        for item in contract_items:
            row_id = _text(item.get("row_id")) or f"row-supplier-result-{_text(item.get('procurement_number'))}"
            by_row_id.setdefault(row_id, {})["contract"] = item

        summaries: list[dict[str, Any]] = []
        for row_id, evidence in by_row_id.items():
            supplier = evidence.get("supplier") or {}
            contract = evidence.get("contract") or {}
            procurement_number = (
                _text(supplier.get("procurement_number"))
                or _text(contract.get("procurement_number"))
            )
            supplier_price = _text(supplier.get("participant_offer_price"))
            contract_price = _text(contract.get("contract_price"))
            price_consistency = "insufficient_data"
            if supplier_price and contract_price:
                price_consistency = "matched" if supplier_price == contract_price else "requires_review"
            warnings = []
            for item in (supplier, contract):
                warnings.extend([_text(warning) for warning in item.get("warnings", []) if _text(warning)])
            if supplier:
                warnings.append("Для подсчета заявок и участников нужен протокол или страница Список заявок.")
            manual_review = bool(warnings) and (
                not _text(contract.get("contract_supplier_name"))
                or "Для подсчета заявок и участников нужен протокол или страница Список заявок." in warnings
            )
            summaries.append(
                {
                    "row_id": row_id,
                    "law": "44-ФЗ",
                    "procurement_number": procurement_number,
                    "found_in_source": bool(supplier or contract),
                    "found_count": 1 if supplier or contract else 0,
                    "status": "temp_parse_ok" if supplier or contract else "not_found",
                    "evidence_storage_status": "temp_download_parsed_deleted",
                    "protocol_documents": [],
                    "supplier_result_evidence": supplier or None,
                    "embedded_contract_evidence": contract or None,
                    "participants_summary": {
                        "participants_total_count": None,
                        "participants_named_count": 0,
                        "participants_with_inn_count": 0,
                        "applications_total_count": None,
                        "admitted_count": 0,
                        "rejected_count": 0,
                    },
                    "price_summary": {
                        "initial_price": "",
                        "winner_price": "",
                        "supplier_result_offer_price": supplier_price,
                        "contract_price": contract_price,
                        "price_consistency_status": price_consistency,
                    },
                    "winner_summary": {
                        "winner_display_name": _text(contract.get("contract_supplier_name")) or None,
                        "winner_display_status": (
                            "supplier_from_embedded_contract_evidence"
                            if _text(contract.get("contract_supplier_name"))
                            else "supplier_result_without_supplier_name"
                        ),
                        "winner_source": (
                            "embedded_contract_registry_evidence"
                            if _text(contract.get("contract_supplier_name"))
                            else "eis_supplier_results"
                        ),
                        "winner_confidence": _text(contract.get("confidence")) or _text(supplier.get("confidence")) or "medium",
                    },
                    "manual_review_required": manual_review,
                    "manual_review_items": warnings if manual_review else [],
                    "warnings": warnings,
                    "next_action": "Временный результат получен; для полного пакета документов создайте проект.",
                    "source_protocol_ids": [],
                }
            )
        return summaries

    def _preview_summaries_for_requested_rows(
        self,
        parsed_summaries: list[dict[str, Any]],
        row_ids: list[str],
    ) -> list[dict[str, Any]]:
        if not row_ids:
            return parsed_summaries
        rows_by_id = {
            _text(row.get("row_id")): row
            for row in self.payload.get("rows", [])
            if isinstance(row, dict) and _text(row.get("row_id"))
        }
        parsed_by_row_id = {
            _text(item.get("row_id")): item
            for item in parsed_summaries
            if _text(item.get("row_id"))
        }
        parsed_by_procurement = {
            _text(item.get("procurement_number")): item
            for item in parsed_summaries
            if _text(item.get("procurement_number"))
        }
        result: list[dict[str, Any]] = []
        parsed_numbers = [
            _text(item.get("procurement_number"))
            for item in parsed_summaries
            if _text(item.get("procurement_number"))
        ]
        for row_id in row_ids:
            selected = rows_by_id.get(_text(row_id))
            if selected is None:
                continue
            procurement_number = _text(
                selected.get("procurement_number")
                or selected.get("purchase_number")
                or selected.get("identifier_display")
                or selected.get("primary_identifier")
            )
            matched = parsed_by_row_id.get(_text(row_id)) or parsed_by_procurement.get(procurement_number)
            if matched:
                card = dict(matched)
                card["row_id"] = _text(row_id)
                result.append(card)
                continue
            result.append(
                {
                    "row_id": _text(row_id),
                    "law": _text(selected.get("law")) or "Не определено",
                    "procurement_number": procurement_number,
                    "found_in_source": False,
                    "found_count": len(parsed_summaries),
                    "status": "temp_parse_source_mismatch" if parsed_summaries else "not_found",
                    "evidence_storage_status": "temp_download_parsed_deleted",
                    "protocol_documents": [],
                    "supplier_result_evidence": None,
                    "embedded_contract_evidence": None,
                    "participants_summary": {
                        "participants_total_count": 0,
                        "participants_named_count": 0,
                        "participants_with_inn_count": 0,
                        "applications_total_count": 0,
                        "admitted_count": 0,
                        "rejected_count": 0,
                    },
                    "price_summary": {
                        "initial_price": "",
                        "winner_price": "",
                        "supplier_result_offer_price": "",
                        "contract_price": "",
                        "price_consistency_status": "source_mismatch",
                    },
                    "winner_summary": {
                        "winner_display_name": None,
                        "winner_display_status": "manual_review_required",
                        "winner_source": "temp_parse_source_mismatch",
                        "winner_confidence": "none",
                    },
                    "manual_review_required": True,
                    "manual_review_items": [
                        "temp_parse_source_procurement_number_mismatch",
                        *[f"parsed_procurement_number:{number}" for number in parsed_numbers],
                    ],
                    "warnings": [
                        "temp_parse_source_procurement_number_mismatch",
                        *[f"parsed_procurement_number:{number}" for number in parsed_numbers],
                    ],
                    "next_action": (
                        "Временный результат разобран, но номер источника не совпадает с выбранной строкой; "
                        "нужна ручная проверка или источник для этой закупки."
                    ),
                    "source_protocol_ids": [],
                }
            )
        return result if result else parsed_summaries

    def _merge_protocol_ui_cards(self, summaries: Any) -> None:
        if not isinstance(summaries, list):
            return
        by_row_id: dict[str, Mapping[str, Any]] = {}
        by_procurement: dict[str, Mapping[str, Any]] = {}
        for item in summaries:
            if not isinstance(item, Mapping):
                continue
            row_id = _text(item.get("row_id"))
            procurement_number = _text(item.get("procurement_number"))
            if row_id:
                by_row_id[row_id] = item
            if procurement_number:
                by_procurement[procurement_number] = item
        for row in self.payload.get("rows", []):
            if not isinstance(row, dict):
                continue
            procurement_number = _text(
                row.get("procurement_number")
                or row.get("purchase_number")
                or row.get("identifier_display")
                or row.get("primary_identifier")
            )
            summary = by_row_id.get(_text(row.get("row_id"))) or by_procurement.get(procurement_number)
            if not summary:
                continue
            row["protocols"] = dict(summary)
            warnings = list(row.get("warnings") or [])
            for warning in summary.get("warnings", []):
                warning_text = _text(warning)
                if warning_text:
                    warnings.append({"code": warning_text, "message": warning_text, "field": "protocols"})
            row["warnings"] = warnings
            row["has_warnings"] = bool(warnings)
            row["card_sections"] = build_card_sections(row)
        selected = self.payload.get("selected_row")
        if isinstance(selected, dict):
            replacement = next(
                (
                    row
                    for row in self.payload.get("rows", [])
                    if isinstance(row, dict) and row.get("row_id") == selected.get("row_id")
                ),
                None,
            )
            if isinstance(replacement, dict):
                self.payload["selected_row"] = replacement
                self.payload["card_sections"] = replacement.get("card_sections", [])

    def _merge_contract_link_cards_from_existing_outputs(self) -> None:
        if self.current_project_root is None:
            return
        try:
            service = ProcurementContractLinkService(project_service=self.project_service)
            result = service.get_contract_link_ui_export(self.current_project_root)
        except Exception:
            return
        self._merge_contract_link_cards(result.get("items") or result.get("ui_rows"))

    def _merge_contract_link_cards(self, ui_rows: Any) -> None:
        if not isinstance(ui_rows, list):
            return
        by_key: dict[str, Mapping[str, Any]] = {}
        for item in ui_rows:
            if not isinstance(item, Mapping):
                continue
            card = item.get("contract_link_card") if isinstance(item.get("contract_link_card"), Mapping) else item
            keys = [
                _text(item.get("row_id")),
                _text(item.get("original_procurement_row_id")),
                _text(item.get("original_contract_row_id")),
                _text(item.get("procurement_number")),
                _text(item.get("contract_registry_number")),
            ]
            if isinstance(item.get("found_contracts"), list):
                keys.extend(_text(value) for value in item.get("found_contracts", []))
            for key in keys:
                if key:
                    by_key.setdefault(key, card)
        for row in self.payload.get("rows", []):
            if not isinstance(row, dict):
                continue
            keys = [
                _text(row.get("row_id")),
                _text(row.get("procurement_number") or row.get("purchase_number") or row.get("identifier_display") or row.get("primary_identifier")),
                _text(row.get("contract_registry_number") or row.get("contract_number")),
            ]
            card = next((by_key[key] for key in keys if key and key in by_key), None)
            if not card:
                continue
            row["contract_link"] = dict(card)
            warnings = list(row.get("warnings") or [])
            for warning in card.get("warnings", []) if isinstance(card.get("warnings"), list) else []:
                warning_text = _text(warning)
                if warning_text:
                    warnings.append({"code": warning_text, "message": warning_text, "field": "contract_link"})
            row["warnings"] = warnings
            row["has_warnings"] = bool(warnings)
            row["card_sections"] = build_card_sections(row)
        selected = self.payload.get("selected_row")
        if isinstance(selected, dict):
            replacement = next(
                (
                    row
                    for row in self.payload.get("rows", [])
                    if isinstance(row, dict) and row.get("row_id") == selected.get("row_id")
                ),
                None,
            )
            if isinstance(replacement, dict):
                self.payload["selected_row"] = replacement
                self.payload["card_sections"] = replacement.get("card_sections", [])

    def _merge_normative_ui_rows(self, ui_rows: Any) -> None:
        if not isinstance(ui_rows, list):
            return
        by_key: dict[str, Mapping[str, Any]] = {}
        for item in ui_rows:
            if not isinstance(item, Mapping):
                continue
            keys = [
                _text(item.get("row_id")),
                _text(item.get("procurement_number")),
            ]
            for key in keys:
                if key:
                    by_key.setdefault(key, item)
        for raw_row in self.current_raw_rows:
            if not isinstance(raw_row, dict):
                continue
            key = _text(raw_row.get("row_id")) or _text(raw_row.get("procurement_number") or raw_row.get("purchase_number"))
            normative = by_key.get(key)
            if not normative:
                continue
            raw_row["normative_status"] = _text(normative.get("normative_status"))
            raw_row["normative_status_source"] = _text(normative.get("normative_status_source"))
            raw_row["normative_enrichment"] = dict(normative.get("normative_enrichment") or {})

        for row in self.payload.get("rows", []):
            if not isinstance(row, dict):
                continue
            keys = [
                _text(row.get("row_id")),
                _text(row.get("procurement_number") or row.get("purchase_number") or row.get("identifier_display") or row.get("primary_identifier")),
            ]
            normative = next((by_key[key] for key in keys if key and key in by_key), None)
            if not normative:
                continue
            row_id = _text(row.get("row_id"))
            row["normative_status"] = _text(normative.get("normative_status")) or row.get("normative_status")
            row["normative_status_source"] = _text(normative.get("normative_status_source")) or row.get("normative_status_source")
            row["normative_enrichment"] = dict(normative.get("normative_enrichment") or {})
            status_block = dict(row.get("status") or {})
            status_block["normative_status"] = row["normative_status"]
            status_block["normative_status_source"] = row["normative_status_source"]
            status_block["right_card_explanation"] = _text(row["normative_enrichment"].get("explanation"))
            status_block["recommended_action"] = _text(row["normative_enrichment"].get("recommended_analyst_action"))
            row["status"] = status_block
            warnings = list(row.get("warnings") or [])
            for warning in normative.get("warnings", []) if isinstance(normative.get("warnings"), list) else []:
                warning_text = _text(warning)
                if warning_text:
                    warnings.append({"code": warning_text, "message": warning_text, "field": "normative"})
            row["warnings"] = warnings
            row["has_warnings"] = bool(warnings)
            if row_id:
                self.normative_results_by_row_id[row_id] = dict(normative)
                enriched = self.enrichment_results_by_row_id.setdefault(row_id, {})
                enriched["normative"] = dict(row["normative_enrichment"])
            row["card_sections"] = build_card_sections(row)
        selected = self.payload.get("selected_row")
        if isinstance(selected, dict):
            replacement = next(
                (
                    row
                    for row in self.payload.get("rows", [])
                    if isinstance(row, dict) and row.get("row_id") == selected.get("row_id")
                ),
                None,
            )
            if isinstance(replacement, dict):
                self.payload["selected_row"] = replacement
                self.payload["card_sections"] = replacement.get("card_sections", [])

    def _merge_download_result_into_payload(self, result: Any) -> None:
        rows_by_id = {row["row_id"]: row for row in self.payload.get("rows", []) if isinstance(row, dict)}
        for row_result in result.rows:
            ui_row = rows_by_id.get(row_result.row_id)
            if not ui_row:
                continue
            documents = dict(ui_row.get("documents") or {})
            status_code = row_result.download_status or row_result.status
            status_label = _download_status_label(status_code)
            check_status = _download_check_status_label(status_code)
            target_folder = row_result.download_target_folder or row_result.document_folder_path
            downloaded_files = _safe_row_downloaded_files(row_result)
            downloaded_count = _safe_row_downloaded_files_count(row_result)
            expected_count = _safe_row_documents_found_count(row_result)
            count_label = f"{downloaded_count} / {expected_count or 0}"
            if downloaded_files or downloaded_count > 0:
                documents["loaded"] = True
                documents["status"] = status_code
                documents["label"] = status_label
                documents["count_label"] = count_label
                documents["folder_path"] = str(target_folder or "")
                documents["can_open_folder"] = bool(target_folder)
            else:
                documents["loaded"] = False
                documents["status"] = status_code
                documents["label"] = status_label
                documents["count_label"] = count_label
                documents["folder_path"] = str(target_folder or "")
                documents["can_open_folder"] = bool(target_folder)
            ui_row["documents_status"] = status_label
            ui_row["documents_count_label"] = count_label
            ui_row["check_status"] = check_status
            ui_row["check_status_reason"] = row_result.download_next_action or status_label
            documents["download_status"] = status_code
            documents["download_attempted_at"] = row_result.download_attempted_at
            documents["download_route"] = row_result.download_route
            documents["download_source_url"] = row_result.download_source_url
            documents["downloaded_count"] = downloaded_count
            documents["expected_count"] = expected_count or 0
            documents["download_files_count"] = downloaded_count
            documents["documents_found_count"] = expected_count or 0
            documents["documents_downloaded_count"] = downloaded_count
            documents["documents_skipped_count"] = max(0, (expected_count or 0) - downloaded_count)
            documents["download_bytes_total"] = row_result.download_bytes_total
            documents["download_error_code"] = row_result.download_error_code
            documents["download_warning"] = row_result.download_warning
            documents["download_next_action"] = row_result.download_next_action
            documents["text"] = row_result.download_next_action or documents.get("text")
            ui_row["documents"] = documents
            ui_row["card_sections"] = build_card_sections(ui_row)
        self.payload["download_manifest_path"] = str(result.manifest_path)
        project_root = self.current_project_root or result.project_root
        download_reports_root = project_root / "08_Рабочая_документация_приложения" / "download_reports"
        self.payload["download_summary"] = {
            "status": result.status,
            "downloaded_rows_count": result.downloaded_rows_count,
            "downloaded_files_count": result.downloaded_files_count,
            "manifest_name": CONTROLLED_DOWNLOAD_MANIFEST_NAME,
            "manifest_path": str(result.manifest_path),
            "download_reports_root": str(download_reports_root),
            "document_download_index_csv": str(download_reports_root / "DOCUMENT_DOWNLOAD_INDEX.csv"),
            "document_download_index_md": str(download_reports_root / "DOCUMENT_DOWNLOAD_INDEX.md"),
            "download_pipeline_run": str(download_reports_root / "DOWNLOAD_PIPELINE_RUN.json"),
        }

    def _mapping_required_response(
        self,
        preview: ImportPreviewResult,
        upload_path: Path,
        *,
        field_mapping: Mapping[str, str] | None = None,
        validation_errors: tuple[str, ...] = (),
    ) -> dict[str, Any]:
        mapping_payload = self._build_mapping_payload(
            preview,
            upload_path,
            field_mapping=field_mapping,
            validation_errors=validation_errors,
            dataset_type=_dataset_type_from_mapping(field_mapping, fallback=_dataset_type_from_preview(preview)),
        )
        self.payload["stage"] = "Требуется сопоставление колонок. Экран сопоставления открыт."
        self.payload["import_preview"] = preview.to_json_dict()
        self.payload["import_mapping"] = mapping_payload
        self.payload["field_mapping_fields"] = mapping_payload["target_fields"]
        self.payload["mapping_modal_open"] = True
        self._set_project_context_on_payload()
        return {
            "status": MAPPING_REQUIRED_STATUS,
            "profile": preview.profile_code or PROFILE_UNKNOWN_REQUIRES_MAPPING,
            "message": "Требуется сопоставление колонок",
            "action": MAPPING_REQUIRED_ACTION,
            "source_file": str(upload_path),
            "sheet_name": preview.sheet_name,
            "header_row": preview.header_row,
            "source_columns": list(preview.columns),
            "sample_values": {column: list(values) for column, values in preview.sample_values.items()},
            "dataset_type_suggestion": mapping_payload["dataset_type_suggestion"],
            "target_fields": mapping_payload["target_fields"],
            "import_session_id": self.last_import_session_id,
            "project_root": str(self.current_project_root or ""),
            "mapping_errors": list(validation_errors),
            "payload": self.payload,
        }

    def _build_mapping_payload(
        self,
        preview: ImportPreviewResult,
        upload_path: Path,
        *,
        field_mapping: Mapping[str, str] | None = None,
        validation_errors: tuple[str, ...] = (),
        dataset_type: str | None = None,
    ) -> dict[str, Any]:
        effective_dataset_type = dataset_type or _dataset_type_from_preview(preview)
        return {
            "status": MAPPING_REQUIRED_STATUS,
            "profile": preview.profile_code or PROFILE_UNKNOWN_REQUIRES_MAPPING,
            "message": "Требуется сопоставление колонок",
            "action": MAPPING_REQUIRED_ACTION,
            "source_file": str(upload_path),
            "source_file_name": upload_path.name,
            "sheet_name": preview.sheet_name,
            "header_row": preview.header_row,
            "source_columns": list(preview.columns),
            "sample_values": {column: list(values) for column, values in preview.sample_values.items()},
            "dataset_type_suggestion": effective_dataset_type,
            "import_session_id": self.last_import_session_id,
            "project_root": str(self.current_project_root or ""),
            "field_mapping": _clean_mapping(field_mapping or preview.suggested_field_mapping),
            "target_fields": mapping_targets_json(effective_dataset_type),
            "mapping_errors": list(validation_errors),
        }

    def _set_project_context_on_payload(self) -> None:
        summary = self.payload.setdefault("project_summary", {})
        summary["project_name"] = self.current_project_name or "Проект не выбран"
        summary["project_root"] = str(self.current_project_root or "")
        self.payload["default_project_root"] = str(self.default_user_projects_root)
        self.payload["project_status"] = "Проект создан" if self.current_project_root else "Проект не создан"
        self.payload["registry_status"] = "Загружен" if self.payload.get("rows") else "Не загружен"
        if len(self.last_upload_paths) > 1:
            self.payload["last_loaded_file_name"] = _file_list_label(self.last_upload_paths)
        elif self.last_upload_path is not None:
            self.payload["last_loaded_file_name"] = self.last_upload_path.name
        self.payload.setdefault("last_loaded_file_name", "")
        self.payload.setdefault("last_imported_at", "")
        if not self.payload.get("field_mapping_fields"):
            self.payload["field_mapping_fields"] = mapping_targets_json(DATASET_PROCUREMENT)
        if self.import_service is not None and self.current_project_root is not None:
            self.payload["column_mappings"] = self.import_service.load_column_mappings(self.current_project_root)
        else:
            self.payload["column_mappings"] = {"schema_version": "tendervestdocs_column_mapping_v1", "mappings": []}
        if self.startup_project_warning:
            self.payload["project_warning"] = self.startup_project_warning

    def _restore_last_project_on_start(self) -> None:
        restored = self.project_service.restore_last_project()
        if restored.status == "ok" and restored.project is not None:
            self.current_project_root = restored.project.project_root
            self.current_project_name = restored.project.project_name
            self.payload = build_empty_ui_shell_payload()
            self.payload["stage"] = "Последний проект открыт. Выберите табличный файл."
            return
        if restored.status == "missing":
            self.startup_project_warning = restored.message
            self.payload["stage"] = restored.message
        elif restored.status == "error":
            self.startup_project_warning = restored.message
            self.payload["stage"] = "Последний проект не открыт. Создайте новый проект или укажите существующую папку."


class LocalAppHandler(BaseHTTPRequestHandler):
    server_version = "TenderVestDocsLocalApp/19.2"

    @property
    def state(self) -> LocalAppState:
        return self.server.state  # type: ignore[attr-defined]

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/api/state":
            self._send_json({"status": "ok", "payload": self.state.payload})
            return
        if parsed.path == "/api/health":
            self._send_json(
                {
                    "status": "ok",
                    "project_root": str(PROJECT_ROOT),
                    "pilot_root": str(self.state.pilot_root),
                    "default_project_root": str(self.state.default_user_projects_root),
                    "allow_network_downloads": self.state.allow_network_downloads,
                    "allow_network_enrichment": self.state.allow_network_enrichment,
                }
            )
            return
        if parsed.path == "/api/document-inventory/summary":
            self._send_json(self.state.get_document_inventory_summary())
            return
        if parsed.path == "/api/text-extraction/summary":
            self._send_json(self.state.get_text_extraction_summary())
            return
        if parsed.path == "/api/protocol-inventory":
            self._send_json(self.state.get_protocol_inventory())
            return
        if parsed.path == "/api/protocol-inventory/summary":
            self._send_json(self.state.get_protocol_inventory_summary())
            return
        if parsed.path == "/api/protocol-data":
            self._send_json(self.state.get_protocol_data())
            return
        if parsed.path == "/api/protocol-data/summary":
            self._send_json(self.state.get_protocol_data_summary())
            return
        if parsed.path == "/api/protocol-data/participants":
            self._send_json(self.state.get_protocol_participants())
            return
        if parsed.path == "/api/protocol-data/price-offers":
            self._send_json(self.state.get_protocol_price_offers())
            return
        if parsed.path == "/api/protocol-data/results":
            self._send_json(self.state.get_protocol_results())
            return
        if parsed.path == "/api/protocol-ui-summary":
            self._send_json(self.state.get_protocol_ui_summary())
            return
        if parsed.path.startswith("/api/protocol-ui-summary/"):
            row_id = parsed.path.rsplit("/", 1)[-1]
            self._send_json(self.state.get_protocol_ui_summary(row_id))
            return
        if parsed.path == "/api/protocol-export":
            self._send_json(self.state.get_protocol_export())
            return
        if parsed.path == "/api/protocol-ai-ready-summary":
            self._send_json(self.state.get_protocol_ai_ready_summary())
            return
        if parsed.path == "/api/normative-export":
            self._send_json(self.state.get_normative_export())
            return
        if parsed.path == "/api/normative-ai-ready-summary":
            self._send_json(self.state.get_normative_ai_ready_summary())
            return
        if parsed.path == "/api/normative-sheet6-readiness":
            self._send_json(self.state.get_normative_sheet6_readiness())
            return
        if parsed.path == "/api/sheet6-export":
            self._send_json(self.state.get_sheet6_export())
            return
        if parsed.path == "/api/contract-link/indexes":
            self._send_json(self.state.get_contract_link_indexes())
            return
        if parsed.path == "/api/contract-link/pre-candidates":
            self._send_json(self.state.get_contract_link_pre_candidates())
            return
        if parsed.path == "/api/contract-link/summary":
            self._send_json(self.state.get_contract_link_summary())
            return
        if parsed.path == "/api/contract-link/matches":
            self._send_json(self.state.get_contract_link_matches())
            return
        if parsed.path == "/api/contract-link/matches/summary":
            self._send_json(self.state.get_contract_link_match_summary())
            return
        if parsed.path == "/api/contract-link/matches/manual-review":
            self._send_json(self.state.get_contract_link_manual_review_matches())
            return
        if parsed.path == "/api/contract-link/matches/conflicts":
            self._send_json(self.state.get_contract_link_conflict_matches())
            return
        if parsed.path == "/api/contract-link/ui-export":
            self._send_json(self.state.get_contract_link_ui_export())
            return
        if parsed.path == "/api/contract-link/coverage-summary":
            self._send_json(self.state.get_contract_link_coverage_summary())
            return
        if parsed.path == "/api/contract-link/ai-ready-summary":
            self._send_json(self.state.get_contract_link_ai_ready_summary())
            return
        if parsed.path == "/api/ai-package-metadata/summary":
            self._send_json(self.state.get_ai_package_metadata_summary())
            return
        if parsed.path == "/api/ai-package-metadata/files":
            self._send_json(self.state.list_ai_package_metadata_files())
            return
        if parsed.path == "/api/ai-ready-zip/summary":
            self._send_json(self.state.get_ai_ready_zip_summary())
            return
        if parsed.path == "/api/upload-queue":
            self._send_json(self.state.get_upload_queue())
            return
        if parsed.path == "/":
            self._send_file(STATIC_DIR / "index.html", "text/html; charset=utf-8")
            return
        static_path = (STATIC_DIR / parsed.path.lstrip("/")).resolve()
        try:
            static_path.relative_to(STATIC_DIR.resolve())
        except ValueError:
            self._send_json({"status": "error", "message": "Недопустимый путь."}, HTTPStatus.BAD_REQUEST)
            return
        if not static_path.is_file():
            self._send_json({"status": "not_found", "message": "Файл не найден."}, HTTPStatus.NOT_FOUND)
            return
        content_type = _content_type(static_path)
        self._send_file(static_path, content_type)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        try:
            if parsed.path == "/api/projects/select-folder":
                selected = select_project_folder()
                if selected is None:
                    self._send_json(
                        {
                            "status": "cancelled",
                            "message": "Выбор папки проекта отменен.",
                            "payload": self.state.payload,
                        }
                    )
                    return
                self._send_json(
                    {
                        "status": "ok",
                        "project_root": str(selected),
                        "message": "Папка для сохранения проекта выбрана.",
                        "payload": self.state.payload,
                    }
                )
                return
            if parsed.path == "/api/projects/create":
                body = self._read_json_body()
                parent_root = Path(_text(body.get("parent_root"))) if _text(body.get("parent_root")) else None
                if parent_root is not None:
                    self._send_json(
                        self.state.create_project_workspace(
                            _text(body.get("project_name")) or DEFAULT_PROJECT_NAME,
                            parent_root=parent_root,
                        )
                    )
                    return
                user_root = Path(_text(body.get("user_root"))) if _text(body.get("user_root")) else None
                self._send_json(
                    self.state.create_project(
                        _text(body.get("project_name")) or DEFAULT_PROJECT_NAME,
                        user_root=user_root,
                    )
                )
                return
            if parsed.path == "/api/projects/open":
                body = self._read_json_body()
                project_root = Path(_text(body.get("project_root"))) if _text(body.get("project_root")) else None
                self._send_json(self.state.open_project(project_root))
                return
            if parsed.path == "/api/import/preview":
                fields = self._read_multipart()
                upload_paths = self._save_uploads(fields)
                if len(upload_paths) == 1:
                    self._send_json(self.state.preview_import(upload_paths[0], sheet_name=_text(fields.get("sheet_name")) or None))
                else:
                    self._send_json(self.state.preview_imports(upload_paths))
                return
            if parsed.path == "/api/import":
                fields = self._read_multipart()
                upload_paths = self._save_uploads(fields)
                mapping = _parse_json_field(fields.get("field_mapping")) or {}
                if len(upload_paths) == 1:
                    self._send_json(
                        self.state.import_uploaded(
                            upload_paths[0],
                            sheet_name=_text(fields.get("sheet_name")) or None,
                            field_mapping=mapping,
                        )
                    )
                else:
                    self._send_json(
                        self.state.import_uploaded_files(
                            upload_paths,
                            field_mapping=mapping,
                        )
                    )
                return
            if parsed.path == "/api/mapping/save":
                body = self._read_json_body()
                mapping = body.get("field_mapping") if isinstance(body.get("field_mapping"), Mapping) else {}
                self._send_json(
                    self.state.save_current_mapping(
                        {str(key): str(value) for key, value in mapping.items()},
                        dataset_type=_text(body.get("dataset_type")) or DATASET_PROCUREMENT,
                    )
                )
                return
            if parsed.path == "/api/mapping/reset":
                self._send_json(self.state.reset_column_mappings())
                return
            if parsed.path == "/api/open-link":
                body = self._read_json_body()
                self._send_json(self.state.open_link(_text(body.get("row_id"))))
                return
            if parsed.path == "/api/refresh-sources":
                body = self._read_json_body()
                row_ids = body.get("row_ids")
                selected = [str(item) for item in row_ids] if isinstance(row_ids, list) else []
                scope_label = _text(body.get("scope_label")) or "ui_current_filter"
                self._send_json(self.state.refresh_sources(selected, scope_label=scope_label))
                return
            if parsed.path == "/api/normative-check":
                body = self._read_json_body()
                row_ids = body.get("row_ids")
                selected = [str(item) for item in row_ids] if isinstance(row_ids, list) else []
                scope_label = _text(body.get("scope_label")) or "ui_all_rows"
                self._send_json(self.state.run_normative_check(selected, scope_label=scope_label))
                return
            if parsed.path == "/api/sheet6-export":
                body = self._read_json_body()
                row_ids = body.get("row_ids")
                selected = [str(item) for item in row_ids] if isinstance(row_ids, list) else []
                scope_label = _text(body.get("scope_label")) or "ui_all_rows"
                self._send_json(self.state.run_sheet6_export(selected, scope_label=scope_label))
                return
            if parsed.path == "/api/download":
                body = self._read_json_body()
                row_ids = body.get("row_ids")
                selected = [str(item) for item in row_ids] if isinstance(row_ids, list) else []
                scope_label = _text(body.get("scope_label")) or "ui_selected_rows"
                self._send_json(self.state.download_rows(selected, scope_label=scope_label))
                return
            if parsed.path == "/api/protocol-results/preview":
                body = self._read_json_body()
                row_ids = body.get("row_ids")
                selected = [str(item) for item in row_ids] if isinstance(row_ids, list) else []
                scope_label = _text(body.get("scope_label")) or "ui_protocol_results_preview"
                self._send_json(self.state.preview_protocol_results(selected, scope_label=scope_label))
                return
            if parsed.path == "/api/document-inventory/refresh":
                self._send_json(self.state.refresh_document_inventory())
                return
            if parsed.path == "/api/document-inventory/row":
                body = self._read_json_body()
                self._send_json(self.state.get_document_inventory_for_row(_text(body.get("row_id"))))
                return
            if parsed.path == "/api/text-extraction/refresh":
                self._send_json(self.state.refresh_text_extraction())
                return
            if parsed.path == "/api/text-extraction/document":
                body = self._read_json_body()
                self._send_json(self.state.get_extracted_text_for_document(_text(body.get("document_id"))))
                return
            if parsed.path == "/api/text-extraction/row-fragments":
                body = self._read_json_body()
                self._send_json(self.state.get_fragments_for_row(_text(body.get("row_id"))))
                return
            if parsed.path == "/api/text-extraction/document-fragments":
                body = self._read_json_body()
                self._send_json(self.state.get_fragments_for_document(_text(body.get("document_id"))))
                return
            if parsed.path == "/api/protocol-inventory/refresh":
                self._send_json(self.state.refresh_protocol_inventory())
                return
            if parsed.path == "/api/protocol-inventory/source-coverage":
                body = self._read_json_body()
                row_ids = body.get("row_ids")
                selected = [str(item) for item in row_ids] if isinstance(row_ids, list) else []
                local_evidence_paths = body.get("local_evidence_paths")
                evidence_paths = (
                    [str(item) for item in local_evidence_paths]
                    if isinstance(local_evidence_paths, list)
                    else []
                )
                self._send_json(
                    self.state.check_protocol_source_coverage(
                        selected,
                        allow_network=bool(body.get("allow_network")),
                        local_evidence_paths=evidence_paths,
                    )
                )
                return
            if parsed.path == "/api/protocol-data/refresh":
                self._send_json(self.state.refresh_protocol_data())
                return
            if parsed.path == "/api/protocol-ui-summary/refresh":
                self._send_json(self.state.refresh_protocol_ui_summary())
                return
            if parsed.path == "/api/contract-link/ui-export/refresh":
                self._send_json(self.state.refresh_contract_link_ui_export())
                return
            if parsed.path == "/api/protocol-export":
                self._send_json(self.state.get_protocol_export())
                return
            if parsed.path == "/api/protocol-ai-ready-summary":
                self._send_json(self.state.get_protocol_ai_ready_summary())
                return
            if parsed.path == "/api/ai-package-metadata/generate":
                self._send_json(self.state.generate_ai_package_metadata())
                return
            if parsed.path == "/api/ai-package-metadata/validate":
                self._send_json(self.state.validate_ai_package_metadata())
                return
            if parsed.path == "/api/ai-ready-zip/generate":
                body = self._read_json_body()
                self._send_json(self.state.generate_ai_ready_zip(_text(body.get("metadata_package_path"))))
                return
            if parsed.path == "/api/ai-ready-zip/validate":
                body = self._read_json_body()
                self._send_json(self.state.validate_ai_ready_zip(_text(body.get("zip_path"))))
                return
        except ProjectServiceError as exc:
            self._send_json(
                {"status": "blocked", "message": str(exc), "payload": self.state.payload},
                HTTPStatus.OK,
            )
            return
        except Exception as exc:
            self._send_json(
                {"status": "error", "message": str(exc), "error_type": type(exc).__name__},
                HTTPStatus.INTERNAL_SERVER_ERROR,
            )
            return
        self._send_json({"status": "not_found", "message": "Раздел API не найден."}, HTTPStatus.NOT_FOUND)

    def log_message(self, format: str, *args: Any) -> None:
        return

    def _read_json_body(self) -> dict[str, Any]:
        length = int(self.headers.get("Content-Length") or "0")
        if length <= 0:
            return {}
        return json.loads(self.rfile.read(length).decode("utf-8"))

    def _read_multipart(self) -> dict[str, Any]:
        content_type = self.headers.get("Content-Type", "")
        if "multipart/form-data" not in content_type:
            raise ValueError("Ожидался multipart/form-data.")
        length = int(self.headers.get("Content-Length") or "0")
        raw = self.rfile.read(length)
        message = BytesParser(policy=policy.default).parsebytes(
            f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("utf-8") + raw
        )
        fields: dict[str, Any] = {}
        for part in message.iter_parts():
            name = part.get_param("name", header="content-disposition")
            if not name:
                continue
            filename = part.get_filename()
            payload = part.get_payload(decode=True) or b""
            if filename:
                file_info = {"filename": filename, "content": payload}
                current = fields.get(name)
                if current is None:
                    fields[name] = file_info
                elif isinstance(current, list):
                    current.append(file_info)
                else:
                    fields[name] = [current, file_info]
            else:
                fields[name] = payload.decode(part.get_content_charset() or "utf-8")
        return fields

    def _save_upload(self, fields: Mapping[str, Any]) -> Path:
        return self._save_uploads(fields)[0]

    def _save_uploads(self, fields: Mapping[str, Any]) -> tuple[Path, ...]:
        file_info = fields.get("file")
        file_items = file_info if isinstance(file_info, list) else [file_info]
        file_items = [item for item in file_items if isinstance(item, Mapping)]
        if not file_items:
            raise ValueError("Файл для импорта не передан.")
        if len(file_items) > 2:
            raise ValueError("Можно выбрать не больше двух файлов Excel для одного импорта.")
        saved_paths: list[Path] = []
        for item in file_items:
            filename = _safe_file_name(_text(item.get("filename")) or "upload")
            target = _next_available_path(self.state.upload_root / filename)
            content = item.get("content")
            if not isinstance(content, bytes):
                raise ValueError("Файл для импорта поврежден.")
            target.write_bytes(content)
            saved_paths.append(target)
        return tuple(saved_paths)

    def _send_file(self, path: Path, content_type: str) -> None:
        data = path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        self.end_headers()
        self.wfile.write(data)

    def _send_json(self, payload: Mapping[str, Any], status: HTTPStatus = HTTPStatus.OK) -> None:
        data = json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)


class TenderVestDocsServer(ThreadingHTTPServer):
    state: LocalAppState


def make_server(
    *,
    host: str,
    port: int,
    pilot_root: Path,
    allow_network_downloads: bool = False,
    allow_network_enrichment: bool = False,
) -> TenderVestDocsServer:
    server = TenderVestDocsServer((host, port), LocalAppHandler)
    server.state = LocalAppState(
        pilot_root=pilot_root,
        allow_network_downloads=allow_network_downloads,
        allow_network_enrichment=allow_network_enrichment,
    )
    return server


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="TenderVestDocs local UI app")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--pilot-root", default=str(DEFAULT_PILOT_ROOT))
    parser.add_argument("--allow-network-downloads", action="store_true")
    parser.add_argument("--allow-network-enrichment", action="store_true")
    args = parser.parse_args(argv)

    server = make_server(
        host=args.host,
        port=args.port,
        pilot_root=Path(args.pilot_root),
        allow_network_downloads=args.allow_network_downloads,
        allow_network_enrichment=args.allow_network_enrichment,
    )
    print(f"http://{args.host}:{args.port}", flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        return 0
    finally:
        server.server_close()
    return 0


def _import_result_summary(result: Any) -> dict[str, Any]:
    return {
        "status": result.status,
        "source_file": str(result.source_file),
        "copied_source_file": str(result.copied_source_file),
        "source_files": [str(path) for path in (getattr(result, "source_files", None) or (result.source_file,))],
        "copied_source_files": [
            str(path)
            for path in (getattr(result, "copied_source_files", None) or (result.copied_source_file,))
        ],
        "normalized_registry_json_path": str(result.normalized_registry_json_path),
        "normalized_registry_csv_path": str(result.normalized_registry_csv_path),
        "normalization_report_json_path": str(result.normalization_report_json_path),
        "normalization_report_md_path": str(result.normalization_report_md_path),
        "sheet_name": result.sheet_name,
        "file_profile": result.file_profile,
        "file_profile_label": result.file_profile_label,
        "sheet_names": list(result.sheet_names),
        "dataset_counts": dict(result.dataset_counts),
        "sheet_profiles": [dict(profile) for profile in result.sheet_profiles],
        "header_row": result.header_row,
        "rows_count": result.rows_count,
        "physical_data_rows_count": result.physical_data_rows_count,
        "skipped_empty_rows_count": result.skipped_empty_rows_count,
        "columns_detected": dict(result.columns_detected),
        "missing_required_columns": list(result.missing_required_columns),
        "warnings": [warning.__dict__ for warning in result.warnings],
        "errors": [error.__dict__ for error in result.errors],
    }


def _rows_to_json_dicts(rows: list[Any] | tuple[Any, ...]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for row in rows:
        if hasattr(row, "to_json_dict"):
            result.append(dict(row.to_json_dict()))
        elif isinstance(row, Mapping):
            result.append(dict(row))
    return result


def _count_records_by_dataset(records: list[ProcurementRecord]) -> dict[str, int]:
    counts = {"procurement": 0, "contract": 0}
    for record in records:
        data = record.to_json_dict() if hasattr(record, "to_json_dict") else {}
        dataset_type = _text(data.get("source_dataset_type")) or "procurement"
        counts[dataset_type] = counts.get(dataset_type, 0) + 1
    return counts


def _now_label() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _record_by_row_id(records: list[ProcurementRecord], row_id: str) -> ProcurementRecord | None:
    return next((record for record in records if record.row_id == row_id), None)


def _download_status_label(status: str) -> str:
    labels = {
        "planned_only": "ожидание разрешения",
        "blocked_no_source": "Нет источника",
        "manual_platform_check": "Требуется ручная проверка",
        "manual_commercial_check": "Требуется ручная проверка",
        "checked_no_public_files_downloaded": "Документы не найдены в публичном источнике",
        "downloaded": "Скачано",
        "partial_download": "Частично скачано",
        "no_public_documents_found": "Документы не найдены в публичном источнике",
        "not_published": "Документы не опубликованы",
        "source_missing_manual_review": "Источник не найден",
        "unsupported_platform": "Неподдержанная площадка",
        "browser_required": "Нужен браузер",
        "auth_required": "Требуется авторизация",
        "captcha_required": "Требуется капча",
        "download_failed": "Ошибка скачивания",
        "manual_review_required": "Требуется ручная проверка",
        "skipped_by_scope": "Пропущено по области отбора",
    }
    return labels.get(status, status)


def _safe_non_negative_int(value: Any) -> int | None:
    if value is None or isinstance(value, bool):
        return None
    try:
        parsed = int(value)
    except (TypeError, ValueError):
        return None
    return max(0, parsed)


def _safe_sequence(value: Any) -> tuple[Any, ...]:
    if value is None:
        return ()
    if isinstance(value, tuple):
        return value
    if isinstance(value, list):
        return tuple(value)
    if isinstance(value, (str, bytes)):
        return (value,)
    try:
        return tuple(value)
    except TypeError:
        return ()


def _safe_row_downloaded_files(row_result: Any) -> tuple[Any, ...]:
    for field_name in ("downloaded_files", "files", "saved_files", "downloaded_file_paths"):
        files = _safe_sequence(getattr(row_result, field_name, None))
        if files:
            return files
        if hasattr(row_result, field_name):
            return ()
    return ()


def _safe_row_downloaded_files_count(row_result: Any) -> int:
    for field_name in ("download_files_count", "downloaded_files_count"):
        parsed = _safe_non_negative_int(getattr(row_result, field_name, None))
        if parsed is not None:
            return parsed
    return len(_safe_row_downloaded_files(row_result))


def _safe_row_documents_found_count(row_result: Any) -> int:
    candidates_count = _safe_non_negative_int(getattr(row_result, "candidates_count", None))
    downloaded_count = _safe_row_downloaded_files_count(row_result)
    if candidates_count is None or candidates_count == 0:
        return downloaded_count
    return candidates_count


def _safe_result_downloaded_files_count(result: Any) -> int:
    parsed = _safe_non_negative_int(getattr(result, "downloaded_files_count", None))
    if parsed is not None:
        return parsed
    return sum(_safe_row_downloaded_files_count(row) for row in getattr(result, "rows", ()))


def _downloaded_file_payload(item: Any) -> dict[str, Any]:
    if hasattr(item, "to_json_dict"):
        payload = item.to_json_dict()
        if isinstance(payload, Mapping):
            return dict(payload)
    if isinstance(item, Mapping):
        return dict(item)
    return {"path": str(item)}


def _downloaded_files_payload(rows: Any) -> list[dict[str, Any]]:
    files: list[dict[str, Any]] = []
    for row_result in rows:
        for item in _safe_row_downloaded_files(row_result):
            files.append(_downloaded_file_payload(item))
    return files


def _download_issue_payload(item: Any) -> dict[str, Any]:
    if hasattr(item, "to_json_dict"):
        payload = item.to_json_dict()
        if isinstance(payload, Mapping):
            return dict(payload)
    if isinstance(item, Mapping):
        return dict(item)
    return {"message": str(item)}


def _download_issues_payload(rows: Any) -> dict[str, list[dict[str, Any]]]:
    warnings: list[dict[str, Any]] = []
    errors: list[dict[str, Any]] = []
    for row_result in rows:
        warnings.extend(_download_issue_payload(item) for item in _safe_sequence(getattr(row_result, "warnings", ())))
        errors.extend(_download_issue_payload(item) for item in _safe_sequence(getattr(row_result, "errors", ())))
    return {"warnings": warnings, "errors": errors}


def _download_check_status_label(status: str) -> str:
    if status == "downloaded":
        return "Документы скачаны"
    if status == "partial_download":
        return "Частично"
    if status in {"no_public_documents_found", "not_published", "checked_no_public_files_downloaded"}:
        return "Нет документов"
    if status in {
        "manual_platform_check",
        "manual_commercial_check",
        "source_missing_manual_review",
        "unsupported_platform",
        "browser_required",
        "auth_required",
        "captcha_required",
        "manual_review_required",
    }:
        return "Требует ручной проверки"
    if status in {"blocked_no_source", "download_failed"}:
        return "Ошибка маршрута"
    return "Не проверено"


def _parse_json_field(value: Any) -> dict[str, str] | None:
    if not value:
        return None
    if isinstance(value, Mapping):
        return {str(key): str(item) for key, item in value.items()}
    parsed = json.loads(str(value))
    if not isinstance(parsed, Mapping):
        return None
    return {str(key): str(item) for key, item in parsed.items()}


def _clean_mapping(field_mapping: Mapping[str, str] | None) -> dict[str, str]:
    if not field_mapping:
        return {}
    result: dict[str, str] = {}
    for field_name, source_column in field_mapping.items():
        field = _text(field_name)
        column = _text(source_column)
        if field and column:
            result[field] = column
    return result


def _dataset_type_from_preview(preview: ImportPreviewResult | None) -> str:
    if preview is not None and preview.source_dataset_type == DATASET_CONTRACT:
        return DATASET_CONTRACT
    return DATASET_PROCUREMENT


def _dataset_type_from_mapping(
    field_mapping: Mapping[str, str] | None,
    *,
    fallback: str = DATASET_PROCUREMENT,
) -> str:
    fields = set(_clean_mapping(field_mapping))
    if fields & {
        "contract_registry_number",
        "contract_url",
        "supplier_name",
        "supplier_inn",
        "contract_price",
        "contract_date",
        "execution_status",
    }:
        return DATASET_CONTRACT
    if fallback == DATASET_CONTRACT:
        return DATASET_CONTRACT
    return DATASET_PROCUREMENT


def _import_session_id(path: Path) -> str:
    try:
        stat = path.stat()
        marker = f"{path.name}-{stat.st_mtime_ns}-{stat.st_size}"
    except OSError:
        marker = path.name
    safe = re.sub(r"[^0-9A-Za-zА-Яа-я_-]+", "_", marker).strip("_")
    return f"import_session_{safe[:96] or 'current'}"


def _import_session_id_for_paths(paths: Sequence[Path]) -> str:
    markers: list[str] = []
    for path in paths:
        try:
            stat = path.stat()
            markers.append(f"{path.name}-{stat.st_mtime_ns}-{stat.st_size}")
        except OSError:
            markers.append(path.name)
    safe = re.sub(r"[^0-9A-Za-zА-Яа-я_-]+", "_", "__".join(markers)).strip("_")
    return f"import_session_{safe[:96] or 'current'}"


def _file_list_label(paths: Sequence[Path]) -> str:
    return "; ".join(path.name for path in paths)


def _content_type(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".html":
        return "text/html; charset=utf-8"
    if suffix == ".css":
        return "text/css; charset=utf-8"
    if suffix == ".js":
        return "application/javascript; charset=utf-8"
    if suffix == ".json":
        return "application/json; charset=utf-8"
    if suffix == ".png":
        return "image/png"
    return "application/octet-stream"


def _safe_file_name(value: str) -> str:
    name = Path(value).name
    safe = SAFE_NAME_RE.sub("_", name).strip(" ._")
    return safe or "upload"


def _next_available_path(path: Path) -> Path:
    if not path.exists():
        return path
    suffix = 2
    while True:
        candidate = path.with_name(f"{path.stem}_{suffix}{path.suffix}")
        if not candidate.exists():
            return candidate
        suffix += 1


def _text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


if __name__ == "__main__":
    raise SystemExit(main())
