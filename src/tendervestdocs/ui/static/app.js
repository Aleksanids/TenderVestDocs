const CARD_SECTION_IDS = [
  "summary",
  "source",
  "documents",
  "actions",
  "protocols_price",
  "contract_link",
  "normative_status",
  "warnings_errors",
  "diagnostics",
];

const UNKNOWN_PRIORITY = "Приоритет не определен";
const DEFAULT_NORMATIVE_STATUS = "Не проверено";
const STAGE_LABEL = "Создайте или откройте проект, затем выберите табличный файл";
const DOCUMENTS_PLACEHOLDER =
  "Документы еще не скачаны. Для строк ЕИС доступно контролируемое скачивание.";
const UNKNOWN_EXPECTED_DOCUMENTS_TEXT = "будет определено при подготовке скачивания";
const UNCHECKED_DOCUMENTS_TEXT = "Не проверено";
const RIGHT_PANE_STORAGE_KEY = "tvd:right-pane-width";
const TABLE_COLUMN_STORAGE_PREFIX = "tvd:table-column-widths:";
const RIGHT_PANE_DEFAULT_WIDTH = 520;
const RIGHT_PANE_MIN_WIDTH = 420;
const RIGHT_PANE_MAX_WIDTH = 1180;
const RIGHT_PANE_MAX_RATIO = 0.58;
const REGISTRY_PANE_MIN_WIDTH = 760;
const TABLE_COLUMN_MIN_WIDTH = 64;
const TABLE_COLUMN_MAX_WIDTH = 520;
const DEFAULT_TABLE_COLUMN_WIDTHS = [76, 150, 76, 132, 210, 300, 126, 112, 116, 142, 128, 96, 150];
const CARD_SECTION_ICONS = {
  summary: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M8 4h8M6 8h12M6 12h8M6 16h10M6 20h6"/></svg>',
  source: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M9 18h6M10 6h4M8 4h8l2 4v10a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2V8z"/></svg>',
  documents: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M7 3h7l4 4v14H7zM14 3v5h5M9 12h6M9 16h6"/></svg>',
  actions: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3v6M12 15v6M4 12h6M14 12h6M7 7l3 3M14 14l3 3M17 7l-3 3M10 14l-3 3"/></svg>',
  protocols_price: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 5h14v14H5zM8 9h8M8 13h5M8 17h3"/></svg>',
  contract_link: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M10 7h-2a5 5 0 0 0 0 10h2M14 7h2a5 5 0 0 1 0 10h-2M9 12h6"/></svg>',
  normative_status: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3l7 3v5c0 5-3 8-7 10-4-2-7-5-7-10V6zM9 12l2 2 4-5"/></svg>',
  warnings_errors: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 4l9 16H3zM12 9v5M12 17h.01"/></svg>',
  diagnostics: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 8a4 4 0 1 1 0 8 4 4 0 0 1 0-8zM12 2v3M12 19v3M4.9 4.9l2.1 2.1M17 17l2.1 2.1M2 12h3M19 12h3M4.9 19.1l2.1-2.1M17 7l2.1-2.1"/></svg>',
};
const ACTION_ICONS = {
  refresh_row: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M20 6v6h-6M4 18v-6h6M18 10a6 6 0 0 0-10-4M6 14a6 6 0 0 0 10 4"/></svg>',
  open_link: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M14 4h6v6M20 4l-9 9M19 14v5H5V5h5"/></svg>',
  download_row: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 4v10M8 10l4 4 4-4M5 20h14"/></svg>',
  open_folder: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 7h7l2 3h9v9H3z"/></svg>',
  normative_check_all: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3l7 3v5c0 5-3 8-7 10-4-2-7-5-7-10V6zM8.5 12.5l2.5 2.5 4.5-6M4 21h16"/></svg>',
  normative_check_row: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3l7 3v5c0 5-3 8-7 10-4-2-7-5-7-10V6zM8.5 12.5l2.5 2.5 4.5-6"/></svg>',
  sheet6_export: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M7 3h10l3 3v15H7zM17 3v5h3M9 11h8M9 15h8M9 19h5M4 7h3v14H4z"/></svg>',
};
const SORT_COLUMN_KEYS = {
  "Приоритет": "priority",
  "Номер": "identifier_display",
  "Закон": "law",
  "Источник": "table_source_label",
  "Источник / маршрут": "table_source_label",
  "Заказчик": "customer",
  "Предмет": "name",
  "НМЦК": "price_sort",
  "НМЦК / НМЦД": "price_sort",
  "НМЦК / Цена": "price_sort",
  "Цена контракта": "price_sort",
  "ОКПД2": "okpd2",
  "КТРУ": "ktru",
  "Нормативный статус": "normative_status",
  "Источник статуса": "normative_status_source",
  "Документы": "documents_count_label",
  "Проверка": "check_status",
};

const FIELD_MAPPING_LABELS = [
  { field: "procurement_number", label: "Номер", required: true, requirement_label: "обязательно" },
  { field: "law_type", label: "Закон", required: true, requirement_label: "обязательно" },
  { field: "source_url", label: "Источник / ссылка", required: true, requirement_label: "обязательно для скачивания" },
  { field: "source_raw", label: "Источник / текст", required: false, requirement_label: "если нет ссылки" },
  { field: "priority", label: "Приоритет", required: false, requirement_label: "опционально" },
  { field: "nmck", label: "НМЦК", required: false, requirement_label: "желательно" },
  { field: "okpd2", label: "ОКПД2", required: false, requirement_label: "желательно" },
  { field: "ktru", label: "КТРУ", required: false, requirement_label: "опционально" },
  { field: "customer_name", label: "Заказчик", required: false, requirement_label: "желательно" },
  { field: "region", label: "Регион", required: false, requirement_label: "опционально" },
  { field: "subject", label: "Предмет", required: false, requirement_label: "желательно" },
];

const DEFAULT_FILTER_OPTIONS = {
  priorities: [UNKNOWN_PRIORITY],
  laws: ["44-ФЗ", "223-ФЗ", "Коммерческая", "Не определено", "Источник не указан"],
  normative_statuses: [DEFAULT_NORMATIVE_STATUS],
  source_types: ["url", "domain", "platform", "identifier", "mixed", "missing", "unknown"],
  documents_statuses: ["Готово к скачиванию", "Документы скачаны", "Найдены, не скачаны", "Документы не найдены", "Частично", "Нет документов", "Не проверено", "Требует ручной проверки", "Ошибка маршрута"],
  check_statuses: ["Обновлено", "Частично обновлено", "Не проверено", "Требует ручной проверки", "Ошибка обновления", "Ошибка маршрута"],
  warning_states: ["Есть предупреждения", "Есть ошибки", "Без замечаний"],
};

const EMPTY_PAYLOAD = {
  app_title: "ТендерВест Документы",
  stage: STAGE_LABEL,
  registry_title: "Реестр закупок",
  registry_status: "Не загружен",
  project_status: "Проект не создан",
  ui_contract_version: "ui_ux_contract_v1_0",
  left_nav: ["Проект", "Настройки"],
  table_columns: [
    "Приоритет",
    "Номер",
    "Закон",
    "Источник",
    "Заказчик",
    "Предмет",
    "НМЦК / НМЦД",
    "ОКПД2",
    "КТРУ",
    "Нормативный статус",
    "Источник статуса",
    "Документы",
    "Проверка",
  ],
  project_summary: {
    project_name: "Проект не выбран",
    project_root: "",
    rows_total: 0,
    rows_done: 0,
    rows_needs_review: 0,
    rows_errors: 0,
  },
  kpis: [
    { label: "Всего строк", value: 0, kind: "neutral" },
    { label: "Обработано", value: 0, kind: "success" },
    { label: "С документами", value: 0, kind: "success" },
    { label: "Частично", value: 0, kind: "warning" },
    { label: "Нет документов", value: 0, kind: "error" },
    { label: "Требует проверки", value: 0, kind: "warning" },
  ],
  top_actions: [],
  filter_options: DEFAULT_FILTER_OPTIONS,
  rows: [],
  selected_row: null,
  card_sections: [],
  last_loaded_file_name: "",
  last_imported_at: "",
};

const DEMO_PAYLOAD = EMPTY_PAYLOAD;

let payload = window.TVD_UI_PAYLOAD || EMPTY_PAYLOAD;
let selectedFile = null;
let selectedFiles = [];
let latestPreview = null;
let mappingVisible = false;
let activeSection = "Проект";
let tableColumnWidths = [];
let tableColumnStorageKey = "";
let filterPopoverOpen = false;
let pendingDownload = null;
let pendingRefresh = null;

const filterState = {
  priorities: [],
  laws: [],
  checkStatuses: [],
  query: "",
};

let registryDatasetView = "";
let selectedRowKey = "";
let sortState = { key: "", direction: "asc" };

function text(value) {
  if (Array.isArray(value)) {
    return value.length ? value.join(", ") : "—";
  }
  if (value === null || value === undefined || value === "") {
    return "—";
  }
  return String(value);
}

function datasetTypeLabel(value) {
  const normalized = String(value || "").toLowerCase();
  if (normalized === "contract") {
    return "реестр контрактов";
  }
  if (normalized === "procurement") {
    return "реестр закупок";
  }
  if (!normalized) {
    return "не определено";
  }
  return String(value);
}

const PROTOCOL_UI_VALUE_LABELS = {
  parse_success: "разобрано",
  parse_partial: "частично разобрано",
  manual_review_required: "требуется ручная проверка",
  not_found: "не найдено",
  persisted_project_artifact: "сохранено в проекте",
  temp_download_parsed_deleted: "временно разобрано, файл удален",
  supplier_result_without_supplier_name: "указан идентификатор участника, имя нужно подтвердить",
  supplier_from_embedded_contract_evidence: "поставщик подтвержден сведениями контракта",
  winner_not_disclosed_223fz: "победитель не раскрыт",
  failed_procurement_no_winner: "закупка не состоялась",
  protocol_data_extraction: "разбор протокола",
  protocol_text: "текст протокола",
  eis_supplier_results: "результат определения поставщика ЕИС",
  embedded_contract_registry_evidence: "сведения о контракте ЕИС",
  eis_44_protocol_bid_list_html: "страница протокола ЕИС «Список заявок»",
  matched: "совпадает",
  mismatch: "расхождение, нужна проверка",
  insufficient_data: "недостаточно данных",
  high: "высокая",
  medium: "средняя",
  low: "низкая",
  none: "нет",
};

function protocolDisplayValue(value) {
  if (value === null || value === undefined || value === "") {
    return "";
  }
  const raw = String(value);
  return PROTOCOL_UI_VALUE_LABELS[raw] || raw;
}

function datasetTypeForRow(row) {
  return row && row.source_dataset_type === "contract" ? "contract" : "procurement";
}

function registryDatasetCounts() {
  return rows().reduce(
    (counts, row) => {
      counts[datasetTypeForRow(row)] += 1;
      return counts;
    },
    { procurement: 0, contract: 0 },
  );
}

function ensureRegistryDatasetView() {
  const counts = registryDatasetCounts();
  if (!(counts.procurement > 0 && counts.contract > 0)) {
    registryDatasetView = "";
    return "";
  }
  if (registryDatasetView !== "procurement" && registryDatasetView !== "contract") {
    registryDatasetView = counts.procurement > 0 ? "procurement" : "contract";
  }
  if (registryDatasetView === "procurement" && counts.procurement === 0) {
    registryDatasetView = "contract";
  }
  if (registryDatasetView === "contract" && counts.contract === 0) {
    registryDatasetView = "procurement";
  }
  return registryDatasetView;
}

function rows() {
  return Array.isArray(payload.rows) ? payload.rows : [];
}

function filterOptions() {
  const options = payload.filter_options || {};
  return {
    priorities: cleanOptions(options.priorities, DEFAULT_FILTER_OPTIONS.priorities),
    laws: cleanOptions(options.laws, DEFAULT_FILTER_OPTIONS.laws),
    normative_statuses: cleanOptions(
      options.normative_statuses,
      DEFAULT_FILTER_OPTIONS.normative_statuses,
    ),
    source_types: cleanOptions(options.source_types, DEFAULT_FILTER_OPTIONS.source_types),
    documents_statuses: cleanOptions(
      options.documents_statuses,
      DEFAULT_FILTER_OPTIONS.documents_statuses,
    ),
    check_statuses: cleanOptions(options.check_statuses, DEFAULT_FILTER_OPTIONS.check_statuses),
    warning_states: cleanOptions(options.warning_states, DEFAULT_FILTER_OPTIONS.warning_states),
  };
}

function cleanOptions(values, fallbackValues) {
  const result = [];
  const source = Array.isArray(values) && values.length ? values : fallbackValues;
  source.forEach((value) => {
    const option = String(value || "").trim();
    if (option && !result.includes(option)) {
      result.push(option);
    }
  });
  return result;
}

function clampNumber(value, min, max) {
  const number = Number(value);
  if (!Number.isFinite(number)) {
    return min;
  }
  return Math.min(max, Math.max(min, number));
}

function readStoredNumber(key, fallbackValue) {
  try {
    return Number(window.localStorage.getItem(key)) || fallbackValue;
  } catch (_error) {
    return fallbackValue;
  }
}

function writeStoredNumber(key, value) {
  try {
    window.localStorage.setItem(key, String(Math.round(value)));
  } catch (_error) {
    // localStorage can be unavailable in restricted browser contexts.
  }
}

function tableColumnKey(columns) {
  return `${TABLE_COLUMN_STORAGE_PREFIX}${columns.map((column) => String(column || "")).join("|")}`;
}

function readStoredColumnWidths(key) {
  try {
    const rawValue = window.localStorage.getItem(key);
    const parsed = rawValue ? JSON.parse(rawValue) : null;
    return Array.isArray(parsed) ? parsed : [];
  } catch (_error) {
    return [];
  }
}

function writeStoredColumnWidths(key, widths) {
  try {
    window.localStorage.setItem(key, JSON.stringify(widths.map((width) => Math.round(width))));
  } catch (_error) {
    // localStorage can be unavailable in restricted browser contexts.
  }
}

function fieldTargets() {
  const targets = Array.isArray(payload.field_mapping_fields) && payload.field_mapping_fields.length
    ? payload.field_mapping_fields
    : FIELD_MAPPING_LABELS;
  return targets.map((entry) => (Array.isArray(entry) ? {
    field: entry[0],
    label: entry[1],
    required: false,
    requirement_label: "опционально",
    group: "Поля приложения",
  } : entry));
}

function targetLabel(fieldName) {
  const target = fieldTargets().find((entry) => entry.field === fieldName);
  return target ? target.label : fieldName;
}

function suggestedFieldForColumn(columnName) {
  const suggested = (latestPreview && latestPreview.suggested_field_mapping) || {};
  return Object.keys(suggested).find((field) => suggested[field] === columnName) || "";
}

function sampleValuesForColumn(columnName) {
  const samples = (latestPreview && latestPreview.sample_values) || {};
  const values = samples[columnName];
  if (Array.isArray(values) && values.length) {
    return values.slice(0, 3).join("; ");
  }
  const previewRows = latestPreview && Array.isArray(latestPreview.preview_rows) ? latestPreview.preview_rows : [];
  return previewRows
    .map((row) => row[columnName])
    .filter(Boolean)
    .slice(0, 3)
    .join("; ");
}

function rowKey(row, index) {
  if (!row) {
    return "";
  }
  if (row.row_id) {
    return `row_id:${row.row_id}`;
  }
  if (row.purchase_number) {
    return `purchase_number:${row.purchase_number}`;
  }
  if (row.primary_identifier) {
    return `primary_identifier:${row.primary_identifier}`;
  }
  return `ui-index:${index}`;
}

function rowExists(key) {
  if (!key) {
    return false;
  }
  return rows().some((row, index) => rowKey(row, index) === key);
}

function initialSelectedRowKey() {
  const allRows = rows();
  if (!allRows.length) {
    return "";
  }
  const selected = payload.selected_row || allRows[0];
  const index = allRows.indexOf(selected);
  return rowKey(selected, index >= 0 ? index : 0);
}

function currentEntry() {
  const entries = getFilteredRowEntries();
  if (!entries.length) {
    return null;
  }
  return entries.find((entry) => rowKey(entry.row, entry.index) === selectedRowKey) || entries[0];
}

function currentRow() {
  const entry = currentEntry();
  return entry ? entry.row : null;
}

function getFilteredRows() {
  return getFilteredRowEntries().map((entry) => entry.row);
}

function getFilteredRowEntries() {
  const datasetView = ensureRegistryDatasetView();
  const query = filterState.query.trim().toLowerCase();
  const selectedPriorities = new Set(filterState.priorities);
  const selectedLaws = new Set(filterState.laws);
  const selectedCheckStatuses = new Set(filterState.checkStatuses);

  const entries = rows()
    .map((row, index) => ({ row, index }))
    .filter(({ row }) => {
      if (datasetView && datasetTypeForRow(row) !== datasetView) {
        return false;
      }
      const priority = row.priority || UNKNOWN_PRIORITY;

      if (selectedPriorities.size && !selectedPriorities.has(priority)) {
        return false;
      }
      if (selectedLaws.size && !selectedLaws.has(row.law || "Не определено")) {
        return false;
      }
      const checkStatus = row.check_status || row.documents_status || "Не проверено";
      if (selectedCheckStatuses.size && !selectedCheckStatuses.has(checkStatus)) {
        return false;
      }
      if (query && !filterTextForRow(row).includes(query)) {
        return false;
      }
      return true;
    });

  return sortedEntries(entries);
}

function sortedEntries(entries) {
  if (!sortState.key) {
    return entries;
  }
  const direction = sortState.direction === "desc" ? -1 : 1;
  return [...entries].sort((left, right) => compareSortValues(
    sortValue(left.row, sortState.key),
    sortValue(right.row, sortState.key),
  ) * direction || left.index - right.index);
}

function sortValue(row, key) {
  if (key === "price_sort") {
    return parseCurrency(row.price_value ?? row.nmck ?? row.contract_price ?? row.price_display);
  }
  const source = row.source || {};
  if (key === "table_source_label") {
    return row.table_source_label || source.source_status || source.source_display || "";
  }
  return row[key] ?? "";
}

function parseCurrency(value) {
  const normalized = String(value ?? "")
    .replace(/\s+/g, "")
    .replace(/[^\d,.-]/g, "")
    .replace(",", ".");
  const number = Number.parseFloat(normalized);
  return Number.isFinite(number) ? number : 0;
}

function compareSortValues(leftValue, rightValue) {
  if (typeof leftValue === "number" || typeof rightValue === "number") {
    return Number(leftValue || 0) - Number(rightValue || 0);
  }
  return String(leftValue || "").localeCompare(String(rightValue || ""), "ru", {
    numeric: true,
    sensitivity: "base",
  });
}

function filterTextForRow(row) {
  if (row.filter_text) {
    return String(row.filter_text).toLowerCase();
  }
  const source = row.source || {};
  const parts = [
    row.identifier_display,
    row.purchase_number,
    row.primary_identifier,
    row.law,
    row.name,
    row.customer,
    source.source_display,
    source.source_text,
    source.source_url,
    row.priority,
    row.okpd2,
    row.normative_status,
    row.normative_status_source,
    row.ktru,
    row.documents_status,
    row.documents_count_label,
    row.check_status,
    row.contract_registry_number,
    row.procurement_number,
    row.supplier_name,
    row.supplier_inn,
  ];
  return parts.flat().filter(Boolean).join(" ").toLowerCase();
}

function warningStateForRow(row) {
  if (row.warning_state) {
    return row.warning_state;
  }
  if (Array.isArray(row.errors) && row.errors.length) {
    return "Есть ошибки";
  }
  if (Array.isArray(row.warnings) && row.warnings.length) {
    return "Есть предупреждения";
  }
  return "Без замечаний";
}

function badgeKind(value) {
  const normalized = String(value || "").toLowerCase();
  if (normalized.includes("запрет") || normalized.includes("ошиб") || normalized.includes("высок") || normalized === "d" || normalized === "1") {
    return "error";
  }
  if (normalized.includes("огранич")) {
    return "restriction";
  }
  if (normalized.includes("преимущ")) {
    return "benefit";
  }
  if (
    normalized.includes("предупреж")
    || normalized.includes("ручн")
    || normalized.includes("частич")
    || normalized.includes("треб")
    || normalized.includes("не скач")
    || normalized.includes("сред")
    || normalized === "b"
    || normalized === "c"
    || normalized === "2"
  ) {
    return "warning";
  }
  if (
    normalized.includes("не определ")
    || normalized.includes("не провер")
    || normalized.includes("не указано")
    || normalized.includes("нет в исходном")
    || normalized.includes("не применимо")
    || normalized.includes("документы не найд")
    || normalized.includes("нет документов")
    || normalized === "3"
  ) {
    return "neutral";
  }
  return "ok";
}

function rowAccentKind(row) {
  if (!row) {
    return "neutral";
  }
  if (row.has_errors || row.check_status === "Ошибка маршрута") {
    return "error";
  }
  const priorityKind = badgeKind(row.priority);
  if (priorityKind === "error" || priorityKind === "warning") {
    return priorityKind;
  }
  if (
    row.has_warnings
    || row.check_status === "Требует ручной проверки"
    || row.check_status === "Частично"
  ) {
    return "warning";
  }
  if (row.check_status === "Готово к скачиванию" || row.check_status === "Документы скачаны") {
    return "ok";
  }
  return priorityKind;
}

function valueAccentKind(label, value) {
  const normalizedLabel = String(label || "").toLowerCase();
  const rawValue = String(Array.isArray(value) ? value.join(" ") : value || "");
  const normalizedValue = rawValue.toLowerCase();
  if (!normalizedValue || normalizedValue === "—") {
    return "";
  }
  if (/https?:\/\//i.test(normalizedValue)) {
    return "";
  }
  const isShortValue = rawValue.length <= 28;
  if (
    isShortValue
    && (
      normalizedLabel.includes("источник")
      || normalizedLabel.includes("маршрут")
      || normalizedLabel.includes("домен")
    )
  ) {
    return "source";
  }
  if (
    isShortValue
    && (
      normalizedLabel.includes("статус")
      || normalizedLabel.includes("провер")
      || normalizedLabel.includes("мер")
      || normalizedLabel.includes("документ")
      || normalizedLabel.includes("предупреж")
      || normalizedLabel.includes("ошиб")
      || normalizedLabel.includes("снижение")
      || normalizedLabel.includes("приоритет")
    )
  ) {
    return badgeKind(normalizedValue);
  }
  if (isShortValue && normalizedLabel.includes("цена")) {
    if (
      normalizedValue.includes("требует")
      || normalizedValue.includes("не определ")
      || normalizedValue.includes("не распоз")
    ) {
      return "warning";
    }
    if (normalizedValue === "не указано") {
      return "";
    }
    return "price";
  }
  return "";
}

function syncSelectionAfterFilters(entries = getFilteredRowEntries()) {
  if (!entries.length) {
    selectedRowKey = "";
    return;
  }
  const selectedStillVisible = entries.some((entry) => rowKey(entry.row, entry.index) === selectedRowKey);
  if (!selectedRowKey || !selectedStillVisible) {
    selectedRowKey = rowKey(entries[0].row, entries[0].index);
  }
}

function renderNav() {
  const nav = document.getElementById("leftNav");
  const projectPanel = document.getElementById("sidebarProjectPanel");
  const settingsPanel = document.getElementById("sidebarSettingsPanel");
  nav.replaceChildren();
  const items = Array.isArray(payload.left_nav) ? payload.left_nav : [];
  items.forEach((item) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `nav-item nav-group${item === activeSection ? " active" : ""}`;
    const label = document.createElement("span");
    label.textContent = item;
    const chevron = document.createElement("span");
    chevron.className = "nav-chevron";
    chevron.textContent = item === activeSection ? "⌃" : "⌄";
    button.append(label, chevron);
    button.addEventListener("click", () => {
      activeSection = item;
      renderAll();
    });
    nav.append(button);
    if (item === "Проект" && projectPanel) {
      nav.append(projectPanel);
    }
    if (item === "Настройки" && settingsPanel) {
      nav.append(settingsPanel);
    }
  });
}

function renderHeader() {
  document.getElementById("appTitle").textContent = payload.registry_title || "Реестр закупок";
  const stageLabel = document.getElementById("stageLabel");
  if (stageLabel) {
    stageLabel.textContent = payload.stage || STAGE_LABEL;
    stageLabel.hidden = rows().length > 0;
  }
  const registryStatus = document.getElementById("registryStatus");
  if (registryStatus) {
    registryStatus.textContent = payload.registry_status || (rows().length ? "Загружен" : "Не загружен");
  }
  const projectStatus = document.getElementById("projectStatus");
  if (projectStatus) {
    projectStatus.textContent = payload.project_status || "Проект не создан";
  }
  const projectSummary = payload.project_summary || {};
  document.getElementById("projectName").textContent =
    projectSummary.project_name || "Проект не выбран";
  const projectRoot = projectSummary.project_root || "";
  const projectName = projectSummary.project_name || "";
  const hasProject = Boolean(projectName || projectRoot);
  const projectPath = document.getElementById("projectPath");
  if (projectPath) {
    projectPath.textContent = projectRoot;
  }
  const uploadCard = document.getElementById("registryUploadCard");
  if (uploadCard) {
    uploadCard.classList.toggle("ready", hasProject);
    uploadCard.classList.toggle("loaded", Boolean(payload.last_loaded_file_name));
  }
  const uploadTargetProject = document.getElementById("uploadTargetProject");
  if (uploadTargetProject) {
    uploadTargetProject.textContent = hasProject
      ? `Куда: ${projectName || "текущий проект"}`
      : "Куда: проект не выбран";
  }
  const totalRows = String((projectSummary.rows_total ?? rows().length) || 0);
  const headerRowsCount = document.getElementById("headerRowsCount");
  if (headerRowsCount) {
    headerRowsCount.textContent = totalRows;
  }
  const lastLoadedFile = document.getElementById("lastLoadedFile");
  if (lastLoadedFile) {
    lastLoadedFile.textContent = payload.last_loaded_file_name || "не выбран";
  }
  const lastImportTime = document.getElementById("lastImportTime");
  if (lastImportTime) {
    lastImportTime.textContent = payload.last_imported_at || "не выполнен";
  }
  const pathInput = document.getElementById("projectPathInput");
  if (pathInput && pathInput.type !== "hidden" && !pathInput.value && payload.default_project_root) {
    pathInput.placeholder = payload.default_project_root;
  }
}

function renderKpis() {
  const root = document.getElementById("kpiRow");
  root.replaceChildren();
  const kpis = Array.isArray(payload.kpis) ? payload.kpis : [];
  kpis.forEach((kpi) => {
    const itemEl = document.createElement("div");
    itemEl.className = `kpi-item ${kpi.kind || "neutral"}`;
    const labelEl = document.createElement("span");
    labelEl.textContent = kpi.label || "";
    const valueEl = document.createElement("strong");
    valueEl.textContent = String(kpi.value ?? 0);
    itemEl.append(labelEl, valueEl);
    root.append(itemEl);
  });
}

function renderRegistryTabs() {
  const root = document.getElementById("registryTabs");
  if (!root) {
    return;
  }
  const counts = registryDatasetCounts();
  const mixed = counts.procurement > 0 && counts.contract > 0;
  root.hidden = !mixed;
  root.replaceChildren();
  if (!mixed) {
    return;
  }
  const activeDataset = ensureRegistryDatasetView();
  [
    { dataset: "procurement", label: "Реестр закупок", count: counts.procurement },
    { dataset: "contract", label: "Реестр контрактов", count: counts.contract },
  ].forEach((tab) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = tab.dataset === activeDataset ? "registry-tab active" : "registry-tab";
    button.setAttribute("aria-pressed", String(tab.dataset === activeDataset));
    button.dataset.registryDataset = tab.dataset;
    button.textContent = `${tab.label}: ${tab.count}`;
    button.addEventListener("click", () => {
      if (registryDatasetView === tab.dataset) {
        return;
      }
      registryDatasetView = tab.dataset;
      selectedRowKey = "";
      renderRegistryTabs();
      renderFilters();
      renderRegistry();
    });
    root.append(button);
  });
}

function renderTopActions() {
  const root = document.getElementById("topActions");
  root.replaceChildren();
  const actions = Array.isArray(payload.top_actions) ? payload.top_actions : [];
  root.hidden = actions.length === 0;
  actions.forEach((action) => {
    const button = document.createElement("button");
    button.type = "button";
    button.disabled = action.enabled === false;
    button.textContent = action.label || "Действие недоступно";
    if (action.action === "refresh_sources_all") {
      button.classList.add("primary-action");
    }
    if (action.action === "download_current_filter") {
      button.classList.add("download-action");
    }
    if (action.action === "normative_check_all") {
      button.classList.add("normative-action");
    }
    if (action.action === "sheet6_export") {
      button.classList.add("sheet6-action");
    }
    if (action.disabled_reason) {
      button.title = action.disabled_reason;
    }
    button.addEventListener("click", () => handleTopAction(action));
    root.append(button);
  });
}

function renderFilters() {
  const options = filterOptions();
  const toggle = document.getElementById("filtersToggle");
  const popover = document.getElementById("filterPopover");
  if (toggle && !toggle.dataset.ready) {
    toggle.dataset.ready = "true";
    toggle.addEventListener("click", () => {
      filterPopoverOpen = !filterPopoverOpen;
      renderFilters();
    });
  }
  if (popover) {
    popover.hidden = !filterPopoverOpen;
  }
  if (toggle) {
    toggle.setAttribute("aria-expanded", String(filterPopoverOpen));
  }
  renderPriorityFilter(options.priorities);
  renderCheckboxGroup("lawFilters", options.laws, filterState.laws, (nextValues) => {
    filterState.laws = nextValues;
    renderFilterScopeInfo();
  });
  renderCheckboxGroup("checkStatusFilters", options.check_statuses, filterState.checkStatuses, (nextValues) => {
    filterState.checkStatuses = nextValues;
    renderFilterScopeInfo();
  }, true);
  renderFilterScopeInfo();

  const search = document.getElementById("filterSearch");
  search.value = filterState.query;
  search.oninput = (event) => {
    filterState.query = event.target.value;
    renderRegistry();
  };

  const resetButton = document.getElementById("resetFilters");
  if (resetButton) {
    resetButton.onclick = resetFilters;
  }
  const applyButton = document.getElementById("applyFilters");
  if (applyButton) {
    applyButton.onclick = () => {
      filterPopoverOpen = false;
      renderFilters();
      renderRegistry();
    };
  }
}

function renderPriorityFilter(priorities) {
  renderCheckboxGroup("priorityFilters", priorities, filterState.priorities, (nextValues) => {
    filterState.priorities = nextValues;
    renderFilterScopeInfo();
  }, true);
}

function renderCheckboxGroup(id, values, selectedValues, onChange, useBadges = false) {
  const root = document.getElementById(id);
  if (!root) {
    return;
  }
  root.replaceChildren();
  values.forEach((value) => {
    const label = document.createElement("label");
    label.className = useBadges ? "filter-check badge-check" : "filter-check";
    const input = document.createElement("input");
    input.type = "checkbox";
    input.value = value;
    input.checked = selectedValues.includes(value);
    input.addEventListener("change", () => {
      const nextValues = [...root.querySelectorAll("input:checked")].map((node) => node.value);
      onChange(nextValues);
      renderActiveFilters();
    });
    const span = document.createElement("span");
    span.className = useBadges ? `badge priority ${badgeKind(value)}` : "";
    span.textContent = value;
    label.append(input, span);
    root.append(label);
  });
}

function renderFilterScopeInfo() {
  const root = document.getElementById("filterScopeInfo");
  if (!root) {
    return;
  }
  const count = getFilteredRowEntries().length;
  root.textContent = `Будет показано строк: ${count}`;
}

function renderSelect(id, values, placeholder, selectedValue, onChange) {
  const select = document.getElementById(id);
  select.replaceChildren();

  const emptyOption = document.createElement("option");
  emptyOption.value = "";
  emptyOption.textContent = placeholder;
  select.append(emptyOption);

  values.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    select.append(option);
  });

  select.value = selectedValue;
  select.onchange = (event) => onChange(event.target.value);
}

function renderRegistry() {
  ensureRegistryDatasetView();
  const entries = getFilteredRowEntries();
  syncSelectionAfterFilters(entries);
  renderTable(entries);
  renderCounters(entries);
  renderActiveFilters();
  renderCard();
}

function renderTable(filteredEntries = getFilteredRowEntries()) {
  const head = document.getElementById("tableHead");
  head.replaceChildren();
  const columns = Array.isArray(payload.table_columns) ? payload.table_columns : [];
  ensureTableColumnWidths(columns);
  renderTableColGroup(columns);
  columns.forEach((column, columnIndex) => {
    const th = document.createElement("th");
    const sortKey = SORT_COLUMN_KEYS[column] || "";
    const label = document.createElement(sortKey ? "button" : "span");
    label.className = sortKey ? "column-sort" : "column-label";
    if (sortKey) {
      label.type = "button";
      label.setAttribute("aria-label", `Сортировать по колонке ${column}`);
      label.addEventListener("click", () => toggleSort(sortKey));
    }
    label.textContent = sortLabel(column, sortKey);
    const resizer = document.createElement("span");
    resizer.className = "column-resizer";
    resizer.setAttribute("role", "separator");
    resizer.setAttribute("aria-orientation", "vertical");
    resizer.setAttribute("aria-label", `Изменить ширину колонки ${column}`);
    resizer.tabIndex = 0;
    resizer.addEventListener("pointerdown", (event) => startColumnResize(event, columnIndex));
    resizer.addEventListener("keydown", (event) => resizeColumnWithKeyboard(event, columnIndex));
    th.append(label, resizer);
    head.append(th);
  });

  const body = document.getElementById("tableBody");
  body.replaceChildren();
  if (!rows().length || !filteredEntries.length) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = Math.max(columns.length, 1);
    td.className = "empty-table";
    td.textContent = rows().length ? "Нет строк по выбранным фильтрам" : "Строки не загружены";
    tr.append(td);
    body.append(tr);
    return;
  }

  filteredEntries.forEach(({ row, index }) => {
    const tr = document.createElement("tr");
    const key = rowKey(row, index);
    tr.className = `row-accent-${rowAccentKind(row)}${key === selectedRowKey ? " selected" : ""}`;
    tr.tabIndex = 0;
    tr.addEventListener("click", () => selectRow(key));
    tr.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        selectRow(key);
      }
    });

    appendBadgeCell(tr, row.priority, row.priority_badge ? row.priority_badge.kind : badgeKind(row.priority));
    appendTextCell(tr, row.identifier_display);
    appendTextCell(tr, row.law);
    appendSourceCell(tr, row);
    appendTextCell(tr, row.customer);
    appendTextCell(tr, row.name);
    appendTextCell(tr, row.price_display || row.nmck);
    appendOkpd2Cell(tr, row);
    appendKtruCell(tr, row);
    appendBadgeCell(tr, row.normative_status || DEFAULT_NORMATIVE_STATUS, badgeKind(row.normative_status || DEFAULT_NORMATIVE_STATUS));
    appendBadgeCell(tr, row.normative_status_source || DEFAULT_NORMATIVE_STATUS, "source");
    appendBadgeCell(tr, documentsTableLabel(row), badgeKind(row.documents_status || documentsTableLabel(row) || "Не проверено"));
    appendBadgeCell(tr, row.check_status || "Не проверено", badgeKind(row.check_status || "Не проверено"));
    body.append(tr);
  });
}

function toggleSort(key) {
  if (sortState.key === key) {
    sortState.direction = sortState.direction === "asc" ? "desc" : "asc";
  } else {
    sortState = { key, direction: "asc" };
  }
  renderRegistry();
}

function sortLabel(column, key) {
  if (!key || sortState.key !== key) {
    return column;
  }
  return `${column} ${sortState.direction === "asc" ? "↑" : "↓"}`;
}

function ensureTableColumnWidths(columns) {
  const nextKey = tableColumnKey(columns);
  if (tableColumnStorageKey === nextKey && tableColumnWidths.length === columns.length) {
    return;
  }
  tableColumnStorageKey = nextKey;
  const storedWidths = readStoredColumnWidths(nextKey);
  tableColumnWidths = columns.map((_, index) => {
    const fallbackWidth = DEFAULT_TABLE_COLUMN_WIDTHS[index] || 140;
    return clampNumber(storedWidths[index] || fallbackWidth, TABLE_COLUMN_MIN_WIDTH, TABLE_COLUMN_MAX_WIDTH);
  });
}

function renderTableColGroup(columns) {
  const colGroup = document.getElementById("tableColGroup");
  const table = colGroup ? colGroup.closest("table") : null;
  if (!colGroup || !table) {
    return;
  }
  const renderedWidths = effectiveTableColumnWidths(table);
  colGroup.replaceChildren();
  columns.forEach((_, index) => {
    const col = document.createElement("col");
    col.style.width = `${Math.round(renderedWidths[index])}px`;
    colGroup.append(col);
  });
  table.style.minWidth = `${Math.round(renderedWidths.reduce((sum, width) => sum + width, 0))}px`;
  table.style.width = "100%";
}

function effectiveTableColumnWidths(table) {
  const baseWidths = tableColumnWidths.length ? tableColumnWidths : DEFAULT_TABLE_COLUMN_WIDTHS;
  const baseTotal = baseWidths.reduce((sum, width) => sum + width, 0);
  if (!baseTotal) {
    return [];
  }
  const tableWrap = table.closest(".table-wrap");
  const availableWidth = Math.floor(tableWrap ? tableWrap.clientWidth : 0);
  const minimumWidth = baseWidths.length * TABLE_COLUMN_MIN_WIDTH;
  const targetWidth = Math.max(availableWidth, minimumWidth);
  const scale = targetWidth / baseTotal;
  return baseWidths.map((width) => clampNumber(width * scale, TABLE_COLUMN_MIN_WIDTH, TABLE_COLUMN_MAX_WIDTH));
}

function startColumnResize(event, columnIndex) {
  event.preventDefault();
  event.stopPropagation();
  const startX = event.clientX;
  const startWidth = tableColumnWidths[columnIndex] || DEFAULT_TABLE_COLUMN_WIDTHS[columnIndex] || 140;
  document.body.classList.add("resizing-column");

  const resizeFromPointer = (moveEvent) => {
    const nextWidth = startWidth + (moveEvent.clientX - startX);
    setTableColumnWidth(columnIndex, nextWidth);
  };

  const finishResize = () => {
    document.body.classList.remove("resizing-column");
    document.removeEventListener("pointermove", resizeFromPointer);
    writeStoredColumnWidths(tableColumnStorageKey, tableColumnWidths);
  };

  document.addEventListener("pointermove", resizeFromPointer);
  document.addEventListener("pointerup", finishResize, { once: true });
}

function resizeColumnWithKeyboard(event, columnIndex) {
  const step = event.shiftKey ? 40 : 12;
  if (event.key === "ArrowLeft") {
    event.preventDefault();
    setTableColumnWidth(columnIndex, (tableColumnWidths[columnIndex] || 140) - step, true);
  }
  if (event.key === "ArrowRight") {
    event.preventDefault();
    setTableColumnWidth(columnIndex, (tableColumnWidths[columnIndex] || 140) + step, true);
  }
}

function setTableColumnWidth(columnIndex, width, persist = false) {
  tableColumnWidths[columnIndex] = clampNumber(width, TABLE_COLUMN_MIN_WIDTH, TABLE_COLUMN_MAX_WIDTH);
  renderTableColGroup(Array.isArray(payload.table_columns) ? payload.table_columns : []);
  if (persist) {
    writeStoredColumnWidths(tableColumnStorageKey, tableColumnWidths);
  }
}

function appendTextCell(tr, value) {
  const td = document.createElement("td");
  td.textContent = text(value);
  tr.append(td);
}

function appendOkpd2Cell(tr, row) {
  const td = document.createElement("td");
  td.className = "code-value-cell";
  const codes = okpd2CodesForRow(row);
  if (!codes.length) {
    appendCodePlaceholder(td, row.okpd2);
    tr.append(td);
    return;
  }

  const wrapper = document.createElement("div");
  wrapper.className = "table-code-links";
  codes.forEach((code) => {
    wrapper.append(makeCodeChip(code));
  });
  td.append(wrapper);
  tr.append(td);
}

function appendKtruCell(tr, row) {
  const td = document.createElement("td");
  td.className = "code-value-cell";
  const codes = ktruCodesForRow(row);
  if (!codes.length) {
    appendCodePlaceholder(td, row.ktru);
    tr.append(td);
    return;
  }

  const wrapper = document.createElement("div");
  wrapper.className = "table-code-links";
  codes.forEach((code) => {
    wrapper.append(makeExternalLink(ktruUrlForCode(code), code, "table-code-link"));
  });
  td.append(wrapper);
  tr.append(td);
}

function appendCodePlaceholder(td, value) {
  const span = document.createElement("span");
  span.className = `table-code-placeholder ${badgeKind(value)}`;
  span.textContent = text(value);
  td.append(span);
}

function makeCodeChip(code) {
  const span = document.createElement("span");
  span.className = "table-code-chip";
  span.textContent = code;
  return span;
}

function okpd2CodesForRow(row) {
  const values = [];
  const append = (value) => {
    if (Array.isArray(value)) {
      value.forEach(append);
      return;
    }
    String(value || "")
      .split(/[,;\s]+/)
      .map((item) => item.trim())
      .filter(Boolean)
      .forEach((item) => {
        if (/^\d{2}\.\d{2}(?:\.\d{1,3}){0,2}$/.test(item) && !values.includes(item)) {
          values.push(item);
        }
      });
  };
  append(row.okpd2);
  const normative = row.normative || {};
  append(normative.okpd2_source);
  append(normative.linked_okpd2_codes);
  return values;
}

function ktruCodesForRow(row) {
  const values = [];
  const append = (value) => {
    if (Array.isArray(value)) {
      value.forEach(append);
      return;
    }
    String(value || "")
      .split(/[,;\s]+/)
      .map((item) => item.trim())
      .filter(Boolean)
      .forEach((item) => {
        if (/^(?:\d{2}\.\d{2}\.\d{2}\.\d{3,6}-\d{8}|\d{21})$/.test(item) && !values.includes(item)) {
          values.push(item);
        }
      });
  };
  append(row.ktru);
  const card = row.ktru_card || {};
  append(card.code);
  append(card.codes);
  return values;
}

function ktruUrlForCode(code) {
  const prefix = "https:" + "//zakupki.gov.ru/epz/ktru/ktruCard/commonInfo.html?itemId=";
  return `${prefix}${encodeURIComponent(code)}`;
}

function appendSourceCell(tr, row) {
  const td = document.createElement("td");
  const source = row.source || {};
  const label = row.table_source_label || source.source_status || row.route_source_label || source.source_display || "Источник не указан";
  const url = sourceUrlForRow(row);
  const wrapper = document.createElement("div");
  wrapper.className = "source-cell";
  if (url) {
    wrapper.append(makeExternalLink(url, label, "badge source source-link"));
  } else {
    const span = document.createElement("span");
    span.className = "badge source";
    span.textContent = text(label);
    wrapper.append(span);
  }
  const refreshLabel = rowRefreshLabel(row);
  if (refreshLabel) {
    const chip = document.createElement("span");
    chip.className = "source-refresh-chip";
    chip.textContent = refreshLabel;
    wrapper.append(chip);
  }
  td.append(wrapper);
  tr.append(td);
}

function rowRefreshLabel(row) {
  if (row.source_refresh_status === "Сведения обновлены" || row.source_update_label) {
    return "Обновлено";
  }
  return "";
}

function documentsTableLabel(row) {
  const value = row.documents_count_label || row.documents_status || UNCHECKED_DOCUMENTS_TEXT;
  return text(value);
}

function appendBadgeCell(tr, value, kind) {
  const td = document.createElement("td");
  const span = document.createElement("span");
  span.className = `badge ${kind || "neutral"}`;
  span.dataset.badgeKind = kind || "neutral";
  span.textContent = text(value);
  td.append(span);
  tr.append(td);
}

function sourceUrlForRow(row) {
  const source = row.source || {};
  const route = row.route || {};
  return firstAllowedPrimarySourceUrl([
    row.contract_url,
    source.contract_url,
    source.source_url,
    route.source_url,
    source.generated_eis_url,
    route.generated_eis_url,
    source.generated_platform_url,
    route.generated_platform_url,
    source.normalized_source_url,
    route.normalized_source_url,
  ]);
}

function firstAllowedPrimarySourceUrl(values) {
  for (const value of values) {
    const normalized = normalizeHttpUrl(value);
    if (normalized && isAllowedPrimarySourceUrl(normalized)) {
      return normalized;
    }
  }
  return "";
}

function firstHttpUrl(values) {
  for (const value of values) {
    const normalized = normalizeHttpUrl(value);
    if (normalized) {
      return normalized;
    }
  }
  return "";
}

function normalizeHttpUrl(value) {
  const candidate = String(value || "").trim();
  return /^https?:\/\//i.test(candidate) ? candidate : "";
}

function isAllowedPrimarySourceUrl(value) {
  try {
    const url = new URL(value);
    return !url.hostname.toLowerCase().includes("rostender");
  } catch (_error) {
    return false;
  }
}

function makeExternalLink(url, label, className = "inline-link") {
  const link = document.createElement("a");
  link.href = url;
  link.target = "_blank";
  link.rel = "noopener noreferrer";
  link.className = className;
  link.textContent = text(label || url);
  link.title = `Открыть первоисточник: ${url}`;
  link.addEventListener("click", (event) => {
    event.stopPropagation();
  });
  return link;
}

function renderCounters(filteredEntries = getFilteredRowEntries()) {
  document.getElementById("totalRowsCount").textContent = String(rows().length);
  document.getElementById("visibleRowsCount").textContent = String(filteredEntries.length);
  const empty = document.getElementById("filteredEmptyState");
  empty.hidden = Boolean(filteredEntries.length) || !rows().length;
}

function renderActiveFilters() {
  const root = document.getElementById("activeFilters");
  root.replaceChildren();
  const filters = activeFilterItems();
  document.getElementById("resetFilters").disabled = filters.length === 0;

  if (!filters.length) {
    const empty = document.createElement("span");
    empty.className = "active-filter-empty";
    empty.textContent = "Фильтры не выбраны";
    root.append(empty);
    return;
  }

  filters.forEach((filter) => {
    const chip = document.createElement("button");
    chip.type = "button";
    chip.className = "active-filter";
    chip.textContent = filter.text;
    chip.addEventListener("click", () => {
      filter.clear();
      renderFilters();
      renderRegistry();
    });
    root.append(chip);
  });
}

function activeFilterItems() {
  const items = filterState.priorities.map((priority) => ({
    text: `Приоритет: ${priority}`,
    clear: () => {
      filterState.priorities = filterState.priorities.filter((item) => item !== priority);
    },
  }));
  filterState.laws.forEach((law) => {
    items.push({
      text: `Закон: ${law}`,
      clear: () => {
        filterState.laws = filterState.laws.filter((item) => item !== law);
      },
    });
  });
  filterState.checkStatuses.forEach((checkStatus) => {
    items.push({
      text: `Проверка: ${checkStatus}`,
      clear: () => {
        filterState.checkStatuses = filterState.checkStatuses.filter((item) => item !== checkStatus);
      },
    });
  });
  if (filterState.query.trim()) {
    items.push({ text: `Поиск: ${filterState.query.trim()}`, clear: () => (filterState.query = "") });
  }
  return items;
}

function resetFilters() {
  filterState.priorities = [];
  filterState.laws = [];
  filterState.checkStatuses = [];
  filterState.query = "";
  selectedRowKey = initialSelectedRowKey();
  renderFilters();
  renderRegistry();
}

// Python build_card_sections() is the source of truth.
// This JS function is fallback only and must be kept in sync with Python section ids.
function cardSections(row) {
  const source = row.source || {};
  const route = row.route || {};
  const documents = row.documents || {};
  const protocols = row.protocols || {};
  const participantSummary = protocols.participants_summary || {};
  const priceSummary = protocols.price_summary || {};
  const winnerSummary = protocols.winner_summary || {};
  const supplierEvidence = protocols.supplier_result_evidence || {};
  const contractEvidence = protocols.embedded_contract_evidence || {};
  const warnings = (row.warnings || []).map((item) => item.message || item.code || "");
  const errors = (row.errors || []).map((item) => item.message || item.code || "");
  const isContract = row.source_dataset_type === "contract";
  let protocolsText = isContract ? "Протоколы по закупке-основанию не анализировались." : "";
  if (winnerSummary.winner_display_status === "supplier_result_without_supplier_name") {
    protocolsText = "Победитель по результатам определения поставщика указан как идентификатор участника. Полное наименование поставщика требуется подтвердить по сведениям о контракте.";
  } else if (winnerSummary.winner_display_status === "supplier_from_embedded_contract_evidence") {
    protocolsText = "В сведениях ЕИС найдено подтверждение по контракту. Полная связь закупка ↔ контракт будет обрабатываться на этапе 28-30.";
  } else if (winnerSummary.winner_display_status === "winner_not_disclosed_223fz") {
    protocolsText = "По 223-ФЗ победитель не раскрыт в доступных протоколах; это не считается ошибкой.";
  }

  return [
    {
      id: CARD_SECTION_IDS[0],
      title: "Сводка",
      default_open: true,
      kind: "key_value",
      items: [
        item(isContract ? "Номер контракта" : "Номер закупки", row.identifier_display || row.contract_registry_number || row.procurement_number),
        item(isContract ? "Номер закупки" : "Закон", isContract ? row.procurement_number : row.law),
        item("Предмет", row.name),
        item("Заказчик", row.customer),
        item(isContract ? "Цена контракта" : row.price_column_label || "НМЦК / НМЦД", row.price_display || row.contract_price || row.nmck),
        item(isContract ? "Поставщик" : "ОКПД2", isContract ? row.supplier_name : row.okpd2),
        item(isContract ? "Дата контракта" : "КТРУ", isContract ? row.contract_date : row.ktru),
        item("Статус строки", row.source_update_label || row.source_refresh_status || row.check_status),
        item("Следующее действие", row.next_action || row.check_status_reason),
      ],
    },
    {
      id: CARD_SECTION_IDS[1],
      title: "Источник",
      default_open: false,
      kind: "key_value",
      items: [
        item("Тип источника", row.card_source_label || source.card_source_status || source.source_status || "Источник не определен"),
        item(isContract ? "Ссылка на карточку контракта / договора" : "Прямая ссылка", row.contract_url || source.contract_url || source.source_url, Boolean(row.contract_url || source.contract_url || source.source_url)),
        item("Маршрут", route.route_display || source.source_display || "Не указано"),
        item("Статус источника", row.check_status || "Не проверено"),
        item("Следующее действие", route.next_action || row.next_action),
        item("Домен / площадка", route.normalized_domain || route.platform_name || source.normalized_domain),
        item("Номер ЕИС / номер договора", source.eis_number || row.procurement_number || row.contract_registry_number),
      ],
    },
    {
      id: CARD_SECTION_IDS[2],
      title: "Документы",
      default_open: false,
      kind: "placeholder",
      text: documents.text || DOCUMENTS_PLACEHOLDER,
      items: [
        item("Статус", documents.label || row.documents_status),
        item("Счетчик", documents.count_label || row.documents_count_label || UNCHECKED_DOCUMENTS_TEXT),
        item("Папка проекта", documents.folder_path),
      ],
    },
    {
      id: CARD_SECTION_IDS[3],
      title: "Действия",
      default_open: false,
      kind: "actions",
      actions: row.actions || [],
    },
    {
      id: CARD_SECTION_IDS[4],
      title: "Протоколы / участники / цена",
      default_open: false,
      kind: "key_value",
      text: protocolsText,
      items: [
        item("Статус протоколов", protocolDisplayValue(protocols.status)),
        item(
          "Режим хранения доказательств",
          protocolDisplayValue(protocols.evidence_storage_status
            || supplierEvidence.evidence_storage_status
            || contractEvidence.evidence_storage_status),
        ),
        item("Протоколов найдено", protocols.found_count),
        item("Подано заявок / участников", participantSummary.applications_total_count || participantSummary.participants_total_count),
        item("Допущено", participantSummary.admitted_count),
        item("Отклонено", participantSummary.rejected_count),
        item("Победитель / отображение", winnerSummary.winner_display_name || "Не подтверждено"),
        item("Статус победителя", protocolDisplayValue(winnerSummary.winner_display_status)),
        item("Источник победителя", protocolDisplayValue(winnerSummary.winner_source)),
        item("Цена результата поставщика", priceSummary.supplier_result_offer_price),
        item("Цена контракта", priceSummary.contract_price),
        item("Согласованность цены", protocolDisplayValue(priceSummary.price_consistency_status)),
        item("Результаты поставщика: участник / идентификатор", supplierEvidence.participant_identifier),
        item("Результаты поставщика: статус", supplierEvidence.rank_text),
        item("Результаты поставщика: предложение", supplierEvidence.participant_offer_price),
        item("Результаты поставщика: протокол", supplierEvidence.protocol_name),
        item("Сведения контракта: поставщик", contractEvidence.contract_supplier_name),
        item("Сведения контракта: контракт", contractEvidence.contract_registry_number),
        item("Сведения контракта: цена", contractEvidence.contract_price),
        item("Ручная проверка", protocols.manual_review_required ? "Да" : "Нет"),
      ],
    },
    {
      id: CARD_SECTION_IDS[5],
      title: isContract ? "Связь с закупкой" : "Связь с контрактом",
      default_open: false,
      kind: "key_value",
      text: isContract && !row.procurement_number ? "Связь с закупкой не проверялась." : "",
      items: [
        item(isContract ? "Номер закупки / извещения" : "Номер контракта", isContract ? row.procurement_number : row.contract_number),
        item("Статус связи", row.contract_link_status || row.procurement_link_status || "Не проверено"),
        item("Основание связи", row.contract_link_reason || row.procurement_link_reason),
        item(isContract ? "Ссылка на закупку" : "Ссылка на контракт", row.procurement_url || row.contract_url, Boolean(row.procurement_url || row.contract_url)),
      ],
    },
    {
      id: CARD_SECTION_IDS[6],
      title: "Нормативный статус",
      default_open: false,
      kind: "key_value",
      items: [
        item("Нормативный статус", row.normative_status || DEFAULT_NORMATIVE_STATUS),
        item("Источник статуса", row.normative_status_source || DEFAULT_NORMATIVE_STATUS),
      ],
    },
    {
      id: CARD_SECTION_IDS[7],
      title: "Предупреждения",
      default_open: false,
      kind: "messages",
      items: [
        item("Предупреждения", warnings.length ? warnings : "Предупреждений нет"),
        item("Ошибки", errors.length ? errors : "Ошибок нет"),
      ],
    },
    {
      id: CARD_SECTION_IDS[8],
      title: "Диагностика",
      default_open: false,
      kind: "key_value",
      items: [
        item("Тип реестра", datasetTypeLabel(row.source_dataset_type)),
        item("Реестровый номер контракта", row.contract_registry_number),
        item("Номер закупки", row.procurement_number),
        item("Ссылка на контракт", row.contract_url, Boolean(row.contract_url)),
        item("Поставщик", row.supplier_name),
        item("ИНН поставщика", row.supplier_inn),
        item("Цена контракта", row.contract_price || row.price_display),
        item("Дата контракта", row.contract_date),
        item("Статус исполнения", row.execution_status),
        item("Код маршрута", route.route || row.route_code),
        item("Тип источника", route.source_type || row.source_type),
        item("Статус маршрута", route.route_status || row.route_status),
        item("Количество документов", row.documents_count_label),
        item("Последнее обновление", row.last_updated_at || row.last_imported_at),
      ],
    },
  ];
}

function item(label, value, copyable = false) {
  return {
    label,
    value,
    empty: isEmpty(value),
    copyable,
  };
}

function isEmpty(value) {
  if (Array.isArray(value)) {
    return value.length === 0;
  }
  return value === null || value === undefined || value === "";
}

function sectionsForRow(row) {
  if (row && Array.isArray(row.card_sections) && row.card_sections.length) {
    return row.card_sections;
  }
  return row ? cardSections(row) : [];
}

function renderCard() {
  const row = currentRow();
  const root = document.getElementById("cardSections");
  root.replaceChildren();

  if (!getFilteredRows().length || !row) {
    document.getElementById("selectedIdentifier").textContent = "Строки не выбраны";
    const empty = document.createElement("p");
    empty.className = "placeholder empty-card";
    empty.textContent = rows().length
      ? "Нет строк по выбранным фильтрам."
      : "Выберите строку после импорта реестра.";
    root.append(empty);
    return;
  }

  document.getElementById("selectedIdentifier").textContent = row.identifier_display || "Идентификатор не указан";
  sectionsForRow(row).forEach((section) => {
    root.append(renderSection(section));
  });
}

function renderSection(section) {
  const details = document.createElement("details");
  details.open = Boolean(section.default_open || section.expanded);
  details.dataset.sectionId = section.id || "";
  details.className = section.kind ? `card-section ${section.kind}` : "card-section";

  const summary = document.createElement("summary");
  summary.append(
    makeIcon(CARD_SECTION_ICONS[section.id] || CARD_SECTION_ICONS.summary, "section-icon"),
    makeTextSpan(section.title || "Раздел карточки", "section-title"),
  );
  details.append(summary);

  const body = document.createElement("div");
  body.className = "section-body";

  if (section.text) {
    const p = document.createElement("p");
    p.className = "placeholder";
    p.textContent = section.text;
    body.append(p);
  }

  if (Array.isArray(section.items)) {
    section.items.forEach((entry) => {
      body.append(renderItem(entry));
    });
  }

  if (Array.isArray(section.actions)) {
    body.append(renderActions(section.actions));
  }

  details.append(body);
  return details;
}

function renderItem(entry) {
  const itemData = Array.isArray(entry)
    ? { label: entry[0], value: entry[1], empty: isEmpty(entry[1]), copyable: false }
    : entry;
  const accent = valueAccentKind(itemData.label, itemData.value);
  const rowEl = document.createElement("div");
  rowEl.className = `kv-row${itemData.empty ? " empty-value" : ""}${itemData.copyable ? " copyable" : ""}${accent ? ` accent-${accent}` : ""}`;

  const labelEl = document.createElement("span");
  labelEl.textContent = itemData.label;

  const valueEl = document.createElement("span");
  if (Array.isArray(itemData.links) && itemData.links.length) {
    valueEl.classList.add("inline-link-list");
    itemData.links.forEach((linkData) => {
      const linkUrl = normalizeHttpUrl(linkData && linkData.url);
      if (linkUrl) {
        valueEl.append(makeExternalLink(linkUrl, linkData.label || linkUrl));
      }
    });
    if (!valueEl.childNodes.length) {
      valueEl.textContent = text(itemData.value);
    }
  } else {
    const url = normalizeHttpUrl(itemData.value);
    if (url) {
      valueEl.append(makeExternalLink(url, itemData.value));
    } else {
      if (accent) {
        valueEl.classList.add("value-chip", accent);
      }
      valueEl.textContent = text(itemData.value);
    }
  }

  rowEl.append(labelEl, valueEl);
  return rowEl;
}

function renderActions(sectionActions) {
  const actions = document.createElement("div");
  actions.className = "action-list";
  sectionActions.forEach((action) => {
    const button = document.createElement("button");
    button.type = "button";
    button.disabled = action.enabled === false;
    button.className = `row-action ${action.action || "generic"}`;
    const icon = ACTION_ICONS[action.action];
    if (icon) {
      button.append(makeIcon(icon, "action-icon"));
    }
    button.append(makeTextSpan(action.text || action.label || "Действие недоступно", "action-label"));
    if (action.enabled === false) {
      button.dataset.enabled = "false";
      if (action.disabled_reason) {
        button.title = action.disabled_reason;
      }
    }
    button.addEventListener("click", () => handleRowAction(action, currentRow()));
    actions.append(button);
  });
  return actions;
}

function makeIcon(markup, className) {
  const icon = document.createElement("span");
  icon.className = className;
  icon.setAttribute("aria-hidden", "true");
  icon.innerHTML = markup;
  return icon;
}

function makeTextSpan(value, className) {
  const span = document.createElement("span");
  span.className = className;
  span.textContent = value;
  return span;
}

function selectRow(key) {
  if (!key) {
    return;
  }
  selectedRowKey = key;
  renderRegistry();
}

function setupResizableWorkspace() {
  const split = document.getElementById("workspaceSplit");
  const resizer = document.getElementById("cardPaneResizer");
  const cardPane = document.getElementById("cardPane");
  if (!split || !resizer || !cardPane) {
    return;
  }

  const storedWidth = readStoredNumber(RIGHT_PANE_STORAGE_KEY, RIGHT_PANE_DEFAULT_WIDTH);
  applyRightPaneWidth(storedWidth);

  let startX = 0;
  let startWidth = 0;

  const resizeFromPointer = (event) => {
    const nextWidth = startWidth + (startX - event.clientX);
    applyRightPaneWidth(nextWidth);
  };

  const finishResize = () => {
    document.body.classList.remove("resizing-pane");
    document.removeEventListener("pointermove", resizeFromPointer);
    writeStoredNumber(RIGHT_PANE_STORAGE_KEY, currentRightPaneWidth());
  };

  resizer.addEventListener("pointerdown", (event) => {
    if (window.matchMedia("(max-width: 1120px)").matches) {
      return;
    }
    event.preventDefault();
    startX = event.clientX;
    startWidth = currentRightPaneWidth();
    document.body.classList.add("resizing-pane");
    document.addEventListener("pointermove", resizeFromPointer);
    document.addEventListener("pointerup", finishResize, { once: true });
  });

  resizer.addEventListener("keydown", (event) => {
    const step = event.shiftKey ? 40 : 16;
    if (event.key === "ArrowLeft") {
      event.preventDefault();
      applyRightPaneWidth(currentRightPaneWidth() + step, true);
    }
    if (event.key === "ArrowRight") {
      event.preventDefault();
      applyRightPaneWidth(currentRightPaneWidth() - step, true);
    }
  });

  window.addEventListener("resize", () => applyRightPaneWidth(currentRightPaneWidth()));
}

function currentRightPaneWidth() {
  const cardPane = document.getElementById("cardPane");
  return cardPane ? cardPane.getBoundingClientRect().width : RIGHT_PANE_DEFAULT_WIDTH;
}

function rightPaneMaxWidth() {
  const split = document.getElementById("workspaceSplit");
  if (!split) {
    return RIGHT_PANE_MAX_WIDTH;
  }
  const availableWidth = split.getBoundingClientRect().width;
  const maxByRegistryWidth = availableWidth - REGISTRY_PANE_MIN_WIDTH;
  const maxByRatio = availableWidth * RIGHT_PANE_MAX_RATIO;
  return Math.max(
    RIGHT_PANE_MIN_WIDTH,
    Math.min(RIGHT_PANE_MAX_WIDTH, maxByRegistryWidth, maxByRatio),
  );
}

function refreshResponsiveTableColumns() {
  renderTableColGroup(Array.isArray(payload.table_columns) ? payload.table_columns : []);
}

function scheduleResponsiveTableColumnsRefresh() {
  window.requestAnimationFrame(refreshResponsiveTableColumns);
}

function applyRightPaneWidth(width, persist = false) {
  const split = document.getElementById("workspaceSplit");
  const resizer = document.getElementById("cardPaneResizer");
  const maxWidth = rightPaneMaxWidth();
  const nextWidth = clampNumber(width, RIGHT_PANE_MIN_WIDTH, maxWidth);
  if (split) {
    split.style.setProperty("--right-pane-width", `${Math.round(nextWidth)}px`);
  }
  if (resizer) {
    resizer.setAttribute("aria-valuemin", String(RIGHT_PANE_MIN_WIDTH));
    resizer.setAttribute("aria-valuemax", String(Math.round(maxWidth)));
    resizer.setAttribute("aria-valuenow", String(Math.round(nextWidth)));
  }
  scheduleResponsiveTableColumnsRefresh();
  if (persist) {
    writeStoredNumber(RIGHT_PANE_STORAGE_KEY, nextWidth);
  }
}

async function initializeApp() {
  setupResizableWorkspace();
  setupWorkflow();
  await loadStateFromLocalApp();
  selectedRowKey = initialSelectedRowKey();
  renderAll();
}

function renderAll() {
  renderHeader();
  renderNav();
  renderSidebarPanels();
  renderSectionNotice();
  renderSettingsPanel();
  renderKpis();
  renderTopActions();
  renderRegistryTabs();
  renderFilters();
  renderRegistry();
}

function renderSidebarPanels() {
  const projectPanel = document.getElementById("sidebarProjectPanel");
  const settingsPanel = document.getElementById("sidebarSettingsPanel");
  if (projectPanel) {
    projectPanel.hidden = activeSection !== "Проект";
  }
  if (settingsPanel) {
    settingsPanel.hidden = activeSection !== "Настройки";
  }
}

function renderSectionNotice() {
  const notice = document.getElementById("sectionNotice");
  if (!notice) {
    return;
  }
  const messages = {
    Проект: "",
    Настройки: "",
  };
  const message = messages[activeSection] || "";
  notice.hidden = !message;
  notice.textContent = message;
}

function renderSettingsPanel() {
  const panel = document.getElementById("settingsPanel");
  if (!panel) {
    return;
  }
  panel.hidden = activeSection !== "Настройки";
  const savedMappings = ((payload.column_mappings || {}).mappings || []).filter(Boolean);
  const count = document.getElementById("savedMappingsCount");
  if (count) {
    count.textContent = `Сохранено: ${savedMappings.length}`;
  }
  const list = document.getElementById("savedMappingsList");
  if (!list) {
    return;
  }
  list.replaceChildren();
  if (latestPreview) {
    const current = document.createElement("div");
    current.className = "mapping-settings-current";
    const title = document.createElement("strong");
    title.textContent = "Текущее сопоставление колонок";
    const meta = document.createElement("p");
    meta.textContent = `Файл: ${selectedFilesLabel() || latestPreview.source_file || "не выбран"}; лист: ${latestPreview.sheet_name || "без листа"}; строка заголовков: ${latestPreview.header_row || "не найдена"}`;
    current.append(title, meta);
    list.append(current);
  }
  if (!savedMappings.length) {
    const empty = document.createElement("p");
    empty.className = "placeholder";
    empty.textContent = "Сохраненных сопоставлений пока нет";
    list.append(empty);
    return;
  }
  savedMappings.forEach((mapping) => {
    const itemEl = document.createElement("div");
    itemEl.className = "saved-mapping-item";
    const title = document.createElement("strong");
    title.textContent = mapping.source_file_name || mapping.mapping_id || "Сопоставление";
    const meta = document.createElement("span");
    meta.textContent = `Лист: ${mapping.sheet_name || "без листа"}; тип: ${datasetTypeLabel(mapping.dataset_type)}; файл настроек: ${mapping.mapping_file || ""}`;
    itemEl.append(title, meta);
    list.append(itemEl);
  });
}

function selectedFilesLabel() {
  if (!selectedFiles.length) {
    return "";
  }
  return selectedFiles.map((file) => file.name).join("; ");
}

function selectedFilesTypeLabel() {
  if (!selectedFiles.length) {
    return "";
  }
  return [...new Set(selectedFiles.map((file) => file.name.split(".").pop().toUpperCase()))].join(", ");
}

function appendSelectedFiles(form) {
  selectedFiles.forEach((file) => {
    form.append("file", file);
  });
}

function setupWorkflow() {
  const createProjectButton = document.getElementById("createProjectButton");
  const openProjectButton = document.getElementById("openProjectButton");
  const projectPathInput = document.getElementById("projectPathInput");
  const fileInput = document.getElementById("fileInput");
  const previewButton = document.getElementById("previewImportButton");
  const importButton = document.getElementById("runImportButton");
  const checkMappingButton = document.getElementById("checkMappingButton");
  const saveMappingButton = document.getElementById("saveMappingButton");
  const applyMappingButton = document.getElementById("applyMappingButton");
  const resetMappingButton = document.getElementById("resetMappingButton");
  const closeMappingModalButton = document.getElementById("closeMappingModalButton");
  const cancelImportButton = document.getElementById("cancelImportButton");
  const cancelDownloadConfirmButton = document.getElementById("cancelDownloadConfirmButton");
  const cancelDownloadConfirmButtonFooter = document.getElementById("cancelDownloadConfirmButtonFooter");
  const downloadAllDocumentsButton = document.getElementById("downloadAllDocumentsButton");
  const downloadProtocolResultsButton = document.getElementById("downloadProtocolResultsButton");
  const cancelRefreshConfirmButton = document.getElementById("cancelRefreshConfirmButton");
  const cancelRefreshConfirmButtonFooter = document.getElementById("cancelRefreshConfirmButtonFooter");
  const confirmRefreshButton = document.getElementById("confirmRefreshButton");
  const chooseProjectFolder = async (message) => {
    setImportStatus(message);
    const selection = await postJson("/api/projects/select-folder", {});
    if (selection.status === "cancelled") {
      setImportStatus(selection.message || "Выбор папки проекта отменен");
      return "";
    }
    if (selection.status !== "ok" || !selection.project_root) {
      setImportStatus(selection.message || "Папка проекта не выбрана");
      return "";
    }
    return selection.project_root;
  };

  createProjectButton.addEventListener("click", async () => {
    const projectRoot = await chooseProjectFolder("Выберите папку, где создать структуру проекта");
    if (!projectRoot) {
      return;
    }
    setImportStatus("Создаю проект в выбранной папке");
    const data = await postJson("/api/projects/create", {
      project_name: "Новый проект TenderVestDocs",
      parent_root: projectRoot,
    });
    applyPayloadResponse(data);
    if (data.status !== "ok" && data.status !== "ok_with_warnings") {
      setImportStatus(data.message || "Проект не создан");
      return;
    }
    setImportStatus(data.message || (data.project_name ? `Проект создан: ${data.project_name}` : "Проект создан"));
  });

  openProjectButton.addEventListener("click", async () => {
    const projectRoot = await chooseProjectFolder("Выберите папку существующего проекта");
    if (!projectRoot) {
      return;
    }
    setImportStatus("Открываю выбранный проект");
    const data = await postJson("/api/projects/open", { project_root: projectRoot });
    applyPayloadResponse(data);
    if (data.status !== "ok") {
      setImportStatus(data.message || "Проект не открыт");
      return;
    }
    setImportStatus(data.message || (data.project_name ? `Проект открыт: ${data.project_name}` : "Проект открыт"));
  });

  fileInput.addEventListener("change", async () => {
    const files = fileInput.files ? Array.from(fileInput.files) : [];
    if (files.length > 2) {
      selectedFiles = [];
      selectedFile = null;
      fileInput.value = "";
      previewButton.disabled = true;
      importButton.disabled = true;
      cancelImportButton.disabled = true;
      latestPreview = null;
      mappingVisible = false;
      renderMappingGrid([]);
      updateMappingButtons();
      renderPreviewSummary(null);
      setImportStatus("Выберите не больше двух Excel-файлов: лист 4 и лист 5");
      return;
    }
    selectedFiles = files;
    selectedFile = selectedFiles.length ? selectedFiles[0] : null;
    previewButton.disabled = !selectedFile;
    importButton.disabled = true;
    cancelImportButton.disabled = !selectedFile;
    latestPreview = null;
    mappingVisible = false;
    renderMappingGrid([]);
    updateMappingButtons();
    renderPreviewSummary(null);
    setImportStatus(selectedFile ? `Выбраны файлы реестра: ${selectedFilesLabel()}` : "Файл реестра не выбран");
    if (selectedFile) {
      await previewSelectedFile();
      if (latestPreview && !latestPreview.manual_mapping_required) {
        await runImportWithCurrentState();
      }
    }
  });

  previewButton.addEventListener("click", previewSelectedFile);

  checkMappingButton.addEventListener("click", () => {
    const result = validateMappingClient();
    renderMappingGrid((payload && payload.field_mapping_fields) || FIELD_MAPPING_LABELS);
    if (result.errors.length) {
      setImportStatus(`Сопоставление требует исправления: ${result.errors.join("; ")}`);
      return;
    }
    setImportStatus(
      `Сопоставление проверено: сопоставлено ${result.mapped}, игнорируется ${result.ignored}, конфликтов ${result.conflicts}`,
    );
  });

  saveMappingButton.addEventListener("click", async () => {
    const validation = validateMappingClient();
    if (validation.errors.length) {
      setImportStatus(`Сопоставление не сохранено: ${validation.errors.join("; ")}`);
      renderMappingGrid((payload && payload.field_mapping_fields) || FIELD_MAPPING_LABELS);
      return;
    }
    const data = await postJson("/api/mapping/save", {
      field_mapping: collectFieldMapping(),
      dataset_type: datasetTypeForMapping(),
    });
    applyPayloadResponse(data);
    mappingVisible = true;
    renderMappingGrid((payload && payload.field_mapping_fields) || FIELD_MAPPING_LABELS);
    updateMappingButtons();
    setImportStatus(data.message || "Сопоставление сохранено");
  });

  resetMappingButton.addEventListener("click", async () => {
    document.querySelectorAll("#mappingGrid select").forEach((select) => {
      select.value = "";
    });
    const data = await postJson("/api/mapping/reset", {});
    applyPayloadResponse(data);
    mappingVisible = Boolean(latestPreview && latestPreview.manual_mapping_required);
    renderMappingGrid((payload && payload.field_mapping_fields) || FIELD_MAPPING_LABELS);
    updateMappingButtons();
    setImportStatus(data.message || "Сопоставление сброшено");
  });

  applyMappingButton.addEventListener("click", runImportWithCurrentState);

  closeMappingModalButton.addEventListener("click", () => {
    mappingVisible = false;
    renderMappingGrid([]);
    updateMappingButtons();
    setImportStatus("Сопоставление закрыто. Навигация и реестр доступны.");
  });

  cancelImportButton.addEventListener("click", () => {
    selectedFile = null;
    selectedFiles = [];
    latestPreview = null;
    mappingVisible = false;
    fileInput.value = "";
    previewButton.disabled = true;
    importButton.disabled = true;
    cancelImportButton.disabled = true;
    renderPreviewSummary(null);
    renderMappingGrid([]);
    updateMappingButtons();
    setImportStatus("Файл реестра не выбран");
  });

  importButton.addEventListener("click", async () => {
    if (!selectedFile) {
      return;
    }
    if (latestPreview && latestPreview.manual_mapping_required) {
      mappingVisible = true;
      renderMappingGrid((payload && payload.field_mapping_fields) || FIELD_MAPPING_LABELS);
      updateMappingButtons();
      setImportStatus("Требуется сопоставление колонок");
      return;
    }
    await runImportWithCurrentState();
  });

  if (cancelDownloadConfirmButton) {
    cancelDownloadConfirmButton.addEventListener("click", hideDownloadConfirmation);
  }
  if (cancelDownloadConfirmButtonFooter) {
    cancelDownloadConfirmButtonFooter.addEventListener("click", hideDownloadConfirmation);
  }
  if (downloadAllDocumentsButton) {
    downloadAllDocumentsButton.addEventListener("click", confirmPendingDownload);
  }
  if (downloadProtocolResultsButton) {
    downloadProtocolResultsButton.addEventListener("click", confirmPendingProtocolResultsPreview);
  }
  if (cancelRefreshConfirmButton) {
    cancelRefreshConfirmButton.addEventListener("click", hideRefreshConfirmation);
  }
  if (cancelRefreshConfirmButtonFooter) {
    cancelRefreshConfirmButtonFooter.addEventListener("click", hideRefreshConfirmation);
  }
  if (confirmRefreshButton) {
    confirmRefreshButton.addEventListener("click", confirmPendingRefresh);
  }
}

async function runImportWithCurrentState() {
  if (!selectedFiles.length) {
    return;
  }
  if (mappingVisible) {
    const validation = validateMappingClient();
    if (validation.errors.length) {
      setImportStatus(`Сопоставление требует исправления: ${validation.errors.join("; ")}`);
      renderMappingGrid((payload && payload.field_mapping_fields) || FIELD_MAPPING_LABELS);
      return;
    }
  }
  setImportStatus("Загружаю реестр в текущий проект");
  const form = new FormData();
  appendSelectedFiles(form);
  form.append("field_mapping", JSON.stringify(collectFieldMapping()));
  const data = await postForm("/api/import", form);
  applyPayloadResponse(data);
  if (data.status === "mapping_required") {
    latestPreview = (data.payload && data.payload.import_preview) || latestPreview;
    mappingVisible = true;
    renderMappingGrid((payload && payload.field_mapping_fields) || FIELD_MAPPING_LABELS);
    updateMappingButtons();
    setImportStatus(data.message || "Требуется сопоставление колонок");
    return;
  }
  if (data.status !== "ok" && data.status !== "ok_with_warnings") {
    setImportStatus(data.message || "Импорт не выполнен");
    return;
  }
  mappingVisible = false;
  renderMappingGrid([]);
  updateMappingButtons();
  const rowsCount = data.import_result ? data.import_result.rows_count : rows().length;
  setImportStatus(`Импорт завершен, строк: ${rowsCount}`);
}

async function previewSelectedFile() {
  if (!selectedFiles.length) {
    return;
  }
  setImportStatus("Проверяю файл");
  const form = new FormData();
  appendSelectedFiles(form);
  const data = await postForm("/api/import/preview", form);
  latestPreview = data.preview || null;
  if (data && data.payload) {
    payload = data.payload;
  }
  mappingVisible = Boolean(latestPreview && latestPreview.manual_mapping_required);
  renderPreviewSummary(latestPreview);
  renderMappingGrid(data.field_mapping_fields || FIELD_MAPPING_LABELS);
  document.getElementById("runImportButton").disabled = false;
  document.getElementById("cancelImportButton").disabled = false;
  updateMappingButtons();
  const profile = latestPreview ? latestPreview.file_profile : "файл";
  const rowsLabel = latestPreview ? latestPreview.data_rows_count : 0;
  const status = latestPreview ? latestPreview.recognition_status : "Файл выбран";
  setImportStatus(`${status}: ${profile}, строк: ${rowsLabel}`);
}

function renderMappingGrid(fields) {
  const root = document.getElementById("mappingGrid");
  const modal = document.getElementById("mappingModal");
  if (modal) {
    modal.hidden = !mappingVisible;
  }
  root.replaceChildren();
  if (!mappingVisible) {
    return;
  }
  const columns = latestPreview && Array.isArray(latestPreview.columns) ? latestPreview.columns : [];
  const targets = fields && fields.length ? fieldTargets() : FIELD_MAPPING_LABELS;
  renderMappingMeta(root);
  const header = document.createElement("div");
  header.className = "mapping-row mapping-header";
  ["№", "Колонка таблицы", "Пример значения", "Поле приложения", "Обязательность", "Статус", "Действие"].forEach((textValue) => {
    const cell = document.createElement("span");
    cell.textContent = textValue;
    header.append(cell);
  });
  root.append(header);

  columns.forEach((column, index) => {
    const row = document.createElement("div");
    row.className = "mapping-row";
    row.dataset.sourceColumn = column;

    const number = document.createElement("span");
    number.textContent = String(index + 1);

    const source = document.createElement("span");
    source.className = "mapping-source-column";
    source.textContent = `${excelColumnLetter(index)} — ${column}`;

    const sample = document.createElement("span");
    sample.textContent = sampleValuesForColumn(column) || "—";

    const fieldCell = document.createElement("span");
    const select = document.createElement("select");
    select.dataset.sourceColumn = column;
    const ignore = document.createElement("option");
    ignore.value = "ignore_column";
    ignore.textContent = "Не импортировать эту колонку";
    select.append(ignore);
    appendTargetOptions(select, targets);
    select.value = suggestedFieldForColumn(column) || "ignore_column";
    select.addEventListener("change", () => updateMappingRowState(row, select));
    fieldCell.append(select);

    const required = document.createElement("span");
    required.className = "mapping-required";

    const status = document.createElement("span");
    status.className = "mapping-status";

    const action = document.createElement("span");
    const clear = document.createElement("button");
    clear.type = "button";
    clear.textContent = "Очистить";
    clear.addEventListener("click", () => {
      select.value = "ignore_column";
      updateMappingRowState(row, select);
    });
    action.append(clear);

    row.append(number, source, sample, fieldCell, required, status, action);
    root.append(row);
    updateMappingRowState(row, select);
  });
}

function updateMappingButtons() {
  const enabled = Boolean(mappingVisible && latestPreview);
  ["checkMappingButton", "saveMappingButton", "applyMappingButton", "resetMappingButton"].forEach((id) => {
    const button = document.getElementById(id);
    if (button) {
      button.disabled = !enabled;
    }
  });
}

function renderMappingMeta(root) {
  if (!latestPreview) {
    return;
  }
  const meta = document.createElement("div");
  meta.className = "mapping-meta";
  const importMapping = payload.import_mapping || {};
  const values = [
    ["Файл", selectedFilesLabel() || importMapping.source_file_name || latestPreview.source_file || ""],
    ["Лист", latestPreview.sheet_name || "без листа"],
    ["Тип таблицы", datasetTypeLabel(latestPreview.source_dataset_type || importMapping.dataset_type_suggestion)],
    ["Строка заголовков", latestPreview.header_row || "не найдена"],
    ["Строк данных", latestPreview.data_rows_count ?? 0],
    ["Статус", latestPreview.recognition_status || "Требуется сопоставление"],
  ];
  values.forEach(([label, value]) => {
    const itemEl = document.createElement("div");
    itemEl.className = "preview-item";
    const labelEl = document.createElement("span");
    labelEl.textContent = label;
    const valueEl = document.createElement("strong");
    valueEl.textContent = text(value);
    itemEl.append(labelEl, valueEl);
    meta.append(itemEl);
  });
  root.append(meta);
}

function appendTargetOptions(select, targets) {
  let currentGroup = "";
  targets.forEach((target) => {
    const group = target.group || "Поля приложения";
    if (group !== currentGroup) {
      currentGroup = group;
      const disabled = document.createElement("option");
      disabled.disabled = true;
      disabled.textContent = `— ${group} —`;
      select.append(disabled);
    }
    const option = document.createElement("option");
    option.value = target.field;
    option.textContent = `${target.label || target.field} (${target.field})`;
    option.dataset.required = target.required ? "true" : "false";
    option.dataset.requirementLabel = target.requirement_label || "опционально";
    select.append(option);
  });
}

function updateMappingRowState(row, select) {
  const selected = select.selectedOptions[0];
  const required = row.querySelector(".mapping-required");
  const status = row.querySelector(".mapping-status");
  const requirement = selected ? selected.dataset.requirementLabel || "опционально" : "опционально";
  required.textContent = select.value === "ignore_column" ? "необязательно" : requirement;
  const duplicate = select.value !== "ignore_column" && duplicateMappedFields().has(select.value);
  if (duplicate) {
    status.textContent = "конфликт";
    status.dataset.kind = "conflict";
    return;
  }
  if (select.value === "ignore_column") {
    status.textContent = "не использовать";
    status.dataset.kind = "ignored";
    return;
  }
  status.textContent = suggestedFieldForColumn(select.dataset.sourceColumn) === select.value ? "сопоставлено" : "требуется проверка";
  status.dataset.kind = status.textContent === "сопоставлено" ? "mapped" : "review";
}

function duplicateMappedFields() {
  const counts = new Map();
  document.querySelectorAll("#mappingGrid select").forEach((select) => {
    if (!select.value || select.value === "ignore_column") {
      return;
    }
    counts.set(select.value, (counts.get(select.value) || 0) + 1);
  });
  return new Set([...counts].filter(([, count]) => count > 1).map(([field]) => field));
}

function excelColumnLetter(index) {
  let value = index + 1;
  let label = "";
  while (value > 0) {
    const mod = (value - 1) % 26;
    label = String.fromCharCode(65 + mod) + label;
    value = Math.floor((value - mod) / 26);
  }
  return label;
}

function renderPreviewSummary(preview) {
  const root = document.getElementById("previewSummary");
  if (!root) {
    return;
  }
  root.replaceChildren();
  if (!preview) {
    root.hidden = true;
    return;
  }
  root.hidden = false;
  const items = [
    ["Файл", selectedFilesLabel()],
    ["Тип", selectedFilesTypeLabel()],
    ["Лист", preview.sheet_name || "без листа"],
    ["Строка заголовков", preview.header_row || "не найдена"],
    ["Строк данных", preview.data_rows_count ?? 0],
    ["Колонок", Array.isArray(preview.columns) ? preview.columns.length : 0],
    ["Профиль", preview.file_profile || "Нестандартный файл"],
    ["Статус", preview.recognition_status || "Нужна настройка сопоставления"],
  ];
  items.forEach(([label, value]) => {
    const itemEl = document.createElement("div");
    itemEl.className = "preview-item";
    const labelEl = document.createElement("span");
    labelEl.textContent = label;
    const valueEl = document.createElement("strong");
    valueEl.textContent = text(value);
    itemEl.append(labelEl, valueEl);
    root.append(itemEl);
  });
  if (preview.file_profile === "Стандартный лист 4" && !preview.manual_mapping_required) {
    const ok = document.createElement("p");
    ok.className = "preview-ok";
    ok.textContent = `Файл распознан как: Лист 4 — реестр закупок товаров. Найдено строк: ${preview.data_rows_count}. Строка заголовков: ${preview.header_row}. Ключевые поля найдены.`;
    root.append(ok);
  }
}

function collectFieldMapping() {
  if (!mappingVisible) {
    return {};
  }
  const mapping = {};
  document.querySelectorAll("#mappingGrid select").forEach((select) => {
    if (select.value && select.value !== "ignore_column") {
      mapping[select.value] = select.dataset.sourceColumn;
    }
  });
  return mapping;
}

function validateMappingClient() {
  const selects = [...document.querySelectorAll("#mappingGrid select")];
  const mapping = collectFieldMapping();
  const fields = new Set(Object.keys(mapping));
  const duplicateFields = duplicateMappedFields();
  const ignored = selects.filter((select) => select.value === "ignore_column").length;
  const errors = [];
  if (datasetTypeForMapping() === "contract") {
    const hasMainId = fields.has("contract_registry_number");
    const hasProcurement = fields.has("procurement_number") || fields.has("purchase_number");
    const hasSource = fields.has("contract_url") || fields.has("source_url") || fields.has("source_raw");
    if (!hasMainId && !(hasProcurement && hasSource)) {
      errors.push("не сопоставлен номер контракта или связка номер закупки + источник");
    }
  } else {
    if (!fields.has("procurement_number") && !fields.has("purchase_number")) {
      errors.push("не сопоставлен номер закупки");
    }
    if (!fields.has("law") && !fields.has("law_type")) {
      errors.push("не сопоставлен закон");
    }
    if (!fields.has("source_url") && !fields.has("source_raw")) {
      errors.push("не сопоставлен источник строки");
    }
    if (!fields.has("customer_name")) {
      errors.push("не сопоставлен заказчик");
    }
    if (!fields.has("subject") && !fields.has("name")) {
      errors.push("не сопоставлен предмет закупки");
    }
  }
  if (duplicateFields.size) {
    errors.push(`конфликт полей: ${[...duplicateFields].join(", ")}`);
  }
  return {
    mapped: Object.keys(mapping).length,
    ignored,
    conflicts: duplicateFields.size,
    errors,
  };
}

function datasetTypeForMapping() {
  const fields = new Set(Object.keys(collectFieldMapping()));
  const previewDataset = latestPreview && latestPreview.source_dataset_type;
  const contractMarkers = [
    "contract_registry_number",
    "contract_url",
    "supplier_name",
    "supplier_inn",
    "contract_price",
    "contract_date",
    "execution_status",
  ];
  if (contractMarkers.some((field) => fields.has(field)) || previewDataset === "contract") {
    return "contract";
  }
  return "procurement";
}

async function loadStateFromLocalApp() {
  try {
    const data = await getJson("/api/state");
    applyPayloadResponse(data);
  } catch (error) {
    payload = EMPTY_PAYLOAD;
    setImportStatus("Локальный API не подключен, показан пустой стартовый экран");
  }
}

function applyPayloadResponse(data) {
  if (data && data.payload) {
    const previousSelectedRowKey = selectedRowKey;
    payload = data.payload;
    if (payload.active_section) {
      activeSection = payload.active_section;
    }
    selectedRowKey = rowExists(previousSelectedRowKey) ? previousSelectedRowKey : initialSelectedRowKey();
    renderAll();
    renderStoredPreviewState();
  }
}

function renderStoredPreviewState() {
  if (!payload.import_preview) {
    renderPreviewSummary(null);
    mappingVisible = false;
    renderMappingGrid([]);
    updateMappingButtons();
    return;
  }
  latestPreview = payload.import_preview;
  mappingVisible = Boolean(latestPreview.manual_mapping_required);
  renderPreviewSummary(latestPreview);
  renderMappingGrid(payload.field_mapping_fields || FIELD_MAPPING_LABELS);
  const rowsLabel = latestPreview ? latestPreview.data_rows_count : 0;
  setImportStatus(`Файл распознан, строк: ${rowsLabel}`);
  ["runImportButton", "cancelImportButton"].forEach((id) => {
    const button = document.getElementById(id);
    if (button) {
      button.disabled = !latestPreview;
    }
  });
  updateMappingButtons();
}

async function handleTopAction(action) {
  if (!action || action.enabled === false) {
    return;
  }
  if (action.action === "refresh_sources_all") {
    const rowIds = rows().map((row) => row.row_id).filter(Boolean);
    showRefreshConfirmation(rowIds, "ui_all_rows");
    return;
  }
  if (action.action === "refresh_sources_filtered") {
    const rowIds = getFilteredRows().map((row) => row.row_id).filter(Boolean);
    showRefreshConfirmation(rowIds, "ui_current_filter");
    return;
  }
  if (action.action === "normative_check_all") {
    const rowIds = rows().map((row) => row.row_id).filter(Boolean);
    await runNormativeCheck(rowIds, "ui_all_rows");
    return;
  }
  if (action.action === "sheet6_export") {
    const rowIds = rows().map((row) => row.row_id).filter(Boolean);
    await runSheet6Export(rowIds, "ui_all_rows");
    return;
  }
  if (action.action === "download_current_filter") {
    const rowIds = getFilteredRows().map((row) => row.row_id).filter(Boolean);
    showDownloadConfirmation(rowIds, "ui_current_filter");
  } else if (action.action === "download_selected_rows") {
    const row = currentRow();
    await runDownload(row && row.row_id ? [row.row_id] : [], "ui_selected_rows");
  }
}

async function refreshSources(rowIds, scopeLabel, options = {}) {
  if (!rowIds.length) {
    setImportStatus("Нет строк для обновления сведений");
    return;
  }
  if (!options.suppressStartStatus) {
    setImportStatus(
      scopeLabel === "ui_current_filter"
        ? `Обновляю данные по текущему фильтру: ${rowIds.length} строк. Документы не скачиваются.`
        : `Обновляю данные по всей таблице: ${rowIds.length} строк. Документы не скачиваются.`,
    );
  }
  const data = await postJson("/api/refresh-sources", { row_ids: rowIds, scope_label: scopeLabel });
  applyPayloadResponse(data);
  renderRegistry();
  setImportStatus(data.message || "Сведения обновлены");
}

function refreshScopeText(scopeLabel) {
  if (scopeLabel === "ui_current_filter") {
    return "Текущий фильтр";
  }
  if (scopeLabel === "ui_single_row") {
    return "Одна строка";
  }
  return "Весь реестр";
}

function showRefreshConfirmation(rowIds, scopeLabel) {
  if (!rowIds.length) {
    setImportStatus("Нет строк для обновления сведений");
    return;
  }
  const modal = document.getElementById("refreshConfirmModal");
  if (!modal) {
    refreshSources(rowIds, scopeLabel);
    return;
  }
  pendingRefresh = { rowIds, scopeLabel };
  setRefreshConfirmBusy(false, "");
  const message = document.getElementById("refreshConfirmMessage");
  const scope = document.getElementById("refreshScopeLabel");
  const rowsCount = document.getElementById("refreshRowsCount");
  if (message) {
    message.textContent =
      scopeLabel === "ui_current_filter"
        ? `Будет обновлено строк по текущему фильтру: ${rowIds.length}. Документы скачиваться не будут.`
        : `Будет обновлен весь реестр: ${rowIds.length} строк. Документы скачиваться не будут.`;
  }
  if (scope) {
    scope.textContent = refreshScopeText(scopeLabel);
  }
  if (rowsCount) {
    rowsCount.textContent = String(rowIds.length);
  }
  modal.hidden = false;
}

function hideRefreshConfirmation() {
  pendingRefresh = null;
  setRefreshConfirmBusy(false, "");
  const modal = document.getElementById("refreshConfirmModal");
  if (modal) {
    modal.hidden = true;
  }
}

function setRefreshConfirmBusy(isBusy, text) {
  const progress = document.getElementById("refreshProgressStatus");
  const confirmButton = document.getElementById("confirmRefreshButton");
  const cancelButton = document.getElementById("cancelRefreshConfirmButton");
  const cancelFooterButton = document.getElementById("cancelRefreshConfirmButtonFooter");
  if (progress) {
    progress.hidden = !isBusy;
    progress.textContent = text || "";
  }
  [confirmButton, cancelButton, cancelFooterButton].forEach((button) => {
    if (button) {
      button.disabled = isBusy;
    }
  });
}

async function confirmPendingRefresh() {
  if (!pendingRefresh) {
    hideRefreshConfirmation();
    return;
  }
  const { rowIds, scopeLabel } = pendingRefresh;
  setRefreshConfirmBusy(
    true,
    `${refreshScopeText(scopeLabel)}: обновление выполняется, строк к обработке ${rowIds.length}. Документы не скачиваются.`,
  );
  setImportStatus(
    scopeLabel === "ui_current_filter"
      ? `Обновляю данные по текущему фильтру: ${rowIds.length} строк. Документы не скачиваются.`
      : `Обновляю данные по всей таблице: ${rowIds.length} строк. Документы не скачиваются.`,
  );
  try {
    await refreshSources(rowIds, scopeLabel, { suppressStartStatus: true });
    setRefreshConfirmBusy(
      true,
      `${refreshScopeText(scopeLabel)}: обновление завершено, обработано строк ${rowIds.length}. Документы не скачивались.`,
    );
    window.setTimeout(hideRefreshConfirmation, 900);
  } catch (error) {
    setRefreshConfirmBusy(false, "");
    throw error;
  }
}

async function handleRowAction(action, row) {
  if (!row || !action || action.enabled === false) {
    return;
  }
  if (action.action === "refresh_row") {
    await refreshSources(row.row_id ? [row.row_id] : [], "ui_single_row");
    return;
  }
  if (action.action === "open_link") {
    let data = {};
    let url = normalizeHttpUrl(action.url) || sourceUrlForRow(row);
    if (!url) {
      data = await postJson("/api/open-link", { row_id: row.row_id });
      url = data.url || "";
    }
    if (url) {
      window.open(url, "_blank", "noopener");
      setImportStatus("Публичная ссылка открыта в новой вкладке");
    } else {
      setImportStatus(data.message || "В строке нет ссылки");
    }
  }
  if (action.action === "download_row") {
    await runDownload([row.row_id], "ui_row_action");
  }
  if (action.action === "normative_check_row") {
    await runNormativeCheck(row.row_id ? [row.row_id] : [], "ui_single_row");
  }
}

function showDownloadConfirmation(rowIds, scopeLabel) {
  const modal = document.getElementById("downloadConfirmModal");
  if (!modal) {
    runProtocolResultsPreview(rowIds, `${scopeLabel}_protocol_results`);
    return;
  }
  pendingDownload = { rowIds, scopeLabel };
  const rowsCount = document.getElementById("downloadRowsCount");
  const docsCount = document.getElementById("downloadExpectedDocsCount");
  const projectName = document.getElementById("downloadProjectName");
  const targetFolder = document.getElementById("downloadTargetFolder");
  const allDocumentsButton = document.getElementById("downloadAllDocumentsButton");
  const protocolResultsButton = document.getElementById("downloadProtocolResultsButton");
  const warning = document.getElementById("downloadModeWarning");
  const allState = downloadAllDocumentsState(rowIds);
  if (rowsCount) {
    rowsCount.textContent = String(rowIds.length);
  }
  if (docsCount) {
    docsCount.textContent = String(expectedDocumentsForRows(rowIds));
  }
  if (projectName) {
    projectName.textContent = (payload.project_summary && payload.project_summary.project_name) || "Проект не выбран";
  }
  if (targetFolder) {
    targetFolder.textContent = "03_Документы_закупок";
  }
  if (allDocumentsButton) {
    allDocumentsButton.disabled = !allState.enabled;
    allDocumentsButton.title = allState.reason;
  }
  if (protocolResultsButton) {
    protocolResultsButton.disabled = !rowIds.length;
    protocolResultsButton.title = rowIds.length
      ? "Временный разбор результатов без создания проекта"
      : "Нет строк для получения результатов";
  }
  if (warning) {
    warning.textContent = allState.enabled
      ? "Полный пакет документов будет сохранен в проект. Только результаты будут разобраны во временной папке."
      : `${allState.reason} Доступен режим "Только результаты / протоколы" без создания проекта.`;
  }
  modal.hidden = false;
}

function hideDownloadConfirmation() {
  pendingDownload = null;
  const modal = document.getElementById("downloadConfirmModal");
  if (modal) {
    modal.hidden = true;
  }
}

async function confirmPendingDownload() {
  if (!pendingDownload) {
    hideDownloadConfirmation();
    return;
  }
  const { rowIds, scopeLabel } = pendingDownload;
  hideDownloadConfirmation();
  await runDownload(rowIds, scopeLabel);
}

async function confirmPendingProtocolResultsPreview() {
  if (!pendingDownload) {
    hideDownloadConfirmation();
    return;
  }
  const { rowIds, scopeLabel } = pendingDownload;
  hideDownloadConfirmation();
  await runProtocolResultsPreview(rowIds, `${scopeLabel}_protocol_results`);
}

function downloadAllDocumentsState(rowIds) {
  if (!rowIds.length) {
    return { enabled: false, reason: "Нет строк для скачивания." };
  }
  const projectRoot = (payload.project_summary && payload.project_summary.project_root) || "";
  if (!projectRoot) {
    return { enabled: false, reason: "Для полного скачивания документов создайте или откройте проект." };
  }
  if (!payload.sources_enriched) {
    return { enabled: false, reason: "Для полного скачивания сначала обновите сведения из источников." };
  }
  const ids = new Set(rowIds);
  const hasDownloadable = rows().some((row) => {
    const documents = row.documents || {};
    return ids.has(row.row_id) && documents.download_available === true;
  });
  if (!hasDownloadable) {
    return { enabled: false, reason: "В выбранных строках нет доступного контролируемого скачивания документов." };
  }
  return { enabled: true, reason: "" };
}

function expectedDocumentsForRows(rowIds) {
  const ids = new Set(rowIds);
  let hasUnknownExpected = false;
  const total = rows()
    .filter((row) => ids.has(row.row_id))
    .reduce((sum, row) => {
      const documents = row.documents || {};
      if (documents.expected_known === false) {
        hasUnknownExpected = true;
        return sum;
      }
      const count = Number(documents.expected_count ?? documents.total ?? documents.count ?? 0);
      if (Number.isFinite(count) && count > 0) {
        return sum + count;
      }
      const countLabel = String(documents.count_label || row.documents_count_label || "");
      if (/будет определено|не проверено/i.test(countLabel)) {
        hasUnknownExpected = true;
        return sum;
      }
      const match = countLabel.match(/\/\s*(\d+)/);
      return sum + (match ? Number(match[1]) : 0);
    }, 0);
  if (hasUnknownExpected) {
    return total > 0 ? `${total} + ${UNKNOWN_EXPECTED_DOCUMENTS_TEXT}` : UNKNOWN_EXPECTED_DOCUMENTS_TEXT;
  }
  return total;
}

async function runDownload(rowIds, scopeLabel) {
  if (!rowIds.length) {
    setImportStatus("Нет строк для скачивания");
    return;
  }
  setImportStatus("Запускаю контролируемое скачивание");
  const data = await postJson("/api/download", { row_ids: rowIds, scope_label: scopeLabel });
  applyPayloadResponse(data);
  if (data.downloaded_files_count !== undefined) {
    const skipped = data.skipped_by_limit ? `, пропущено по лимиту: ${data.skipped_by_limit}` : "";
    setImportStatus(
      `Скачивание завершено, выбрано строк: ${data.selected_rows_count || rowIds.length}, файлов: ${data.downloaded_files_count}${skipped}`,
    );
  } else {
    setImportStatus(data.message || "Скачивание проверено");
  }
}

async function runProtocolResultsPreview(rowIds, scopeLabel) {
  if (!rowIds.length) {
    setImportStatus("Нет строк для получения результатов протокола");
    return;
  }
  setImportStatus("Получаю результаты протокола во временном режиме");
  const data = await postJson("/api/protocol-results/preview", { row_ids: rowIds, scope_label: scopeLabel });
  applyPayloadResponse(data);
  const deleted = data.temp_files_deleted_count ?? 0;
  const parsed = data.protocols_parsed_count ?? 0;
  if (data.status === "ok") {
    setImportStatus(
      `Результаты протокола получены: записей ${parsed}, временных файлов удалено ${deleted}. Проект не требуется.`,
    );
  } else {
    setImportStatus(data.message || "Результаты протокола не получены");
  }
}

async function runNormativeCheck(rowIds, scopeLabel) {
  if (!rowIds.length) {
    setImportStatus("Нет строк для нормативной проверки");
    return;
  }
  setImportStatus(`Выполняется нормативная проверка: ${rowIds.length} строк. Документы не скачиваются.`);
  const data = await postJson("/api/normative-check", { row_ids: rowIds, scope_label: scopeLabel });
  applyPayloadResponse(data);
  const checked = data.selected_rows_count ?? rowIds.length;
  const total = data.total_rows_count ?? rows().length;
  if (data.status === "ok") {
    setImportStatus(`Нормативная проверка завершена: ${checked} из ${total} строк. Экспорт и AI-ready файлы сохранены в проект.`);
  } else {
    setImportStatus(data.message || "Нормативная проверка не выполнена");
  }
}

async function runSheet6Export(rowIds, scopeLabel) {
  if (!rowIds.length) {
    setImportStatus("Нет строк для экспорта данных");
    return;
  }
  setImportStatus(`Формируется экспорт данных: ${rowIds.length} строк входного реестра. Документы и GPT-анализ не запускаются.`);
  const data = await postJson("/api/sheet6-export", { row_ids: rowIds, scope_label: scopeLabel });
  applyPayloadResponse(data);
  const rowsCount = data.control_rows_count ?? data.summary?.control_rows_count ?? 0;
  const warnings = data.warnings_count ?? data.summary?.warnings_count ?? 0;
  const missing = data.missing_inputs_count ?? data.summary?.missing_inputs_count ?? 0;
  const path = data.xlsx_path || data.paths?.xlsx || "";
  if (data.status === "ok" || data.status === "ok_with_warnings" || data.status === "export_data_ready" || data.status === "export_data_ready_with_warnings") {
    const userStatus = data.status === "export_data_ready_with_warnings" || warnings ? "успешно с предупреждениями" : "успешно";
    setImportStatus(`Экспорт данных создан, статус: ${userStatus}. Контрольных узлов: ${rowsCount}, warnings: ${warnings}, missing inputs: ${missing}. Файл: ${path}`);
  } else {
    setImportStatus(data.message || "Экспорт данных не создан");
  }
}

async function getJson(url) {
  const response = await fetch(url, { method: "GET" });
  if (!response.ok) {
    throw new Error(`GET ${url}: ${response.status}`);
  }
  return response.json();
}

async function postJson(url, body) {
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body || {}),
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data.message || `POST ${url}: ${response.status}`);
  }
  return data;
}

async function postForm(url, form) {
  const response = await fetch(url, { method: "POST", body: form });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data.message || `POST ${url}: ${response.status}`);
  }
  return data;
}

function setImportStatus(message) {
  const target = document.getElementById("importStatus");
  if (target) {
    target.textContent = message;
  }
}

selectedRowKey = initialSelectedRowKey();
initializeApp();
