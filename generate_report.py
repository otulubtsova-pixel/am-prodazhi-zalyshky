"""
Формирует сводный файл по товарам и магазинам на основании
"АМ+продажі+залишки_Міленіум_контрагент.xlsx".

Оба входных файла (файл контрагента и графік поставок) принимаются как
в формате .xlsx, так и в старом .xls - формат определяется по содержимому
файла (сигнатуре), а не по расширению в имени.

Один товар + один магазин = одна строка.

Логика:
- Лист "АМ": для каждого магазина берём "количество АМ" -> колонка АМ,
  "Остатки на дату, шт." -> колонка Залишок.
- Лист "Продажі 2025": для каждого магазина суммируем "Сумма по столбцу
  Кол-во Шт" по месяцам, распределяя их в два периода:
    "11 Лис. - 01 Січ."  -> месяцы 11, 12, 01
    "02 Лют. - 10 Жов."  -> месяцы 02..10
  В файле реально есть только Січ-Серп 2025 (рос. названия), поэтому
  первый период фактически = сумма за январь, второй = сумма Лют-Серп.
  Если для магазина в этот период вообще нет данных (ни одного значения) -
  ячейка остаётся пустой.
- "НОВА АМ" пока оставляем пустым.
- Товар и магазин объединяются по ключу "Но_"; описательные поля
  (Поставщик, Артикул и т.д.) берутся с листа АМ, а если товара там нет -
  с листа Продажі.
- Строятся ВСЕ комбинации товар x магазин (объединение списков магазинов
  и товаров из обоих листов) - так видно и то, где сейчас нет ни АМ,
  ни продаж (кандидаты на НОВА АМ).
- "Графік поставок ...xls": даёт поле товарного (не магазинного) уровня
  "В дорозі" - джойн по "Артикул" (а не "Но_"), используются только листы,
  где есть и артикул, и статус, и количество: Chicco, Kids2, Offspring,
  Kendamil. Остальные листы (без количества и/или статуса, либо вовсе без
  артикула) пропускаются. "В дорозі" = сумма количества "в пути" по всем
  совпавшим листам. Пусто и 0 - РАЗНЫЕ случаи и оба показываются как есть:
  пусто = артикула вообще нет в графике поставок; 0 = артикул найден, но
  сейчас по нему ничего не едет.
- "Новинка" НЕ определяется джойном по артикулу с уже существующим
  ассортиментом контрагента: товар, который уже есть в списке контрагента
  (на листах "АМ"/"Продажі"), по определению не может быть новинкой -
  у него уже есть статус в системе. Ищется ТОЛЬКО на листе "Chicco" -
  только там подтверждено, что "пустой Статус артикула" реально значит
  новинку (см. novelty_reliable в DELIVERY_SHEETS); на Kids2/Offspring/
  Kendamil пустой статус может быть по другим причинам, это не
  проверялось, поэтому они не используются для новинок (для "В дорозі"
  используются все 4). Строка считается новинкой, если на листе Chicco
  "Статус артикула" пуст, количество "в дорозі" больше нуля (иначе товар
  пока не интересен) и бренд строки (CHICCO) входит в бренды контрагента -
  если у контрагента вообще нет Chicco в ассортименте, список новинок
  будет пустым, и это ожидаемо (например, для БебіШоп). Найденные строки
  добавляются в отчёт ОТДЕЛЬНЫМИ строками в самом низу: заполняются
  "Наименование товара", "Артикул" (только если он реально есть в графике -
  если его там нет, поле остаётся пустым, ничем не подменяется) и
  "В дорозі". Остальные поля пустые.
"""

import io

import openpyxl
import xlrd
from openpyxl.styles import PatternFill

NOVA_AM_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

SRC = "АМ+продажі+залишки_Міленіум_контрагент.xlsx"
DELIVERY_SRC = "Графік поставок 04,09,2026.xls"
OUT = "Звід_АМ_продажі_залишки_Міленіум.xlsx"

XLS_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"  # OLE2/Compound File - старий .xls
XLSX_MAGIC = b"PK"  # ZIP - сучасний .xlsx


def detect_excel_format(data):
    """Визначає реальний формат за вмістом файлу, а не за розширенням."""
    if data[:2] == XLSX_MAGIC:
        return "xlsx"
    if data[:8] == XLS_MAGIC:
        return "xls"
    raise ValueError("Не вдалося розпізнати формат файлу - очікується .xlsx або .xls")


def read_bytes(file):
    """file - шлях на диску (str), файлоподібний об'єкт, або вже готові bytes."""
    if isinstance(file, (bytes, bytearray)):
        return bytes(file)
    if hasattr(file, "read"):
        data = file.read()
        if hasattr(file, "seek"):
            try:
                file.seek(0)
            except (OSError, ValueError):
                pass
        return data
    with open(file, "rb") as f:
        return f.read()


class _Cell:
    __slots__ = ("value",)

    def __init__(self, value):
        self.value = value


class _XlsAsOpenpyxlSheet:
    """Оборачивает лист xlrd (.xls, 0-индексация) под API листа openpyxl (1-индексация)."""

    def __init__(self, xlrd_sheet):
        self._sheet = xlrd_sheet
        self.max_row = xlrd_sheet.nrows
        self.max_column = xlrd_sheet.ncols

    def cell(self, row, column):
        r, c = row - 1, column - 1
        if r < 0 or c < 0 or r >= self._sheet.nrows or c >= self._sheet.ncols:
            return _Cell(None)
        value = self._sheet.cell_value(r, c)
        return _Cell(value if value != "" else None)


class _XlsAsOpenpyxlWorkbook:
    def __init__(self, xlrd_book):
        self._book = xlrd_book

    def __getitem__(self, name):
        return _XlsAsOpenpyxlSheet(self._book.sheet_by_name(name))


class _XlsxAsXlrdSheet:
    """Оборачивает лист openpyxl (.xlsx, 1-індексація) під API листа xlrd (0-індексація)."""

    def __init__(self, ws):
        self._ws = ws
        self.nrows = ws.max_row
        self.ncols = ws.max_column

    def cell_value(self, row, col):
        value = self._ws.cell(row=row + 1, column=col + 1).value
        return value if value is not None else ""


class _XlsxAsXlrdBook:
    def __init__(self, wb):
        self._wb = wb

    def sheet_names(self):
        return self._wb.sheetnames

    def sheet_by_name(self, name):
        return _XlsxAsXlrdSheet(self._wb[name])

DESC_COLS = [
    "Поставщик", "Код Группы", "Торговая Марка", "Но_",
    "Наименование товара", "Артикул", "Штрихкод", "Статус товара",
]

PRODUCT_FIELDS = ["В дорозі"]

# На каждом листе: колонка названия, артикула, статуса, и способ получить
# количество "в пути" - либо номер одной готовой колонки-итога (int),
# либо список колонок для суммирования, либо "dynamic" - колонки
# определяются по названию в шапке (для "плавающих" колонок, где их
# число меняется со временем, как "Замовлення N" у Offspring).
# "brand" - лист целиком про один бренд (используется для фильтрации
# новинок по бренду контрагента). "brand_col" - лист про несколько
# брендов, бренд читается из этой колонки построчно (Kids2).
# "novelty_reliable" - на этом листе "пустой Статус артикула" реально
# значит новинку (подтверждено только для Chicco; на остальных листах
# статус может быть пуст по совсем другим причинам, не значит "новый").
# Используется для "В дорозі" (все 4 листа), но для поиска новинок -
# только листы с novelty_reliable=True.
DELIVERY_SHEETS = {
    "Chicco": {"header_row": 2, "data_start": 4, "name_col": 0, "art_col": 1, "status_col": 3, "qty": 5,
               "brand": "CHICCO", "novelty_reliable": True},
    "Kids2": {"header_row": 2, "data_start": 4, "name_col": 0, "art_col": 1, "status_col": 3, "qty": [7, 8, 9],
              "brand_col": 5},
    "Kendamil": {"header_row": 1, "data_start": 2, "name_col": 1, "art_col": 0, "status_col": 2, "qty": 6,
                 "brand": "KENDAMIL"},
    "Offspring": {"header_row": 1, "data_start": 2, "name_col": 1, "art_col": 0, "status_col": 2, "qty": "dynamic",
                  "brand": "OFFSPRING"},
}


def normalize_brand(value):
    if value in ("", None):
        return None
    return "".join(ch for ch in str(value).upper() if ch.isalnum())


def _qty_cols(ws, cfg):
    if cfg["qty"] == "dynamic":
        header0 = [ws.cell_value(cfg["header_row"] - 1, c) for c in range(ws.ncols)]
        return [c for c, v in enumerate(header0) if str(v).strip().lower().startswith("замовл")]
    if isinstance(cfg["qty"], list):
        return cfg["qty"]
    return [cfg["qty"]]


def _row_qty(ws, r, cols):
    qty = 0
    for c in cols:
        v = ws.cell_value(r, c)
        if isinstance(v, (int, float)):
            qty += v
    return qty

GROUP1_LABEL = "11 Лис. - 01 Січ."
GROUP2_LABEL = "02 Лют. - 10 Жов."
GROUP1_MONTHS = {11, 12, 1}
GROUP2_MONTHS = {2, 3, 4, 5, 6, 7, 8, 9, 10}

EXCLUDED_STORES = {"(пусто)", None}


def norm(value):
    if value is None:
        return None
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def month_num(period_label):
    return int(str(period_label).strip()[:2])


def is_excluded_store(name):
    return name in EXCLUDED_STORES or (name and str(name).startswith("Итог"))


def parse_am_sheet(ws):
    store_cols = {}
    c = 9
    while c <= ws.max_column:
        store = ws.cell(row=2, column=c).value
        if not is_excluded_store(store):
            m1 = ws.cell(row=3, column=c).value
            m2 = ws.cell(row=3, column=c + 1).value if c + 1 <= ws.max_column else None
            if m1 == "количество АМ" and m2 == "Остатки на дату, шт.":
                store_cols[store] = (c, c + 1)
                c += 2
                continue
        c += 1

    desc = {}
    data = {}
    for r in range(4, ws.max_row + 1):
        key = norm(ws.cell(row=r, column=4).value)
        if key is None:
            continue
        desc[key] = {name: ws.cell(row=r, column=i + 1).value for i, name in enumerate(DESC_COLS)}
        for store, (ca, cz) in store_cols.items():
            av = ws.cell(row=r, column=ca).value
            zv = ws.cell(row=r, column=cz).value
            if av is not None or zv is not None:
                data[(key, store)] = (av, zv)

    return store_cols, desc, data


def parse_sales_sheet(ws):
    store_periods = {}
    for c in range(9, ws.max_column + 1):
        store = ws.cell(row=2, column=c).value
        period = ws.cell(row=3, column=c).value
        measure = ws.cell(row=4, column=c).value
        if is_excluded_store(store):
            continue
        if measure == "Сумма по столбцу Кол-во Шт":
            store_periods.setdefault(store, {})[period] = c

    desc = {}
    data = {}
    for r in range(5, ws.max_row + 1):
        key = norm(ws.cell(row=r, column=4).value)
        if key is None:
            continue
        desc[key] = {name: ws.cell(row=r, column=i + 1).value for i, name in enumerate(DESC_COLS)}
        for store, periods in store_periods.items():
            vals = {}
            for period, col in periods.items():
                v = ws.cell(row=r, column=col).value
                if v is not None:
                    vals[period] = v
            if vals:
                data[(key, store)] = vals

    return store_periods, desc, data


def group_sum(period_values, months):
    matched = [v for period, v in period_values.items() if month_num(period) in months]
    if not matched:
        return None
    return sum(matched)


def open_delivery_workbook(file):
    """file - путь на диске, файлоподобный объект либо байты (.xls или .xlsx)."""
    data = read_bytes(file)
    fmt = detect_excel_format(data)
    if fmt == "xls":
        return xlrd.open_workbook(file_contents=data)
    return _XlsxAsXlrdBook(openpyxl.load_workbook(io.BytesIO(data), data_only=True))


def parse_delivery_schedule(wb):
    """Возвращает {Артикул: кол-во в дорозі}, суммируя совпадения по всем листам."""
    result = {}

    for sheet_name, cfg in DELIVERY_SHEETS.items():
        if sheet_name not in wb.sheet_names():
            continue
        ws = wb.sheet_by_name(sheet_name)
        cols = _qty_cols(ws, cfg)

        for r in range(cfg["data_start"], ws.nrows):
            art = ws.cell_value(r, cfg["art_col"])
            if art in ("", None):
                continue
            art = norm(art)
            qty = _row_qty(ws, r, cols)
            result[art] = result.get(art, 0) + qty

    return result


def find_novelty_candidates(wb, existing_articles, contragent_brands):
    """
    Кандидаты в новинки: строки листов графика поставок с
    novelty_reliable=True (сейчас только Chicco - только там подтверждено,
    что "пустой Статус артикула" реально значит новинку; на остальных
    листах пустой статус может быть по другим причинам, не проверялось)
    без "Статус артикула" и с ненулевым количеством "в дорозі" (иначе товар
    не интересен - его пока даже не заказали). Строка учитывается, только
    если её бренд входит в contragent_brands (бренды, которые реально есть
    в ассортименте этого контрагента) - если у контрагента нет Chicco,
    список новинок будет пустым, и это ожидаемо. Товар, уже присутствующий
    в ассортименте контрагента (existing_articles), новинкой считаться не
    может - исключаем. Если у товара нет "Артикул" - оставляем пустым,
    ничем не подменяем (значит его ещё не завели в базу).
    Возвращает список (Наименование товара, Артикул или None, кол-во).
    """
    candidates = []
    seen = set()

    for sheet_name, cfg in DELIVERY_SHEETS.items():
        if not cfg.get("novelty_reliable"):
            continue
        if sheet_name not in wb.sheet_names():
            continue
        ws = wb.sheet_by_name(sheet_name)
        cols = _qty_cols(ws, cfg)
        fixed_brand = normalize_brand(cfg.get("brand"))

        for r in range(cfg["data_start"], ws.nrows):
            art_raw = ws.cell_value(r, cfg["art_col"])
            status = ws.cell_value(r, cfg["status_col"])
            name = ws.cell_value(r, cfg["name_col"])
            qty = _row_qty(ws, r, cols)

            if status not in ("", None):
                continue
            if qty <= 0:
                continue
            if name in ("", None):
                continue

            row_brand = fixed_brand if "brand" in cfg else normalize_brand(ws.cell_value(r, cfg["brand_col"]))
            if row_brand is None or row_brand not in contragent_brands:
                continue

            art = norm(art_raw) if art_raw not in ("", None) else None
            if art is not None:
                if art in existing_articles or art in seen:
                    continue
                seen.add(art)

            candidates.append((name, art, qty))

    return candidates


def build_report(src_file, delivery_file):
    """
    src_file / delivery_file - путь на диске (str), файлоподобный объект
    (BytesIO, Streamlit UploadedFile) либо сырые bytes. Формат каждого
    файла (.xlsx или .xls) определяется по содержимому автоматически,
    расширение в имени файла роли не играет.
    Возвращает (openpyxl.Workbook с готовым отчётом, dict со статистикой).
    """
    src_bytes = read_bytes(src_file)
    src_fmt = detect_excel_format(src_bytes)
    if src_fmt == "xlsx":
        wb = openpyxl.load_workbook(io.BytesIO(src_bytes), data_only=True)
    else:
        wb = _XlsAsOpenpyxlWorkbook(xlrd.open_workbook(file_contents=src_bytes))

    am_store_cols, am_desc, am_data = parse_am_sheet(wb["АМ"])
    pr_store_periods, pr_desc, pr_data = parse_sales_sheet(wb["Продажі 2025"])

    delivery_wb = open_delivery_workbook(delivery_file)
    delivery_data = parse_delivery_schedule(delivery_wb)

    all_stores = sorted(set(am_store_cols) | set(pr_store_periods))
    all_keys = sorted(set(am_desc) | set(pr_desc))

    STORE_FIELDS = [GROUP1_LABEL, GROUP2_LABEL, "АМ", "Залишок", "НОВА АМ"]

    out_wb = openpyxl.Workbook()
    out_ws = out_wb.active
    out_ws.title = "Звід"

    n_desc_cols = len(DESC_COLS) + len(PRODUCT_FIELDS)

    # Header row 1: store name repeated across its block of columns
    for i, store in enumerate(all_stores):
        start_col = n_desc_cols + 1 + i * len(STORE_FIELDS)
        for j in range(len(STORE_FIELDS)):
            out_ws.cell(row=1, column=start_col + j, value=store)

    # Header row 2: descriptive + product-level field names, then repeating store field names
    for i, name in enumerate(DESC_COLS + PRODUCT_FIELDS):
        out_ws.cell(row=2, column=i + 1, value=name)
    for i, store in enumerate(all_stores):
        start_col = n_desc_cols + 1 + i * len(STORE_FIELDS)
        for j, field in enumerate(STORE_FIELDS):
            out_ws.cell(row=2, column=start_col + j, value=field)

    out_ws.freeze_panes = out_ws.cell(row=3, column=n_desc_cols + 1)

    n_rows = 0
    for r, key in enumerate(all_keys, start=3):
        desc = am_desc.get(key) or pr_desc.get(key) or {}
        for i, name in enumerate(DESC_COLS):
            out_ws.cell(row=r, column=i + 1, value=key if name == "Но_" else desc.get(name))

        artikul = norm(desc.get("Артикул"))
        delivery_qty = delivery_data.get(artikul)  # None = артикула нет в графіку; 0 = є, але зараз нічого не їде
        out_ws.cell(row=r, column=len(DESC_COLS) + 1, value=delivery_qty)

        for i, store in enumerate(all_stores):
            pr_vals = pr_data.get((key, store))
            am_val, zal_val = am_data.get((key, store), (None, None))

            g1 = group_sum(pr_vals, GROUP1_MONTHS) if pr_vals else None
            g2 = group_sum(pr_vals, GROUP2_MONTHS) if pr_vals else None

            start_col = n_desc_cols + 1 + i * len(STORE_FIELDS)
            out_ws.cell(row=r, column=start_col, value=g1)
            out_ws.cell(row=r, column=start_col + 1, value=g2)
            out_ws.cell(row=r, column=start_col + 2, value=am_val)
            out_ws.cell(row=r, column=start_col + 3, value=zal_val)
            out_ws.cell(row=r, column=start_col + 4, value=None)
        n_rows += 1

    # Кандидаты в новинки - отдельными строками внизу: название, артикул
    # (если он есть - без подмены) и кол-во "в дорозі". Остальные поля пустые.
    # Ищем только среди брендов, которые реально есть у этого контрагента.
    existing_articles = {norm((am_desc.get(k) or pr_desc.get(k) or {}).get("Артикул")) for k in all_keys}
    existing_articles.discard(None)
    contragent_brands = {
        normalize_brand((am_desc.get(k) or pr_desc.get(k) or {}).get("Торговая Марка")) for k in all_keys
    }
    contragent_brands.discard(None)
    novelty_candidates = find_novelty_candidates(delivery_wb, existing_articles, contragent_brands)

    name_col = DESC_COLS.index("Наименование товара") + 1
    art_col = DESC_COLS.index("Артикул") + 1
    vdorozi_col = len(DESC_COLS) + 1
    r = 3 + n_rows
    for name, art, qty in novelty_candidates:
        out_ws.cell(row=r, column=name_col, value=name)
        out_ws.cell(row=r, column=art_col, value=art)
        out_ws.cell(row=r, column=vdorozi_col, value=qty)
        r += 1
    last_row = r - 1

    # "НОВА АМ" - светло-зелёная заливка (заголовок + все строки данных)
    nova_am_field_index = STORE_FIELDS.index("НОВА АМ")
    for i in range(len(all_stores)):
        start_col = n_desc_cols + 1 + i * len(STORE_FIELDS)
        nova_am_col = start_col + nova_am_field_index
        out_ws.cell(row=2, column=nova_am_col).fill = NOVA_AM_FILL
        for row in range(3, last_row + 1):
            out_ws.cell(row=row, column=nova_am_col).fill = NOVA_AM_FILL

    stats = {
        "n_products": len(all_keys),
        "n_stores": len(all_stores),
        "n_rows": n_rows,
        "n_novelty": len(novelty_candidates),
    }
    return out_wb, stats


def main():
    out_wb, stats = build_report(SRC, DELIVERY_SRC)
    out_wb.save(OUT)
    print(f"Готово: {OUT}")
    print(f"Товаров: {stats['n_products']}, магазинов: {stats['n_stores']}, строк: {stats['n_rows']}")
    print(f"Кандидатов в новинки (без статуса, кол-во > 0): {stats['n_novelty']}")


if __name__ == "__main__":
    main()
