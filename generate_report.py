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
- Лист продаж (название ищем по подстроке "Продаж", год в названии не
  хардкодим - см. find_sheet_by_name_part): для каждого магазина суммируем
  "Сумма по столбцу Кол-во Шт" по месяцам, распределяя их в два периода:
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
import math
import re
import zipfile
from datetime import date

import openpyxl
import xlrd
from openpyxl.styles import PatternFill

NOVA_AM_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
NOVA_AM_PEACH_FILL = PatternFill(start_color="FFDAB9", end_color="FFDAB9", fill_type="solid")
NOVA_AM_YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

SRC = "АМ+продажі+залишки_Міленіум_контрагент.xlsx"
DELIVERY_SRC = "Графік поставок 04,09,2026.xls"
PRICE_SRC = "Прайс-лист 11.09.2026.xlsx"
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


def load_xlsx(data):
    """
    openpyxl.load_workbook з обходом файлів, де частина архіву
    xl/sharedStrings.xml має неправильний регістр (наприклад
    xl/SharedStrings.xml) - таке трапляється у деяких вивантаженнях з 1С.
    Сам архів при цьому валідний, просто openpyxl очікує точну назву.
    """
    try:
        return openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    except KeyError as e:
        if "sharedStrings" not in str(e):
            raise
        fixed = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(data)) as zin, zipfile.ZipFile(fixed, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                content = zin.read(item.filename)
                name = item.filename
                if name.lower() == "xl/sharedstrings.xml" and name != "xl/sharedStrings.xml":
                    name = "xl/sharedStrings.xml"
                zout.writestr(name, content)
        fixed.seek(0)
        return openpyxl.load_workbook(fixed, data_only=True)


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

    @property
    def sheetnames(self):
        return self._book.sheet_names()


class _XlsxAsXlrdSheet:
    """Оборачивает лист openpyxl (.xlsx, 1-індексація) під API листа xlrd (0-індексація)."""

    def __init__(self, ws):
        self._ws = ws
        self.nrows = ws.max_row
        self.ncols = ws.max_column

    def cell_value(self, row, col):
        value = self._ws.cell(row=row + 1, column=col + 1).value
        return value if value is not None else ""

    @property
    def merged_cells(self):
        """Список (row1, row2, col1, col2) у конвенції xlrd (0-індексація, кінець не включно)."""
        return [
            (rng.min_row - 1, rng.max_row, rng.min_col - 1, rng.max_col)
            for rng in self._ws.merged_cells.ranges
        ]


class _XlsxAsXlrdBook:
    def __init__(self, wb):
        self._wb = wb

    def sheet_names(self):
        return self._wb.sheetnames

    def sheet_by_name(self, name):
        return _XlsxAsXlrdSheet(self._wb[name])

    def sheet_by_index(self, index):
        return _XlsxAsXlrdSheet(self._wb[self._wb.sheetnames[index]])

DESC_COLS = [
    "Поставщик", "Код Группы", "Торговая Марка", "Но_",
    "Наименование товара", "Артикул", "Штрихкод", "Статус товара",
]

PRODUCT_FIELDS = ["В дорозі", "Статус артикула", "Склад", "Категорія"]

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
    # Kids2: колонки задані не абсолютними номерами, а зсувом від колонки
    # "Номенклатура" (anchor_header) - у вивантаженні від 25.09.2026 перед
    # "Номенклатура" з'явилась ще одна колонка "Артикул", і все зсунулось
    # на 1 позицію праворуч; абсолютні номери зламались би, а відносний
    # зсув від "Номенклатура" лишився той самий в обох файлах.
    "Kids2": {"header_row": 2, "data_start": 4, "anchor_header": "Номенклатура",
              "name_offset": 0, "art_offset": 1, "status_offset": 3, "brand_offset": 5,
              "qty_offsets": [7, 8, 9]},
    "Kendamil": {"header_row": 1, "data_start": 2, "name_col": 1, "art_col": 0, "status_col": 2, "qty": 6,
                 "brand": "KENDAMIL"},
    "Offspring": {"header_row": 1, "data_start": 2, "name_col": 1, "art_col": 0, "status_col": 2, "qty": "dynamic",
                  "brand": "OFFSPRING"},
}


def _resolve_delivery_cfg(ws, cfg):
    """
    Якщо cfg заданий через "anchor_header" + відносні зсуви (зараз - тільки
    Kids2) - шукає колонку з таким заголовком (у header_row) і повертає
    новий cfg з уже обчисленими абсолютними "name_col"/"art_col"/
    "status_col"/"brand_col"/"qty". Так лист лишається робочим, навіть
    якщо 1С додасть чи прибере колонку ПЕРЕД якірною - зсув від якоря не
    зміниться. Якщо cfg і так заданий абсолютними номерами - повертає без
    змін.
    """
    if "anchor_header" not in cfg:
        return cfg
    header_row = cfg["header_row"]
    anchor = _norm_header(cfg["anchor_header"])
    anchor_col = None
    for c in range(ws.ncols):
        if _norm_header(ws.cell_value(header_row, c)) == anchor:
            anchor_col = c
            break
    if anchor_col is None:
        raise KeyError(f'Не знайдено колонку "{cfg["anchor_header"]}" у рядку заголовків листа')

    resolved = dict(cfg)
    resolved["name_col"] = anchor_col + cfg["name_offset"]
    resolved["art_col"] = anchor_col + cfg["art_offset"]
    resolved["status_col"] = anchor_col + cfg["status_offset"]
    if "brand_offset" in cfg:
        resolved["brand_col"] = anchor_col + cfg["brand_offset"]
    resolved["qty"] = [anchor_col + o for o in cfg["qty_offsets"]]
    return resolved


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


def find_sheet_by_name_part(wb, name_part):
    """
    Знаходить перший лист книги, назва якого містить name_part
    (регістронезалежно) - наприклад "Продажі 2025" чи "Продажі 2026"
    обидва знайдуться за name_part="Продаж", без хардкоду конкретного
    року. Якщо збігів декілька - береться перший за порядком листів у
    книзі. KeyError зі списком наявних листів, якщо жодного не знайдено.
    """
    name_part_lower = name_part.lower()
    for name in wb.sheetnames:
        if name_part_lower in name.lower():
            return wb[name]
    raise KeyError(
        f'Не знайдено лист, назва якого містить "{name_part}" '
        f'(наявні листи: {", ".join(wb.sheetnames)})'
    )


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
        # formatting_info нужен, чтобы получить merged_cells (для довідника Епіцентру - АМ).
        return xlrd.open_workbook(file_contents=data, formatting_info=True)
    return _XlsxAsXlrdBook(load_xlsx(data))


def parse_delivery_schedule(wb):
    """Возвращает {Артикул: кол-во в дорозі}, суммируя совпадения по всем листам."""
    result = {}

    for sheet_name, raw_cfg in DELIVERY_SHEETS.items():
        if sheet_name not in wb.sheet_names():
            continue
        ws = wb.sheet_by_name(sheet_name)
        cfg = _resolve_delivery_cfg(ws, raw_cfg)
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

    for sheet_name, raw_cfg in DELIVERY_SHEETS.items():
        if not raw_cfg.get("novelty_reliable"):
            continue
        if sheet_name not in wb.sheet_names():
            continue
        ws = wb.sheet_by_name(sheet_name)
        cfg = _resolve_delivery_cfg(ws, raw_cfg)
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


# --- Прайс-лист з 1С УТП (звіт "Прайс-лист (СКД)") - для "Артикул", "Бренд",
# "Статус артикула" і "Склад" ---
#
# Формат НЕ фіксований (звіт СКД можна перебудувати з іншими налаштуваннями),
# тому колонки шукаються за назвою заголовка, а не за позицією:
#   - ключ джойну - "Артикул" (для Антошки) або "Артикул в сети" (для
#     Епіцентру, джойн з "Артикул мережі" - див. key_header у
#     parse_price_list);
#   - "Артикул" і "Бренд" - віддаються як є (для Епіцентру - там своїх
#     немає; для Антошки, де key_header="Артикул", "Артикул" збігається
#     з ключем і фактично не використовується);
#   - "Номенклатура" - за її наявністю відрізняємо рядок товару від рядка-
#     заголовка товарної категорії (в СКД-звіті категорії йдуть окремими
#     рядками, де заповнена лише перша колонка);
#   - "Статус артикула";
#   - "Остаток" у групі "Итого" (в рядку заголовків груп над заголовками
#     колонок стоїть "Итого" саме над потрібною колонкою "Остаток" - таких
#     колонок "Остаток" декілька, по одній на кожен склад/канал, тому
#     важливо взяти саме ту, що під "Итого").

PRICE_KEY_HEADER = "артикул"
PRICE_NETWORK_KEY_HEADER = "артикул в сети"
PRICE_NAME_HEADER = "номенклатура"
PRICE_BRAND_HEADER = "бренд"
PRICE_STATUS_HEADER = "статус артикула"
PRICE_STOCK_HEADER = "остаток"
PRICE_STOCK_GROUP = "итого"


def parse_price_list(wb, key_header=PRICE_KEY_HEADER):
    """
    Розбирає прайс-лист. key_header - яка колонка є ключем джойну
    ("артикул" за замовчуванням, або "артикул в сети" для Епіцентру).
    Повертає {ключ: (артикул, бренд, статус, склад)}.
    Порожній словник, якщо на листі немає потрібного заголовка-ключа.
    """
    ws = wb.sheet_by_index(0)

    header_row = None
    for r in range(ws.nrows):
        for c in range(ws.ncols):
            if _norm_header(ws.cell_value(r, c)) == key_header:
                header_row = r
                break
        if header_row is not None:
            break
    if header_row is None:
        return {}

    key_col = art_col = name_col = brand_col = status_col = stock_col = None
    for c in range(ws.ncols):
        h = _norm_header(ws.cell_value(header_row, c))
        if h == key_header and key_col is None:
            key_col = c
        elif h == PRICE_KEY_HEADER and art_col is None:
            art_col = c
        elif h == PRICE_NAME_HEADER and name_col is None:
            name_col = c
        elif h == PRICE_BRAND_HEADER and brand_col is None:
            brand_col = c
        elif h == PRICE_STATUS_HEADER and status_col is None:
            status_col = c
        elif h == PRICE_STOCK_HEADER and stock_col is None:
            group = _norm_header(ws.cell_value(header_row - 1, c)) if header_row > 0 else None
            if group == PRICE_STOCK_GROUP:
                stock_col = c

    if key_col is None:
        return {}

    result = {}
    for r in range(header_row + 1, ws.nrows):
        if name_col is not None and ws.cell_value(r, name_col) in ("", None):
            continue  # рядок-заголовок товарної категорії, не товар
        key = ws.cell_value(r, key_col)
        if key in ("", None):
            continue
        art = ws.cell_value(r, art_col) if art_col is not None else None
        brand = ws.cell_value(r, brand_col) if brand_col is not None else None
        status = ws.cell_value(r, status_col) if status_col is not None else None
        stock = ws.cell_value(r, stock_col) if stock_col is not None else None
        result[norm(key)] = (
            norm(art) if art not in ("", None) else None,
            str(brand).strip() if brand not in ("", None) else None,
            status if status not in ("", None) else None,
            stock if isinstance(stock, (int, float)) else None,
        )

    return result


# --- Підрахунок СКЮ іграшок у прайс-листі (СКД) ---
#
# Прайс-лист - ієрархічний звіт: категорії товарів вкладені одна в одну
# (бренд -> категорія -> підкатегорія -> ...), і ця вкладеність закодована
# не текстом, а Excel-групуванням рядків (row_dimensions[r].outlineLevel -
# те саме "згорнути/розгорнути" з панелі зліва в Excel). Категорія - рядок,
# де заповнена тільки колонка "Артикул" (там лежить назва категорії, а не
# сам артикул); товар - рядок, де заповнена "Номенклатура".
#
# Кількість і склад брендів у файлі НЕ фіксовані (сьогодні Chicco/
# Clementoni/KIDS2, завтра можуть бути інші), тому пошук не прив'язаний до
# конкретних назв брендів - береться будь-яка категорія (на будь-якому
# рівні вкладеності), де в назві є "іграш" або "игрушк" (регістронезалежно,
# покриває "Іграшки"/"іграшка"/"Игрушки"/укр. і рос. написання). Усі товари
# в піддереві такої категорії (разом з вкладеними підкатегоріями, навіть
# якщо в ЇХ власній назві "іграш"/"игрушк" немає) рахуються як іграшки -
# піддерево закінчується на першій наступній категорії того самого рівня
# вкладеності або вищого. Бренд для розбивки результату береться з поля
# "Бренд" самого товару (а не з назви категорії, де він фізично лежить) -
# трапляються крос-брендові товари (наприклад товар під категорією
# "Ingenuity іграшка", але з "Бренд" = "Bright Starts" - обидва суббренди
# одного правовласника).
#
# Винятки (щоб не зловити аксесуари ДЛЯ іграшок, а не самі іграшки,
# наприклад "Контейнери для іграшок Stokke® MuTable"): категорія НЕ
# вважається іграшковою, навіть з "іграш"/"игрушк" у назві, якщо там
# також є "для іграшок", "для игрушек" або "контейнер".

TOY_NAME_PATTERNS = ("іграш", "игрушк")
TOY_NAME_EXCLUDE_PATTERNS = ("для іграшок", "для игрушек", "контейнер")


def _is_toy_category_name(name):
    if not isinstance(name, str):
        return False
    lname = name.strip().lower()
    if any(p in lname for p in TOY_NAME_EXCLUDE_PATTERNS):
        return False
    return any(p in lname for p in TOY_NAME_PATTERNS)


def count_toy_skus(file):
    """
    Розбирає прайс-лист (звіт "Прайс-лист (СКД)" з 1С УТП - той самий
    файл і формат, що і parse_price_list) і рахує кількість СКЮ іграшок,
    окремо по кожному бренду ("Бренд" товару): загальну кількість і
    скільки з них мають заповнене "Артикул в сети".
    Повертає {бренд: {"total": ..., "with_artikul_v_seti": ...}}. Порожній
    словник, якщо на листі немає заголовка "Артикул" або "Номенклатура".
    """
    wb = load_xlsx(read_bytes(file))
    ws = wb[wb.sheetnames[0]]

    header_row = None
    for r in range(1, min(ws.max_row, 10) + 1):
        for c in range(1, ws.max_column + 1):
            if _norm_header(ws.cell(row=r, column=c).value) == PRICE_KEY_HEADER:
                header_row = r
                break
        if header_row is not None:
            break
    if header_row is None:
        return {}

    key_col = name_col = brand_col = network_key_col = None
    for c in range(1, ws.max_column + 1):
        h = _norm_header(ws.cell(row=header_row, column=c).value)
        if h == PRICE_KEY_HEADER and key_col is None:
            key_col = c
        elif h == PRICE_NAME_HEADER and name_col is None:
            name_col = c
        elif h == PRICE_BRAND_HEADER and brand_col is None:
            brand_col = c
        elif h == PRICE_NETWORK_KEY_HEADER and network_key_col is None:
            network_key_col = c

    if key_col is None or name_col is None:
        return {}

    counts = {}
    in_toy_subtree = False
    toy_level = None
    for r in range(header_row + 1, ws.max_row + 1):
        col1 = ws.cell(row=r, column=key_col).value
        nomenklatura = ws.cell(row=r, column=name_col).value
        is_category = col1 not in ("", None) and nomenklatura in ("", None)

        if is_category:
            rd = ws.row_dimensions.get(r)
            lvl = rd.outlineLevel if rd else 0
            if in_toy_subtree and lvl <= toy_level:
                in_toy_subtree = False
                toy_level = None
            if not in_toy_subtree and _is_toy_category_name(col1):
                in_toy_subtree = True
                toy_level = lvl
            continue

        if not in_toy_subtree:
            continue

        brand = ws.cell(row=r, column=brand_col).value if brand_col else None
        brand = str(brand).strip() if brand not in ("", None) else "(без бренду)"
        bucket = counts.setdefault(brand, {"total": 0, "with_artikul_v_seti": 0})
        bucket["total"] += 1
        network_key = ws.cell(row=r, column=network_key_col).value if network_key_col else None
        if network_key not in ("", None):
            bucket["with_artikul_v_seti"] += 1

    return counts


# --- АВС-аналіз продажів (звіт "АВС-аналіз продажів (за номенклатурою)" з
# 1С УТП) - для поля "Категорія" ---
#
# Формат НЕ фіксований, тому колонки шукаються за назвою заголовка:
#   - "Артикул" - основний ключ джойну;
#   - "АВС-клас" - значення, яке стає "Категорія";
#   - "Групування" - назва товару (в цьому звіті колонка з назвою для
#     рядків-товарів називається саме так, а не "Номенклатура" - тому
#     що вона ж використовується для підсумкових рядків класів A/B/C,
#     де замість назви стоїть "A - класс" тощо).
# Рядок вважається товаром, тільки якщо в ньому є "Артикул" - підсумкові
# рядки класів (де замість артикула просто підсумок по класу) пропускаються.
# Можна завантажити декілька файлів АВС-аналізу відразу (по одному на
# бренд/категорію) - результати об'єднуються.

ABC_KEY_HEADER = "артикул"
ABC_CLASS_HEADER = "авс-клас"
ABC_NAME_HEADER = "групування"
ABC_QTY_HEADER = "продажі, шт"


def parse_abc_analysis(wb):
    """
    Розбирає один файл АВС-аналізу продажів. Повертає
    (by_article, by_name, qty_by_article, qty_by_name):
      by_article - {артикул: АВС-клас}
      by_name    - {точна назва товару: АВС-клас} - для резервного джойну,
                   якщо артикул не співпаде (без жодних відхилень - точний
                   рядок)
      qty_by_article / qty_by_name - те саме, але значення "Продажі, шт"
                   (для сортування товарів усередині класу за кількістю -
                   автозаповнення "НОВА АМ").
    Порожні словники, якщо на листі немає заголовка "Артикул".
    """
    ws = wb.sheet_by_index(0)

    header_row = None
    for r in range(ws.nrows):
        for c in range(ws.ncols):
            if _norm_header(ws.cell_value(r, c)) == ABC_KEY_HEADER:
                header_row = r
                break
        if header_row is not None:
            break
    if header_row is None:
        return {}, {}, {}, {}

    key_col = class_col = name_col = qty_col = None
    for c in range(ws.ncols):
        h = _norm_header(ws.cell_value(header_row, c))
        if h == ABC_KEY_HEADER and key_col is None:
            key_col = c
        elif h == ABC_CLASS_HEADER and class_col is None:
            class_col = c
        elif h == ABC_NAME_HEADER and name_col is None:
            name_col = c
        elif h == ABC_QTY_HEADER and qty_col is None:
            qty_col = c

    if key_col is None or class_col is None:
        return {}, {}, {}, {}

    by_article = {}
    by_name = {}
    qty_by_article = {}
    qty_by_name = {}
    for r in range(header_row + 1, ws.nrows):
        art = ws.cell_value(r, key_col)
        if art in ("", None):
            continue  # підсумковий рядок класу (A/B/C), не товар
        cls = ws.cell_value(r, class_col)
        if cls in ("", None):
            continue
        art_norm = norm(art)
        by_article[art_norm] = str(cls).strip()
        name = ws.cell_value(r, name_col) if name_col is not None else None
        name = str(name).strip() if name not in ("", None) else None
        if name is not None:
            by_name[name] = str(cls).strip()
        qty = ws.cell_value(r, qty_col) if qty_col is not None else None
        if isinstance(qty, (int, float)):
            qty_by_article[art_norm] = qty
            if name is not None:
                qty_by_name[name] = qty

    return by_article, by_name, qty_by_article, qty_by_name


def parse_abc_files(files):
    """
    files - список файлів (шлях/файлоподібний об'єкт/bytes), кожен - один
    файл АВС-аналізу. Об'єднує результати всіх файлів в один
    (by_article, by_name, qty_by_article, qty_by_name); при дублікатах
    пізніший файл у списку переважає.
    """
    by_article, by_name = {}, {}
    qty_by_article, qty_by_name = {}, {}
    for f in files:
        wb = open_delivery_workbook(f)
        a, n, qa, qn = parse_abc_analysis(wb)
        by_article.update(a)
        by_name.update(n)
        qty_by_article.update(qa)
        qty_by_name.update(qn)
    return by_article, by_name, qty_by_article, qty_by_name


def lookup_category(art, name, abc_by_article, abc_by_name):
    """Категорія за артикулом, або (якщо не знайдено) за точною назвою."""
    if art is not None:
        cat = abc_by_article.get(art)
        if cat is not None:
            return cat
    if name is not None:
        return abc_by_name.get(str(name).strip())
    return None


def lookup_abc_qty(art, name, qty_by_article, qty_by_name):
    """"Продажі, шт" за тим самим товаром, за тією ж логікою (артикул,
    резервно - точна назва), що і lookup_category."""
    if art is not None:
        qty = qty_by_article.get(art)
        if qty is not None:
            return qty
    if name is not None:
        return qty_by_name.get(str(name).strip())
    return None


def _class_letter(value):
    """Витягує літеру класу (A/B/C) зі значення "Категорія" на кшталт
    "A - класс" чи "B - клас". None, якщо не розпізнано."""
    if not isinstance(value, str):
        return None
    v = value.strip().upper()
    return v[0] if v[:1] in ("A", "B", "C") else None


def _quantity_cutoff(quantities, fraction):
    """Поріг кількості для "верхні `fraction`*100% позицій класу за
    спаданням кількості": товар входить у цю частку, якщо його власна
    кількість >= порогу. None, якщо quantities порожній."""
    if not quantities:
        return None
    sorted_desc = sorted(quantities, reverse=True)
    cutoff_idx = min(len(sorted_desc), max(1, math.ceil(len(sorted_desc) * fraction))) - 1
    return sorted_desc[cutoff_idx]


def class_quantity_cutoff(abc_by_article, qty_by_article, target_class, fraction):
    """Рахує поріг _quantity_cutoff серед усіх товарів з класом
    target_class (за даними усіх завантажених файлів АВС-аналізу разом)."""
    quantities = [
        qty_by_article[art] for art, cls in abc_by_article.items()
        if _class_letter(cls) == target_class and art in qty_by_article
    ]
    return _quantity_cutoff(quantities, fraction)


def _has_sales(*period_values):
    return any(isinstance(v, (int, float)) and v > 0 for v in period_values)


def _has_stock(value):
    return isinstance(value, (int, float)) and value > 0


def antoshka_assortment_match(store_class, product_class, product_qty, c_cutoff):
    """Чи має товар бути представлений у магазині цього класу (Антошка):
    A - усі класи A/B/C; B - класи A/B; C - класи A/B, і з класу C -
    тільки верхні 75% за кількістю (c_cutoff - поріг з class_quantity_cutoff).
    None, якщо клас магазину чи товару невідомий (немає з чим звіряти)."""
    if store_class is None or product_class is None:
        return None
    if store_class == "A":
        return product_class in ("A", "B", "C")
    if store_class == "B":
        return product_class in ("A", "B")
    if store_class == "C":
        if product_class in ("A", "B"):
            return True
        if product_class == "C":
            return c_cutoff is not None and product_qty is not None and product_qty >= c_cutoff
        return False
    return None


def epicentr_assortment_match(store_class, product_class, product_qty, b_cutoff):
    """Те саме для Епіцентру: A - категорії A/B; B - категорія A, і з
    категорії B - тільки верхні 50% за кількістю (b_cutoff); C - тільки
    категорія A. Категорія C товару в асортимент Епіцентру не входить
    ні в одному класі магазину."""
    if store_class is None or product_class is None:
        return None
    if product_class == "C":
        return False
    if store_class == "A":
        return product_class in ("A", "B")
    if store_class == "B":
        if product_class == "A":
            return True
        if product_class == "B":
            return b_cutoff is not None and product_qty is not None and product_qty >= b_cutoff
        return False
    if store_class == "C":
        return product_class == "A"
    return None


def build_report(src_file, delivery_file, price_file=None, abc_files=None, network_sales_file=None):
    """
    src_file / delivery_file - путь на диске (str), файлоподобный объект
    (BytesIO, Streamlit UploadedFile) либо сырые bytes. Формат каждого
    файла (.xlsx или .xls) определяется по содержимому автоматически,
    расширение в имени файла роли не играет.
    network_sales_file (необов'язковий) - файл "Мережі продажі" (звіт
    1С "Аналіз продаж по сетям товарів на реалізації"). Дає АВС-аналіз
    магазинів за категорією "Іграшка": новий рядок над назвами магазинів
    (АВС-клас кожного) і один лист Total (повний список магазинів цього
    контрагента з "Продажі", "Серед.продажі в міс." і АВС-класом,
    відсортований за спаданням середніх продажів).
    Без цього файлу нова строка і лист Total просто не додаються.
    Возвращает (openpyxl.Workbook с готовым отчётом, dict со статистикой).
    """
    src_bytes = read_bytes(src_file)
    src_fmt = detect_excel_format(src_bytes)
    if src_fmt == "xlsx":
        wb = load_xlsx(src_bytes)
    else:
        wb = _XlsAsOpenpyxlWorkbook(xlrd.open_workbook(file_contents=src_bytes))

    am_store_cols, am_desc, am_data = parse_am_sheet(wb["АМ"])
    pr_store_periods, pr_desc, pr_data = parse_sales_sheet(find_sheet_by_name_part(wb, "Продаж"))

    delivery_wb = open_delivery_workbook(delivery_file)
    delivery_data = parse_delivery_schedule(delivery_wb)

    price_data = parse_price_list(open_delivery_workbook(price_file)) if price_file else {}
    abc_by_article, abc_by_name, qty_by_article, qty_by_name = (
        parse_abc_files(abc_files) if abc_files else ({}, {}, {}, {})
    )
    c_cutoff = class_quantity_cutoff(abc_by_article, qty_by_article, "C", 0.75)

    all_stores = sorted(set(am_store_cols) | set(pr_store_periods))
    all_keys = sorted(set(am_desc) | set(pr_desc))

    STORE_FIELDS = [GROUP1_LABEL, GROUP2_LABEL, "АМ", "Залишок", "НОВА АМ"]

    out_wb = openpyxl.Workbook()
    out_ws = out_wb.active
    out_ws.title = "Звід"

    n_desc_cols = len(DESC_COLS) + len(PRODUCT_FIELDS)

    network_data = {}
    if network_sales_file:
        network_data = parse_network_sales_category(open_delivery_workbook(network_sales_file), "Іграшка")
    network_stores = {k: v for k, v in network_data.items() if k in set(all_stores)}
    abc_result = classify_store_abc(network_stores) if network_stores else {}
    if network_stores:
        total_ws = out_wb.create_sheet(title="Total")
        _write_store_totals_sheet(total_ws, network_stores, abc_result)

    # Header row 1: АВС-клас магазину (з "Мережі продажі"), над назвою магазину
    for i, store in enumerate(all_stores):
        start_col = n_desc_cols + 1 + i * len(STORE_FIELDS)
        for j in range(len(STORE_FIELDS)):
            out_ws.cell(row=1, column=start_col + j, value=abc_result.get(_normalize_store_key(store)))

    # Header row 2: store name repeated across its block of columns
    for i, store in enumerate(all_stores):
        start_col = n_desc_cols + 1 + i * len(STORE_FIELDS)
        for j in range(len(STORE_FIELDS)):
            out_ws.cell(row=2, column=start_col + j, value=store)

    # Header row 3: descriptive + product-level field names, then repeating store field names
    for i, name in enumerate(DESC_COLS + PRODUCT_FIELDS):
        out_ws.cell(row=3, column=i + 1, value=name)
    for i, store in enumerate(all_stores):
        start_col = n_desc_cols + 1 + i * len(STORE_FIELDS)
        for j, field in enumerate(STORE_FIELDS):
            out_ws.cell(row=3, column=start_col + j, value=field)

    out_ws.freeze_panes = out_ws.cell(row=4, column=n_desc_cols + 1)

    n_rows = 0
    for r, key in enumerate(all_keys, start=4):
        desc = am_desc.get(key) or pr_desc.get(key) or {}
        for i, name in enumerate(DESC_COLS):
            out_ws.cell(row=r, column=i + 1, value=key if name == "Но_" else desc.get(name))

        artikul = norm(desc.get("Артикул"))
        delivery_qty = delivery_data.get(artikul)  # None = артикула нет в графіку; 0 = є, але зараз нічого не їде
        _, _, price_status, price_sklad = price_data.get(artikul, (None, None, None, None))
        category = lookup_category(artikul, desc.get("Наименование товара"), abc_by_article, abc_by_name)
        out_ws.cell(row=r, column=len(DESC_COLS) + 1, value=delivery_qty)
        out_ws.cell(row=r, column=len(DESC_COLS) + 2, value=price_status)
        out_ws.cell(row=r, column=len(DESC_COLS) + 3, value=price_sklad)
        out_ws.cell(row=r, column=len(DESC_COLS) + 4, value=category)

        status_is_8 = price_status == 8
        product_class = _class_letter(category)
        product_qty = lookup_abc_qty(artikul, desc.get("Наименование товара"), qty_by_article, qty_by_name)

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

            nova_am_cell = out_ws.cell(row=r, column=start_col + 4)
            if status_is_8:
                nova_am_cell.value = 0
            elif not _has_sales(g1, g2):
                if _has_stock(zal_val):
                    nova_am_cell.value = 0
                else:
                    nova_am_cell.fill = NOVA_AM_PEACH_FILL
            else:
                store_class = abc_result.get(_normalize_store_key(store))
                match = antoshka_assortment_match(store_class, product_class, product_qty, c_cutoff)
                if match is True:
                    nova_am_cell.value = 1
                else:
                    nova_am_cell.fill = NOVA_AM_YELLOW_FILL
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
    status_col = len(DESC_COLS) + 2
    r = 4 + n_rows
    for name, art, qty in novelty_candidates:
        out_ws.cell(row=r, column=name_col, value=name)
        out_ws.cell(row=r, column=art_col, value=art)
        out_ws.cell(row=r, column=vdorozi_col, value=qty)
        out_ws.cell(row=r, column=status_col, value="NEW")
        for i in range(len(all_stores)):
            start_col = n_desc_cols + 1 + i * len(STORE_FIELDS)
            out_ws.cell(row=r, column=start_col + 4, value=1)
        r += 1
    last_row = r - 1

    # Заголовок "НОВА АМ" - світло-зелена заливка як орієнтир по колонці.
    # Заливка самих даних - за результатом розрахунку вище (NOVA_AM_PEACH_FILL
    # / NOVA_AM_YELLOW_FILL для незаповнених клітинок, без заливки - для
    # клітинок з уже проставленим 0/1).
    nova_am_field_index = STORE_FIELDS.index("НОВА АМ")
    for i in range(len(all_stores)):
        start_col = n_desc_cols + 1 + i * len(STORE_FIELDS)
        nova_am_col = start_col + nova_am_field_index
        out_ws.cell(row=3, column=nova_am_col).fill = NOVA_AM_FILL

    stats = {
        "n_products": len(all_keys),
        "n_stores": len(all_stores),
        "n_rows": n_rows,
        "n_novelty": len(novelty_candidates),
        "n_network_stores": len(network_stores),
    }
    return out_wb, stats


def main():
    out_wb, stats = build_report(SRC, DELIVERY_SRC, PRICE_SRC)
    out_wb.save(OUT)
    print(f"Готово: {OUT}")
    print(f"Товаров: {stats['n_products']}, магазинов: {stats['n_stores']}, строк: {stats['n_rows']}")
    print(f"Кандидатов в новинки (без статуса, кол-во > 0): {stats['n_novelty']}")


# ============================================================
# Епіцентр
#
# Джерело - два звіти "Звіт про продажу товарів" (кожен: один аркуш,
# нема нормальної табличної структури): магазин (рядок, де у колонці B
# стоїть "-", а в колонці A - код магазину, наприклад "01", "BR", "CHB")
# і під ним рядки товарів (колонка A - 8-значний числовий код товару,
# колонка B - назва, колонка D - "Розхід" = кількість продажу за період
# файлу, колонка E - "Кінцевий залишок"). Рядок загального підсумку
# (колонка A - число 1, не рядок з кодом магазину) і рядки-роздільники
# пропускаються - вони не мають назви товару.
#
# Два файли: "сезон" (проксі - січень) і "не сезон" (проксі - серпень) -
# кожен дає свою кількість продажу за свій період. "Залишок" - єдина
# колонка на магазин: береться НЕ завжди з файлу "не сезон" - для кожного
# з двох файлів читаємо дату з шапки "Кінцевий залишок на <дата>" (рядок 4,
# колонка E) і беремо залишок з того файлу, де ця дата пізніша (реально
# найсвіжіший знімок на момент формування звіту, а не за умовчанням
# "не сезон"). Нюанс: ця клітинка об'єднана на 4 рядки вниз (рядки 4-7),
# текст лежить тільки в "якірній" клітинці (рядок 4) - решта порожні.
#
# "Артикул мережі" = той самий 8-значний код (єдиний ідентифікатор товару,
# який взагалі є в цих файлах). "Артикул", "Бренд", "Статус артикула" і
# "Склад" - якщо переданий файл-прайс, підтягуються з нього за збігом
# "Артикул мережі" ↔ "Артикул в сети" (parse_price_list з
# key_header="артикул в сети"). "В дорозі" - якщо переданий графік
# поставок, рахується так само, як у Антошки (parse_delivery_schedule),
# але джойн за щойно підтягнутим "Артикул" (не за "Артикул мережі").
# "АМ" і "НОВА АМ" - див. розділ довідника нижче.
# ============================================================

EPICENTR_DESC_COLS = [
    "Артикул", "Артикул мережі", "Бренд", "Назва",
    "Статус артикула", "Склад", "В дорозі", "Категорія",
]
EPICENTR_STORE_FIELDS = ["Сезон", "Не сезон", "АМ", "Залишок", "НОВА АМ"]

ZALYSHOK_HEADER_ROW = 4
ZALYSHOK_HEADER_COL = 4
ZALYSHOK_DATE_RE = re.compile(r"(\d{1,2})\.(\d{1,2})\.(\d{2,4})")


def extract_zalyshok_date(ws):
    """
    Дата з заголовка "Кінцевий залишок на <дата>" (рядок 4, колонка E -
    "якірна" клітинка об'єднаного діапазону). Повертає datetime.date або
    None, якщо в заголовку не вдалося розпізнати дату.
    """
    text = ws.cell_value(ZALYSHOK_HEADER_ROW, ZALYSHOK_HEADER_COL)
    if not text:
        return None
    m = ZALYSHOK_DATE_RE.search(str(text))
    if not m:
        return None
    d, mo, y = (int(x) for x in m.groups())
    if y < 100:
        y += 2000
    try:
        return date(y, mo, d)
    except ValueError:
        return None


def parse_epicentr_sales(wb):
    """
    Розбирає один файл "Звіт про продажу товарів" Епіцентру.
    Повертає (names, data, stores, zalyshok_date):
      names         - {артикул: назва}
      data          - {(артикул, магазин): (розхід, залишок)}
      stores        - множина кодів магазинів, що зустрілися
      zalyshok_date - дата "Кінцевий залишок на" з шапки файлу (або None)
    """
    ws = wb.sheet_by_index(0)
    names = {}
    data = {}
    current_store = None

    for r in range(ws.nrows):
        a = ws.cell_value(r, 0)
        b = ws.cell_value(r, 1)

        if isinstance(a, str) and b == "-":
            current_store = a.strip()
            continue

        if isinstance(a, str) and len(a) == 8 and a.isdigit() and b not in ("", None) and current_store:
            d = ws.cell_value(r, 3)
            e = ws.cell_value(r, 4)
            rozhid = d if isinstance(d, (int, float)) else None
            zalyshok = e if isinstance(e, (int, float)) else None
            names.setdefault(a, b)
            data[(a, current_store)] = (rozhid, zalyshok)

    stores = {store for (_, store) in data}
    zalyshok_date = extract_zalyshok_date(ws)
    return names, data, stores, zalyshok_date


def build_epicentr_report(
    season_file, offseason_file,
    reference_file=None, price_file=None, delivery_file=None, abc_files=None,
):
    """
    season_file / offseason_file - те саме, що src_file у build_report:
    шлях, файлоподібний об'єкт або bytes; формат (.xls/.xlsx) визначається
    автоматично. reference_file (необов'язковий) - файл-довідник попереднього
    періоду для заповнення "АМ". price_file (необов'язковий) - прайс-лист
    для "Артикул"/"Бренд"/"Статус артикула"/"Склад". delivery_file
    (необов'язковий) - графік поставок для "В дорозі" (потребує, щоб
    price_file вже дав "Артикул" - без нього "В дорозі" лишиться порожнім).
    abc_files (необов'язковий список) - файли АВС-аналізу продажів для
    "Категорія" (джойн за "Артикул", підтягнутим з прайс-листа, або, якщо
    не співпав, за точною "Назва").
    Повертає (openpyxl.Workbook, dict зі статистикою).
    """
    out_wb = openpyxl.Workbook()
    out_ws = out_wb.active
    out_ws.title = "Епіцентр"

    ref_wb = open_delivery_workbook(reference_file) if reference_file else None
    price_wb = open_delivery_workbook(price_file) if price_file else None
    delivery_wb = open_delivery_workbook(delivery_file) if delivery_file else None
    abc_by_article, abc_by_name, qty_by_article, qty_by_name = (
        parse_abc_files(abc_files) if abc_files else ({}, {}, {}, {})
    )
    b_cutoff = class_quantity_cutoff(abc_by_article, qty_by_article, "B", 0.5)
    stats = _write_epicentr_sheet(
        out_ws, season_file, offseason_file, ref_wb, price_wb, delivery_wb, abc_by_article, abc_by_name,
        qty_by_article=qty_by_article, qty_by_name=qty_by_name, b_cutoff=b_cutoff,
    )
    return out_wb, stats


def build_epicentr_combined_report(
    mt_season, mt_offseason, bsh_season, bsh_offseason,
    reference_file=None, price_file=None, delivery_file=None, abc_files=None,
    network_sales_file=None,
):
    """
    Один вихідний файл з двома листами - "МТ" (Мілленіум Трейд) і "БШ"
    (БебіШоп), кожен зі своєю парою файлів продажу (сезон/не сезон), плюс
    спільні необов'язкові файли: reference_file - файл-довідник попереднього
    періоду для "АМ" (шукає найкраще відповідний лист довідника окремо для
    кожного контрагента за перетином "Артикул мережі" - назви листів
    довідника при цьому не важливі); price_file - прайс-лист для
    "Артикул"/"Бренд"/"Статус артикула"/"Склад"; delivery_file - графік
    поставок для "В дорозі" (за щойно підтягнутим "Артикул"); abc_files -
    список файлів АВС-аналізу продажів для "Категорія" (спільний для обох
    листів - джойн за "Артикул", резервно за точною "Назва"); network_sales_file -
    файл "Мережі продажі" для АВС-аналізу магазинів за категорією "Іграшка":
    новий рядок над кодами магазинів на листах "МТ"/"БШ", і два листи
    Total_МТ/Total_БШ з ОДНАКОВИМ вмістом (повний список магазинів Епіцентру,
    знайдених і в "Мережі продажі", і хоча б в одному з чотирьох файлів
    продажу, з "Продажі", "Серед.продажі в міс." і АВС-класом).
    Повертає (openpyxl.Workbook, {"МТ": stats, "БШ": stats}).
    """
    ref_wb = open_delivery_workbook(reference_file) if reference_file else None
    price_wb = open_delivery_workbook(price_file) if price_file else None
    delivery_wb = open_delivery_workbook(delivery_file) if delivery_file else None
    abc_by_article, abc_by_name, qty_by_article, qty_by_name = (
        parse_abc_files(abc_files) if abc_files else ({}, {}, {}, {})
    )
    b_cutoff = class_quantity_cutoff(abc_by_article, qty_by_article, "B", 0.5)

    epicentr_stores = set()
    for f in (mt_season, mt_offseason, bsh_season, bsh_offseason):
        _, _, stores, _ = parse_epicentr_sales(open_delivery_workbook(f))
        epicentr_stores |= stores

    network_data = {}
    if network_sales_file:
        network_data = parse_network_sales_category(open_delivery_workbook(network_sales_file), "Іграшка")
    network_stores = {k: v for k, v in network_data.items() if k in epicentr_stores}

    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active)

    network_abc = classify_store_abc(network_stores) if network_stores else {}

    all_stats = {}
    for label, season_f, offseason_f in [
        ("МТ", mt_season, mt_offseason),
        ("БШ", bsh_season, bsh_offseason),
    ]:
        out_ws = out_wb.create_sheet(title=label)
        all_stats[label] = _write_epicentr_sheet(
            out_ws, season_f, offseason_f, ref_wb, price_wb, delivery_wb, abc_by_article, abc_by_name,
            network_abc, qty_by_article=qty_by_article, qty_by_name=qty_by_name, b_cutoff=b_cutoff,
        )

    if network_stores:
        for sheet_name in ("Total_МТ", "Total_БШ"):
            out_ws = out_wb.create_sheet(title=sheet_name)
            _write_store_totals_sheet(out_ws, network_stores, network_abc)

    return out_wb, all_stats


# --- Довідник попереднього періоду (тільки для "АМ" = "Нова АМ" довідника) ---
#
# Формат довідника НЕ фіксований - наступного разу файл трохи зміниться,
# тому все шукається за назвами заголовків (без урахування регістру), а не
# за номерами колонок/рядків:
#   - рядок заголовків шукається як той, де є клітинка "артикул мережі";
#   - у цьому рядку шукаються колонка "артикул мережі" (ключ) і всі колонки
#     "нова ам" (по одній на кожен магазин);
#   - код магазину для кожної колонки "нова ам" береться з об'єднаної
#     клітинки НАД рядком заголовків (там, де на практиці стоїть код на
#     кшталт "01", "CV", "CHB") - шукається через реальні межі об'єднання
#     (merged_cells), а не за фіксованим зсувом рядків, оскільки з
#     об'єднаних клітинок значення читається лише в "якірній" (лівій верхній).
# Довідник може містити декілька листів (по одному на контрагента) - який
# лист відповідає якому контрагенту, визначається автоматично за перетином
# множини "Артикул мережі" з товарами поточного зводу (не псується
# новинками, яких у довіднику ще нема - вони просто не додають перетину).
#
# "АМ" на новому періоді = "Нова АМ" зі старого довідника (не стара "АМ") -
# бо "Нова АМ" з минулого разу це і є вже прийняте й актуальне на зараз
# рішення по асортименту. Якщо "Нова АМ" в довіднику порожня для якогось
# товару/магазину - "АМ" лишається порожнім (без відкату до старої "АМ").
#
# "Артикул" і "Бренд" з довідника НЕ підтягуються (скасовано - буде інший
# підхід).

REFERENCE_KEY_HEADER = "артикул мережі"
REFERENCE_AM_HEADER = "нова ам"


def _norm_header(value):
    return value.strip().lower() if isinstance(value, str) else None


def _resolve_merge(merged_cells, row, col):
    """Якщо (row, col) входить у злиття - повертає координати його якірної клітинки."""
    for r1, r2, c1, c2 in merged_cells:
        if r1 <= row < r2 and c1 <= col < c2:
            return r1, c1
    return row, col


def _find_label_above(ws, merged_cells, header_row, col):
    """Шукає найближче непорожнє значення у стовпці col вище header_row,
    з урахуванням об'єднаних клітинок (для визначення коду магазину)."""
    for r in range(header_row - 1, -1, -1):
        anchor_r, anchor_c = _resolve_merge(merged_cells, r, col)
        v = ws.cell_value(anchor_r, anchor_c)
        if v not in ("", None):
            return str(v).strip()
    return None


def parse_epicentr_reference_sheet(ws):
    """
    Розбирає один лист довідника. Повертає am_map:
      {(артикул мережі, магазин): значення з колонки "Нова АМ" - воно
       стає "АМ" на новому періоді}
    Порожній словник, якщо на листі немає заголовка "артикул мережі".
    """
    header_row = None
    for r in range(ws.nrows):
        for c in range(ws.ncols):
            if _norm_header(ws.cell_value(r, c)) == REFERENCE_KEY_HEADER:
                header_row = r
                break
        if header_row is not None:
            break
    if header_row is None:
        return {}

    key_col = None
    am_cols = []
    for c in range(ws.ncols):
        h = _norm_header(ws.cell_value(header_row, c))
        if h == REFERENCE_KEY_HEADER and key_col is None:
            key_col = c
        elif h == REFERENCE_AM_HEADER:
            am_cols.append(c)

    if key_col is None:
        return {}

    merged_cells = getattr(ws, "merged_cells", [])
    store_by_col = {c: _find_label_above(ws, merged_cells, header_row, c) for c in am_cols}
    store_by_col = {c: s for c, s in store_by_col.items() if s}

    am_map = {}
    for r in range(header_row + 1, ws.nrows):
        key = ws.cell_value(r, key_col)
        if key in ("", None):
            continue
        key_norm = norm(key)
        for c, store in store_by_col.items():
            am_val = ws.cell_value(r, c)
            if isinstance(am_val, (int, float)):
                am_map[(key_norm, store)] = am_val

    return am_map


def find_best_reference_sheet(ref_wb, articles_needed):
    """
    Перебирає всі листи довідника, повертає (am_map, sheet_name, overlap)
    для листа з найбільшим перетином "Артикул мережі" з articles_needed.
    Якщо жоден лист не дав перетину - ({}, None, 0).
    """
    best_sheet, best_am_map, best_overlap = None, {}, 0
    for sheet_name in ref_wb.sheet_names():
        ws = ref_wb.sheet_by_name(sheet_name)
        am_map = parse_epicentr_reference_sheet(ws)
        keys_found = {k for k, _ in am_map}
        overlap = len(keys_found & articles_needed)
        if overlap > best_overlap:
            best_sheet, best_am_map, best_overlap = sheet_name, am_map, overlap
    return best_am_map, best_sheet, best_overlap


# --- Мережі продажі (АВС-аналіз магазинів) ---
#
# Файл - звіт 1С "Аналіз продаж по сетям товарів на реалізації": рядки
# згруповані по категорії / магазину / адресі, колонки - по місяцях
# (останній блок - "Итого"). Формат НЕ фіксований (може зміщуватись між
# вивантаженнями), тому все шукається за назвами заголовків, а не за
# номерами рядків/колонок:
#   - рядок заголовків - той, де в колонці A є "Магазин сети";
#   - у цьому рядку шукаються колонка "Адреса" і всі колонки "Сумма
#     продано" (по одній на місяць плюс одна підсумкова - визначається
#     за міткою в рядку над заголовками: "Итого" чи дата місяця);
#   - потрібна категорія шукається як об'єднана комірка в колонці A
#     (ширша за 2 колонки) з точним значенням її назви, десь нижче
#     заголовків; кінець блоку категорії - наступна така сама об'єднана
#     комірка (інша категорія або підсумковий "Итого" по всьому файлу);
#   - рядок одразу під заголовком категорії, якщо він без назви магазину,
#     пропускається як порожній рядок (загальне правило для будь-якого
#     рядка без назви) - на практиці там іноді лишаються "нічийні" суми
#     через те, що 1С під час вивантаження не розпізнала адресу окремого
#     магазину і приплюсувала їх сюди; якщо ж назва там колись таки буде -
#     рядок обробиться як звичайний магазин, а не пропуститься наосліп.
NETWORK_STORE_HEADER = "магазин сети"
NETWORK_ADDRESS_HEADER = "адреса"
NETWORK_SUM_HEADER = "сумма продано"
NETWORK_TOTAL_LABEL = "итого"


def _normalize_store_key(value):
    """Ключ магазину для звірки з іншими файлами: числа (код магазину на
    кшталт 1, 2, 3) доповнюються нулем зліва до 2 знаків, як "01", "02" -
    саме так вони записані в файлах продажу Епіцентру."""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    if isinstance(value, int):
        return f"{value:02d}"
    return str(value).strip()


def parse_network_sales_category(wb, category_name):
    """
    Розбирає файл "Мережі продажі" і повертає дані тільки по одній
    категорії (category_name, наприклад "Іграшка"):
      {ключ_магазину: {"name": назва як у файлі, "address": ...,
                        "total": "Сумма продано" з блоку "Итого",
                        "avg_month": total / кількість місяців, де
                        "Сумма продано" не порожнє (None, якщо таких
                        місяців нема)}}
    Порожній словник, якщо потрібна категорія або потрібні заголовки не
    знайдені.
    """
    ws = wb.sheet_by_index(0)
    merged_cells = ws.merged_cells

    header_row = None
    for r in range(ws.nrows):
        if _norm_header(ws.cell_value(r, 0)) == NETWORK_STORE_HEADER:
            header_row = r
            break
    if header_row is None:
        return {}

    address_col = None
    sum_cols = []  # (col, is_total)
    for c in range(ws.ncols):
        h = _norm_header(ws.cell_value(header_row, c))
        if h == NETWORK_ADDRESS_HEADER:
            address_col = c
        elif h == NETWORK_SUM_HEADER:
            anchor_r, anchor_c = _resolve_merge(merged_cells, header_row - 1, c)
            month_label = ws.cell_value(anchor_r, anchor_c)
            sum_cols.append((c, _norm_header(month_label) == NETWORK_TOTAL_LABEL))

    total_col = next((c for c, is_total in sum_cols if is_total), None)
    monthly_cols = [c for c, is_total in sum_cols if not is_total]
    if total_col is None:
        return {}

    category_row = None
    for r1, r2, c1, c2 in merged_cells:
        if c1 == 0 and (c2 - c1) > 2 and r1 > header_row:
            v = ws.cell_value(r1, 0)
            if isinstance(v, str) and v.strip() == category_name:
                category_row = r1
                break
    if category_row is None:
        return {}

    end_row = ws.nrows
    for r1, r2, c1, c2 in merged_cells:
        if c1 == 0 and (c2 - c1) > 2 and category_row < r1 < end_row:
            end_row = r1

    result = {}
    for r in range(category_row + 1, end_row):  # +1: пропускаємо сам рядок заголовка категорії
        name = ws.cell_value(r, 0)
        if name in ("", None):
            continue  # порожня назва - або "нічийний" рядок з артефактними сумами, або зайвий пробіл
        address = ws.cell_value(r, address_col) if address_col is not None else None
        total = ws.cell_value(r, total_col)
        total = total if isinstance(total, (int, float)) else 0
        active_months = sum(1 for c in monthly_cols if isinstance(ws.cell_value(r, c), (int, float)))
        store_key = _normalize_store_key(name)
        result[store_key] = {
            "name": store_key if isinstance(name, (int, float)) else name.strip(),
            "address": address if address not in ("", None) else None,
            "total": total,
            "avg_month": (total / active_months) if active_months else None,
        }
    return result


def classify_store_abc(store_data):
    """Класичний АВС-аналіз за "Серед.продажі в міс.": сортування за
    спаданням, потім кумулятивний поріг 70% / 20% / 10%. Магазини без
    жодного місяця продажів (avg_month=None) трактуються як 0 і йдуть в
    кінець. Повертає {ключ_магазину: "A"|"B"|"C"}."""
    items = sorted(store_data.items(), key=lambda kv: kv[1]["avg_month"] or 0, reverse=True)
    total = sum(v["avg_month"] or 0 for _, v in items)
    result, cum = {}, 0
    for key, info in items:
        cum += info["avg_month"] or 0
        share = (cum / total) if total else 1
        result[key] = "A" if share <= 0.7 else ("B" if share <= 0.9 else "C")
    return result


NETWORK_TOTALS_HEADERS = ["Магазин", "Адреса", "Продажі", "Серед.продажі в міс.", "АВС-клас"]


def _write_store_totals_sheet(out_ws, store_data, abc_result):
    """Пише лист Total (Антошка) / Total_МТ, Total_БШ (Епіцентр): магазини,
    відсортовані за спаданням "Серед.продажі в міс.", з їх АВС-класом."""
    for i, h in enumerate(NETWORK_TOTALS_HEADERS, start=1):
        out_ws.cell(row=1, column=i, value=h)
    ordered = sorted(store_data.items(), key=lambda kv: kv[1]["avg_month"] or 0, reverse=True)
    for r, (key, info) in enumerate(ordered, start=2):
        out_ws.cell(row=r, column=1, value=info["name"])
        out_ws.cell(row=r, column=2, value=info["address"])
        out_ws.cell(row=r, column=3, value=info["total"])
        out_ws.cell(row=r, column=4, value=info["avg_month"])
        out_ws.cell(row=r, column=5, value=abc_result.get(key))


# --- Дозаповнення листів Total після ручного коригування "НОВА АМ" ---
#
# Окремий, наступний крок: користувач завантажує вже готовий звід (той,
# що згенерував сам застосунок і в якому вручну доправив персикові й
# жовті клітинки "НОВА АМ" в Excel), і отримує той самий файл, де на
# листах Total/Total_МТ/Total_БШ дозаповнені лічильники по кожному
# магазину. Структура листа-матриці шукається за назвами заголовків
# (рядок з "НОВА АМ"), а не за фіксованими номерами рядків - так це не
# зламається, навіть якщо користувач вручну додав/прибрав рядок.

TOTAL_SHEET_PAIRS = [("Звід", "Total"), ("МТ", "Total_МТ"), ("БШ", "Total_БШ")]
TOTAL_NEW_HEADERS = [
    "A", "B", "C", "Новинки", "Без категорії", "Всього СКЮ",
    "Старий асортимент", "Залишок активного асортименту не в АМ",
]
TOTAL_PRICE_HEADERS = [
    "Всього скю прайс", "Всього скю прайс мережі",
    "% представленості прайс", "% представленості до акт ас з ПРАЙСУ",
]


def _empty_counts_bucket():
    return {
        "A": 0, "B": 0, "C": 0, "Новинки": 0, "Без категорії": 0,
        "Старий асортимент": 0, "Залишок не в АМ": 0,
    }


def _count_nova_am_by_class(ws):
    """
    Розбирає лист-матрицю (Звід / МТ / БШ) вже готового зводу. Повертає
    {назва_магазину: {...}} з лічильниками по КОЖНОМУ магазину матриці
    (навіть якщо в нього все по нулях - наприклад, магазин не знайшовся у
    "Мережі продажі" і тому не потрапив у лист Total на етапі побудови
    зводу; тут він все одно з'явиться, з нульовими лічильниками):
      "A"/"B"/"C"    - товари з "НОВА АМ" = 1 у цього магазину і класом
                       ("Категорія") A/B/C відповідно;
      "Новинки"      - товари з "НОВА АМ" = 1 у цього магазину і "Статус
                       артикула" = "NEW";
      "Без категорії" - товари з "НОВА АМ" = 1 у цього магазину, чий клас
                       не розпізнано як A/B/C (наприклад "Категорія" не
                       заповнена - товар не потрапив у жоден з файлів
                       АВС-аналізу) і це не новинка - без цього поля такі
                       товари взагалі не потрапляли б у жоден лічильник,
                       хоча людина вручну вирішила тримати їх в асортименті;
      "Старий асортимент" - товари зі "Статус артикула" = 8 (не входять у
                       жоден з лічильників вище), у яких "Залишок" у
                       цього магазину > 0;
      "Залишок не в АМ" - товари зі "Статус артикула" = 1, "НОВА АМ" = 0
                       і "Залишок" > 0 у цього магазину (актуальний товар,
                       якому вирішили не бути в новій АМ, але залишок ще
                       фізично є).
    "Залишок" для конкретного магазину береться з колонки одразу ліворуч
    від його "НОВА АМ" (в обох форматах - Антошка і Епіцентр - поля
    магазину завжди йдуть у порядку "... АМ, Залишок, НОВА АМ").
    """
    header_row = None
    for r in range(1, min(ws.max_row, 10) + 1):
        for c in range(1, ws.max_column + 1):
            if ws.cell(row=r, column=c).value == "НОВА АМ":
                header_row = r
                break
        if header_row is not None:
            break
    if header_row is None:
        return {}

    store_row = header_row - 1
    nova_am_cols = [c for c in range(1, ws.max_column + 1) if ws.cell(row=header_row, column=c).value == "НОВА АМ"]

    category_col = status_col = None
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if v == "Категорія" and category_col is None:
            category_col = c
        elif v == "Статус артикула" and status_col is None:
            status_col = c

    # Ініціалізуємо нулями кожен магазин матриці одразу - навіть якщо в нього
    # взагалі немає жодної клітинки, що щось інкрементує (наприклад "НОВА АМ"
    # скрізь порожня), він все одно має потрапити в результат.
    counts = {}
    for col in nova_am_cols:
        store = ws.cell(row=store_row, column=col).value
        if store is not None:
            counts.setdefault(store, _empty_counts_bucket())

    for r in range(header_row + 1, ws.max_row + 1):
        status = ws.cell(row=r, column=status_col).value if status_col else None
        status_is_8 = status == 8
        status_is_1 = status == 1
        category = ws.cell(row=r, column=category_col).value if category_col else None
        cls = _class_letter(category)
        is_new = isinstance(status, str) and status.strip().upper() == "NEW"

        for col in nova_am_cols:
            store = ws.cell(row=store_row, column=col).value
            if store is None:
                continue
            nova_am_val = ws.cell(row=r, column=col).value
            zal_val = ws.cell(row=r, column=col - 1).value
            has_stock = isinstance(zal_val, (int, float)) and zal_val > 0

            if status_is_8:
                if has_stock:
                    counts.setdefault(store, _empty_counts_bucket())["Старий асортимент"] += 1
                continue  # статус 8 не входить у жоден інший лічильник

            if nova_am_val == 1:
                bucket = counts.setdefault(store, _empty_counts_bucket())
                if cls in ("A", "B", "C"):
                    bucket[cls] += 1
                elif is_new:
                    bucket["Новинки"] += 1
                else:
                    bucket["Без категорії"] += 1
            elif status_is_1 and nova_am_val == 0 and has_stock:
                counts.setdefault(store, _empty_counts_bucket())["Залишок не в АМ"] += 1
    return counts


def _write_counts_row(total_ws, r, cols, c, price_sku_total, price_sku_network):
    """Пише лічильники c (з counts_by_store) у рядок r листа Total.
    Повертає "Всього СКЮ" цього рядка."""
    total = c["A"] + c["B"] + c["C"] + c["Новинки"] + c["Без категорії"]
    total_ws.cell(row=r, column=cols["A"], value=c["A"])
    total_ws.cell(row=r, column=cols["B"], value=c["B"])
    total_ws.cell(row=r, column=cols["C"], value=c["C"])
    total_ws.cell(row=r, column=cols["Новинки"], value=c["Новинки"])
    total_ws.cell(row=r, column=cols["Без категорії"], value=c["Без категорії"])
    total_ws.cell(row=r, column=cols["Всього СКЮ"], value=total)
    total_ws.cell(row=r, column=cols["Старий асортимент"], value=c["Старий асортимент"])
    total_ws.cell(
        row=r, column=cols["Залишок активного асортименту не в АМ"], value=c["Залишок не в АМ"]
    )

    if price_sku_total is not None or price_sku_network is not None:
        total_ws.cell(row=r, column=cols["Всього скю прайс"], value=price_sku_total)
        total_ws.cell(row=r, column=cols["Всього скю прайс мережі"], value=price_sku_network)

        pct1 = total_ws.cell(row=r, column=cols["% представленості прайс"])
        pct1.value = (total / price_sku_total) if price_sku_total else None
        pct1.number_format = "0.0%"

        pct2 = total_ws.cell(row=r, column=cols["% представленості до акт ас з ПРАЙСУ"])
        pct2.value = (total / price_sku_network) if price_sku_network else None
        pct2.number_format = "0.0%"

    return total


def _augment_total_sheet(total_ws, counts_by_store, price_totals=None):
    """Дописує в total_ws колонки TOTAL_NEW_HEADERS (або оновлює значення
    в них, якщо вони там вже є - повторний запуск не плодить дублі).
    price_totals (необов'язковий) - {"total": ..., "with_artikul_v_seti": ...}
    з count_toy_skus (сумарно по всіх брендах прайс-листа) - якщо переданий,
    додатково дописує TOTAL_PRICE_HEADERS: те саме число СКЮ прайс-листа в
    кожному рядку магазину (воно з прайс-листа, не з конкретного магазину),
    і два відсотки представленості від нього.
    Магазини, які є в counts_by_store (тобто реально є в листі-матриці), але
    яких немає серед наявних рядків total_ws (не знайшлися свого часу в
    "Мережі продажі" і тому не потрапили в лист Total при побудові зводу) -
    дописуються новими рядками в кінець, з порожніми "Адреса"/"Продажі"/
    "Серед.продажі в міс."/"АВС-клас" (даних звідти нема), але з нормально
    порахованими лічильниками "НОВА АМ" - вони ж все одно є в цьому магазині."""
    header_row = 1
    existing = {total_ws.cell(row=header_row, column=c).value: c for c in range(1, total_ws.max_column + 1)}
    store_col = existing.get("Магазин", 1)

    headers_to_ensure = list(TOTAL_NEW_HEADERS)
    if price_totals is not None:
        headers_to_ensure += TOTAL_PRICE_HEADERS

    next_col = total_ws.max_column + 1
    cols = {}
    for name in headers_to_ensure:
        if name in existing:
            cols[name] = existing[name]
        else:
            cols[name] = next_col
            total_ws.cell(row=header_row, column=next_col, value=name)
            next_col += 1

    price_sku_total = price_totals["total"] if price_totals else None
    price_sku_network = price_totals["with_artikul_v_seti"] if price_totals else None

    n_stores = 0
    n_sku_total = 0
    seen_stores = set()
    last_row = 1
    for r in range(2, total_ws.max_row + 1):
        store = total_ws.cell(row=r, column=store_col).value
        last_row = r
        if store is None:
            continue
        n_stores += 1
        seen_stores.add(store)
        c = counts_by_store.get(store, _empty_counts_bucket())
        n_sku_total += _write_counts_row(total_ws, r, cols, c, price_sku_total, price_sku_network)

    r = last_row + 1
    for store, c in counts_by_store.items():
        if store in seen_stores:
            continue
        total_ws.cell(row=r, column=store_col, value=store)
        n_stores += 1
        n_sku_total += _write_counts_row(total_ws, r, cols, c, price_sku_total, price_sku_network)
        r += 1

    result = {"n_stores": n_stores, "n_sku_total": n_sku_total}
    if price_totals is not None:
        result["price_sku_total"] = price_sku_total
        result["price_sku_network"] = price_sku_network
    return result


def fill_total_summary(file, price_file=None):
    """
    Приймає вже готовий звід (той, що згенерував build_report /
    build_epicentr_combined_report, і в якому користувач вручну доправив
    персикові й жовті клітинки "НОВА АМ" в Excel) і дозаповнює наявні
    листи Total/Total_МТ/Total_БШ колонками "A", "B", "C", "Новинки",
    "Всього СКЮ", "Старий асортимент", "Залишок активного асортименту не
    в АМ" (лічильники товарів по кожному магазину). Обробляються тільки ті
    пари (лист матриці, лист Total), які реально є у файлі - Антошка дає
    ("Звід", "Total"), Епіцентр дає ("МТ", "Total_МТ") і ("БШ", "Total_БШ");
    порожньо, якщо жодної пари не знайдено.
    price_file (необов'язковий) - прайс-лист будь-якої мережі (той самий
    формат, що бере count_toy_skus). Якщо переданий - додатково дописує
    "Всього скю прайс" / "Всього скю прайс мережі" (те саме число СКЮ
    прайс-листа - усього і з "Артикул в сети" - в кожному рядку кожного
    листа Total) і два відсотки представленості від "Всього СКЮ" цього
    магазину.
    Повертає (openpyxl.Workbook, {назва_листа_Total: {статистика}}).
    """
    wb = load_xlsx(read_bytes(file))

    price_totals = None
    if price_file is not None:
        by_brand = count_toy_skus(price_file)
        price_totals = {
            "total": sum(v["total"] for v in by_brand.values()),
            "with_artikul_v_seti": sum(v["with_artikul_v_seti"] for v in by_brand.values()),
        }

    stats = {}
    for matrix_name, total_name in TOTAL_SHEET_PAIRS:
        if matrix_name not in wb.sheetnames or total_name not in wb.sheetnames:
            continue
        counts_by_store = _count_nova_am_by_class(wb[matrix_name])
        stats[total_name] = _augment_total_sheet(wb[total_name], counts_by_store, price_totals)

    return wb, stats


def _write_epicentr_sheet(
    out_ws, season_file, offseason_file,
    ref_wb=None, price_wb=None, delivery_wb=None, abc_by_article=None, abc_by_name=None,
    network_abc=None, qty_by_article=None, qty_by_name=None, b_cutoff=None,
):
    """Пише один лист зводу Епіцентру в out_ws (уже створений). network_abc
    (необов'язковий) - {код_магазину: АВС-клас} з "Мережі продажі", для
    нового рядка над кодами магазинів і для автозаповнення "НОВА АМ".
    qty_by_article/qty_by_name/b_cutoff - те саме, що дає parse_abc_files і
    class_quantity_cutoff, теж для "НОВА АМ". Повертає stats."""
    network_abc = network_abc or {}
    qty_by_article = qty_by_article or {}
    qty_by_name = qty_by_name or {}
    season_wb = open_delivery_workbook(season_file)
    offseason_wb = open_delivery_workbook(offseason_file)

    season_names, season_data, season_stores, season_zdate = parse_epicentr_sales(season_wb)
    offseason_names, offseason_data, offseason_stores, offseason_zdate = parse_epicentr_sales(offseason_wb)

    # Залишок беремо з файлу, чия дата "Кінцевий залишок на" пізніша -
    # це не обов'язково файл "не сезон" (наприклад, якщо його завантажили
    # раніше і дата в шапці старіша, ніж у файлі "сезон").
    if season_zdate and offseason_zdate:
        zalyshok_data = offseason_data if offseason_zdate >= season_zdate else season_data
        zalyshok_period = "не сезон" if offseason_zdate >= season_zdate else "сезон"
        zalyshok_date = max(season_zdate, offseason_zdate)
    elif offseason_zdate:
        zalyshok_data, zalyshok_period, zalyshok_date = offseason_data, "не сезон", offseason_zdate
    elif season_zdate:
        zalyshok_data, zalyshok_period, zalyshok_date = season_data, "сезон", season_zdate
    else:
        zalyshok_data, zalyshok_period, zalyshok_date = offseason_data, "не сезон (за умовчанням)", None

    all_stores = sorted(season_stores | offseason_stores)
    all_names = dict(season_names)
    all_names.update(offseason_names)
    all_articles = sorted(all_names)

    am_map, ref_sheet_name, ref_overlap = {}, None, 0
    if ref_wb is not None:
        am_map, ref_sheet_name, ref_overlap = find_best_reference_sheet(ref_wb, set(all_articles))

    price_map = parse_price_list(price_wb, key_header=PRICE_NETWORK_KEY_HEADER) if price_wb is not None else {}
    delivery_data = parse_delivery_schedule(delivery_wb) if delivery_wb is not None else {}

    n_desc_cols = len(EPICENTR_DESC_COLS)

    # Рядок 1: АВС-клас магазину (з "Мережі продажі"), над кодом магазину
    for i, store in enumerate(all_stores):
        start_col = n_desc_cols + 1 + i * len(EPICENTR_STORE_FIELDS)
        for j in range(len(EPICENTR_STORE_FIELDS)):
            out_ws.cell(row=1, column=start_col + j, value=network_abc.get(_normalize_store_key(store)))

    for i, store in enumerate(all_stores):
        start_col = n_desc_cols + 1 + i * len(EPICENTR_STORE_FIELDS)
        for j in range(len(EPICENTR_STORE_FIELDS)):
            out_ws.cell(row=2, column=start_col + j, value=store)

    for i, name in enumerate(EPICENTR_DESC_COLS):
        out_ws.cell(row=3, column=i + 1, value=name)
    for i, store in enumerate(all_stores):
        start_col = n_desc_cols + 1 + i * len(EPICENTR_STORE_FIELDS)
        for j, field in enumerate(EPICENTR_STORE_FIELDS):
            out_ws.cell(row=3, column=start_col + j, value=field)

    out_ws.freeze_panes = out_ws.cell(row=4, column=n_desc_cols + 1)

    n_rows = 0
    price_matches = 0
    delivery_matches = 0
    existing_articles = set()
    contragent_brands = set()
    for r, art in enumerate(all_articles, start=4):
        p_art, p_brand, p_status, p_sklad = price_map.get(art, (None, None, None, None))
        vdorozi = delivery_data.get(p_art) if p_art else None
        if p_art is not None:
            price_matches += 1
            existing_articles.add(p_art)
        if p_brand is not None:
            contragent_brands.add(normalize_brand(p_brand))
        if vdorozi is not None:
            delivery_matches += 1

        out_ws.cell(row=r, column=1, value=p_art)  # Артикул (з прайсу, якщо є)
        out_ws.cell(row=r, column=2, value=art)  # Артикул мережі
        out_ws.cell(row=r, column=3, value=p_brand)  # Бренд (з прайсу, якщо є)
        out_ws.cell(row=r, column=4, value=all_names.get(art))  # Назва
        out_ws.cell(row=r, column=5, value=p_status)  # Статус артикула (з прайсу, якщо є)
        out_ws.cell(row=r, column=6, value=p_sklad)  # Склад (з прайсу, якщо є)
        out_ws.cell(row=r, column=7, value=vdorozi)  # В дорозі (з графіка поставок, за Артикул)
        category = lookup_category(p_art, all_names.get(art), abc_by_article or {}, abc_by_name or {})
        out_ws.cell(row=r, column=8, value=category)  # Категорія (з АВС-аналізу)

        status_is_8 = p_status == 8
        product_class = _class_letter(category)
        product_qty = lookup_abc_qty(p_art, all_names.get(art), qty_by_article, qty_by_name)

        for i, store in enumerate(all_stores):
            season_val = season_data.get((art, store))
            offseason_val = offseason_data.get((art, store))
            zalyshok_val = zalyshok_data.get((art, store))
            sezon_qty = season_val[0] if season_val else None
            ne_sezon_qty = offseason_val[0] if offseason_val else None
            zalyshok = zalyshok_val[1] if zalyshok_val else None
            am_val = am_map.get((art, store))

            start_col = n_desc_cols + 1 + i * len(EPICENTR_STORE_FIELDS)
            out_ws.cell(row=r, column=start_col, value=sezon_qty)
            out_ws.cell(row=r, column=start_col + 1, value=ne_sezon_qty)
            out_ws.cell(row=r, column=start_col + 2, value=am_val)
            out_ws.cell(row=r, column=start_col + 3, value=zalyshok)

            nova_am_cell = out_ws.cell(row=r, column=start_col + 4)
            if status_is_8:
                nova_am_cell.value = 0
            elif not _has_sales(sezon_qty, ne_sezon_qty):
                if _has_stock(zalyshok):
                    nova_am_cell.value = 0
                else:
                    nova_am_cell.fill = NOVA_AM_PEACH_FILL
            else:
                store_class = network_abc.get(_normalize_store_key(store))
                match = epicentr_assortment_match(store_class, product_class, product_qty, b_cutoff)
                if match is True:
                    nova_am_cell.value = 1
                else:
                    nova_am_cell.fill = NOVA_AM_YELLOW_FILL
        n_rows += 1

    # Кандидати в новинки - так само, як у Антошки: пошук ведеться на
    # графіку поставок (тільки лист Chicco, novelty_reliable=True), з
    # фільтром за брендами і артикулами, які вже підтягнуті прайс-листом
    # для цього конкретного зводу (МТ/БШ). Без графіка поставок або без
    # прайс-листа (тоді contragent_brands порожній) новинок не буде.
    novelty_candidates = []
    if delivery_wb is not None:
        novelty_candidates = find_novelty_candidates(delivery_wb, existing_articles, contragent_brands)

    name_col = EPICENTR_DESC_COLS.index("Назва") + 1
    art_col = EPICENTR_DESC_COLS.index("Артикул") + 1
    vdorozi_col = EPICENTR_DESC_COLS.index("В дорозі") + 1
    status_col = EPICENTR_DESC_COLS.index("Статус артикула") + 1
    r = 4 + n_rows
    for name, art, qty in novelty_candidates:
        out_ws.cell(row=r, column=name_col, value=name)
        out_ws.cell(row=r, column=art_col, value=art)
        out_ws.cell(row=r, column=vdorozi_col, value=qty)
        out_ws.cell(row=r, column=status_col, value="NEW")
        for i in range(len(all_stores)):
            start_col = n_desc_cols + 1 + i * len(EPICENTR_STORE_FIELDS)
            out_ws.cell(row=r, column=start_col + 4, value=1)
        r += 1
    last_row = r - 1

    # Заголовок "НОВА АМ" - світло-зелена заливка як орієнтир по колонці.
    # Заливка даних - за результатом розрахунку вище.
    nova_am_idx = EPICENTR_STORE_FIELDS.index("НОВА АМ")
    for i in range(len(all_stores)):
        start_col = n_desc_cols + 1 + i * len(EPICENTR_STORE_FIELDS)
        col = start_col + nova_am_idx
        out_ws.cell(row=3, column=col).fill = NOVA_AM_FILL

    return {
        "n_products": len(all_articles),
        "n_stores": len(all_stores),
        "n_rows": n_rows,
        "zalyshok_period": zalyshok_period,
        "zalyshok_date": zalyshok_date.isoformat() if zalyshok_date else None,
        "season_date": season_zdate.isoformat() if season_zdate else None,
        "offseason_date": offseason_zdate.isoformat() if offseason_zdate else None,
        "reference_sheet": ref_sheet_name,
        "reference_overlap": ref_overlap,
        "price_matches": price_matches,
        "delivery_matches": delivery_matches,
        "n_novelty": len(novelty_candidates),
    }


if __name__ == "__main__":
    main()
