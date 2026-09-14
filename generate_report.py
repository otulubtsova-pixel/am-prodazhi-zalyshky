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
import re
import zipfile
from datetime import date

import openpyxl
import xlrd
from openpyxl.styles import PatternFill

NOVA_AM_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

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
        # formatting_info нужен, чтобы получить merged_cells (для довідника Епіцентру - АМ).
        return xlrd.open_workbook(file_contents=data, formatting_info=True)
    return _XlsxAsXlrdBook(load_xlsx(data))


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


def parse_abc_analysis(wb):
    """
    Розбирає один файл АВС-аналізу продажів. Повертає (by_article, by_name):
      by_article - {артикул: АВС-клас}
      by_name    - {точна назва товару: АВС-клас} - для резервного джойну,
                   якщо артикул не співпаде (без жодних відхилень - точний
                   рядок)
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
        return {}, {}

    key_col = class_col = name_col = None
    for c in range(ws.ncols):
        h = _norm_header(ws.cell_value(header_row, c))
        if h == ABC_KEY_HEADER and key_col is None:
            key_col = c
        elif h == ABC_CLASS_HEADER and class_col is None:
            class_col = c
        elif h == ABC_NAME_HEADER and name_col is None:
            name_col = c

    if key_col is None or class_col is None:
        return {}, {}

    by_article = {}
    by_name = {}
    for r in range(header_row + 1, ws.nrows):
        art = ws.cell_value(r, key_col)
        if art in ("", None):
            continue  # підсумковий рядок класу (A/B/C), не товар
        cls = ws.cell_value(r, class_col)
        if cls in ("", None):
            continue
        by_article[norm(art)] = str(cls).strip()
        if name_col is not None:
            name = ws.cell_value(r, name_col)
            if name not in ("", None):
                by_name[str(name).strip()] = str(cls).strip()

    return by_article, by_name


def parse_abc_files(files):
    """
    files - список файлів (шлях/файлоподібний об'єкт/bytes), кожен - один
    файл АВС-аналізу. Об'єднує результати всіх файлів в один (by_article,
    by_name); при дублікатах пізніший файл у списку переважає.
    """
    by_article = {}
    by_name = {}
    for f in files:
        wb = open_delivery_workbook(f)
        a, n = parse_abc_analysis(wb)
        by_article.update(a)
        by_name.update(n)
    return by_article, by_name


def lookup_category(art, name, abc_by_article, abc_by_name):
    """Категорія за артикулом, або (якщо не знайдено) за точною назвою."""
    if art is not None:
        cat = abc_by_article.get(art)
        if cat is not None:
            return cat
    if name is not None:
        return abc_by_name.get(str(name).strip())
    return None


def build_report(src_file, delivery_file, price_file=None, abc_files=None):
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
        wb = load_xlsx(src_bytes)
    else:
        wb = _XlsAsOpenpyxlWorkbook(xlrd.open_workbook(file_contents=src_bytes))

    am_store_cols, am_desc, am_data = parse_am_sheet(wb["АМ"])
    pr_store_periods, pr_desc, pr_data = parse_sales_sheet(wb["Продажі 2025"])

    delivery_wb = open_delivery_workbook(delivery_file)
    delivery_data = parse_delivery_schedule(delivery_wb)

    price_data = parse_price_list(open_delivery_workbook(price_file)) if price_file else {}
    abc_by_article, abc_by_name = parse_abc_files(abc_files) if abc_files else ({}, {})

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
        _, _, price_status, price_sklad = price_data.get(artikul, (None, None, None, None))
        category = lookup_category(artikul, desc.get("Наименование товара"), abc_by_article, abc_by_name)
        out_ws.cell(row=r, column=len(DESC_COLS) + 1, value=delivery_qty)
        out_ws.cell(row=r, column=len(DESC_COLS) + 2, value=price_status)
        out_ws.cell(row=r, column=len(DESC_COLS) + 3, value=price_sklad)
        out_ws.cell(row=r, column=len(DESC_COLS) + 4, value=category)

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
    abc_by_article, abc_by_name = parse_abc_files(abc_files) if abc_files else ({}, {})
    stats = _write_epicentr_sheet(
        out_ws, season_file, offseason_file, ref_wb, price_wb, delivery_wb, abc_by_article, abc_by_name
    )
    return out_wb, stats


def build_epicentr_combined_report(
    mt_season, mt_offseason, bsh_season, bsh_offseason,
    reference_file=None, price_file=None, delivery_file=None, abc_files=None,
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
    листів - джойн за "Артикул", резервно за точною "Назва").
    Повертає (openpyxl.Workbook, {"МТ": stats, "БШ": stats}).
    """
    ref_wb = open_delivery_workbook(reference_file) if reference_file else None
    price_wb = open_delivery_workbook(price_file) if price_file else None
    delivery_wb = open_delivery_workbook(delivery_file) if delivery_file else None
    abc_by_article, abc_by_name = parse_abc_files(abc_files) if abc_files else ({}, {})

    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active)

    all_stats = {}
    for label, season_f, offseason_f in [
        ("МТ", mt_season, mt_offseason),
        ("БШ", bsh_season, bsh_offseason),
    ]:
        out_ws = out_wb.create_sheet(title=label)
        all_stats[label] = _write_epicentr_sheet(
            out_ws, season_f, offseason_f, ref_wb, price_wb, delivery_wb, abc_by_article, abc_by_name
        )

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


def _write_epicentr_sheet(
    out_ws, season_file, offseason_file,
    ref_wb=None, price_wb=None, delivery_wb=None, abc_by_article=None, abc_by_name=None,
):
    """Пише один лист зводу Епіцентру в out_ws (уже створений). Повертає stats."""
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

    for i, store in enumerate(all_stores):
        start_col = n_desc_cols + 1 + i * len(EPICENTR_STORE_FIELDS)
        for j in range(len(EPICENTR_STORE_FIELDS)):
            out_ws.cell(row=1, column=start_col + j, value=store)

    for i, name in enumerate(EPICENTR_DESC_COLS):
        out_ws.cell(row=2, column=i + 1, value=name)
    for i, store in enumerate(all_stores):
        start_col = n_desc_cols + 1 + i * len(EPICENTR_STORE_FIELDS)
        for j, field in enumerate(EPICENTR_STORE_FIELDS):
            out_ws.cell(row=2, column=start_col + j, value=field)

    out_ws.freeze_panes = out_ws.cell(row=3, column=n_desc_cols + 1)

    n_rows = 0
    price_matches = 0
    delivery_matches = 0
    existing_articles = set()
    contragent_brands = set()
    for r, art in enumerate(all_articles, start=3):
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
            out_ws.cell(row=r, column=start_col + 4, value=None)
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
    r = 3 + n_rows
    for name, art, qty in novelty_candidates:
        out_ws.cell(row=r, column=name_col, value=name)
        out_ws.cell(row=r, column=art_col, value=art)
        out_ws.cell(row=r, column=vdorozi_col, value=qty)
        r += 1
    last_row = r - 1

    nova_am_idx = EPICENTR_STORE_FIELDS.index("НОВА АМ")
    for i in range(len(all_stores)):
        start_col = n_desc_cols + 1 + i * len(EPICENTR_STORE_FIELDS)
        col = start_col + nova_am_idx
        out_ws.cell(row=2, column=col).fill = NOVA_AM_FILL
        for row in range(3, last_row + 1):
            out_ws.cell(row=row, column=col).fill = NOVA_AM_FILL

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
