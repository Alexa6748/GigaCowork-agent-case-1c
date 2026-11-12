#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Распределение новых ФУЛ между логистами (навык GigaCowork «ful-distribution»).

Вход: выгрузка ФУЛ из 1С, реестр Cloud, список сотрудников из 1С (.xlsx / .xlsm / .xls / .csv).
Порядок файлов и имена листов не важны: тип файла и нужные колонки определяются по заголовкам.

Выход (в папку --out):
  Распределение_ФУЛ_<дата>.xlsx — новые строки Cloud, повторные отправки, возможные пропуски,
                                   «Требует внимания», журнал решений, загрузка логистов, сводка;
  Распределение_ФУЛ_<дата>.json — то же в машиночитаемом виде;
  stdout                         — готовый короткий ответ для чата.

Зависимости: стандартная библиотека Python 3, openpyxl, xlrd (только для .xls). Сеть не используется.

Запуск:
  python3 distribute.py --inputs <файл или папка> [<файл или папка> ...] --out <папка результата>
"""
import argparse
import csv
import json
import os
import re
import sys
import time
from collections import Counter, OrderedDict
from datetime import date, datetime, timedelta

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_CONFIG = os.path.join(SKILL_DIR, "references", "config.json")
OUTPUT_MARK_SHEET = "Журнал решений"   # по этому листу узнаём собственный результат и не читаем его как вход

KIND_NEW = "Новый"
KIND_REPEAT = "Повторная отправка"
KIND_DUP = "Дубль в пачке"
KIND_DONE = "Уже учтён в Cloud"
KIND_MISS = "Возможный пропуск"
KIND_SKIP = "Не обработан"


# --------------------------------------------------------------------------- утилиты

def norm_text(value):
    """Нормализация для сравнения: регистр, ё→е, знаки препинания и пробелы."""
    if value is None:
        return ""
    s = str(value).replace("\xa0", " ").lower().replace("ё", "е")
    s = re.sub(r"[\W_]+", " ", s)
    return " ".join(s.split())


def clean(value):
    """Пустые строки → None, строки без лишних пробелов."""
    if value is None:
        return None
    if isinstance(value, str):
        s = " ".join(value.replace("\xa0", " ").split())
        return s or None
    return value


def as_str(value):
    v = clean(value)
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    if isinstance(v, datetime):
        return v.strftime("%d.%m.%Y %H:%M")
    if isinstance(v, date):
        return v.strftime("%d.%m.%Y")
    return str(v)


DATE_FORMATS = ["%d.%m.%Y %H:%M:%S", "%d.%m.%Y %H:%M", "%d.%m.%Y", "%d.%m.%y",
                "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d", "%Y-%m-%dT%H:%M:%S",
                "%d/%m/%Y %H:%M", "%d/%m/%Y"]


def parse_dt(value):
    v = clean(value)
    if v is None:
        return None
    if isinstance(v, datetime):
        return v
    if isinstance(v, date):
        return datetime(v.year, v.month, v.day)
    if isinstance(v, (int, float)) and 20000 < float(v) < 80000:   # серийная дата Excel
        return datetime(1899, 12, 30) + timedelta(days=float(v))
    s = str(v).strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def fmt_date(d):
    return d.strftime("%d.%m.%Y") if d else None


# --------------------------------------------------------------------------- чтение файлов

def read_workbook(path):
    """Возвращает список (имя листа, строки) для файла любого поддерживаемого формата."""
    ext = os.path.splitext(path)[1].lower()
    sheets = []
    if ext in (".xlsx", ".xlsm"):
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        for ws in wb.worksheets:
            rows = [list(r) for r in ws.iter_rows(values_only=True)]
            sheets.append((ws.title, rows))
        wb.close()
    elif ext == ".xls":
        import xlrd
        book = xlrd.open_workbook(path)
        for sh in book.sheets():
            rows = []
            for r in range(sh.nrows):
                row = []
                for c in range(sh.ncols):
                    cell = sh.cell(r, c)
                    if cell.ctype == xlrd.XL_CELL_DATE:
                        row.append(xlrd.xldate.xldate_as_datetime(cell.value, book.datemode))
                    elif cell.ctype == xlrd.XL_CELL_EMPTY:
                        row.append(None)
                    elif cell.ctype == xlrd.XL_CELL_NUMBER and float(cell.value).is_integer():
                        row.append(int(cell.value))
                    else:
                        row.append(cell.value)
                rows.append(row)
            sheets.append((sh.name, rows))
    elif ext in (".csv", ".txt"):
        raw = None
        for enc in ("utf-8-sig", "cp1251"):
            try:
                with open(path, encoding=enc) as fh:
                    raw = fh.read()
                break
            except UnicodeDecodeError:
                continue
        if raw is None:
            return []
        try:
            dialect = csv.Sniffer().sniff(raw[:5000], delimiters=";,\t")
        except csv.Error:
            dialect = csv.excel
            dialect.delimiter = ";"
        rows = [list(r) for r in csv.reader(raw.splitlines(), dialect)]
        sheets.append((os.path.basename(path), rows))
    return sheets


def is_own_output(sheets):
    return any(name == OUTPUT_MARK_SHEET for name, _ in sheets)


def build_specs(cfg):
    specs = {}
    for kind, spec in cfg["headers"].items():
        cols = OrderedDict()
        for key, syns in spec["columns"].items():
            cols[key] = {norm_text(s) for s in syns}
        specs[kind] = {"required": spec["required"], "columns": cols}
    return specs


def find_header(rows, spec, scan_rows=30):
    """Ищет строку заголовков: все обязательные колонки + максимум найденных."""
    best = None
    for i, row in enumerate(rows[:scan_rows]):
        mapping = {}
        for j, cell in enumerate(row):
            h = norm_text(cell)
            if not h:
                continue
            for key, syns in spec["columns"].items():
                if key not in mapping and h in syns:
                    mapping[key] = (j, str(cell).strip())
                    break
        if all(k in mapping for k in spec["required"]):
            if best is None or len(mapping) > len(best[1]):
                best = (i, mapping)
    return best


def extract_records(rows, header_idx, mapping):
    recs = []
    for r_i in range(header_idx + 1, len(rows)):
        row = rows[r_i]
        rec = {}
        for key, (j, _) in mapping.items():
            rec[key] = clean(row[j]) if j < len(row) else None
        if all(v is None for v in rec.values()):
            continue
        rec["_row"] = r_i + 1          # номер строки в исходном листе (как в Excel)
        recs.append(rec)
    return recs


def collect_inputs(paths):
    files = []
    for p in paths:
        if os.path.isdir(p):
            for name in sorted(os.listdir(p)):
                full = os.path.join(p, name)
                if os.path.isfile(full) and not name.startswith("~$") and \
                        os.path.splitext(name)[1].lower() in (".xlsx", ".xlsm", ".xls", ".csv"):
                    files.append(full)
        elif os.path.isfile(p):
            files.append(p)
    seen, out = set(), []
    for f in files:
        a = os.path.abspath(f)
        if a not in seen:
            seen.add(a)
            out.append(f)
    return out


def detect_tables(files, specs):
    """Для каждого типа (ful / cloud / staff) выбирает лист с лучшим совпадением заголовков."""
    found = {k: [] for k in specs}
    skipped = []
    for path in files:
        try:
            sheets = read_workbook(path)
        except Exception as exc:  # noqa: BLE001 — сообщаем пользователю, а не падаем
            skipped.append("%s: не удалось прочитать (%s)" % (os.path.basename(path), exc))
            continue
        if is_own_output(sheets):
            skipped.append("%s: это результат прошлого запуска агента, пропущен" % os.path.basename(path))
            continue
        for sheet_name, rows in sheets:
            for kind, spec in specs.items():
                hdr = find_header(rows, spec)
                if hdr:
                    recs = extract_records(rows, hdr[0], hdr[1])
                    found[kind].append({"file": path, "sheet": sheet_name, "header_row": hdr[0] + 1,
                                        "mapping": hdr[1], "records": recs, "score": len(hdr[1])})
    chosen, problems = {}, []
    for kind, cands in found.items():
        if not cands:
            continue
        cands.sort(key=lambda c: (-c["score"], -len(c["records"]), c["file"], c["sheet"]))
        chosen[kind] = cands[0]
        if len(cands) > 1:
            problems.append("Для «%s» подошло несколько листов, взят %s / %s"
                            % (KIND_TITLES[kind], os.path.basename(cands[0]["file"]), cands[0]["sheet"]))
    return chosen, problems, skipped


KIND_TITLES = {"ful": "выгрузка ФУЛ из 1С", "cloud": "реестр Cloud", "staff": "сотрудники из 1С"}


# --------------------------------------------------------------------------- справочники и сопоставление

class Rules:
    def __init__(self, cfg):
        self.cfg = cfg
        self.undef = cfg["undefined"]
        self.groups = cfg["groups"]
        self.method_map = {}
        for canon, syns in cfg["methods"].items():
            for s in syns + [canon]:
                self.method_map[norm_text(s)] = canon
                self.method_map[norm_text(s).replace(" ", "")] = canon
        self.air = set(cfg["air_methods"])
        self.china = {norm_text(c) for c in cfg["china_countries"]} | \
                     {norm_text(c).replace(" ", "") for c in cfg["china_countries"]}
        self.order_re = re.compile(cfg["order_regex"])
        self.legal = {norm_text(t) for t in cfg["supplier_legal_tokens"]} | set(cfg["supplier_legal_tokens"])
        self.in_work = {norm_text(v) for v in cfg["cloud_in_work_values"]}
        self.skip_status = [norm_text(k) for k in cfg["ful_skip_status_keywords"]]

    # --- способ доставки, страна, группа
    def method(self, raw):
        if raw is None:
            return None
        n = norm_text(raw)
        return self.method_map.get(n) or self.method_map.get(n.replace(" ", ""))

    def is_china(self, raw):
        if raw is None:
            return None
        n = norm_text(raw)
        return n in self.china or n.replace(" ", "") in self.china

    def group(self, method_raw, method, country_raw):
        """Возвращает (ключ группы или None, объяснение)."""
        if method_raw is None:
            return None, "не указан способ доставки"
        if method is None:
            return None, "неизвестный способ доставки «%s»" % method_raw
        if method in self.air:
            return "air", "способ доставки %s → авиаперевозки (страна не влияет)" % method
        if country_raw is None:
            return None, "способ доставки %s (не авиа), но не указана страна отправления — без неё группу не выбрать" % method
        if self.is_china(country_raw):
            return "ved", "способ доставки %s (не авиа), страна «%s» — Китай → ВЭД" % (method, country_raw)
        return "multi", "способ доставки %s (не авиа), страна «%s» — не Китай → мультимодальные перевозки" % (method, country_raw)

    # --- номер заказа
    def orders_in(self, text):
        if text is None:
            return [], False
        found, cyr = [], False
        for letter, digits in self.order_re.findall(str(text)):
            if letter in "Сс":
                cyr = True
            o = "S" + digits
            if o not in found:
                found.append(o)
        return found, cyr

    # --- поставщик
    def supplier_key(self, raw):
        if raw is None:
            return None
        tokens = norm_text(raw).split()
        # хвост из однобуквенных токенов вида «S.r.l.», «A.Ş.», «Co.»
        k = len(tokens)
        while k > 0 and len(tokens[k - 1]) == 1:
            k -= 1
        if k < len(tokens) and "".join(tokens[k:]) in self.legal:
            tokens = tokens[:k]
        tokens = [t for t in tokens if t not in self.legal]
        key = "".join(tokens)
        return key or norm_text(raw).replace(" ", "")


class Employee:
    def __init__(self, rec, rules):
        cfg = rules.cfg
        self.fio = as_str(rec.get("name")) or ""
        self.state = as_str(rec.get("state")) or ""
        self.position = as_str(rec.get("position")) or ""
        self.code = as_str(rec.get("code")) or ""
        parts = self.fio.split()
        self.surname = parts[0] if parts else ""
        self.first = parts[1] if len(parts) > 1 else ""
        self.patr = parts[2] if len(parts) > 2 else ""
        initials = "".join(p[0].upper() + "." for p in (self.first, self.patr) if p)
        self.short = (self.surname + " " + initials).strip()
        pos = norm_text(self.position)
        self.excluded = any(norm_text(k) in pos for k in cfg["excluded_position_keywords"])
        self.group = None
        for gkey, g in rules.groups.items():
            if any(norm_text(k) in pos for k in g["position_keywords"]):
                self.group = gkey
                break
        st = norm_text(self.state)
        states = cfg["staff_states"]
        if st in {norm_text(s) for s in states["available"]}:
            self.state_kind = "works"
        elif any(norm_text(k) in st for k in states["dismissed_keywords"]):
            self.state_kind = "dismissed"
        elif any(norm_text(k) in st for k in states["leave_keywords"]):
            self.state_kind = "leave"
        else:
            self.state_kind = "unknown"
        self.available = self.state_kind == "works"
        self.distributes = self.group is not None and not self.excluded

    def sort_key(self):
        return (norm_text(self.fio), self.code)


class NameMatcher:
    """Сопоставляет короткие имена из Cloud («Кузнецова М.», «Федоров П.И.») с ФИО из 1С."""

    def __init__(self, staff):
        self.staff = staff
        self.cache = {}

    def match(self, raw):
        if raw is None:
            return None, "пусто"
        if raw in self.cache:
            return self.cache[raw]
        tokens = norm_text(raw).split()
        cands = []
        for e in self.staff:
            sn = norm_text(e.surname)
            if not sn or sn not in tokens:
                continue
            rest = list(tokens)
            rest.remove(sn)
            parts = [norm_text(e.first), norm_text(e.patr)]
            ok = len(rest) <= 2
            for i, t in enumerate(rest):
                p = parts[i] if i < len(parts) else ""
                if not p or (len(t) == 1 and p[0] != t) or (len(t) > 1 and p != t):
                    ok = False
                    break
            if ok:
                cands.append(e)
        if len(cands) == 1:
            res = (cands[0], "ok")
        elif cands:
            res = (None, "неоднозначно: " + ", ".join(c.fio for c in cands))
        else:
            res = (None, "нет в списке сотрудников")
        self.cache[raw] = res
        return res


# --------------------------------------------------------------------------- основной алгоритм

class Ful:
    pass


def run(inputs, out_dir, config_path):
    t0 = time.time()
    with open(config_path, encoding="utf-8") as fh:
        cfg = json.load(fh)
    rules = Rules(cfg)
    U = rules.undef
    specs = build_specs(cfg)

    files = collect_inputs(inputs)
    if not files:
        return fail("не найдено ни одного файла Excel/CSV среди приложенных.")
    tables, detect_problems, skipped = detect_tables(files, specs)
    missing = [KIND_TITLES[k] for k in ("ful", "cloud", "staff") if k not in tables]
    if missing:
        return fail("не удалось найти: %s. Проверьте, что приложены выгрузка ФУЛ из 1С, реестр Cloud и "
                    "список сотрудников из 1С, и что в них есть строка заголовков." % ", ".join(missing),
                    files, skipped)

    general_attention = []   # замечания, не привязанные к конкретному ФУЛ
    for p in detect_problems:
        general_attention.append({"what": p, "action": "Использован лист с лучшим совпадением заголовков",
                                  "decide": "Проверить, что взят нужный файл"})

    # --- сотрудники
    staff = [Employee(r, rules) for r in tables["staff"]["records"] if clean(r.get("name"))]
    matcher = NameMatcher(staff)

    # --- реестр Cloud
    cloud_rows = []
    cloud_by_order = {}
    supplier_spelling = {}
    unmatched_cloud_names = Counter()
    for r in tables["cloud"]["records"]:
        order_raw = as_str(r.get("order"))
        if order_raw is None:
            continue
        found, _ = rules.orders_in(order_raw)
        order = found[0] if len(found) == 1 else order_raw.upper()
        logist_raw = as_str(r.get("logist"))
        emp, why = matcher.match(logist_raw) if logist_raw else (None, "пусто")
        if logist_raw and emp is None:
            unmatched_cloud_names[(logist_raw, why)] += 1
        sup_raw = as_str(r.get("supplier"))
        skey = rules.supplier_key(sup_raw)
        if skey:
            supplier_spelling.setdefault(skey, Counter())[sup_raw] += 1
        row = {"order": order, "supplier": sup_raw, "skey": skey, "logist": logist_raw, "emp": emp,
               "in_work": norm_text(r.get("status")) in rules.in_work,
               "ul_date": parse_dt(r.get("ul_date")), "comment": as_str(r.get("comment")) or "",
               "_row": r["_row"]}
        cloud_rows.append(row)
        cloud_by_order.setdefault(order, []).append(row)
    ul_dates = [c["ul_date"].date() for c in cloud_rows if c["ul_date"]]
    last_cloud_date = max(ul_dates) if ul_dates else None
    if last_cloud_date is None:
        general_attention.append({"what": "В реестре Cloud нет дат получения УЛ",
                                  "action": "Проверка возможных пропусков не выполнялась",
                                  "decide": "Проверить колонку «Дата получения УЛ» в реестре"})
    for (name, why), n in sorted(unmatched_cloud_names.items()):
        if why.startswith("неоднозначно"):
            general_attention.append({"what": "Логист «%s» в Cloud (%d стр.) — %s" % (name, n, why),
                                      "action": "Эти строки не учтены ни в загрузке, ни в истории",
                                      "decide": "Уточнить, кто это, и записать в Cloud однозначно"})

    load = Counter()       # «в работе» по сотруднику (id)
    history = {}           # skey → Counter(id сотрудника)
    hist_unmatched = {}    # skey → Counter(имя из Cloud) для пояснений
    for c in cloud_rows:
        if c["emp"] is not None:
            if c["in_work"]:
                load[id(c["emp"])] += 1
            if c["skey"]:
                history.setdefault(c["skey"], Counter())[id(c["emp"])] += 1
        elif c["logist"] and c["skey"]:
            hist_unmatched.setdefault(c["skey"], Counter())[c["logist"]] += 1
    load_initial = Counter(load)

    # --- ФУЛ
    fuls = []
    for r in tables["ful"]["records"]:
        f = Ful()
        f.row = r["_row"]
        f.number = as_str(r.get("number")) or U
        f.dt = parse_dt(r.get("date"))
        f.date_raw = as_str(r.get("date"))
        f.cutoff_dt = parse_dt(r.get("cutoff"))
        f.cutoff_raw = as_str(r.get("cutoff"))
        f.status = as_str(r.get("status"))
        f.comment = as_str(r.get("comment"))
        f.orders, f.cyr = rules.orders_in(f.comment)
        f.order = f.orders[0] if len(f.orders) == 1 else None
        f.method_raw = as_str(r.get("method"))
        f.method = rules.method(f.method_raw)
        f.brand = as_str(r.get("brand"))
        f.logist_1c = as_str(r.get("logist"))
        f.supplier_raw = as_str(r.get("supplier"))
        f.skey = rules.supplier_key(f.supplier_raw)
        f.country = as_str(r.get("country"))
        f.kind = None
        f.group = None
        f.group_why = ""
        f.emp = None
        f.reason = []
        f.notes = []
        f.attention = []        # (что, что сделал агент, что решить)
        f.cloud_comment = ""
        f.candidates = ""
        f.primary = None
        f.dups = []
        fuls.append(f)

    def sort_key(f):
        return (f.dt or datetime.max, f.number, f.row)

    # 1) ФУЛ, которые нельзя обработать
    for f in fuls:
        st = norm_text(f.status)
        if f.status and any(k in st for k in rules.skip_status):
            f.kind = KIND_SKIP
            f.reason.append("Статус ФУЛ в 1С «%s» — документ не обрабатывается" % f.status)
            f.attention.append(("ФУЛ со статусом «%s»" % f.status, "Строка не создана",
                                "Подтвердить, что ФУЛ действительно отменён"))
        elif len(f.orders) > 1:
            f.kind = KIND_SKIP
            f.reason.append("В комментарии несколько номеров заказа (%s) — заказ не определён" % ", ".join(f.orders))
            f.attention.append(("Несколько номеров заказа в комментарии: %s" % ", ".join(f.orders),
                                "Строка не создана, заказ: %s" % U, "Указать, к какому заказу относится ФУЛ"))
        elif f.order is None:
            f.kind = KIND_SKIP
            f.reason.append("В комментарии нет номера заказа (ожидается «Номер: S0001234 от ДД.ММ.ГГГГ»), "
                            "комментарий: «%s»" % (f.comment or ""))
            f.attention.append(("Нет номера заказа", "Строка не создана, заказ: %s" % U,
                                "Найти номер заказа и внести вручную"))
        if f.cyr and f.order:
            f.notes.append("номер заказа записан с кириллической «С», приведён к %s" % f.order)

    # 2) повтор / новый / дубль / пропуск
    by_order = OrderedDict()
    for f in sorted([x for x in fuls if x.kind is None], key=sort_key):
        by_order.setdefault(f.order, []).append(f)

    repeats_rows = []
    cloud_comment_acc = {}
    for order, group_fuls in by_order.items():
        if order in cloud_by_order:
            crow = cloud_by_order[order][0]
            if len(cloud_by_order[order]) > 1:
                general_attention.append({"what": "Заказ %s записан в Cloud %d раза" % (order, len(cloud_by_order[order])),
                                          "action": "Для сверки взята первая строка",
                                          "decide": "Убрать дубль в реестре Cloud"})
            acc = cloud_comment_acc.setdefault(order, crow["comment"])
            for f in group_fuls:
                d = fmt_date(f.dt)
                text = cfg["repeat_comment_template"].format(date=d or U)
                if f.dt and crow["ul_date"] and f.dt.date() == crow["ul_date"].date():
                    f.kind = KIND_DONE
                    f.reason.append("Заказ %s уже есть в Cloud, дата получения УЛ в Cloud %s совпадает с датой ФУЛ — "
                                    "это исходный ФУЛ, уже перенесённый" % (order, fmt_date(crow["ul_date"])))
                    continue
                if d and norm_text(text) in norm_text(crow["comment"]):
                    f.kind = KIND_DONE
                    f.reason.append("Заказ %s уже есть в Cloud, и в комментарии уже отмечено «%s»" % (order, text))
                    continue
                f.kind = KIND_REPEAT
                f.cloud_comment = text
                f.reason.append("Заказ %s уже есть в Cloud (строка %d, логист %s) — новая строка не создаётся, "
                                "в комментарий дописывается «%s»" % (order, crow["_row"], crow["logist"] or "не указан", text))
                if not f.dt:
                    f.attention.append(("Повторная отправка без даты ФУЛ", "В комментарий записано «%s»" % text,
                                        "Указать дату повторной отправки"))
                if f.skey and crow["skey"] and f.skey != crow["skey"]:
                    f.attention.append(("Поставщик в 1С «%s» не совпадает с Cloud «%s»" % (f.supplier_raw, crow["supplier"]),
                                        "Отмечено как повторная отправка по номеру заказа",
                                        "Проверить номер заказа и поставщика"))
                before = acc
                acc = (acc + "; " + text) if acc else text
                repeats_rows.append({"f": f, "type": "заказ уже в Cloud", "logist": crow["logist"] or "",
                                     "current": before, "append": text, "after": acc, "order": order})
            cloud_comment_acc[order] = acc
        else:
            primary = group_fuls[0]
            if primary.dt is None:
                primary.kind = KIND_NEW
                primary.attention.append(("Нет даты ФУЛ", "Строка создана, дата получения УЛ: %s" % U,
                                          "Указать дату ФУЛ"))
            elif last_cloud_date and primary.dt.date() < last_cloud_date:
                primary.kind = KIND_MISS
                primary.reason.append("Заказа %s нет в Cloud, а ФУЛ от %s пришёл раньше последней записи Cloud (%s) — "
                                      "возможно, ФУЛ пропустили. В новые строки не включён" %
                                      (order, fmt_date(primary.dt), fmt_date(last_cloud_date)))
            else:
                primary.kind = KIND_NEW
            for f in group_fuls[1:]:
                f.kind = KIND_DUP
                f.primary = primary
                d = fmt_date(f.dt)
                text = cfg["repeat_comment_template"].format(date=d or U)
                f.cloud_comment = text
                primary.dups.append(f)
                f.reason.append("Заказ %s повторяется внутри пачки: первый ФУЛ %s от %s. Отдельная строка не создаётся, "
                                "в комментарий строки заказа записано «%s»" % (order, primary.number, fmt_date(primary.dt) or U, text))
            if primary.dups:
                primary.cloud_comment = "; ".join(x.cloud_comment for x in primary.dups)

    # 3) группа
    for f in fuls:
        if f.kind in (KIND_NEW, KIND_MISS):
            f.group, f.group_why = rules.group(f.method_raw, f.method, f.country)

    # 4) выбор логиста: один поставщик (в пределах группы) — один логист
    by_id = {id(e): e for e in staff}

    def emp_label(e):
        return e.short

    def state_label(e):
        if e.excluded:
            return "не распределяет (%s)" % e.position
        return e.state or "состояние не указано"

    def choose(gkey, skey):
        members = sorted([e for e in staff if e.group == gkey and e.distributes], key=Employee.sort_key)
        avail = [e for e in members if e.available]
        hist = history.get(skey, Counter())
        hist_members = {id(e): hist[id(e)] for e in members if hist[id(e)] > 0}
        lines = []
        for e in members:
            h = hist_members.get(id(e), 0)
            if e.available:
                lines.append("%s: в работе %d, история с поставщиком %d" % (emp_label(e), load[id(e)], h))
            else:
                lines.append("%s: %s, история с поставщиком %d" % (emp_label(e), state_label(e), h))
        excluded = [e for e in staff if e.group == gkey and e.excluded]
        if excluded:
            lines.append("не распределяют: " + ", ".join("%s (%s)" % (emp_label(e), e.position) for e in excluded))
        cand_text = "; ".join(lines)
        reason = []
        hu = hist_unmatched.get(skey)
        if hu:
            reason.append("в истории поставщика есть логисты не из списка сотрудников: " +
                          ", ".join("%s (%d)" % (n, c) for n, c in sorted(hu.items())) + " — не учитываются")

        def least(pool):
            return min(pool, key=lambda e: (load[id(e)],) + e.sort_key())

        if not avail:
            reason.append("в группе «%s» нет доступных логистов (%s) — логист не назначен, в другую группу заказ не передаётся"
                          % (rules.groups[gkey]["title"],
                             ", ".join("%s — %s" % (emp_label(e), state_label(e)) for e in members) or "в списке сотрудников никого нет"))
            return None, reason, cand_text, "none"
        if hist_members:
            top = max(hist_members.values())
            top_set = [by_id[i] for i, c in hist_members.items() if c == top]
            top_avail = [e for e in top_set if e.available]
            names = ", ".join(emp_label(e) for e in sorted(top_set, key=Employee.sort_key))
            if top_avail:
                e = least(top_avail)
                if len(top_set) > 1:
                    reason.append("с поставщиком чаще всех работали %s (по %d заказов) — из них выбран наименее загруженный "
                                  "доступный: %s (в работе %d)" % (names, top, emp_label(e), load[id(e)]))
                else:
                    reason.append("предпочтительный логист по истории — %s (%d заказ(ов) с поставщиком), доступен"
                                  % (emp_label(e), top))
                return e, reason, cand_text, "history"
            e = least(avail)
            reason.append("предпочтительный по истории — %s (%d заказ(ов) с поставщиком), но недоступен: %s. "
                          "Выбран наименее загруженный доступный логист группы: %s (в работе %d)"
                          % (names, top, "; ".join(state_label(x) for x in top_set), emp_label(e), load[id(e)]))
            return e, reason, cand_text, "fallback_unavailable"
        e = least(avail)
        reason.append("у логистов группы нет истории с этим поставщиком — выбран наименее загруженный доступный: "
                      "%s (в работе %d)" % (emp_label(e), load[id(e)]))
        return e, reason, cand_text, "least_loaded"

    unit_choice = {}
    batch_assigned = Counter()
    miss_assigned = Counter()

    def assign(kind, counter):
        units = OrderedDict()
        for f in sorted([x for x in fuls if x.kind == kind], key=sort_key):
            if f.group is None:
                f.reason.insert(0, "Группа: %s — %s" % (U, f.group_why))
                f.reason.append("логист: %s — группа не определена" % U)
                what = f.group_why[0].upper() + f.group_why[1:]
                f.attention.append((what, "Строка %s, группа и логист: %s" %
                                    ("создана" if kind == KIND_NEW else "предложена", U),
                                    "Уточнить данные ФУЛ в 1С и выбрать логиста вручную"))
                continue
            if f.skey is None:
                f.reason.insert(0, "Группа: %s (%s)" % (rules.groups[f.group]["title"], f.group_why))
                f.reason.append("логист: %s — не указан поставщик, правила истории и «один поставщик — один логист» неприменимы" % U)
                f.attention.append(("Не указан поставщик (контрагент)", "Логист не назначен",
                                    "Указать поставщика и выбрать логиста"))
                continue
            units.setdefault((f.skey, f.group), []).append(f)
        for (skey, gkey), unit in units.items():
            if (skey, gkey) in unit_choice:
                e, reason, cand, how = unit_choice[(skey, gkey)]
                reason = ["%s — тот же логист, что у других заказов этого поставщика в пачке (один поставщик — один логист)"
                          % (("назначен " + e.short) if e else "логист не назначен")]
            else:
                e, reason, cand, how = choose(gkey, skey)
                unit_choice[(skey, gkey)] = (e, reason, cand, how)
            for f in unit:
                f.emp = e
                f.candidates = cand
                f.final = ("Итог: логист %s (%s)" % (e.short, e.fio)) if e else ("Итог: логист %s" % U)
                f.reason.insert(0, "Группа: %s (%s)" % (rules.groups[gkey]["title"], f.group_why))
                f.reason.extend(reason)
                if len(unit) > 1:
                    f.reason.append("все %d заказ(а) поставщика в этой группе в пачке переданы одному логисту" % len(unit))
                if e is None:
                    f.attention.append(("В группе «%s» нет доступных логистов" % rules.groups[gkey]["title"],
                                        "Строка %s без логиста" % ("создана" if kind == KIND_NEW else "предложена"),
                                        "Назначить логиста вручную (агент не передаёт заказ в другую группу)"))
            if e is not None:
                load[id(e)] += len(unit)
                counter[id(e)] += len(unit)

    assign(KIND_NEW, batch_assigned)
    assign(KIND_MISS, miss_assigned)

    # поставщик в нескольких группах — пояснение
    groups_by_sup = {}
    for f in fuls:
        if f.kind in (KIND_NEW, KIND_MISS) and f.group and f.skey:
            groups_by_sup.setdefault(f.skey, set()).add(f.group)
    for f in fuls:
        if f.kind in (KIND_NEW, KIND_MISS) and f.skey and len(groups_by_sup.get(f.skey, ())) > 1:
            f.notes.append("у поставщика в пачке заказы разных групп — логист выбран отдельно в каждой группе")
        if f.logist_1c and f.kind in (KIND_NEW, KIND_MISS):
            f.notes.append("в 1С у ФУЛ указан логист «%s» — правилами не используется" % f.logist_1c)

    # остальные пропущенные данные в строке Cloud
    for f in fuls:
        if f.kind == KIND_NEW and f.group == "air" and f.country is None:
            f.attention.append(("Не указана страна отправления", "Группа определена по способу доставки (air), "
                                "в строке Cloud страна: %s" % U, "Внести страну в Cloud"))
        if f.kind == KIND_NEW and f.brand is None:
            f.attention.append(("Не указан бренд (бизнес-единица)", "В строке Cloud бренд: %s" % U, "Внести бренд"))
        if f.kind == KIND_NEW and f.cutoff_raw is None:
            f.attention.append(("Не указан cut-off", "В строке Cloud cut-off: %s" % U, "Внести cut-off"))

    # --- строки Cloud
    form_keys = cfg["cloud_form_order"]
    cloud_map = tables["cloud"]["mapping"]
    present = sorted([k for k in form_keys if k in cloud_map], key=lambda k: cloud_map[k][0])
    absent = [k for k in form_keys if k not in cloud_map]
    form_cols = present + absent    # порядок как в реестре; отсутствующие в реестре колонки — в конце
    form_titles = [cloud_map[k][1] if k in cloud_map else cfg["cloud_form_default_titles"][k] for k in form_cols]
    if absent:
        general_attention.append({"what": "В реестре Cloud не найдены колонки: " +
                                  ", ".join(cfg["cloud_form_default_titles"][k] for k in absent),
                                  "action": "Колонки добавлены в конец формы с типовыми названиями",
                                  "decide": "Проверить форму реестра"})

    batch_spelling = {}
    for f in sorted(fuls, key=sort_key):
        if f.skey and f.supplier_raw:
            batch_spelling.setdefault(f.skey, f.supplier_raw)

    def supplier_display(f):
        """Написание как в Cloud (самое частое); для нового поставщика — как в первом ФУЛ пачки."""
        if f.supplier_raw is None:
            return U
        sp = supplier_spelling.get(f.skey)
        if sp:
            return sorted(sp.items(), key=lambda kv: (-kv[1], kv[0]))[0][0]
        return batch_spelling.get(f.skey, f.supplier_raw)

    def cloud_row(f):
        vals = {
            "order": f.order,
            "supplier": supplier_display(f),
            "brand": f.brand or U,
            "logist": f.emp.short if f.emp else "",
            "status": cfg["new_row_status"],
            "confirm_date": "",
            "cutoff": (f.cutoff_dt.date() if f.cutoff_dt else (f.cutoff_raw or U)),
            "ul_date": (f.dt.date() if f.dt else U),
            "country": f.country or U,
            "method": f.method or f.method_raw or U,
            "docs_date": "",
            "comment": f.cloud_comment,
        }
        return [vals[k] for k in form_cols]

    for f in fuls:
        if f.kind in (KIND_NEW, KIND_MISS):
            f.reason.append(getattr(f, "final", "Итог: логист %s" % U))
        if f.notes:
            f.reason.append("Примечание: " + "; ".join(f.notes))
        f.reason = [x[0].upper() + x[1:] if x else x for x in f.reason]

    new_list = sorted([f for f in fuls if f.kind == KIND_NEW], key=sort_key)
    miss_list = sorted([f for f in fuls if f.kind == KIND_MISS], key=sort_key)
    attention = []
    for f in sorted(fuls, key=lambda x: x.row):
        for what, action, decide in f.attention:
            attention.append({"ful": f.number, "row": f.row, "date": fmt_date(f.dt) or U, "order": f.order or U,
                              "supplier": f.supplier_raw or U, "what": what, "action": action, "decide": decide})
    for g in general_attention:
        attention.append(dict({"ful": "—", "row": "", "date": "", "order": "", "supplier": ""}, **g))

    # --- выгрузка
    os.makedirs(out_dir, exist_ok=True)
    dts = [f.dt for f in fuls if f.dt]
    stem = "Распределение_ФУЛ_%s" % (max(dts).strftime("%Y-%m-%d") if dts else "без_даты")
    xlsx_path = os.path.join(out_dir, stem + ".xlsx")
    json_path = os.path.join(out_dir, stem + ".json")

    def journal_record(f):
        assigned_kind = f.kind in (KIND_NEW, KIND_MISS)
        return OrderedDict([
            ("row", f.row), ("ful", f.number), ("date", fmt_date(f.dt) or U), ("order", f.order or U),
            ("supplier_1c", f.supplier_raw or U),
            ("supplier_cloud", supplier_display(f) if assigned_kind else "—"),
            ("method", f.method_raw or U), ("country", f.country or U), ("kind", f.kind),
            ("group", (rules.groups[f.group]["title"] if f.group else U) if assigned_kind else "—"),
            ("logist_fio", (f.emp.fio if f.emp else U) if assigned_kind else "—"),
            ("logist_cloud", (f.emp.short if f.emp else "") if assigned_kind else "—"),
            ("comment", f.cloud_comment), ("reason", ". ".join(x.rstrip(".") for x in f.reason) + "."), ("candidates", f.candidates),
            ("attention", "да" if f.attention else "нет"),
        ])

    journal = [journal_record(f) for f in sorted(fuls, key=lambda x: x.row)]
    staff_load = []
    for e in sorted(staff, key=lambda e: ((list(rules.groups).index(e.group) if e.group else 9),) + e.sort_key()):
        staff_load.append(OrderedDict([
            ("group", rules.groups[e.group]["title"] if e.group else "—"), ("fio", e.fio), ("short", e.short),
            ("position", e.position), ("state", e.state),
            ("distributes", "да" if (e.distributes and e.available) else "нет"),
            ("in_work", load_initial[id(e)]), ("assigned_new", batch_assigned[id(e)]),
            ("assigned_miss", miss_assigned[id(e)]),
            ("total", load_initial[id(e)] + batch_assigned[id(e)])]))

    counts = OrderedDict([
        ("total_ful", len(fuls)),
        ("new_rows", len(new_list)),
        ("new_with_logist", sum(1 for f in new_list if f.emp)),
        ("new_without_logist", sum(1 for f in new_list if not f.emp)),
        ("repeats", sum(1 for f in fuls if f.kind in (KIND_REPEAT, KIND_DUP))),
        ("repeats_in_cloud", sum(1 for f in fuls if f.kind == KIND_REPEAT)),
        ("dups_in_batch", sum(1 for f in fuls if f.kind == KIND_DUP)),
        ("already_in_cloud", sum(1 for f in fuls if f.kind == KIND_DONE)),
        ("possible_misses", len(miss_list)),
        ("not_processed", sum(1 for f in fuls if f.kind == KIND_SKIP)),
        ("attention", len(attention)),
        ("attention_ful", len({a["ful"] for a in attention if a["ful"] != "—"})),
    ])
    inputs_info = [OrderedDict([("type", KIND_TITLES[k]), ("file", os.path.basename(t["file"])), ("sheet", t["sheet"]),
                                ("header_row", t["header_row"]), ("rows", len(t["records"])),
                                ("columns", ", ".join("%s→%s" % (v[1], k) for k, v in sorted(t["mapping"].items(), key=lambda kv: kv[1][0])))])
                   for k, t in tables.items()]
    result = OrderedDict([
        ("counts", counts), ("last_cloud_date", fmt_date(last_cloud_date)), ("inputs", inputs_info),
        ("skipped_files", skipped), ("cloud_form", form_titles),
        ("new_rows", [dict(zip(form_titles, [fmt_cell(v) for v in cloud_row(f)]), **{"_ful": f.number}) for f in new_list]),
        ("repeats", [OrderedDict([("order", r["order"]), ("ful", r["f"].number), ("date", fmt_date(r["f"].dt) or U),
                                  ("type", r["type"]), ("logist", r["logist"]), ("append", r["append"]),
                                  ("after", r["after"])]) for r in repeats_rows] +
                    [OrderedDict([("order", f.order), ("ful", f.number), ("date", fmt_date(f.dt) or U),
                                  ("type", "дубль внутри пачки"), ("logist", f.primary.emp.short if f.primary.emp else ""),
                                  ("append", f.cloud_comment), ("after", f.primary.cloud_comment)])
                     for f in sorted([x for x in fuls if x.kind == KIND_DUP], key=sort_key)]),
        ("possible_misses", [dict(zip(form_titles, [fmt_cell(v) for v in cloud_row(f)]), **{"_ful": f.number}) for f in miss_list]),
        ("attention", attention), ("journal", journal), ("staff_load", staff_load),
        ("elapsed_sec", None), ("xlsx", os.path.basename(xlsx_path)),
    ])

    write_xlsx(xlsx_path, result, form_titles, [cloud_row(f) for f in new_list], new_list, miss_list,
               [cloud_row(f) for f in miss_list], last_cloud_date, U)
    result["elapsed_sec"] = round(time.time() - t0, 3)
    with open(json_path, "w", encoding="utf-8") as fh:
        json.dump(result, fh, ensure_ascii=False, indent=1, default=str)
    msg = chat_message(result, xlsx_path, cfg)
    print(msg)
    return 0


def fmt_cell(v):
    if isinstance(v, (date, datetime)):
        return v.strftime("%d.%m.%Y")
    return v


def fail(text, files=None, skipped=None):
    lines = ["ОШИБКА: " + text]
    if files:
        lines.append("Приложенные файлы: " + ", ".join(os.path.basename(f) for f in files))
    for s in skipped or []:
        lines.append("— " + s)
    print("\n".join(lines))
    return 2


# --------------------------------------------------------------------------- Excel

def write_xlsx(path, result, form_titles, new_rows, new_list, miss_list, miss_rows, last_cloud_date, U):
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    head_fill = PatternFill("solid", fgColor="1F4E78")
    head_font = Font(bold=True, color="FFFFFF")
    warn_fill = PatternFill("solid", fgColor="FFF2CC")
    bad_fill = PatternFill("solid", fgColor="F8CBAD")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def sheet(title, headers, rows, widths=None, date_cols=(), first=False):
        ws = wb.active if first else wb.create_sheet()
        ws.title = title
        ws.append(headers)
        for c in range(1, len(headers) + 1):
            cell = ws.cell(1, c)
            cell.fill, cell.font, cell.border = head_fill, head_font, border
            cell.alignment = Alignment(wrap_text=True, vertical="center")
        for row in rows:
            ws.append(list(row))
        for r in range(2, ws.max_row + 1):
            for c in range(1, len(headers) + 1):
                cell = ws.cell(r, c)
                cell.border = border
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                if isinstance(cell.value, (date, datetime)):
                    cell.number_format = "DD.MM.YYYY"
                if cell.value == U:
                    cell.fill = warn_fill
        for i, h in enumerate(headers, 1):
            w = (widths or {}).get(h) or max(10, min(45, max([len(str(h))] + [len(str(ws.cell(r, i).value or "")) for r in range(2, min(ws.max_row, 200) + 1)]) + 2))
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"
        if ws.max_row > 1:
            ws.auto_filter.ref = ws.dimensions
        return ws

    c = result["counts"]
    summary = [
        ("Всего ФУЛ в выгрузке", c["total_ful"]),
        ("Новых строк для Cloud", c["new_rows"]),
        ("  из них с логистом", c["new_with_logist"]),
        ("  из них без логиста", c["new_without_logist"]),
        ("Повторных отправок (всего)", c["repeats"]),
        ("  заказ уже в Cloud", c["repeats_in_cloud"]),
        ("  дубль внутри пачки", c["dups_in_batch"]),
        ("Уже учтены в Cloud (ничего делать не нужно)", c["already_in_cloud"]),
        ("Возможных пропусков", c["possible_misses"]),
        ("Не обработано (нет номера заказа, отменён и т. п.)", c["not_processed"]),
        ("Пунктов «Требует внимания»", c["attention"]),
        ("Последняя дата получения УЛ в Cloud", result["last_cloud_date"] or U),
        ("", ""),
        ("Агент ничего не записывает в Cloud и 1С. Строки переносит сотрудник после проверки.", ""),
    ]
    for inp in result["inputs"]:
        summary.append(("Источник: " + inp["type"], "%s, лист «%s», заголовки в строке %d, записей %d"
                        % (inp["file"], inp["sheet"], inp["header_row"], inp["rows"])))
    for s in result["skipped_files"]:
        summary.append(("Пропущен файл", s))
    sheet("Сводка", ["Показатель", "Значение"], summary, widths={"Показатель": 55, "Значение": 70}, first=True)

    ws = sheet("Новые строки Cloud", form_titles, new_rows)
    li = form_titles.index(next(t for t in form_titles if norm_text(t) == "логист")) + 1 if any(norm_text(t) == "логист" for t in form_titles) else None
    if li:
        for r in range(2, ws.max_row + 1):
            if not ws.cell(r, li).value:
                ws.cell(r, li).fill = bad_fill

    rep = result["repeats"]
    sheet("Повторные отправки",
          ["Заказ", "Номер ФУЛ", "Дата ФУЛ", "Тип", "Логист заказа", "Дописать в комментарий", "Комментарий после изменения"],
          [(r["order"], r["ful"], r["date"], r["type"], r["logist"], r["append"], r["after"]) for r in rep],
          widths={"Дописать в комментарий": 38, "Комментарий после изменения": 60})

    miss_headers = form_titles + ["Номер ФУЛ", "Дата ФУЛ", "Последняя запись Cloud", "Почему в этом списке"]
    sheet("Возможные пропуски", miss_headers,
          [r + [f.number, fmt_date(f.dt), fmt_date(last_cloud_date),
                "ФУЛ пришёл раньше последней записи Cloud, заказа в Cloud нет. Решите, вносить ли строку; "
                "группа и логист предложены по тем же правилам"] for r, f in zip(miss_rows, miss_list)])

    sheet("Требует внимания", ["Номер ФУЛ", "Строка 1С", "Дата ФУЛ", "Заказ", "Поставщик", "Что не так", "Что сделал агент", "Что решить"],
          [(a["ful"], a["row"], a["date"], a["order"], a["supplier"], a["what"], a["action"], a["decide"]) for a in result["attention"]],
          widths={"Что не так": 50, "Что сделал агент": 40, "Что решить": 45})

    jh = ["Строка 1С", "Номер ФУЛ", "Дата ФУЛ", "Заказ", "Поставщик (1С)", "Поставщик (для Cloud)", "Способ доставки",
          "Страна", "Решение", "Группа", "Логист (ФИО)", "Логист (для Cloud)", "Комментарий для Cloud", "Основание",
          "Кандидаты группы", "Требует внимания"]
    sheet(OUTPUT_MARK_SHEET, jh, [list(j.values()) for j in result["journal"]],
          widths={"Основание": 90, "Кандидаты группы": 70})

    sheet("Загрузка логистов", ["Группа", "Сотрудник", "Как пишется в Cloud", "Должность", "Состояние",
                                "Принимает заказы", "В работе (Cloud)", "Назначено сейчас", "Предложено по пропускам",
                                "В работе после пачки"],
          [list(s.values()) for s in result["staff_load"]])
    wb.save(path)


# --------------------------------------------------------------------------- ответ в чат

def chat_message(result, xlsx_path, cfg):
    c = result["counts"]
    lines = [
        "Распределение пачки ФУЛ готово.",
        "",
        "Обработано ФУЛ: %d" % c["total_ful"],
        "• Новых строк для Cloud: %d (с логистом — %d, без логиста — %d)" % (c["new_rows"], c["new_with_logist"], c["new_without_logist"]),
        "• Повторных отправок: %d (заказ уже в Cloud — %d, дубль внутри пачки — %d)" % (c["repeats"], c["repeats_in_cloud"], c["dups_in_batch"]),
        "• Возможных пропусков: %d" % c["possible_misses"],
        "• Требует внимания: %d вопрос(ов) по %d ФУЛ" % (c["attention"], c["attention_ful"]),
    ]
    if c["already_in_cloud"]:
        lines.append("• Уже учтены в Cloud, действий не нужно: %d" % c["already_in_cloud"])
    lines += ["", "Файл: %s" % os.path.basename(xlsx_path)]
    att = result["attention"]
    if att:
        lines += ["", "Главное, что нужно решить:"]
        for a in att[:7]:
            ref = "ФУЛ %s" % a["ful"] if a["ful"] != "—" else "Общее"
            if a.get("order") and a["order"] != cfg["undefined"]:
                ref += " (заказ %s)" % a["order"]
            lines.append("— %s: %s. %s." % (ref, a["what"], a["decide"]))
        if len(att) > 7:
            lines.append("— и ещё %d — на листе «Требует внимания»." % (len(att) - 7))
    lines += ["", "Агент ничего не записывал в Cloud и 1С: строки переносите после проверки. "
                  "Основание по каждому ФУЛ — на листе «Журнал решений»."]
    return "\n".join(lines)


def main():
    ap = argparse.ArgumentParser(description="Распределение новых ФУЛ между логистами")
    ap.add_argument("--inputs", nargs="+", required=True, help="файлы или папки с выгрузками (в любом порядке)")
    ap.add_argument("--out", required=True, help="папка для результата")
    ap.add_argument("--config", default=DEFAULT_CONFIG, help="справочник правил (references/config.json)")
    args = ap.parse_args()
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:  # noqa: BLE001
        pass
    sys.exit(run(args.inputs, args.out, args.config))


if __name__ == "__main__":
    main()
