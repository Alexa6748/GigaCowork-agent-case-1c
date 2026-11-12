#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Контрольные прогоны навыка и метрики приёмки.

Для каждого набора данных N раз запускает scripts/distribute.py, читает получившийся Excel
(как это сделал бы проверяющий) и сравнивает с эталоном (эталон.json из generator.py).

Метрики (по БТ):
  1. Точность распределения по полям: решение (статус ФУЛ), заказ, группа, логист, комментарий; доля ФУЛ,
     совпавших по всем полям (порог ≥ 95 %).
  2. Форма Cloud: колонки и порядок как в реестре, набор заказов = эталону, нет старых ФУЛ и дублей,
     точность по ячейкам (порог 100 % прогонов).
  3. Выдуманные значения на ловушках (порог 0).
  4. Воспроизводимость: содержимое результата одинаково во всех прогонах.
  5. Обоснование у каждого назначения логиста (порог 100 %).
  6. Время ответа (порог ≤ 60 с).
  + полнота «Требует внимания» (каждая ловушка, требующая решения человека, попала в список).

Запуск: python3 evaluate.py --skill ../skill/ful-distribution --datasets datasets --runs 3 --out results
"""
import argparse
import hashlib
import json
import os
import subprocess
import sys
import time
from datetime import date, datetime

import openpyxl

U = "не определено"
FIELDS = ["kind", "order", "group", "logist_fio", "comment"]
FIELD_TITLES = {"kind": "Решение (статус ФУЛ)", "order": "Заказ", "group": "Группа", "logist_fio": "Логист",
                "comment": "Комментарий о повторе"}
JOURNAL_COLS = {"ful": "Номер ФУЛ", "kind": "Решение", "order": "Заказ", "group": "Группа",
                "logist_fio": "Логист (ФИО)", "comment": "Комментарий для Cloud", "reason": "Основание",
                "logist_cloud": "Логист (для Cloud)"}


def cell_str(v):
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%d.%m.%Y")
    if isinstance(v, date):
        return v.strftime("%d.%m.%Y")
    return str(v).strip()


def read_sheet(wb, name):
    ws = wb[name]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    if not rows:
        return [], []
    head = [cell_str(h) for h in rows[0]]
    data = [[cell_str(v) for v in r] for r in rows[1:] if any(v is not None for v in r)]
    return head, data


def read_result(xlsx):
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    res = {}
    jh, jd = read_sheet(wb, "Журнал решений")
    idx = {k: jh.index(t) for k, t in JOURNAL_COLS.items()}
    res["journal"] = {r[idx["ful"]]: {k: r[i] for k, i in idx.items()} for r in jd}
    res["new_head"], res["new_rows"] = read_sheet(wb, "Новые строки Cloud")
    mh, md = read_sheet(wb, "Возможные пропуски")
    res["miss_head"], res["miss_rows"] = mh, md
    ah, ad = read_sheet(wb, "Требует внимания")
    res["attention_fuls"] = sorted({r[ah.index("Номер ФУЛ")] for r in ad if r[ah.index("Номер ФУЛ")] not in ("", "—")})
    rh, rd = read_sheet(wb, "Повторные отправки")
    res["repeats"] = [dict(zip(rh, r)) for r in rd]
    wb.close()
    return res


def content_hash(res):
    return hashlib.sha256(json.dumps(res, ensure_ascii=False, sort_keys=True).encode()).hexdigest()[:16]


def expected_attention(exp):
    s = set()
    for e in exp["fuls"]:
        if e["kind"] == "Не обработан":
            s.add(e["ful"])
        elif e["kind"] in ("Новый", "Возможный пропуск") and (e["group"] == U or e["logist_fio"] == U):
            s.add(e["ful"])
        elif "Требует внимания" in e["must"]:
            s.add(e["ful"])
    return s


def evaluate_run(exp, res, cloud_orders):
    m = {}
    per_field = {f: 0 for f in FIELDS}
    full = 0
    mismatches = []
    for e in exp["fuls"]:
        a = res["journal"].get(e["ful"])
        ok_all = True
        for f in FIELDS:
            av = a[f] if a else "<нет в журнале>"
            if av == e[f]:
                per_field[f] += 1
            else:
                ok_all = False
                mismatches.append("%s: %s — ожидалось «%s», получено «%s»" % (e["ful"], FIELD_TITLES[f], e[f], av))
        full += ok_all
    n = len(exp["fuls"])
    m["n_ful"] = n
    m["full_match"] = full / n
    m["per_field"] = {FIELD_TITLES[f]: per_field[f] / n for f in FIELDS}

    # форма Cloud
    form_ok = res["new_head"] == exp["cloud_form"]
    order_col = exp["cloud_form"].index("Заказ") if "Заказ" in exp["cloud_form"] else 0
    got_orders = [r[order_col] for r in res["new_rows"]] if form_ok else []
    exp_orders = [r["Заказ"] for r in exp["new_rows"]]
    dup_or_old = [o for o in got_orders if o in cloud_orders] + [o for o in set(got_orders) if got_orders.count(o) > 1]
    cells_total = cells_ok = 0
    exp_by_order = {r["Заказ"]: r for r in exp["new_rows"]}
    cell_mismatch = []
    for r in res["new_rows"] if form_ok else []:
        er = exp_by_order.get(r[order_col])
        for j, t in enumerate(exp["cloud_form"]):
            cells_total += 1
            if er and r[j] == er[t]:
                cells_ok += 1
            elif er:
                cell_mismatch.append("%s / %s: ожидалось «%s», получено «%s»" % (r[order_col], t, er[t], r[j]))
    m["form_columns_ok"] = form_ok
    m["form_orders_ok"] = sorted(got_orders) == sorted(exp_orders)
    m["form_no_old_or_dups"] = not dup_or_old
    m["form_cells_accuracy"] = (cells_ok / cells_total) if cells_total else 0.0
    m["form_pass"] = form_ok and m["form_orders_ok"] and m["form_no_old_or_dups"] and cells_ok == cells_total
    mismatches += cell_mismatch

    # пропуски
    miss_got = sorted(r[order_col] for r in res["miss_rows"]) if res["miss_head"][:len(exp["cloud_form"])] == exp["cloud_form"] else None
    m["misses_ok"] = miss_got == sorted(r["Заказ"] for r in exp["possible_misses"])

    # выдуманные значения
    invented = []
    for t in exp["undefined_traps"]:
        a = res["journal"].get(t["ful"], {})
        for f, want in t["fields"].items():
            if a.get(f) not in (want, "", None):
                invented.append("%s: %s = «%s»" % (t["ful"], FIELD_TITLES[f], a.get(f)))
        if "logist_fio" in t["fields"] and a.get("logist_cloud") not in ("", None):
            invented.append("%s: в строке Cloud логист «%s»" % (t["ful"], a.get("logist_cloud")))
    m["trap_checks"] = sum(len(t["fields"]) for t in exp["undefined_traps"])
    m["invented_values"] = len(invented)

    # обоснование назначений
    assigned = [a for a in res["journal"].values() if a["kind"] in ("Новый", "Возможный пропуск") and a["logist_fio"] not in (U, "", "—")]
    justified = [a for a in assigned if a["reason"] and a["logist_cloud"] and a["logist_cloud"] in a["reason"]]
    m["assigned"] = len(assigned)
    m["justified_share"] = (len(justified) / len(assigned)) if assigned else 1.0

    # внимание
    ea = expected_attention(exp)
    got = set(res["attention_fuls"])
    m["attention_recall"] = len(ea & got) / len(ea) if ea else 1.0
    m["attention_extra"] = sorted(got - ea)
    m["attention_missed"] = sorted(ea - got)
    m["mismatches"] = mismatches + ["выдумано: " + x for x in invented]
    return m


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--skill", required=True)
    ap.add_argument("--datasets", required=True)
    ap.add_argument("--runs", type=int, default=3)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    script = os.path.join(a.skill, "scripts", "distribute.py")
    report = {"runs": a.runs, "datasets": {}}
    for name in sorted(os.listdir(a.datasets)):
        folder = os.path.join(a.datasets, name)
        exp_path = os.path.join(folder, "эталон.json")
        if not os.path.isfile(exp_path):
            continue
        with open(exp_path, encoding="utf-8") as fh:
            exp = json.load(fh)
        # заказы, которые уже есть в Cloud (для проверки «без старых ФУЛ»)
        cloud_orders = set(exp["cloud_orders"])
        runs = []
        for i in range(1, a.runs + 1):
            out = os.path.join(a.out, name, "run%d" % i)
            os.makedirs(out, exist_ok=True)
            t0 = time.time()
            p = subprocess.run([sys.executable, script, "--inputs", os.path.join(folder, "входные_файлы"), "--out", out],
                               capture_output=True, text=True)
            elapsed = time.time() - t0
            with open(os.path.join(out, "ответ_в_чат.txt"), "w", encoding="utf-8") as fh:
                fh.write(p.stdout)
            xlsx = [f for f in os.listdir(out) if f.endswith(".xlsx")]
            res = read_result(os.path.join(out, xlsx[0]))
            m = evaluate_run(exp, res, cloud_orders)
            m["elapsed_sec"] = round(elapsed, 2)
            m["hash"] = content_hash(res)
            m["exit_code"] = p.returncode
            runs.append(m)
        hashes = {r["hash"] for r in runs}
        report["datasets"][name] = {"runs": runs, "reproducible": len(hashes) == 1}
    os.makedirs(a.out, exist_ok=True)
    with open(os.path.join(a.out, "metrics.json"), "w", encoding="utf-8") as fh:
        json.dump(report, fh, ensure_ascii=False, indent=1)

    # таблица метрик
    lines = ["| Набор | Прогон | Полное совпадение | Решение | Заказ | Группа | Логист | Комментарий | Форма Cloud | "
             "Ячейки Cloud | Пропуски | Выдумано | Обоснование | Внимание (полнота) | Время, с | Хеш |",
             "|" + "---|" * 16]
    for name, d in report["datasets"].items():
        for i, r in enumerate(d["runs"], 1):
            pf = r["per_field"]
            lines.append("| %s | %d | %.0f%% | %.0f%% | %.0f%% | %.0f%% | %.0f%% | %.0f%% | %s | %.0f%% | %s | %d из %d | %.0f%% | %.0f%% | %.2f | %s |" % (
                name, i, 100 * r["full_match"], 100 * pf["Решение (статус ФУЛ)"], 100 * pf["Заказ"], 100 * pf["Группа"],
                100 * pf["Логист"], 100 * pf["Комментарий о повторе"], "да" if r["form_pass"] else "НЕТ",
                100 * r["form_cells_accuracy"], "да" if r["misses_ok"] else "НЕТ", r["invented_values"], r["trap_checks"],
                100 * r["justified_share"], 100 * r["attention_recall"], r["elapsed_sec"], r["hash"]))
        lines.append("| %s | — | воспроизводимость: %s |" % (name, "одинаково во всех прогонах" if d["reproducible"] else "РАЗЛИЧАЕТСЯ"))
    text = "\n".join(lines)
    with open(os.path.join(a.out, "metrics.md"), "w", encoding="utf-8") as fh:
        fh.write(text + "\n")
    print(text)
    for name, d in report["datasets"].items():
        mm = d["runs"][0]["mismatches"]
        if mm:
            print("\n%s — расхождения с эталоном (прогон 1):" % name)
            for x in mm:
                print("  ", x)
        if d["runs"][0]["attention_extra"]:
            print("%s — дополнительно в «Требует внимания»: %s" % (name, ", ".join(d["runs"][0]["attention_extra"])))


if __name__ == "__main__":
    main()
