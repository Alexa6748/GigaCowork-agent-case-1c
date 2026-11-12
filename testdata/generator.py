#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Генератор синтетических данных для агента распределения ФУЛ + эталон.

Наборы:
  A_основной   — 26 ФУЛ, 45 строк Cloud, 12 сотрудников; все обязательные ловушки из задания.
  B_формат     — те же данные, что A, но другой порядок колонок, лишние колонки, другие имена листов,
                 заголовок не в первой строке, даты строкой, сотрудники в .xls, безликие имена файлов.
                 Эталон по ФУЛ совпадает с A, форма Cloud — в порядке колонок реестра B.
  C_крайние    — 20 ФУЛ, 32 строки Cloud, 11 сотрудников; дополнительные крайние случаи.

Эталон (ожидаемое решение по каждому ФУЛ) записан вручную в описании сценария — он не вычисляется
скриптом агента. Все компании, бренды и сотрудники вымышленные.

Запуск: python3 generator.py --out <папка>
"""
import argparse
import json
import os
from datetime import date, datetime, timedelta

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

U = "не определено"
Y = 2026


def d(day, month):
    return date(Y, month, day)


def dt(day, month, hh, mm):
    return datetime(Y, month, day, hh, mm, 0)


# ============================================================================ НАБОР A

STAFF_A = [
    # ФИО, состояние, должность, таб. №
    ("Орлова Анна Сергеевна", "Уволен", "Специалист по авиаперевозкам", "00101"),
    ("Белов Игорь Николаевич", "Ежегодный оплачиваемый отпуск", "Специалист по авиаперевозкам", "00102"),
    ("Ершова Дарья Олеговна", "Отпуск без сохранения заработной платы", "Специалист по авиаперевозкам", "00103"),
    ("Миронов Олег Петрович", "Работает", "Руководитель группы авиаперевозок", "00105"),
    ("Ким Виктор Андреевич", "Уволен", "Специалист по работе с ВЭД документацией", "00201"),
    ("Лаптева Елена Викторовна", "Работает", "Специалист по работе с ВЭД документацией", "00202"),
    ("Кузнецова Мария Петровна", "Работает", "Специалист по работе с ВЭД документацией", "00203"),
    ("Кузнецова Ольга Игоревна", "Работает", "Ассистент специалиста по работе с ВЭД документацией", "00204"),
    ("Фёдоров Павел Ильич", "Ежегодный оплачиваемый отпуск", "Специалист по мультимодальным перевозкам", "00301"),
    ("Соколова Ирина Юрьевна", "Работает", "Специалист по мультимодальным перевозкам", "00302"),
    ("Громов Артём Денисович", "Работает", "Специалист по мультимодальным перевозкам", "00303"),
    ("Захарова Нина Борисовна", "Работает", "Руководитель группы мультимодальных перевозок", "00304"),
]

SHZ = "Shenzhen Bright Textile Co., Ltd"
GZ = "Guangzhou Silkway Garment Co., Ltd"
HKP = "Hong Kong Pearl Apparel Ltd"
XM = "Xiamen Lucky Shoes Co., Ltd"
YW = "Yiwu Star Accessories"
ANA = "Anatolia Tekstil A.Ş."
DHK = "Dhaka Cotton Mills Ltd"
SEV = "ООО «Северный Трикотаж»"
HAN = "Hanoi Garment JSC"
BUR = "Bursa Denim Sanayi"
MIL = "Milano Moda S.r.l."
SEO = "Seoul Fashion Lab"
PAR = "Paris Atelier SAS"

# Заказ, Поставщик, Бренд, Логист (как в Cloud), статус, дата УЛ, страна, способ, комментарий
CLOUD_A = [
    ("S0010001", SHZ, "NordLine", "Ким В.А.", "confirm", d(5, 9), "Китай", "sea", ""),
    ("S0010002", SHZ, "UrbanFox", "Ким В.А.", "confirm", d(12, 9), "Китай", "sea", ""),
    ("S0010003", SHZ, "NordLine", "Ким В.А.", "confirm", d(26, 9), "Китай", "railway", ""),
    ("S0010004", SHZ, "NordLine", "Ким В.А.", "в работе", d(20, 10), "Китай", "sea", ""),
    ("S0010006", SHZ, "UrbanFox", "Лаптева Е.", "confirm", d(18, 9), "КНР", "sea", ""),
    ("S0010005", GZ, "Bellamare", "Кузнецова М.", "в работе", d(15, 10), "Китай", "sea", ""),
    ("S0010007", GZ, "Bellamare", "Кузнецова М.", "в работе", d(22, 10), "Китай", "sea", ""),
    ("S0010008", GZ, "Bellamare", "Кузнецова М.", "confirm", d(10, 9), "Китай", "sea", ""),
    ("S0010009", HKP, "Lumen", "Лаптева Е.В.", "в работе", d(25, 10), "Гонконг", "sea", ""),
    ("S0010010", HKP, "Lumen", "Лаптева Е.В.", "confirm", d(8, 9), "Гонконг", "sea", ""),
    ("S0010011", XM, "StepUp", "Кузнецова М.П.", "в работе", d(27, 10), "Китай", "sea", ""),
    ("S0010012", XM, "StepUp", "Кузнецова М.П.", "в работе", d(29, 10), "Китай", "railway", ""),
    ("S0010013", XM, "StepUp", "Кузнецова М.", "в работе", d(30, 10), "Китай", "sea", ""),
    ("S0010014", XM, "StepUp", "Петренко С.", "confirm", d(2, 9), "Китай", "sea", ""),
    ("S0010015", XM, "StepUp", "Петренко С.", "confirm", d(4, 9), "Китай", "sea", ""),
    ("S0010016", YW, "Lumen", "Лаптева Е.В.", "в работе", d(21, 10), "Китай", "railway", ""),
    ("S0010017", YW, "Lumen", "Лаптева Е.", "в работе", d(24, 10), "Китай", "sea", ""),
    ("S0010018", YW, "Lumen", "Лаптева Е.В.", "в работе", d(28, 10), "Китай", "sea", ""),
    ("S0010019", YW, "Lumen", "Кузнецова М.", "confirm", d(14, 9), "Китай", "sea", ""),
    ("S0010030", ANA, "UrbanFox", "Федоров П.И.", "в работе", d(19, 10), "Турция", "truck", ""),
    ("S0010031", ANA, "UrbanFox", "Федоров П.И.", "confirm", d(3, 9), "Турция", "truck", ""),
    ("S0010032", ANA, "NordLine", "Федоров П.И.", "confirm", d(17, 9), "Турция", "truck", ""),
    ("S0010033", ANA, "NordLine", "Громов А.Д.", "confirm", d(1, 9), "Турция", "truck", ""),
    ("S0010034", DHK, "Bellamare", "Соколова И.Ю.", "в работе", d(23, 10), "Бангладеш", "sea", ""),
    ("S0010035", DHK, "Bellamare", "Соколова И.Ю.", "confirm", d(11, 9), "Бангладеш", "sea", ""),
    ("S0010036", DHK, "Bellamare", "Громов А.Д.", "confirm", d(9, 9), "Бангладеш", "sea", ""),
    ("S0010037", SEV, "Lumen", "Громов А.", "в работе", d(26, 10), "Россия", "truck", ""),
    ("S0010038", SEV, "Lumen", "Громов А.", "confirm", d(15, 9), "Россия", "direct truck", ""),
    ("S0010039", HAN, "StepUp", "соколова и.ю.", "в работе", d(16, 10), "Вьетнам", "sea", ""),
    ("S0010040", HAN, "StepUp", "соколова и.ю.", "в работе", d(28, 10), "Вьетнам", "sea",
     "Повторная отправка ФУЛа, 30.10.2026"),
    ("S0010041", HAN, "StepUp", "Громов А.Д.", "в работе", d(20, 10), "Вьетнам", "sea", ""),
    ("S0010042", HAN, "StepUp", "Громов А.Д.", "в работе", d(31, 10), "Вьетнам", "sea", ""),
    ("S0010043", HAN, "StepUp", "Федоров П.И.", "в работе", d(13, 10), "Вьетнам", "sea", ""),
    ("S0010044", HAN, "StepUp", "Громов А.Д.", "confirm", d(6, 9), "Вьетнам", "sea", ""),
    ("S0010045", BUR, "NordLine", "Громов А.Д.", "в работе", d(2, 11), "Турция", "truck", ""),
    ("S0010046", BUR, "NordLine", "Соколова И.Ю.", "confirm", d(19, 9), "Турция", "truck", ""),
    ("S0010050", MIL, "Bellamare", "Орлова А.С.", "в работе", d(22, 10), "Италия", "air", ""),
    ("S0010051", MIL, "Bellamare", "Орлова А.С.", "confirm", d(7, 9), "Италия", "air", ""),
    ("S0010052", MIL, "Bellamare", "Белов И.Н.", "confirm", d(16, 9), "Италия", "air", ""),
    ("S0010053", SEO, "UrbanFox", "Белов И.Н.", "в работе", d(27, 10), "Республика Корея", "air", ""),
    ("S0010054", SEO, "UrbanFox", "Белов И.Н.", "в работе", d(30, 10), "Республика Корея", "air", ""),
    ("S0010055", SEO, "UrbanFox", "Ершова Д.О.", "в работе", d(29, 10), "Республика Корея", "air", ""),
    ("S0010056", SEO, "UrbanFox", "Орлова А.С.", "в работе", d(1, 11), "Республика Корея", "air", ""),
    ("S0010057", PAR, "Lumen", "Белов И.Н.", "в работе", d(31, 10), "Франция", "air", ""),
    ("S0010058", PAR, "Lumen", "Ершова Д.О.", "confirm", d(12, 9), "Франция", "air", ""),
]

LAP = "Лаптева Елена Викторовна"
KMP = "Кузнецова Мария Петровна"
SOK = "Соколова Ирина Юрьевна"
GRO = "Громов Артём Денисович"
NO = ""   # логист не назначен

GV, GM, GA = "ВЭД", "Мультимодальные перевозки", "Авиаперевозки"


def ful(num, when, order, sup, brand, method, country, exp, *, comment=None, cutoff=None, status="Проведен",
        logist_1c=None, exp_supplier=None, trap="", must=""):
    """exp = (решение, группа, логист ФИО, комментарий для Cloud)."""
    if comment is None:
        comment = "Номер: %s от %s" % (order, (when - timedelta(days=7)).strftime("%d.%m.%Y")) if order else ""
    if cutoff is None:
        cutoff = (when + timedelta(days=21)).date()
    elif cutoff == "-":
        cutoff = None
    return dict(num=num, when=when, order=order, comment=comment, sup=sup, brand=brand, method=method,
                country=country, status=status, logist_1c=logist_1c, cutoff=cutoff,
                kind=exp[0], group=exp[1], logist=exp[2], cloud_comment=exp[3],
                exp_supplier=exp_supplier, trap=trap, must=must)


R = "Повторная отправка ФУЛа, %s"
DASH = "—"

FUL_A = [
    ful("УЛ-1001", dt(3, 11, 9, 15), "S0011001", SHZ, "NordLine", "sea", "Китай", ("Новый", GV, LAP, ""),
        trap="Предпочтительный логист уволен", must="История: Ким В.А. (4) — уволен → наименее загруженный в ВЭД: Лаптева (4) < Кузнецова М.П. (5). Кузнецова О.И. — ассистент, не участвует"),
    ful("УЛ-1002", dt(3, 11, 9, 40), "S0011002", "SHENZHEN BRIGHT TEXTILE CO LTD", "UrbanFox", "sea", "Китай",
        ("Новый", GV, LAP, ""), exp_supplier=SHZ,
        trap="Поставщик записан иначе (регистр, нет знаков)", must="Тот же поставщик → тот же логист, что у УЛ-1001; в Cloud — написание из реестра"),
    ful("УЛ-1003", dt(4, 11, 11, 5), "S0011003", "  Shenzhen  Bright-Textile Co.,Ltd. ", "NordLine", "railway", "КНР",
        ("Новый", GV, LAP, ""), exp_supplier=SHZ,
        trap="Поставщик с лишними пробелами и дефисом; страна «КНР»", must="Тот же поставщик; КНР = Китай → ВЭД"),
    ful("УЛ-1004", dt(3, 11, 10, 20), "S0011004", GZ, "Bellamare", "sea", "Китай", ("Новый", GV, KMP, ""),
        trap="Короткое имя «Кузнецова М.» в Cloud при двух Кузнецовых в 1С", must="Сопоставить с Кузнецовой Марией Петровной (не с ассистентом Ольгой); предпочтительный доступен"),
    ful("УЛ-1005", dt(3, 11, 12, 0), "S0011005", HKP, "Lumen", "sea", "Гонконг", ("Новый", GV, LAP, ""),
        trap="Страна — Гонконг", must="Гонконг относится к Китаю → ВЭД; предпочтительный Лаптева"),
    ful("УЛ-1006", dt(4, 11, 8, 30), "S0011006", "Ningbo Ocean Knit Co., Ltd", "UrbanFox", "sea", "Китай",
        ("Новый", GV, KMP, ""), trap="Новый поставщик", must="Нет истории → наименее загруженный: Кузнецова М.П. (6) < Лаптева (8)"),
    ful("УЛ-1007", dt(4, 11, 8, 45), "S0011007", "Ningbo Ocean Knit Co Ltd", "UrbanFox", "sea", "Китай",
        ("Новый", GV, KMP, ""), exp_supplier="Ningbo Ocean Knit Co., Ltd",
        trap="Один поставщик — два заказа", must="Тот же логист, что у УЛ-1006"),
    ful("УЛ-1008", dt(5, 11, 13, 0), "S0011008", "Qingdao Blue Wave Fashion Co., Ltd", "Lumen", "truck", "Hong Kong",
        ("Новый", GV, KMP, ""), trap="Ничья по загрузке (8 = 8); страна «Hong Kong»",
        must="Ничья решается по алфавиту ФИО → Кузнецова М.П."),
    ful("УЛ-1009", dt(3, 11, 13, 10), "S0011009", ANA, "UrbanFox", "truck", "Турция", ("Новый", GM, SOK, ""),
        trap="Предпочтительный логист в отпуске; в Cloud «Федоров» через «е»",
        must="Фёдоров (3) в отпуске → наименее загруженный: Соколова (3) < Громов (4)"),
    ful("УЛ-1010", dt(3, 11, 14, 0), "S0011010", DHK, "Bellamare", "sea", "Бангладеш",
        ("Новый", GM, SOK, R % "05.11.2026"), trap="Позже в пачке придёт повтор этого заказа",
        must="Одна строка; в комментарий — «Повторная отправка ФУЛа, 05.11.2026»"),
    ful("УЛ-1011", dt(4, 11, 9, 50), "S0011011", 'ООО "Северный трикотаж"', "Lumen", "direct truck", "Россия",
        ("Новый", GM, GRO, ""), exp_supplier=SEV, trap="Поставщик: другие кавычки и регистр",
        must="Сопоставить с «ООО «Северный Трикотаж»», предпочтительный Громов"),
    ful("УЛ-1012", dt(4, 11, 15, 30), "S0011012", "Tashkent Textile Group", "NordLine", "railway", "Узбекистан",
        ("Новый", GM, GRO, ""), trap="Новый поставщик, ничья по загрузке (5 = 5)", must="По алфавиту → Громов"),
    ful("УЛ-1013", dt(5, 11, 10, 0), "S0011013", "Tashkent Textile Group", "NordLine", "railway", "Узбекистан",
        ("Новый", GM, GRO, ""), trap="Один поставщик — один логист", must="Тот же логист, что у УЛ-1012"),
    ful("УЛ-1014", dt(5, 11, 16, 20), "S0011010", DHK, "Bellamare", "sea", "Бангладеш",
        ("Дубль в пачке", DASH, DASH, R % "05.11.2026"), trap="Повтор заказа внутри одной пачки",
        must="Не создавать вторую строку"),
    ful("УЛ-1015", dt(3, 11, 11, 30), "S0011015", MIL, "Bellamare", "air", "Италия", ("Новый", GA, NO, ""),
        trap="Вся группа авиаперевозок недоступна (уволена, отпуска, руководитель)",
        must="Строка без логиста, «Требует внимания», в другую группу не передавать"),
    ful("УЛ-1016", dt(4, 11, 12, 15), "S0011016", SHZ, "NordLine", "AIR", "Китай", ("Новый", GA, NO, ""),
        trap="Китай, но air; способ доставки заглавными",
        must="air → авиаперевозки (не ВЭД); группа недоступна → без логиста"),
    ful("УЛ-1017", dt(3, 11, 15, 45), "S0010005", GZ, "Bellamare", "sea", "Китай",
        ("Повторная отправка", DASH, DASH, R % "03.11.2026"), trap="Заказ уже есть в Cloud",
        must="Не создавать строку; дописать комментарий"),
    ful("УЛ-1018", dt(5, 11, 9, 0), "S0010030", "Anatolia Tekstil A.S.", "UrbanFox", "truck", "Турция",
        ("Повторная отправка", DASH, DASH, R % "05.11.2026"), trap="Повтор; поставщик «A.S.» вместо «A.Ş.»",
        must="Не создавать строку; поставщик тот же — без замечания"),
    ful("УЛ-1019", dt(6, 11, 9, 0), "S0010005", GZ, "Bellamare", "sea", "Китай",
        ("Повторная отправка", DASH, DASH, R % "06.11.2026"), trap="Второй повтор того же заказа",
        must="Дописать вторую отметку после первой"),
    ful("УЛ-0990", dt(30, 10, 10, 0), "S0010040", HAN, "StepUp", "sea", "Вьетнам",
        ("Уже учтён в Cloud", DASH, DASH, ""), trap="Повтор уже отмечен в комментарии Cloud",
        must="Ничего не делать, второй раз не дописывать"),
    ful("УЛ-0975", dt(31, 10, 14, 20), "S0011020", DHK, "Bellamare", "sea", "Бангладеш",
        ("Возможный пропуск", GM, SOK, ""), trap="ФУЛ раньше последней записи Cloud (02.11), заказа в Cloud нет",
        must="Не включать в новые строки; показать в «Возможные пропуски» с предложенным логистом"),
    ful("УЛ-0980", dt(1, 11, 9, 30), "S0011021", SEO, "UrbanFox", "air", "Республика Корея",
        ("Возможный пропуск", GA, NO, ""), trap="Пропуск + группа недоступна", must="В пропуски, логист пустой"),
    ful("УЛ-1024", dt(4, 11, 10, 10), None, YW, "Lumen", "sea", "Китай", ("Не обработан", DASH, DASH, ""),
        comment="Номер:  от 04.11.2026", trap="Нет номера заказа", must="Не создавать строку; заказ «не определено»; «Требует внимания»"),
    ful("УЛ-1025", dt(5, 11, 11, 0), "S0011025", BUR, "NordLine", None, "Турция", ("Новый", U, NO, ""),
        trap="Нет способа доставки", must="Группа «не определено», логист пустой, способ доставки «не определено»"),
    ful("УЛ-1026", dt(5, 11, 12, 30), "S0011026", HAN, "StepUp", "sea", None, ("Новый", U, NO, ""),
        trap="Нет страны при неавиа-доставке", must="Группа «не определено», логист пустой, страна «не определено»"),
    ful("УЛ-1027", dt(5, 11, 14, 0), "S0011027", PAR, "Lumen", "air", None, ("Новый", GA, NO, ""),
        trap="Нет страны при air", must="Группа определяется по air; страна «не определено»; группа недоступна → без логиста"),
]

# ============================================================================ НАБОР C

STAFF_C = [
    ("Орлова Анна Сергеевна", "Работает", "Специалист по авиаперевозкам", "10101"),
    ("Белов Игорь Николаевич", "Работает", "Ведущий специалист по авиаперевозкам", "10102"),
    ("Ершова Дарья Олеговна", "Отпуск по беременности и родам", "Специалист по авиаперевозкам", "10103"),
    ("Ким Виктор Андреевич", "Уволен", "Специалист по работе с ВЭД документацией", "10201"),
    ("Лаптева Елена Викторовна", "Ежегодный отпуск", "Специалист по работе с ВЭД документацией", "10202"),
    ("Кузнецова Ольга Игоревна", "Работает", "Ассистент специалиста по работе с ВЭД документацией", "10204"),
    ("Миронов Олег Петрович", "Работает", "Руководитель группы ВЭД", "10205"),
    ("Соколова Ирина Юрьевна", "Работает", "Специалист по мультимодальным перевозкам", "10302"),
    ("Громов Артём Денисович", "Работает", "Специалист по мультимодальным перевозкам", "10303"),
    ("Белова Ирина Николаевна", "Работает", "Специалист по мультимодальным перевозкам", "10305"),
    ("Захарова Нина Борисовна", "Работает", "Начальник отдела логистики", "10304"),
]

TOK = "Tokyo Denim Works"
SHC = "Shanghai Cloud Silk Co., Ltd"
TPE = "Taipei Fabric Co"
IZM = "Izmir Cotton A.Ş."
ANK = "Ankara Wool Ltd"
KAZ = "Kazan Textile"

CLOUD_C = [
    ("S0020001", SEO, "UrbanFox", "Белов И.Н.", "в работе", d(20, 10), "Республика Корея", "air", ""),
    ("S0020002", SEO, "UrbanFox", "Белов И.Н.", "confirm", d(10, 9), "Республика Корея", "air", ""),
    ("S0020003", TOK, "NordLine", "Орлова А.С.", "confirm", d(5, 9), "Япония", "air", ""),
    ("S0020004", TOK, "NordLine", "Орлова А.", "confirm", d(15, 9), "Япония", "air", ""),
    ("S0020005", TOK, "NordLine", "Белов И.Н.", "в работе", d(25, 10), "Япония", "air", ""),
    ("S0020006", TOK, "NordLine", "Белов И.", "confirm", d(12, 9), "Япония", "air", ""),
    ("S0020007", MIL, "Bellamare", "Орлова А.С.", "в работе", d(28, 10), "Италия", "air", ""),
    ("S0020008", MIL, "Bellamare", "Орлова А.С.", "в работе", d(30, 10), "Италия", "air", ""),
    ("S0020009", PAR, "Lumen", "Орлова А.С.", "в работе", d(1, 11), "Франция", "air", ""),
    ("S0020010", PAR, "Lumen", "Ершова Д.О.", "в работе", d(2, 11), "Франция", "air", ""),
    ("S0020012", MIL, "Bellamare", "Ершова Д.О.", "в работе", d(27, 10), "Италия", "air", ""),
    ("S0020013", SHC, "Lumen", "Лаптева Е.В.", "в работе", d(21, 10), "Китай", "sea", ""),
    ("S0020014", SHC, "Lumen", "Лаптева Е.В.", "в работе", d(29, 10), "Китай", "sea", ""),
    ("S0020015", SHC, "Lumen", "Ким В.А.", "в работе", d(18, 10), "Китай", "sea", ""),
    ("S0020016", "Ningbo Ocean Knit Co., Ltd", "UrbanFox", "Ким В.А.", "confirm", d(9, 9), "Китай", "sea", ""),
    ("S0020017", TPE, "StepUp", "Соколова И.Ю.", "в работе", d(22, 10), "Тайвань", "sea", ""),
    ("S0020018", HAN, "StepUp", "Соколова И.Ю.", "в работе", d(24, 10), "Вьетнам", "sea", ""),
    ("S0020019", IZM, "NordLine", "Петренко С.", "confirm", d(3, 9), "Турция", "truck", ""),
    ("S0020020", IZM, "NordLine", "Петренко С.", "confirm", d(30, 10), "Турция", "truck", ""),
    ("S0020021", IZM, "NordLine", "Петренко С.", "confirm", d(17, 9), "Турция", "truck", ""),
    ("S0020022", ANK, "Bellamare", "Белова И.", "в работе", d(26, 10), "Турция", "truck", ""),
    ("S0020023", ANK, "Bellamare", "Белова И.", "confirm", d(11, 9), "Турция", "truck", ""),
    ("S0020024", HAN, "StepUp", "Белова И.Н.", "в работе", d(31, 10), "Вьетнам", "sea", ""),
    ("S0020025", ANK, "Bellamare", "Белова И.Н.", "в работе", d(2, 11), "Турция", "truck", ""),
    ("S0020026", HAN, "StepUp", "Белова И.Н.", "в работе", d(3, 11), "Вьетнам", "sea", ""),
    ("S0020027", KAZ, "Lumen", "Громов А.Д.", "confirm", d(8, 9), "Россия", "truck", ""),
    ("S0020028", HAN, "StepUp", "Громов А.Д.", "в работе", d(19, 10), "Вьетнам", "sea", ""),
    ("S0020029", HAN, "StepUp", "Громов А.", "в работе", d(23, 10), "Вьетнам", "sea", ""),
    ("S0020030", HAN, "StepUp", "Громов А.Д.", "в работе", d(28, 10), "Вьетнам", "sea", ""),
    ("S0020031", HAN, "StepUp", "Соколова И.Ю.", "confirm", d(1, 9), "Вьетнам", "sea", ""),
    ("S0020032", PAR, "Lumen", "Ершова Д.О.", "confirm", d(2, 9), "Франция", "air", ""),
    ("S0020033", "Ningbo Ocean Knit Co., Ltd", "UrbanFox", "Лаптева Е.В.", "confirm", d(13, 9), "Китай", "sea", ""),
]

BEL = "Белов Игорь Николаевич"
ORL = "Орлова Анна Сергеевна"
BLV = "Белова Ирина Николаевна"

FUL_C = [
    ful("УЛ-2001", dt(4, 11, 9, 0), "S0021001", SEO, "UrbanFox", "air", None, ("Новый", GA, BEL, ""),
        trap="Нет страны, но air", must="Группа по air, логист назначается (Белов — предпочтительный), страна «не определено», «Требует внимания»"),
    ful("УЛ-2002", dt(4, 11, 9, 30), "S0020012", "Milano Moda Srl", "Bellamare", "air", "Италия",
        ("Повторная отправка", DASH, DASH, R % "04.11.2026"), comment="Номер: С0020012 от 28.10.2026",
        trap="Номер заказа с кириллической «С»", must="Распознать S0020012 → повторная отправка"),
    ful("УЛ-1990", dt(30, 10, 11, 0), "S0020020", IZM, "NordLine", "truck", "Турция",
        ("Уже учтён в Cloud", DASH, DASH, ""), trap="Исходный ФУЛ, уже перенесённый в Cloud (та же дата УЛ)",
        must="Не считать повторной отправкой"),
    ful("УЛ-2004", dt(4, 11, 10, 0), "S0021004", KAZ, "Lumen", "truck", "Россия", ("Не обработан", DASH, DASH, ""),
        status="Помечен на удаление", trap="ФУЛ помечен на удаление в 1С", must="Строку не создавать, «Требует внимания»"),
    ful("УЛ-2005", dt(4, 11, 10, 30), "S0021005", "Lyon Soie SARL", "Lumen", "air", "Франция",
        ("Новый", GA, ORL, ""), logist_1c="Белов И.Н.",
        trap="В 1С у ФУЛ указан логист", must="Логист из 1С не используется; новый поставщик → наименее загруженный: Орлова (3) < Белов (6)"),
    ful("УЛ-2006", dt(4, 11, 11, 0), "S0021006", TOK, "NordLine", "air", "Япония", ("Новый", GA, BEL, ""),
        trap="Ничья в истории: Орлова 2, Белов 2", must="Из равных по истории — наименее загруженный: Белов (2) < Орлова (3)"),
    ful("УЛ-2007", dt(4, 11, 12, 0), "S0021007", SHC, "Lumen", "sea", "Китай", ("Новый", GV, NO, ""),
        trap="ВЭД: уволен, отпуск, ассистент и руководитель «Работает»", must="Логист пустой; ассистента и руководителя не назначать"),
    ful("УЛ-2008", dt(4, 11, 13, 0), None, "Hanoi Garment JSC", "StepUp", "sea", "Вьетнам",
        ("Не обработан", DASH, DASH, ""), comment="Номер: S0021008 от 30.10.2026, S0021009 от 30.10.2026",
        trap="Два номера заказа в комментарии", must="Заказ «не определено», строку не создавать"),
    ful("УЛ-2009", dt(4, 11, 14, 0), "S0021010", KAZ, "Lumen", "courier", "Россия", ("Новый", U, NO, ""),
        trap="Неизвестный способ доставки «courier»", must="Группа «не определено», логист пустой, способ доставки — как в 1С"),
    ful("УЛ-2010", dt(5, 11, 9, 0), "S0021011", TPE, "StepUp", "sea", "Тайвань", ("Новый", GM, SOK, ""),
        trap="Тайвань", must="По принятому допущению не Китай → мультимодальные; предпочтительный Соколова"),
    ful("УЛ-2011", dt(5, 11, 9, 30), "S0021012", IZM, "NordLine", "truck", "Турция", ("Новый", GM, GRO, ""),
        trap="История только у логиста не из списка сотрудников (Петренко С.)",
        must="Считать новым для группы → наименее загруженный: ничья Соколова 3 = Громов 3 → по алфавиту Громов"),
    ful("УЛ-2012", dt(5, 11, 10, 0), "S0021013", "ANKARA WOOL LTD.", "Bellamare", "truck", "Турция",
        ("Новый", GM, BLV, ""), exp_supplier=ANK, trap="«Белова И.» в Cloud при Белове И.Н. в 1С",
        must="Сопоставить с Беловой Ириной Николаевной (мультимодальные), не с Беловым"),
    ful("УЛ-2013", dt(5, 11, 11, 0), "S0021014", KAZ, "Lumen", "Direct Truck ", "Россия", ("Новый", GM, GRO, ""),
        trap="Способ доставки с заглавными и пробелом", must="direct truck; предпочтительный Громов"),
    ful("УЛ-2015", dt(6, 11, 8, 0), "S0021015", BUR, "NordLine", "truck", "Турция",
        ("Дубль в пачке", DASH, DASH, R % "06.11.2026"), trap="Дубль стоит в файле раньше исходного ФУЛ",
        must="Первым считается ФУЛ с более ранней датой (УЛ-2014), этот — дубль"),
    ful("УЛ-2014", dt(5, 11, 12, 0), "S0021015", BUR, "NordLine", "truck", "Турция",
        ("Новый", GM, SOK, R % "06.11.2026"), trap="Исходный ФУЛ для дубля", must="Одна строка, дата УЛ 05.11; наименее загруженный Соколова (3)"),
    ful("УЛ-2016", dt(5, 11, 13, 0), "S0020025", IZM, "NordLine", "truck", "Турция",
        ("Повторная отправка", DASH, DASH, R % "05.11.2026"), trap="Заказ в Cloud у другого поставщика",
        must="Повтор по номеру + «Требует внимания»: поставщик не совпадает"),
    ful("УЛ-2017", dt(5, 11, 14, 0), "S0021017", None, "Bellamare", "truck", "Турция", ("Новый", GM, NO, ""),
        trap="Нет поставщика", must="Поставщик «не определено», логист пустой"),
    ful("УЛ-2018", dt(6, 11, 9, 0), "S0021018", SEO, None, "air", "Республика Корея", ("Новый", GA, BEL, ""),
        cutoff="-", trap="Нет бренда и cut-off", must="Бренд и cut-off «не определено», «Требует внимания»; логист тот же, что у УЛ-2001"),
    ful("УЛ-2019", dt(3, 11, 10, 0), "S0021019", TOK, "NordLine", "air", "Япония", ("Новый", GA, BEL, ""),
        trap="Дата ФУЛ равна последней дате Cloud (03.11)", must="Не пропуск (не раньше), а новый; тот же логист, что у УЛ-2006"),
    ful("УЛ-1985", dt(1, 11, 16, 0), "S0021020", TPE, "StepUp", "sea", "Тайвань", ("Возможный пропуск", GM, SOK, ""),
        trap="Пропуск", must="В пропуски, предложен тот же логист, что у УЛ-2010"),
]


# ============================================================================ запись файлов

def short(fio):
    p = fio.split()
    return p[0] + " " + "".join(x[0] + "." for x in p[1:3])


METHOD_CANON = {"air": "air", "sea": "sea", "truck": "truck", "railway": "railway", "direct truck": "direct truck"}


def canon_method(m):
    if m is None:
        return U
    k = " ".join(m.lower().split())
    return METHOD_CANON.get(k, m.strip())


def cloud_rows_full(cloud):
    out = []
    for order, sup, brand, logist, status, ul, country, method, comment in cloud:
        confirm = (ul + timedelta(days=3)) if status == "confirm" else None
        docs = (ul + timedelta(days=10)) if status == "confirm" else None
        out.append(dict(order=order, supplier=sup, brand=brand, logist=logist, status=status, confirm=confirm,
                        cutoff=ul + timedelta(days=20), ul=ul, country=country, method=method, docs=docs, comment=comment))
    return out


CLOUD_COLS_STD = [("Заказ", "order"), ("Поставщик", "supplier"), ("Бренд", "brand"), ("Логист", "logist"),
                  ("статус", "status"), ("дата подтверждения", "confirm"), ("cut-off", "cutoff"),
                  ("Дата получения УЛ", "ul"), ("Страна отправления", "country"), ("Способ доставки", "method"),
                  ("Документы получены", "docs"), ("Комментарий", "comment")]
FUL_COLS_STD = [("Дата", "when"), ("Cut-off date", "cutoff"), ("Номер", "num"), ("Статус", "status"),
                ("Комментарий", "comment"), ("Способ доставки", "method"), ("Бизнес единица", "brand"),
                ("Логист", "logist_1c"), ("Контрагент", "sup"), ("Страна", "country")]
STAFF_COLS_STD = [("Наименование", 0), ("Текущее состояние", 1), ("Должность", 2), ("Код", 3)]


def write_sheet(ws, cols, rows, title_rows=(), date_as_text=()):
    for t in title_rows:
        ws.append(list(t))
    ws.append([c[0] for c in cols])
    hdr = ws.max_row
    for c in range(1, len(cols) + 1):
        ws.cell(hdr, c).font = Font(bold=True)
    for r in rows:
        vals = []
        for title, key in cols:
            v = r.get(key) if isinstance(r, dict) else (r[key] if key is not None else None)
            if key in date_as_text and isinstance(v, (date, datetime)):
                v = v.strftime("%d.%m.%Y %H:%M:%S") if isinstance(v, datetime) else v.strftime("%d.%m.%Y")
            vals.append(v)
        ws.append(vals)
    for r in range(hdr + 1, ws.max_row + 1):
        for c in range(1, len(cols) + 1):
            cell = ws.cell(r, c)
            if isinstance(cell.value, datetime):
                cell.number_format = "DD.MM.YYYY HH:MM:SS"
            elif isinstance(cell.value, date):
                cell.number_format = "DD.MM.YYYY"
    for i in range(1, len(cols) + 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 22


def write_dataset(out, name, staff, cloud, fuls, variant="std"):
    folder = os.path.join(out, name)
    inp = os.path.join(folder, "входные_файлы")
    os.makedirs(inp, exist_ok=True)
    cloud_full = cloud_rows_full(cloud)
    if variant == "std":
        cloud_cols, ful_cols = CLOUD_COLS_STD, FUL_COLS_STD
        wb = openpyxl.Workbook(); write_sheet(wb.active, ful_cols, fuls); wb.active.title = "ФУЛ"
        wb.save(os.path.join(inp, "Выгрузка_ФУЛ_1С.xlsx"))
        wb = openpyxl.Workbook(); write_sheet(wb.active, cloud_cols, cloud_full); wb.active.title = "Cloud"
        wb.save(os.path.join(inp, "Реестр_Cloud.xlsx"))
        wb = openpyxl.Workbook(); write_sheet(wb.active, STAFF_COLS_STD, staff); wb.active.title = "Сотрудники"
        wb.save(os.path.join(inp, "Сотрудники_1С.xlsx"))
    else:
        # другой порядок, лишние колонки, другие имена листов, заголовок не в первой строке, даты строкой
        ful_cols = [("Номер", "num"), ("Контрагент ", "sup"), ("Страна", "country"), ("Способ  доставки", "method"),
                    ("Дата", "when"), ("Комментарий", "comment"), ("Статус", "status"), ("Бизнес-единица", "brand"),
                    ("Склад", "_wh"), ("Cut off date", "cutoff"), ("Логист", "logist_1c"), ("Ответственный", "_resp")]
        rows = [dict(f, _wh="Подольск", _resp="Иванов И.И.") for f in fuls]
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Выгрузка 1С"
        write_sheet(ws, ful_cols, rows, title_rows=[("Упаковочные листы за период 30.10.2026 – 06.11.2026",), ()],
                    date_as_text=("when",))
        wb.save(os.path.join(inp, "export_1.xlsx"))
        cloud_cols = [("Заказ", "order"), ("Логист", "logist"), ("Поставщик", "supplier"), ("Бренд", "brand"),
                      ("Статус", "status"), ("Cut-off", "cutoff"), ("Дата подтверждения", "confirm"),
                      ("Страна отправления", "country"), ("Способ доставки", "method"), ("Дата получения УЛ", "ul"),
                      ("Документы получены", "docs"), ("Комментарий", "comment"), ("Менеджер бренда", "_mgr")]
        wb = openpyxl.Workbook(); ws0 = wb.active; ws0.title = "Инструкция"
        ws0.append(["Реестр заказов. Заполняйте лист «Реестр 2026»."])
        ws = wb.create_sheet("Реестр 2026")
        write_sheet(ws, cloud_cols, [dict(r, _mgr="Смирнов А.") for r in cloud_full])
        wb.save(os.path.join(inp, "export_2.xlsx"))
        import xlwt  # только для генерации .xls; агенту xlwt не нужен
        book = xlwt.Workbook(encoding="utf-8")
        sh = book.add_sheet("TDSheet")
        heads = ["Код (Таб. №)", "Наименование", "Подразделение", "Должность", "Текущее состояние"]
        for j, h in enumerate(heads):
            sh.write(0, j, h)
        for i, (fio, st, pos, code) in enumerate(staff, 1):
            for j, v in enumerate([code, fio, "Отдел логистики", pos, st]):
                sh.write(i, j, v)
        book.save(os.path.join(inp, "export_3.xls"))

    # ----- эталон
    form_titles = [t for t, k in cloud_cols if not k.startswith("_")]
    form_keys = [k for t, k in cloud_cols if not k.startswith("_")]
    exp_fuls, new_rows, misses = [], [], []
    for f in fuls:
        lg = f["logist"]
        e = dict(ful=f["num"], date=f["when"].strftime("%d.%m.%Y"), order=f["order"] or U, kind=f["kind"],
                 group=f["group"], logist_fio=(lg if lg else U) if f["kind"] in ("Новый", "Возможный пропуск") else DASH,
                 comment=f["cloud_comment"], trap=f["trap"], must=f["must"])
        exp_fuls.append(e)
        if f["kind"] in ("Новый", "Возможный пропуск"):
            vals = dict(order=f["order"], supplier=f["exp_supplier"] or (" ".join(f["sup"].split()) if f["sup"] else U),
                        brand=f["brand"] or U, logist=short(lg) if lg else "", status="в работе", confirm="",
                        cutoff=f["cutoff"].strftime("%d.%m.%Y") if f["cutoff"] else U,
                        ul=f["when"].strftime("%d.%m.%Y"), country=f["country"] or U, method=canon_method(f["method"]),
                        docs="", comment=f["cloud_comment"])
            row = {t: vals[k] for t, k in zip(form_titles, form_keys)}
            row["_ful"] = f["num"]
            (new_rows if f["kind"] == "Новый" else misses).append(row)
    # undefined-ловушки: какие поля обязаны остаться «не определено»/пустыми
    traps = []
    for f in fuls:
        fields = {}
        if f["order"] is None:
            fields["order"] = U
        if f["kind"] in ("Новый", "Возможный пропуск"):
            if f["group"] == U:
                fields["group"] = U
            if not f["logist"]:
                fields["logist_fio"] = U
        if fields:
            traps.append(dict(ful=f["num"], fields=fields))
    expected = dict(dataset=name, cloud_form=form_titles, cloud_orders=[c[0] for c in cloud], fuls=exp_fuls, new_rows=new_rows, possible_misses=misses,
                    undefined_traps=traps)
    with open(os.path.join(folder, "эталон.json"), "w", encoding="utf-8") as fh:
        json.dump(expected, fh, ensure_ascii=False, indent=1)

    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "Эталон по ФУЛ"
    heads = ["Номер ФУЛ", "Дата ФУЛ", "Заказ", "Решение", "Группа", "Логист (ФИО)", "Комментарий для Cloud",
             "Ловушка", "Что агент обязан сделать"]
    ws.append(heads)
    for e in exp_fuls:
        ws.append([e["ful"], e["date"], e["order"], e["kind"], e["group"], e["logist_fio"], e["comment"], e["trap"], e["must"]])
    ws2 = wb.create_sheet("Эталон новых строк Cloud")
    ws2.append(form_titles + ["Номер ФУЛ"])
    for r in new_rows:
        ws2.append([r[t] for t in form_titles] + [r["_ful"]])
    ws3 = wb.create_sheet("Эталон пропусков")
    ws3.append(form_titles + ["Номер ФУЛ"])
    for r in misses:
        ws3.append([r[t] for t in form_titles] + [r["_ful"]])
    for w in (ws, ws2, ws3):
        for c in w[1]:
            c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="1F4E78")
            c.alignment = Alignment(wrap_text=True)
        for i in range(1, w.max_column + 1):
            w.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 20
    ws.column_dimensions["H"].width = 45; ws.column_dimensions["I"].width = 70
    wb.save(os.path.join(folder, "Эталон.xlsx"))
    return expected


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    for name, staff, cloud, fuls, var in [("A_основной", STAFF_A, CLOUD_A, FUL_A, "std"),
                                          ("B_формат", STAFF_A, CLOUD_A, FUL_A, "shuffled"),
                                          ("C_крайние", STAFF_C, CLOUD_C, FUL_C, "std")]:
        e = write_dataset(a.out, name, staff, cloud, fuls, var)
        print("%s: ФУЛ %d, Cloud %d, сотрудников %d, новых строк %d" % (name, len(fuls), len(cloud), len(staff), len(e["new_rows"])))


if __name__ == "__main__":
    main()
