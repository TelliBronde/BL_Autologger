import os
import re
import subprocess
import sys
import tempfile
import threading
import time
from collections import defaultdict
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

import config
import connector
import parser

try:
    import gsheets
except ImportError:
    gsheets = None

HF = "[h]:mm"
C_HEAD = "1F3864"
C_SUB  = "2F5597"
C_DAY  = "D9E1F2"
C_OK   = "C6EFCE"
C_REP  = "FFF2CC"
C_FAIL = "F8CBAD"
C_WARN = "FCE4D6"
C_GREY = "F2F2F2"

ILLEGAL_XML_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

if not hasattr(sys, "_bambu_live_state"):
    sys._bambu_live_state = {
        "active": False,
        "interval": 60,
        "last_update": "—",
        "file": "Farm_Report_LIVE.xlsx",
        "status_msg": "",
        "done_cnt": 0,
        "fail_cnt": 0,
        "last_cloud_sync": "—",
        "cloud_status": "Выкл",
    }


def get_live_state():
    return sys._bambu_live_state


def sanitize_excel_value(v, allow_formula=False):
    if not isinstance(v, str):
        return v
    s = ILLEGAL_XML_RE.sub("", v)
    if not allow_formula and s.startswith(("=", "+", "-", "@", "\t", "\r")):
        return "'" + s
    return s


def put(ws, ref, v, bold=False, fmt=None, fill=None, color=None, align=None, size=10, wrap=False, formula=False):
    c = ws[ref]
    safe_v = sanitize_excel_value(v, allow_formula=formula)
    c.value = safe_v
    if isinstance(safe_v, str) and not formula:
        c.data_type = "s"
    c.font = Font(name="Arial", size=size, bold=bold, color=color)
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = PatternFill("solid", fgColor=fill)
    if align or wrap:
        c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    return c


def header(ws, row, titles, fill_color=C_HEAD):
    thin = Side(style="thin", color="BFBFBF")
    for i, t in enumerate(titles):
        c = put(ws, f"{get_column_letter(i + 1)}{row}", t, bold=True, fill=fill_color, color="FFFFFF",
                align="center", wrap=True)
        c.border = Border(top=thin, bottom=thin, left=thin, right=thin)
    ws.row_dimensions[row].height = 32


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths):
        ws.column_dimensions[get_column_letter(i + 1)].width = w


def parse_user_date(raw_str: str, default_date: str) -> str:
    s = (raw_str or "").strip()
    if not s:
        return default_date
    if re.match(r"^\d{4}-\d{2}-\d{2}$", s):
        return s
    m_full = re.match(r"^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$", s)
    if m_full:
        d, m, y = int(m_full.group(1)), int(m_full.group(2)), int(m_full.group(3))
        return f"{y:04d}-{m:02d}-{d:02d}"
    m_short = re.match(r"^(\d{1,2})[./-](\d{1,2})$", s)
    if m_short:
        d, m = int(m_short.group(1)), int(m_short.group(2))
        y = datetime.now().year
        return f"{y:04d}-{m:02d}-{d:02d}"
    return default_date


VBS_REFRESH_TEMPLATE = """
On Error Resume Next
If WScript.Arguments.Count < 3 Then WScript.Quit 1
Dim targetName, tempPath, targetPath
targetName = WScript.Arguments(0)
tempPath = WScript.Arguments(1)
targetPath = WScript.Arguments(2)

Dim xl, wb, activeSheetName, activeCellAddr, found
found = False
activeSheetName = ""
activeCellAddr = "A1"

Set xl = GetObject(, "Excel.Application")
If Err.Number = 0 And Not (xl Is Nothing) Then
    For Each wb In xl.Workbooks
        If LCase(wb.Name) = LCase(targetName) Then
            found = True
            activeSheetName = wb.ActiveSheet.Name
            activeCellAddr = xl.Selection.Address
            xl.DisplayAlerts = False
            wb.Close False
            If Err.Number <> 0 Then WScript.Quit 2
            Exit For
        End If
    Next
End If
Err.Clear

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(targetPath) Then fso.DeleteFile targetPath, True
If Err.Number <> 0 Then WScript.Quit 3
fso.MoveFile tempPath, targetPath
If Err.Number <> 0 Then WScript.Quit 4

If found And Not (xl Is Nothing) Then
    Dim newWb
    Set newWb = xl.Workbooks.Open(targetPath)
    If activeSheetName <> "" Then
        newWb.Worksheets(activeSheetName).Activate
        xl.Range(activeCellAddr).Select
    End If
    xl.DisplayAlerts = True
End If
WScript.Quit 0
"""


def refresh_open_excel_via_vbs(temp_path: Path, target_path: Path) -> bool:
    if sys.platform != "win32":
        return False
    vbs_path = None
    try:
        with tempfile.NamedTemporaryFile("w", suffix=".vbs", delete=False, encoding="cp1251", errors="ignore") as tf:
            tf.write(VBS_REFRESH_TEMPLATE)
            vbs_path = Path(tf.name)
        res = subprocess.run(
            ["cscript", "//NoLogo", str(vbs_path), target_path.name, str(temp_path), str(target_path)],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=6,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        return res.returncode == 0
    except Exception:
        return False
    finally:
        if vbs_path and vbs_path.exists():
            try:
                vbs_path.unlink()
            except Exception:
                pass


def safe_save_and_open(wb: Workbook, export_path: Path, open_file=True, is_live_tick=False):
    temp_path = export_path.with_suffix(f".{os.getpid()}.tmp.xlsx")
    wb.save(temp_path)

    try:
        if export_path.exists():
            os.replace(temp_path, export_path)
        else:
            temp_path.rename(export_path)
        if open_file and sys.platform == "win32":
            os.startfile(export_path)
        return export_path, True, "ОК"
    except PermissionError:
        pass

    if refresh_open_excel_via_vbs(temp_path, export_path):
        return export_path, True, "ОК"

    try:
        if temp_path.exists():
            temp_path.unlink()
    except Exception:
        pass

    if is_live_tick:
        return export_path, True, "Excel занят вводом в ячейку"

    ts = datetime.now().strftime("%H-%M-%S")
    alt_path = export_path.with_name(f"{export_path.stem}_{ts}{export_path.suffix}")
    wb.save(alt_path)
    if open_file and sys.platform == "win32":
        os.startfile(alt_path)
    return alt_path, True, "Сохранено в копию"


def build_period_printers_stat(filtered_parsed, full_printers_stat):
    p_stat = defaultdict(lambda: {
        "tasks": 0, "copies": 0.0, "stopped_cnt": 0.0,
        "weight_good": 0.0, "weight_waste": 0.0,
        "secs_good": 0.0, "secs_waste": 0.0,
        "secs_after_sync": 0.0,
    })
    for pname, full_st in full_printers_stat.items():
        p_stat[pname]["secs_after_sync"] = full_st.get("secs_after_sync", 0.0)
        p_stat[pname]["total_all_time_secs"] = full_st.get("secs_good", 0.0) + full_st.get("secs_waste", 0.0)

    for x in filtered_parsed:
        devs = [d.strip() for d in (x.get("printers") or "").split(",") if d.strip() and d.strip() != "—"]
        if not devs:
            continue
        n = len(devs)
        for dn in devs:
            if x.get("is_failed"):
                p_stat[dn]["stopped_cnt"] += x["qty"] / n
                p_stat[dn]["weight_waste"] += x["weight_waste"] / n
                p_stat[dn]["secs_waste"] += x["tot_secs"] / n
            else:
                p_stat[dn]["tasks"] += 1
                p_stat[dn]["copies"] += x["qty"] / n
                p_stat[dn]["weight_good"] += x["weight_good"] / n
                p_stat[dn]["secs_good"] += x["tot_secs"] / n
    return p_stat


def generate_report(date_from=None, date_to=None, custom_filename=None, open_file=True, is_live_tick=False, quiet=False):
    with config.get_lock():
        cfg = config.load_config()
        loc_name = cfg.get("location_name") or "Farm"
        loc_tag = config.sanitize_filename_tag(loc_name)
        start_date_cfg = (cfg.get("start_date") or "").strip()
        to_interval = float(cfg.get("to_interval_hours") or 300)

        parsed_all, full_printers_stat, devices_hw = parser.load_and_parse(quiet=(quiet or is_live_tick))
        odometers = config.load_odometers()
        order_rules = config.load_orders_map()

        is_period = bool(date_from or date_to)
        if is_period:
            d_min = date_from or "0000-00-00"
            d_max = date_to or "9999-99-99"
            parsed = [x for x in (parsed_all or []) if d_min <= x["date"] <= d_max]
            period_label = f"период {d_min} — {d_max}"
        else:
            parsed = list(parsed_all or [])
            scope_txt = f"с {start_date_cfg}" if start_date_cfg else "вся история"
            period_label = f"{scope_txt} (обн. {datetime.now():%H:%M:%S})"

        printers_stat = build_period_printers_stat(parsed, full_printers_stat) if is_period else (full_printers_stat or {})

        done_cnt = sum(x["qty"] for x in parsed if not x.get("is_failed"))
        fail_cnt = sum(x["qty"] for x in parsed if x.get("is_failed"))
        sys._bambu_live_state["done_cnt"] = done_cnt
        sys._bambu_live_state["fail_cnt"] = fail_cnt

        safe_custom = Path(custom_filename).name if custom_filename else None
        if safe_custom:
            export_path = Path.home() / "Desktop" / safe_custom
        elif is_period:
            export_path = Path.home() / "Desktop" / f"Farm_Report_{loc_tag}_{d_min}_to_{d_max}.xlsx"
        else:
            export_path = Path.home() / "Desktop" / f"Farm_Report_{loc_tag}_{datetime.now():%Y-%m-%d}.xlsx"

        wb = Workbook()
        sh_sum = wb.active
        sh_sum.title = "Общая сводка"
        sh_map = wb.create_sheet("Справочник заказов")
        sh_all = wb.create_sheet("Все задачи")
        sh_df  = wb.create_sheet("По дням и файлам")
        sh_dyn = wb.create_sheet("Динамика по дням")
        sh_pr  = wb.create_sheet("Принтеры и Моточасы")

        thin = Side(style="thin", color="D9D9D9")
        med_bottom = Side(style="medium", color="8EA9DB")
        box = Border(top=thin, bottom=thin, left=thin, right=thin)
        day_box = Border(top=med_bottom, bottom=med_bottom, left=thin, right=thin)

        unique_files = sorted({x["file"] for x in parsed})
        map_end = max(500, 4 + len(unique_files) + 100)
        last_all = max(4, 3 + len(parsed))

        # 2. СПРАВОЧНИК ЗАКАЗОВ
        put(sh_map, "A1", f"Справочник заказов — {loc_name}", bold=True, size=14)
        put(
            sh_map, "A2",
            "Чтобы объединить несколько файлов в один заказ: впиши одинаковое название в столбец B и (при необходимости) укажи кол-во деталей на столе в столбце C. Лист «Общая сводка» пересчитается автоматически!",
            color="595959", size=9,
        )
        header(sh_map, 3, ["Имя файла (задача)", "Заказ", "Деталей на столе, шт"])

        discovered_orders = []
        seen_orders = set()
        for r in order_rules:
            o = r["order"]
            if o not in seen_orders:
                discovered_orders.append(o)
                seen_orders.add(o)

        for idx, fname in enumerate(unique_files):
            r_idx = 4 + idx
            ord_name, ppp = parser.match_order_rule(fname, order_rules)
            if ord_name != "Не распределено" and ord_name not in seen_orders:
                discovered_orders.append(ord_name)
                seen_orders.add(ord_name)
            put(sh_map, f"A{r_idx}", fname)
            put(sh_map, f"B{r_idx}", ord_name, fill=C_REP if ord_name == "Не распределено" else C_OK)
            put(sh_map, f"C{r_idx}", ppp, align="center", fmt="0")
            for col in "ABC":
                sh_map[f"{col}{r_idx}"].border = box

        sh_map.auto_filter.ref = f"A3:C{max(4, 3 + len(unique_files))}"
        sh_map.freeze_panes = "A4"
        widths(sh_map, [46, 28, 22])

        # 3. ВСЕ ЗАДАЧИ
        put(sh_all, "A1", f"Журнал печатей и срывов — {loc_name} ({len(parsed)} записей | {period_label})", bold=True, size=14)
        header(sh_all, 3, [
            "Локация", "Дата", "Время", "Тип", "Статус / Прогресс",
            "Заказ", "Файл (задача)", "Принтер",
            "Столов, шт", "Деталей на столе", "Готовых деталей, шт",
            "Вес готовых, г", "Вес срывов, г", "Время печати"
        ])

        for i, t in enumerate(parsed):
            r_idx = 4 + i
            st_fill = C_FAIL if t.get("is_failed") else (C_REP if t.get("is_reprint") else C_OK)
            f_order = (
                f'=IFERROR(IF(INDEX(\'Справочник заказов\'!$B$4:$B${map_end},'
                f'MATCH(G{r_idx},\'Справочник заказов\'!$A$4:$A${map_end},0))="","Не распределено",'
                f'INDEX(\'Справочник заказов\'!$B$4:$B${map_end},'
                f'MATCH(G{r_idx},\'Справочник заказов\'!$A$4:$A${map_end},0))),"Не распределено")'
            )
            f_ppp = (
                f'=IFERROR(IF(INDEX(\'Справочник заказов\'!$C$4:$C${map_end},'
                f'MATCH(G{r_idx},\'Справочник заказов\'!$A$4:$A${map_end},0))="",1,'
                f'INDEX(\'Справочник заказов\'!$C$4:$C${map_end},'
                f'MATCH(G{r_idx},\'Справочник заказов\'!$A$4:$A${map_end},0))),1)'
            )
            f_parts = f'=IF(D{r_idx}="Готово",I{r_idx}*J{r_idx},0)'

            put(sh_all, f"A{r_idx}", loc_name, align="center")
            put(sh_all, f"B{r_idx}", t["date"], align="center")
            put(sh_all, f"C{r_idx}", t["time"], align="center")
            put(sh_all, f"D{r_idx}", t["type"], bold=True, fill=st_fill, align="center")
            put(sh_all, f"E{r_idx}", t["status"], fill=st_fill, align="center")
            put(sh_all, f"F{r_idx}", f_order, formula=True)
            put(sh_all, f"G{r_idx}", t["file"])
            put(sh_all, f"H{r_idx}", t["printers"])
            put(sh_all, f"I{r_idx}", t["qty"], bold=True, align="center", fmt="0")
            put(sh_all, f"J{r_idx}", f_ppp, align="center", fmt="0", formula=True)
            put(sh_all, f"K{r_idx}", f_parts, bold=True, align="center", fmt="0", formula=True)
            put(sh_all, f"L{r_idx}", round(t["weight_good"], 1), fmt="0.0")
            put(sh_all, f"M{r_idx}", round(t["weight_waste"], 1), fmt="0.0", color="9C0006" if t["weight_waste"] > 0 else None)
            put(sh_all, f"N{r_idx}", t["tot_secs"] / 86400 if t["tot_secs"] else 0, fmt=HF, align="center")
            for col in "ABCDEFGHIJKLMN":
                sh_all[f"{col}{r_idx}"].border = box

        sh_all.auto_filter.ref = f"A3:N{last_all}"
        sh_all.freeze_panes = "A4"
        widths(sh_all, [14, 12, 8, 11, 20, 22, 40, 20, 11, 15, 17, 15, 14, 14])

        # 1. ОБЩАЯ СВОДКА
        put(sh_sum, "A1", f"Общая сводка — {loc_name} ({period_label})", bold=True, size=15)
        put(sh_sum, "A2", f"Обновлено: {datetime.now():%Y-%m-%d %H:%M:%S} | Порог готово/срыв: {config.SUCCESS_THRESHOLD_PCT}%", color="595959")

        header(sh_sum, 4, ["Ключевой показатель локации", "Значение"], fill_color=C_HEAD)
        kpis = [
            ("Локация", loc_name, None, False, None, False),
            ("Общее время печати (всего)", f"=SUM('Все задачи'!$N$4:$N${last_all})", HF, True, C_DAY, True),
            ("Завершённых столов, шт", f"=SUMIFS('Все задачи'!$I$4:$I${last_all},'Все задачи'!$D$4:$D${last_all},\"Готово\")", "#,##0", True, C_OK, True),
            ("Изготовлено деталей, шт", f"=SUM('Все задачи'!$K$4:$K${last_all})", "#,##0", True, C_OK, True),
            ("Потрачено пластика (готовые), кг", f"=SUM('Все задачи'!$L$4:$L${last_all})/1000", "0.00", True, C_OK, True),
            ("Сорванных столов (брак), шт", f"=SUMIFS('Все задачи'!$I$4:$I${last_all},'Все задачи'!$D$4:$D${last_all},\"Срыв\")", "#,##0", True, C_FAIL, True),
            ("Вес сорванных печатей (брак), кг", f"=SUM('Все задачи'!$M$4:$M${last_all})/1000", "0.00", True, C_FAIL, True),
            ("Итого потрачено пластика, кг", "=B9+B11", "0.00", True, C_DAY, True),
            ("% брака по весу", "=IF(B12>0,B11/B12,0)", "0.0%", True, C_WARN, True),
        ]
        for idx, (label, val, num_fmt, is_bold, fill_c, is_f) in enumerate(kpis):
            r_k = 5 + idx
            put(sh_sum, f"A{r_k}", label, bold=is_bold, fill=C_GREY if idx % 2 else None)
            put(sh_sum, f"B{r_k}", val, bold=is_bold, fmt=num_fmt, fill=fill_c, align="center", formula=is_f)
            sh_sum[f"A{r_k}"].border = box
            sh_sum[f"B{r_k}"].border = box
            sh_sum.row_dimensions[r_k].height = 20

        put(sh_sum, "A15", "Сводка по заказам (добавь имя заказа на листе «Справочник заказов» и впиши его в пустую строку ниже)", bold=True, size=12)
        header(sh_sum, 16, [
            "Заказ", "Файлов в заказе", "Выполнено столов, шт", "Готовых деталей, шт",
            "Вес готовых, кг", "Сорвано столов, шт", "Вес срывов, кг",
            "Всего пластик, кг", "% брака (вес)", "Время печати"
        ], fill_color=C_SUB)

        order_rows_list = list(discovered_orders) + ["Не распределено"] + [""] * 25
        ord_start = 17
        ord_end = ord_start + len(order_rows_list) - 1

        for idx, ord_label in enumerate(order_rows_list):
            r_o = ord_start + idx
            is_unassigned = (ord_label == "Не распределено")
            row_fill = C_REP if is_unassigned else None

            put(sh_sum, f"A{r_o}", ord_label, bold=bool(ord_label), fill=row_fill)
            put(sh_sum, f"B{r_o}", f'=IF(A{r_o}="","",COUNTIF(\'Справочник заказов\'!$B$4:$B${map_end},A{r_o}))', align="center", fmt="0", fill=row_fill, formula=True)
            put(sh_sum, f"C{r_o}", f'=IF(A{r_o}="","",SUMIFS(\'Все задачи\'!$I$4:$I${last_all},\'Все задачи\'!$F$4:$F${last_all},A{r_o},\'Все задачи\'!$D$4:$D${last_all},"Готово"))', bold=True, align="center", fmt="0", fill=row_fill, formula=True)
            put(sh_sum, f"D{r_o}", f'=IF(A{r_o}="","",SUMIFS(\'Все задачи\'!$K$4:$K${last_all},\'Все задачи\'!$F$4:$F${last_all},A{r_o}))', bold=True, align="center", fmt="0", fill=row_fill, formula=True)
            put(sh_sum, f"E{r_o}", f'=IF(A{r_o}="","",SUMIFS(\'Все задачи\'!$L$4:$L${last_all},\'Все задачи\'!$F$4:$F${last_all},A{r_o})/1000)', fmt="0.00", fill=row_fill, formula=True)
            put(sh_sum, f"F{r_o}", f'=IF(A{r_o}="","",SUMIFS(\'Все задачи\'!$I$4:$I${last_all},\'Все задачи\'!$F$4:$F${last_all},A{r_o},\'Все задачи\'!$D$4:$D${last_all},"Срыв"))', align="center", fmt="0", fill=row_fill, formula=True)
            put(sh_sum, f"G{r_o}", f'=IF(A{r_o}="","",SUMIFS(\'Все задачи\'!$M$4:$M${last_all},\'Все задачи\'!$F$4:$F${last_all},A{r_o})/1000)', fmt="0.00", fill=row_fill, formula=True)
            put(sh_sum, f"H{r_o}", f'=IF(A{r_o}="","",E{r_o}+G{r_o})', bold=True, fmt="0.00", fill=row_fill, formula=True)
            put(sh_sum, f"I{r_o}", f'=IF(A{r_o}="","",IF(H{r_o}>0,G{r_o}/H{r_o},0))', align="center", fmt="0.0%", fill=row_fill, formula=True)
            put(sh_sum, f"J{r_o}", f'=IF(A{r_o}="","",SUMIFS(\'Все задачи\'!$N$4:$N${last_all},\'Все задачи\'!$F$4:$F${last_all},A{r_o}))', align="center", fmt=HF, fill=row_fill, formula=True)
            for col in "ABCDEFGHIJ":
                sh_sum[f"{col}{r_o}"].border = box

        ord_tot = ord_end + 1
        put(sh_sum, f"A{ord_tot}", "ИТОГО ПО ЗАКАЗАМ", bold=True, fill=C_GREY)
        for col, fmt_c in [("B", "0"), ("C", "0"), ("D", "0"), ("E", "0.00"), ("F", "0"), ("G", "0.00"), ("H", "0.00"), ("J", HF)]:
            put(sh_sum, f"{col}{ord_tot}", f"=SUM({col}{ord_start}:{col}{ord_end})", bold=True, fmt=fmt_c, fill=C_GREY, align="center" if col in "BCDFJ" else None, formula=True)
        put(sh_sum, f"I{ord_tot}", f"=IF(H{ord_tot}>0,G{ord_tot}/H{ord_tot},0)", bold=True, fmt="0.0%", fill=C_GREY, align="center", formula=True)
        for col in "ABCDEFGHIJ":
            sh_sum[f"{col}{ord_tot}"].border = day_box
        widths(sh_sum, [34, 18, 18, 18, 16, 17, 15, 17, 14, 15])

        # 4. ПО ДНЯМ И ФАЙЛАМ
        by_day = defaultdict(list)
        for x in parsed:
            by_day[x["date"]].append(x)
        sorted_days = sorted(by_day.keys(), reverse=True)

        put(sh_df, "A1", f"Разбивка по дням и файлам — {loc_name} ({period_label})", bold=True, size=14)
        header(sh_df, 4, [
            "Локация", "Дата", "Файл (задача)", "Готово, столов",
            "Пластик готовых, г", "Срывы, столов", "Вес срывов, г", "Общее время печати"
        ])
        rr = 5
        grand_done = grand_fail = grand_secs = 0
        grand_wg = grand_ww = 0.0

        for d in sorted_days:
            day_items = by_day[d]
            by_file = defaultdict(lambda: {"done": 0, "fail": 0, "wg": 0.0, "ww": 0.0, "secs": 0})
            for x in day_items:
                rec = by_file[x["file"]]
                if x.get("is_failed"):
                    rec["fail"] += x["qty"]
                    rec["ww"] += x["weight_waste"]
                else:
                    rec["done"] += x["qty"]
                    rec["wg"] += x["weight_good"]
                rec["secs"] += x["tot_secs"]

            d_done = sum(v["done"] for v in by_file.values())
            d_fail = sum(v["fail"] for v in by_file.values())
            d_wg = sum(v["wg"] for v in by_file.values())
            d_ww = sum(v["ww"] for v in by_file.values())
            d_secs = sum(v["secs"] for v in by_file.values())
            grand_done += d_done
            grand_fail += d_fail
            grand_wg += d_wg
            grand_ww += d_ww
            grand_secs += d_secs

            put(sh_df, f"A{rr}", loc_name, bold=True, fill=C_DAY, align="center")
            put(sh_df, f"B{rr}", d, bold=True, fill=C_DAY, align="center")
            put(sh_df, f"C{rr}", f"ИТОГО ЗА {d} (файлов: {len(by_file)})", bold=True, fill=C_DAY)
            put(sh_df, f"D{rr}", d_done, bold=True, fill=C_DAY, align="center")
            put(sh_df, f"E{rr}", round(d_wg, 1), bold=True, fmt="0.0", fill=C_DAY)
            put(sh_df, f"F{rr}", d_fail, bold=True, fill=C_DAY, align="center")
            put(sh_df, f"G{rr}", round(d_ww, 1), bold=True, fmt="0.0", fill=C_DAY)
            put(sh_df, f"H{rr}", d_secs / 86400 if d_secs else 0, bold=True, fmt=HF, fill=C_DAY, align="center")
            for col in "ABCDEFGH":
                sh_df[f"{col}{rr}"].border = day_box
            sh_df.row_dimensions[rr].height = 22
            rr += 1

            for fname in sorted(by_file.keys()):
                st = by_file[fname]
                put(sh_df, f"A{rr}", loc_name, align="center", color="595959")
                put(sh_df, f"B{rr}", d, align="center", color="595959")
                put(sh_df, f"C{rr}", fname)
                put(sh_df, f"D{rr}", st["done"], bold=True, align="center")
                put(sh_df, f"E{rr}", round(st["wg"], 1), fmt="0.0")
                put(sh_df, f"F{rr}", st["fail"], align="center", color="9C0006" if st["fail"] else None)
                put(sh_df, f"G{rr}", round(st["ww"], 1), fmt="0.0", color="9C0006" if st["ww"] else None)
                put(sh_df, f"H{rr}", st["secs"] / 86400 if st["secs"] else 0, fmt=HF, align="center")
                for col in "ABCDEFGH":
                    sh_df[f"{col}{rr}"].border = box
                rr += 1

        put(sh_df, f"A{rr}", loc_name, bold=True, fill=C_GREY, align="center")
        put(sh_df, f"B{rr}", "ВСЕГО", bold=True, fill=C_GREY, align="center")
        put(sh_df, f"C{rr}", f"Итого ({len(sorted_days)} дн.)", bold=True, fill=C_GREY)
        put(sh_df, f"D{rr}", grand_done, bold=True, fill=C_GREY, align="center")
        put(sh_df, f"E{rr}", round(grand_wg, 1), bold=True, fmt="0.0", fill=C_GREY)
        put(sh_df, f"F{rr}", grand_fail, bold=True, fill=C_GREY, align="center")
        put(sh_df, f"G{rr}", round(grand_ww, 1), bold=True, fmt="0.0", fill=C_GREY)
        put(sh_df, f"H{rr}", grand_secs / 86400 if grand_secs else 0, bold=True, fmt=HF, fill=C_GREY, align="center")
        for col in "ABCDEFGH":
            sh_df[f"{col}{rr}"].border = box
        sh_df.auto_filter.ref = f"A4:H{rr}"
        sh_df.freeze_panes = "A5"
        widths(sh_df, [14, 13, 42, 15, 18, 14, 15, 18])

        # 5. ДИНАМИКА ПО ДНЯМ
        put(sh_dyn, "A1", f"Динамика по дням — {loc_name} ({period_label})", bold=True, size=14)
        header(sh_dyn, 3, [
            "Локация", "Дата", "Готово, столов", "Пластик готовых, кг",
            "Срывы, столов", "Вес срывов, кг", "% брака (вес)", "Общее время печати"
        ])
        first_dyn = 4
        for i, d in enumerate(sorted_days):
            r_idx = first_dyn + i
            items = by_day[d]
            q_ok = sum(x["qty"] for x in items if not x.get("is_failed"))
            q_fl = sum(x["qty"] for x in items if x.get("is_failed"))
            wg_kg = sum(x["weight_good"] for x in items) / 1000.0
            ww_kg = sum(x["weight_waste"] for x in items) / 1000.0
            s_d = sum(x["tot_secs"] for x in items)

            put(sh_dyn, f"A{r_idx}", loc_name, align="center")
            put(sh_dyn, f"B{r_idx}", d, bold=True, align="center")
            put(sh_dyn, f"C{r_idx}", q_ok, bold=True, align="center")
            put(sh_dyn, f"D{r_idx}", round(wg_kg, 2), fmt="0.00")
            put(sh_dyn, f"E{r_idx}", q_fl, align="center")
            put(sh_dyn, f"F{r_idx}", round(ww_kg, 2), fmt="0.00")
            put(sh_dyn, f"G{r_idx}", f"=IF((D{r_idx}+F{r_idx})>0,F{r_idx}/(D{r_idx}+F{r_idx}),0)", fmt="0.0%", align="center", formula=True)
            put(sh_dyn, f"H{r_idx}", s_d / 86400 if s_d else 0, fmt=HF, align="center")
            for col in "ABCDEFGH":
                sh_dyn[f"{col}{r_idx}"].border = box

        last_dyn = max(first_dyn, first_dyn + len(sorted_days) - 1)
        tot_dyn = last_dyn + 1
        put(sh_dyn, f"A{tot_dyn}", loc_name, bold=True, fill=C_GREY, align="center")
        put(sh_dyn, f"B{tot_dyn}", "ИТОГО", bold=True, fill=C_GREY, align="center")
        for col, fmt_c in [("C", "0"), ("D", "0.00"), ("E", "0"), ("F", "0.00"), ("H", HF)]:
            put(sh_dyn, f"{col}{tot_dyn}", f"=SUM({col}{first_dyn}:{col}{last_dyn})", bold=True, fmt=fmt_c, fill=C_GREY, align="center" if col in "CEH" else None, formula=True)
        put(sh_dyn, f"G{tot_dyn}", f"=IF((D{tot_dyn}+F{tot_dyn})>0,F{tot_dyn}/(D{tot_dyn}+F{tot_dyn}),0)", bold=True, fmt="0.0%", fill=C_GREY, align="center", formula=True)
        for col in "ABCDEFGH":
            sh_dyn[f"{col}{tot_dyn}"].border = box
        sh_dyn.freeze_panes = "A4"
        widths(sh_dyn, [14, 13, 16, 18, 15, 15, 14, 18])

        # 6. ПРИНТЕРЫ И МОТОЧАСЫ
        put(sh_pr, "A1", f"Принтеры, моточасы и ТО — {loc_name} ({period_label})", bold=True, size=14)
        header(sh_pr, 4, [
            "Локация", "Принтер", "Модель", "Сопло", "IP / Wi-Fi",
            "Готовых, шт", "Срывов, шт", "Пластик, кг",
            "Часы за период, ч" if is_period else "По логам фермы, ч",
            "База с экрана, ч", "Дата сверки",
            "Прирост после сверки, ч", "АКТУАЛЬНО ОДОМЕТР, ч", "Статус ТО / HMS"
        ])

        for i, (pname, pstat) in enumerate(sorted(printers_stat.items())):
            r_idx = 5 + i
            hw = devices_hw.get(pname, {})
            odo = odometers.get(pname, {})

            period_h = (pstat["secs_good"] + pstat["secs_waste"]) / 3600.0
            all_time_log_h = pstat.get("total_all_time_secs", pstat["secs_good"] + pstat["secs_waste"]) / 3600.0
            tot_kg = (pstat.get("weight_good", 0.0) + pstat.get("weight_waste", 0.0)) / 1000.0

            has_odo = "base_hours" in odo
            base_h = float(odo["base_hours"]) if has_odo else None
            sync_str = odo.get("sync_time", "—") if has_odo else "Не сверялся"
            delta_h = (pstat["secs_after_sync"] / 3600.0) if has_odo else 0.0
            actual_odo_h = (base_h + delta_h) if has_odo else all_time_log_h
            last_to_h = float(odo.get("last_to_hours") or 0.0)
            hours_since_to = actual_odo_h - last_to_h

            ip_wifi = f"{hw.get('ip', '')} ({hw.get('wifi', '')})".strip(" ()") if hw else "—"
            hms_cnt = hw.get("hms_errors", 0)
            if hms_cnt > 0:
                to_status = f"⚠️ Ошибка HMS ({hms_cnt})"
                st_fill = C_WARN
            elif hours_since_to >= to_interval:
                to_status = f"🔧 Пора на ТО (+{hours_since_to:.0f}ч)"
                st_fill = C_REP
            else:
                to_status = f"ОК (до ТО {max(0, int(to_interval - hours_since_to))} ч)"
                st_fill = C_OK

            put(sh_pr, f"A{r_idx}", loc_name, align="center")
            put(sh_pr, f"B{r_idx}", pname, bold=True)
            put(sh_pr, f"C{r_idx}", hw.get("model", "—"), align="center")
            put(sh_pr, f"D{r_idx}", hw.get("nozzle", "—"), align="center")
            put(sh_pr, f"E{r_idx}", ip_wifi, align="center")
            put(sh_pr, f"F{r_idx}", round(pstat["copies"], 1), fmt="0.0", align="center")
            put(sh_pr, f"G{r_idx}", round(pstat["stopped_cnt"], 1), fmt="0.0", align="center")
            put(sh_pr, f"H{r_idx}", round(tot_kg, 2), fmt="0.00")
            put(sh_pr, f"I{r_idx}", round(period_h, 1), fmt="0.0")
            put(sh_pr, f"J{r_idx}", round(base_h, 1) if has_odo else "—", fmt="0.0" if has_odo else None, align="center" if not has_odo else None)
            put(sh_pr, f"K{r_idx}", sync_str, align="center")
            put(sh_pr, f"L{r_idx}", round(delta_h, 1) if has_odo else "—", fmt="0.0" if has_odo else None, align="center" if not has_odo else None)
            put(sh_pr, f"M{r_idx}", round(actual_odo_h, 1), bold=True, fmt="0.0", fill=C_DAY)
            put(sh_pr, f"N{r_idx}", to_status, bold=True, fill=st_fill, align="center")
            for col in "ABCDEFGHIJKLMN":
                sh_pr[f"{col}{r_idx}"].border = box

        sh_pr.auto_filter.ref = f"A4:N{max(5, 4 + len(printers_stat))}"
        sh_pr.freeze_panes = "A5"
        widths(sh_pr, [14, 15, 12, 18, 22, 13, 12, 13, 17, 15, 17, 19, 20, 22])

        for sh in wb.worksheets:
            sh.page_setup.orientation = "landscape"
            sh.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
            sh.page_setup.fitToWidth = 1
            sh.page_setup.fitToHeight = 0

        final_path, save_ok, save_msg = safe_save_and_open(wb, export_path, open_file=open_file, is_live_tick=is_live_tick)
        if not quiet and not is_live_tick:
            print(f"✅ Excel-отчёт готов ({period_label}): {final_path}")
        return final_path, save_ok, save_msg


def generate_period_report_interactive():
    parsed, _, _ = parser.load_and_parse(quiet=True)
    all_dates = sorted({x["date"] for x in (parsed or []) if x.get("date") and x["date"] != "Неизвестно"})
    min_d = all_dates[0] if all_dates else datetime.now().strftime("%Y-%m-%d")
    max_d = all_dates[-1] if all_dates else datetime.now().strftime("%Y-%m-%d")

    print("\n" + "=" * 78)
    print(" 📅 ВЫГРУЗКА ОТЧЁТА ЗА ВЫБРАННЫЙ ПЕРИОД")
    print("=" * 78)
    print(f" Доступные даты в базе: с {min_d} по {max_d}")
    print(" Формат ввода любой: ДД.ММ.ГГГГ (01.12.2026), ДД.ММ (01.12) или ГГГГ-ММ-ДД")
    print("-" * 78)
    raw_from = input(f" Начальная дата [Enter = с начала ({min_d})]: ").strip()
    raw_to = input(f" Конечная дата  [Enter = по сегодня ({max_d})]: ").strip()

    d_from = parse_user_date(raw_from, min_d)
    d_to = parse_user_date(raw_to, max_d)
    if d_from > d_to:
        d_from, d_to = d_to, d_from

    print(f"\n⏳ Формируем отчёт за период: {d_from} — {d_to}...")
    generate_report(date_from=d_from, date_to=d_to, open_file=True)


def _background_live_worker():
    state = sys._bambu_live_state
    last_excel_ts = time.time()
    last_cloud_ts = 0.0

    while state["active"]:
        for _ in range(30):
            if not state["active"]:
                connector.send_heartbeat(live=False)
                return
            time.sleep(1)

        if not state["active"]:
            connector.send_heartbeat(live=False)
            return

        try:
            now_ts = time.time()
            cfg = config.load_config()
            hb_ok = True
            hb_msg = state.get("status_msg") or "ОК"

            if (now_ts - last_excel_ts) >= state["interval"]:
                _, ok, msg = generate_report(
                    custom_filename=state["file"],
                    open_file=False,
                    is_live_tick=True,
                    quiet=True,
                )
                last_excel_ts = now_ts
                state["last_update"] = datetime.now().strftime("%H:%M:%S")
                state["status_msg"] = msg
                hb_ok = ok
                hb_msg = msg
            else:
                parser.sync_live_into_archive(quiet=True, scan_external_dumps=False)

            if gsheets and cfg.get("google_enabled") and cfg.get("google_webhook_url"):
                cloud_interval = max(5, int(cfg.get("google_sync_interval_min") or 30)) * 60
                if (now_ts - last_cloud_ts) >= cloud_interval:
                    c_ok, c_msg = gsheets.push_to_google(
                        live_active=True,
                        note=state.get("status_msg") or "ОК",
                        quiet=True,
                    )
                    last_cloud_ts = now_ts
                    state["cloud_status"] = c_msg
                    if c_ok:
                        state["last_cloud_sync"] = datetime.now().strftime("%H:%M")
                    else:
                        hb_ok = False
                        hb_msg = c_msg

            connector.send_heartbeat(
                live=True, ok=hb_ok, msg=hb_msg,
                last_save=state["last_update"],
                done=state["done_cnt"], fail=state["fail_cnt"],
            )
        except Exception as e:
            err_txt = str(e)[:60]
            state["status_msg"] = f"Ошибка: {err_txt}"
            connector.send_heartbeat(
                live=True, ok=False, msg=err_txt,
                last_save=state["last_update"],
                done=state["done_cnt"], fail=state["fail_cnt"],
            )


def toggle_live_excel_background(auto_mode=False):
    state = sys._bambu_live_state
    cfg = config.load_config()
    loc_tag = config.sanitize_filename_tag(cfg.get("location_name") or "Farm")
    state["file"] = f"Farm_Report_{loc_tag}_LIVE.xlsx"

    if state["active"] and not auto_mode:
        state["active"] = False
        connector.send_heartbeat(live=False)
        print("\n⏹️ Фоновая Live-запись ОСТАНОВЛЕНА.")
        time.sleep(0.8)
        return

    if state["active"] and auto_mode:
        return

    if auto_mode:
        interval = max(30, int(cfg.get("live_interval_sec") or 60))
    else:
        print("\n" + "=" * 78)
        print(" 🔴 ЗАПУСК ФОНОВОЙ LIVE-ЗАПИСИ (АРХИВ + EXCEL + GOOGLE SHEETS + ПУЛЬС)")
        print("=" * 78)
        def_i = int(cfg.get("live_interval_sec") or 60)
        sec_str = input(f"Интервал обновления Excel в секундах [Enter = {def_i} сек]: ").strip()
        interval = int(sec_str) if sec_str.isdigit() and int(sec_str) >= 30 else def_i
        cfg["live_interval_sec"] = interval
        config.save_config(cfg)

    state["interval"] = interval
    state["active"] = True
    state["status_msg"] = "Запуск..."

    if not auto_mode:
        print(f"⏳ Создаём живой отчёт {state['file']}...")
    try:
        _, ok, msg = generate_report(
            custom_filename=state["file"],
            open_file=(not auto_mode),
            is_live_tick=auto_mode,
            quiet=True,
        )
        state["last_update"] = datetime.now().strftime("%H:%M:%S")
        state["status_msg"] = msg
        connector.send_heartbeat(
            live=True, ok=ok, msg=msg,
            last_save=state["last_update"],
            done=state["done_cnt"], fail=state["fail_cnt"],
        )
    except Exception as e:
        state["status_msg"] = f"Ошибка: {e}"
        connector.send_heartbeat(live=True, ok=False, msg=str(e)[:60])

    t = threading.Thread(target=_background_live_worker, daemon=True)
    t.start()
    if not auto_mode:
        print(f"✅ Фоновая Live-запись активна! ({state['file']}, каждые {interval} сек).")
        time.sleep(1.2)


live_excel_monitor_loop = toggle_live_excel_background


def update_single_printer_odometer(pname, val_str, printers_stat, odometers):
    try:
        new_h = float(val_str.replace(",", "."))
        if new_h < 0 or new_h > 100000:
            raise ValueError
    except ValueError:
        print("❌ Некорректное число моточасов.")
        return

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    ps = printers_stat.get(pname, {"secs_good": 0, "secs_waste": 0, "secs_after_sync": 0, "events": []})
    old_rec = odometers.get(pname, {})

    if "base_hours" in old_rec and old_rec.get("sync_time"):
        prev_h = float(old_rec["base_hours"])
        prev_dt = parser.parse_iso_local(old_rec["sync_time"])
        screen_delta = new_h - prev_h
        log_delta = sum(s for dt, s in ps.get("events", []) if prev_dt and dt > prev_dt) / 3600.0
        drift_h = round(screen_delta - log_delta, 1)
        print(f"\n🔍 СВЕРКА [{pname}]: Экран +{screen_delta:.1f}ч | Логи +{log_delta:.1f}ч | Расхождение: {drift_h:+.1f}ч")
    else:
        log_total = (ps["secs_good"] + ps["secs_waste"]) / 3600.0
        drift_h = round(new_h - log_total, 1)
        print(f"\n✅ Первичная привязка [{pname}]: экран = {new_h:.1f}ч.")

    odometers[pname] = {
        "base_hours": new_h,
        "sync_time": now_str,
        "last_to_hours": old_rec.get("last_to_hours", 0.0),
        "last_to_date": old_rec.get("last_to_date", ""),
        "last_drift_hours": drift_h,
    }


def odometer_reconciliation_menu():
    _, printers_stat, _ = parser.load_and_parse(quiet=True)
    printers_stat = printers_stat or {}
    odometers = config.load_odometers()
    all_printers = sorted(set(list(printers_stat.keys()) + list(odometers.keys())))

    while True:
        print("\n" + "=" * 88)
        print(" 🔧 СВЕРКА МОТОЧАСОВ И РУЧНОЙ ОДОМЕТР С ЭКРАНА ПРИНТЕРА")
        print("=" * 88)
        print(f"{'№':>2} | {'Принтер':<12} | {'По логам':>9} | {'База экрана':>11} | {'Дата сверки':<16} | {'Прирост':>8} | {'ИТОГО ОДОМЕТР':>13} | {'Расхожд.'}")
        print("-" * 88)

        for idx, pname in enumerate(all_printers, 1):
            ps = printers_stat.get(pname, {"secs_good": 0, "secs_waste": 0, "secs_after_sync": 0})
            log_h = (ps["secs_good"] + ps["secs_waste"]) / 3600.0
            odo = odometers.get(pname)
            if odo and "base_hours" in odo:
                base_h = float(odo["base_hours"])
                sync_s = odo.get("sync_time", "—")[:16]
                delta_h = ps["secs_after_sync"] / 3600.0
                curr_odo = base_h + delta_h
                drift = odo.get("last_drift_hours")
                drift_s = f"{drift:+.1f} ч" if drift is not None else "—"
                print(f"{idx:2d} | {pname:<12} | {log_h:7.1f} ч | {base_h:9.1f} ч | {sync_s:<16} | +{delta_h:5.1f} ч | {curr_odo:11.1f} ч | {drift_s}")
            else:
                print(f"{idx:2d} | {pname:<12} | {log_h:7.1f} ч | {'(не задан)':>11} | {'—':<16} | {'—':>8} | {log_h:11.1f} ч | —")

        print("-" * 88)
        ch = input("Номер принтера / [A] Обход всех / [T] Отметить ТО / [Enter] Назад: ").strip()
        if not ch:
            break
        elif ch.upper() == "A":
            for pname in all_printers:
                val = input(f"  Одометр [{pname}] (Enter=пропуск, Q=выход): ").strip()
                if val.upper() == "Q":
                    break
                if val:
                    update_single_printer_odometer(pname, val, printers_stat, odometers)
            config.save_odometers(odometers)
            _, printers_stat, _ = parser.load_and_parse(quiet=True)
        elif ch.upper() == "T":
            num = input("Номер принтера после ТО: ").strip()
            if num.isdigit() and 1 <= int(num) <= len(all_printers):
                pname = all_printers[int(num) - 1]
                ps = printers_stat.get(pname, {"secs_good": 0, "secs_waste": 0, "secs_after_sync": 0})
                odo = odometers.get(pname, {})
                curr_h = (float(odo["base_hours"]) + ps["secs_after_sync"] / 3600.0) if "base_hours" in odo else ((ps["secs_good"] + ps["secs_waste"]) / 3600.0)
                odo["last_to_hours"] = round(curr_h, 1)
                odo["last_to_date"] = datetime.now().strftime("%Y-%m-%d")
                odometers[pname] = odo
                config.save_odometers(odometers)
                print(f"✅ ТО для [{pname}] зафиксировано на {curr_h:.1f} ч!")
        elif ch.isdigit() and 1 <= int(ch) <= len(all_printers):
            pname = all_printers[int(ch) - 1]
            val = input(f"Введи моточасы с экрана [{pname}] (ч): ").strip()
            if val:
                update_single_printer_odometer(pname, val, printers_stat, odometers)
                config.save_odometers(odometers)
                _, printers_stat, _ = parser.load_and_parse(quiet=True)
