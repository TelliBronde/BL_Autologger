import os
import re
import subprocess
import sys
import time
from collections import defaultdict
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

import config
import parser

HF = "[h]:mm"
C_HEAD = "1F3864"
C_DAY  = "D9E1F2"
C_OK   = "C6EFCE"
C_REP  = "FFF2CC"
C_WORK = "DDEBF7"
C_WARN = "FCE4D6"
C_GREY = "F2F2F2"


def put(ws, ref, v, bold=False, fmt=None, fill=None, color=None, align=None, size=10, wrap=False):
    c = ws[ref]
    c.value = v
    c.font = Font(name="Arial", size=size, bold=bold, color=color)
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = PatternFill("solid", fgColor=fill)
    if align or wrap:
        c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    return c


def header(ws, row, titles):
    thin = Side(style="thin", color="BFBFBF")
    for i, t in enumerate(titles):
        c = put(ws, f"{get_column_letter(i + 1)}{row}", t, bold=True, fill=C_HEAD, color="FFFFFF",
                align="center", wrap=True)
        c.border = Border(top=thin, bottom=thin, left=thin, right=thin)
    ws.row_dimensions[row].height = 32


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths):
        ws.column_dimensions[get_column_letter(i + 1)].width = w


def parse_user_date(raw_str: str, default_date: str) -> str:
    """Понимает форматы: ГГГГ-ММ-ДД, ДД.ММ.ГГГГ, ДД/ММ/ГГГГ и короткий ДД.ММ."""
    s = (raw_str or "").strip()
    if not s:
        return default_date
    # ГГГГ-ММ-ДД
    if re.match(r"^\d{4}-\d{2}-\d{2}$", s):
        return s
    # ДД.ММ.ГГГГ или ДД/ММ/ГГГГ
    m_full = re.match(r"^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$", s)
    if m_full:
        d, m, y = int(m_full.group(1)), int(m_full.group(2)), int(m_full.group(3))
        return f"{y:04d}-{m:02d}-{d:02d}"
    # Короткий формат ДД.ММ (год берём текущий)
    m_short = re.match(r"^(\d{1,2})[./-](\d{1,2})$", s)
    if m_short:
        d, m = int(m_short.group(1)), int(m_short.group(2))
        y = datetime.now().year
        return f"{y:04d}-{m:02d}-{d:02d}"
    return default_date


def close_workbook_in_excel_com(file_path: Path) -> bool:
    """Если файл открыт в Excel, мягко закрывает только эту книгу через COM, чтобы снять блокировку записи."""
    if sys.platform != "win32":
        return False
    fname = file_path.name.replace("'", "''")
    ps_script = (
        "$ErrorActionPreference='SilentlyContinue'; "
        "$xl=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application'); "
        "if($xl){ "
        "  $xl.DisplayAlerts=$false; "
        f"  foreach($wb in $xl.Workbooks){{ if($wb.Name -eq '{fname}'){{ $wb.Close($false); $xl.DisplayAlerts=$true; exit 0 }} }} "
        "  $xl.DisplayAlerts=$true; "
        "} exit 1"
    )
    try:
        res = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_script],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=4,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        if res.returncode == 0:
            time.sleep(0.3)
            return True
    except Exception:
        pass
    return False


def safe_save_and_open(wb: Workbook, export_path: Path, open_file=True) -> Path:
    """Безопасно сохраняет Excel-файл без падения скрипта, даже если старый отчёт открыт в Excel."""
    try:
        wb.save(export_path)
    except PermissionError:
        print(f"🔄 Файл {export_path.name} открыт в Excel — обновляем открытую книгу...")
        close_workbook_in_excel_com(export_path)
        try:
            wb.save(export_path)
        except PermissionError:
            # Если в Excel активен курсор ввода внутри ячейки, сохраняем с меткой времени
            ts = datetime.now().strftime("%H-%M-%S")
            alt_path = export_path.with_name(f"{export_path.stem}_{ts}{export_path.suffix}")
            wb.save(alt_path)
            export_path = alt_path
            print(f"⚠️ В Excel редактируется ячейка. Отчёт сохранён в новый файл: {export_path.name}")

    if open_file and sys.platform == "win32":
        try:
            os.startfile(export_path)
        except Exception as e:
            print(f"⚠️ Файл сохранён ({export_path}), но не удалось открыть автоматически: {e}")
    return export_path


def build_Period_printers_stat(filtered_parsed, full_printers_stat):
    """Пересчитывает выработку по принтерам строго за выбранный диапазон дат."""
    p_stat = defaultdict(lambda: {
        "tasks": 0, "copies": 0.0, "stopped_cnt": 0.0,
        "weight": 0.0, "secs_good": 0.0, "secs_waste": 0.0,
        "secs_after_sync": 0.0,
    })
    # Переносим актуальный прирост после сверки для корректного отображения общего одометра
    for pname, full_st in full_printers_stat.items():
        p_stat[pname]["secs_after_sync"] = full_st.get("secs_after_sync", 0.0)
        p_stat[pname]["total_all_time_secs"] = full_st.get("secs_good", 0.0) + full_st.get("secs_waste", 0.0)

    for x in filtered_parsed:
        devs = [d.strip() for d in (x.get("printers") or "").split(",") if d.strip() and d.strip() != "—"]
        if not devs:
            continue
        n = len(devs)
        for dn in devs:
            p_stat[dn]["tasks"] += 1
            p_stat[dn]["copies"] += x["qty"] / n
            p_stat[dn]["weight"] += x["tot_weight"] / n
            p_stat[dn]["secs_good"] += x["tot_secs"] / n

    return p_stat


def generate_report(date_from=None, date_to=None, custom_filename=None, open_file=True):
    res = parser.load_and_parse(quiet=bool(custom_filename))
    if not res or not res[0]:
        return None
    parsed_all, full_printers_stat, devices_hw = res
    odometers = config.load_odometers()
    cfg = config.load_config()
    to_interval = float(cfg.get("to_interval_hours") or 300)

    is_period = bool(date_from or date_to)
    d_min = date_from or "0000-00-00"
    d_max = date_to or "9999-99-99"

    parsed = [x for x in parsed_all if d_min <= x["date"] <= d_max]
    if not parsed:
        print(f"⚠️ За выбранный период ({d_min} ... {d_max}) завершённых задач не найдено.")
        return None

    printers_stat = build_Period_printers_stat(parsed, full_printers_stat) if is_period else full_printers_stat
    period_label = f"за период {d_min} — {d_max}" if is_period else f"за все дни (обновлено {datetime.now():%H:%M:%S})"

    by_day = defaultdict(list)
    for x in parsed:
        by_day[x["date"]].append(x)

    if custom_filename:
        export_path = Path.home() / "Desktop" / custom_filename
    elif is_period:
        export_path = Path.home() / "Desktop" / f"farm_report_{d_min}_to_{d_max}.xlsx"
    else:
        export_path = Path.home() / "Desktop" / f"farm_report_{datetime.now():%Y-%m-%d}.xlsx"

    wb = Workbook()
    sh_df = wb.active
    sh_df.title = "По дням и файлам"
    sh_dyn = wb.create_sheet("Динамика по дням")
    sh_all = wb.create_sheet("Все задачи")
    sh_pr = wb.create_sheet("Принтеры и Моточасы")

    thin = Side(style="thin", color="D9D9D9")
    med_bottom = Side(style="medium", color="8EA9DB")
    box = Border(top=thin, bottom=thin, left=thin, right=thin)
    day_box = Border(top=med_bottom, bottom=med_bottom, left=thin, right=thin)
    sorted_days = sorted(by_day.keys(), reverse=True)

    # 1. ПО ДНЯМ И ФАЙЛАМ
    put(sh_df, "A1", f"Разбивка выполненных задач по дням и файлам ({period_label})", bold=True, size=14)
    header(sh_df, 4, ["Дата", "Файл (задача)", "Запусков / Повторов", "Напечатано, шт",
                      "Вес 1 шт, г", "Всего пластик, г", "Время 1 шт", "Общее время печати"])
    rr = 5
    grand_runs = grand_qty = grand_secs = 0
    grand_weight = 0.0

    for d in sorted_days:
        day_items = by_day[d]
        by_file = defaultdict(lambda: {"runs": 0, "qty": 0, "w_one": 0.0, "w_tot": 0.0, "s_one": 0, "s_tot": 0})
        for x in day_items:
            f_rec = by_file[x["file"]]
            f_rec["runs"] += 1
            f_rec["qty"] += x["qty"]
            f_rec["w_one"] = x["weight_one"]
            f_rec["w_tot"] += x["tot_weight"]
            f_rec["s_one"] = x["secs_one"]
            f_rec["s_tot"] += x["tot_secs"]

        d_runs = sum(v["runs"] for v in by_file.values())
        d_qty = sum(v["qty"] for v in by_file.values())
        d_weight = sum(v["w_tot"] for v in by_file.values())
        d_secs = sum(v["s_tot"] for v in by_file.values())
        grand_runs += d_runs
        grand_qty += d_qty
        grand_weight += d_weight
        grand_secs += d_secs

        put(sh_df, f"A{rr}", d, bold=True, fill=C_DAY, align="center")
        put(sh_df, f"B{rr}", f"ИТОГО ЗА {d}  (уникальных файлов: {len(by_file)})", bold=True, fill=C_DAY)
        put(sh_df, f"C{rr}", d_runs, bold=True, fill=C_DAY, align="center")
        put(sh_df, f"D{rr}", d_qty, bold=True, fill=C_DAY, align="center")
        put(sh_df, f"E{rr}", "", fill=C_DAY)
        put(sh_df, f"F{rr}", round(d_weight, 1), bold=True, fmt="0.0", fill=C_DAY)
        put(sh_df, f"G{rr}", "", fill=C_DAY)
        put(sh_df, f"H{rr}", d_secs / 86400 if d_secs else 0, bold=True, fmt=HF, fill=C_DAY, align="center")
        for col in "ABCDEFGH":
            sh_df[f"{col}{rr}"].border = day_box
        sh_df.row_dimensions[rr].height = 22
        rr += 1

        for fname in sorted(by_file.keys()):
            st = by_file[fname]
            put(sh_df, f"A{rr}", d, align="center", color="595959")
            put(sh_df, f"B{rr}", fname)
            put(sh_df, f"C{rr}", st["runs"], align="center")
            put(sh_df, f"D{rr}", st["qty"], bold=True, align="center")
            put(sh_df, f"E{rr}", round(st["w_one"], 1), fmt="0.0")
            put(sh_df, f"F{rr}", round(st["w_tot"], 1), bold=True, fmt="0.0")
            put(sh_df, f"G{rr}", st["s_one"] / 86400 if st["s_one"] else None, fmt=HF, align="center")
            put(sh_df, f"H{rr}", st["s_tot"] / 86400 if st["s_tot"] else None, fmt=HF, align="center")
            for col in "ABCDEFGH":
                sh_df[f"{col}{rr}"].border = box
            rr += 1

    put(sh_df, f"A{rr}", "ВСЕГО", bold=True, fill=C_GREY, align="center")
    put(sh_df, f"B{rr}", f"Итого ({len(sorted_days)} дн. | {period_label})", bold=True, fill=C_GREY)
    put(sh_df, f"C{rr}", grand_runs, bold=True, fill=C_GREY, align="center")
    put(sh_df, f"D{rr}", grand_qty, bold=True, fill=C_GREY, align="center")
    put(sh_df, f"E{rr}", "", fill=C_GREY)
    put(sh_df, f"F{rr}", round(grand_weight, 1), bold=True, fmt="0.0", fill=C_GREY)
    put(sh_df, f"G{rr}", "", fill=C_GREY)
    put(sh_df, f"H{rr}", grand_secs / 86400 if grand_secs else 0, bold=True, fmt=HF, fill=C_GREY, align="center")
    for col in "ABCDEFGH":
        sh_df[f"{col}{rr}"].border = box
    sh_df.auto_filter.ref = f"A4:H{rr}"
    sh_df.freeze_panes = "A5"
    widths(sh_df, [14, 42, 18, 16, 14, 18, 14, 18])

    # 2. ДИНАМИКА ПО ДНЯМ
    put(sh_dyn, "A1", f"Сводная статистика по дням ({period_label})", bold=True, size=14)
    header(sh_dyn, 3, ["Дата", "Запусков тасок", "Напечатано, шт", "Уник. файлов", "Пластик, г", "Пластик, кг", "Общее время печати"])
    first_dyn = 4
    for i, d in enumerate(sorted_days):
        r_idx = first_dyn + i
        items = by_day[d]
        q_d = sum(x["qty"] for x in items)
        w_d = sum(x["tot_weight"] for x in items)
        s_d = sum(x["tot_secs"] for x in items)
        u_files = len({x["file"] for x in items})
        put(sh_dyn, f"A{r_idx}", d, bold=True, align="center")
        put(sh_dyn, f"B{r_idx}", len(items), align="center")
        put(sh_dyn, f"C{r_idx}", q_d, bold=True, align="center")
        put(sh_dyn, f"D{r_idx}", u_files, align="center")
        put(sh_dyn, f"E{r_idx}", round(w_d, 1), bold=True, fmt="0.0")
        put(sh_dyn, f"F{r_idx}", round(w_d / 1000.0, 2), fmt="0.00")
        put(sh_dyn, f"G{r_idx}", s_d / 86400 if s_d else 0, fmt=HF, align="center")
        for col in "ABCDEFG":
            sh_dyn[f"{col}{r_idx}"].border = box

    last_dyn = first_dyn + len(sorted_days) - 1
    tot_dyn = last_dyn + 1
    put(sh_dyn, f"A{tot_dyn}", "ИТОГО", bold=True, fill=C_GREY, align="center")
    for col in "BCD":
        put(sh_dyn, f"{col}{tot_dyn}", f"=SUM({col}{first_dyn}:{col}{last_dyn})", bold=True, fill=C_GREY, align="center")
    put(sh_dyn, f"E{tot_dyn}", f"=SUM(E{first_dyn}:E{last_dyn})", bold=True, fmt="0.0", fill=C_GREY)
    put(sh_dyn, f"F{tot_dyn}", f"=SUM(F{first_dyn}:F{last_dyn})", bold=True, fmt="0.00", fill=C_GREY)
    put(sh_dyn, f"G{tot_dyn}", f"=SUM(G{first_dyn}:G{last_dyn})", bold=True, fmt=HF, fill=C_GREY, align="center")
    for col in "ABCDEFG":
        sh_dyn[f"{col}{tot_dyn}"].border = box
    sh_dyn.freeze_panes = "A4"
    widths(sh_dyn, [14, 16, 16, 15, 15, 14, 18])

    # 3. ВСЕ ЗАДАЧИ
    put(sh_all, "A1", f"Журнал задач и повторов ({len(parsed)} шт. | {period_label})", bold=True, size=14)
    header(sh_all, 3, ["Дата", "Время", "Статус", "Файл (задача)", "Принтеры в пуле",
                       "Напечатано, шт", "План. время (1 шт)", "Вес 1 шт, г", "Всего пластик, г"])
    for i, t in enumerate(parsed):
        r_idx = 4 + i
        st_fill = C_REP if t.get("is_reprint") else (C_WORK if "В работе" in t["status"] else C_OK)
        put(sh_all, f"A{r_idx}", t["date"], align="center")
        put(sh_all, f"B{r_idx}", t["time"], align="center")
        put(sh_all, f"C{r_idx}", t["status"], bold=True, fill=st_fill, align="center")
        put(sh_all, f"D{r_idx}", t["file"])
        put(sh_all, f"E{r_idx}", t["printers"])
        put(sh_all, f"F{r_idx}", t["qty"], bold=True, align="center")
        put(sh_all, f"G{r_idx}", t["secs_one"] / 86400 if t["secs_one"] else None, fmt=HF, align="center")
        put(sh_all, f"H{r_idx}", round(t["weight_one"], 1), fmt="0.0")
        put(sh_all, f"I{r_idx}", round(t["tot_weight"], 1), bold=True, fmt="0.0")
        for col in "ABCDEFGHI":
            sh_all[f"{col}{r_idx}"].border = box
    sh_all.auto_filter.ref = f"A3:I{3 + len(parsed)}"
    sh_all.freeze_panes = "A4"
    widths(sh_all, [14, 9, 22, 42, 28, 15, 16, 14, 16])

    # 4. ПРИНТЕРЫ И МОТОЧАСЫ
    put(sh_pr, "A1", f"Сверка моточасов, выработка ({period_label}) и статус ТО", bold=True, size=14)
    header(sh_pr, 4, [
        "Принтер", "Модель", "Сопло", "IP / Wi-Fi",
        "Готовых, шт", "Пластик, кг",
        "Часы за период, ч" if is_period else "По логам фермы, ч",
        "База с экрана, ч", "Дата сверки",
        "Прирост после сверки, ч", "АКТУАЛЬНО ОДОМЕТР, ч",
        "Сверка (Экран−Лог), ч", "Статус ТО / HMS"
    ])

    for i, (pname, pstat) in enumerate(sorted(printers_stat.items())):
        r_idx = 5 + i
        hw = devices_hw.get(pname, {})
        odo = odometers.get(pname, {})

        period_h = (pstat["secs_good"] + pstat["secs_waste"]) / 3600.0
        all_time_log_h = pstat.get("total_all_time_secs", pstat["secs_good"] + pstat["secs_waste"]) / 3600.0

        has_odo = "base_hours" in odo
        base_h = float(odo["base_hours"]) if has_odo else None
        sync_str = odo.get("sync_time", "—") if has_odo else "Не сверялся"
        delta_h = (pstat["secs_after_sync"] / 3600.0) if has_odo else 0.0
        actual_odo_h = (base_h + delta_h) if has_odo else all_time_log_h
        drift_h = odo.get("last_drift_hours")
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

        put(sh_pr, f"A{r_idx}", pname, bold=True)
        put(sh_pr, f"B{r_idx}", hw.get("model", "—"), align="center")
        put(sh_pr, f"C{r_idx}", hw.get("nozzle", "—"), align="center")
        put(sh_pr, f"D{r_idx}", ip_wifi, align="center")
        put(sh_pr, f"E{r_idx}", round(pstat["copies"], 1), fmt="0.0", align="center")
        put(sh_pr, f"F{r_idx}", round(pstat["weight"] / 1000.0, 2), fmt="0.00")
        put(sh_pr, f"G{r_idx}", round(period_h, 1), fmt="0.0")
        put(sh_pr, f"H{r_idx}", round(base_h, 1) if has_odo else "—", fmt="0.0" if has_odo else None, align="center" if not has_odo else None)
        put(sh_pr, f"I{r_idx}", sync_str, align="center")
        put(sh_pr, f"J{r_idx}", round(delta_h, 1) if has_odo else "—", fmt="0.0" if has_odo else None, align="center" if not has_odo else None)
        put(sh_pr, f"K{r_idx}", round(actual_odo_h, 1), bold=True, fmt="0.0", fill=C_DAY)
        put(sh_pr, f"L{r_idx}", round(drift_h, 1) if drift_h is not None else "—", fmt="+0.0;-0.0;0.0" if drift_h is not None else None, align="center")
        put(sh_pr, f"M{r_idx}", to_status, bold=True, fill=st_fill, align="center")

        for col in "ABCDEFGHIJKLM":
            sh_pr[f"{col}{r_idx}"].border = box

    sh_pr.auto_filter.ref = f"A4:M{4 + len(printers_stat)}"
    sh_pr.freeze_panes = "A5"
    widths(sh_pr, [15, 12, 18, 22, 14, 13, 18, 16, 17, 19, 21, 18, 22])

    for sh in wb.worksheets:
        sh.page_setup.orientation = "landscape"
        sh.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        sh.page_setup.fitToWidth = 1
        sh.page_setup.fitToHeight = 0

    final_path = safe_save_and_open(wb, export_path, open_file=open_file)
    print(f"✅ Excel-отчёт готов ({period_label}): {final_path}")
    return final_path


def generate_period_report_interactive():
    """Меню выбора дат вручную для формирования отчёта за конкретный период."""
    res = parser.load_and_parse(quiet=True)
    if not res or not res[0]:
        print("❌ Нет данных для отчёта.")
        return
    all_dates = sorted({x["date"] for x in res[0] if x.get("date") and x["date"] != "Неизвестно"})
    min_d = all_dates[0] if all_dates else datetime.now().strftime("%Y-%m-%d")
    max_d = all_dates[-1] if all_dates else datetime.now().strftime("%Y-%m-%d")

    print("\n" + "=" * 78)
    print(" 📅 ВЫГРУЗКА ОТЧЁТА ЗА ВЫБРАННЫЙ ПЕРИОД")
    print("=" * 78)
    print(f" Доступные даты в базе: с {min_d} по {max_d}")
    print(" Формат ввода любой: ДД.ММ.ГГГГ (10.11.2026), ДД.ММ (10.11) или ГГГГ-ММ-ДД")
    print("-" * 78)
    raw_from = input(f" Начальная дата [Enter = с начала ({min_d})]: ").strip()
    raw_to = input(f" Конечная дата  [Enter = по сегодня ({max_d})]: ").strip()

    d_from = parse_user_date(raw_from, min_d)
    d_to = parse_user_date(raw_to, max_d)
    if d_from > d_to:
        d_from, d_to = d_to, d_from

    print(f"\n⏳ Формируем отчёт за период: {d_from} — {d_to}...")
    generate_report(date_from=d_from, date_to=d_to, open_file=True)


def live_excel_monitor_loop():
    """Режим живой таблицы Excel: автоматически обновляет открытый файл каждые N секунд."""
    print("\n" + "=" * 78)
    print(" 🔴 РЕЖИМ ЖИВОЙ ТАБЛИЦЫ EXCEL (REAL-TIME AUTO-UPDATE)")
    print("=" * 78)
    sec_str = input("Интервал автообновления в секундах [Enter = 30 сек]: ").strip()
    interval = int(sec_str) if sec_str.isdigit() and int(sec_str) >= 10 else 30

    live_filename = "farm_report_LIVE.xlsx"
    print(f"\n🚀 Запущен живой мониторинг! Файл на Рабочем столе: {live_filename}")
    print("   Таблица в Excel будет сама обновляться каждые", interval, "сек.")
    print("   Для остановки нажми Ctrl+C в этом окне.\n")

    try:
        while True:
            generate_report(custom_filename=live_filename, open_file=True)
            for remaining in range(interval, 0, -1):
                sys.stdout.write(f"\r⏳ Следующее обновление таблицы через {remaining:2d} сек... (Ctrl+C — стоп) ")
                sys.stdout.flush()
                time.sleep(1)
            print()
    except KeyboardInterrupt:
        print("\n\n⏹️ Живое обновление остановлено. Возврат в меню.")


def update_single_printer_odometer(pname, val_str, printers_stat, odometers):
    try:
        new_h = float(val_str.replace(",", "."))
    except ValueError:
        print("❌ Некорректное число.")
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
        print(f"\n✅ Первичная привязка [{pname}]: экран = {new_h:.1f}ч (в логах было {log_total:.1f}ч).")

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
