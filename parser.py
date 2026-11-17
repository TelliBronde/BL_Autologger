import json
import re
from collections import defaultdict
from datetime import datetime, timedelta

import config
import connector

TASK_RE = re.compile(
    r"^(?P<name>.*?)_(?:(?P<h>\d+)h(?:(?P<m>\d+)m)?|(?P<m2>\d+)m)_(?P<g>\d+(?:[.,]\d+)?)g$",
    re.I,
)
CTRL_CHARS_RE = re.compile(r"[\x00-\x1f\x7f]")


def clean_str(val) -> str:
    """Вычищает управляющие символы, способные сломать XML внутри Excel."""
    return CTRL_CHARS_RE.sub("", str(val or "")).strip()


def safe_int(val, default=0) -> int:
    try:
        return int(float(val or 0))
    except Exception:
        return default


def safe_float(val, default=0.0) -> float:
    try:
        return float(str(val or 0).replace(",", "."))
    except Exception:
        return default


def strip_ext(name):
    n = clean_str(name)
    for ext in (".gcode.3mf", ".3mf", ".gcode"):
        if n.lower().endswith(ext):
            n = n[: -len(ext)]
    return n


def parse_task_name(raw):
    n = strip_ext(raw)
    m = TASK_RE.match(n)
    if not m:
        return n, None, None
    secs = safe_int(m.group("h")) * 3600 + safe_int(m.group("m") or m.group("m2")) * 60
    grams = safe_float(m.group("g"))
    return m.group("name"), secs, grams


def fmt_h(sec):
    sec = safe_int(sec)
    return f"{sec // 3600}ч{(sec % 3600) // 60:02d}м"


def parse_iso_local(ts_str):
    if not ts_str or str(ts_str).startswith(("1970", "0001")):
        return None
    try:
        dt = datetime.fromisoformat(str(ts_str).replace("Z", "+00:00"))
        return dt.astimezone().replace(tzinfo=None)
    except Exception:
        try:
            return datetime.strptime(str(ts_str)[:16], "%Y-%m-%d %H:%M")
        except Exception:
            return None


def extract_metrics(task, raw_name):
    weight, secs = None, None
    try:
        f3mf = task.get("f3mf_info") or {}
        plates = f3mf.get("slice_info_config", {}).get("plates", [])
        if plates and isinstance(plates, list):
            meta = plates[0].get("metadata", {})
            fil = sum(safe_float(x.get("used_g")) for x in plates[0].get("filaments", []) if isinstance(x, dict))
            if fil > 0:
                weight = fil
            elif meta.get("weight"):
                weight = safe_float(meta["weight"])
            if meta.get("prediction"):
                secs = safe_int(meta["prediction"])
    except Exception:
        pass

    clean_name, name_secs, name_weight = parse_task_name(raw_name)
    if weight is None:
        weight = name_weight or 0.0
    if secs is None:
        secs = name_secs or 0
    return clean_name, secs, weight


def find_all_json_files(quiet=False):
    live = connector.pull_live_data_from_farm()
    if live and not quiet:
        print("⚡ Живые данные автоматически получены из открытого окна Farm Manager!")
    found = []
    for d in config.SEARCH_DIRS:
        if not d.exists():
            continue
        for pat in ("history*", "timelogger*"):
            for p in d.glob(pat):
                try:
                    if p.is_symlink() or not p.is_file():
                        continue
                    if p.suffix.lower() not in (".json", ".txt"):
                        continue
                    sz = p.stat().st_size
                    if 100 < sz <= config.MAX_JSON_BYTES:
                        found.append(p.resolve())
                except Exception:
                    pass
    return list(set(found))


def walk_tasks(o):
    if isinstance(o, dict):
        if "task_name" in o and ("f3mf_info" in o or "statistics" in o):
            yield o
            return
        for v in o.values():
            yield from walk_tasks(v)
    elif isinstance(o, list):
        for v in o:
            yield from walk_tasks(v)


def walk_reprints(o):
    if isinstance(o, dict):
        if "__reprints_list__" in o and isinstance(o["__reprints_list__"], list):
            for item in o["__reprints_list__"]:
                if isinstance(item, dict):
                    yield item
            return
        for v in o.values():
            yield from walk_reprints(v)
    elif isinstance(o, list):
        for v in o:
            yield from walk_reprints(v)


def walk_devices_info(o):
    if isinstance(o, dict):
        if "__devices_info__" in o and isinstance(o["__devices_info__"], dict):
            for dev in o["__devices_info__"].values():
                if isinstance(dev, dict):
                    yield dev
            return
        if "devices" in o and isinstance(o["devices"], list):
            for d in o["devices"]:
                if isinstance(d, dict) and ("dev_id" in d or "report_status" in d):
                    rs = d.get("report_status") or {}
                    dev_name = clean_str(d.get("name")) or clean_str(d.get("dev_name")) or clean_str(d.get("dev_id"))
                    noz_t = clean_str(rs.get("nozzle_type"))
                    noz_str = f"{clean_str(rs.get('nozzle_diameter', '0.4'))} ({'Hardened' if noz_t == 'hardened_steel' else 'Stainless'})" if noz_t else ""
                    yield {
                        "dev_id": clean_str(d.get("dev_id")),
                        "name": dev_name,
                        "model": "Bambu A1" if d.get("dev_model") == "N2S" else ("Bambu P1S" if d.get("dev_model") == "C12" else clean_str(d.get("dev_model"))),
                        "ip": clean_str(d.get("dev_ip")),
                        "fw": clean_str((d.get("upgrade_info") or {}).get("current_ver")),
                        "nozzle": noz_str,
                        "wifi": clean_str(rs.get("wifi_signal")),
                        "hms_errors": len(rs.get("hms") or []),
                        "state": clean_str(rs.get("gcode_state")),
                        "mc_percent": safe_float(rs.get("mc_percent")),
                        "mc_remaining_time": safe_float(rs.get("mc_remaining_time")),
                        "current_file": clean_str(rs.get("subtask_name") or rs.get("gcode_file")),
                    }
        for v in o.values():
            yield from walk_devices_info(v)
    elif isinstance(o, list):
        for v in o:
            yield from walk_devices_info(v)


def load_and_parse(quiet=False):
    with config.get_lock():
        json_files = find_all_json_files(quiet=quiet)
        if not json_files:
            if not quiet:
                print("❌ Не найден ни один файл history*.json.")
            return None, None, None

        best = {}
        all_reprints_raw = []
        devices_hw = {}
        used_files = []

        for jp in sorted(json_files, key=lambda p: p.stat().st_mtime):
            try:
                with open(jp, "r", encoding="utf-8-sig") as f:
                    raw_data = json.load(f)
                used_files.append(jp.name)
                for t in walk_tasks(raw_data):
                    st = t.get("statistics") if isinstance(t.get("statistics"), dict) else {}
                    score = (safe_int(st.get("finished")) + safe_int(st.get("stopped")), str(t.get("finish_time") or ""))
                    key = str(t.get("id") or id(t))
                    if key not in best or score >= best[key][0]:
                        best[key] = (score, t)
                all_reprints_raw.extend(list(walk_reprints(raw_data)))
                for d_info in walk_devices_info(raw_data):
                    d_name = clean_str(d_info.get("name"))
                    if d_name:
                        devices_hw[d_name] = d_info
            except Exception as e:
                if not quiet:
                    print(f"⚠️ Пропущен повреждённый файл {jp.name}: {e}")

        if not quiet and used_files:
            print(f"📂 Файлы истории ({len(used_files)}): {', '.join(used_files)}")

        tasks_list = [v[1] for v in best.values()]
        if not tasks_list and not devices_hw:
            return None, None, None

        odometers = config.load_odometers()
        metrics_by_file = {}
        for t in tasks_list:
            f3mf = t.get("f3mf_info") if isinstance(t.get("f3mf_info"), dict) else {}
            raw_name = strip_ext(t.get("task_name") or f3mf.get("f3mf_name") or t.get("name") or "Unknown")
            clean_name, secs_one, weight_one = extract_metrics(t, raw_name)
            if weight_one > 0 or secs_one > 0:
                metrics_by_file[raw_name.lower()] = (raw_name, clean_name, secs_one, weight_one)
                if f3mf.get("f3mf_name"):
                    metrics_by_file[strip_ext(f3mf["f3mf_name"]).lower()] = (raw_name, clean_name, secs_one, weight_one)

        printer_history = defaultdict(list)
        parsed = []
        printers_stat = defaultdict(lambda: {
            "tasks": 0, "copies": 0.0, "stopped_cnt": 0.0,
            "weight": 0.0, "secs_good": 0.0, "secs_waste": 0.0,
            "secs_after_sync": 0.0, "events": []
        })

        def add_printer_time(dn, event_dt, secs_val, is_waste=False):
            if is_waste:
                printers_stat[dn]["secs_waste"] += secs_val
            else:
                printers_stat[dn]["secs_good"] += secs_val
            if event_dt:
                printers_stat[dn]["events"].append((event_dt, secs_val))
                odo_rec = odometers.get(dn)
                if odo_rec and odo_rec.get("sync_time"):
                    sync_dt = parse_iso_local(odo_rec["sync_time"])
                    if sync_dt and event_dt > sync_dt:
                        printers_stat[dn]["secs_after_sync"] += secs_val

        for t in tasks_list:
            stats = t.get("statistics") if isinstance(t.get("statistics"), dict) else {}
            qty = safe_int(stats.get("finished"), 1 if not stats else 0)
            stopped = safe_int(stats.get("stopped"))
            printing_now = safe_int(stats.get("printing"))
            total_p = safe_int(stats.get("total"), qty)

            f3mf = t.get("f3mf_info") if isinstance(t.get("f3mf_info"), dict) else {}
            raw_name = strip_ext(t.get("task_name") or f3mf.get("f3mf_name") or t.get("name") or "Unknown")
            clean_name, secs_one, weight_one = extract_metrics(t, raw_name)

            if (weight_one == 0 or secs_one == 0) and raw_name.lower() in metrics_by_file:
                _, clean_name, s_ref, w_ref = metrics_by_file[raw_name.lower()]
                weight_one = weight_one or w_ref
                secs_one = secs_one or s_ref

            c_dt = parse_iso_local(t.get("create_time")) or parse_iso_local(t.get("start_time"))
            f_dt = parse_iso_local(t.get("finish_time"))
            expected_f_dt = (c_dt + timedelta(seconds=secs_one)) if (c_dt and secs_one) else c_dt

            if not f_dt:
                dt = expected_f_dt
            elif expected_f_dt and (f_dt - expected_f_dt).total_seconds() > 3 * 3600:
                dt = expected_f_dt
            else:
                dt = f_dt

            pool = t.get("device_pool") or []
            dev_names = [clean_str(d.get("dev_name") or d.get("name") or d.get("dev_id")) for d in pool if isinstance(d, dict)]
            dev_names = [x for x in dev_names if x]

            if stopped > 0 and dev_names and c_dt:
                actual_dur = (f_dt - c_dt).total_seconds() if (f_dt and f_dt > c_dt) else (secs_one * 0.3)
                waste_one = max(60.0, min(float(secs_one or 3600), actual_dur))
                tot_waste_secs = waste_one * stopped
                for dn in dev_names:
                    printers_stat[dn]["stopped_cnt"] += stopped / len(dev_names)
                    add_printer_time(dn, f_dt or c_dt, tot_waste_secs / len(dev_names), is_waste=True)

            if t.get("state") == "ongoing" and printing_now > 0 and dev_names:
                for dn in dev_names:
                    hw = devices_hw.get(dn)
                    if hw and hw.get("mc_percent", 0) > 0 and secs_one > 0:
                        add_printer_time(dn, datetime.now(), secs_one * (hw["mc_percent"] / 100.0), is_waste=False)

            if qty <= 0:
                continue

            date_str = dt.strftime("%Y-%m-%d") if dt else "Неизвестно"
            time_str = dt.strftime("%H:%M") if dt else ""
            tot_weight = weight_one * qty
            tot_secs = secs_one * qty
            printers_str = ", ".join(dev_names) or "—"

            for dn in dev_names:
                if c_dt:
                    printer_history[dn].append({
                        "create_dt": c_dt, "finish_dt": dt, "file": raw_name,
                        "clean_name": clean_name, "secs_one": secs_one, "weight_one": weight_one,
                    })

            done_devs = [
                clean_str(d.get("dev_name") or d.get("name") or d.get("dev_id"))
                for d in pool if isinstance(d, dict) and not d.get("is_in_running_task")
            ]
            target_devs = [x for x in done_devs if x] or dev_names
            if target_devs:
                for dn in target_devs:
                    printers_stat[dn]["tasks"] += 1
                    printers_stat[dn]["copies"] += qty / len(target_devs)
                    printers_stat[dn]["weight"] += tot_weight / len(target_devs)
                    add_printer_time(dn, dt, tot_secs / len(target_devs), is_waste=False)

            state_str = f"В работе ({qty}/{total_p})" if t.get("state") == "ongoing" else (f"Готово (отмен: {stopped})" if stopped > 0 else "Готово")
            parsed.append({
                "dt": dt or datetime.min, "date": date_str, "time": time_str,
                "status": state_str, "is_reprint": False,
                "file": raw_name, "clean_name": clean_name, "qty": qty,
                "weight_one": weight_one, "tot_weight": tot_weight,
                "secs_one": secs_one, "tot_secs": tot_secs, "printers": printers_str,
            })

        for dn in printer_history:
            printer_history[dn].sort(key=lambda x: x["create_dt"])

        seen_reprints = set()
        for ev in all_reprints_raw:
            dn = clean_str(ev.get("dev_name") or ev.get("dev_id") or "Unknown")
            f_dt = parse_iso_local(ev.get("finish_time"))
            s_dt = parse_iso_local(ev.get("start_time"))
            if not f_dt:
                continue
            dedup_key = (dn, f_dt.strftime("%Y-%m-%d %H:%M"))
            if dedup_key in seen_reprints:
                continue
            seen_reprints.add(dedup_key)

            is_farm_finish = False
            last_known_task = None
            for h in printer_history.get(dn, []):
                if h["create_dt"] <= f_dt:
                    last_known_task = h
                if h["finish_dt"] and abs((f_dt - h["finish_dt"]).total_seconds()) <= 600:
                    is_farm_finish = True
                    break
                if s_dt and abs((s_dt - h["create_dt"]).total_seconds()) <= 600:
                    is_farm_finish = True
                    break
            if is_farm_finish:
                continue

            raw_ev_file = strip_ext(ev.get("file") or "")
            if raw_ev_file and raw_ev_file.lower() in metrics_by_file:
                raw_name, clean_name, secs_one, weight_one = metrics_by_file[raw_ev_file.lower()]
            elif last_known_task:
                raw_name, clean_name, secs_one, weight_one = (
                    last_known_task["file"], last_known_task["clean_name"],
                    last_known_task["secs_one"], last_known_task["weight_one"]
                )
            else:
                raw_name = raw_ev_file or "Повтор (неизв. файл)"
                clean_name, secs_one, weight_one = parse_task_name(raw_name)
                secs_one, weight_one = secs_one or 0, weight_one or 0.0

            printers_stat[dn]["tasks"] += 1
            printers_stat[dn]["copies"] += 1
            printers_stat[dn]["weight"] += weight_one
            add_printer_time(dn, f_dt, secs_one, is_waste=False)
            parsed.append({
                "dt": f_dt, "date": f_dt.strftime("%Y-%m-%d"), "time": f_dt.strftime("%H:%M"),
                "status": "Повтор с принтера", "is_reprint": True,
                "file": raw_name, "clean_name": clean_name, "qty": 1,
                "weight_one": weight_one, "tot_weight": weight_one,
                "secs_one": secs_one, "tot_secs": secs_one, "printers": dn,
            })

        for item in config.load_manual_reprints():
            m_date, m_printer, m_file, m_qty = item["date"], item["printer"], item["file"], safe_int(item["qty"])
            if m_qty <= 0:
                continue
            try:
                m_dt = datetime.strptime(f"{m_date} 20:00", "%Y-%m-%d %H:%M")
            except ValueError:
                continue
            raw_m = strip_ext(m_file)
            if raw_m.lower() in metrics_by_file:
                raw_name, clean_name, secs_one, weight_one = metrics_by_file[raw_m.lower()]
            else:
                raw_name = raw_m
                clean_name, secs_one, weight_one = parse_task_name(raw_name)
                secs_one, weight_one = secs_one or 0, weight_one or 0.0

            printers_stat[m_printer]["tasks"] += m_qty
            printers_stat[m_printer]["copies"] += m_qty
            printers_stat[m_printer]["weight"] += weight_one * m_qty
            add_printer_time(m_printer, m_dt, secs_one * m_qty, is_waste=False)
            parsed.append({
                "dt": m_dt, "date": m_date, "time": "повторы",
                "status": "Повтор с принтера", "is_reprint": True,
                "file": raw_name, "clean_name": clean_name, "qty": m_qty,
                "weight_one": weight_one, "tot_weight": weight_one * m_qty,
                "secs_one": secs_one, "tot_secs": secs_one * m_qty, "printers": m_printer,
            })

        for d_name in list(devices_hw.keys()) + list(odometers.keys()):
            _ = printers_stat[d_name]

        parsed.sort(key=lambda x: x["dt"], reverse=True)
        return parsed, printers_stat, devices_hw
