import fnmatch
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


def match_order_rule(filename: str, rules: list):
    """Ищет соответствие файла правилам orders_map.json (поддерживает маски * и ?)."""
    fn_low = (filename or "").lower()
    for r in rules:
        pat = (r.get("pattern") or "").lower()
        if pat and (fn_low == pat or fnmatch.fnmatch(fn_low, pat)):
            return r["order"], int(r.get("parts_per_plate", 1))
    return "Не распределено", 1


def _load_archive():
    default_arch = {"events": {}, "tasks": {}, "metrics": {}, "devices": {}}
    data = config.load_json(config.ARCHIVE_FILE, default_arch)
    if not isinstance(data, dict):
        return default_arch
    for k in default_arch:
        if not isinstance(data.get(k), dict):
            data[k] = {}
    return data


def _save_archive(arch: dict):
    config.save_json(config.ARCHIVE_FILE, arch)


def sync_live_into_archive(quiet=True):
    """Забирает живые данные из открытого окна фермы и сохраняет их в накопительный farm_archive.json."""
    with config.get_lock():
        live_path = connector.pull_live_data_from_farm()
        arch = _load_archive()
        if not live_path or not live_path.exists():
            return arch

        try:
            raw_data = json.loads(live_path.read_text(encoding="utf-8-sig"))
        except Exception:
            return arch

        changed = False
        if isinstance(raw_data, list):
            for block in raw_data:
                if not isinstance(block, dict):
                    continue
                if "hits" in block and isinstance(block["hits"], list):
                    for t in block["hits"]:
                        if not isinstance(t, dict) or not t.get("id"):
                            continue
                        tid = str(t["id"])
                        st = t.get("statistics") if isinstance(t.get("statistics"), dict) else {}
                        new_score = (safe_int(st.get("finished")) + safe_int(st.get("stopped")), str(t.get("finish_time") or ""))
                        old_t = arch["tasks"].get(tid)
                        if old_t:
                            ost = old_t.get("statistics") if isinstance(old_t.get("statistics"), dict) else {}
                            old_score = (safe_int(ost.get("finished")) + safe_int(ost.get("stopped")), str(old_t.get("finish_time") or ""))
                        else:
                            old_score = (-1, "")
                        if new_score >= old_score:
                            arch["tasks"][tid] = t
                            changed = True

                if "__events_log__" in block and isinstance(block["__events_log__"], list):
                    for ev in block["__events_log__"]:
                        if not isinstance(ev, dict) or not ev.get("finish_time"):
                            continue
                        f_dt = parse_iso_local(ev.get("finish_time"))
                        if not f_dt:
                            continue
                        ev_key = f"{clean_str(ev.get('dev_id') or ev.get('dev_name'))}|{f_dt.strftime('%Y-%m-%d %H:%M')}|{ev.get('outcome')}"
                        if ev_key not in arch["events"]:
                            arch["events"][ev_key] = ev
                            changed = True

                if "__devices_info__" in block and isinstance(block["__devices_info__"], dict):
                    for dname, dinfo in block["__devices_info__"].items():
                        if isinstance(dinfo, dict) and dname:
                            arch["devices"][clean_str(dname)] = dinfo
                            changed = True

                if "__metrics_cache__" in block and isinstance(block["__metrics_cache__"], dict):
                    for fn_k, mval in block["__metrics_cache__"].items():
                        if isinstance(mval, dict) and (safe_float(mval.get("w")) > 0 or safe_int(mval.get("s")) > 0):
                            arch["metrics"][fn_k.lower()] = mval
                            changed = True

        if changed:
            _save_archive(arch)
            if not quiet:
                print(f"💾 Архив обновлён: {len(arch['events'])} живых событий, {len(arch['devices'])} принтеров.")
        return arch


def load_and_parse(quiet=False):
    """
    Собирает чистую Live-историю начиная строго с config['start_date'].
    Старая история до start_date игнорируется.
    """
    with config.get_lock():
        cfg = config.load_config()
        start_date_str = cfg.get("start_date") or "2026-12-01"
        arch = sync_live_into_archive(quiet=quiet)

        odometers = config.load_odometers()
        devices_hw = arch.get("devices", {})
        metrics_by_file = {}

        # Заполняем справочник весов и времени из кэша и задач
        for fn_low, mv in arch.get("metrics", {}).items():
            metrics_by_file[fn_low] = (fn_low, fn_low, safe_int(mv.get("s")), safe_float(mv.get("w")))

        for t in arch.get("tasks", {}).values():
            f3mf = t.get("f3mf_info") if isinstance(t.get("f3mf_info"), dict) else {}
            raw_name = strip_ext(t.get("task_name") or f3mf.get("f3mf_name") or t.get("name") or "Unknown")
            clean_name, secs_one, weight_one = extract_metrics(t, raw_name)
            if weight_one > 0 or secs_one > 0:
                metrics_by_file[raw_name.lower()] = (raw_name, clean_name, secs_one, weight_one)
                if f3mf.get("f3mf_name"):
                    metrics_by_file[strip_ext(f3mf["f3mf_name"]).lower()] = (raw_name, clean_name, secs_one, weight_one)

        parsed = []
        printers_stat = defaultdict(lambda: {
            "tasks": 0, "copies": 0.0, "stopped_cnt": 0.0,
            "weight_good": 0.0, "weight_waste": 0.0,
            "secs_good": 0.0, "secs_waste": 0.0,
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

        # Счётчики живых событий по task_id для защиты от двойного учёта со стороны Tasks
        live_done_by_task = defaultdict(int)
        live_fail_by_task = defaultdict(int)
        live_by_printer_dt = defaultdict(list)

        # 1. ПЕРВИЧНЫЙ ИСТОЧНИК: ЖИВЫЕ СОБЫТИЯ (Готово + Срывы с точным %)
        for ev in arch.get("events", {}).values():
            f_dt = parse_iso_local(ev.get("finish_time"))
            if not f_dt:
                continue
            date_str = f_dt.strftime("%Y-%m-%d")
            if date_str < start_date_str:
                continue

            dn = clean_str(ev.get("dev_name") or ev.get("dev_id") or "Unknown")
            outcome = "done" if ev.get("outcome") == "done" else "failed"
            pct = 100 if outcome == "done" else max(1, min(99, safe_int(ev.get("pct"), 50)))
            raw_ev_file = strip_ext(ev.get("file") or "")

            weight_one = safe_float(ev.get("weight_g"))
            secs_one = safe_int(ev.get("pred_secs"))
            clean_name = raw_ev_file or "Неизвестный файл"

            if (weight_one <= 0 or secs_one <= 0) and raw_ev_file.lower() in metrics_by_file:
                ref_raw, ref_clean, s_ref, w_ref = metrics_by_file[raw_ev_file.lower()]
                weight_one = weight_one or w_ref
                secs_one = secs_one or s_ref
                if ref_raw and ref_raw != raw_ev_file.lower():
                    raw_ev_file = ref_raw
                    clean_name = ref_clean

            if weight_one <= 0 or secs_one <= 0:
                c_n, s_n, w_n = parse_task_name(raw_ev_file)
                clean_name = c_n or clean_name
                weight_one = weight_one or (w_n or 0.0)
                secs_one = secs_one or (s_n or 0)

            ratio = 1.0 if outcome == "done" else (pct / 100.0)
            used_weight = weight_one * ratio
            used_secs = int(secs_one * ratio)

            tid = str(ev.get("task_id") or "")
            if tid:
                if outcome == "done":
                    live_done_by_task[tid] += 1
                else:
                    live_fail_by_task[tid] += 1
            live_by_printer_dt[dn].append((f_dt, outcome))

            if outcome == "done":
                printers_stat[dn]["tasks"] += 1
                printers_stat[dn]["copies"] += 1.0
                printers_stat[dn]["weight_good"] += used_weight
                add_printer_time(dn, f_dt, used_secs, is_waste=False)
                ev_type = "Готово"
                st_text = "Готово (100%)" if tid else "Повтор с принтера"
            else:
                printers_stat[dn]["stopped_cnt"] += 1.0
                printers_stat[dn]["weight_waste"] += used_weight
                add_printer_time(dn, f_dt, used_secs, is_waste=True)
                ev_type = "Срыв"
                st_text = f"Срыв на {pct}%"

            parsed.append({
                "dt": f_dt,
                "date": date_str,
                "time": f_dt.strftime("%H:%M"),
                "type": ev_type,
                "status": st_text,
                "is_failed": (outcome == "failed"),
                "is_reprint": (outcome == "done" and not tid),
                "file": raw_ev_file or "Неизвестный файл",
                "clean_name": clean_name,
                "qty": 1,
                "weight_one": weight_one,
                "weight_good": used_weight if outcome == "done" else 0.0,
                "weight_waste": used_weight if outcome == "failed" else 0.0,
                "secs_one": secs_one,
                "tot_secs": used_secs,
                "printers": dn,
            })

        # 2. СТРАХОВКА ИЗ TASKS: если задача завершилась, пока окно клиента было закрыто
        for tid, t in arch.get("tasks", {}).items():
            stats = t.get("statistics") if isinstance(t.get("statistics"), dict) else {}
            fin_cnt = safe_int(stats.get("finished"))
            stp_cnt = safe_int(stats.get("stopped"))

            f3mf = t.get("f3mf_info") if isinstance(t.get("f3mf_info"), dict) else {}
            raw_name = strip_ext(t.get("task_name") or f3mf.get("f3mf_name") or t.get("name") or "Unknown")
            clean_name, secs_one, weight_one = extract_metrics(t, raw_name)

            c_dt = parse_iso_local(t.get("create_time")) or parse_iso_local(t.get("start_time"))
            f_dt = parse_iso_local(t.get("finish_time"))
            expected_f_dt = (c_dt + timedelta(seconds=secs_one)) if (c_dt and secs_one) else c_dt
            dt = f_dt if (f_dt and (not expected_f_dt or (f_dt - expected_f_dt).total_seconds() <= 3 * 3600)) else expected_f_dt
            if not dt:
                continue

            date_str = dt.strftime("%Y-%m-%d")
            if date_str < start_date_str:
                continue

            pool = t.get("device_pool") or []
            dev_names = [clean_str(d.get("dev_name") or d.get("name") or d.get("dev_id")) for d in pool if isinstance(d, dict)]
            dev_names = [x for x in dev_names if x]
            printers_str = ", ".join(dev_names) or "—"

            # Проверяем сколько финишей уже поймал живой хук (по task_id или по совпадению принтера и времени)
            matched_done = live_done_by_task.get(tid, 0)
            if matched_done == 0 and dev_names:
                for dn in dev_names:
                    for ev_dt, ev_out in live_by_printer_dt.get(dn, []):
                        if ev_out == "done" and abs((ev_dt - dt).total_seconds()) <= 900:
                            matched_done += 1

            missing_done = max(0, fin_cnt - matched_done)
            if missing_done > 0:
                tot_w = weight_one * missing_done
                tot_s = secs_one * missing_done
                if dev_names:
                    for dn in dev_names:
                        printers_stat[dn]["tasks"] += 1
                        printers_stat[dn]["copies"] += missing_done / len(dev_names)
                        printers_stat[dn]["weight_good"] += tot_w / len(dev_names)
                        add_printer_time(dn, dt, tot_s / len(dev_names), is_waste=False)
                parsed.append({
                    "dt": dt, "date": date_str, "time": dt.strftime("%H:%M"),
                    "type": "Готово", "status": "Готово (фоново)",
                    "is_failed": False, "is_reprint": False,
                    "file": raw_name, "clean_name": clean_name, "qty": missing_done,
                    "weight_one": weight_one, "weight_good": tot_w, "weight_waste": 0.0,
                    "secs_one": secs_one, "tot_secs": tot_s, "printers": printers_str,
                })

            matched_fail = live_fail_by_task.get(tid, 0)
            if matched_fail == 0 and dev_names:
                for dn in dev_names:
                    for ev_dt, ev_out in live_by_printer_dt.get(dn, []):
                        if ev_out == "failed" and c_dt and (c_dt - timedelta(minutes=5)) <= ev_dt <= (dt + timedelta(minutes=15)):
                            matched_fail += 1

            missing_fail = max(0, stp_cnt - matched_fail)
            if missing_fail > 0:
                est_w = weight_one * 0.5 * missing_fail
                est_s = int(secs_one * 0.5 * missing_fail)
                if dev_names:
                    for dn in dev_names:
                        printers_stat[dn]["stopped_cnt"] += missing_fail / len(dev_names)
                        printers_stat[dn]["weight_waste"] += est_w / len(dev_names)
                        add_printer_time(dn, dt, est_s / len(dev_names), is_waste=True)
                parsed.append({
                    "dt": dt, "date": date_str, "time": dt.strftime("%H:%M"),
                    "type": "Срыв", "status": "Срыв (офлайн ~50%)",
                    "is_failed": True, "is_reprint": False,
                    "file": raw_name, "clean_name": clean_name, "qty": missing_fail,
                    "weight_one": weight_one, "weight_good": 0.0, "weight_waste": est_w,
                    "secs_one": secs_one, "tot_secs": est_s, "printers": printers_str,
                })

        # 3. РУЧНЫЕ ПОВТОРЫ (если добавлены через меню)
        for item in config.load_manual_reprints():
            m_date, m_printer, m_file, m_qty = item["date"], item["printer"], item["file"], safe_int(item["qty"])
            if m_qty <= 0 or m_date < start_date_str:
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

            tot_w = weight_one * m_qty
            tot_s = secs_one * m_qty
            printers_stat[m_printer]["tasks"] += m_qty
            printers_stat[m_printer]["copies"] += m_qty
            printers_stat[m_printer]["weight_good"] += tot_w
            add_printer_time(m_printer, m_dt, tot_s, is_waste=False)
            parsed.append({
                "dt": m_dt, "date": m_date, "time": "ручной",
                "type": "Готово", "status": "Повтор (ручной)",
                "is_failed": False, "is_reprint": True,
                "file": raw_name, "clean_name": clean_name, "qty": m_qty,
                "weight_one": weight_one, "weight_good": tot_w, "weight_waste": 0.0,
                "secs_one": secs_one, "tot_secs": tot_s, "printers": m_printer,
            })

        for d_name in list(devices_hw.keys()) + list(odometers.keys()):
            _ = printers_stat[d_name]

        parsed.sort(key=lambda x: x["dt"], reverse=True)
        return parsed, printers_stat, devices_hw
