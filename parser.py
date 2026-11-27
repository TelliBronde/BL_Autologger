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
    fn_low = (filename or "").lower()
    for r in rules:
        pat = (r.get("pattern") or "").lower()
        if pat and (fn_low == pat or fnmatch.fnmatch(fn_low, pat)):
            return r["order"], int(r.get("parts_per_plate", 1))
    return "Не распределено", 1


def find_history_json_files():
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
    return sorted(set(found), key=lambda p: p.stat().st_mtime)


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


def walk_events(o):
    if isinstance(o, dict):
        if "__events_log__" in o and isinstance(o["__events_log__"], list):
            for item in o["__events_log__"]:
                if isinstance(item, dict):
                    yield item
            return
        for v in o.values():
            yield from walk_events(v)
    elif isinstance(o, list):
        for v in o:
            yield from walk_events(v)


def walk_legacy_reprints(o):
    if isinstance(o, dict):
        if "__reprints_list__" in o and isinstance(o["__reprints_list__"], list):
            for item in o["__reprints_list__"]:
                if isinstance(item, dict):
                    yield item
            return
        for v in o.values():
            yield from walk_legacy_reprints(v)
    elif isinstance(o, list):
        for v in o:
            yield from walk_legacy_reprints(v)


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


def _load_archive():
    default_arch = {"events": {}, "tasks": {}, "metrics": {}, "devices": {}, "legacy_reprints": {}}
    data = config.load_json(config.ARCHIVE_FILE, default_arch)
    if not isinstance(data, dict):
        return default_arch
    for k in default_arch:
        if not isinstance(data.get(k), dict):
            data[k] = {}
    return data


def _save_archive(arch: dict):
    config.save_json(config.ARCHIVE_FILE, arch)


def _merge_raw_payload_into_archive(raw_data, arch: dict) -> bool:
    changed = False

    for t in walk_tasks(raw_data):
        tid = str(t.get("id") or "")
        if not tid:
            continue
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

    for ev in walk_events(raw_data):
        f_dt = parse_iso_local(ev.get("finish_time"))
        if not f_dt:
            continue
        ev_key = f"{clean_str(ev.get('dev_id') or ev.get('dev_name'))}|{f_dt.strftime('%Y-%m-%d %H:%M')}|{ev.get('outcome')}"
        if ev_key not in arch["events"]:
            arch["events"][ev_key] = ev
            changed = True

    for rep in walk_legacy_reprints(raw_data):
        f_dt = parse_iso_local(rep.get("finish_time"))
        if not f_dt:
            continue
        r_key = f"{clean_str(rep.get('dev_name') or rep.get('dev_id'))}|{f_dt.strftime('%Y-%m-%d %H:%M')}"
        if r_key not in arch["legacy_reprints"]:
            arch["legacy_reprints"][r_key] = rep
            changed = True

    for dinfo in walk_devices_info(raw_data):
        dname = clean_str(dinfo.get("name"))
        if dname:
            arch["devices"][dname] = dinfo
            changed = True

    if isinstance(raw_data, list):
        for block in raw_data:
            if isinstance(block, dict) and isinstance(block.get("__metrics_cache__"), dict):
                for fn_k, mval in block["__metrics_cache__"].items():
                    if isinstance(mval, dict) and (safe_float(mval.get("w")) > 0 or safe_int(mval.get("s")) > 0):
                        arch["metrics"][fn_k.lower()] = mval
                        changed = True

    return changed


def sync_live_into_archive(quiet=True, scan_external_dumps=False):
    with config.get_lock():
        arch = _load_archive()
        changed = False
        used_files = []

        live_path = connector.pull_live_data_from_farm()
        if live_path and live_path.exists():
            try:
                raw_live = json.loads(live_path.read_text(encoding="utf-8-sig"))
                if _merge_raw_payload_into_archive(raw_live, arch):
                    changed = True
                used_files.append(live_path.name)
            except Exception:
                pass

        if scan_external_dumps:
            for jp in find_history_json_files():
                if live_path and jp.resolve() == live_path.resolve():
                    continue
                try:
                    raw_dump = json.loads(jp.read_text(encoding="utf-8-sig"))
                    if _merge_raw_payload_into_archive(raw_dump, arch):
                        changed = True
                    used_files.append(jp.name)
                except Exception as e:
                    if not quiet:
                        print(f"⚠️ Пропущен повреждённый файл {jp.name}: {e}")

        if changed:
            _save_archive(arch)

        if not quiet and used_files:
            print(f"📂 Источники ({len(used_files)}): {', '.join(used_files)} | В архиве задач: {len(arch['tasks'])}, живых событий: {len(arch['events'])}")
        return arch


def load_and_parse(quiet=False):
    with config.get_lock():
        cfg = config.load_config()
        start_date_str = (cfg.get("start_date") or "").strip()
        loc_tag = config.sanitize_filename_tag(cfg.get("location_name") or "Farm")

        arch = sync_live_into_archive(quiet=quiet, scan_external_dumps=True)
        odometers = config.load_odometers()
        devices_hw = arch.get("devices", {})
        metrics_by_file = {}
        task_create_dt = {}

        for fn_low, mv in arch.get("metrics", {}).items():
            metrics_by_file[fn_low] = (fn_low, fn_low, safe_int(mv.get("s")), safe_float(mv.get("w")))

        for tid, t in arch.get("tasks", {}).items():
            ct = parse_iso_local(t.get("create_time")) or parse_iso_local(t.get("start_time"))
            if ct:
                task_create_dt[str(tid)] = ct
            f3mf = t.get("f3mf_info") if isinstance(t.get("f3mf_info"), dict) else {}
            raw_name = strip_ext(t.get("task_name") or f3mf.get("f3mf_name") or t.get("name") or "Unknown")
            clean_name, secs_one, weight_one = extract_metrics(t, raw_name)
            if weight_one > 0 or secs_one > 0:
                metrics_by_file[raw_name.lower()] = (raw_name, clean_name, secs_one, weight_one)
                if f3mf.get("f3mf_name"):
                    metrics_by_file[strip_ext(f3mf["f3mf_name"]).lower()] = (raw_name, clean_name, secs_one, weight_one)

        parsed = []
        printer_history = defaultdict(list)
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

        live_done_by_task = defaultdict(int)
        live_fail_by_task = defaultdict(int)
        live_by_printer_dt = defaultdict(list)

        sorted_events = sorted(
            arch.get("events", {}).items(),
            key=lambda kv: parse_iso_local(kv[1].get("finish_time")) or datetime.min
        )
        seen_subtask_done = set()
        seen_subtask_fail = set()
        last_done_by_printer_file = {}
        last_fail_by_printer_file = {}
        bad_keys_to_purge = []

        for ev_key, ev in sorted_events:
            f_dt = parse_iso_local(ev.get("finish_time"))
            if not f_dt:
                continue
            s_dt = parse_iso_local(ev.get("start_time"))
            dur_sec = (f_dt - s_dt).total_seconds() if s_dt else None

            dn = clean_str(ev.get("dev_name") or ev.get("dev_id") or "Unknown")
            dev_id_key = (clean_str(ev.get("dev_id")) or dn).lower()
            outcome = "done" if ev.get("outcome") == "done" else "failed"
            pct = 100 if outcome == "done" else max(1, min(99, safe_int(ev.get("pct"), 50)))
            raw_ev_file = strip_ext(ev.get("file") or "")

            # 1. Отсеиваем служебные калибровки принтера
            if "auto_cali" in raw_ev_file.lower() or "calibration" in raw_ev_file.lower():
                bad_keys_to_purge.append(ev_key)
                continue

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

            tid = str(ev.get("task_id") or "")
            sub = str(ev.get("subtask_id") or "")

            # 2. Отсеиваем фантомный "done" на старте печати (устаревший mc_percent=100% -> 0%)
            if outcome == "done":
                if tid and tid != "0" and tid in task_create_dt:
                    since_create = (f_dt - task_create_dt[tid]).total_seconds()
                    min_real = max(180, int(secs_one * 0.35)) if secs_one > 0 else 180
                    if since_create < min_real:
                        bad_keys_to_purge.append(ev_key)
                        continue
                elif dur_sec is not None and secs_one >= 600 and dur_sec <= 90:
                    bad_keys_to_purge.append(ev_key)
                    continue

            # 3. Жёсткая дедупликация по subtask_id (отдельно для done и failed)
            if sub and sub != "0":
                sub_key = (dev_id_key, tid, sub)
                if outcome == "done":
                    if sub_key in seen_subtask_done:
                        bad_keys_to_purge.append(ev_key)
                        continue
                    seen_subtask_done.add(sub_key)
                else:
                    if sub_key in seen_subtask_fail:
                        bad_keys_to_purge.append(ev_key)
                        continue
                    seen_subtask_fail.add(sub_key)

            # 4. Временная дедупликация по (принтер + файл) для done и failed
            pf_key = (dev_id_key, raw_ev_file.lower())
            if outcome == "done":
                min_gap_sec = max(180, int(secs_one * 0.6)) if secs_one > 0 else 300
                prev_dt = last_done_by_printer_file.get(pf_key)
                if prev_dt and 0 <= (f_dt - prev_dt).total_seconds() < min_gap_sec:
                    bad_keys_to_purge.append(ev_key)
                    continue
                last_done_by_printer_file[pf_key] = f_dt
            else:
                prev_dt = last_fail_by_printer_file.get(pf_key)
                if prev_dt and 0 <= (f_dt - prev_dt).total_seconds() < 600:
                    bad_keys_to_purge.append(ev_key)
                    continue
                last_fail_by_printer_file[pf_key] = f_dt

            date_str = f_dt.strftime("%Y-%m-%d")
            if start_date_str and date_str < start_date_str:
                continue

            ratio = 1.0 if outcome == "done" else (pct / 100.0)
            used_weight = weight_one * ratio
            used_secs = int(secs_one * ratio)

            if tid and tid != "0":
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
                st_text = "Готово (100%)" if (tid and tid != "0") else "Повтор с принтера"
            else:
                printers_stat[dn]["stopped_cnt"] += 1.0
                printers_stat[dn]["weight_waste"] += used_weight
                add_printer_time(dn, f_dt, used_secs, is_waste=True)
                ev_type = "Срыв"
                st_text = f"Срыв на {pct}%"

            parsed.append({
                "event_id": f"{loc_tag}|live|{ev_key}",
                "source": "live",
                "dt": f_dt,
                "date": date_str,
                "time": f_dt.strftime("%H:%M"),
                "type": ev_type,
                "status": st_text,
                "is_failed": (outcome == "failed"),
                "is_reprint": (outcome == "done" and (not tid or tid == "0")),
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

        if bad_keys_to_purge:
            for dk in bad_keys_to_purge:
                arch["events"].pop(dk, None)
            _save_archive(arch)
            if not quiet:
                print(f"🧹 Автоочистка: удалено {len(bad_keys_to_purge)} дублей/фантомных записей из farm_archive.json.")

        for tid, t in arch.get("tasks", {}).items():
            stats = t.get("statistics") if isinstance(t.get("statistics"), dict) else {}
            fin_cnt = safe_int(stats.get("finished"), 1 if not stats else 0)
            stp_cnt = safe_int(stats.get("stopped"))
            printing_now = safe_int(stats.get("printing"))
            total_p = safe_int(stats.get("total"), fin_cnt)

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

            date_str = dt.strftime("%Y-%m-%d") if dt else "Неизвестно"
            if start_date_str and date_str != "Неизвестно" and date_str < start_date_str:
                continue

            pool = t.get("device_pool") or []
            dev_names = [clean_str(d.get("dev_name") or d.get("name") or d.get("dev_id")) for d in pool if isinstance(d, dict)]
            dev_names = [x for x in dev_names if x]
            printers_str = ", ".join(dev_names) or "—"

            for dn in dev_names:
                if c_dt:
                    printer_history[dn].append({
                        "create_dt": c_dt, "finish_dt": dt, "file": raw_name,
                        "clean_name": clean_name, "secs_one": secs_one, "weight_one": weight_one,
                    })

            if t.get("state") == "ongoing" and printing_now > 0 and dev_names:
                for dn in dev_names:
                    hw = devices_hw.get(dn)
                    if hw and hw.get("mc_percent", 0) > 0 and secs_one > 0:
                        add_printer_time(dn, datetime.now(), secs_one * (hw["mc_percent"] / 100.0), is_waste=False)

            matched_done = live_done_by_task.get(tid, 0)
            if matched_done == 0 and dev_names and dt:
                for dn in dev_names:
                    for ev_dt, ev_out in live_by_printer_dt.get(dn, []):
                        if ev_out == "done" and abs((ev_dt - dt).total_seconds()) <= 900:
                            matched_done += 1

            missing_done = max(0, fin_cnt - matched_done)
            if missing_done > 0:
                tot_w = weight_one * missing_done
                tot_s = secs_one * missing_done
                done_devs = [
                    clean_str(d.get("dev_name") or d.get("name") or d.get("dev_id"))
                    for d in pool if isinstance(d, dict) and not d.get("is_in_running_task")
                ]
                target_devs = [x for x in done_devs if x] or dev_names
                if target_devs:
                    for dn in target_devs:
                        printers_stat[dn]["tasks"] += 1
                        printers_stat[dn]["copies"] += missing_done / len(target_devs)
                        printers_stat[dn]["weight_good"] += tot_w / len(target_devs)
                        add_printer_time(dn, dt, tot_s / len(target_devs), is_waste=False)

                st_label = f"В работе ({fin_cnt}/{total_p})" if t.get("state") == "ongoing" else "Готово (история)"
                parsed.append({
                    "event_id": f"{loc_tag}|task|{tid}|done",
                    "source": "task",
                    "dt": dt or datetime.min,
                    "date": date_str,
                    "time": dt.strftime("%H:%M") if dt else "",
                    "type": "Готово",
                    "status": st_label,
                    "is_failed": False,
                    "is_reprint": False,
                    "file": raw_name,
                    "clean_name": clean_name,
                    "qty": missing_done,
                    "weight_one": weight_one,
                    "weight_good": tot_w,
                    "weight_waste": 0.0,
                    "secs_one": secs_one,
                    "tot_secs": tot_s,
                    "printers": printers_str,
                })

            matched_fail = live_fail_by_task.get(tid, 0)
            if matched_fail == 0 and dev_names and c_dt and dt:
                for dn in dev_names:
                    for ev_dt, ev_out in live_by_printer_dt.get(dn, []):
                        if ev_out == "failed" and (c_dt - timedelta(minutes=5)) <= ev_dt <= (dt + timedelta(minutes=15)):
                            matched_fail += 1

            missing_fail = max(0, stp_cnt - matched_fail)
            if missing_fail > 0:
                if total_p <= 1 and c_dt and f_dt and f_dt > c_dt and secs_one > 0:
                    actual_s = max(60.0, min(float(secs_one), (f_dt - c_dt).total_seconds()))
                    ratio = min(0.95, actual_s / float(secs_one))
                else:
                    ratio = 0.5
                est_w = weight_one * ratio * missing_fail
                est_s = int(secs_one * ratio * missing_fail)
                if dev_names:
                    for dn in dev_names:
                        printers_stat[dn]["stopped_cnt"] += missing_fail / len(dev_names)
                        printers_stat[dn]["weight_waste"] += est_w / len(dev_names)
                        add_printer_time(dn, dt or c_dt, est_s / len(dev_names), is_waste=True)

                parsed.append({
                    "event_id": f"{loc_tag}|task|{tid}|fail",
                    "source": "task",
                    "dt": dt or c_dt or datetime.min,
                    "date": date_str,
                    "time": dt.strftime("%H:%M") if dt else "",
                    "type": "Срыв",
                    "status": f"Срыв (история ~{int(ratio * 100)}%)",
                    "is_failed": True,
                    "is_reprint": False,
                    "file": raw_name,
                    "clean_name": clean_name,
                    "qty": missing_fail,
                    "weight_one": weight_one,
                    "weight_good": 0.0,
                    "weight_waste": est_w,
                    "secs_one": secs_one,
                    "tot_secs": est_s,
                    "printers": printers_str,
                })

        for dn in printer_history:
            printer_history[dn].sort(key=lambda x: x["create_dt"])

        for r_key, ev in arch.get("legacy_reprints", {}).items():
            dn = clean_str(ev.get("dev_name") or ev.get("dev_id") or "Unknown")
            f_dt = parse_iso_local(ev.get("finish_time"))
            s_dt = parse_iso_local(ev.get("start_time"))
            if not f_dt:
                continue
            date_str = f_dt.strftime("%Y-%m-%d")
            if start_date_str and date_str < start_date_str:
                continue

            if any(abs((f_dt - ev_dt).total_seconds()) <= 600 for ev_dt, _ in live_by_printer_dt.get(dn, [])):
                continue

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
            printers_stat[dn]["copies"] += 1.0
            printers_stat[dn]["weight_good"] += weight_one
            add_printer_time(dn, f_dt, secs_one, is_waste=False)
            parsed.append({
                "event_id": f"{loc_tag}|rep|{r_key}",
                "source": "reprint",
                "dt": f_dt,
                "date": date_str,
                "time": f_dt.strftime("%H:%M"),
                "type": "Готово",
                "status": "Повтор с принтера",
                "is_failed": False,
                "is_reprint": True,
                "file": raw_name,
                "clean_name": clean_name,
                "qty": 1,
                "weight_one": weight_one,
                "weight_good": weight_one,
                "weight_waste": 0.0,
                "secs_one": secs_one,
                "tot_secs": secs_one,
                "printers": dn,
            })

        for item in config.load_manual_reprints():
            m_date, m_printer, m_file, m_qty = item["date"], item["printer"], item["file"], safe_int(item["qty"])
            if m_qty <= 0 or (start_date_str and m_date < start_date_str):
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
                "event_id": f"{loc_tag}|manual|{m_date}|{m_printer}|{raw_name}",
                "source": "manual",
                "dt": m_dt,
                "date": m_date,
                "time": "ручной",
                "type": "Готово",
                "status": "Повтор (ручной)",
                "is_failed": False,
                "is_reprint": True,
                "file": raw_name,
                "clean_name": clean_name,
                "qty": m_qty,
                "weight_one": weight_one,
                "weight_good": tot_w,
                "weight_waste": 0.0,
                "secs_one": secs_one,
                "tot_secs": tot_s,
                "printers": m_printer,
            })

        for d_name in list(devices_hw.keys()) + list(odometers.keys()):
            _ = printers_stat[d_name]

        parsed.sort(key=lambda x: x["dt"], reverse=True)
        return parsed, printers_stat, devices_hw
