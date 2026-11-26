import json
import re
import urllib.error
import urllib.request
from datetime import datetime

import config
import parser

YMD_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def _load_sync_state():
    data = config.load_json(config.GOOGLE_SYNC_STATE_FILE, {"sent_ids": []})
    if not isinstance(data, dict) or not isinstance(data.get("sent_ids"), list):
        return {"sent_ids": []}
    return data


def _save_sync_state(state: dict):
    if len(state["sent_ids"]) > 30000:
        state["sent_ids"] = state["sent_ids"][-30000:]
    config.save_json(config.GOOGLE_SYNC_STATE_FILE, state)


def _post_webhook(url: str, payload: dict, timeout: float = 25.0):
    raw = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=raw,
        headers={
            "Content-Type": "application/json; charset=utf-8",
            "User-Agent": f"BambuFarmControl/{config.VERSION}",
        },
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        text = resp.read().decode("utf-8-sig", errors="replace")
        return json.loads(text)


def _expand_record_to_atomic_events(rec: dict, order_rules: list):
    """
    Раскладывает многостоловые записи (из Tasks или ручных повторов) на атомарные строки (qty=1),
    чтобы при допечатывании новых столов в ongoing-задаче не появлялось дублей в Google Sheets.
    """
    base_id = rec.get("event_id")
    date_s = str(rec.get("date") or "")
    if not base_id or not YMD_RE.match(date_s):
        return []

    ord_hint, ppp_hint = parser.match_order_rule(rec["file"], order_rules)
    qty = max(1, int(rec.get("qty") or 1))
    src = rec.get("source") or "live"

    if src in ("live", "reprint") and qty == 1:
        return [{
            "event_id": base_id,
            "date": date_s,
            "time": rec.get("time", ""),
            "type": rec.get("type", "Готово"),
            "status": rec.get("status", ""),
            "file": rec.get("file", ""),
            "printers": rec.get("printers", ""),
            "qty": 1,
            "weight_good": round(float(rec.get("weight_good", 0.0)), 1),
            "weight_waste": round(float(rec.get("weight_waste", 0.0)), 1),
            "tot_secs": int(rec.get("tot_secs", 0)),
            "order_hint": ord_hint,
            "parts_hint": ppp_hint,
        }]

    unit_wg = round(float(rec.get("weight_good", 0.0)) / qty, 1)
    unit_ww = round(float(rec.get("weight_waste", 0.0)) / qty, 1)
    unit_secs = int(float(rec.get("tot_secs", 0)) / qty)

    items = []
    for i in range(1, qty + 1):
        items.append({
            "event_id": f"{base_id}#{i}",
            "date": date_s,
            "time": rec.get("time", ""),
            "type": rec.get("type", "Готово"),
            "status": rec.get("status", ""),
            "file": rec.get("file", ""),
            "printers": rec.get("printers", ""),
            "qty": 1,
            "weight_good": unit_wg,
            "weight_waste": unit_ww,
            "tot_secs": unit_secs,
            "order_hint": ord_hint,
            "parts_hint": ppp_hint,
        })
    return items


def push_to_google(parsed=None, devices_hw=None, live_active=True, note="ОК", quiet=False):
    with config.get_lock():
        cfg = config.load_config()
        url = (cfg.get("google_webhook_url") or "").strip()
        tok = (cfg.get("google_token") or "").strip()
        loc = (cfg.get("location_name") or "").strip()

        if not url or not tok or not loc:
            return False, "Не задан URL, токен или локация для Google"

        if parsed is None or devices_hw is None:
            parsed, _, devices_hw = parser.load_and_parse(quiet=True)

        parsed = parsed or []
        devices_hw = devices_hw or {}
        order_rules = config.load_orders_map()

        done_cnt = sum(x["qty"] for x in parsed if not x.get("is_failed"))
        fail_cnt = sum(x["qty"] for x in parsed if x.get("is_failed"))

        state = _load_sync_state()
        sent_set = set(state["sent_ids"])

        unsent = []
        for x in parsed:
            for ev in _expand_record_to_atomic_events(x, order_rules):
                if ev["event_id"] not in sent_set:
                    unsent.append(ev)

        status_block = {
            "live": bool(live_active),
            "printers": len(devices_hw),
            "done": done_cnt,
            "fail": fail_cnt,
            "version": config.VERSION,
            "note": note,
        }

        batches = [unsent[i:i + 150] for i in range(0, len(unsent), 150)] or [[]]
        total_added = 0

        try:
            for batch in batches:
                res = _post_webhook(url, {
                    "location": loc,
                    "token": tok,
                    "events": batch,
                    "status": status_block,
                })
                if not isinstance(res, dict) or not res.get("ok"):
                    err = (res or {}).get("error", "unknown_error")
                    if not quiet:
                        print(f"❌ Google Sheets отклонил запрос: {err}")
                    return False, f"Google: {err}"

                total_added += int(res.get("added", 0))
                if batch:
                    for item in batch:
                        state["sent_ids"].append(item["event_id"])
                    _save_sync_state(state)

            msg = f"Облако ОК (+{total_added})" if total_added else "Облако ОК (пинг)"
            if not quiet:
                print(f"✅ Google Sheets синхронизирован ({datetime.now():%H:%M:%S}): добавлено {total_added} новых строк.")
            return True, msg

        except urllib.error.HTTPError as e:
            err_msg = f"HTTP {e.code}"
            if not quiet:
                print(f"❌ Ошибка HTTP при отправке в Google Sheets: {err_msg}")
            return False, f"Google {err_msg}"
        except Exception as e:
            err_msg = str(e)[:50]
            if not quiet:
                print(f"❌ Нет связи с Google Sheets: {err_msg}")
            return False, f"Нет сети ({err_msg})"


def reset_sync_cache():
    config.save_json(config.GOOGLE_SYNC_STATE_FILE, {"sent_ids": []})
    print("✅ Локальный кэш отправленных ID очищен. Следующая синхронизация перепроверит весь архив.")
