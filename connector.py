import json
import os
import subprocess
import sys
import tempfile
import time
import urllib.request
from pathlib import Path

import websocket
import config

# JS-радар v7:
# - Конечный автомат по /devices2 (PAUSE не считается срывом, порог готово/срыв = 95%)
# - Кэш метрик файлов (вес/время) для повторов с экрана и расчёта веса срывов
# - Ротация localStorage
# - Двусторонний статус-бар Heartbeat (Зелёный / Жёлтый / Красный баннер)
JS_HOOK = r"""
(() => {
  const THRESHOLD = 95;
  window.__tasksMap = JSON.parse(localStorage.getItem('__bambu_tasks_v3') || '{}');
  window.__eventsLog = JSON.parse(localStorage.getItem('__bambu_events_v3') || '[]');
  window.__devicesInfo = JSON.parse(localStorage.getItem('__bambu_devices_v3') || '{}');
  window.__metricsCache = JSON.parse(localStorage.getItem('__bambu_metrics_v3') || '{}');
  window.__pState = window.__pState || {};
  window.__bambuHb = window.__bambuHb || {
    ts: 0, live: false, ok: true, msg: 'Ожидание подключения...', loc: '', last_save: '—', done: 0, fail: 0
  };

  const MAX_TASKS = 1500;
  const MAX_EVENTS = 2500;

  const stripExt = (n) => String(n || '').trim().replace(/\.(gcode\.3mf|3mf|gcode)$/i, '');

  const extractTaskMetrics = (t) => {
    let w = 0, s = 0;
    try {
      const plates = t && t.f3mf_info && t.f3mf_info.slice_info_config && t.f3mf_info.slice_info_config.plates;
      if (Array.isArray(plates) && plates.length > 0) {
        const p0 = plates[0];
        if (Array.isArray(p0.filaments)) {
          w = p0.filaments.reduce((acc, x) => acc + (Number(x.used_g) || 0), 0);
        }
        if (!w && p0.metadata && p0.metadata.weight) w = Number(p0.metadata.weight) || 0;
        if (p0.metadata && p0.metadata.prediction) s = Number(p0.metadata.prediction) || 0;
      }
    } catch (e) {}
    return { w, s };
  };

  const cacheTaskMetrics = (t) => {
    if (!t || typeof t !== 'object') return;
    const { w, s } = extractTaskMetrics(t);
    if (w <= 0 && s <= 0) return;
    const n1 = stripExt(t.task_name || t.name || '').toLowerCase();
    const n2 = stripExt(t.f3mf_info && t.f3mf_info.f3mf_name || '').toLowerCase();
    if (n1) window.__metricsCache[n1] = { w, s };
    if (n2) window.__metricsCache[n2] = { w, s };
  };

  const compactTask = (t) => ({
    id: t.id,
    task_name: t.task_name || t.name,
    state: t.state,
    create_time: t.create_time,
    start_time: t.start_time,
    finish_time: t.finish_time,
    statistics: t.statistics,
    device_pool: Array.isArray(t.device_pool) ? t.device_pool.map(d => ({
      dev_id: d.dev_id,
      dev_name: d.dev_name || d.name,
      gcode_state: d.gcode_state,
      is_in_running_task: d.is_in_running_task
    })) : [],
    f3mf_info: t.f3mf_info ? {
      f3mf_name: t.f3mf_info.f3mf_name,
      slice_info_config: t.f3mf_info.slice_info_config
    } : {}
  });

  const trimStorage = () => {
    if (window.__eventsLog.length > MAX_EVENTS) {
      window.__eventsLog = window.__eventsLog.slice(-MAX_EVENTS);
    }
    const keys = Object.keys(window.__tasksMap);
    if (keys.length > MAX_TASKS) {
      keys.sort((a, b) => Number(a) - Number(b));
      keys.slice(0, keys.length - MAX_TASKS).forEach(k => delete window.__tasksMap[k]);
    }
  };

  const saveAll = () => {
    try {
      trimStorage();
      localStorage.setItem('__bambu_tasks_v3', JSON.stringify(window.__tasksMap));
      localStorage.setItem('__bambu_events_v3', JSON.stringify(window.__eventsLog));
      localStorage.setItem('__bambu_devices_v3', JSON.stringify(window.__devicesInfo));
      localStorage.setItem('__bambu_metrics_v3', JSON.stringify(window.__metricsCache));
    } catch (e) {
      try {
        window.__eventsLog = window.__eventsLog.slice(-800);
        localStorage.setItem('__bambu_events_v3', JSON.stringify(window.__eventsLog));
      } catch (_) {}
    }
    if (window.__updateBambuUI) window.__updateBambuUI();
  };

  const emitEvent = (prev, id, devName, outcome, pctVal, fallbackFile, fallbackTaskId) => {
    const fname = prev.file || fallbackFile || '';
    const cached = window.__metricsCache[stripExt(fname).toLowerCase()] || {};
    const weightG = prev.weight_g || cached.w || 0;
    const predSecs = prev.pred_secs || cached.s || 0;
    window.__eventsLog.push({
      dev_id: id,
      dev_name: devName,
      outcome: outcome, // 'done' | 'failed'
      pct: outcome === 'done' ? 100 : Math.max(1, Math.min(99, Math.round(pctVal || 1))),
      start_time: prev.start_time || null,
      finish_time: new Date().toISOString(),
      file: fname,
      task_id: prev.task_id || fallbackTaskId || '',
      subtask_id: prev.subtask_id || '',
      weight_g: weightG,
      pred_secs: predSecs
    });
  };

  const processDevicesArray = (devices) => {
    if (!Array.isArray(devices)) return;
    let changed = false;

    devices.forEach(d => {
      if (!d || typeof d !== 'object') return;
      const id = d.dev_id || d.sn;
      const rs = d.report_status || {};
      const state = String(rs.gcode_state || d.gcode_state || '').toUpperCase();
      if (!id || !state) return;

      const devName = (d.name && d.name.trim()) || (d.dev_name && d.dev_name.trim()) || id;
      const pct = Number(rs.mc_percent ?? 0);
      const remMin = Number(rs.mc_remaining_time ?? 0);
      const subtaskId = String(rs.subtask_id || '');
      const taskId = String(rs.task_id || (d.running_task && d.running_task.id) || '');
      const fileName = rs.subtask_name || rs.gcode_file || (d.running_task && d.running_task.task_name) || '';

      let taskWeight = 0, taskSecs = 0;
      if (d.running_task && d.running_task.id) {
        window.__tasksMap[String(d.running_task.id)] = compactTask(d.running_task);
        cacheTaskMetrics(d.running_task);
        const m = extractTaskMetrics(d.running_task);
        taskWeight = m.w;
        taskSecs = m.s;
        changed = true;
      }

      const prevInfo = window.__devicesInfo[devName] || {};
      window.__devicesInfo[devName] = {
        dev_id: id,
        name: devName,
        model: d.dev_model === 'N2S' ? 'Bambu A1' : (d.dev_model === 'C12' ? 'Bambu P1S' : (d.dev_model || prevInfo.model || '')),
        ip: d.dev_ip || prevInfo.ip || '',
        fw: (d.upgrade_info && d.upgrade_info.current_ver) || prevInfo.fw || '',
        nozzle: rs.nozzle_type ? `${rs.nozzle_diameter || '0.4'} (${rs.nozzle_type === 'hardened_steel' ? 'Hardened' : 'Stainless'})` : (prevInfo.nozzle || ''),
        wifi: rs.wifi_signal || prevInfo.wifi || '',
        hms_errors: Array.isArray(rs.hms) ? rs.hms.length : 0,
        state: state,
        mc_percent: pct,
        mc_remaining_time: remMin,
        current_file: fileName
      };

      const prev = window.__pState[id] || { active: false, max_pct: 0, state: '' };
      let next = Object.assign({}, prev, { state, pct, dev_name: devName });

      if (state === 'RUNNING' || state === 'PREPARE') {
        if (!prev.active) {
          next.active = true;
          next.start_time = new Date().toISOString();
          next.max_pct = pct;
          next.file = fileName;
          next.task_id = taskId;
          next.subtask_id = subtaskId;
          next.weight_g = taskWeight;
          next.pred_secs = taskSecs;
        } else {
          // Быстрый перезапуск печати поверх завершённой (без промежуточного IDLE/FINISH)
          if (prev.max_pct >= THRESHOLD && pct <= 15) {
            emitEvent(prev, id, devName, 'done', 100, fileName, taskId);
            changed = true;
            next.start_time = new Date().toISOString();
            next.max_pct = pct;
            next.file = fileName;
            next.task_id = taskId;
            next.subtask_id = subtaskId;
            next.weight_g = taskWeight;
            next.pred_secs = taskSecs;
          } else if (prev.max_pct >= 1 && prev.max_pct < THRESHOLD && pct <= 5 && (prev.max_pct - pct) >= 15 && state === 'RUNNING') {
            // Быстрый перезапуск после срыва
            emitEvent(prev, id, devName, 'failed', prev.max_pct, fileName, taskId);
            changed = true;
            next.start_time = new Date().toISOString();
            next.max_pct = pct;
            next.file = fileName;
            next.task_id = taskId;
            next.subtask_id = subtaskId;
            next.weight_g = taskWeight;
            next.pred_secs = taskSecs;
          } else {
            next.max_pct = Math.max(prev.max_pct || 0, pct);
            if (fileName) next.file = fileName;
            if (taskId) next.task_id = taskId;
            if (subtaskId) next.subtask_id = subtaskId;
            if (taskWeight > 0) next.weight_g = taskWeight;
            if (taskSecs > 0) next.pred_secs = taskSecs;
          }
        }
      } else if (state === 'PAUSE') {
        // Пауза — не срыв! Сохраняем активный контекст и ждём возобновления или отмены
        next.active = prev.active;
        next.max_pct = Math.max(prev.max_pct || 0, pct);
      } else if (state === 'FINISH') {
        if (prev.active) {
          emitEvent(prev, id, devName, 'done', 100, fileName, taskId);
          changed = true;
        }
        next.active = false;
        next.max_pct = 0;
      } else if (state === 'FAILED') {
        if (prev.active) {
          const outcome = (prev.max_pct >= THRESHOLD) ? 'done' : 'failed';
          emitEvent(prev, id, devName, outcome, prev.max_pct || pct, fileName, taskId);
          changed = true;
        }
        next.active = false;
        next.max_pct = 0;
      } else if (state === 'IDLE') {
        if (prev.active) {
          if (prev.max_pct >= THRESHOLD) {
            emitEvent(prev, id, devName, 'done', 100, fileName, taskId);
            changed = true;
          } else if (prev.max_pct >= 1) {
            emitEvent(prev, id, devName, 'failed', prev.max_pct, fileName, taskId);
            changed = true;
          }
        }
        next.active = false;
        next.max_pct = 0;
      }

      window.__pState[id] = next;
    });

    if (changed) saveAll();
    else if (window.__updateBambuUI) window.__updateBambuUI();
  };

  const findHits = (obj) => {
    if (!obj || typeof obj !== 'object') return null;
    if (Array.isArray(obj.hits) && obj.hits.length > 0 && obj.hits[0].task_name !== undefined) return obj.hits;
    if (Array.isArray(obj)) {
      if (obj.length > 0 && obj[0] && obj[0].task_name !== undefined) return obj;
      for (const item of obj) { const r = findHits(item); if (r) return r; }
      return null;
    }
    for (const v of Object.values(obj)) {
      if (v && typeof v === 'object') { const r = findHits(v); if (r) return r; }
    }
    return null;
  };

  window.__bambuProcessPayload = (u, t) => {
    try {
      if (!t || typeof t !== 'string') return;
      const isDevices2 = u.includes('devices2') || t.includes('"report_status"');
      const hasTasks = t.includes('"task_name"');
      if (!isDevices2 && !hasTasks) return;

      const data = JSON.parse(t);
      if (isDevices2 && data && Array.isArray(data.devices)) {
        processDevicesArray(data.devices);
        return;
      }
      if (hasTasks) {
        const hits = findHits(data);
        if (hits) {
          hits.forEach(task => {
            if (task && task.id) {
              window.__tasksMap[String(task.id)] = compactTask(task);
              cacheTaskMetrics(task);
            }
          });
          saveAll();
        }
      }
    } catch (e) {}
  };

  if (!window.__bambuHookedV7) {
    window.__bambuHookedV7 = true;
    const oOpen = XMLHttpRequest.prototype.open;
    XMLHttpRequest.prototype.open = function (m, u) {
      this.__u = String(u);
      this.addEventListener('load', () => {
        try {
          let t = (this.responseType === '' || this.responseType === 'text')
            ? this.responseText
            : (typeof this.response === 'string' ? this.response : JSON.stringify(this.response));
          window.__bambuProcessPayload(this.__u, t);
        } catch (e) {}
      });
      return oOpen.apply(this, arguments);
    };

    const oFetch = window.fetch;
    window.fetch = function (u, o) {
      const url = String(typeof u === 'string' ? u : (u && u.url) || '');
      return oFetch.apply(this, arguments).then(r => {
        try { r.clone().text().then(t => window.__bambuProcessPayload(url, t)).catch(() => {}); } catch (e) {}
        return r;
      });
    };
  }

  window.__bambuSetHeartbeat = (payload) => {
    window.__bambuHb = Object.assign(window.__bambuHb || {}, payload || {}, { ts: Date.now() });
    if (window.__updateBambuUI) window.__updateBambuUI();
  };

  let box = document.getElementById('__bambu_logger_box');
  if (!box && document.body) {
    box = document.createElement('div');
    box.id = '__bambu_logger_box';
    box.style.cssText = 'position:fixed;bottom:14px;right:14px;z-index:999999;background:#1F3864;color:#fff;padding:9px 14px;border-radius:8px;font-family:Arial,sans-serif;font-size:12px;font-weight:600;box-shadow:0 4px 16px rgba(0,0,0,0.4);display:flex;align-items:center;gap:10px;border:2px solid rgba(255,255,255,0.2);transition:background 0.3s;';
    const badge = document.createElement('span');
    box.appendChild(badge);
    document.body.appendChild(box);

    window.__updateBambuUI = () => {
      const hb = window.__bambuHb || {};
      const ageSec = hb.ts ? Math.floor((Date.now() - hb.ts) / 1000) : 999999;
      const pCount = Object.keys(window.__devicesInfo || {}).length;
      const locTag = hb.loc ? `[${hb.loc}] ` : '';

      if (!hb.ts || !hb.live) {
        box.style.background = '#b8860b';
        box.style.borderColor = '#ffd700';
        badge.textContent = `🟡 ${locTag}LIVE ВЫКЛ — включите Live-запись в Farm_Control | 🖨️ Принтеров: ${pCount}`;
      } else if (hb.ok === false) {
        box.style.background = '#c0392b';
        box.style.borderColor = '#ff7675';
        badge.textContent = `🔴 ${locTag}СБОЙ АВТОЛОГГЕРА: ${hb.msg || 'Ошибка записи'} — проверьте Farm_Control!`;
      } else if (ageSec > 300) {
        box.style.background = '#c0392b';
        box.style.borderColor = '#ff7675';
        badge.textContent = `🔴 ${locTag}СБОЙ: скрипт не отвечает ${Math.floor(ageSec / 60)} мин! Перезапустите Farm_Control`;
      } else if (ageSec > 120) {
        box.style.background = '#d35400';
        box.style.borderColor = '#fdcb6e';
        badge.textContent = `🟡 ${locTag}Задержка пульса (${ageSec}с) | Посл. запись: ${hb.last_save || '—'} | 🖨️ ${pCount}`;
      } else {
        box.style.background = '#1e7e34';
        box.style.borderColor = '#2ecc71';
        const extra = hb.msg && hb.msg !== 'ОК' ? ` | ⚠️ ${hb.msg}` : '';
        badge.textContent = `🟢 ${locTag}LIVE OK | Запись: ${hb.last_save || '—'} | ✅ ${hb.done ?? 0} | ❌ Срывы: ${hb.fail ?? 0} | 🖨️ ${pCount}${extra}`;
      }
    };
    window.__updateBambuUI();

    if (!window.__bambuHbTimer) {
      window.__bambuHbTimer = setInterval(() => {
        if (window.__updateBambuUI) window.__updateBambuUI();
      }, 10000);
    }
  }
  return Object.keys(window.__eventsLog).length;
})();
"""


def is_farm_process_running() -> bool:
    try:
        out = subprocess.check_output(
            ["tasklist", "/FI", f"IMAGENAME eq {config.FARM_EXE_NAME}", "/NH"],
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            stderr=subprocess.DEVNULL,
        ).decode("cp866", errors="ignore")
        return "Bambu Farm Manager" in out
    except Exception:
        return False


def kill_farm_processes():
    subprocess.run(
        ["taskkill", "/F", "/IM", config.FARM_EXE_NAME, "/T"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    time.sleep(1.0)


def check_port_network_exposure() -> str:
    try:
        out = subprocess.check_output(
            ["netstat", "-ano", "-p", "TCP"],
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            stderr=subprocess.DEVNULL,
        ).decode("ascii", errors="ignore")
        for line in out.splitlines():
            if "LISTENING" not in line:
                continue
            parts = line.split()
            if len(parts) < 2:
                continue
            local_addr = parts[1]
            for port in config.CDP_SCAN_PORTS:
                if local_addr.endswith(f":{port}"):
                    ip_part = local_addr.rsplit(":", 1)[0]
                    if ip_part not in ("127.0.0.1", "[::1]"):
                        return f"🛑 ВНИМАНИЕ: Порт {port} открыт в сеть ({local_addr})!"
    except Exception:
        pass
    return ""


def _is_valid_farm_target(port: int, target: dict) -> bool:
    t_type = target.get("type", "")
    t_url = (target.get("url") or "").lower()
    t_title = (target.get("title") or "").lower()
    if t_type != "page" or "devtools://" in t_url or "webSocketDebuggerUrl" not in target:
        return False

    is_electron = False
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/json/version", timeout=0.8) as resp:
            ver_info = json.loads(resp.read().decode("utf-8"))
            ua = (ver_info.get("User-Agent") or "").lower()
            if "electron" in ua:
                is_electron = True
    except Exception:
        pass

    has_marker = any(m in t_title or m in t_url for m in config.FARM_MARKERS)
    return is_electron or has_marker


def get_cdp_ws_url():
    if not is_farm_process_running():
        return None
    for port in config.CDP_SCAN_PORTS:
        try:
            with urllib.request.urlopen(f"http://127.0.0.1:{port}/json", timeout=1.0) as resp:
                targets = json.loads(resp.read().decode("utf-8"))
            for t in targets:
                if _is_valid_farm_target(port, t):
                    return t["webSocketDebuggerUrl"]
        except Exception:
            pass
    return None


def cdp_eval(expression, timeout=5.0, bring_to_front=False):
    ws_url = get_cdp_ws_url()
    if not ws_url:
        return None
    try:
        ws = websocket.create_connection(ws_url, timeout=timeout, suppress_origin=True)
        if bring_to_front:
            ws.send(json.dumps({"id": 0, "method": "Page.bringToFront"}))
            ws.recv()
        msg = {"id": 1, "method": "Runtime.evaluate", "params": {"expression": expression, "returnByValue": True}}
        ws.send(json.dumps(msg))
        raw = ws.recv()
        ws.close()
        res = json.loads(raw)
        return res.get("result", {}).get("result", {}).get("value")
    except Exception:
        return None


def send_heartbeat(live: bool, ok: bool = True, msg: str = "ОК", last_save: str = "—", done: int = 0, fail: int = 0):
    """Отправляет пульс и статистику в статус-бар окна Bambu Farm Manager."""
    if not get_cdp_ws_url():
        return False
    cfg = config.load_config()
    payload = {
        "live": bool(live),
        "ok": bool(ok),
        "msg": str(msg or ""),
        "loc": cfg.get("location_name") or "",
        "last_save": str(last_save or "—"),
        "done": int(done or 0),
        "fail": int(fail or 0),
    }
    js = f"if(!window.__bambuHookedV7){{{JS_HOOK}}} if(window.__bambuSetHeartbeat){{window.__bambuSetHeartbeat({json.dumps(payload, ensure_ascii=False)});}}"
    return cdp_eval(js, timeout=4.0) is not None or True


def pull_live_data_from_farm():
    with config.get_lock():
        if not get_cdp_ws_url():
            return None
        cdp_eval(JS_HOOK)
        expr = (
            "JSON.stringify(["
            "{total: Object.keys(window.__tasksMap||{}).length, hits: Object.values(window.__tasksMap||{})},"
            "{__events_log__: window.__eventsLog||[]},"
            "{__devices_info__: window.__devicesInfo||{}},"
            "{__metrics_cache__: window.__metricsCache||{}}"
            "])"
        )
        raw_json = cdp_eval(expr, timeout=10.0)
        if raw_json:
            try:
                data = json.loads(raw_json)
                if isinstance(data, list) and len(data) >= 3:
                    has_any = (
                        len(data[0].get("hits", [])) > 0
                        or len(data[1].get("__events_log__", [])) > 0
                        or len(data[2].get("__devices_info__", {})) > 0
                    )
                    if has_any:
                        tmp = config.LIVE_DUMP_FILE.with_suffix(".tmp.json")
                        tmp.write_text(raw_json, encoding="utf-8")
                        tmp.replace(config.LIVE_DUMP_FILE)
                        return config.LIVE_DUMP_FILE
            except Exception:
                pass
        return None


def start_farm_with_hook(auto_mode=False):
    proc_alive = is_farm_process_running()
    ws_url = get_cdp_ws_url()

    if proc_alive and ws_url:
        cnt = cdp_eval(JS_HOOK, bring_to_front=(not auto_mode))
        if not auto_mode:
            print(f"✅ Ферма уже подключена! (Событий в памяти: {cnt})")
            ans = input("Перезапустить ферму с нуля? [R = Перезапустить / Enter = Оставить]: ").strip()
            if ans.upper() != "R":
                return True
            kill_farm_processes()
        else:
            return True
    elif proc_alive and not ws_url:
        if not auto_mode:
            print("⚠️ Ферма запущена без порта автологгера. Перезапускаем...")
        kill_farm_processes()

    exe_path = config.FARM_EXE
    if not exe_path.exists():
        if auto_mode:
            return False
        raw = input("Укажи путь к Bambu Farm Manager Client.exe:\n> ").strip().strip('"')
        exe_path = Path(raw)
    if not exe_path.exists():
        if not auto_mode:
            print("❌ Файл клиента фермы не найден!")
        return False

    if not auto_mode:
        print(f"🚀 Запускаем {exe_path.name} на 127.0.0.1:{config.CDP_PORT}...")
    subprocess.Popen(
        [str(exe_path), f"--remote-debugging-port={config.CDP_PORT}"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )

    for _ in range(30):
        time.sleep(1)
        if get_cdp_ws_url():
            time.sleep(1.5)
            cnt = cdp_eval(JS_HOOK, bring_to_front=(not auto_mode))
            if cnt is not None:
                if not auto_mode:
                    print(f"✅ Ферма открыта, автологгер подключён! (Событий в памяти: {cnt})")
                    warn = check_port_network_exposure()
                    if warn:
                        print(warn)
                return True
    if not auto_mode:
        print("❌ Не удалось дождаться ответа от окна фермы.")
    return False


# ==============================================================================
# АВТОНАСТРОЙКА СИСТЕМЫ: АВТОЗАГРУЗКА WINDOWS И ПАТЧ ЯРЛЫКОВ BAMBU
# ==============================================================================
PS_SHORTCUT_SCRIPT = r"""
$ErrorActionPreference = 'Stop'
$mode = $env:BF_MODE
$portFlag = "--remote-debugging-port=$env:BF_PORT"
$farmExe = $env:BF_FARM_EXE.ToLower()
$startupLnk = Join-Path ([Environment]::GetFolderPath('Startup')) 'Farm_Control_Auto.lnk'
$wsh = New-Object -ComObject WScript.Shell

if ($mode -eq 'install_startup') {
    $sc = $wsh.CreateShortcut($startupLnk)
    $sc.TargetPath = $env:BF_TARGET
    $sc.Arguments = $env:BF_ARGS
    $sc.WorkingDirectory = $env:BF_WORKDIR
    $sc.Description = 'Bambu Farm Control Auto-Start'
    $sc.Save()
    Write-Output "STARTUP_OK:$startupLnk"
    exit 0
}

if ($mode -eq 'remove_startup') {
    if (Test-Path $startupLnk) { Remove-Item $startupLnk -Force }
    Write-Output "STARTUP_REMOVED"
    exit 0
}

$dirs = @(
    [Environment]::GetFolderPath('Desktop'),
    [Environment]::GetFolderPath('CommonDesktopDirectory'),
    [Environment]::GetFolderPath('Programs'),
    [Environment]::GetFolderPath('CommonPrograms'),
    (Join-Path $env:APPDATA 'Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar')
) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -Unique

$patched = 0
$needAdmin = @()

foreach ($d in $dirs) {
    $lnks = Get-ChildItem -Path $d -Filter '*.lnk' -Recurse -ErrorAction SilentlyContinue
    foreach ($f in $lnks) {
        try {
            $sc = $wsh.CreateShortcut($f.FullName)
            $tp = ($sc.TargetPath + '').ToLower()
            if ($tp.EndsWith($farmExe)) {
                $args = ($sc.Arguments + '').Trim()
                # Чистим старый небезопасный флаг и старые порты
                $cleanArgs = ($args -replace '--remote-allow-origins=\S*', '' -replace '--remote-debugging-port=\d+', '').Trim()
                if ($mode -eq 'patch_bambu') {
                    $newArgs = if ($cleanArgs) { "$cleanArgs $portFlag" } else { $portFlag }
                } else {
                    $newArgs = $cleanArgs
                }
                if ($args -ne $newArgs) {
                    $sc.Arguments = $newArgs
                    try {
                        $sc.Save()
                        $patched++
                        Write-Output "PATCHED:$($f.FullName)"
                    } catch {
                        $needAdmin += $f.FullName
                    }
                } else {
                    Write-Output "ALREADY_OK:$($f.FullName)"
                }
            }
        } catch {}
    }
}
Write-Output "SUMMARY:$patched|$($needAdmin.Count)"
if ($needAdmin.Count -gt 0) {
    Write-Output ("ADMIN_NEEDED:" + ($needAdmin -join ';'))
}
"""


def _run_ps_shortcut_helper(mode: str, elevate: bool = False):
    if sys.platform != "win32":
        return ""
    env = os.environ.copy()
    env["BF_MODE"] = mode
    env["BF_PORT"] = str(config.CDP_PORT)
    env["BF_FARM_EXE"] = config.FARM_EXE_NAME
    env["BF_WORKDIR"] = str(config.HERE)

    if getattr(sys, "frozen", False):
        env["BF_TARGET"] = sys.executable
        env["BF_ARGS"] = "--auto"
    else:
        env["BF_TARGET"] = sys.executable
        env["BF_ARGS"] = f'"{config.HERE / "main.py"}" --auto'

    ps_file = None
    try:
        with tempfile.NamedTemporaryFile("w", suffix=".ps1", delete=False, encoding="utf-8-sig") as tf:
            tf.write(PS_SHORTCUT_SCRIPT)
            ps_file = Path(tf.name)

        if not elevate:
            res = subprocess.run(
                ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(ps_file)],
                env=env,
                capture_output=True,
                text=True,
                timeout=15,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            return res.stdout or ""
        else:
            # Передаём только безопасные env-переменные через wrapper-вызов с UAC
            wrapper = (
                f"$env:BF_MODE='{mode}';"
                f"$env:BF_PORT='{config.CDP_PORT}';"
                f"$env:BF_FARM_EXE='{config.FARM_EXE_NAME}';"
                f"& '{str(ps_file)}'"
            )
            subprocess.run(
                [
                    "powershell", "-NoProfile", "-Command",
                    f"Start-Process powershell -Verb RunAs -Wait -ArgumentList '-NoProfile -ExecutionPolicy Bypass -Command \"{wrapper}\"'"
                ],
                timeout=30,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            return "ELEVATED_DONE"
    except Exception as e:
        return f"ERROR:{e}"
    finally:
        if ps_file and ps_file.exists():
            try:
                ps_file.unlink()
            except Exception:
                pass


def setup_autostart_and_shortcuts(interactive=True):
    """Прописывает Farm_Control --auto в Автозагрузку и добавляет порт 19222 в ярлыки Bambu."""
    out1 = _run_ps_shortcut_helper("install_startup", elevate=False)
    if interactive:
        if "STARTUP_OK" in out1:
            print("  ✅ Автозагрузка Windows настроена (Farm_Control_Auto.lnk).")
        else:
            print(f"  ⚠️ Автозагрузка: {out1.strip()}")

    out2 = _run_ps_shortcut_helper("patch_bambu", elevate=False)
    if interactive:
        for line in out2.splitlines():
            if line.startswith("PATCHED:"):
                print(f"  ✅ Обновлён ярлык фермы: {line.split(':', 1)[1]}")
            elif line.startswith("ALREADY_OK:"):
                print(f"  ℹ️ Ярлык уже настроен: {line.split(':', 1)[1]}")

    if "ADMIN_NEEDED:" in out2 and interactive:
        print("\n🔐 Часть ярлыков Bambu лежит в общей папке (Public Desktop / ProgramData).")
        ans = input("Запросить права администратора (окно UAC), чтобы обновить и их? [Enter = Да / N = Нет]: ").strip()
        if ans.upper() != "N":
            _run_ps_shortcut_helper("patch_bambu", elevate=True)
            print("  ✅ Системные ярлыки Bambu обновлены!")


def ensure_shortcuts_silently():
    """Тихая проверка пользовательских ярлыков при старте (на случай обновления клиента Bambu)."""
    _run_ps_shortcut_helper("patch_bambu", elevate=False)


def remove_autostart_and_shortcuts():
    """Полностью убирает автозапуск и очищает флаг порта из всех ярлыков Bambu."""
    _run_ps_shortcut_helper("remove_startup", elevate=False)
    print("  ✅ Ярлык из Автозагрузки Windows удалён.")
    out = _run_ps_shortcut_helper("unpatch_bambu", elevate=False)
    if "ADMIN_NEEDED:" in out:
        print("🔐 Очистка системных ярлыков (Public Desktop) требует прав администратора...")
        _run_ps_shortcut_helper("unpatch_bambu", elevate=True)
    print("  ✅ Все ярлыки Bambu Farm Manager возвращены в исходное состояние.")
