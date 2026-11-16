import json
import subprocess
import time
import urllib.request
from pathlib import Path

import websocket
import config

JS_HOOK = r"""
(() => {
  window.__tasksMap = JSON.parse(localStorage.getItem('__bambu_tasks_v2') || '{}');
  window.__reprintsList = JSON.parse(localStorage.getItem('__bambu_reprints_v2') || '[]');
  window.__devicesInfo = JSON.parse(localStorage.getItem('__bambu_devices_v2') || '{}');
  window.__pState = window.__pState || {};

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

  const saveAll = () => {
    try {
      localStorage.setItem('__bambu_tasks_v2', JSON.stringify(window.__tasksMap));
      localStorage.setItem('__bambu_reprints_v2', JSON.stringify(window.__reprintsList));
      localStorage.setItem('__bambu_devices_v2', JSON.stringify(window.__devicesInfo));
    } catch (e) {}
    if (window.__updateBambuUI) window.__updateBambuUI();
  };

  const processDevicesArray = (devices) => {
    if (!Array.isArray(devices)) return;
    let changed = false;

    devices.forEach(d => {
      if (!d || typeof d !== 'object') return;
      const id = d.dev_id || d.sn;
      const rs = d.report_status || {};
      const state = rs.gcode_state || d.gcode_state;
      if (!id || !state) return;

      const devName = (d.name && d.name.trim()) || (d.dev_name && d.dev_name.trim()) || id;
      const pct = Number(rs.mc_percent ?? 0);
      const remMin = Number(rs.mc_remaining_time ?? 0);
      const subtaskId = String(rs.subtask_id || '');
      const taskId = String(rs.task_id || (d.running_task && d.running_task.id) || '');
      const fileName = rs.subtask_name || rs.gcode_file || (d.running_task && d.running_task.task_name) || '';

      if (d.running_task && d.running_task.id) {
        window.__tasksMap[String(d.running_task.id)] = compactTask(d.running_task);
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

      const prev = window.__pState[id] || {};
      if ((prev.state && prev.state !== 'RUNNING' && state === 'RUNNING') ||
          (state === 'RUNNING' && prev.pct >= 85 && pct <= 15)) {
        prev.start_time = new Date().toISOString();
      }

      const isNormalFinish = (prev.state === 'RUNNING' && state === 'FINISH');
      const isFastRestartFinish = (prev.state === 'RUNNING' && prev.pct >= 90 &&
                                   ((state === 'RUNNING' && pct <= 15) || state === 'IDLE'));

      if (isNormalFinish || isFastRestartFinish) {
        window.__reprintsList.push({
          dev_id: id,
          dev_name: devName,
          start_time: prev.start_time || null,
          finish_time: new Date().toISOString(),
          file: fileName || prev.file || '',
          task_id: taskId || prev.task_id || '',
          subtask_id: subtaskId || prev.subtask_id || ''
        });
        changed = true;
      }

      window.__pState[id] = {
        state: state,
        pct: pct,
        dev_name: devName,
        file: fileName || prev.file || '',
        task_id: taskId,
        subtask_id: subtaskId,
        start_time: prev.start_time
      };
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
            if (task && task.id) window.__tasksMap[String(task.id)] = compactTask(task);
          });
          saveAll();
        }
      }
    } catch (e) {}
  };

  if (!window.__bambuHookedV5) {
    window.__bambuHookedV5 = true;
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

  let box = document.getElementById('__bambu_logger_box');
  if (!box && document.body) {
    box = document.createElement('div');
    box.id = '__bambu_logger_box';
    box.style.cssText = 'position:fixed;bottom:14px;right:14px;z-index:999999;background:#1F3864;color:#fff;padding:8px 12px;border-radius:8px;font-family:Arial,sans-serif;font-size:12px;box-shadow:0 4px 12px rgba(0,0,0,0.35);display:flex;align-items:center;gap:10px;';
    const badge = document.createElement('span');
    window.__updateBambuUI = () => {
      badge.textContent = `🟢 Автологгер | 📦 Задач: ${Object.keys(window.__tasksMap).length} | 🔄 Финишей: ${window.__reprintsList.length} | 🖨️ Принтеров: ${Object.keys(window.__devicesInfo).length}`;
    };
    window.__updateBambuUI();

    const dlBtn = document.createElement('button');
    dlBtn.textContent = '💾 Скачать JSON';
    dlBtn.style.cssText = 'background:#28a745;color:#fff;border:none;padding:5px 10px;border-radius:5px;cursor:pointer;font-weight:bold;font-size:12px;';
    dlBtn.onclick = () => {
      const payload = [
        { total: Object.keys(window.__tasksMap).length, hits: Object.values(window.__tasksMap) },
        { __reprints_list__: window.__reprintsList },
        { __devices_info__: window.__devicesInfo }
      ];
      const blob = new Blob([JSON.stringify(payload)], { type: 'application/json' });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = 'history.json';
      a.click();
      URL.revokeObjectURL(a.href);
    };
    box.appendChild(badge);
    box.appendChild(dlBtn);
    document.body.appendChild(box);
  }
  return Object.keys(window.__tasksMap).length;
})();
"""


def is_farm_process_running() -> bool:
    """Проверяет через системный tasklist, реально ли запущен Bambu Farm Manager Client.exe."""
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
    """Принудительно закрывает зависшие фоновые процессы фермы."""
    subprocess.run(
        ["taskkill", "/F", "/IM", config.FARM_EXE_NAME, "/T"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    time.sleep(1.0)


def get_cdp_ws_url():
    """Возвращает WebSocket URL только если сам процесс фермы реально запущен."""
    if not is_farm_process_running():
        return None

    for port in config.CDP_SCAN_PORTS:
        try:
            with urllib.request.urlopen(f"http://127.0.0.1:{port}/json", timeout=1.0) as resp:
                targets = json.loads(resp.read().decode("utf-8"))
            for t in targets:
                t_type = t.get("type")
                t_url = t.get("url", "")
                if t_type == "page" and "devtools://" not in t_url and "webSocketDebuggerUrl" in t:
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


def pull_live_data_from_farm():
    if not get_cdp_ws_url():
        return None
    cdp_eval(JS_HOOK)
    expr = (
        "JSON.stringify(["
        "{total: Object.keys(window.__tasksMap||{}).length, hits: Object.values(window.__tasksMap||{})},"
        "{__reprints_list__: window.__reprintsList||[]},"
        "{__devices_info__: window.__devicesInfo||{}}"
        "])"
    )
    raw_json = cdp_eval(expr, timeout=10.0)
    if raw_json:
        try:
            data = json.loads(raw_json)
            if isinstance(data, list) and len(data) > 0:
                if len(data[0].get("hits", [])) > 0 or len(data[2].get("__devices_info__", {})) > 0:
                    config.LIVE_DUMP_FILE.write_text(raw_json, encoding="utf-8")
                    return config.LIVE_DUMP_FILE
        except Exception:
            pass
    return None


def start_farm_with_hook():
    proc_alive = is_farm_process_running()
    ws_url = get_cdp_ws_url()

    if proc_alive and ws_url:
        print("ℹ️ Процесс фермы найден в системе. Разворачиваем окно и обновляем хук...")
        cnt = cdp_eval(JS_HOOK, bring_to_front=True)
        print(f"✅ Автологгер активен! (В памяти задач: {cnt})")
        ans = input("Если окно фермы зависло в трее, перезапустить её с нуля? [R = Перезапустить / Enter = Оставить]: ").strip()
        if ans.upper() != "R":
            return
        kill_farm_processes()
    elif proc_alive and not ws_url:
        print("⚠️ Ферма запущена без порта автологгера (или зависла в фоне). Перезапускаем...")
        kill_farm_processes()

    exe_path = config.FARM_EXE
    if not exe_path.exists():
        raw = input("Укажи путь к Bambu Farm Manager Client.exe:\n> ").strip().strip('"')
        exe_path = Path(raw)
    if not exe_path.exists():
        print("❌ Файл клиента фермы не найден!")
        return

    print(f"🚀 Запускаем {exe_path.name} на порту {config.CDP_PORT}...")
    subprocess.Popen(
        [str(exe_path), f"--remote-debugging-port={config.CDP_PORT}", "--remote-allow-origins=*"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )

    for _ in range(25):
        time.sleep(1)
        if get_cdp_ws_url():
            time.sleep(1.5)
            cnt = cdp_eval(JS_HOOK, bring_to_front=True)
            if cnt is not None:
                print(f"✅ Ферма открыта, автологгер подключён! (В памяти задач: {cnt})")
                return
    print("❌ Не удалось дождаться ответа от окна фермы.")