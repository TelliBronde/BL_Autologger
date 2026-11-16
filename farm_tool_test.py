import base64
import json
import os
import re
import shutil
import subprocess
import sys
import time
import traceback
import urllib.error
import urllib.request
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path

import websocket
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

# ==============================================================================
# НАСТРОЙКИ ВЕРСИИ И РЕПОЗИТОРИЯ TelliBronde/BL_Autologger
# ==============================================================================
VERSION = "2.1.1-test"
GITHUB_REPO = "TelliBronde/BL_Autologger"
DEFAULT_STABLE_URL = f"https://raw.githubusercontent.com/{GITHUB_REPO}/main/farm_tool.py"
DEFAULT_TEST_URL = f"https://raw.githubusercontent.com/{GITHUB_REPO}/main/farm_tool_test.py"

if getattr(sys, "frozen", False):
    HERE = Path(sys.executable).resolve().parent
else:
    HERE = Path(__file__).resolve().parent

VERSIONS_DIR = HERE / "versions"
VERSIONS_DIR.mkdir(exist_ok=True)

CONFIG_FILE = HERE / "farm_config.json"
ODOMETER_FILE = HERE / "odometer_base.json"
MANUAL_FILE = HERE / "manual_reprints.json"
LIVE_DUMP_FILE = HERE / "history_live.json"

ACTIVE_SCRIPT = HERE / "farm_tool_active.py"
STABLE_SCRIPT = VERSIONS_DIR / "stable_backup.py"

SEARCH_DIRS = [HERE, Path.cwd(), Path.home() / "Downloads", Path.home() / "Desktop"]
FARM_EXE = Path(r"C:\Program Files\Bambu Farm Manager Client\Bambu Farm Manager Client.exe")
CDP_PORT = 9222

HF = "[h]:mm"
C_HEAD = "1F3864"
C_DAY  = "D9E1F2"
C_OK   = "C6EFCE"
C_REP  = "FFF2CC"
C_WORK = "DDEBF7"
C_WARN = "FCE4D6"
C_GREY = "F2F2F2"

TASK_RE = re.compile(
    r"^(?P<name>.*?)_(?:(?P<h>\d+)h(?:(?P<m>\d+)m)?|(?P<m2>\d+)m)_(?P<g>\d+(?:[.,]\d+)?)g$",
    re.I,
)
VER_RE = re.compile(r'^\s*VERSION\s*=\s*["\']([^"\']+)["\']', re.M)


# ==============================================================================
# 0. СИСТЕМА ВЕРСИОННОСТИ, ПРИВАТНОГО GITHUB И АВТОДЕПЛОЯ
# ==============================================================================
def extract_version_from_code(code_text):
    m = VER_RE.search(code_text or "")
    return m.group(1) if m else "unknown"


def get_script_version(path: Path):
    if not path.exists():
        return "—"
    try:
        return extract_version_from_code(path.read_text(encoding="utf-8-sig"))
    except Exception:
        return "ошибка чтения"


def load_config():
    default_cfg = {
        "channel": "stable",
        "github_stable_url": DEFAULT_STABLE_URL,
        "github_test_url": DEFAULT_TEST_URL,
        "github_token": "",
        "auto_update_on_start": False,
        "to_interval_hours": 300,
    }
    if CONFIG_FILE.exists():
        try:
            saved = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
            if not saved.get("github_stable_url"):
                saved["github_stable_url"] = DEFAULT_STABLE_URL
            if not saved.get("github_test_url"):
                saved["github_test_url"] = DEFAULT_TEST_URL
            default_cfg.update(saved)
        except Exception:
            pass
    else:
        CONFIG_FILE.write_text(json.dumps(default_cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    return default_cfg


def save_config(cfg):
    CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


def normalize_github_url(url):
    u = (url or "").strip()
    if "github.com/" in u and "/blob/" in u:
        u = u.replace("github.com/", "raw.githubusercontent.com/").replace("/blob/", "/")
    return u


def backup_current_version(tag="auto"):
    if not ACTIVE_SCRIPT.exists():
        return None
    try:
        code = ACTIVE_SCRIPT.read_text(encoding="utf-8-sig")
        ver = re.sub(r"[^a-zA-Z0-9._-]", "_", extract_version_from_code(code))
        ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        bak_path = VERSIONS_DIR / f"v_{ver}_{ts}_{tag}.py"
        bak_path.write_text(code, encoding="utf-8")

        archives = sorted(VERSIONS_DIR.glob("v_*.py"), key=lambda p: p.stat().st_mtime, reverse=True)
        for old_p in archives[20:]:
            try:
                old_p.unlink()
            except Exception:
                pass
        return bak_path
    except Exception:
        return None


def mark_current_as_stable():
    if ACTIVE_SCRIPT.exists():
        shutil.copy2(ACTIVE_SCRIPT, STABLE_SCRIPT)
        ver = get_script_version(STABLE_SCRIPT)
        print(f"📌 Версия v{ver} зафиксирована как эталонная СТАБИЛЬНАЯ сборка!")
    else:
        src = Path(__file__).resolve()
        if src.exists() and src.suffix == ".py":
            shutil.copy2(src, STABLE_SCRIPT)
            print(f"📌 Текущий файл v{VERSION} сохранён как эталонная СТАБИЛЬНАЯ сборка!")
        else:
            print(f"ℹ️ Сейчас работает встроенное в .exe ядро (v{VERSION}) — оно всегда доступно для отката.")


def rollback_to_stable():
    if STABLE_SCRIPT.exists():
        backup_current_version(tag="before_rollback")
        shutil.copy2(STABLE_SCRIPT, ACTIVE_SCRIPT)
        ver = get_script_version(ACTIVE_SCRIPT)
        print(f"⏪ Выполнен откат на сохранённую стабильную сборку: v{ver}!")
        return True
    elif ACTIVE_SCRIPT.exists():
        backup_current_version(tag="broken_removed")
        ACTIVE_SCRIPT.unlink()
        print(f"⏪ Внешний билд отключён. Возврат к заводскому ядру внутри .exe (v{VERSION})!")
        return True
    else:
        print("ℹ️ Вы уже находитесь на базовой заводской версии .exe.")
        return False


def update_from_github(silent=False):
    cfg = load_config()
    ch = cfg.get("channel", "stable")
    url_key = "github_test_url" if ch == "test" else "github_stable_url"
    url = normalize_github_url(cfg.get(url_key) or DEFAULT_STABLE_URL)
    token = (cfg.get("github_token") or "").strip()

    sep = "&" if "?" in url else "?"
    fetch_url = f"{url}{sep}_t={int(time.time())}"
    headers = {
        "User-Agent": "BambuFarmControl/2.1",
        "Cache-Control": "no-cache",
    }
    if token:
        headers["Authorization"] = f"token {token}"

    try:
        if not silent:
            print(f"🌐 Канал [{ch.upper()}]: скачиваем билд из {GITHUB_REPO}...\n   {url}")
        req = urllib.request.Request(fetch_url, headers=headers)
        with urllib.request.urlopen(req, timeout=8.0) as resp:
            code_text = resp.read().decode("utf-8-sig")

        if "def main_menu" not in code_text:
            if not silent:
                print("❌ ОТМЕНА: Скачанный файл не содержит функции main_menu.")
            return False

        try:
            compile(code_text, "<github_build>", "exec")
        except SyntaxError as se:
            if not silent:
                print(f"❌ ОТМЕНА ОБНОВЛЕНИЯ! В новом билде на GitHub синтаксическая ошибка:")
                print(f"   Строка {se.lineno}: {se.msg} -> {se.text}")
                print("   Текущая рабочая версия НЕ затронута.")
            return False

        backup_current_version(tag="pre_update")
        if ch == "stable" and not STABLE_SCRIPT.exists() and ACTIVE_SCRIPT.exists():
            shutil.copy2(ACTIVE_SCRIPT, STABLE_SCRIPT)

        ACTIVE_SCRIPT.write_text(code_text, encoding="utf-8")
        new_ver = extract_version_from_code(code_text)

        if ch == "stable":
            shutil.copy2(ACTIVE_SCRIPT, STABLE_SCRIPT)

        if not silent:
            print(f"✅ Успешно установлена версия v{new_ver} (канал {ch.upper()})!")
        return True
    except urllib.error.HTTPError as he:
        if not silent:
            if he.code == 404:
                print("❌ Ошибка 404: Файл ещё не загружен в репозиторий ИЛИ репозиторий Private (нужен GitHub Token либо сделать репозиторий Public).")
            else:
                print(f"❌ HTTP Ошибка {he.code}: {he.reason}")
        return False
    except Exception as e:
        if not silent:
            print(f"❌ Ошибка скачивания с GitHub: {e}")
        return False


def github_api_upload_file(repo, filename, content_str, token, commit_msg):
    """Загружает или обновляет файл напрямую в GitHub через REST API (даже без установленного Git!)."""
    api_url = f"https://api.github.com/repos/{repo}/contents/{filename}"
    headers = {
        "User-Agent": "BambuFarmControl/2.1",
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json",
        "Content-Type": "application/json",
    }
    sha = None
    try:
        req_get = urllib.request.Request(api_url, headers=headers, method="GET")
        with urllib.request.urlopen(req_get, timeout=8.0) as resp:
            existing = json.loads(resp.read().decode("utf-8"))
            sha = existing.get("sha")
    except Exception:
        pass

    payload = {
        "message": commit_msg,
        "content": base64.b64encode(content_str.encode("utf-8")).decode("ascii"),
        "branch": "main",
    }
    if sha:
        payload["sha"] = sha

    req_put = urllib.request.Request(
        api_url,
        data=json.dumps(payload).encode("utf-8"),
        headers=headers,
        method="PUT",
    )
    with urllib.request.urlopen(req_put, timeout=12.0) as resp:
        return resp.status in (200, 201)


def auto_turnkey_deploy():
    """Делает всё за пользователя: создаёт файлы, заливает в BL_Autologger и собирает Farm_Control.exe."""
    print("\n" + "=" * 78)
    print(f" ⚡ АВТО-НАСТРОЙКА ПОД КЛЮЧ ДЛЯ РЕПОЗИТОРИЯ {GITHUB_REPO}")
    print("=" * 78)

    src_file = HERE / "farm_tool.py"
    if not src_file.exists() and not getattr(sys, "frozen", False):
        src_file = Path(__file__).resolve()

    if not src_file.exists():
        print("❌ Запусти этот пункт из исходного файла python farm_tool.py!")
        return

    code_stable = src_file.read_text(encoding="utf-8-sig")
    code_test = VER_RE.sub('VERSION = "2.1.1-test"', code_stable, count=1)

    test_file = HERE / "farm_tool_test.py"
    test_file.write_text(code_test, encoding="utf-8")
    STABLE_SCRIPT.write_text(code_stable, encoding="utf-8")
    print("✅ 1/3. Созданы локальные файлы: farm_tool.py (v2.1.0) и farm_tool_test.py (v2.1.1-test).")

    cfg = load_config()
    cfg["github_stable_url"] = DEFAULT_STABLE_URL
    cfg["github_test_url"] = DEFAULT_TEST_URL
    save_config(cfg)

    print("\n--- Шаг 2/3. Загрузка файлов в твой GitHub (TelliBronde/BL_Autologger) ---")
    print("Так как репозиторий сейчас Private, выбери удобный способ:")
    print("  [1] Вставить GitHub Token (скрипт САМ зальёт файлы в приватный репо и настроит автообновление)")
    print("  [2] Загрузить через установленный Git (командами git init / git push)")
    print("  [3] Я просто перетащу файлы мышкой в открытое окно браузера GitHub (и сделаю репо Public)")
    way = input("Выбери [1 / 2 / 3]: ").strip() or "3"

    if way == "1":
        print("\nЧтобы создать токен за 15 секунд, открой: https://github.com/settings/tokens/new")
        print("(В поле Note напиши farm, поставь галочку [v] repo и нажми Generate token внизу)")
        tok = input("Вставь скопированный токен (ghp_...): ").strip()
        if tok:
            cfg["github_token"] = tok
            save_config(cfg)
            try:
                print("⏳ Загружаем farm_tool.py и farm_tool_test.py в TelliBronde/BL_Autologger...")
                github_api_upload_file(GITHUB_REPO, "farm_tool.py", code_stable, tok, "Init stable v2.1.0")
                github_api_upload_file(GITHUB_REPO, "farm_tool_test.py", code_test, tok, "Init test v2.1.1-test")
                print("✅ Оба файла автоматически загружены в твой приватный GitHub!")
            except Exception as e:
                print(f"⚠️ Не удалось загрузить через API: {e}")
    elif way == "2":
        try:
            subprocess.run(["git", "init"], cwd=str(HERE), check=False)
            subprocess.run(["git", "branch", "-M", "main"], cwd=str(HERE), check=False)
            subprocess.run(["git", "remote", "remove", "origin"], cwd=str(HERE), check=False, stderr=subprocess.DEVNULL)
            subprocess.run(["git", "remote", "add", "origin", f"https://github.com/{GITHUB_REPO}.git"], cwd=str(HERE), check=False)
            subprocess.run(["git", "add", "farm_tool.py", "farm_tool_test.py"], cwd=str(HERE), check=False)
            subprocess.run(["git", "commit", "-m", "Auto deploy v2.1.0"], cwd=str(HERE), check=False)
            subprocess.run(["git", "push", "-u", "origin", "main", "--force"], cwd=str(HERE), check=True)
            print("✅ Файлы успешно отправлены через Git!")
        except Exception as e:
            print(f"⚠️ Git не настроен или не установлен ({e}). Открываю папку для перетаскивания мышкой...")
            way = "3"

    if way == "3":
        print("\n📂 Сейчас откроется папка с файлами и страница твоего репозитория.")
        print("Просто перетащи мышкой farm_tool.py и farm_tool_test.py в окно браузера и нажми Commit changes!")
        try:
            os.startfile(str(HERE))
            os.startfile(f"https://github.com/{GITHUB_REPO}/upload/main")
        except Exception:
            pass

    print("\n--- Шаг 3/3. Автоматическая сборка Farm_Control.exe ---")
    ans_b = input("Собрать готовый Farm_Control.exe прямо сейчас? [Enter = Да / N = Нет]: ").strip()
    if ans_b.upper() != "N":
        print("⏳ Собираем Farm_Control.exe через PyInstaller (подожди ~30 секунд)...")
        res = subprocess.run(
            [sys.executable, "-m", "PyInstaller", "--onefile", "--name", "Farm_Control", str(src_file)],
            cwd=str(HERE),
        )
        built_exe = HERE / "dist" / "Farm_Control.exe"
        target_exe = HERE / "Farm_Control.exe"
        if res.returncode == 0 and built_exe.exists():
            try:
                shutil.copy2(built_exe, target_exe)
                print(f"\n🎉 ГОТОВО! Файл скопирован прямо сюда: {target_exe}")
            except PermissionError:
                print(f"\n🎉 ГОТОВО! Забери готовый файл здесь: {built_exe}")
        else:
            print("⚠️ Ошибка при сборке PyInstaller.")


def execute_active_script_safely():
    if not ACTIVE_SCRIPT.exists():
        return
    try:
        code = ACTIVE_SCRIPT.read_text(encoding="utf-8-sig")
        g = globals().copy()
        g["__HOT_RELOADED__"] = True
        g["__ACTIVE_VER__"] = extract_version_from_code(code)
        g["__file__"] = str(ACTIVE_SCRIPT)
        exec(compile(code, str(ACTIVE_SCRIPT), "exec"), g)
        sys.exit(0)
    except SystemExit:
        raise
    except KeyboardInterrupt:
        sys.exit(0)
    except Exception as e:
        bad_ver = get_script_version(ACTIVE_SCRIPT)
        print("\n" + "!" * 78)
        print(f" 🚨 ВНИМАНИЕ: СБОЙ В БИЛДЕ v{bad_ver} (СРАБОТАЛА ЗАЩИТА WATCHDOG)")
        print("!" * 78)
        print(f"Ошибка: {type(e).__name__}: {e}\n")
        traceback.print_exc(limit=4)
        print("\n" + "-" * 78)
        print("  [1] ⏪ Мгновенно откатиться на СТАБИЛЬНУЮ версию (Рекомендуется)")
        print("  [2] 🗂️ Выбрать другую версию из архива бэкапов")
        print(f"  [3] 🛡️ Запустить встроенное заводское ядро .exe (v{VERSION})")
        print("-" * 78)
        ans = input("> ").strip() or "1"

        backup_current_version(tag="BROKEN")
        if ans == "2":
            choose_version_from_archive()
        elif ans == "3":
            if ACTIVE_SCRIPT.exists():
                ACTIVE_SCRIPT.unlink()
        else:
            rollback_to_stable()

        input("\nНажми Enter для запуска восстановленной версии...")


def choose_version_from_archive():
    archives = sorted(VERSIONS_DIR.glob("v_*.py"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not archives:
        print("Архив версий пока пуст.")
        return False
    print("\n--- 🗂️ АРХИВ СОХРАНЁННЫХ ВЕРСИЙ ---")
    for idx, p in enumerate(archives, 1):
        ver = get_script_version(p)
        dt_str = datetime.fromtimestamp(p.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
        print(f"  [{idx:2d}] v{ver:<12} | Сохранена: {dt_str} | Файл: {p.name}")
    print("-----------------------------------")
    num = input("Введи номер версии для восстановления (или Enter для отмены): ").strip()
    if num.isdigit() and 1 <= int(num) <= len(archives):
        chosen = archives[int(num) - 1]
        shutil.copy2(chosen, ACTIVE_SCRIPT)
        print(f"✅ Восстановлена версия v{get_script_version(ACTIVE_SCRIPT)} из файла {chosen.name}!")
        return True
    return False


def version_control_menu():
    cfg = load_config()
    while True:
        active_ver = get_script_version(ACTIVE_SCRIPT) if ACTIVE_SCRIPT.exists() else f"{VERSION} (ядро .exe)"
        stable_ver = get_script_version(STABLE_SCRIPT) if STABLE_SCRIPT.exists() else f"{VERSION} (ядро .exe)"
        ch = cfg.get("channel", "stable").upper()
        archives_cnt = len(list(VERSIONS_DIR.glob("v_*.py")))
        has_token = "Задан (Private OK)" if cfg.get("github_token") else "Не задан (только для Public репо)"

        print("\n" + "=" * 78)
        print(f" 🌐 УПРАВЛЕНИЕ ВЕРСИЯМИ И GITHUB ({GITHUB_REPO})")
        print("=" * 78)
        print(f" • Текущая рабочая версия:   v{active_ver}")
        print(f" • Эталонная STABLE версия:  v{stable_ver}")
        print(f" • Активный канал:           [{ch}]  | Архивных бэкапов: {archives_cnt} шт.")
        print(f" • Доступ к GitHub:          {has_token}")
        print(f" • URL STABLE: {cfg.get('github_stable_url')}")
        print(f" • URL TEST:   {cfg.get('github_test_url')}")
        print("-" * 78)
        print(f"  [1] 🔄 Скачать и применить обновление с GitHub (канал {ch})")
        print("  [2] ⏪ ОТКАТИТЬСЯ НА СТАБИЛЬНУЮ ВЕРСИЮ В 1 КЛИК (Rollback)")
        print("  [3] 📌 Сделать текущую версию эталонной СТАБИЛЬНОЙ")
        print("  [4] 🔀 Переключить канал: STABLE <-> TEST (песочница)")
        print(f"  [5] 🗂️ Архив версий (выбрать любую из {archives_cnt} прошлых сборок)")
        print("  [6] 🔑 Указать GitHub Token (для приватного репо) / Настроить ссылки")
        print(f"  [7] 🛡️ Сбросить до заводского ядра внутри .exe (v{VERSION})")
        print("  [Enter] Назад в главное меню")
        ans = input("> ").strip()

        if ans == "1":
            if update_from_github(silent=False):
                input("\nНажми Enter для перезагрузки в новую версию...")
                execute_active_script_safely()
        elif ans == "2":
            if rollback_to_stable():
                input("\nНажми Enter для перезагрузки в стабильную версию...")
                if ACTIVE_SCRIPT.exists():
                    execute_active_script_safely()
                else:
                    return
        elif ans == "3":
            mark_current_as_stable()
        elif ans == "4":
            cfg["channel"] = "test" if cfg.get("channel") == "stable" else "stable"
            save_config(cfg)
            print(f"🔀 Активный канал переключён на: [{cfg['channel'].upper()}]")
        elif ans == "5":
            if choose_version_from_archive():
                input("\nНажми Enter для запуска выбранной версии...")
                execute_active_script_safely()
        elif ans == "6":
            print("\n1 — Ввести GitHub Personal Access Token (для приватного репозитория)")
            print("2 — Вкл/Выкл автообновление при запуске программы")
            print("3 — Изменить URL ссылок")
            sub = input("> ").strip()
            if sub == "1":
                tok = input("Вставь токен (ghp_... или github_pat_...): ").strip()
                cfg["github_token"] = tok
                save_config(cfg)
                print("✅ Токен сохранён!")
            elif sub == "2":
                cfg["auto_update_on_start"] = not cfg.get("auto_update_on_start", False)
                save_config(cfg)
            elif sub == "3":
                u1 = input(f"URL STABLE (Enter = {DEFAULT_STABLE_URL}): ").strip() or DEFAULT_STABLE_URL
                u2 = input(f"URL TEST (Enter = {DEFAULT_TEST_URL}): ").strip() or DEFAULT_TEST_URL
                cfg["github_stable_url"] = normalize_github_url(u1)
                cfg["github_test_url"] = normalize_github_url(u2)
                save_config(cfg)
        elif ans == "7":
            if ACTIVE_SCRIPT.exists():
                backup_current_version(tag="before_factory_reset")
                ACTIVE_SCRIPT.unlink()
                print(f"🛡️ Внешние скрипты отключены. Работает заводское ядро v{VERSION}.")
                return
        else:
            break


# ==============================================================================
# 1. JS-РАДАР И СВЯЗЬ С ФЕРМОЙ ЧЕРЕЗ ПОРТ 9222
# ==============================================================================
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
        const tid = String(d.running_task.id);
        window.__tasksMap[tid] = compactTask(d.running_task);
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
        const ev = {
          dev_id: id,
          dev_name: devName,
          start_time: prev.start_time || null,
          finish_time: new Date().toISOString(),
          file: fileName || prev.file || '',
          task_id: taskId || prev.task_id || '',
          subtask_id: subtaskId || prev.subtask_id || ''
        };
        window.__reprintsList.push(ev);
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


def get_cdp_ws_url():
    for port in (CDP_PORT, 9333):
        try:
            with urllib.request.urlopen(f"http://127.0.0.1:{port}/json", timeout=1.2) as resp:
                targets = json.loads(resp.read().decode("utf-8"))
            for t in targets:
                if t.get("type") == "page" and "webSocketDebuggerUrl" in t:
                    return t["webSocketDebuggerUrl"]
        except Exception:
            pass
    return None


def cdp_eval(expression, timeout=5.0):
    ws_url = get_cdp_ws_url()
    if not ws_url:
        return None
    try:
        ws = websocket.create_connection(ws_url, timeout=timeout, suppress_origin=True)
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
                    LIVE_DUMP_FILE.write_text(raw_json, encoding="utf-8")
                    return LIVE_DUMP_FILE
        except Exception:
            pass
    return None


def start_farm_with_hook():
    if not get_cdp_ws_url():
        exe_path = FARM_EXE
        if not exe_path.exists():
            raw = input("Укажи путь к Bambu Farm Manager Client.exe:\n> ").strip().strip('"')
            exe_path = Path(raw)
        if not exe_path.exists():
            print("❌ Файл клиента фермы не найден!")
            return
        print("Запускаем Bambu Farm Manager Client с портом автологгера...")
        subprocess.Popen(
            [str(exe_path), f"--remote-debugging-port={CDP_PORT}", "--remote-allow-origins=*"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    else:
        print("Клиент уже запущен, подключаемся...")

    for _ in range(25):
        time.sleep(1)
        if get_cdp_ws_url():
            time.sleep(1.5)
            cnt = cdp_eval(JS_HOOK)
            if cnt is not None:
                print(f"✅ Автологгер успешно подключён! (В памяти задач: {cnt})")
                return
    print("❌ Не удалось подключиться. Если ферма была открыта через старый ярлык — закрой её и нажми [1] ещё раз.")


# ==============================================================================
# 2. БАЗА ОДОМЕТРОВ И РУЧНЫХ ПОВТОРОВ
# ==============================================================================
def load_odometers():
    if ODOMETER_FILE.exists():
        try:
            return json.loads(ODOMETER_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def save_odometers(data):
    ODOMETER_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def load_manual_reprints():
    if not MANUAL_FILE.exists():
        default_list = [
            {"date": "2026-11-12", "printer": "109 AMS", "file": "Poland_test_numero_uno", "qty": 4}
        ]
        MANUAL_FILE.write_text(json.dumps(default_list, ensure_ascii=False, indent=2), encoding="utf-8")
        return default_list
    try:
        return json.loads(MANUAL_FILE.read_text(encoding="utf-8"))
    except Exception:
        return []


# ==============================================================================
# 3. ПАРСЕР ИСТОРИИ И РАСЧЁТ ДЕЛЬТЫ ПОСЛЕ СВЕРКИ ОДОМЕТРА
# ==============================================================================
def strip_ext(name):
    n = (name or "").strip()
    for ext in (".gcode.3mf", ".3mf", ".gcode"):
        if n.lower().endswith(ext):
            n = n[: -len(ext)]
    return n


def parse_task_name(raw):
    n = strip_ext(raw)
    m = TASK_RE.match(n)
    if not m:
        return n, None, None
    secs = int(m.group("h") or 0) * 3600 + int(m.group("m") or m.group("m2") or 0) * 60
    grams = float(m.group("g").replace(",", "."))
    return m.group("name"), secs, grams


def fmt_h(sec):
    sec = int(sec or 0)
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
            fil = sum(float(x.get("used_g", 0) or 0) for x in plates[0].get("filaments", []))
            if fil:
                weight = fil
            elif meta.get("weight"):
                weight = float(meta["weight"])
            if meta.get("prediction"):
                secs = int(float(meta["prediction"]))
    except Exception:
        pass

    clean_name, name_secs, name_weight = parse_task_name(raw_name)
    if weight is None:
        weight = name_weight or 0.0
    if secs is None:
        secs = name_secs or 0
    return clean_name, secs, weight


def find_all_json_files():
    live = pull_live_data_from_farm()
    if live:
        print("⚡ Живые данные автоматически получены из открытого окна Farm Manager!")
    found = []
    for d in SEARCH_DIRS:
        if not d.exists():
            continue
        for pat in ("history*", "timelogger*"):
            for p in d.glob(pat):
                if p.is_file() and p.suffix.lower() in (".json", ".txt", "") and p.stat().st_size > 100:
                    found.append(p.resolve())
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


def Number_or_zero(val):
    try:
        return float(val or 0)
    except Exception:
        return 0.0


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
                    dev_name = (d.get("name") or "").strip() or (d.get("dev_name") or "").strip() or d.get("dev_id")
                    noz_t = rs.get("nozzle_type") or ""
                    noz_str = f"{rs.get('nozzle_diameter', '0.4')} ({'Hardened' if noz_t == 'hardened_steel' else 'Stainless'})" if noz_t else ""
                    yield {
                        "dev_id": d.get("dev_id", ""),
                        "name": dev_name,
                        "model": "Bambu A1" if d.get("dev_model") == "N2S" else ("Bambu P1S" if d.get("dev_model") == "C12" else d.get("dev_model", "")),
                        "ip": d.get("dev_ip", ""),
                        "fw": (d.get("upgrade_info") or {}).get("current_ver", ""),
                        "nozzle": noz_str,
                        "wifi": rs.get("wifi_signal", ""),
                        "hms_errors": len(rs.get("hms") or []),
                        "state": rs.get("gcode_state", ""),
                        "mc_percent": Number_or_zero(rs.get("mc_percent")),
                        "mc_remaining_time": Number_or_zero(rs.get("mc_remaining_time")),
                        "current_file": rs.get("subtask_name") or rs.get("gcode_file") or "",
                    }
        for v in o.values():
            yield from walk_devices_info(v)
    elif isinstance(o, list):
        for v in o:
            yield from walk_devices_info(v)


def load_and_parse(quiet=False):
    json_files = find_all_json_files()
    if not json_files:
        if not quiet:
            print("❌ Не найден ни один файл history*.json.")
        return None, None, None

    best = {}
    all_reprints_raw = []
    devices_hw = {}

    for jp in sorted(json_files, key=lambda p: p.stat().st_mtime):
        try:
            with open(jp, "r", encoding="utf-8-sig") as f:
                raw_data = json.load(f)
            for t in walk_tasks(raw_data):
                st = t.get("statistics") if isinstance(t.get("statistics"), dict) else {}
                score = (int(st.get("finished") or 0) + int(st.get("stopped") or 0), str(t.get("finish_time") or ""))
                key = str(t.get("id") or id(t))
                if key not in best or score >= best[key][0]:
                    best[key] = (score, t)
            all_reprints_raw.extend(list(walk_reprints(raw_data)))
            for d_info in walk_devices_info(raw_data):
                d_name = d_info.get("name")
                if d_name:
                    devices_hw[d_name] = d_info
        except Exception:
            pass

    tasks_list = [v[1] for v in best.values()]
    if not tasks_list and not devices_hw:
        return None, None, None

    odometers = load_odometers()
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
        qty = int(stats.get("finished") or 0) if stats else 1
        stopped = int(stats.get("stopped") or 0) if stats else 0
        printing_now = int(stats.get("printing") or 0) if stats else 0
        total_p = int(stats.get("total") or qty) if stats else qty

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
        dev_names = [(d.get("dev_name") or d.get("name") or d.get("dev_id") or "").strip() for d in pool if isinstance(d, dict)]
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
            (d.get("dev_name") or d.get("name") or d.get("dev_id") or "").strip()
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
        dn = (ev.get("dev_name") or ev.get("dev_id") or "Unknown").strip()
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

    for item in load_manual_reprints():
        m_date, m_printer, m_file, m_qty = item["date"], item["printer"], item["file"], int(item["qty"])
        if m_qty <= 0:
            continue
        raw_m = strip_ext(m_file)
        if raw_m.lower() in metrics_by_file:
            raw_name, clean_name, secs_one, weight_one = metrics_by_file[raw_m.lower()]
        else:
            raw_name = raw_m
            clean_name, secs_one, weight_one = parse_task_name(raw_name)
            secs_one, weight_one = secs_one or 0, weight_one or 0.0

        m_dt = datetime.strptime(f"{m_date} 20:00", "%Y-%m-%d %H:%M")
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


# ==============================================================================
# 4. ИНТЕРАКТИВНАЯ СВЕРКА МОТОЧАСОВ И РУЧНОЙ ОДОМЕТР
# ==============================================================================
def odometer_reconciliation_menu():
    _, printers_stat, _ = load_and_parse(quiet=True)
    if printers_stat is None:
        printers_stat = {}
    odometers = load_odometers()
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
        print("Команды:")
        print("  • Введи НОМЕР принтера (например, 12) — внести моточасы с экрана и выполнить сверку")
        print("  • [A] Поочерёдный обход всех принтеров   [T] Отметить проведение ТО на принтере")
        print("  • [Enter] Вернуться в главное меню")
        ch = input("> ").strip()

        if not ch:
            break
        elif ch.upper() == "A":
            print("\nОбход принтеров (жми Enter, чтобы пропустить принтер, или Q для выхода):")
            for pname in all_printers:
                val = input(f"  Одометр с экрана [{pname}]: ").strip()
                if val.upper() == "Q":
                    break
                if val:
                    update_single_printer_odometer(pname, val, printers_stat, odometers)
            save_odometers(odometers)
            _, printers_stat, _ = load_and_parse(quiet=True)
        elif ch.upper() == "T":
            num = input("Номер принтера, на котором провели ТО: ").strip()
            if num.isdigit() and 1 <= int(num) <= len(all_printers):
                pname = all_printers[int(num) - 1]
                ps = printers_stat.get(pname, {"secs_good": 0, "secs_waste": 0, "secs_after_sync": 0})
                odo = odometers.get(pname, {})
                curr_h = (float(odo["base_hours"]) + ps["secs_after_sync"] / 3600.0) if "base_hours" in odo else ((ps["secs_good"] + ps["secs_waste"]) / 3600.0)
                odo["last_to_hours"] = round(curr_h, 1)
                odo["last_to_date"] = datetime.now().strftime("%Y-%m-%d")
                odometers[pname] = odo
                save_odometers(odometers)
                print(f"✅ Для [{pname}] зафиксировано ТО на отметке {curr_h:.1f} ч!")
        elif ch.isdigit() and 1 <= int(ch) <= len(all_printers):
            pname = all_printers[int(ch) - 1]
            val = input(f"Введи текущие моточасы с экрана [{pname}] (ч): ").strip()
            if val:
                update_single_printer_odometer(pname, val, printers_stat, odometers)
                save_odometers(odometers)
                _, printers_stat, _ = load_and_parse(quiet=True)


def update_single_printer_odometer(pname, val_str, printers_stat, odometers):
    try:
        new_h = float(val_str.replace(",", "."))
    except ValueError:
        print("❌ Некорректное число.")
        return

    now_dt = datetime.now()
    now_str = now_dt.strftime("%Y-%m-%d %H:%M")
    ps = printers_stat.get(pname, {"secs_good": 0, "secs_waste": 0, "secs_after_sync": 0, "events": []})
    old_rec = odometers.get(pname, {})

    drift_h = None
    if "base_hours" in old_rec and old_rec.get("sync_time"):
        prev_h = float(old_rec["base_hours"])
        prev_dt = parse_iso_local(old_rec["sync_time"])
        screen_delta = new_h - prev_h
        log_delta = sum(s for dt, s in ps.get("events", []) if prev_dt and dt > prev_dt) / 3600.0
        drift_h = round(screen_delta - log_delta, 1)
        print(f"\n🔍 РЕЗУЛЬТАТ СВЕРКИ [{pname}] (с {old_rec['sync_time']} по {now_str}):")
        print(f"   • Пробег по экрану принтера: +{screen_delta:.1f} ч (было {prev_h:.1f} -> стало {new_h:.1f})")
        print(f"   • Пробег по логам фермы:     +{log_delta:.1f} ч")
        print(f"   • Расхождение (Экран - Лог): {drift_h:+.1f} ч " + (
            "(принтер печатал повторы мимо логов!)" if drift_h > 1.5 else "(сходится отлично!)"
        ))
    else:
        log_total = (ps["secs_good"] + ps["secs_waste"]) / 3600.0
        drift_h = round(new_h - log_total, 1)
        print(f"\n✅ Первичная привязка одометра [{pname}]: экран = {new_h:.1f} ч (в логах с сентября было {log_total:.1f} ч).")

    odometers[pname] = {
        "base_hours": new_h,
        "sync_time": now_str,
        "last_to_hours": old_rec.get("last_to_hours", 0.0),
        "last_to_date": old_rec.get("last_to_date", ""),
        "last_drift_hours": drift_h,
    }


# ==============================================================================
# 5. ТЕСТОВЫЙ СТЕНД И ПРОВЕРКА ЗАЩИТЫ ОТ СБОЕВ
# ==============================================================================
def test_bench_menu():
    while True:
        print("\n=== 🧪 ТЕСТОВЫЙ РЕЖИМ И ДИАГНОСТИКА ===")
        print("  [1] 📡 Проверить живой ответ от фермы (показать, что прямо сейчас печатают принтеры)")
        print("  [2] 🔄 Добавить/удалить ручные повторы за прошлые дни")
        print("  [3] 💥 Симулировать сбой тестового билда (проверить автооткат Watchdog)")
        print("  [Enter] Назад")
        ch = input("> ").strip()

        if ch == "1":
            ws_url = get_cdp_ws_url()
            print(f"Порт отладки фермы: {ws_url or 'НЕ НАЙДЕН'}")
            res = load_and_parse(quiet=True)
            if res and res[2]:
                devices_hw = res[2]
                print(f"\nНайдено принтеров в телеметрии /devices2: {len(devices_hw)}")
                for name, hw in sorted(devices_hw.items()):
                    print(f"  • {name:<10} | {hw.get('model'):<9} | Статус: {hw.get('state'):<8} | "
                          f"{hw.get('mc_percent', 0):3.0f}% (ост. {hw.get('mc_remaining_time', 0):.0f} мин) | Файл: {hw.get('current_file')}")
            else:
                print("Нет данных /devices2.")
        elif ch == "2":
            items = load_manual_reprints()
            for idx, it in enumerate(items, 1):
                print(f"  {idx}. {it['date']} | Принтер: {it['printer']} | Файл: {it['file']} | Повторов: {it['qty']} шт.")
            sub = input("[1] Добавить запись  [2] Удалить запись  [Enter] Назад: ").strip()
            if sub == "1":
                d_def = datetime.now().strftime("%Y-%m-%d")
                d_in = input(f"Дата (Enter = {d_def}): ").strip() or d_def
                pr_in = input("Имя принтера (например, 109 AMS): ").strip()
                f_in = input("Имя файла (например, Poland_test_numero_uno): ").strip()
                q_in = input("Количество повторов (шт): ").strip()
                if pr_in and f_in and q_in.isdigit():
                    items.append({"date": d_in, "printer": pr_in, "file": f_in, "qty": int(q_in)})
                    MANUAL_FILE.write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")
                    print("✅ Запись сохранена!")
            elif sub == "2":
                num = input("Номер для удаления: ").strip()
                if num.isdigit() and 1 <= int(num) <= len(items):
                    items.pop(int(num) - 1)
                    MANUAL_FILE.write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")
        elif ch == "3":
            print("\nСоздаём искусственно сломанный билд v9.9.9-broken и пробуем его запустить...")
            backup_current_version(tag="before_crash_test")
            ACTIVE_SCRIPT.write_text('VERSION = "9.9.9-broken"\ndef main_menu():\n    raise RuntimeError("Тестовая ошибка в новом билде!")\nmain_menu()\n', encoding="utf-8")
            execute_active_script_safely()
        else:
            break


# ==============================================================================
# 6. ГЕНЕРАЦИЯ EXCEL-ОТЧЁТА
# ==============================================================================
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


def generate_report():
    res = load_and_parse()
    if not res or not res[0]:
        return
    parsed, printers_stat, devices_hw = res
    odometers = load_odometers()
    cfg = load_config()
    to_interval = float(cfg.get("to_interval_hours") or 300)

    by_day = defaultdict(list)
    for x in parsed:
        by_day[x["date"]].append(x)

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

    put(sh_df, "A1", "Разбивка выполненных задач по дням и файлам", bold=True, size=15)
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
    put(sh_df, f"B{rr}", f"За все дни ({len(sorted_days)} дн.)", bold=True, fill=C_GREY)
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

    put(sh_dyn, "A1", "Сводная статистика по дням", bold=True, size=15)
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

    put(sh_all, "A1", f"Журнал задач и повторов ({len(parsed)} шт.)", bold=True, size=15)
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

    put(sh_pr, "A1", "Сверка моточасов (Одометр экрана + Прирост), выработка и ТО", bold=True, size=15)
    put(sh_pr, "A2", "Если задан ручной одометр с экрана, итоговый пробег считается как: База экрана + Прирост после даты сверки.", color="7F7F7F")
    header(sh_pr, 4, [
        "Принтер", "Модель", "Сопло", "IP / Wi-Fi",
        "Готовых, шт", "Пластик, кг",
        "По логам фермы, ч", "База с экрана, ч", "Дата сверки",
        "Прирост после сверки, ч", "АКТУАЛЬНО ОДОМЕТР, ч",
        "Сверка (Экран−Лог), ч", "Статус ТО / HMS"
    ])

    for i, (pname, pstat) in enumerate(sorted(printers_stat.items())):
        r_idx = 5 + i
        hw = devices_hw.get(pname, {})
        odo = odometers.get(pname, {})

        log_total_h = (pstat["secs_good"] + pstat["secs_waste"]) / 3600.0
        has_odo = "base_hours" in odo
        base_h = float(odo["base_hours"]) if has_odo else None
        sync_str = odo.get("sync_time", "—") if has_odo else "Не сверялся"
        delta_h = (pstat["secs_after_sync"] / 3600.0) if has_odo else 0.0
        actual_odo_h = (base_h + delta_h) if has_odo else log_total_h
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
        put(sh_pr, f"G{r_idx}", round(log_total_h, 1), fmt="0.0")
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
    widths(sh_pr, [15, 12, 18, 22, 14, 13, 17, 16, 17, 19, 21, 18, 22])

    for sh in wb.worksheets:
        sh.page_setup.orientation = "landscape"
        sh.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        sh.page_setup.fitToWidth = 1
        sh.page_setup.fitToHeight = 0

    wb.save(export_path)
    print(f"\n✅ Excel-отчёт со сверкой моточасов готов: {export_path}")
    if sys.platform == "win32":
        os.startfile(export_path)


def main_menu():
    cfg = load_config()
    if cfg.get("auto_update_on_start") and not globals().get("__HOT_RELOADED__"):
        update_from_github(silent=True)

    if not globals().get("__HOT_RELOADED__") and ACTIVE_SCRIPT.exists():
        execute_active_script_safely()

    while True:
        os.system("cls" if os.name == "nt" else "clear")
        cfg = load_config()
        cdp_status = "🟢 ПОДКЛЮЧЕНА" if get_cdp_ws_url() else "⚪ Не запущена (или без порта)"
        ch_str = cfg.get("channel", "stable").upper()
        if globals().get("__HOT_RELOADED__"):
            ver_badge = f"v{globals().get('__ACTIVE_VER__', VERSION)} [🔥 Внешний билд | Канал: {ch_str}]"
        else:
            ver_badge = f"v{VERSION} [🛡️ Заводское ядро .exe | Канал: {ch_str}]"

        print("==========================================================================")
        print(f"   BAMBU FARM CONTROL — {ver_badge}")
        print(f"   Репозиторий: github.com/{GITHUB_REPO}")
        print("==========================================================================")
        print(f" Связь с фермой: {cdp_status} | Одометров в базе: {len(load_odometers())} шт.")
        print("--------------------------------------------------------------------------")
        print("  [1] 🚀 Запустить Ферму + Автологгер (захват задач и /devices2)")
        print("  [2] 📊 Сформировать и открыть Excel-отчёт (выработка + Одометр)")
        print("  [3] 🔧 Сверка и ручной ввод моточасов с экрана принтера (Одометр + ТО)")
        print("  [4] 🧪 Тестовый стенд (диагностика + тест защиты от сбоев)")
        print("  [5] 🌐 Версии, Бэкапы и GitHub (Обновление / Откат на стабильную в 1 клик)")
        if not getattr(sys, "frozen", False):
            print("  [6] ⚡ АВТО-НАСТРОЙКА ПОД КЛЮЧ (Залить в BL_Autologger + Собрать .exe)")
        print("  [0] Выход")
        print("==========================================================================")
        choice = input("Выбери пункт меню: ").strip()

        if choice == "1":
            start_farm_with_hook()
            input("\nНажми Enter, чтобы вернуться в меню...")
        elif choice == "2":
            generate_report()
            input("\nНажми Enter, чтобы вернуться в меню...")
        elif choice == "3":
            odometer_reconciliation_menu()
        elif choice == "4":
            test_bench_menu()
        elif choice == "5":
            version_control_menu()
        elif choice == "6" and not getattr(sys, "frozen", False):
            auto_turnkey_deploy()
            input("\nНажми Enter, чтобы вернуться в меню...")
        elif choice == "0":
            break


if __name__ == "__main__":
    main_menu()