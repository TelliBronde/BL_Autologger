import getpass
import importlib
import importlib.util
import os
import shutil
import sys
import time
import traceback
import urllib.request
from datetime import datetime

import config
import connector
import parser
import report


def reload_disk_modules():
    """Подгружает внешние .py модули с диска только если это разрешено в конфиге."""
    cfg = config.load_config()
    if getattr(sys, "frozen", False) and not cfg.get("allow_disk_override", True):
        return
    for mod_file in ("config.py", "connector.py", "parser.py", "report.py"):
        mod_name = mod_file[:-3]
        disk_path = config.HERE / mod_file
        if disk_path.exists():
            try:
                spec = importlib.util.spec_from_file_location(mod_name, str(disk_path))
                mod = importlib.util.module_from_spec(spec)
                sys.modules[mod_name] = mod
                spec.loader.exec_module(mod)
                globals()[mod_name] = mod
            except Exception as e:
                print(f"⚠️ Ошибка в модуле {mod_file}: {e}")


def restart_self():
    print("🔄 Перезапуск меню с обновлённым кодом...")
    time.sleep(0.8)
    if getattr(sys, "frozen", False):
        os.execv(sys.executable, [sys.executable] + sys.argv[1:])
    else:
        os.execv(sys.executable, [sys.executable, str(config.HERE / "main.py")])


def sync_modules_from_github(target_branch=None):
    cfg = config.load_config()
    if target_branch in ("main", "stable"):
        branch = "main"
        cfg["channel"] = "stable"
    elif target_branch == "test":
        branch = "test"
        cfg["channel"] = "test"
    else:
        branch = "test" if cfg.get("channel") == "test" else "main"

    cfg["allow_disk_override"] = True
    config.save_config(cfg)

    repo = cfg.get("github_repo", config.GITHUB_REPO)
    token = (cfg.get("github_token") or "").strip()
    headers = {"User-Agent": "BambuFarmControl/2.5", "Cache-Control": "no-cache"}
    if token:
        headers["Authorization"] = f"token {token}"

    print(f"\n🌐 Скачиваем модули из ветки [{branch.upper()}] ({repo})...")
    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    bak_dir = config.VERSIONS_DIR / f"backup_{ts}_{branch}"
    bak_dir.mkdir(exist_ok=True)
    for m in config.MODULES:
        p = config.HERE / m
        if p.exists():
            shutil.copy2(p, bak_dir / m)

    updated = 0
    for m in config.MODULES:
        url = f"https://raw.githubusercontent.com/{repo}/{branch}/{m}?_t={int(time.time())}"
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=6.0) as resp:
                code_text = resp.read().decode("utf-8-sig")
            compile(code_text, m, "exec")
            (config.HERE / m).write_text(code_text, encoding="utf-8")
            updated += 1
            print(f"  ✅ Обновлён модуль: {m} (ветка {branch})")
        except Exception as e:
            print(f"  ⚠️ Пропущен {m}: {e}")

    if updated:
        print(f"\n🎉 Успешно обновлено модулей: {updated} из ветки [{branch.upper()}]!")
        print(f"   Бэкап прошлой версии: versions/{bak_dir.name}")
        restart_self()
    else:
        print("\n❌ Ни один модуль не обновлён. Проверь интернет или GitHub Token.")


def rollback_modules():
    backups = sorted([d for d in config.VERSIONS_DIR.iterdir() if d.is_dir()], key=lambda p: p.stat().st_mtime, reverse=True)
    if not backups:
        print("Архив бэкапов пока пуст.")
        return
    print("\n--- 🗂️ ДОСТУПНЫЕ БЭКАПЫ ДЛЯ ОТКАТА ---")
    for idx, b in enumerate(backups[:10], 1):
        print(f"  [{idx}] {b.name}")
    ch = input("Выбери номер бэкапа для отката (Enter = 1-й самый свежий): ").strip() or "1"
    if ch.isdigit() and 1 <= int(ch) <= len(backups):
        chosen = backups[int(ch) - 1]
        for f in chosen.glob("*.py"):
            shutil.copy2(f, config.HERE / f.name)
        print(f"⏪ Все модули успешно восстановлены из {chosen.name}!")
        restart_self()


def main_menu():
    cfg = config.load_config()
    disk_main = config.HERE / "main.py"
    if getattr(sys, "frozen", False) and cfg.get("allow_disk_override", True) and not globals().get("__DISK_MAIN_ACTIVE__") and disk_main.exists():
        try:
            reload_disk_modules()
            code = disk_main.read_text(encoding="utf-8-sig")
            g = globals().copy()
            g["__DISK_MAIN_ACTIVE__"] = True
            g["__file__"] = str(disk_main)
            exec(compile(code, str(disk_main), "exec"), g)
            return
        except Exception as e:
            print(f"\n🚨 Сбой во внешнем main.py ({e})! Запускаем встроенное ядро .exe...")
            time.sleep(2)

    reload_disk_modules()

    while True:
        os.system("cls" if os.name == "nt" else "clear")
        cfg = config.load_config()
        proc_ok = connector.is_farm_process_running()
        cdp_ok = bool(connector.get_cdp_ws_url())
        net_warn = connector.check_port_network_exposure() if hasattr(connector, "check_port_network_exposure") else ""

        if proc_ok and cdp_ok:
            status_str = "🟢 ЗАПУЩЕНА И ПОДКЛЮЧЕНА (127.0.0.1)"
        elif proc_ok and not cdp_ok:
            status_str = "🟡 Запущена БЕЗ порта автологгера (нажми [1] для перезапуска)"
        else:
            status_str = "⚪ Выключена"

        is_test = cfg.get("channel") == "test"
        branch_badge = "🧪 ВЕТКА: TEST (test)" if is_test else "🟢 ВЕТКА: STABLE (main)"

        live_st = report.get_live_state() if hasattr(report, "get_live_state") else {"active": False}
        if live_st.get("active"):
            live_header = f"🔴 LIVE-ЗАПИСЬ: ВКЛ (каждые {live_st['interval']}с | обновл. в {live_st['last_update']} | {live_st['file']})"
            btn4_text = "⏹️ Остановить фоновую Live-запись Excel"
        else:
            live_header = "⚪ LIVE-ЗАПИСЬ: ВЫКЛ"
            btn4_text = "🔴 Включить фоновую Live-запись Excel (автообновление в фоне)"

        print("==========================================================================")
        print(f"   BAMBU FARM CONTROL — v{config.VERSION} | {branch_badge}")
        print("==========================================================================")
        print(f" Статус фермы:  {status_str} | Одометров: {len(config.load_odometers())} шт.")
        print(f" Живая таблица: {live_header}")
        if net_warn:
            print(f" {net_warn}")
        print("--------------------------------------------------------------------------")
        print("  [1] 🚀 Запустить / Перезапустить Ферму + Автологгер")
        print("  [2] 📊 Сформировать полный Excel-отчёт (безопасно поверх открытого)")
        print("  [3] 📅 Отчёт за период (указать начальную и конечную дату вручную)")
        print(f"  [4] {btn4_text}")
        print("  [5] 🔧 Сверка и ручной ввод моточасов с экрана принтера (Одометр + ТО)")
        print("  [6] 🔄 Ручные повторы за прошлые дни")
        print("  [7] 🌐 GitHub: Обновить из STABLE (main) / TEST (test) / ⏪ Откатить")
        print("  [0] Выход")
        print("==========================================================================")
        choice = input("Выбери пункт меню: ").strip()

        try:
            if choice == "1":
                reload_disk_modules()
                connector.start_farm_with_hook()
                input("\nНажми Enter, чтобы вернуться в меню...")
            elif choice == "2":
                reload_disk_modules()
                report.generate_report()
                input("\nНажми Enter, чтобы вернуться в меню...")
            elif choice == "3":
                reload_disk_modules()
                report.generate_period_report_interactive()
                input("\nНажми Enter, чтобы вернуться в меню...")
            elif choice == "4":
                reload_disk_modules()
                report.toggle_live_excel_background()
            elif choice == "5":
                reload_disk_modules()
                report.odometer_reconciliation_menu()
            elif choice == "6":
                items = config.load_manual_reprints()
                if not items:
                    print("  (Список ручных повторов пуст)")
                for i, it in enumerate(items, 1):
                    print(f"  {i}. {it['date']} | {it['printer']} | {it['file']} | {it['qty']} шт.")
                sub = input("[1] Добавить  [2] Удалить  [Enter] Назад: ").strip()
                if sub == "1":
                    d_raw = input(f"Дата (Enter={datetime.now():%Y-%m-%d}): ").strip()
                    d = report.parse_user_date(d_raw, f"{datetime.now():%Y-%m-%d}")
                    pr = parser.clean_str(input("Принтер (109 AMS): "))
                    fl = parser.clean_str(input("Файл: "))
                    q = input("Штук: ").strip()
                    if pr and fl and q.isdigit() and int(q) > 0:
                        items.append({"date": d, "printer": pr, "file": fl, "qty": int(q)})
                        config.save_manual_reprints(items)
                elif sub == "2":
                    n = input("Номер строки: ").strip()
                    if n.isdigit() and 1 <= int(n) <= len(items):
                        items.pop(int(n) - 1)
                        config.save_manual_reprints(items)
            elif choice == "7":
                override_st = "ВКЛ" if cfg.get("allow_disk_override", True) else "ВЫКЛ (только ядро .exe)"
                print("\n--- 🌐 ВЫБОР ВЕТКИ И БЕЗОПАСНОСТЬ ---")
                print("  [1] 🟢 Обновить из STABLE ветки (main)")
                print("  [2] 🧪 Обновить из TEST ветки (test)")
                print("  [3] ⏪ Откатить модули на предыдущий бэкап")
                print("  [4] 🔑 Указать Fine-grained GitHub Token (Read-Only)")
                print(f"  [5] 🔒 Подгрузка внешних .py файлов в .exe: [{override_st}]")
                sub = input("> ").strip()
                if sub == "1":
                    sync_modules_from_github("main")
                elif sub == "2":
                    sync_modules_from_github("test")
                elif sub == "3":
                    rollback_modules()
                elif sub == "4":
                    tok = getpass.getpass("Вставь GitHub Token (ввод скрыт): ").strip() or input("Token: ").strip()
                    cfg["github_token"] = tok
                    config.save_config(cfg)
                    print("✅ Токен сохранён!")
                elif sub == "5":
                    cfg["allow_disk_override"] = not cfg.get("allow_disk_override", True)
                    config.save_config(cfg)
                    print(f"✅ Подгрузка внешних .py переключена: {'ВКЛ' if cfg['allow_disk_override'] else 'ВЫКЛ'}")
                input("\nНажми Enter...")
            elif choice == "0":
                if hasattr(report, "get_live_state"):
                    report.get_live_state()["active"] = False
                break
        except Exception as e:
            print(f"\n⚠️ Перехвачена ошибка (программа продолжает работу): {e}")
            traceback.print_exc(limit=3)
            input("\nНажми Enter, чтобы вернуться в меню...")


if __name__ == "__main__":
    main_menu()
