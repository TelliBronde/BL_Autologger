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
    """Динамически подгружает свежие .py модули с диска (даже если запущен из .exe)."""
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

    config.save_config(cfg)
    repo = cfg.get("github_repo", config.GITHUB_REPO)
    token = (cfg.get("github_token") or "").strip()
    headers = {"User-Agent": "BambuFarmControl/2.4", "Cache-Control": "no-cache"}
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
        reload_disk_modules()
        print(f"\n🎉 Успешно обновлено модулей: {updated} из ветки [{branch.upper()}]!")
        print(f"   Бэкап прошлой версии сохранён в: versions/{bak_dir.name}")
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
        reload_disk_modules()
        print(f"⏪ Все модули успешно восстановлены из {chosen.name}!")


def main_menu():
    disk_main = config.HERE / "main.py"
    if getattr(sys, "frozen", False) and not globals().get("__DISK_MAIN_ACTIVE__") and disk_main.exists():
        try:
            reload_disk_modules()
            code = disk_main.read_text(encoding="utf-8-sig")
            g = globals().copy()
            g["__DISK_MAIN_ACTIVE__"] = True
            g["__file__"] = str(disk_main)
            exec(compile(code, str(disk_main), "exec"), g)
            return
        except Exception as e:
            print(f"\n🚨 Сбой во внешнем main.py ({e})! Запускаем встроенное аварийное меню .exe...")
            time.sleep(2)

    reload_disk_modules()

    if "--stable" in sys.argv:
        sync_modules_from_github("main")
    elif "--test" in sys.argv:
        sync_modules_from_github("test")

    while True:
        os.system("cls" if os.name == "nt" else "clear")
        cfg = config.load_config()
        proc_ok = connector.is_farm_process_running()
        cdp_ok = bool(connector.get_cdp_ws_url())

        if proc_ok and cdp_ok:
            status_str = "🟢 ЗАПУЩЕНА И ПОДКЛЮЧЕНА"
        elif proc_ok and not cdp_ok:
            status_str = "🟡 Запущена БЕЗ порта автологгера (нажми [1] для перезапуска)"
        else:
            status_str = "⚪ Выключена"

        is_test = cfg.get("channel") == "test"
        branch_badge = "🧪 ВЕТКА: TEST (test)" if is_test else "🟢 ВЕТКА: STABLE (main)"

        live_st = report.get_live_state()
        if live_st["active"]:
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
                for i, it in enumerate(items, 1):
                    print(f"  {i}. {it['date']} | {it['printer']} | {it['file']} | {it['qty']} шт.")
                sub = input("[1] Добавить  [2] Удалить  [Enter] Назад: ").strip()
                if sub == "1":
                    d = input(f"Дата (Enter={datetime.now():%Y-%m-%d}): ").strip() or f"{datetime.now():%Y-%m-%d}"
                    pr = input("Принтер (109 AMS): ").strip()
                    fl = input("Файл: ").strip()
                    q = input("Штук: ").strip()
                    if pr and fl and q.isdigit():
                        items.append({"date": d, "printer": pr, "file": fl, "qty": int(q)})
                        config.save_manual_reprints(items)
                elif sub == "2":
                    n = input("Номер строки: ").strip()
                    if n.isdigit() and 1 <= int(n) <= len(items):
                        items.pop(int(n) - 1)
                        config.save_manual_reprints(items)
            elif choice == "7":
                print("\n--- 🌐 ВЫБОР ВЕТКИ ОБНОВЛЕНИЯ GITHUB ---")
                print("  [1] 🟢 Обновить из STABLE ветки (main)")
                print("  [2] 🧪 Обновить из TEST ветки (test)")
                print("  [3] ⏪ Откатить модули на предыдущий бэкап")
                print("  [4] 🔑 Указать GitHub Token (для приватного репозитория)")
                sub = input("> ").strip()
                if sub == "1":
                    sync_modules_from_github("main")
                elif sub == "2":
                    sync_modules_from_github("test")
                elif sub == "3":
                    rollback_modules()
                elif sub == "4":
                    cfg["github_token"] = input("GitHub Token: ").strip()
                    config.save_config(cfg)
                    print("✅ Токен сохранён!")
                input("\nНажми Enter...")
            elif choice == "0":
                report.get_live_state()["active"] = False
                break
        except Exception as e:
            print(f"\n⚠️ Перехвачена ошибка (программа продолжает работу): {e}")
            traceback.print_exc(limit=3)
            input("\nНажми Enter, чтобы вернуться в меню...")


if __name__ == "__main__":
    main_menu()
