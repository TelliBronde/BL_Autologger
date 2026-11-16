import importlib
import importlib.util
import os
import shutil
import sys
import time
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


def sync_modules_from_github():
    cfg = config.load_config()
    branch = "test" if cfg.get("channel") == "test" else "main"
    repo = cfg.get("github_repo", config.GITHUB_REPO)
    token = (cfg.get("github_token") or "").strip()
    headers = {"User-Agent": "BambuFarmControl/2.2", "Cache-Control": "no-cache"}
    if token:
        headers["Authorization"] = f"token {token}"

    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    bak_dir = config.VERSIONS_DIR / f"backup_{ts}"
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
            compile(code_text, m, "exec")  # Проверка синтаксиса перед записью
            (config.HERE / m).write_text(code_text, encoding="utf-8")
            updated += 1
            print(f"  ✅ Обновлён модуль: {m}")
        except Exception as e:
            print(f"  ⚠️ Пропущен {m}: {e}")

    if updated:
        reload_disk_modules()
        print(f"\n🎉 Обновлено модулей: {updated}. Бэкап прошлой версии: {bak_dir.name}")


def rollback_modules():
    backups = sorted([d for d in config.VERSIONS_DIR.iterdir() if d.is_dir()], key=lambda p: p.stat().st_mtime, reverse=True)
    if not backups:
        print("Архив бэкапов пока пуст.")
        return
    for idx, b in enumerate(backups[:10], 1):
        print(f"  [{idx}] {b.name}")
    ch = input("Выбери номер бэкапа для отката (Enter = 1-й самый свежий): ").strip() or "1"
    if ch.isdigit() and 1 <= int(ch) <= len(backups):
        chosen = backups[int(ch) - 1]
        for f in chosen.glob("*.py"):
            shutil.copy2(f, config.HERE / f.name)
        reload_disk_modules()
        print(f"⏪ Модули успешно откачены из {chosen.name}!")


def main_menu():
    reload_disk_modules()
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

        print("==========================================================================")
        print(f"   BAMBU FARM CONTROL — v{config.VERSION} | Репо: {cfg.get('github_repo')}")
        print("==========================================================================")
        print(f" Статус фермы: {status_str} | Одометров: {len(config.load_odometers())} шт.")
        print("--------------------------------------------------------------------------")
        print("  [1] 🚀 Запустить / Перезапустить Ферму + Автологгер")
        print("  [2] 📊 Сформировать и открыть Excel-отчёт (выработка + Моточасы)")
        print("  [3] 🔧 Сверка и ручной ввод моточасов с экрана принтера (Одометр + ТО)")
        print("  [4] 🔄 Ручные повторы за прошлые дни")
        print("  [5] 🌐 Обновить модули с GitHub / ⏪ Откатить из бэкапа")
        print("  [0] Выход")
        print("==========================================================================")
        choice = input("Выбери пункт меню: ").strip()

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
            report.odometer_reconciliation_menu()
        elif choice == "4":
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
        elif choice == "5":
            print("\n[1] 🔄 Обновить модули с GitHub  [2] ⏪ Откатить на прошлый бэкап  [3] 🔑 Ввести GitHub Token")
            sub = input("> ").strip()
            if sub == "1":
                sync_modules_from_github()
            elif sub == "2":
                rollback_modules()
            elif sub == "3":
                cfg["github_token"] = input("GitHub Token: ").strip()
                config.save_config(cfg)
            input("\nНажми Enter...")
        elif choice == "0":
            break


if __name__ == "__main__":
    main_menu()