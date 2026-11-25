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


def first_run_wizard(force=False):
    """Окно приветствия и первичной настройки локации + автономного режима."""
    cfg = config.load_config()
    if cfg.get("setup_done") and cfg.get("location_name") and not force:
        return

    os.system("cls" if os.name == "nt" else "clear")
    print("=" * 78)
    print(" 👋 ДОБРО ПОЖАЛОВАТЬ В BAMBU FARM CONTROL v" + config.VERSION)
    print("=" * 78)
    print(" Первичная настройка компьютера фермы (выполняется один раз).\n")

    cur_loc = cfg.get("location_name") or ""
    while True:
        prompt_loc = f" 1. Укажите название локации для отчётов (например, Odesa / Kyiv) [{cur_loc}]: " if cur_loc else " 1. Укажите название локации для отчётов (например, Odesa / Kyiv): "
        loc = parser.clean_str(input(prompt_loc)) or cur_loc
        if loc:
            cfg["location_name"] = loc
            break
        print("    ❌ Название локации не может быть пустым.")

    cur_sd = cfg.get("start_date") or "2026-12-01"
    today_s = datetime.now().strftime("%Y-%m-%d")
    raw_sd = input(f" 2. Дата старта чистого Live-учёта [Enter = {cur_sd}, или '{today_s}' для теста]: ").strip()
    cfg["start_date"] = report.parse_user_date(raw_sd, cur_sd)

    print("\n 3. Включить автономный режим на этом ПК?")
    print("    • Добавит Farm_Control --auto в Автозагрузку Windows")
    print("    • Допишет порт отладки 19222 в ярлыки Bambu Farm Manager Client")
    ans_auto = input("    Настроить автозапуск и ярлыки сейчас? [Enter = Да / N = Пропустить]: ").strip()
    if ans_auto.upper() != "N":
        connector.setup_autostart_and_shortcuts(interactive=True)

    cfg["setup_done"] = True
    config.save_config(cfg)
    print(f"\n✅ Локация [{cfg['location_name']}] настроена! Учёт задач ведётся с {cfg['start_date']}.")
    input("\nНажми Enter для перехода в главное меню...")


def orders_map_menu():
    while True:
        rules = config.load_orders_map()
        print("\n" + "=" * 78)
        print(" 📦 СПРАВОЧНИК АВТОГРУППИРОВКИ ЗАКАЗОВ (orders_map.json)")
        print("=" * 78)
        print(" Поддерживаются маски * (например, CHYK_4.5.0* -> Заказ CHYK 4.5.0)")
        print("-" * 78)
        if not rules:
            print("  (Правил пока нет — все новые файлы можно группировать прямо в Excel)")
        for idx, r in enumerate(rules, 1):
            print(f"  [{idx:2d}] Маска: {r['pattern']:<28} -> Заказ: {r['order']:<22} | Деталей/стол: {r['parts_per_plate']}")
        print("-" * 78)
        sub = input("[1] Добавить правило  [2] Удалить правило  [Enter] Назад: ").strip()
        if sub == "1":
            pat = parser.clean_str(input("Имя файла или маска (например, CHYK_4.5.0*): "))
            ord_n = parser.clean_str(input("Название заказа: "))
            ppp_s = input("Деталей на 1 столе [Enter = 1]: ").strip()
            ppp = int(ppp_s) if ppp_s.isdigit() and int(ppp_s) > 0 else 1
            if pat and ord_n:
                rules.append({"pattern": pat, "order": ord_n, "parts_per_plate": ppp})
                config.save_orders_map(rules)
                print("✅ Правило сохранено!")
        elif sub == "2":
            n = input("Номер правила для удаления: ").strip()
            if n.isdigit() and 1 <= int(n) <= len(rules):
                rules.pop(int(n) - 1)
                config.save_orders_map(rules)
        else:
            break


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
    headers = {"User-Agent": "BambuFarmControl/3.0", "Cache-Control": "no-cache"}
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
    first_run_wizard(force=False)

    # Автономный запуск при старте Windows (--auto)
    if "--auto" in sys.argv:
        print("⚡ Автономный запуск (--auto): проверка ярлыков, запуск фермы и Live-записи...")
        connector.ensure_shortcuts_silently()
        connector.start_farm_with_hook(auto_mode=True)
        report.toggle_live_excel_background(auto_mode=True)

    while True:
        os.system("cls" if os.name == "nt" else "clear")
        cfg = config.load_config()
        loc_name = cfg.get("location_name") or "Не задана"
        start_d = cfg.get("start_date") or "2026-12-01"
        proc_ok = connector.is_farm_process_running()
        cdp_ok = bool(connector.get_cdp_ws_url())
        net_warn = connector.check_port_network_exposure()

        if proc_ok and cdp_ok:
            status_str = "🟢 ЗАПУЩЕНА И ПОДКЛЮЧЕНА (127.0.0.1)"
        elif proc_ok and not cdp_ok:
            status_str = "🟡 Запущена БЕЗ порта автологгера (нажми [1] для перезапуска)"
        else:
            status_str = "⚪ Выключена"

        is_test = cfg.get("channel") == "test"
        branch_badge = "🧪 TEST" if is_test else "🟢 STABLE"

        live_st = report.get_live_state()
        if live_st.get("active"):
            live_header = (
                f"🔴 LIVE-ЗАПИСЬ: ВКЛ (обн. {live_st['last_update']} | "
                f"✅ {live_st.get('done_cnt', 0)} / ❌ {live_st.get('fail_cnt', 0)} | {live_st['file']})"
            )
            btn4_text = "⏹️ Остановить фоновую Live-запись"
        else:
            live_header = "⚪ LIVE-ЗАПИСЬ: ВЫКЛ"
            btn4_text = "🔴 Включить фоновую Live-запись (архив + Excel + пульс)"

        print("==========================================================================")
        print(f"   BAMBU FARM CONTROL v{config.VERSION} | Локация: {loc_name} | {branch_badge}")
        print("==========================================================================")
        print(f" Статус фермы:  {status_str} | Учёт с: {start_d}")
        print(f" Живая таблица: {live_header}")
        if net_warn:
            print(f" {net_warn}")
        print("--------------------------------------------------------------------------")
        print("  [1] 🚀 Запустить / Перезапустить Ферму + Автологгер")
        print("  [2] 📊 Сформировать и открыть Excel-отчёт (Общая сводка + Заказы)")
        print("  [3] 📅 Отчёт за период (указать начальную и конечную дату вручную)")
        print(f"  [4] {btn4_text}")
        print("  [5] 🔧 Сверка моточасов с экрана принтера (Одометр + ТО)")
        print("  [6] 📦 Справочник автогруппировки заказов / Ручные повторы")
        print("  [7] ⚙️ Локация, Автозагрузка, Ярлыки и Обновления GitHub")
        print("  [0] Выход")
        print("==========================================================================")
        choice = input("Выбери пункт меню: ").strip()

        try:
            if choice == "1":
                reload_disk_modules()
                connector.start_farm_with_hook(auto_mode=False)
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
                report.toggle_live_excel_background(auto_mode=False)
            elif choice == "5":
                reload_disk_modules()
                report.odometer_reconciliation_menu()
            elif choice == "6":
                print("\n[1] 📦 Правила автогруппировки заказов (маски + деталей/стол)  [2] 🔄 Ручные повторы")
                sub = input("> ").strip()
                if sub == "1":
                    orders_map_menu()
                elif sub == "2":
                    items = config.load_manual_reprints()
                    if not items:
                        print("  (Список ручных повторов пуст)")
                    for i, it in enumerate(items, 1):
                        print(f"  {i}. {it['date']} | {it['printer']} | {it['file']} | {it['qty']} шт.")
                    sub2 = input("[1] Добавить  [2] Удалить  [Enter] Назад: ").strip()
                    if sub2 == "1":
                        d_raw = input(f"Дата (Enter={datetime.now():%Y-%m-%d}): ").strip()
                        d = report.parse_user_date(d_raw, f"{datetime.now():%Y-%m-%d}")
                        pr = parser.clean_str(input("Принтер (109 AMS): "))
                        fl = parser.clean_str(input("Файл: "))
                        q = input("Штук: ").strip()
                        if pr and fl and q.isdigit() and int(q) > 0:
                            items.append({"date": d, "printer": pr, "file": fl, "qty": int(q)})
                            config.save_manual_reprints(items)
                    elif sub2 == "2":
                        n = input("Номер строки: ").strip()
                        if n.isdigit() and 1 <= int(n) <= len(items):
                            items.pop(int(n) - 1)
                            config.save_manual_reprints(items)
            elif choice == "7":
                override_st = "ВКЛ" if cfg.get("allow_disk_override", True) else "ВЫКЛ (только ядро .exe)"
                print("\n--- ⚙️ СИСТЕМА, ЛОКАЦИЯ И GITHUB ---")
                print("  [1] 🟢 Обновить модули с GitHub — STABLE (main)")
                print("  [2] 🧪 Обновить модули с GitHub — TEST (test)")
                print("  [3] ⏪ Откатить модули на предыдущий бэкап")
                print("  [4] 🏷️ Изменить локацию / дату старта учёта")
                print("  [5] 🤖 Настроить Автозапуск + ярлыки Bambu (повторить мастер)")
                print("  [6] 🧹 Удалить Автозапуск и очистить ярлыки Bambu")
                print("  [7] 🔑 Указать GitHub Token (Read-Only)")
                print(f"  [8] 🔒 Подгрузка внешних .py файлов в .exe: [{override_st}]")
                sub = input("> ").strip()
                if sub == "1":
                    sync_modules_from_github("main")
                elif sub == "2":
                    sync_modules_from_github("test")
                elif sub == "3":
                    rollback_modules()
                elif sub == "4":
                    first_run_wizard(force=True)
                elif sub == "5":
                    connector.setup_autostart_and_shortcuts(interactive=True)
                elif sub == "6":
                    connector.remove_autostart_and_shortcuts()
                elif sub == "7":
                    tok = getpass.getpass("Вставь GitHub Token (ввод скрыт): ").strip() or input("Token: ").strip()
                    cfg["github_token"] = tok
                    config.save_config(cfg)
                    print("✅ Токен сохранён!")
                elif sub == "8":
                    cfg["allow_disk_override"] = not cfg.get("allow_disk_override", True)
                    config.save_config(cfg)
                    print(f"✅ Подгрузка внешних .py: {'ВКЛ' if cfg['allow_disk_override'] else 'ВЫКЛ'}")
                input("\nНажми Enter...")
            elif choice == "0":
                report.get_live_state()["active"] = False
                connector.send_heartbeat(live=False)
                break
        except Exception as e:
            print(f"\n⚠️ Перехвачена ошибка (программа продолжает работу): {e}")
            traceback.print_exc(limit=3)
            input("\nНажми Enter, чтобы вернуться в меню...")


if __name__ == "__main__":
    main_menu()
