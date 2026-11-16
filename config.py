import json
import sys
from pathlib import Path

VERSION = "2.2.0-modular"
GITHUB_REPO = "TelliBronde/BL_Autologger"
MODULES = ["config.py", "connector.py", "parser.py", "report.py", "main.py"]

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

SEARCH_DIRS = [HERE, Path.cwd(), Path.home() / "Downloads", Path.home() / "Desktop"]
FARM_EXE_NAME = "Bambu Farm Manager Client.exe"
FARM_EXE = Path(rf"C:\Program Files\Bambu Farm Manager Client\{FARM_EXE_NAME}")

# Используем отдельный порт 19222 для запуска, чтобы не пересекаться с Chrome/Edge на 9222
CDP_PORT = 19222
CDP_SCAN_PORTS = (19222, 9222, 9333)


def load_json(path: Path, default):
    if path.exists():
        try:
            return json.loads(path.read_text(encoding="utf-8-sig"))
        except Exception:
            pass
    return default


def save_json(path: Path, data):
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def load_config():
    default_cfg = {
        "channel": "stable",  # "stable" (ветка main) или "test" (ветка test)
        "github_repo": GITHUB_REPO,
        "github_token": "",
        "auto_update_on_start": False,
        "to_interval_hours": 300,
    }
    cfg = load_json(CONFIG_FILE, {})
    default_cfg.update(cfg)
    if not CONFIG_FILE.exists():
        save_json(CONFIG_FILE, default_cfg)
    return default_cfg


def save_config(cfg):
    save_json(CONFIG_FILE, cfg)


def load_odometers():
    return load_json(ODOMETER_FILE, {})


def save_odometers(data):
    save_json(ODOMETER_FILE, data)


def load_manual_reprints():
    default_list = [
        {"date": "2026-11-12", "printer": "109 AMS", "file": "Poland_test_numero_uno", "qty": 4}
    ]
    if not MANUAL_FILE.exists():
        save_json(MANUAL_FILE, default_list)
        return default_list
    return load_json(MANUAL_FILE, [])


def save_manual_reprints(items):
    save_json(MANUAL_FILE, items)