import json
import sys
import threading
from pathlib import Path

VERSION = "2.5.0-sec"
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

# Строгий список папок и лимит размера файла (200 МБ)
SEARCH_DIRS = [HERE, Path.home() / "Downloads", Path.home() / "Desktop"]
MAX_JSON_BYTES = 200 * 1024 * 1024

FARM_EXE_NAME = "Bambu Farm Manager Client.exe"
FARM_EXE = Path(rf"C:\Program Files\Bambu Farm Manager Client\{FARM_EXE_NAME}")

CDP_PORT = 19222
CDP_SCAN_PORTS = (19222, 9222)
FARM_MARKERS = ("bambu", "farm", "app.asar")

# Единая межпоточная блокировка (защита от гонки потоков в режиме Live Excel)
if not hasattr(sys, "_bambu_rlock"):
    sys._bambu_rlock = threading.RLock()


def get_lock() -> threading.RLock:
    return sys._bambu_rlock


def load_json(path: Path, default):
    if path.exists():
        try:
            if path.stat().st_size > MAX_JSON_BYTES:
                print(f"⚠️ Файл {path.name} превышает 200 МБ и пропущен.")
                return default
            return json.loads(path.read_text(encoding="utf-8-sig"))
        except Exception as e:
            print(f"⚠️ Повреждён JSON {path.name} ({e}) — пропущен.")
    return default


def save_json(path: Path, data):
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(path)


def load_config():
    default_cfg = {
        "channel": "stable",
        "github_repo": GITHUB_REPO,
        "github_token": "",
        "allow_disk_override": True,  # Можно отключить на цеховых ПК, чтобы .exe работал строго изнутри
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
    # Без тестовых фантомных записей: пустой список, пока пользователь сам не добавит повтор
    if not MANUAL_FILE.exists():
        return []
    raw = load_json(MANUAL_FILE, [])
    valid = []
    if isinstance(raw, list):
        for it in raw:
            if isinstance(it, dict) and it.get("printer") and it.get("file"):
                try:
                    q = int(it.get("qty", 0))
                    if q > 0:
                        valid.append({
                            "date": str(it.get("date", ""))[:10],
                            "printer": str(it["printer"]).strip(),
                            "file": str(it["file"]).strip(),
                            "qty": q,
                        })
                except Exception:
                    pass
    return valid


def save_manual_reprints(items):
    save_json(MANUAL_FILE, items)
