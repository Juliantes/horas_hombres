"""Arranque portable: Streamlit en localhost, datos en ../data."""
from __future__ import annotations

import os
import socket
import subprocess
import sys
import time
import webbrowser
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
APP_DIR = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
APP_SCRIPT = APP_DIR / "horas_hombre.py"
if not APP_SCRIPT.is_file():
    APP_SCRIPT = ROOT / "horas_hombre.py"


def _puerto_libre(desde: int = 8501, intentos: int = 20) -> int:
    for p in range(desde, desde + intentos):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", p))
                return p
            except OSError:
                continue
    return desde


def main() -> int:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    mpl_cache = DATA_DIR / ".matplotlib"
    mpl_cache.mkdir(parents=True, exist_ok=True)

    os.environ["HH_DATA_DIR"] = str(DATA_DIR)
    os.environ["MPLCONFIGDIR"] = str(mpl_cache)
    os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
    os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"

    port = _puerto_libre()
    url = f"http://127.0.0.1:{port}"

    cmd = [
        sys.executable,
        "-m",
        "streamlit",
        "run",
        str(APP_SCRIPT),
        "--server.address",
        "127.0.0.1",
        "--server.port",
        str(port),
        "--server.headless",
        "true",
        "--browser.gatherUsageStats",
        "false",
        "--server.fileWatcherType",
        "none",
    ]

    print("Horas Hombre — no cierre esta ventana mientras usa la aplicación.")
    print(f"Datos: {DATA_DIR}")
    print(f"Si el navegador no abre solo, vaya a: {url}")

    proc = subprocess.Popen(cmd, cwd=str(APP_DIR))
    time.sleep(2.5)
    webbrowser.open(url)

    try:
        return proc.wait()
    except KeyboardInterrupt:
        proc.terminate()
        proc.wait(timeout=10)
        return 0


if __name__ == "__main__":
    raise SystemExit(main())