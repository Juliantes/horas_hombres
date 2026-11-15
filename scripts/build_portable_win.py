#!/usr/bin/env python3
"""Arma dist/HorasHombre para Windows 64-bit (sin admin, sin red en destino)."""
from __future__ import annotations

import re
import shutil
import subprocess
import sys
import urllib.request
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DIST = ROOT / "dist" / "HorasHombre"
PYTHON_VER = "3.12.8"
EMBED_URL = (
    f"https://www.python.org/ftp/python/{PYTHON_VER}/"
    f"python-{PYTHON_VER}-embed-amd64.zip"
)
REQUIREMENTS = ROOT / "requirements-portable.txt"
LEEME_SRC = ROOT / "dist-templates" / "LEEME.txt"


def _download(url: str, dest: Path) -> None:
    dest.parent.mkdir(parents=True, exist_ok=True)
    print(f"Descargando {url} …")
    urllib.request.urlretrieve(url, dest)


def _patch_pth(python_dir: Path) -> None:
    pth_files = list(python_dir.glob("python*._pth"))
    if not pth_files:
        raise FileNotFoundError(f"No se encontró python*._pth en {python_dir}")
    pth = pth_files[0]
    text = pth.read_text(encoding="utf-8")
    text = re.sub(r"#?\s*import site", "import site", text)
    site_line = "Lib\\site-packages"
    if site_line not in text:
        text = text.rstrip() + f"\n{site_line}\n"
    pth.write_text(text, encoding="utf-8")
    (python_dir / "Lib" / "site-packages").mkdir(parents=True, exist_ok=True)


def _pip_install(site_packages: Path) -> None:
    cmd = [
        sys.executable,
        "-m",
        "pip",
        "install",
        "-r",
        str(REQUIREMENTS),
        "--target",
        str(site_packages),
        "--platform",
        "win_amd64",
        "--python-version",
        "312",
        "--implementation",
        "cp",
        "--only-binary",
        ":all:",
        "--upgrade",
    ]
    print("Instalando dependencias Windows (wheels amd64) …")
    subprocess.run(cmd, check=True)


def _copy_app() -> None:
    app = DIST / "app"
    app.mkdir(parents=True, exist_ok=True)
    for name in ("horas_hombre.py", "sap_import.py"):
        shutil.copy2(ROOT / name, app / name)
    shutil.copy2(ROOT / "app" / "launch.py", app / "launch.py")


def _copy_launcher_files() -> None:
    shutil.copy2(ROOT / "HorasHombre.bat", DIST / "HorasHombre.bat")
    if LEEME_SRC.is_file():
        shutil.copy2(LEEME_SRC, DIST / "LEEME.txt")
    else:
        (DIST / "LEEME.txt").write_text(
            "Copie esta carpeta a la PC Windows y ejecute HorasHombre.bat.\n",
            encoding="utf-8",
        )


def main() -> None:
    if DIST.exists():
        shutil.rmtree(DIST)
    DIST.mkdir(parents=True)

    zip_path = ROOT / "build" / f"python-{PYTHON_VER}-embed-amd64.zip"
    _download(EMBED_URL, zip_path)

    py_dir = DIST / "python"
    py_dir.mkdir(parents=True)
    with zipfile.ZipFile(zip_path) as zf:
        zf.extractall(py_dir)

    _patch_pth(py_dir)
    site = py_dir / "Lib" / "site-packages"
    _pip_install(site)

    _copy_app()
    _copy_launcher_files()
    (DIST / "data").mkdir(parents=True, exist_ok=True)

    print(f"\nListo: {DIST}")
    print("Copie la carpeta HorasHombre a un USB o PC Windows 64-bit y ejecute HorasHombre.bat.")


if __name__ == "__main__":
    main()