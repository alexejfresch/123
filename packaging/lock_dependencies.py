"""Закрепляет установленное окружение Python 3.12 и SHA256 из официального PyPI.

Запускать вручную после осознанного обновления зависимостей в чистом venv.
Включает дополнительные зависимости Windows, отсутствующие на Linux.
"""
from concurrent.futures import ThreadPoolExecutor
from importlib.metadata import distributions
import json
from pathlib import Path
import urllib.request

ROOT = Path(__file__).resolve().parents[1]
MARKERS = {"secretstorage": 'sys_platform == "linux"', "jeepney": 'sys_platform == "linux"',
           "pefile": 'sys_platform == "win32"', "pywin32-ctypes": 'sys_platform == "win32"',
           "colorama": 'sys_platform == "win32"'}


def fetch(item):
    name, version = item
    with urllib.request.urlopen(f"https://pypi.org/pypi/{name}/{version}/json", timeout=30) as response:
        metadata = json.load(response)
    hashes = sorted({f["digests"]["sha256"] for f in metadata["urls"] if not f.get("yanked")})
    if not hashes:
        raise RuntimeError(f"Нет проверяемых артефактов: {name}")
    requirement = f"{name}=={version}"
    if name in MARKERS:
        requirement += "; " + MARKERS[name]
    return requirement + " \\\n" + " \\\n".join("    --hash=sha256:" + h for h in hashes)


def main():
    packages = {d.metadata["Name"].lower().replace("_", "-"): d.version for d in distributions()
                if d.metadata["Name"].lower() not in {"pip", "uik-reconcile"}}
    packages.update({"pefile": "2024.8.26", "pywin32-ctypes": "0.2.3", "colorama": "0.4.6"})
    with ThreadPoolExecutor(max_workers=6) as pool:
        lines = list(pool.map(fetch, sorted(packages.items())))
    (ROOT / "requirements.lock").write_text(
        "# Python 3.12, Windows x64 / Linux x64. SHA256: official pypi.org metadata.\n"
        "# Regenerate: python packaging/lock_dependencies.py (clean project venv only).\n\n"
        + "\n\n".join(lines) + "\n", encoding="utf-8")
    print(f"Закреплено пакетов: {len(lines)}")


if __name__ == "__main__":
    main()
