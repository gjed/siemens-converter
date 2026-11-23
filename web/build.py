"""Build the static site: python web/build.py --out build/site

The site is the page in ``web/``, the converter source and ``template.xlsx`` of the
checked-out commit, a pinned Pyodide and the pinned wheels micropip installs. Every
file the page loads comes from the site itself, never from a CDN.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import subprocess
import sys
import tarfile
import urllib.request
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PACKAGE = ROOT / "src" / "siemens_converter"

PYODIDE_VERSION = "314.0.8"  # runs Python 3.14.2
CORE_URL = (
    f"https://github.com/pyodide/pyodide/releases/download/{PYODIDE_VERSION}/"
    f"pyodide-core-{PYODIDE_VERSION}.tar.bz2"
)
CORE_SHA256 = "ff3e5f9d9bc8e98ff30b99515535fc24d375994f47a978c61975d653e162ca37"
# Pyodide's own builds of these two, checked against the lock file inside the core tarball
PACKAGES_URL = f"https://cdn.jsdelivr.net/pyodide/v{PYODIDE_VERSION}/full/"
PYODIDE_PACKAGES = ("micropip", "packaging")
CORE_FILES = (
    "pyodide.mjs",
    "pyodide.asm.mjs",
    "pyodide.asm.wasm",
    "python_stdlib.zip",
    "pyodide-lock.json",
)
PAGE_FILES = ("index.html", "app.js", "worker.js", "style.css")


def download(url: str, destination: Path, sha256: str) -> None:
    """Fetch *url* into *destination*, unless it is already there with that hash."""
    if destination.exists() and _sha256(destination) == sha256:
        return
    destination.parent.mkdir(parents=True, exist_ok=True)
    with urllib.request.urlopen(url, timeout=120) as response:  # noqa: S310
        destination.write_bytes(response.read())
    found = _sha256(destination)
    if found != sha256:
        destination.unlink()
        raise SystemExit(f"{url}: sha256 is {found}, expected {sha256}")


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def build_pyodide(out: Path, cache: Path) -> None:
    core = cache / f"pyodide-core-{PYODIDE_VERSION}.tar.bz2"
    download(CORE_URL, core, CORE_SHA256)
    target = out / "pyodide"
    target.mkdir(parents=True)
    with tarfile.open(core) as archive:
        for name in CORE_FILES:
            member = archive.extractfile(f"pyodide/{name}")
            if member is None:
                raise SystemExit(f"{core.name} has no {name}")
            (target / name).write_bytes(member.read())
    packages = json.loads((target / "pyodide-lock.json").read_text())["packages"]
    for name in PYODIDE_PACKAGES:
        wheel, sha256 = packages[name]["file_name"], packages[name]["sha256"]
        download(PACKAGES_URL + wheel, cache / wheel, sha256)
        shutil.copy(cache / wheel, target / wheel)


def build_wheels(out: Path) -> list[str]:
    wheels = out / "wheels"
    subprocess.run(
        [
            sys.executable,
            "-m",
            "pip",
            "download",
            "--no-deps",
            "--only-binary=:all:",
            "--require-hashes",
            "--requirement",
            str(ROOT / "web" / "wheels.txt"),
            "--dest",
            str(wheels),
        ],
        check=True,
    )
    return sorted(f"wheels/{wheel.name}" for wheel in wheels.glob("*.whl"))


def build_source(out: Path) -> None:
    sources = sorted(PACKAGE.glob("*.py")) + [PACKAGE / "template.xlsx"]
    with zipfile.ZipFile(out / "siemens_converter.zip", "w", zipfile.ZIP_DEFLATED) as zf:
        for path in sources:
            zf.write(path, f"siemens_converter/{path.name}")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--out", type=Path, default=ROOT / "build" / "site")
    parser.add_argument(
        "--cache",
        type=Path,
        default=ROOT / "build" / "cache",
        help="where the Pyodide downloads are kept between builds",
    )
    args = parser.parse_args()

    if args.out.exists():
        shutil.rmtree(args.out)
    args.out.mkdir(parents=True)
    for name in PAGE_FILES:
        shutil.copy(ROOT / "web" / name, args.out / name)
    build_source(args.out)
    build_pyodide(args.out, args.cache)
    wheels = build_wheels(args.out)
    (args.out / "config.json").write_text(json.dumps({"wheels": wheels}, indent=2) + "\n")
    print(f"Site built in {args.out}")


if __name__ == "__main__":
    main()
