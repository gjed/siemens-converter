"""Headless browser test of the built site: python -m pytest web/tests

Build the site first (``python web/build.py``); ``SITE_DIR`` names another build.
Needs the ``web`` extra (``pip install -e ".[web]"``) and ``playwright install chromium``.
The truth for every expectation is what the command line writes and prints for the
same files.
"""

from __future__ import annotations

import functools
import http.server
import os
import re
import shutil
import subprocess
import sys
import threading
from pathlib import Path
from urllib.parse import urlparse

import openpyxl
import pytest
from playwright.sync_api import Page, expect, sync_playwright

ROOT = Path(__file__).resolve().parents[2]
SITE = Path(os.environ.get("SITE_DIR", ROOT / "build" / "site"))
FIXTURES = ROOT / "tests" / "fixtures"
FC_REPORT = FIXTURES / "FC_report_TEST_9999_2026-01-15.xls"
STATIC_DATA = FIXTURES / "static_data_TEST.xlsx"
OUTPUT_NAME = "Riparto_TE99999999_15-01-2026.xlsx"
PROGRESS = re.compile(r"^(Lettura del report|Lettura dei dati statici da|Scrittura di) ")
SLOW = 180_000  # ms: Pyodide has to load


class Handler(http.server.SimpleHTTPRequestHandler):
    extensions_map = {
        **http.server.SimpleHTTPRequestHandler.extensions_map,
        ".mjs": "text/javascript",
        ".wasm": "application/wasm",
    }

    def log_message(self, *args) -> None:
        pass


@pytest.fixture(scope="session")
def origin():
    if not (SITE / "index.html").exists():
        pytest.fail(f"No site in {SITE}: run python web/build.py --out {SITE}")
    handler = functools.partial(Handler, directory=str(SITE))
    server = http.server.ThreadingHTTPServer(("127.0.0.1", 0), handler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    yield f"http://127.0.0.1:{server.server_port}"
    server.shutdown()


@pytest.fixture(scope="session")
def browser():
    with sync_playwright() as playwright:
        chromium = playwright.chromium.launch()
        yield chromium
        chromium.close()


class Visit:
    """A page opened on the site, with everything it requested and logged."""

    def __init__(self, page: Page, requests: list[str], problems: list[str]) -> None:
        self.page = page
        self.requests = requests
        self.problems = problems


@pytest.fixture
def visit(browser, origin):
    context = browser.new_context(accept_downloads=True)
    context.set_default_timeout(SLOW)
    requests: list[str] = []
    problems: list[str] = []
    context.on("request", lambda request: requests.append(request.url))
    page = context.new_page()
    page.on("pageerror", lambda error: problems.append(f"pageerror: {error}"))
    page.on(
        "console",
        lambda m: problems.append(f"console {m.type}: {m.text}") if m.type == "error" else None,
    )
    page.goto(origin)
    yield Visit(page, requests, problems)
    context.close()


def cli(*files: Path) -> tuple[int, str]:
    """Run the command line on *files*; return its exit code and its output."""
    done = subprocess.run(
        [sys.executable, "-m", "siemens_converter", *map(str, files)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        check=False,
    )
    return done.returncode, done.stdout


def cli_message(output: str) -> str:
    """The error message the CLI printed, without its progress lines."""
    return "\n".join(line for line in output.splitlines() if not PROGRESS.match(line)).strip()


def cli_warnings(output: str) -> list[str]:
    _, _, tail = output.partition("Attenzione:\n")
    return [line[2:] for line in tail.splitlines() if line.startswith("- ")]


def copy(source: Path, directory: Path, name: str | None = None) -> Path:
    directory.mkdir(parents=True, exist_ok=True)
    return Path(shutil.copy(source, directory / (name or source.name)))


def cells(path: Path) -> dict[str, list[list[tuple]]]:
    """Every sheet as rows of (value, number format), formulas as written."""
    workbook = openpyxl.load_workbook(path)
    return {
        ws.title: [[(c.value, c.number_format) for c in row] for row in ws.iter_rows()]
        for ws in workbook
    }


def choose(visit: Visit, *files: Path) -> None:
    visit.page.wait_for_selector("#loading", state="hidden")
    visit.page.set_input_files("#files", [str(f) for f in files])


def refusal_on_page(visit: Visit) -> str:
    error = visit.page.locator("#outcome .message.error")
    expect(error).to_be_visible()
    assert visit.page.locator("#download").count() == 0
    return error.inner_text().strip()


@pytest.mark.parametrize("reverse", [False, True], ids=["report-first", "static-first"])
def test_page_writes_the_workbook_the_cli_writes(visit, tmp_path, reverse):
    report = copy(FC_REPORT, tmp_path / "cli")
    static = copy(STATIC_DATA, tmp_path / "cli")
    code, output = cli(report, static)
    assert code == 0
    expected = report.parent / OUTPUT_NAME

    page = visit.page
    assert page.locator("#loading").is_visible()  # the loading state shows first
    files = [STATIC_DATA, FC_REPORT] if reverse else [FC_REPORT, STATIC_DATA]
    choose(visit, *files)
    expect(page.locator("#download")).to_be_visible()
    with page.expect_download() as download:
        page.click("#download")
    assert download.value.suggested_filename == OUTPUT_NAME
    got = tmp_path / "page.xlsx"
    download.value.save_as(got)

    assert cells(got) == cells(expected)
    assert [li.inner_text() for li in page.locator("#warnings li").all()] == cli_warnings(output)
    assert visit.problems == []


def test_warnings_are_listed_under_attenzione_like_the_cli_prints_them(visit, tmp_path):
    report = copy(FC_REPORT, tmp_path)
    text = report.read_text(encoding="utf-8")
    report.write_text(text.replace("<td>Ok</td>", "<td>Errore</td>", 1), encoding="utf-8")
    static = copy(STATIC_DATA, tmp_path)
    code, output = cli(report, static)
    assert code == 0
    assert cli_warnings(output)

    choose(visit, report, static)

    expect(visit.page.locator("#download")).to_be_visible()
    assert visit.page.locator("#warnings h2").inner_text() == "Attenzione"
    shown = [li.inner_text() for li in visit.page.locator("#warnings li").all()]
    assert shown == cli_warnings(output)


def test_page_makes_no_request_outside_its_own_origin(visit, origin):
    choose(visit, FC_REPORT, STATIC_DATA)
    expect(visit.page.locator("#download")).to_be_visible()

    host = urlparse(origin).netloc
    outside = [
        u for u in visit.requests if urlparse(u).scheme not in ("blob", "data")
        and urlparse(u).netloc != host
    ]
    print(f"\n{len(visit.requests)} requests seen:\n" + "\n".join(sorted(set(visit.requests))))
    assert outside == []
    assert any(u.endswith("pyodide.asm.wasm") for u in visit.requests)  # workers are seen
    assert any(u.endswith(".whl") and "openpyxl" in u for u in visit.requests)


def test_only_the_report_is_refused_like_the_cli(visit, tmp_path):
    report = copy(FC_REPORT, tmp_path)
    code, output = cli(report)
    assert code == 1

    choose(visit, report)
    visit.page.click("#convert")

    assert refusal_on_page(visit) == cli_message(output)
    assert "Manca il file dei dati statici" in cli_message(output)


def test_malformed_static_data_is_refused_like_the_cli(visit, tmp_path):
    report = copy(FC_REPORT, tmp_path)
    static = copy(STATIC_DATA, tmp_path, "dati_difettosi.xlsx")
    workbook = openpyxl.load_workbook(static)
    workbook["Millesimali"]["C5"] = None  # a blank millesimi cell
    workbook.save(static)
    code, output = cli(report, static)
    assert code == 1

    choose(visit, report, static)

    message = refusal_on_page(visit)
    assert message == cli_message(output)
    assert message.startswith("Errore nel file dati statici: ")
    assert "cella C5" in message


def test_a_file_that_is_not_a_report_is_refused_like_the_cli(visit, tmp_path):
    fake = tmp_path / "FC_report_finto.xls"
    fake.write_text("questo non è un report", encoding="utf-8")
    static = copy(STATIC_DATA, tmp_path)
    code, output = cli(fake, static)
    assert code == 1

    choose(visit, fake, static)

    message = refusal_on_page(visit)
    assert message == cli_message(output)
    assert message.startswith("Errore nel report FC: ")


def test_choosing_the_second_file_later_converts(visit):
    choose(visit, FC_REPORT)
    expect(visit.page.locator("#outcome .message")).to_have_count(0)  # waits for the other file
    visit.page.set_input_files("#files", [str(STATIC_DATA)])
    expect(visit.page.locator("#download")).to_be_visible()
    assert visit.page.locator("#download").get_attribute("download") == OUTPUT_NAME


def test_repicking_the_static_file_converts_the_new_one(visit, tmp_path):
    report = copy(FC_REPORT, tmp_path)
    broken = copy(STATIC_DATA, tmp_path, "dati_difettosi.xlsx")
    workbook = openpyxl.load_workbook(broken)
    workbook["Millesimali"]["C5"] = None
    workbook.save(broken)

    choose(visit, report, broken)
    assert refusal_on_page(visit).startswith("Errore nel file dati statici: ")
    visit.page.set_input_files("#files", [str(STATIC_DATA)])

    expect(visit.page.locator("#download")).to_be_visible()
    names = [li.inner_text() for li in visit.page.locator("#chosen li").all()]
    assert names == [report.name, STATIC_DATA.name]
