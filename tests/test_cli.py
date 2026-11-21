"""CLI: запуск из чужого рабочего каталога, `-o папка` для одного входа, команды из README."""
from __future__ import annotations

import os
import subprocess
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEMO = os.path.join(ROOT, "examples", "demo_raw.xlsx")

pytestmark = pytest.mark.skipif(not os.path.exists(DEMO), reason="нет examples/demo_raw.xlsx")


def run(args, cwd):
    # UTF-8 для stdout дочернего процесса: на Windows (CI) пайп иначе в cp1252,
    # и печать кириллических имён файлов упадёт с UnicodeEncodeError
    env = {**os.environ, "PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8"}
    return subprocess.run([sys.executable, *args], cwd=cwd, capture_output=True,
                          text=True, encoding="utf-8", env=env, timeout=120)


def test_cli_from_other_cwd(tmp_path):
    r = run([os.path.join(ROOT, "smeta_render.py"), DEMO, "-o", "out.xlsx"], cwd=tmp_path)
    assert r.returncode == 0, r.stderr
    assert (tmp_path / "out.xlsx").exists()


@pytest.mark.parametrize("out", ["result/", "existing"])
def test_cli_single_input_output_dir(tmp_path, out):
    if out == "existing":
        (tmp_path / out).mkdir()
    r = run([os.path.join(ROOT, "smeta_render.py"), DEMO, "-o", out], cwd=tmp_path)
    assert r.returncode == 0, r.stderr
    files = os.listdir(tmp_path / out.rstrip("/"))
    assert len(files) == 1 and files[0].startswith("Сводник") and files[0].endswith(".xlsx")


def test_parser_cli_defaults_to_demo(tmp_path):
    r = run([os.path.join(ROOT, "smeta_parser.py")], cwd=tmp_path)
    assert r.returncode == 0, r.stderr
    assert "Распарсено смет: 3" in r.stdout
