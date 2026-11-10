import json
import subprocess
import sys
from pathlib import Path

from pptx import Presentation

from agent.scope import in_scope
from agent.specs import normalize_field
from agent.units import to_si

ROOT = Path(__file__).resolve().parents[1]


def test_units():
    assert round(to_si(211.9, "in")[0]) == 5382
    assert round(to_si(420, "hp")[0]) == 313
    assert round(to_si(460, "lb-ft")[0]) == 624
    assert round(to_si(17, "mpg")[0], 1) == 13.8
    assert to_si(None, "in") == (None, None)


def test_provenance_rules():
    assert normalize_field("power", {"value": 400, "unit": "hp"})["value"] is None
    assert normalize_field("power", {"value": 400, "unit": "hp", "tier": 4, "source_url": "a"})["value"] is None
    ok = normalize_field("power", {"value": 400, "unit": "hp", "tier": 4, "source_urls": ["a", "b"]})
    assert ok["display"] == "298 kW"


def test_scope():
    assert in_scope({"maker": "Genesis", "body": "Sedan"}, 4900)[0]
    assert in_scope({"maker": "X", "platform": "BOF", "body": "SUV"}, 5382)[0]
    assert not in_scope({"maker": "X", "platform": "BOF", "body": "SUV"}, 6096)[0]
    assert not in_scope({"maker": "X", "platform": "Monocoque", "body": "SUV"}, 5000)[0]


def test_build(tmp_path):
    r = subprocess.run([sys.executable, "run_daily.py", "tests/fixtures/sample_brief.json",
                        "--specs", "/nonexistent", "--out", str(tmp_path)],
                       cwd=ROOT, capture_output=True, text=True)
    assert r.returncode == 0, r.stderr
    out = tmp_path / "2099-01-01" / "BOF_Genesis_Brief_20990101.pptx"
    prs = Presentation(out)
    assert len(prs.slides) == 9


def test_out_of_scope_rival_dropped(tmp_path):
    r = subprocess.run([sys.executable, "run_daily.py", "tests/fixtures/sample_brief.json",
                        "--specs", "/nonexistent", "--out", str(tmp_path)],
                       cwd=ROOT, capture_output=True, text=True)
    assert "'r2' 대상 범위 밖" in r.stdout
