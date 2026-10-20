"""Run: python -m pytest tests  (logic only; Excel I/O needs a real Excel run)."""
import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1] / "src"))
from sheets import plan_unhide  # noqa: E402


def test_unhides_hidden_and_very_hidden():
    sheets = [
        {"name": "Summary", "visibility": "Visible"},
        {"name": "Sales", "visibility": "Hidden"},
        {"name": "Secrets", "visibility": "VeryHidden"},
    ]
    assert plan_unhide(sheets) == ["Sales", "Secrets"]


def test_nothing_hidden():
    assert plan_unhide([{"name": "A", "visibility": "Visible"}]) == []
