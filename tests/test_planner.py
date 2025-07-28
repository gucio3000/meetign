import os
import sys
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parents[1]))
from meeting.v1 import meeting_planner_template as mpt


def test_main_creates_files(tmp_path, monkeypatch):
    with monkeypatch.context() as m:
        m.chdir(tmp_path)
        mpt.main()
    expected = [
        "meeting_planner_template.xlsx",
        "directory_template.csv",
        "roster_template.csv",
        "meeting_invite.ics",
        "decision_log.csv",
    ]
    for name in expected:
        assert (tmp_path / name).exists()
