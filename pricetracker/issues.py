"""Collects input data issues and writes the issues file."""
from __future__ import annotations

import threading
from pathlib import Path

from .models import ERROR, Issue

ISSUE_HEADERS = ["Kind", "Input file", "Row", "Item or competitor", "What is wrong",
                 "Value in the client's data", "Value seen", "What to change"]
NO_ISSUES = "No issues found"


class IssueLog:
    def __init__(self) -> None:
        self.issues: list[Issue] = []
        self._seen: set[tuple] = set()
        self._lock = threading.Lock()

    def add(self, kind, file, row, subject, problem, client_value="", seen_value="", fix="") -> None:
        issue = Issue(kind, file, row, str(subject), problem, str(client_value or ""),
                      str(seen_value or ""), fix)
        key = (issue.kind, issue.file, issue.row, issue.subject, issue.problem, issue.seen_value)
        with self._lock:
            if key not in self._seen:
                self._seen.add(key)
                self.issues.append(issue)

    def __len__(self) -> int:
        return len(self.issues)

    def count(self, kind: str) -> int:
        return sum(1 for i in self.issues if i.kind == kind)

    def rows(self) -> list[list]:
        ordered = sorted(self.issues, key=lambda i: (i.kind != ERROR, i.file, i.row or 0))
        return [[i.kind, i.file, i.row, i.subject, i.problem, i.client_value or None,
                 i.seen_value or None, i.fix or None] for i in ordered]

    def write(self, path: Path) -> Path:
        from .excel_io import write_table
        rows = self.rows() or [[NO_ISSUES]]
        return write_table(Path(path), "Data issues", ISSUE_HEADERS, rows)
