"""Progress lines in the command window."""
from __future__ import annotations

import sys
import threading
from collections import Counter
from pathlib import Path

from .models import Competitor, Item, Label, PriceResult


class Progress:
    def __init__(self, total: int, out=None) -> None:
        self.total, self.done = total, 0
        self.out = out if out is not None else sys.stdout
        self._lock = threading.Lock()

    def say(self, message: str = "") -> None:
        """Print one line; characters the console cannot show (such as £) are replaced."""
        encoding = getattr(self.out, "encoding", None) or "utf-8"
        try:
            message.encode(encoding)
        except (UnicodeEncodeError, LookupError):
            message = message.replace("£", "GBP ").encode("ascii", "replace").decode("ascii")
        with self._lock:
            self.out.write(message + "\n")
            try:
                self.out.flush()
            except Exception:
                pass

    def line(self, item: Item, comp: Competitor, result: PriceResult) -> None:
        with self._lock:
            self.done += 1
            n = self.done
        width = len(str(self.total))
        detail = f"£{result.price_ex:.2f} ex VAT" if result.price_ex is not None else result.remark
        self.say(f"{n:>{width}}/{self.total}  {item.code:<16}  {comp.company:<24}  "
                 f"{result.label.value:<12}  {detail}".rstrip())

    def summary(self, results: list[PriceResult], issues: int, issues_path: Path) -> None:
        counts = Counter(r.label for r in results)
        found = sum(counts[l] for l in (Label.CLIENT_LINK, Label.VERIFIED, Label.MANUAL))
        self.say("")
        self.say(f"Finished: {len(results)} of {self.total} lookups.")
        self.say(f"  Prices found:      {found}")
        self.say(f"  Unverified:        {counts[Label.UNVERIFIED]}")
        self.say(f"  Not found:         {counts[Label.NOT_FOUND]}")
        self.say(f"  Blocked:           {counts[Label.BLOCKED]}")
        self.say(f"  Errors:            {counts[Label.ERROR]}")
        self.say(f"  Data issues:       {issues}" + ("" if issues else "  (none - nothing to correct)"))
        self.say(f"  Issues file:       {issues_path}")
