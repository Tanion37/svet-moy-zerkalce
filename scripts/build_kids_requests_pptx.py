"""Deprecated: kids-only requests PnP.

Kids marker `*` removed from canon (2026-08-25). Kept as stub so old
commands fail loudly with a pointer to archive.
"""
from __future__ import annotations


def build():
    raise SystemExit(
        "Kids PnP disabled: marker `*` removed. "
        "See archive/rejected-request-parts.md"
    )


if __name__ == "__main__":
    build()
