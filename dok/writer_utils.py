"""
dok.writer_utils
~~~~~~~~~~~~~~~~
Shared utilities for docx_writer and html_writer.
"""

from __future__ import annotations
from .models import RunModel


def group_runs_by_hyperlink(
    runs: list[RunModel],
) -> list[tuple[str | None, list[RunModel]]]:
    """Group consecutive runs by hyperlink_url.

    Returns a list of (url_or_none, [runs]) tuples.
    Consecutive runs with the same non-None hyperlink_url are grouped together.
    Runs without a hyperlink get url=None and are grouped individually.
    """
    groups: list[tuple[str | None, list[RunModel]]] = []
    i = 0
    while i < len(runs):
        run = runs[i]
        if run.hyperlink_url:
            url = run.hyperlink_url
            group: list[RunModel] = []
            while i < len(runs) and runs[i].hyperlink_url == url:
                group.append(runs[i])
                i += 1
            groups.append((url, group))
        else:
            groups.append((None, [run]))
            i += 1
    return groups
