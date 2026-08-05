# -*- coding: utf-8 -*-
"""Fetch Japanese national holidays (内閣府由来) for company calendar initialization."""

from __future__ import annotations

import json
import logging
import urllib.error
import urllib.request
from datetime import date

from planning_core.core.attendance_paths import japanese_holidays_cache_path

JP_CALENDAR_API_YEAR = "https://api.jp-calendar.com/v1/holidays/{year}.json"
FETCH_TIMEOUT_SEC = 30

logger = logging.getLogger(__name__)


def fetch_national_holidays_for_year(year: int, force_online: bool = False) -> list[dict]:
    """
    Returns [{ "date": "YYYY-MM-DD", "name": "元日" }, ...].
    Uses cache when available unless force_online; tries API then cache on failure.
    """
    cached = _load_cache(year)
    if cached and not force_online:
        return cached
    online = _fetch_online(year)
    if online:
        _write_cache(year, online)
        return online
    if cached:
        logger.warning(
            "祝日 API 取得失敗。キャッシュ %s を使用します。",
            japanese_holidays_cache_path(year),
        )
        return cached
    raise RuntimeError(
        f"{year} 年の国民の祝日を取得できません（ネットワークとキャッシュともに失敗）"
    )


def _fetch_online(year: int) -> list[dict] | None:
    url = JP_CALENDAR_API_YEAR.format(year=year)
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "pm-ai-desktop/1.0"})
        with urllib.request.urlopen(req, timeout=FETCH_TIMEOUT_SEC) as resp:
            raw = resp.read().decode("utf-8")
        payload = json.loads(raw)
        if isinstance(payload, dict):
            out: list[dict] = []
            for k, name in sorted(payload.items()):
                if isinstance(name, str):
                    out.append({"date": k, "name": name})
            return out
    except (urllib.error.URLError, TimeoutError, json.JSONDecodeError, OSError) as e:
        logger.warning("祝日 API 取得失敗 (%s): %s", url, e)
    return None


def _load_cache(year: int) -> list[dict] | None:
    path = japanese_holidays_cache_path(year)
    if not path.is_file():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        if isinstance(data, list):
            return data
    except (OSError, json.JSONDecodeError):
        return None
    return None


def _write_cache(year: int, holidays: list[dict]) -> None:
    path = japanese_holidays_cache_path(year)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(holidays, ensure_ascii=False, indent=2), encoding="utf-8")
