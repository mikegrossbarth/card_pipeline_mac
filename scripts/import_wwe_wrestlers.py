from __future__ import annotations

import html
import json
import re
import subprocess
import unicodedata
import urllib.request
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
DATA_PATH = ROOT / "assignment_player_sport_data.js"
SOURCE_URL = "https://www.whenitwascool.com/201-greatest-pro-wrestlers-of-all-time-list-pro-wrestlings-greatest-ever"
SOURCE_NAME = "When It Was Cool 201 greatest pro wrestlers"


def main() -> None:
    data = load_data()
    source_names = fetch_wrestler_names()
    import_names = expand_aliases(source_names)
    players: dict[str, dict[str, Any]] = data.setdefault("players", {})
    for name in sorted(import_names, key=str.lower):
        key = player_key(name)
        if key:
            players[key] = {"sport": "wwe", "displayName": name}

    sources = [
        source
        for source in data.get("sources", [])
        if str(source.get("name") or "") != SOURCE_NAME
    ]
    sources.append({
        "name": SOURCE_NAME,
        "sport": "wwe",
        "url": SOURCE_URL,
        "count": len(source_names),
    })
    data["sources"] = sources
    data["generatedAt"] = datetime.now(timezone.utc).isoformat(timespec="milliseconds").replace("+00:00", "Z")
    write_data(data)
    print(f"Imported {len(source_names)} WWE wrestler names and {len(import_names)} match aliases from {SOURCE_URL}")


def fetch_wrestler_names() -> list[str]:
    raw = fetch_text(SOURCE_URL)
    names: list[str] = []
    seen: set[str] = set()
    pattern = re.compile(r"<strong>\s*(\d{1,3})\s*:\s*(.*?)</strong>", re.I | re.S)
    for match in pattern.finditer(raw):
        rank = int(match.group(1))
        if rank < 1 or rank > 201:
            continue
        name = clean_label(match.group(2))
        if useful_name(name) and name.lower() not in seen:
            names.append(name)
            seen.add(name.lower())
    return names


def expand_aliases(names: list[str]) -> set[str]:
    aliases: set[str] = set()
    for name in names:
        aliases.add(name)
        for alias in alias_candidates(name):
            if useful_name(alias):
                aliases.add(alias)
    return aliases


def alias_candidates(name: str) -> set[str]:
    candidates: set[str] = set()
    text = normalize_quotes(name)

    without_quotes = re.sub(r'"[^"]+"', " ", text)
    without_quotes = re.sub(r"\s+", " ", without_quotes).strip()
    if without_quotes and without_quotes != text:
        candidates.add(without_quotes)

    without_parens = re.sub(r"\([^)]*\)", " ", without_quotes or text)
    without_parens = re.sub(r"\s+", " ", without_parens).strip()
    if without_parens:
        candidates.add(without_parens)

    paren_matches = re.findall(r"\(([^)]{2,60})\)", text)
    for value in paren_matches:
        candidate = re.sub(r"\s+", " ", value).strip()
        if candidate:
            candidates.add(candidate)

    quote_matches = re.findall(r'"([^"]{2,60})"', text)
    for value in quote_matches:
        candidate = re.sub(r"\s+", " ", value).strip()
        if candidate:
            candidates.add(candidate)

    return candidates


def fetch_text(url: str) -> str:
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/125.0 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.9",
        },
    )
    try:
        with urllib.request.urlopen(request, timeout=45) as response:
            return response.read().decode("utf-8", "ignore")
    except Exception:
        result = subprocess.run(
            [
                "curl.exe",
                "-L",
                "-A",
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
                url,
            ],
            check=True,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
        )
        return result.stdout


def clean_label(value: str) -> str:
    text = re.sub(r"<[^>]+>", " ", value)
    text = html.unescape(text)
    text = normalize_quotes(text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_quotes(value: str) -> str:
    return (
        str(value or "")
        .replace("“", '"')
        .replace("”", '"')
        .replace("‘", "'")
        .replace("’", "'")
    )


def useful_name(value: str) -> bool:
    text = str(value or "").strip()
    lowered = text.lower()
    if len(text) < 2 or len(text) > 90:
        return False
    blocked = {"yes", "no", "sources used"}
    return lowered not in blocked and not lowered.isdigit()


def load_data() -> dict[str, Any]:
    raw = DATA_PATH.read_text(encoding="utf-8")
    match = re.search(r"window\.AutoSheetReviewPlayerSports\s*=\s*(\{.*?\})\s*;\s*\}\)\(\);", raw, re.S)
    if not match:
        raise RuntimeError(f"Could not parse {DATA_PATH}")
    return json.loads(match.group(1))


def write_data(data: dict[str, Any]) -> None:
    payload = json.dumps(data, indent=2, ensure_ascii=False, sort_keys=True)
    DATA_PATH.write_text(
        "// Generated by category import scripts. Do not edit by hand.\n"
        "(function exposePlayerSportData() {\n"
        f"  window.AutoSheetReviewPlayerSports = {payload};\n"
        "})();\n",
        encoding="utf-8",
    )


def player_key(value: str) -> str:
    text = strip_accents(value).lower()
    text = re.sub(r"[^a-z0-9.' -]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def strip_accents(value: str) -> str:
    return "".join(
        char
        for char in unicodedata.normalize("NFKD", str(value or ""))
        if not unicodedata.combining(char)
    )


if __name__ == "__main__":
    main()
