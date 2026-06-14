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
SOURCE_URL = "https://f1frogblog.wordpress.com/2022/06/03/f1-all-time-driver-rankings-1950-2022/"
SOURCE_NAME = "F1 Frog all-time driver rankings"
CURRENT_SOURCE_NAME = "User-provided current F1 drivers"
CURRENT_DRIVERS = {
    "Pierre Gasly",
    "Franco Colapinto",
    "Fernando Alonso",
    "Lance Stroll",
    "Nico Hülkenberg",
    "Gabriel Bortoleto",
    "Sergio Pérez",
    "Valtteri Bottas",
    "Charles Leclerc",
    "Lewis Hamilton",
    "Esteban Ocon",
    "Oliver Bearman",
    "Lando Norris",
    "Oscar Piastri",
    "George Russell",
    "Kimi Antonelli",
    "Liam Lawson",
    "Arvid Lindblad",
    "Max Verstappen",
    "Isack Hadjar",
    "Carlos Sainz",
    "Alex Albon",
}


def main() -> None:
    data = load_data()
    source_names = fetch_f1_names()
    names = sorted(set(source_names) | CURRENT_DRIVERS, key=str.lower)
    players: dict[str, dict[str, Any]] = data.setdefault("players", {})
    for name in names:
        key = player_key(name)
        if key:
            players[key] = {"sport": "f1", "displayName": name}

    sources = [
        source
        for source in data.get("sources", [])
        if str(source.get("name") or "") not in {SOURCE_NAME, CURRENT_SOURCE_NAME}
    ]
    sources.append({
        "name": SOURCE_NAME,
        "sport": "f1",
        "url": SOURCE_URL,
        "count": len(source_names),
    })
    sources.append({
        "name": CURRENT_SOURCE_NAME,
        "sport": "f1",
        "url": "user message in Codex thread",
        "count": len(CURRENT_DRIVERS),
    })
    data["sources"] = sources
    data["generatedAt"] = datetime.now(timezone.utc).isoformat(timespec="milliseconds").replace("+00:00", "Z")
    write_data(data)
    print(f"Imported {len(source_names)} F1 Frog driver names plus {len(CURRENT_DRIVERS)} current drivers")


def fetch_f1_names() -> set[str]:
    raw = fetch_text(SOURCE_URL)
    names = set(extract_ranked_names(raw))
    names.update(extract_considered_names(raw))
    return {name for name in names if useful_name(name)}


def extract_ranked_names(raw: str) -> set[str]:
    text = html.unescape(raw)
    names: set[str] = set()
    for match in re.finditer(r">\s*(\d{1,3})\s*[–-]\s*([^<]+?)\s*<", text, re.I):
        rank = int(match.group(1))
        if 1 <= rank <= 100:
            name = clean_label(match.group(2))
            if useful_name(name):
                names.add(name)
    return names


def extract_considered_names(raw: str) -> set[str]:
    names: set[str] = set()
    marker = "Drivers considered for the list but not quite making the cut include"
    start = raw.find(marker)
    end = raw.find("Sources for the information", start)
    if start < 0 or end < 0:
        return names
    section = raw[start:end]
    paragraphs = re.findall(r"<p[^>]*>(.*?)</p>", section, re.I | re.S)
    for paragraph in paragraphs:
        text = clean_label(paragraph)
        if not text or text.startswith("Drivers considered"):
            continue
        first_sentence = re.split(r"\.\s+", text, maxsplit=1)[0]
        for value in re.split(r"\s*,\s*", first_sentence):
            name = clean_label(value)
            if useful_name(name):
                names.add(name)
    return names


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
    text = normalize_text(text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_text(value: str) -> str:
    return (
        str(value or "")
        .replace("“", '"')
        .replace("”", '"')
        .replace("‘", "'")
        .replace("’", "'")
        .replace("\xa0", " ")
    )


def useful_name(value: str) -> bool:
    text = str(value or "").strip()
    lowered = text.lower()
    if len(text) < 2 or len(text) > 90:
        return False
    blocked = {"thank you", "share this", "like loading", "related"}
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
