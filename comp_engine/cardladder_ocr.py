from __future__ import annotations

import base64
import json
import os
import re
import time
from pathlib import Path

try:
    from dotenv import load_dotenv
except Exception:
    load_dotenv = None

try:
    from google import genai
    from google.genai import types
except Exception:
    genai = None
    types = None


ROOT = Path(__file__).resolve().parent.parent
DEBUG_DIR = ROOT / "work" / "cardladder-ocr"
PROJECT_ENV = ROOT / ".env"
PHOTO_TOOL_ENV = ROOT / "photo_tool" / "app" / ".env"


def extract_cl_value_from_data_url(data_url: str) -> dict:
    if genai is None or types is None:
        return {"ok": False, "value": None, "error": "google-genai is not available"}

    load_env()
    api_key = os.environ.get("GOOGLE_API_KEY", "").strip()
    if not api_key:
        return {"ok": False, "value": None, "error": "GOOGLE_API_KEY is not configured"}

    mime_type, image_bytes = parse_data_url(data_url)
    debug_id = f"{time.strftime('%Y%m%d-%H%M%S')}-{time.time_ns() % 1_000_000_000:09d}"
    DEBUG_DIR.mkdir(parents=True, exist_ok=True)
    image_path = DEBUG_DIR / f"cardladder-{debug_id}.png"
    image_path.write_bytes(image_bytes)
    client = genai.Client(api_key=api_key)
    prompt = (
        "Read this Card Ladder Sales History screenshot. Return only JSON with keys: "
        "value, label_seen, profile_title, profile_grader, profile_grade, evidence, comps. "
        "Find the Card Ladder value shown near the label 'CL Value'. Only set value if the dollar amount "
        "is visually adjacent to the CL Value label. If the label is not visible or ambiguous, return value null. "
        "Read the full filter line above CL Value. If it contains text like "
        "'Grade: 10, Grader: PSA, Profile: 2020 Panini Contenders Optic 120 Tyrese Maxey Autograph-Blue (Pop 24)', "
        "return profile_grade as 10, profile_grader as PSA, and profile_title as only the profile text without the "
        "'Profile:' label and without the '(Pop ...)' text. "
        "For comps, extract the most recent visible sales rows in page order, up to 5. For each comp return "
        "source, title, date_sold, sale_type, and price. Only use text visible in the screenshot."
    )

    last_error = ""
    response = None
    for attempt in range(3):
        try:
            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=[
                    types.Part.from_bytes(data=image_bytes, mime_type=mime_type),
                    prompt,
                ],
            )
            break
        except Exception as error:
            last_error = str(error)
            time.sleep(1.5 * (attempt + 1))

    if response is None:
        result = {
            "ok": False,
            "value": None,
            "labelSeen": False,
            "error": f"OCR request failed: {last_error}",
            "debugImage": str(image_path),
        }
        (DEBUG_DIR / f"cardladder-{debug_id}.json").write_text(json.dumps(result, indent=2), encoding="utf-8")
        return result
    raw = response.text or ""
    parsed = parse_jsonish(raw)
    value = parse_money(parsed.get("value"))
    label_seen = parse_bool(parsed.get("label_seen"))
    comps = normalize_comps(parsed.get("comps"))
    profile_title = clean_profile_title(parsed.get("profile_title") or parsed.get("profile") or parsed.get("card_title"))
    profile_grader = clean_grader(parsed.get("profile_grader") or parsed.get("grader"))
    profile_grade = clean_grade(parsed.get("profile_grade") or parsed.get("grade"))
    result = {
        "ok": value is not None and label_seen,
        "value": value if label_seen else None,
        "labelSeen": label_seen,
        "profileTitle": profile_title,
        "profileGrader": profile_grader,
        "profileGrade": profile_grade,
        "comps": comps,
        "evidence": str(parsed.get("evidence") or raw)[:500],
        "raw": raw[:1000],
        "debugImage": str(image_path),
    }
    (DEBUG_DIR / f"cardladder-{debug_id}.json").write_text(json.dumps(result, indent=2), encoding="utf-8")
    return result


def load_env() -> None:
    if load_dotenv:
        load_dotenv(PROJECT_ENV, override=False)
        load_dotenv(PHOTO_TOOL_ENV, override=False)


def parse_data_url(data_url: str) -> tuple[str, bytes]:
    match = re.match(r"^data:(.*?);base64,(.*)$", data_url or "", re.S)
    if not match:
        raise ValueError("Expected screenshot data URL")
    return match.group(1) or "image/png", base64.b64decode(match.group(2))


def parse_jsonish(raw: str) -> dict:
    text = str(raw or "").strip()
    text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text, flags=re.I | re.S).strip()
    try:
        return repair_parsed_jsonish(json.loads(text), text)
    except Exception:
        match = re.search(r"\{.*\}", text, re.S)
        if match:
            try:
                return repair_parsed_jsonish(json.loads(match.group(0)), text)
            except Exception:
                pass
    return salvage_jsonish(text)


def repair_parsed_jsonish(parsed: dict, text: str) -> dict:
    if not isinstance(parsed, dict):
        return salvage_jsonish(text)
    parsed_comps = parsed.get("comps") if isinstance(parsed.get("comps"), list) else []
    source_count = len(re.findall(r'"source"\s*:', text, flags=re.I))
    if source_count > len(parsed_comps):
        salvaged = salvage_comps(text)
        if len(salvaged) > len(parsed_comps):
            parsed = dict(parsed)
            parsed["comps"] = salvaged
    return parsed


def salvage_jsonish(text: str) -> dict:
    return {
        "value": first_jsonish_field(text, "value"),
        "label_seen": first_jsonish_field(text, "label_seen"),
        "profile_title": first_jsonish_field(text, "profile_title"),
        "profile_grader": first_jsonish_field(text, "profile_grader"),
        "profile_grade": first_jsonish_field(text, "profile_grade"),
        "evidence": first_jsonish_field(text, "evidence") or text,
        "comps": salvage_comps(text),
    }


def first_jsonish_field(text: str, key: str):
    pattern = rf'"{re.escape(key)}"\s*:\s*(true|false|null|"(?:[^"\\]|\\.)*"|[-+]?\d+(?:\.\d+)?)'
    match = re.search(pattern, text, flags=re.I | re.S)
    if not match:
        return None
    value = match.group(1).strip()
    if value.lower() == "true":
        return True
    if value.lower() == "false":
        return False
    if value.lower() == "null":
        return None
    if value.startswith('"') and value.endswith('"'):
        return value[1:-1].replace('\\"', '"')
    return value


def salvage_comps(text: str) -> list[dict]:
    comps: list[dict] = []
    for source_match in re.finditer(r'"source"\s*:', text, flags=re.I):
        block = text[source_match.start():]
        next_source = re.search(r',\s*"source"\s*:', block[10:], flags=re.I)
        if next_source:
            block = block[: 10 + next_source.start()]
        comp = {
            "source": first_jsonish_field(block, "source"),
            "title": first_jsonish_field(block, "title"),
            "date_sold": first_jsonish_field(block, "date_sold") or first_jsonish_field(block, "date"),
            "sale_type": first_jsonish_field(block, "sale_type") or first_jsonish_field(block, "type"),
            "price": first_jsonish_field(block, "price"),
        }
        if comp["source"] and comp["title"] and comp["date_sold"] and comp["price"]:
            comps.append(comp)
        if len(comps) >= 5:
            break
    return comps


def parse_money(value) -> float | None:
    if value is None or value == "":
        return None
    match = re.search(r"[\d,]+(?:\.\d{1,2})?", str(value))
    if not match:
        return None
    try:
        return float(match.group(0).replace(",", ""))
    except ValueError:
        return None


def parse_bool(value) -> bool:
    if isinstance(value, bool):
        return value
    text = str(value or "").strip().lower()
    return text in {"true", "yes", "y", "1", "label seen", "cl value"}


def normalize_comps(value) -> list[dict]:
    if not isinstance(value, list):
        return []
    comps: list[dict] = []
    for item in value[:5]:
        if not isinstance(item, dict):
            continue
        comps.append(
            {
                "source": clean_text(item.get("source")),
                "title": clean_text(item.get("title")),
                "date_sold": clean_text(item.get("date_sold") or item.get("date")),
                "sale_type": clean_text(item.get("sale_type") or item.get("type")),
                "price": clean_text(item.get("price")),
            }
        )
    return comps


def clean_text(value) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def clean_profile_title(value) -> str:
    text = clean_text(value)
    text = re.sub(r"^profile\s*:\s*", "", text, flags=re.I)
    text = re.sub(r"\s*\(pop\s*[^)]*\)\s*$", "", text, flags=re.I)
    return clean_text(text)


def clean_grader(value) -> str:
    text = clean_text(value).upper()
    aliases = {"BECKETT": "BGS", "BVG": "BGS", "PSA": "PSA", "BGS": "BGS", "SGC": "SGC", "CGC": "CGC"}
    return aliases.get(text, text)


def clean_grade(value) -> str:
    text = clean_text(value)
    matches = re.findall(r"\d+(?:\.\d+)?", text)
    return matches[-1] if matches else ""
