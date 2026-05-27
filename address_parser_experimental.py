from __future__ import annotations

import json
import os
import re
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import pandas as pd
import requests


ROOT = Path(__file__).resolve().parent

RAW_COLUMN_ALIASES = {
    "order_id": ["order id", "order_id", "orderid"],
    "client_id": ["client id", "client_id", "clientid"],
    "site_id": ["site id", "site_id", "siteid"],
    "site_name": ["site name", "site_name", "sitename"],
    "ship_to_city": ["ship to city", "ship_to_city", "ship to_city", "city", "עיר", "יישוב"],
    "ship_to_street1": [
        "ship to street 1",
        "ship_to_street1",
        "ship to street1",
        "address line 1",
        "line1",
        "address1",
        "כתובת 1",
        "כתובת1",
        "Street_Name",
        "street name",
    ],
    "ship_to_street2": [
        "ship to street 2",
        "ship_to_street2",
        "ship to street2",
        "address line 2",
        "line2",
        "address2",
        "כתובת 2",
        "כתובת2",
        "House_Number",
        "house number",
    ],
    "required_delivery_date": ["required delivery date", "required_delivery_date", "delivery date"],
    "comments": ["comments", "comment", "notes"],
}

CITY_FALLBACK_INDEX = 3
STREET1_FALLBACK_INDEX = 4
STREET2_FALLBACK_INDEX = 5

TYPO_FIXES = {
    "רבי עקיבה": "רבי עקיבא",
    "ר עקיבה": "רבי עקיבא",
    "ר עקיבא": "רבי עקיבא",
    "חזוניש": "חזון איש",
    "חזו\"א": "חזון איש",
    "חזו׳א": "חזון איש",
    "machec chochma": "משך חכמה",
    "רשבי": "רבי שמעון בר יוחאי",
    "רשב\"י": "רבי שמעון בר יוחאי",
    "רשב\"\"י": "רבי שמעון בר יוחאי",
    "רשב׳י": "רבי שמעון בר יוחאי",
    "תל אביביפו": "תל אביב יפו",
    "מודיעין עלית": "מודיעין עילית",
}

STREET_PREFIX_PATTERN = re.compile(r"^(?:רחוב|רח[׳'\"]?|שדרות|שד[׳'\"]?)\s+")
DATE_TIME_PATTERN = re.compile(
    r"(\b\d{1,2}[:.]\d{2}(?:[:.]\d{2})?\b|"
    r"\b\d{4}[-/.]\d{1,2}[-/.]\d{1,2}(?:\s+\d{1,2}[:.]\d{2}(?::\d{2})?)?\b|"
    r"\b\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}\b|"
    r"\b(?:am|pm|שעה|תאריך|יום)\b)",
    flags=re.IGNORECASE,
)
PHONE_PATTERN = re.compile(r"(?:\+?972|0)(?:[-\s]?\d){8,9}")

CITY_GEOCODE_ALIASES = {
    "אחיסמך": ["ahisamakh", "achisamach", "ahisamach"],
    "אלישמע": ["elishama"],
    "אלעד": ["elad"],
    "אלפי מנשה": ["alfei menashe"],
    "אשדוד": ["ashdod"],
    "אשתאול": ["eshtaol"],
    "באר יעקב": ["beer yaakov", "be'er ya'akov", "be'er yaakov"],
    "בית מאיר": ["beit meir", "beit me'ir"],
    "בית שמש": ["beit shemesh", "bet shemesh"],
    "ביתר עילית": ["beitar illit", "betar illit", "beitar ilit"],
    "בת ים": ["bat yam"],
    "דולב": ["dolev"],
    "חולון": ["holon"],
    "טבריה": ["tiberias", "tverya"],
    "יד בנימין": ["yad binyamin"],
    "ירושלים": ["jerusalem"],
    "כרמיאל": ["karmiel", "carmiel"],
    "לוד": ["lod"],
    "מבוא חורון": ["mevo horon", "mevo choron"],
    "מודיעין": ["modiin", "modi'in", "modi'in-maccabim-re'ut", "modiin-maccabim-reut"],
    "מודיעין עילית": ["modi'in illit", "modiin illit", "modi'in ilit", "modiin ilit", "modi in illit"],
    "נריה": ["nerya", "neriya"],
    "נתניה": ["netanya"],
    "עלי": ["eli"],
    "עפולה": ["afula"],
    "פתח תקווה": ["petah tikva", "petach tikva"],
    "קריית אתא": ["kiryat ata", "qiryat ata"],
    "קריית גת": ["kiryat gat", "qiryat gat"],
    "קריית טבעון": ["kiryat tivon", "qiryat tivon", "kiryat tiv'on"],
    "קריית ספר": ["kiryat sefer", "qiryat sefer"],
    "רכסים": ["rekhasim", "rechasim", "rakhassim"],
    "רמלה": ["ramla", "ramle"],
    "שילה": ["shilo", "shiloh"],
}

STREET_GEOCODE_ALIASES = {
    "רבי שמעון בר יוחאי": ["רשב\"י", "רשב״י", "Rashbi", "Rabbi Shimon Bar Yochai", "Rabi Shimon Bar Yochai"],
    "הרב מפונוביז": ["הרב מפוניבז", "שדרות הרב מפוניבז", "Harav Miponovezh", "Ponevezh Rav"],
    "הרב מפוניבז": ["הרב מפונוביז", "שדרות הרב מפוניבז", "Harav Miponovezh", "Ponevezh Rav"],
    "אהרונסון": ["Aharonson", "Aaronson"],
    "יחד שבטי ישראל": ["שדרות יחד שבטי ישראל", "Yahad Shivtei Israel", "Yachad Shivtei Israel"],
    "שדרות יחד שבטי ישראל": ["יחד שבטי ישראל", "Yahad Shivtei Israel", "Yachad Shivtei Israel"],
    "לאה אמנו": ["Lea Imenu", "Leah Imenu"],
    "משה רבינו": ["Moshe Rabbeinu", "Moshe Rabeinu"],
    "הרב אלישיב": ["Harav Elyashiv", "Rav Elyashiv"],
    "בעל שם טוב": ["Baal Shem Tov", "Besht"],
}

PRECISE_GEOCODE_TYPES = {"street_address", "premise", "subpremise"}
PRECISE_LOCATION_TYPES = {"ROOFTOP", "RANGE_INTERPOLATED"}
CITY_COMPONENT_TYPES = {"locality", "postal_town", "administrative_area_level_3", "administrative_area_level_2"}
MAX_GEOCODE_ATTEMPTS_PER_ADDRESS = 8
REFERENCE_METADATA_PATH = ROOT / ".agents" / "skills" / "israeli-address-autocomplete" / "references" / "data" / "metadata.json"


@dataclass
class CleanedAddress:
    city: str
    street: str
    house_number: str
    apartment: str
    floor: str
    entrance: str
    secondary_notes: str
    secondary_classification: str
    needs_review: bool
    merged_address: str
    cleaned_address: str
    status: str
    confidence: str
    parser_output: str = ""

    @property
    def is_valid(self) -> bool:
        return self.confidence in {"high", "medium"} and bool(self.city and self.street and self.house_number)


@dataclass(frozen=True)
class GeocodeAttempt:
    query: str
    language: str
    reason: str


def get_lookup_script() -> str | None:
    script_path = ROOT / ".agents" / "skills" / "israeli-address-autocomplete" / "scripts" / "lookup_address.py"
    return str(script_path) if script_path.exists() else None


SKILL_SCRIPT = get_lookup_script()


def clean_cell(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip()
    if text.lower() in {"nan", "none", "null"}:
        return ""
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return re.sub(r"\s+", " ", text).strip()


def normalize_column_name(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value).strip().lower().replace("_", " "))


def resolve_columns(df: pd.DataFrame) -> dict[str, str]:
    normalized = {normalize_column_name(col): col for col in df.columns}
    resolved: dict[str, str] = {}
    for canonical, aliases in RAW_COLUMN_ALIASES.items():
        for alias in aliases:
            col = normalized.get(normalize_column_name(alias))
            if col is not None:
                resolved[canonical] = col
                break
    if "ship_to_city" not in resolved and "site_name" in resolved:
        resolved["ship_to_city"] = resolved["site_name"]
    if "ship_to_city" not in resolved and "City" in df.columns:
        resolved["ship_to_city"] = "City"
    if "ship_to_city" not in resolved and CITY_FALLBACK_INDEX < len(df.columns):
        resolved["ship_to_city"] = df.columns[CITY_FALLBACK_INDEX]
    if "ship_to_street1" not in resolved and STREET1_FALLBACK_INDEX < len(df.columns):
        resolved["ship_to_street1"] = df.columns[STREET1_FALLBACK_INDEX]
    if "ship_to_street2" not in resolved and STREET2_FALLBACK_INDEX < len(df.columns):
        resolved["ship_to_street2"] = df.columns[STREET2_FALLBACK_INDEX]
    return resolved


def parse_address(address_string: str) -> str | None:
    if not SKILL_SCRIPT:
        return None
    try:
        result = subprocess.run(
            [sys.executable, str(SKILL_SCRIPT), "format", address_string],
            capture_output=True,
            text=True,
            check=True,
            encoding="utf-8",
            errors="replace",
        )
        return result.stdout
    except subprocess.CalledProcessError:
        return None


def parse_lookup_output(text: str) -> dict[str, str | bool | None]:
    parsed: dict[str, str | bool | None] = {"street": None, "number": None, "city": None, "formatted": False}
    for line in text.splitlines():
        line = line.strip()
        if line.startswith("Street:"):
            parsed["street"] = line.split(":", 1)[1].strip()
        elif line.startswith("Number:"):
            parsed["number"] = line.split(":", 1)[1].strip()
        elif line.startswith("City:"):
            parsed["city"] = line.split(":", 1)[1].strip()
        elif line.startswith("Formatted address:"):
            parsed["formatted"] = True
    if isinstance(parsed["city"], str) and parsed["city"].lower() == "not identified":
        parsed["city"] = None
    if isinstance(parsed["number"], str) and parsed["number"].lower() == "not found":
        parsed["number"] = None
    return parsed


def normalize_geocode_token(value: Any) -> str:
    text = clean_cell(value).lower()
    text = text.replace("-", " ").replace("'", "").replace('"', "")
    text = text.replace("״", "").replace("׳", "")
    return re.sub(r"[^a-zא-ת0-9]+", " ", text).strip()


def unique_clean_values(values: list[Any]) -> list[str]:
    seen: set[str] = set()
    unique: list[str] = []
    for value in values:
        text = clean_cell(value)
        key = normalize_geocode_token(text)
        if text and key and key not in seen:
            seen.add(key)
            unique.append(text)
    return unique


def fix_known_typos(text: str) -> str:
    for typo, correct in TYPO_FIXES.items():
        text = text.replace(typo, correct)
    return text


def normalize_marks(text: str) -> str:
    text = text.replace("“", '"').replace("”", '"').replace("`", "'")
    text = re.sub(r"^[\"'׳]+|[\"'׳]+$", "", text)
    return text.replace('""', '"')


def strip_city_from_address(address: str, city: str) -> str:
    if not city:
        return address
    return re.sub(re.escape(city), " ", address, flags=re.IGNORECASE).strip()


def remove_noise(text: str) -> str:
    text = PHONE_PATTERN.sub(" ", text)
    text = DATE_TIME_PATTERN.sub(" ", text)
    text = re.sub(r"\b(?:טל|טלפון|נייד|פלאפון|phone|mobile)\b.*$", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"\b(?:נא\s+לתאם|לתאם|הערה|הערות|comments?)\b.*$", " ", text, flags=re.IGNORECASE)
    text = text.replace("\\", " ").replace("|", " ").replace(",", " ").replace(";", " ")
    return re.sub(r"\s+", " ", text).strip()


def normalize_address_text(text: Any, city: Any = "") -> str:
    value = clean_cell(text)
    value = normalize_marks(value)
    value = strip_city_from_address(value, clean_cell(city))
    value = fix_known_typos(value)
    value = remove_noise(value)
    value = re.sub(r"(\d)(קומה|floor)\b", r"\1 \2", value, flags=re.IGNORECASE)
    value = re.sub(r"^(\d+[A-Za-zא-ת/-]*)\s+(דירה|דיר[׳'\"]?|apt\.?|apartment|unit)\s*$", r"\2 \1", value, flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", value).strip(" ,/")


def normalize_street_name(street: str) -> str:
    street = normalize_marks(clean_cell(street))
    street = STREET_PREFIX_PATTERN.sub("", street).strip()
    street = fix_known_typos(street)
    street = normalize_marks(street)
    street = fix_known_typos(street)
    street = re.sub(r"^ר\s+", "רבי ", street)
    return re.sub(r"\s+", " ", street).strip(" ,-/")


def extract_secondary_details(text: str) -> tuple[dict[str, str], str]:
    details = {"apartment": "", "floor": "", "entrance": ""}
    patterns = [
        (r"(?:דירה|דיר[׳'\"]?|apt\.?|apartment|unit)\s*(?:מספר|מס[׳'\"]?)?\s*[-:]?\s*([\wא-ת/-]+)", "apartment"),
        (r"(?:קומה|floor)\s*:?\s*(-?[\wא-ת/]+-?)", "floor"),
        (r"(?:כניסה|entrance)\s*[-:]?\s*([\wא-ת/-]+)", "entrance"),
    ]
    cleaned = text
    for pattern, key in patterns:
        match = re.search(pattern, cleaned, flags=re.IGNORECASE)
        if not match:
            continue
        value = match.group(1).strip()
        if key == "floor":
            raw_match = match.group(0)
            if f"-{value}" in raw_match or f"{value}-" in raw_match:
                value = f"-{value.strip('-')}"
        details[key] = value
        cleaned = re.sub(pattern, " ", cleaned, flags=re.IGNORECASE)
    return details, re.sub(r"\s+", " ", cleaned).strip(" ,/")


def merge_detail(primary: str, secondary: str) -> str:
    return clean_cell(primary) or clean_cell(secondary)


def append_class(classes: list[str], value: str) -> None:
    if value and value not in classes:
        classes.append(value)


def normalize_floor_value(value: str) -> str:
    value = normalize_address_text(value)
    value = re.sub(r"^(?:קומה|floor)\s+", "", value, flags=re.IGNORECASE).strip()
    value = re.sub(r"^(?:מינוס|minus)\s+", "-", value, flags=re.IGNORECASE).strip()
    if re.fullmatch(r"\d+-", value):
        return "-" + value[:-1]
    return value


def numeric_value(value: str) -> int | None:
    value = normalize_floor_value(value)
    if re.fullmatch(r"-?\d+", value):
        return int(value)
    return None


def is_floor_like(value: str) -> bool:
    number = numeric_value(value)
    return number is not None and -5 <= number <= 40


def normalize_number_token(value: str) -> str:
    value = clean_cell(value).strip(" ,/")
    if re.fullmatch(r"\d+-", value):
        return "-" + value[:-1]
    return value


def classify_secondary_part(
    text: str,
    primary_house: str = "",
    primary_apartment: str = "",
    building_found_in_primary: bool = False,
) -> tuple[dict[str, str], str, list[str], bool]:
    details, remainder = extract_secondary_details(text)
    remainder = normalize_address_text(remainder)
    classes: list[str] = []
    needs_review = False
    if details["apartment"]:
        append_class(classes, "apartment_labeled")
    if details["floor"]:
        details["floor"] = normalize_floor_value(details["floor"])
        append_class(classes, "floor_labeled")
    if details["entrance"]:
        append_class(classes, "entrance_labeled")
    slash_pair = re.fullmatch(r"(-?\d+-?)\s*/\s*(-?\d+-?)", remainder)
    if slash_pair:
        details["apartment"] = merge_detail(details["apartment"], normalize_number_token(slash_pair.group(2)))
        append_class(classes, "building_and_apartment_from_street2_slash")
        remainder = ""
    bare_number = normalize_floor_value(remainder)
    if re.fullmatch(r"-?\d+", bare_number):
        if primary_house and bare_number == primary_house:
            append_class(classes, "duplicate_building_number")
            remainder = ""
        elif primary_apartment and bare_number == primary_apartment:
            append_class(classes, "duplicate_apartment_number")
            remainder = ""
        elif primary_house and primary_apartment and is_floor_like(bare_number):
            details["floor"] = merge_detail(details["floor"], bare_number)
            append_class(classes, "floor_from_bare_secondary_number")
            remainder = ""
        elif building_found_in_primary and numeric_value(bare_number) is not None and numeric_value(bare_number) < 0:
            details["floor"] = merge_detail(details["floor"], bare_number)
            append_class(classes, "floor_from_negative_secondary_number")
            remainder = ""
        elif building_found_in_primary and not details["apartment"]:
            details["apartment"] = bare_number
            append_class(classes, "apartment_from_bare_secondary_number")
            remainder = ""
        else:
            append_class(classes, "ambiguous_secondary_number")
            needs_review = True
    if remainder:
        append_class(classes, "secondary_note")
    return details, remainder, classes, needs_review


def remove_secondary_tail(text: str) -> str:
    text = re.sub(r"\b(?:מספר|מס[׳'\"]?)\s+(?=\d)", "", text)
    text = re.sub(
        r"\b(?:דירה|דיר[׳'\"]?|קומה|כניסה|בניין|בנין|יחידה|דלת|apt\.?|apartment|floor|entrance|unit)\b.*$",
        "",
        text,
        flags=re.IGNORECASE,
    )
    return re.sub(r"\s+", " ", text).strip(" ,-/")


def is_date_like_number(value: str) -> bool:
    value = clean_cell(value)
    return bool(re.fullmatch(r"\d{4}[-/]\d{1,2}[-/]\d{1,2}", value) or re.fullmatch(r"\d{1,2}[-/]\d{1,2}[-/]\d{2,4}", value))


def number_tokens(text: str) -> list[re.Match[str]]:
    return list(re.finditer(r"(?<!\d)-?\d+-?(?!\d)", text))


def parse_primary_address_part(address: str) -> tuple[str, str, str, list[str], str]:
    text = remove_secondary_tail(normalize_address_text(address))
    classes: list[str] = []
    if not text:
        return "", "", "", classes, "empty address"
    text = re.sub(r"\s*/\s*", "/", text)
    slash_match = re.match(r"^(?P<street>.*?[^\d\s])\s*(?P<house>-?\d+-?)\s*/\s*(?P<apt>-?\d+-?)$", text)
    if slash_match:
        street = normalize_street_name(slash_match.group("street"))
        house = normalize_number_token(slash_match.group("house"))
        apartment = normalize_number_token(slash_match.group("apt"))
        append_class(classes, "apartment_from_slash_in_street1")
        if street and house:
            return street, house, apartment, classes, ""
    matches = number_tokens(text)
    if not matches:
        return normalize_street_name(text), "", "", classes, "missing house number"
    first = matches[0]
    house = normalize_number_token(first.group(0))
    if is_date_like_number(house):
        return normalize_street_name(text), "", "", classes, "date-like value is not a house number"
    if len(matches) >= 2:
        second = matches[1]
        apartment = normalize_number_token(second.group(0))
        before_first = text[: first.start()].strip(" ,-/")
        between = text[first.end() : second.start()].strip(" ,-/")
        after_second = text[second.end() :].strip(" ,-/")
        street = normalize_street_name(before_first or between or after_second)
        if street and not re.fullmatch(r"[\d\s/-]+", street):
            append_class(classes, "apartment_from_second_number_in_street1")
            return street, house, apartment, classes, ""
    before = text[: first.start()].strip(" ,-/")
    after = text[first.end() :].strip(" ,-/")
    street = normalize_street_name(before or after)
    if not street or re.fullmatch(r"[\d\s/-]+", street):
        return "", house, "", classes, "missing street"
    return street, house, "", classes, ""


def confidence_for(city: str, street: str, house_number: str, parse_error: str) -> tuple[str, str]:
    if city and street and house_number:
        return "high", "cleaned"
    if street and house_number:
        return "medium", "missing city"
    if street and not house_number:
        return "low", parse_error or "missing house number"
    return "none", parse_error or "could not parse address"


def format_cleaned_address(street: str, house_number: str) -> str:
    return " ".join(part for part in [clean_cell(street), clean_cell(house_number)] if part)


def clean_raw_address(city: Any, street1: Any, street2: Any, use_skill: bool = True) -> CleanedAddress:
    city_text = fix_known_typos(clean_cell(city))
    part1 = normalize_address_text(street1, city_text)
    part2 = normalize_address_text(street2, city_text)
    merged = " ".join(part for part in [part1, part2] if part)
    primary_street, primary_number, primary_apartment, primary_classes, _primary_error = parse_primary_address_part(part1)
    secondary_details, secondary_notes, secondary_classes, needs_review = classify_secondary_part(
        part2,
        primary_house=primary_number,
        primary_apartment=primary_apartment,
        building_found_in_primary=bool(primary_street and primary_number),
    )
    merged_secondary_details, merged_without_secondary = extract_secondary_details(merged)
    secondary_details = {
        "apartment": merge_detail(primary_apartment, merge_detail(secondary_details["apartment"], merged_secondary_details["apartment"])),
        "floor": merge_detail(secondary_details["floor"], normalize_floor_value(merged_secondary_details["floor"])),
        "entrance": merge_detail(secondary_details["entrance"], merged_secondary_details["entrance"]),
    }
    classifications = primary_classes + secondary_classes
    street = ""
    house_number = ""
    parse_error = "empty address"
    cleaned_candidate = ""
    seen: set[str] = set()
    for candidate in [part1, merged_without_secondary, merged, part2]:
        candidate = normalize_address_text(candidate, city_text)
        if not candidate or candidate in seen:
            continue
        seen.add(candidate)
        cand_street, cand_number, cand_apartment, cand_classes, cand_error = parse_primary_address_part(candidate)
        parse_error = cand_error
        if cand_street and cand_number:
            street = cand_street
            house_number = cand_number
            secondary_details["apartment"] = merge_detail(secondary_details["apartment"], cand_apartment)
            for cand_class in cand_classes:
                append_class(classifications, cand_class)
            cleaned_candidate = candidate
            break
        if not street and cand_street:
            street = cand_street
            cleaned_candidate = candidate
    if secondary_details["apartment"] == house_number and normalize_address_text(part2) == house_number:
        secondary_details["apartment"] = ""
    if secondary_notes in {house_number, secondary_details["apartment"], secondary_details["floor"]}:
        secondary_notes = ""
    if not secondary_notes:
        classifications = [value for value in classifications if value != "secondary_note"]
    if secondary_notes:
        needs_review = needs_review or bool(re.fullmatch(r"-?\d+", normalize_floor_value(secondary_notes)))
    confidence, status = confidence_for(city_text, street, house_number, parse_error)
    if classifications:
        status = f"{status}; {'; '.join(classifications)}"
    skill_output = ""
    if use_skill and street and house_number:
        skill_query = " ".join(part for part in [street, house_number, city_text] if part)
        skill_output = parse_address(skill_query) or ""
        skill_parsed = parse_lookup_output(skill_output) if skill_output else {}
        if skill_parsed.get("formatted") and confidence == "medium":
            confidence = "high"
            status = "cleaned and skill formatted"
        elif skill_output and status == "cleaned":
            status = "cleaned and skill checked"
    return CleanedAddress(
        city=city_text,
        street=street,
        house_number=house_number,
        apartment=secondary_details["apartment"],
        floor=secondary_details["floor"],
        entrance=secondary_details["entrance"],
        secondary_notes=secondary_notes,
        secondary_classification="; ".join(classifications),
        needs_review=needs_review,
        merged_address=merged,
        cleaned_address=format_cleaned_address(street, house_number) or cleaned_candidate,
        status=status,
        confidence=confidence,
        parser_output=skill_output,
    )


def city_match_options(city: str) -> list[str]:
    city = clean_cell(city)
    options = [city] + CITY_GEOCODE_ALIASES.get(city, [])
    if city == "מודיעין":
        options.extend(["מודיעין-מכבים-רעות", "מודיעין מכבים רעות"])
    if city == "תל אביב":
        options.extend(["תל אביב-יפו", "תל אביב יפו"])
    return unique_clean_values(options)


def street_match_options(street: str) -> list[str]:
    street = normalize_street_name(street)
    options = [street] + STREET_GEOCODE_ALIASES.get(street, [])
    if street.startswith("שדרות "):
        options.append(street.replace("שדרות ", "", 1))
    elif street.startswith("רחוב "):
        options.append(street.replace("רחוב ", "", 1))
    else:
        options.append(f"רחוב {street}")
        options.append(f"שדרות {street}")
    return unique_clean_values(options)


def formatted_address_matches_city(formatted_address: str, city: str) -> bool:
    formatted_norm = normalize_geocode_token(formatted_address)
    for option in city_match_options(city):
        option_norm = normalize_geocode_token(option)
        if option_norm and option_norm in formatted_norm:
            return True
    return not clean_cell(city)


def google_result_matches_city(result: dict[str, Any], city: str) -> bool:
    if not clean_cell(city):
        return True
    city_options = [normalize_geocode_token(option) for option in city_match_options(city)]
    component_values: list[str] = []
    for component in result.get("address_components", []):
        types = set(component.get("types", []))
        if types & CITY_COMPONENT_TYPES:
            component_values.extend([component.get("long_name", ""), component.get("short_name", "")])
    for value in component_values:
        value_norm = normalize_geocode_token(value)
        if value_norm and any(option and (option == value_norm or option in value_norm or value_norm in option) for option in city_options):
            return True
    return formatted_address_matches_city(str(result.get("formatted_address", "")), city)


def geocode_empty_response(status: str, query: str = "", failure_reason: str = "", attempt_count: int = 0) -> dict[str, Any]:
    return {
        "LAT": "",
        "LNG": "",
        "Coordinates": "",
        "Geocode_Query": query,
        "Geocode_Status": status,
        "Geocode_Precise": "no",
        "Geocode_Usable": "no",
        "Geocode_Formatted": "",
        "Geocode_Attempt_Count": attempt_count,
        "Geocode_Query_Used": query,
        "Geocode_Result_Types": "",
        "Geocode_Location_Type": "",
        "Geocode_Diagnostic_Coordinates": "",
        "Geocode_Source": "none",
        "Geocode_Estimated": "no",
        "Geocode_Failure_Reason": failure_reason or status,
        "Google_Street": "",
        "Google_House_Number": "",
        "Google_City": "",
    }


def build_geocode_attempts(address: str, city: str, street: str = "", house_number: str = "") -> list[GeocodeAttempt]:
    address = clean_cell(address)
    city = clean_cell(city)
    street = normalize_street_name(street)
    house_number = clean_cell(house_number)
    city_options = city_match_options(city) or [city]
    primary_city = city_options[0] if city_options else city
    alias_cities = city_options[1:]
    street_options = street_match_options(street) if street else []
    primary_street = street_options[0] if street_options else street
    alias_streets = street_options[1:]
    attempts: list[GeocodeAttempt] = []
    seen: set[tuple[str, str]] = set()

    def add(query: str, language: str, reason: str) -> None:
        query = re.sub(r"\s+", " ", clean_cell(query)).strip(" ,")
        key = (normalize_geocode_token(query), language)
        if query and key not in seen:
            seen.add(key)
            attempts.append(GeocodeAttempt(query=query, language=language, reason=reason))

    add(", ".join(part for part in [address, primary_city, "Israel"] if part), "he", "cleaned address")
    if primary_street and house_number and primary_city:
        add(f"{primary_street} {house_number}, {primary_city}, Israel", "he", "structured exact address")
    if house_number and primary_city:
        for street_option in alias_streets:
            add(f"{street_option} {house_number}, {primary_city}, Israel", "he", "street alias")
            add(f"{street_option} {house_number}, {primary_city}, Israel", "en", "street alias english")
    for city_option in alias_cities:
        add(", ".join(part for part in [address, city_option, "Israel"] if part), "he", "city alias")
        if primary_street and house_number:
            add(f"{primary_street} {house_number}, {city_option}, Israel", "he", "structured city alias")
            add(f"{primary_street} {house_number}, {city_option}, Israel", "en", "structured city alias english")
    for city_option in city_options:
        add(", ".join(part for part in [address, city_option, "Israel"] if part), "en", "cleaned address english")
    return attempts


def call_google_geocode(attempt: GeocodeAttempt, api_key: str) -> dict[str, Any]:
    response = requests.get(
        "https://maps.googleapis.com/maps/api/geocode/json",
        params={
            "address": attempt.query,
            "components": "country:IL",
            "region": "il",
            "language": attempt.language,
            "key": api_key,
        },
        timeout=30,
    )
    return response.json()


def sanitize_error_text(text: Any) -> str:
    return re.sub(r"([?&]key=)[^&\\s)]+", r"\1<redacted>", str(text))


def extract_google_address_components(result: dict[str, Any]) -> dict[str, str]:
    extracted = {"street": "", "house_number": "", "city": ""}
    for component in result.get("address_components", []):
        types = set(component.get("types", []))
        value = clean_cell(component.get("long_name", ""))
        if not value:
            continue
        if "route" in types and not extracted["street"]:
            extracted["street"] = normalize_street_name(value)
        elif "street_number" in types and not extracted["house_number"]:
            extracted["house_number"] = clean_cell(value)
        elif types & CITY_COMPONENT_TYPES and not extracted["city"]:
            extracted["city"] = clean_cell(value)
    return extracted


def geocode_result_row(
    result: dict[str, Any],
    attempt: GeocodeAttempt,
    attempt_count: int,
    is_precise: bool,
    usable: bool,
    failure_reason: str = "",
) -> dict[str, Any]:
    formatted = str(result.get("formatted_address", ""))
    components = extract_google_address_components(result)
    location = result.get("geometry", {}).get("location", {})
    raw_lat = location.get("lat", "")
    raw_lng = location.get("lng", "")
    lat = raw_lat if usable else ""
    lng = raw_lng if usable else ""
    result_types = sorted(str(value) for value in result.get("types", []))
    location_type = str(result.get("geometry", {}).get("location_type", ""))
    status_word = "precise" if is_precise else "fallback"
    return {
        "LAT": lat,
        "LNG": lng,
        "Coordinates": f"{lat},{lng}" if lat != "" and lng != "" else "",
        "Geocode_Query": attempt.query,
        "Geocode_Status": f"google {status_word} match: {formatted}",
        "Geocode_Precise": "yes" if is_precise else "no",
        "Geocode_Usable": "yes" if usable else "no",
        "Geocode_Formatted": formatted,
        "Geocode_Attempt_Count": attempt_count,
        "Geocode_Query_Used": attempt.query,
        "Geocode_Result_Types": "; ".join(result_types),
        "Geocode_Location_Type": location_type,
        "Geocode_Diagnostic_Coordinates": f"{raw_lat},{raw_lng}" if raw_lat != "" and raw_lng != "" and not usable else "",
        "Geocode_Source": "google_precise" if usable else "google_fallback",
        "Geocode_Estimated": "no",
        "Geocode_Failure_Reason": failure_reason,
        "Google_Street": components["street"],
        "Google_House_Number": components["house_number"],
        "Google_City": components["city"],
    }


def geocode_address(address: str, city: str, api_key: str, street: str = "", house_number: str = "") -> dict[str, Any]:
    all_attempts = build_geocode_attempts(address, city, street=street, house_number=house_number)
    attempts = all_attempts[:MAX_GEOCODE_ATTEMPTS_PER_ADDRESS]
    if not attempts:
        return geocode_empty_response("empty geocode query")
    fallback: dict[str, Any] | None = None
    wrong_city_precise = False
    last_failure = ""
    completed_attempts = 0
    for attempt in attempts:
        completed_attempts += 1
        try:
            data = call_google_geocode(attempt, api_key)
        except Exception as exc:
            last_failure = f"geocode request failed: {sanitize_error_text(exc)}"
            continue
        status = data.get("status", "no result")
        if status != "OK" or not data.get("results"):
            last_failure = f"google geocode failed: {status}"
            if status in {"REQUEST_DENIED", "OVER_DAILY_LIMIT", "OVER_QUERY_LIMIT", "INVALID_REQUEST"}:
                return geocode_empty_response(last_failure, attempt.query, last_failure, completed_attempts)
            continue
        for result in data.get("results", []):
            location = result.get("geometry", {}).get("location", {})
            if location.get("lat") is None or location.get("lng") is None:
                last_failure = "google result missing coordinates"
                continue
            result_types = set(result.get("types", []))
            location_type = str(result.get("geometry", {}).get("location_type", ""))
            is_precise = bool(result_types & PRECISE_GEOCODE_TYPES) or location_type in PRECISE_LOCATION_TYPES
            city_matches = google_result_matches_city(result, city)
            if is_precise and city_matches:
                return geocode_result_row(result, attempt, completed_attempts, is_precise=True, usable=True)
            if is_precise and not city_matches:
                wrong_city_precise = True
                last_failure = "google found a precise address in a different city"
                continue
            if fallback is None and city_matches:
                fallback = geocode_result_row(
                    result,
                    attempt,
                    completed_attempts,
                    is_precise=False,
                    usable=False,
                    failure_reason="google returned only a non-address fallback",
                )
    if fallback is not None:
        fallback["Geocode_Attempt_Count"] = completed_attempts
        return fallback
    failure = "google found a precise address in a different city" if wrong_city_precise else (last_failure or "google could not validate this address")
    if len(all_attempts) > MAX_GEOCODE_ATTEMPTS_PER_ADDRESS:
        failure = f"{failure}; stopped after {MAX_GEOCODE_ATTEMPTS_PER_ADDRESS} geocode attempts"
    return geocode_empty_response(failure, attempts[0].query, failure, completed_attempts)


def numeric_house_number(value: Any) -> int | None:
    match = re.match(r"^\s*(\d+)", clean_cell(value))
    return int(match.group(1)) if match else None


def float_or_none(value: Any) -> float | None:
    try:
        if clean_cell(value) == "":
            return None
        return float(value)
    except (TypeError, ValueError):
        return None


def geocode_group_key(row: dict[str, Any]) -> tuple[str, str]:
    return (normalize_geocode_token(row.get("City", "")), normalize_geocode_token(row.get("Street_Name", "")))


def estimate_failed_geocodes_from_neighbors(rows: list[dict[str, Any]]) -> int:
    successful_by_street: dict[tuple[str, str], list[dict[str, Any]]] = {}
    for row in rows:
        if clean_cell(row.get("Geocode_Usable", "")).lower() != "yes":
            continue
        house = numeric_house_number(row.get("House_Number", ""))
        lat = float_or_none(row.get("LAT", ""))
        lng = float_or_none(row.get("LNG", ""))
        if house is None or lat is None or lng is None:
            continue
        successful_by_street.setdefault(geocode_group_key(row), []).append({"house": house, "lat": lat, "lng": lng, "source_row": row.get("source_row", "")})
    estimate_count = 0
    for row in rows:
        if clean_cell(row.get("Geocode_Usable", "")).lower() == "yes":
            continue
        house = numeric_house_number(row.get("House_Number", ""))
        if house is None:
            continue
        neighbors = sorted(successful_by_street.get(geocode_group_key(row), []), key=lambda item: item["house"])
        if not neighbors:
            continue
        lower = [item for item in neighbors if item["house"] <= house]
        upper = [item for item in neighbors if item["house"] >= house]
        before = lower[-1] if lower else None
        after = upper[0] if upper else None
        method = "same_street_nearest_house"
        if before and after and before["house"] != after["house"]:
            span = after["house"] - before["house"]
            ratio = (house - before["house"]) / span
            lat = before["lat"] + (after["lat"] - before["lat"]) * ratio
            lng = before["lng"] + (after["lng"] - before["lng"]) * ratio
            source = f"between house {before['house']} row {before['source_row']} and house {after['house']} row {after['source_row']}"
            method = "same_street_interpolated"
        else:
            nearest = min(neighbors, key=lambda item: abs(item["house"] - house))
            if abs(nearest["house"] - house) > 20:
                continue
            lat = nearest["lat"]
            lng = nearest["lng"]
            source = f"nearest house {nearest['house']} row {nearest['source_row']}"
        row["LAT"] = round(lat, 7)
        row["LNG"] = round(lng, 7)
        row["Coordinates"] = f"{row['LAT']},{row['LNG']}"
        row["Geocode_Status"] = f"estimated coordinates from {method}: {source}"
        row["Geocode_Usable"] = "review"
        row["Geocode_Source"] = method
        row["Geocode_Estimated"] = "yes"
        row["Geocode_Failure_Reason"] = f"google did not return an exact address; estimated from {source}"
        estimate_count += 1
    return estimate_count


def coordinate_confidence(row: dict[str, Any]) -> str:
    usable = clean_cell(row.get("Geocode_Usable", "")).lower()
    if usable == "yes":
        return "exact_google"
    if usable == "review":
        return "estimated_review"
    return "missing"


def is_route_ready(row: dict[str, Any]) -> bool:
    return coordinate_confidence(row) in {"exact_google", "estimated_review"} and bool(clean_cell(row.get("LAT", "")) and clean_cell(row.get("LNG", "")))

