from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import pandas as pd


RAW_COLUMN_ALIASES = {
    "order_id": ["order id", "order_id", "orderid"],
    "client_id": ["client id", "client_id", "clientid"],
    "site_id": ["site id", "site_id", "siteid"],
    "site_name": ["site name", "site_name", "sitename"],
    "ship_to_street1": ["ship to street 1", "ship_to_street1", "ship to street1"],
    "ship_to_street2": ["ship to street 2", "ship_to_street2", "ship to street2"],
    "required_delivery_date": [
        "required delivery date",
        "required_delivery_date",
        "delivery date",
    ],
    "comments": ["comments", "comment", "notes"],
}

KNOWN_CITY_CODES = {
    "ירושלים": 3000,
    "תל אביב": 5000,
    "תל אביב יפו": 5000,
    "חיפה": 4000,
    "ראשון לציון": 8300,
    "פתח תקווה": 7900,
    "אשדוד": 70,
    "נתניה": 7400,
    "באר שבע": 9000,
    "חולון": 6600,
    "בני ברק": 6100,
    "בית שמש": 2610,
    "מודיעין": 1200,
    "מודיעין עילית": 3797,
    "לוד": 7000,
    "עפולה": 7700,
    "רכסים": 922,
}

STREET_NORMALIZATIONS = [
    (r"[\"״׳']", ""),
    (r"\bר\s+", "רבי "),
    (r"\bרבי עקיבה\b", "רבי עקיבא"),
    (r"\bרשבי\b", "רבי שמעון בר יוחאי"),
    (r"\bרשבא\b", "רשב״א"),
    (r"\bהריטבא\b", "הריטב״א"),
    (r"\bשד\s+", "שדרות "),
]


def clean_cell(value: Any) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def normalize_column_name(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value).strip().lower().replace("_", " "))


def resolve_columns(df: pd.DataFrame) -> dict[str, str]:
    normalized = {normalize_column_name(col): col for col in df.columns}
    resolved = {}
    for canonical, aliases in RAW_COLUMN_ALIASES.items():
        for alias in aliases:
            col = normalized.get(normalize_column_name(alias))
            if col is not None:
                resolved[canonical] = col
                break
    return resolved


def strip_city_from_address(address: str, city: str) -> str:
    if not city:
        return address
    return re.sub(re.escape(city), " ", address, flags=re.IGNORECASE).strip()


def normalize_address_text(street1: Any, street2: Any, city: Any) -> str:
    text = " ".join(part for part in [clean_cell(street1), clean_cell(street2)] if part)
    text = strip_city_from_address(text, clean_cell(city))
    text = text.replace("\\", " ").replace("|", " ")
    text = re.sub(r"[,:;]+", " ", text)
    text = re.sub(r"\b(טל|טלפון|נייד|פלאפון)\b.*$", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def parse_street_and_number(address: str) -> tuple[str, str, str]:
    if not address:
        return "", "", "empty address"

    address = re.sub(r"\s*/\s*", "/", address)
    match = re.search(r"\d+[\wא-ת/-]*", address)
    if not match:
        return address.strip(), "", "missing house number"

    number = match.group(0).strip()
    before = address[: match.start()].strip(" ,-/")
    after = address[match.end() :].strip(" ,-/")

    if before:
        street = before
    else:
        street = after

    street = normalize_street_name(street)
    street = re.sub(r"\s+", " ", street)
    if not street:
        return "", number, "missing street"
    return street, number, ""


def normalize_street_name(street: str) -> str:
    street = re.sub(r"^(רחוב|רח׳|רח'|רח)\s+", "", clean_cell(street)).strip()
    for pattern, replacement in STREET_NORMALIZATIONS:
        street = re.sub(pattern, replacement, street)
    return re.sub(r"\s+", " ", street).strip()


def validate_components(city: str, street: str, house_number: str) -> tuple[bool, str]:
    if not city:
        return False, "missing city"
    if not street:
        return False, "missing street"
    if not house_number:
        return False, "missing house number"
    if not re.match(r"^\d+", house_number):
        return False, "invalid house number"
    return True, "verified by Israeli address parser"


def clean_raw_orders(input_path: Path, run_dir: Path) -> dict[str, Path]:
    df = pd.read_excel(input_path).fillna("")
    columns = resolve_columns(df)
    missing = [
        name
        for name in ["order_id", "client_id", "site_id", "site_name", "ship_to_street1", "ship_to_street2", "comments"]
        if name not in columns
    ]
    if missing:
        raise RuntimeError(f"Stage 1 input is missing columns: {', '.join(missing)}")

    if "required_delivery_date" in columns:
        df = df.drop(columns=[columns["required_delivery_date"]])

    df = df.sort_values(by=[columns["site_name"]], kind="stable").reset_index(drop=True)

    good_address_rows = []
    good_original_rows = []
    failed_rows = []

    for source_index, row in df.iterrows():
        city = clean_cell(row[columns["site_name"]])
        merged = normalize_address_text(
            row[columns["ship_to_street1"]],
            row[columns["ship_to_street2"]],
            city,
        )
        street, house_number, parse_error = parse_street_and_number(merged)
        valid, status = validate_components(city, street, house_number)
        city_code = KNOWN_CITY_CODES.get(city, "")

        result = row.to_dict()
        result.update(
            {
                "source_row": source_index + 2,
                "City": city,
                "Street_Name": street,
                "House_Number": house_number,
                "merged_address": merged,
                "city_code": city_code,
                "cleanup_status": status if valid else parse_error or status,
            }
        )

        if valid:
            good_address_rows.append(
                {
                    "City": city,
                    "Street_Name": street,
                    "House_Number": house_number,
                }
            )
            good_original_rows.append(result)
        else:
            failed_rows.append(result)

    result_columns = list(df.columns) + [
        "source_row",
        "City",
        "Street_Name",
        "House_Number",
        "merged_address",
        "city_code",
        "cleanup_status",
    ]
    good_addresses = pd.DataFrame(good_address_rows, columns=["City", "Street_Name", "House_Number"])
    good_original = pd.DataFrame(good_original_rows, columns=result_columns)
    failed = pd.DataFrame(failed_rows, columns=result_columns)

    failed_path = run_dir / "01a_failed_addresses.xlsx"
    good_path = run_dir / "01b_addresses_for_geocoding.xlsx"
    original_path = run_dir / "01c_good_orders_original_format.xlsx"

    failed.to_excel(failed_path, index=False)
    good_addresses.to_excel(good_path, index=False)
    good_original.to_excel(original_path, index=False)

    return {
        "failed": failed_path,
        "good_addresses": good_path,
        "good_original": original_path,
    }


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(description="Clean Routecraft raw order addresses.")
    parser.add_argument("--input", required=True)
    parser.add_argument("--output-dir", default=".")
    args = parser.parse_args()

    paths = clean_raw_orders(Path(args.input), Path(args.output_dir))
    for name, path in paths.items():
        print(f"{name}: {path}")


if __name__ == "__main__":
    main()
