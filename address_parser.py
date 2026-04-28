#!/usr/bin/env python3
"""
Israeli Address Parser - Wrapper script for address decoding

This script uses the israeli-address-autocomplete skill to parse and validate
Israeli addresses, returning structured components.

Usage:
    python address_parser.py "רחוב הרצל 15, תל אביב"
    python address_parser.py "Rabbi Akiva 14, Modi'in Illit"
    python address_parser.py --file address_input.txt
    python address_parser.py -f address_input.txt
    python address_parser.py --excel input.xlsx
    python address_parser.py -x input.xlsx
    python address_parser.py --excel input.xlsx --output output.xlsx

Interactive mode (no arguments):
    python address_parser.py
"""

import argparse
import subprocess
import sys
import re
from pathlib import Path

import pandas as pd


def get_lookup_script():
    """Find the lookup_address.py script location."""
    current_dir = Path(__file__).parent
    script_path = current_dir / ".agents" / "skills" / "israeli-address-autocomplete" / "scripts" / "lookup_address.py"
    if script_path.exists():
        return str(script_path)
    abs_path = Path(r"c:\dev\routecraft\.agents\skills\israeli-address-autocomplete\scripts\lookup_address.py")
    if abs_path.exists():
        return str(abs_path)
    return None


SKILL_SCRIPT = get_lookup_script()

TYPO_FIXES = {
    "רבי עקיבה": "רבי עקיבא",
    "חזוניש": "חזון איש",
    "תל אביביפו": "תל אביב יפו",
}

CITY_COLUMN_PATTERNS = ["city", "עיר", "יישוב", "city name", "cityname"]
LINE1_COLUMN_PATTERNS = ["address line 1", "line1", "address1", "כתובת 1", "כתובת1", "line 1"]
LINE2_COLUMN_PATTERNS = ["address line 2", "line2", "address2", "כתובת 2", "כתובת2", "line 2"]

DATE_TIME_PATTERN = re.compile(
    r"(\b\d{1,2}[:.]\d{2}([:.]\d{2})?\b|\b\d{1,2}[-/.]\d{1,2}([-/.]\d{2,4})?\b|\b(?:am|pm|שעה|תאריך|יום)\b)",
    flags=re.IGNORECASE,
)


def parse_address(address_string):
    """Parse an Israeli address using the lookup_address script."""
    if not SKILL_SCRIPT:
        print(f"Error: lookup_address.py script not found in any expected location", file=sys.stderr)
        return None

    try:
        result = subprocess.run(
            [sys.executable, str(SKILL_SCRIPT), "format", address_string],
            capture_output=True,
            text=True,
            check=True,
        )
        return result.stdout
    except subprocess.CalledProcessError as e:
        print(f"Error parsing address: {e.stderr}", file=sys.stderr)
        return None


def parse_lookup_output(text):
    """Parse plain text output from lookup_address.py into fields."""
    parsed = {
        "street": None,
        "number": None,
        "city": None,
        "formatted": False,
        "raw_output": text,
    }
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

    if parsed["city"] and parsed["city"].lower() == "not identified":
        parsed["city"] = None
    if parsed["number"] and parsed["number"].lower() == "not found":
        parsed["number"] = None
    return parsed


def remove_date_time_tokens(text):
    return DATE_TIME_PATTERN.sub("", text).strip()


def is_probably_date_or_time(text):
    if not text:
        return True
    if DATE_TIME_PATTERN.search(text) and not re.search(r"[א-תA-Za-z]", text):
        return True
    if len(re.findall(r"\d", text)) >= 6 and not re.search(r"[א-תA-Za-z]", text):
        return True
    return False


def extract_apartment_floor(text):
    info = {"apartment": None, "floor": None}
    patterns = [
        (r"דירה\s*[-–]?\s*(\d+)", "apartment"),
        (r"apt\.?\s*[-–]?\s*(\d+)", "apartment"),
        (r"unit\s*[-–]?\s*(\d+)", "apartment"),
        (r"קומה\s*[-–]?\s*(-?\d+)", "floor"),
        (r"floor\s*[-–]?\s*(-?\d+)", "floor"),
    ]
    for pattern, key in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            info[key] = match.group(1)
            text = re.sub(pattern, "", text, flags=re.IGNORECASE)
    return info, text.strip()


def fix_known_typos(text):
    for typo, correct in TYPO_FIXES.items():
        text = text.replace(typo, correct)
    return text


def normalize_address(address):
    address = str(address or "").strip()
    if not address:
        return ""
    address = address.replace("\n", " ").replace(",", " ").replace("/", " ")
    address = fix_known_typos(address)
    address = remove_date_time_tokens(address)
    address = re.sub(r"\s+", " ", address)
    return address.strip()


def get_column_by_name_or_index(df, patterns, fallback_index):
    for col in df.columns:
        if isinstance(col, str):
            col_norm = col.strip().lower()
            for pattern in patterns:
                if pattern in col_norm:
                    return col
    if fallback_index < len(df.columns):
        return df.columns[fallback_index]
    return None


def process_file(file_path):
    file_obj = Path(file_path)
    if not file_obj.exists() or not file_obj.is_file():
        print(f"Error: File not found or not a file: {file_path}", file=sys.stderr)
        return

    with open(file_obj, "r", encoding="utf-8") as f:
        addresses = [line.strip() for line in f if line.strip()]

    if not addresses:
        print(f"Warning: File is empty: {file_path}")
        return

    print(f"\n{'='*60}")
    print(f"Processing {len(addresses)} address(es) from: {file_path}")
    print(f"{'='*60}\n")

    for idx, address in enumerate(addresses, 1):
        print(f"[{idx}] Parsing address: {address}")
        result = parse_address(address)
        format_output(result)


def build_result_strings(parsed, apartment_info, original_city):
    street = parsed.get("street") or ""
    number = parsed.get("number") or ""
    city = parsed.get("city") or original_city or ""

    street = street.strip()
    number = number.strip()
    city = city.strip()

    main_address = " ".join(part for part in [street, number, city] if part)

    friendly_parts = []
    if street:
        friendly_parts.append(f"רחוב {street}")
    if number:
        friendly_parts.append(f"מספר בניין {number}")
    if apartment_info.get("apartment"):
        friendly_parts.append(f"מספר דירה {apartment_info['apartment']}")
    if apartment_info.get("floor"):
        friendly_parts.append(f"קומה {apartment_info['floor']}")
    if city:
        friendly_parts.append(f"עיר {city}")

    friendly = ", ".join(friendly_parts)
    if not friendly:
        friendly = "כתובת לא מפוענחת"

    if parsed.get("formatted") and street and number and city:
        confidence = "high"
    elif street and number:
        confidence = "medium"
    elif street or number:
        confidence = "low"
    else:
        confidence = "none"

    if confidence != "high" and original_city:
        friendly += f" (מקורית: {original_city})"

    return main_address, friendly, confidence


def process_excel(input_path, output_path=None):
    input_file = Path(input_path)
    if not input_file.exists() or not input_file.is_file():
        print(f"Error: Excel file not found: {input_path}", file=sys.stderr)
        return

    if output_path:
        output_file = Path(output_path)
    else:
        output_file = input_file.with_name(input_file.stem + "_parsed.xlsx")

    try:
        df = pd.read_excel(input_file, engine="openpyxl", dtype=str)
    except Exception as e:
        print(f"Error reading Excel file: {e}", file=sys.stderr)
        return

    city_col = get_column_by_name_or_index(df, CITY_COLUMN_PATTERNS, 3)
    line1_col = get_column_by_name_or_index(df, LINE1_COLUMN_PATTERNS, 4)
    line2_col = get_column_by_name_or_index(df, LINE2_COLUMN_PATTERNS, 5)

    if city_col is None and line1_col is None and line2_col is None:
        print("Error: Could not identify city or address columns in Excel.", file=sys.stderr)
        return

    results = []
    for _, row in df.iterrows():
        original_city = str(row.get(city_col, "") if city_col is not None else "").strip()
        line1 = str(row.get(line1_col, "") if line1_col is not None else "").strip()
        line2 = str(row.get(line2_col, "") if line2_col is not None else "").strip()

        combined_raw = " ".join(part for part in [original_city, line1, line2] if part).strip()
        combined = normalize_address(combined_raw)
        if not combined or is_probably_date_or_time(combined):
            continue

        apartment_info, cleaned_address = extract_apartment_floor(combined)
        cleaned_address = normalize_address(cleaned_address)
        if not cleaned_address:
            continue

        parsed_text = parse_address(cleaned_address)
        parsed = parse_lookup_output(parsed_text or "") if parsed_text else {
            "street": None,
            "number": None,
            "city": None,
            "formatted": False,
            "raw_output": "",
        }

        main_address, friendly, confidence = build_result_strings(parsed, apartment_info, original_city)

        row_data = row.to_dict()
        row_data["ParsedAddress"] = main_address
        row_data["FriendlyAddress"] = friendly
        row_data["Confidence"] = confidence
        row_data["ParserOutput"] = parsed_text or ""
        results.append(row_data)

    if not results:
        print("No valid address rows found in the Excel file.", file=sys.stderr)
        return

    output_df = pd.DataFrame(results)
    try:
        output_df.to_excel(output_file, index=False, engine="openpyxl")
        print(f"Excel parsing completed. Output written to: {output_file}")
    except Exception as e:
        print(f"Error writing Excel file: {e}", file=sys.stderr)


def format_output(parsed_result):
    if parsed_result:
        print("\n" + "="*60)
        print("PARSED ADDRESS COMPONENTS")
        print("="*60)
        print(parsed_result)
        print("="*60 + "\n")


def main():
    parser = argparse.ArgumentParser(description="Israeli Address Parser")
    parser.add_argument("address", nargs="*", help="Address to parse directly")
    parser.add_argument("--file", "-f", help="Read addresses from a plain text file")
    parser.add_argument("--excel", "-x", help="Read addresses from an Excel file")
    parser.add_argument("--output", "-o", help="Output Excel file path")
    args = parser.parse_args()

    if args.excel:
        process_excel(args.excel, args.output)
        return

    if args.file:
        process_file(args.file)
        return

    if args.address:
        address = " ".join(args.address)
        print(f"Parsing address: {address}")
        result = parse_address(address)
        format_output(result)
        return

    print("\n" + "="*60)
    print("ISRAELI ADDRESS PARSER")
    print("="*60)
    print("Enter Israeli addresses to parse (Hebrew or English)")
    print("Type 'exit' to quit\n")

    while True:
        try:
            address = input("Enter address > ").strip()
            if not address:
                continue
            if address.lower() == "exit":
                print("Goodbye!")
                break

            result = parse_address(address)
            format_output(result)
        except KeyboardInterrupt:
            print("\n\nGoodbye!")
            break
        except Exception as e:
            print(f"Error: {e}\n", file=sys.stderr)


if __name__ == "__main__":
    main()
