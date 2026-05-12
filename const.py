from __future__ import annotations

import argparse
import contextlib
import hashlib
import html
import importlib.util
import importlib.machinery
import io
import json
import os
import re
import sys
import types
import traceback
import warnings
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any
from unittest.mock import patch

import numpy as np
import pandas as pd
import requests


ROOT = Path(__file__).resolve().parent
RUNS_DIR = ROOT / "runs"
MAX_GROUP_SIZE = 30
MOCK_DATA_DIR = ROOT / "addresses" / "mock_google_data"
CONST_MOCK_DATA_DIR = ROOT / "addresses" / "mock_google_data_const"
CONSTRAINTS_PARAMETERS_PATH = ROOT / "constraints_parameters_const.json"
WAREHOUSE_NAME = "WAREHOUSE ESHTAOL"
WAREHOUSE_LAT = 31.77927525
WAREHOUSE_LNG = 35.0105885


@dataclass
class StageResult:
    number: int
    name: str
    status: str
    output_files: list[Path] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)


def load_env_file(env_path: Path = ROOT / ".env") -> None:
    if not env_path.exists():
        return
    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def import_script(path: Path, module_name: str) -> Any:
    spec = importlib.util.spec_from_file_location(module_name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot import {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def timestamp_slug() -> str:
    return datetime.now().strftime("%Y-%m-%d_%H-%M-%S")


def make_run_dir(base_dir: Path) -> Path:
    slug = timestamp_slug()
    run_dir = base_dir / slug
    suffix = 2
    while run_dir.exists():
        run_dir = base_dir / f"{slug}_{suffix:02d}"
        suffix += 1
    run_dir.mkdir(parents=True, exist_ok=False)
    return run_dir


def require_file(path: Path, label: str) -> Path:
    resolved = path.expanduser().resolve()
    if not resolved.exists():
        raise FileNotFoundError(f"{label} does not exist: {resolved}")
    return resolved


def load_constraints_parameters(path: Path = CONSTRAINTS_PARAMETERS_PATH) -> dict[str, Any]:
    if not path.exists():
        return {}
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise RuntimeError(f"Constraints parameters file must contain a JSON object: {path}")
    return data


def without_local_ortools_shadow() -> list[str]:
    original_path = list(sys.path)
    root_text = str(ROOT)
    sys.path = [
        path
        for path in sys.path
        if path not in {"", ".", root_text} and Path(path or ".").resolve() != ROOT
    ]
    sys.modules.pop("ortools", None)
    return original_path


def constraint_value(args: argparse.Namespace, name: str, default: Any) -> Any:
    constraints = getattr(args, "constraints", {}) or {}
    return constraints.get(name, default)


def sanitize_filename(value: str) -> str:
    safe_value = re.sub(r'[<>:"/\\|?*]+', "_", value.strip())
    safe_value = re.sub(r"\s+", "_", safe_value)
    return safe_value or "unknown"


def get_google_key(args: argparse.Namespace) -> str:
    key = (
        args.google_api_key
        or os.getenv("GOOGLE_MAPS_API_KEY")
        or os.getenv("GEOCODING_API_KEY")
    )
    if not key:
        raise RuntimeError(
            "Missing Google Maps API key. Set GOOGLE_MAPS_API_KEY in .env, "
            "or pass --google-api-key."
        )
    return key


def get_distance_key(args: argparse.Namespace) -> str:
    key = (
        args.distance_api_key
        or os.getenv("DISTANCE_MATRIX_API_KEY")
    )
    if key:
        return key
    return get_google_key(args)


def stable_int(value: str, modulo: int) -> int:
    digest = hashlib.sha256(value.encode("utf-8")).hexdigest()
    return int(digest[:12], 16) % modulo


def mock_lat_lng(address: str) -> tuple[float, float]:
    lat_offset = stable_int(f"lat:{address}", 24000) / 1_000_000
    lng_offset = stable_int(f"lng:{address}", 24000) / 1_000_000
    return 31.880000 + lat_offset, 35.000000 + lng_offset


def mock_coords(address: str, _key: str = "") -> tuple[dict[str, float], str]:
    lat, lng = mock_lat_lng(address)
    return {"lat": lat, "lng": lng}, "mock"


def mock_distance_matrix(full_addresses: list[str], coords_dict: dict[str, Any]) -> pd.DataFrame:
    matrix_data = []
    for i, origin in enumerate(full_addresses):
        row_minutes = []
        for j, dest in enumerate(full_addresses):
            if i == j:
                row_minutes.append(0)
                continue
            origin_coords = coords_dict.get(origin)
            dest_coords = coords_dict.get(dest)
            if not origin_coords or not dest_coords:
                row_minutes.append(999)
                continue
            meters = calculate_haversine(
                origin_coords["lat"],
                origin_coords["lng"],
                dest_coords["lat"],
                dest_coords["lng"],
            )
            row_minutes.append(max(1, round((meters / 500) + 3)))
        matrix_data.append(row_minutes)
    return pd.DataFrame(matrix_data, index=full_addresses, columns=full_addresses)


def load_mock_geocode_cache() -> dict[tuple[str, str], dict[str, Any]]:
    path = CONST_MOCK_DATA_DIR / "geocoding_cache.xlsx"
    if not path.exists():
        path = MOCK_DATA_DIR / "geocoding_cache.xlsx"
    if not path.exists():
        return {}
    df = pd.read_excel(path).fillna("")
    cache = {}
    for _, row in df.iterrows():
        cache[(str(row.get("query", "")), str(row.get("city", "")))] = row.to_dict()
    return cache


def load_mock_coords_cache() -> dict[str, dict[str, float]]:
    path = CONST_MOCK_DATA_DIR / "distance_coords_cache.xlsx"
    if not path.exists():
        path = MOCK_DATA_DIR / "distance_coords_cache.xlsx"
    if not path.exists():
        return {}
    df = pd.read_excel(path).fillna("")
    cache = {}
    for _, row in df.iterrows():
        address = str(row.get("address", "")).strip()
        if not address:
            continue
        lat = pd.to_numeric(row.get("lat"), errors="coerce")
        lng = pd.to_numeric(row.get("lng"), errors="coerce")
        if pd.notna(lat) and pd.notna(lng):
            cache[address] = {"lat": float(lat), "lng": float(lng)}
    return cache


def load_origin_duration_cache() -> dict[tuple[float, float, float, float], float]:
    path = CONST_MOCK_DATA_DIR / "origin_duration_cache.xlsx"
    if not path.exists():
        path = MOCK_DATA_DIR / "origin_duration_cache.xlsx"
    cache: dict[tuple[float, float, float, float], float] = {}
    if not path.exists():
        return cache
    df = pd.read_excel(path).fillna("")
    for _, row in df.iterrows():
        try:
            key = (
                round(float(row.get("origin_lat")), 6),
                round(float(row.get("origin_lng")), 6),
                round(float(row.get("dest_lat")), 6),
                round(float(row.get("dest_lng")), 6),
            )
            cache[key] = float(row.get("minutes"))
        except Exception:
            continue
    return cache


def mock_coords_from_cache(address: str, _key: str = "") -> tuple[dict[str, float], str]:
    cached = load_mock_coords_cache().get(address)
    if cached:
        return cached, "mock cache"
    return mock_coords(address, _key)


def mock_matrix_cache_path(group_value: Any) -> Path:
    safe_group = re.sub(r'[<>:"/\\|?*]+', "_", str(group_value).strip())
    safe_group = re.sub(r"\s+", "_", safe_group) or "unknown"
    const_path = CONST_MOCK_DATA_DIR / "distance_matrices" / f"05_distance_matrix_group-{safe_group}.xlsx"
    if const_path.exists():
        return const_path
    return MOCK_DATA_DIR / "distance_matrices" / f"05_distance_matrix_group-{safe_group}.xlsx"


def save_real_mock_data(run_dir: Path) -> list[Path]:
    target_dir = CONST_MOCK_DATA_DIR
    target_dir.mkdir(parents=True, exist_ok=True)
    saved_paths: list[Path] = []

    geocode_path = run_dir / "02_geocoding_cache.xlsx"
    if geocode_path.exists():
        path = target_dir / "geocoding_cache.xlsx"
        pd.read_excel(geocode_path).to_excel(path, index=False)
        saved_paths.append(path)

    coords_path = run_dir / "05_distance_coords_cache.xlsx"
    if coords_path.exists():
        path = target_dir / "distance_coords_cache.xlsx"
        pd.read_excel(coords_path).to_excel(path, index=False)
        saved_paths.append(path)

    matrix_dir = target_dir / "distance_matrices"
    matrix_dir.mkdir(parents=True, exist_ok=True)
    for matrix_path in sorted(run_dir.glob("05_distance_matrix_group-*.xlsx")):
        target = matrix_dir / matrix_path.name
        pd.read_excel(matrix_path, index_col=0).to_excel(target)
        saved_paths.append(target)

    origin_source = run_dir / "06_origin_duration_cache.xlsx"
    if origin_source.exists():
        origin_target = target_dir / "origin_duration_cache.xlsx"
        pd.read_excel(origin_source).to_excel(origin_target, index=False)
        saved_paths.append(origin_target)

    return saved_paths


def clean_cell(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def split_street_and_number(address_str: Any) -> tuple[str, str]:
    address = clean_cell(address_str)
    if not address:
        return "", ""
    address = re.split(r"\s*/\s*|\s*\(", address, maxsplit=1)[0].strip()
    match = re.search(r"\d", address)
    if not match:
        return address, ""
    first_digit_idx = match.start()
    if first_digit_idx == 0:
        end_of_num = re.search(r"^\d+[\w\-/]*", address)
        number = end_of_num.group(0).strip() if end_of_num else ""
        return address[len(number) :].strip().rstrip(","), number
    return address[:first_digit_idx].strip().rstrip(","), address[first_digit_idx:].strip()


def get_lat_lng(address: str, key: str) -> tuple[Any, Any]:
    response = requests.get(
        "https://maps.googleapis.com/maps/api/geocode/json",
        params={"address": address, "key": key},
        timeout=30,
    ).json()
    if response.get("status") == "OK":
        location = response["results"][0]["geometry"]["location"]
        return location["lat"], location["lng"]
    return None, None


def formatted_address_matches_city(formatted_address: str, city: str) -> bool:
    formatted = clean_cell(formatted_address)
    city = clean_cell(city)
    if not city:
        return True
    if city in formatted:
        return True
    aliases = {
        "מודיעין": ["מודיעין-מכבים-רעות", "מודיעין מכבים רעות"],
        "תל אביב": ["תל אביב-יפו", "תל אביב יפו"],
    }
    return any(alias in formatted for alias in aliases.get(city, []))


def geocode_precise_address(address: str, key: str, city: str = "") -> tuple[Any, Any, bool, str]:
    response = requests.get(
        "https://maps.googleapis.com/maps/api/geocode/json",
        params={
            "address": address,
            "components": "country:IL",
            "language": "he",
            "key": key,
        },
        timeout=30,
    ).json()
    if response.get("status") != "OK" or not response.get("results"):
        return None, None, False, f"google could not accurately find this address: {response.get('status', 'no result')}"

    precise_types = {"street_address", "premise", "subpremise"}
    wrong_city_precise = False
    for result in response["results"]:
        result_types = set(result.get("types", []))
        if not result_types.intersection(precise_types):
            continue
        formatted = result.get("formatted_address", "")
        if not formatted_address_matches_city(formatted, city):
            wrong_city_precise = True
            continue

        location = result.get("geometry", {}).get("location", {})
        lat = location.get("lat")
        lng = location.get("lng")
        if lat is None or lng is None:
            return None, None, False, "google could not accurately find this address: missing coordinates"
        return lat, lng, True, f"google precise match: {formatted}"

    if wrong_city_precise:
        return None, None, False, "google found a precise address, but not in the requested city"
    return None, None, False, "google could not accurately find this address"


def record_geocode_failures(run_dir: Path, failed_rows: list[dict[str, Any]]) -> Path | None:
    if not failed_rows:
        return None

    failure_path = run_dir / "02a_failed_geocoding_addresses.xlsx"
    failed_df = pd.DataFrame(failed_rows)
    failed_df.to_excel(failure_path, index=False)

    cumulative_path = run_dir / "01a_failed_addresses.xlsx"
    if cumulative_path.exists():
        existing = pd.read_excel(cumulative_path).fillna("")
        cumulative = pd.concat([existing, failed_df], ignore_index=True, sort=False)
    else:
        cumulative = failed_df
    cumulative.to_excel(cumulative_path, index=False)
    return failure_path


def calculate_haversine(lat1: Any, lon1: Any, lat2: Any, lon2: Any) -> float:
    if pd.isna(lat1) or pd.isna(lon1) or pd.isna(lat2) or pd.isna(lon2):
        return float("inf")
    radius_m = 6371000
    phi1, phi2 = np.radians(float(lat1)), np.radians(float(lat2))
    dphi = np.radians(float(lat2) - float(lat1))
    dlambda = np.radians(float(lon2) - float(lon1))
    a = np.sin(dphi / 2) ** 2 + np.cos(phi1) * np.cos(phi2) * np.sin(dlambda / 2) ** 2
    return float(2 * radius_m * np.arctan2(np.sqrt(a), np.sqrt(1 - a)))


def normalize_raw_col(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value).strip().lower().replace("_", " "))


def find_raw_col(df: pd.DataFrame, *names: str) -> str | None:
    normalized = {normalize_raw_col(col): col for col in df.columns}
    for name in names:
        col = normalized.get(normalize_raw_col(name))
        if col is not None:
            return col
    return None


def parse_planning_date(value: Any) -> pd.Timestamp:
    text = clean_cell(value)
    if not text or text.lower() == "today":
        return pd.Timestamp.today().normalize()
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        parsed = pd.to_datetime(text, errors="coerce", dayfirst=False)
    if pd.isna(parsed):
        raise RuntimeError(f"Could not parse planning date: {value}")
    return pd.Timestamp(parsed).normalize()


def parse_required_delivery_date(value: Any) -> pd.Timestamp | None:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return None
    if isinstance(value, (pd.Timestamp, datetime)):
        parsed = pd.Timestamp(value).normalize()
        if parsed.month <= 12 and parsed.day <= 12:
            return pd.Timestamp(year=parsed.year, month=parsed.day, day=parsed.month)
        return parsed
    text = clean_cell(value)
    if not text:
        return None
    for date_format in ("%d/%m/%Y", "%d/%m/%y"):
        parsed = pd.to_datetime(text, format=date_format, errors="coerce")
        if pd.notna(parsed):
            return pd.Timestamp(parsed).normalize()
    if re.fullmatch(r"\d{4}-\d{1,2}-\d{1,2}(?:\s+\d{1,2}:\d{2}:\d{2})?", text):
        parsed = pd.to_datetime(text, errors="coerce")
        if pd.notna(parsed):
            return pd.Timestamp(parsed).normalize()
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        parsed = pd.to_datetime(text, errors="coerce", dayfirst=True)
    if pd.isna(parsed):
        return None
    return pd.Timestamp(parsed).normalize()


def format_required_delivery_date(value: Any) -> str:
    due_date = parse_required_delivery_date(value)
    if due_date is None:
        return clean_cell(value)
    return f"{due_date.day}/{due_date.month}/{due_date.year}"


def delivery_priority_for_date(due_date: pd.Timestamp | None, args: argparse.Namespace) -> int:
    if due_date is None:
        return int(constraint_value(args, "no_due_date_priority", 999999))
    planning_date = parse_planning_date(constraint_value(args, "planning_date", "today"))
    days_until_due = int((due_date - planning_date).days)
    if days_until_due < 0:
        return int(constraint_value(args, "overdue_priority", 0))
    if days_until_due == 0:
        return 1
    return 2


def delivery_priority(value: Any, args: argparse.Namespace) -> int:
    return delivery_priority_for_date(parse_required_delivery_date(value), args)


def package_count_from_row(row: pd.Series) -> int:
    value = pd.to_numeric(row.get("total_orders_in_cluster", ""), errors="coerce")
    if pd.notna(value) and int(value) > 0:
        return int(value)
    value = pd.to_numeric(row.get("source_order_count", ""), errors="coerce")
    if pd.notna(value) and int(value) > 0:
        return int(value)
    value = pd.to_numeric(row.get("package_count", ""), errors="coerce")
    if pd.notna(value) and int(value) > 0:
        return int(value)
    return 1


def select_orders_for_current_run(
    df: pd.DataFrame, run_dir: Path, args: argparse.Namespace
) -> tuple[pd.DataFrame, pd.DataFrame, Path, Path]:
    due_col = str(constraint_value(args, "due_date_column", "required_delivery_date"))
    if due_col not in df.columns:
        df[due_col] = ""
    parsed_due_dates = df[due_col].apply(parse_required_delivery_date)
    df["required_delivery_date"] = df[due_col].apply(format_required_delivery_date)
    df["delivery_priority"] = parsed_due_dates.apply(lambda value: delivery_priority_for_date(value, args))
    df["package_count"] = 1
    df["_selection_order"] = range(len(df))
    df["_parsed_required_delivery_date"] = parsed_due_dates
    capacity = (
        int(constraint_value(args, "drivers", 3))
        * int(constraint_value(args, "max_packages_per_driver", 100))
        * int(constraint_value(args, "selection_capacity_multiplier", 2))
    )
    sorted_df = df.sort_values(
        by=["delivery_priority", "_parsed_required_delivery_date", "_selection_order"],
        kind="stable",
    ).reset_index(drop=True)
    selected = sorted_df.head(capacity).drop(columns=["_selection_order", "_parsed_required_delivery_date"])
    deferred = sorted_df.iloc[capacity:].drop(columns=["_selection_order", "_parsed_required_delivery_date"])
    selected_path = run_dir / "02b_selected_orders_for_run.xlsx"
    deferred_path = run_dir / "02c_deferred_orders_next_run.xlsx"
    selected.to_excel(selected_path, index=False)
    deferred.to_excel(deferred_path, index=False)
    return selected, deferred, selected_path, deferred_path


def stage_1_cleanup(input_path: Path, run_dir: Path, args: argparse.Namespace) -> tuple[Path, StageResult]:
    cleanup = import_script(ROOT / "data.cleanup.py", "routecraft_data_cleanup")
    df = pd.read_excel(input_path).fillna("")
    city_col = find_raw_col(df, "City", "site name", "site_name")
    street_col = find_raw_col(df, "Street_Name", "street name", "ship to street 1", "ship_to_street1")
    number_col = find_raw_col(df, "House_Number", "house number", "building number", "ship to street 2", "ship_to_street2")
    if city_col and street_col and number_col:
        rows = []
        failed_rows = []
        for source_index, row in df.iterrows():
            city = clean_cell(row.get(city_col, ""))
            street = clean_cell(row.get(street_col, ""))
            house_number = clean_cell(row.get(number_col, "")).replace(".0", "")
            valid, status = cleanup.validate_components(city, street, house_number)
            result = row.to_dict()
            result.update(
                {
                    "source_row": source_index + 2,
                    "City": city,
                    "Street_Name": street,
                    "House_Number": house_number,
                    "cleanup_status": status,
                }
            )
            if valid:
                rows.append(result)
            else:
                failed_rows.append(result)
        good_df = pd.DataFrame(rows)
        failed_df = pd.DataFrame(failed_rows)
        shaped_df = good_df.copy()
        lead_cols = ["City", "Street_Name", "House_Number"]
        shaped_df = shaped_df[lead_cols + [col for col in shaped_df.columns if col not in lead_cols]]

        failed_path = run_dir / "01a_failed_addresses.xlsx"
        shaped_path = run_dir / "01b_addresses_for_geocoding.xlsx"
        original_path = run_dir / "01c_good_orders_original_format.xlsx"
        failed_df.to_excel(failed_path, index=False)
        shaped_df.to_excel(shaped_path, index=False)
        good_df.to_excel(original_path, index=False)
    else:
        paths = cleanup.clean_raw_orders(input_path, run_dir)
        failed_path = paths["failed"]
        shaped_path = paths["good_addresses"]
        original_path = paths["good_original"]
        due_col = find_raw_col(df, "required_delivery_date", "required delivery date")
        if due_col:
            original_df = pd.read_excel(original_path).fillna("")
            source_due = df[[due_col]].copy()
            source_due["source_row"] = source_due.index + 2
            merged = original_df.merge(source_due, on="source_row", how="left")
            merged = merged.rename(columns={due_col: "required_delivery_date"})
            merged.to_excel(original_path, index=False)
            shaped_df = pd.read_excel(shaped_path).fillna("")
            extra_cols = ["source_row", "required_delivery_date"]
            shaped_extra = merged[["City", "Street_Name", "House_Number", *extra_cols]]
            shaped_df = shaped_df.merge(shaped_extra, on=["City", "Street_Name", "House_Number"], how="left")
            shaped_df.to_excel(shaped_path, index=False)

    return shaped_path, StageResult(
        1,
        "Clean raw order addresses",
        "completed",
        [failed_path, shaped_path, original_path],
        [
            "Deleted required delivery date, grouped by city, merged ship_to_street1/2, "
            "and wrote failed, flow-ready, and original-format good-address files."
        ],
    )


def stage_2_geocode(input_path: Path, run_dir: Path, args: argparse.Namespace) -> tuple[Path, StageResult]:
    google_key = "" if args.mock_google else get_google_key(args)
    geocode_cache = load_mock_geocode_cache()
    df = pd.read_excel(input_path)
    if df.shape[1] < 3:
        raise RuntimeError("Stage 2 input must have at least 3 columns: city, street, house number.")

    city_col, street_col, number_col = df.columns[:3]
    rows = []
    failed_rows = []
    geocode_cache_rows = []
    for _, row in df.iterrows():
        city = clean_cell(row[city_col])
        street = clean_cell(row[street_col])
        house_number = clean_cell(row[number_col])
        query = f"{street} {house_number}, {city}".strip()
        cached = geocode_cache.get((query, city))
        if cached:
            lat = cached.get("lat")
            lng = cached.get("lng")
            is_precise = bool(cached.get("is_precise", True))
            status = str(cached.get("status", "cached precise match"))
        elif args.mock_google:
                lat, lng = mock_lat_lng(query)
                is_precise = True
                status = "mock precise match"
        else:
            lat, lng, is_precise, status = geocode_precise_address(query, google_key, city)

        geocode_cache_rows.append(
            {
                "query": query,
                "city": city,
                "street": street,
                "house_number": house_number,
                "lat": lat,
                "lng": lng,
                "is_precise": is_precise,
                "status": status,
            }
        )

        if not is_precise:
            failed_rows.append(
                {
                    "City": city,
                    "Street_Name": street,
                    "House_Number": house_number,
                    "merged_address": query,
                    "cleanup_status": status,
                }
            )
            continue

        rows.append(
            {
                "City": city,
                "Street_Name": street,
                "House_Number": house_number,
                "LAT": lat,
                "LNG": lng,
                "geocode_status": status,
                **{
                    str(col): row.get(col)
                    for col in df.columns[3:]
                    if str(col) not in {"LAT", "LNG", "geocode_status"}
                },
            }
        )

    output_path = run_dir / "02_geocoded_addresses.xlsx"
    if not rows:
        record_geocode_failures(run_dir, failed_rows)
        raise RuntimeError("Stage 2 did not find any addresses with precise Google geocoding matches.")

    geocoded_df = pd.DataFrame(rows)
    selected_df, deferred_df, selected_path, deferred_path = select_orders_for_current_run(geocoded_df, run_dir, args)
    geocoded_df.to_excel(output_path, index=False)
    if geocode_cache_rows:
        pd.DataFrame(geocode_cache_rows).to_excel(run_dir / "02_geocoding_cache.xlsx", index=False)
    failure_path = record_geocode_failures(run_dir, failed_rows)
    output_files = [output_path, selected_path, deferred_path, run_dir / "02_geocoding_cache.xlsx"] + ([failure_path] if failure_path else [])
    notes = []
    notes.append(f"Selected {len(selected_df)} orders for this run; deferred {len(deferred_df)} orders for next run.")
    if failed_rows:
        notes.append(
            f"{len(failed_rows)} addresses were moved to the error file because Google could not accurately find them."
        )
    return selected_path, StageResult(2, "Geocode addresses and select loadable orders", "completed", output_files, notes)


def stage_3_cluster(input_path: Path, run_dir: Path, args: argparse.Namespace) -> tuple[Path, StageResult]:
    df = pd.read_excel(input_path)
    required = {"City", "Street_Name", "House_Number", "LAT", "LNG"}
    missing = required.difference(df.columns)
    if missing:
        raise RuntimeError(f"Stage 3 input is missing columns: {', '.join(sorted(missing))}")

    df["House_Number"] = pd.to_numeric(df["House_Number"], errors="coerce")
    df = df.dropna(subset=["House_Number"]).reset_index(drop=True)
    if df.empty:
        raise RuntimeError("Stage 3 has no rows with numeric House_Number.")

    df["House_Number"] = df["House_Number"].astype(int)
    df = df.sort_values(by=["City", "Street_Name", "House_Number"]).reset_index(drop=True)

    clusters: list[list[dict[str, Any]]] = []
    for _city, city_df in df.groupby("City", sort=False):
        city_df = city_df.sort_values(by=["Street_Name", "House_Number"]).reset_index(drop=True)
        anchor = city_df.iloc[0]
        current = [anchor.to_dict()]
        for i in range(1, len(city_df)):
            row = city_df.iloc[i]
            distance = calculate_haversine(anchor["LAT"], anchor["LNG"], row["LAT"], row["LNG"])
            close_distance = float(constraint_value(args, "close_building_distance_meters", args.cluster_threshold_meters))
            if row["Street_Name"] == anchor["Street_Name"] and distance <= close_distance:
                current.append(row.to_dict())
            else:
                clusters.append(current)
                anchor = row
                current = [anchor.to_dict()]
        clusters.append(current)

    reps = []
    for cluster_id, cluster_rows in enumerate(clusters):
        rep = cluster_rows[0].copy()
        rep["cluster_id"] = cluster_id
        rep["total_orders_in_cluster"] = len(cluster_rows)
        rep["detailed_addresses"] = ", ".join(
            f"{item['Street_Name']} {int(item['House_Number'])}" for item in cluster_rows
        )
        priority_values = pd.to_numeric(
            pd.Series([item.get("delivery_priority", 999999) for item in cluster_rows]),
            errors="coerce",
        ).dropna()
        rep["delivery_priority"] = int(priority_values.min()) if not priority_values.empty else 999999
        rep["required_delivery_date"] = ", ".join(
            clean_cell(item.get("required_delivery_date", ""))
            for item in cluster_rows
            if clean_cell(item.get("required_delivery_date", ""))
        )
        rep["source_order_count"] = len(cluster_rows)
        reps.append(rep)

    output_path = run_dir / "03_nearby_address_clusters.xlsx"
    pd.DataFrame(reps).to_excel(output_path, index=False)
    return output_path, StageResult(3, "Cluster nearby addresses", "completed", [output_path])


def stage_4_group(input_path: Path, run_dir: Path, args: argparse.Namespace) -> tuple[Path, StageResult]:
    df = pd.read_excel(input_path)
    required = {"City", "LAT", "LNG"}
    missing = required.difference(df.columns)
    if missing:
        raise RuntimeError(f"Stage 4 input is missing columns: {', '.join(sorted(missing))}")

    original_path = without_local_ortools_shadow()
    try:
        from k_means_constrained import KMeansConstrained
    finally:
        sys.path = original_path
    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    df["cluster_group"] = pd.NA
    global_cluster_id = 0
    total_clusters = 0
    notes = []
    for city, city_df in df.groupby("City"):
        city_indices = city_df.index
        coords = city_df[["LAT", "LNG"]].values
        num_addresses = len(city_df)
        if num_addresses <= MAX_GROUP_SIZE:
            df.loc[city_indices, "cluster_group"] = global_cluster_id
            notes.append(f"{city}: {num_addresses} rows fit in group {global_cluster_id}.")
            global_cluster_id += 1
            total_clusters += 1
            continue

        n_clusters = int(np.ceil(num_addresses / MAX_GROUP_SIZE))
        clf = KMeansConstrained(
            n_clusters=n_clusters,
            size_min=1,
            size_max=MAX_GROUP_SIZE,
            n_init=50,
            max_iter=500,
            random_state=42,
        )
        local_labels = clf.fit_predict(coords)
        df.loc[city_indices, "cluster_group"] = local_labels + global_cluster_id
        notes.append(f"{city}: created groups {global_cluster_id}-{global_cluster_id + n_clusters - 1}.")
        global_cluster_id += n_clusters
        total_clusters += n_clusters

    df["cluster_group"] = df["cluster_group"].astype(int)

    output_path = run_dir / "04_clustered_delivery_groups.xlsx"
    df.to_excel(output_path, index=False)

    map_paths = []
    for city, city_df in df.groupby("City"):
        city_coords = city_df[["LAT", "LNG"]].values
        city_clusters = city_df["cluster_group"]
        fig, ax = plt.subplots(figsize=(10, 7))
        scatter = ax.scatter(city_coords[:, 1], city_coords[:, 0], c=city_clusters, cmap="viridis", s=50)
        ax.set_title(f"Delivery Clusters - {city} ({city_clusters.nunique()} Groups)", fontsize=14)
        ax.set_xlabel("Longitude (LNG)")
        ax.set_ylabel("Latitude (LAT)")
        ax.grid(True, linestyle="--", alpha=0.6)
        fig.colorbar(scatter, ax=ax, label="Global Group Number")
        fig.tight_layout()
        map_path = run_dir / f"04_cluster_map_{sanitize_filename(str(city))}.png"
        fig.savefig(map_path, dpi=160)
        plt.close(fig)
        map_paths.append(map_path)

    notes.append(f"Created {total_clusters} groups across all cities.")
    return output_path, StageResult(4, "Create delivery groups", "completed", [output_path, *map_paths], notes)


def stage_5_distance_matrices(input_path: Path, run_dir: Path, args: argparse.Namespace) -> tuple[list[Path], StageResult]:
    distance = import_script(ROOT / "distance.matrix.py", "routecraft_distance_matrix")
    if args.mock_google:
        distance.API_KEY = "mock-google"
        distance.get_coords = mock_coords_from_cache
        distance.build_matrix_for_addresses = mock_distance_matrix
    else:
        distance.API_KEY = get_distance_key(args)

    df = pd.read_excel(input_path)
    df.columns = df.columns.str.strip()
    if "cluster_group" not in df.columns:
        raise RuntimeError("Stage 5 input must include a cluster_group column.")

    group_col = "cluster_group"
    grouped_frames = {
        group_value: group_df.reset_index(drop=True)
        for group_value, group_df in df.groupby(group_col, dropna=True)
    }
    if not grouped_frames:
        raise RuntimeError("Stage 5 found no groups in column I.")

    all_addresses = []
    for group_df in grouped_frames.values():
        all_addresses.extend(distance.build_full_addresses(group_df))
    unique_addresses = list(dict.fromkeys(all_addresses))

    coords_dict = {}
    coords_rows = []
    coords_cache = load_mock_coords_cache()
    for address in unique_addresses:
        cached_coords = coords_cache.get(address)
        if cached_coords:
            coords, _status = cached_coords, "cached"
        else:
            coords, _status = distance.get_coords(address, distance.API_KEY)
        coords_dict[address] = coords
        coords_rows.append(
            {
                "address": address,
                "lat": coords.get("lat") if coords else None,
                "lng": coords.get("lng") if coords else None,
                "status": _status,
            }
        )

    coords_cache_path = run_dir / "05_distance_coords_cache.xlsx"
    pd.DataFrame(coords_rows).to_excel(coords_cache_path, index=False)

    matrix_paths = []
    for group_value, group_df in grouped_frames.items():
        full_addresses = distance.build_full_addresses(group_df)
        safe_group = distance.sanitize_group_value(group_value)
        matrix_path = run_dir / f"05_distance_matrix_group-{safe_group}.xlsx"
        cached_matrix_path = mock_matrix_cache_path(group_value)
        if cached_matrix_path.exists():
            matrix_df = pd.read_excel(cached_matrix_path, index_col=0)
            matrix_df = matrix_df.reindex(index=full_addresses, columns=full_addresses)
            if matrix_df.isna().any().any():
                matrix_df = distance.build_matrix_for_addresses(full_addresses, coords_dict)
        else:
            matrix_df = distance.build_matrix_for_addresses(full_addresses, coords_dict)
        matrix_df.to_excel(matrix_path)
        matrix_paths.append(matrix_path)

    return matrix_paths, StageResult(5, "Build distance matrices", "completed", [coords_cache_path, *matrix_paths])


def matrix_label(row: pd.Series) -> str:
    street = clean_cell(row.get("Street_Name", ""))
    house = clean_cell(row.get("House_Number", "")).replace(".0", "")
    city = clean_cell(row.get("City", ""))
    return ", ".join(part for part in [" ".join(part for part in [street, house] if part), city] if part)


def route_parts_for_row(row: pd.Series) -> list[str]:
    detailed = clean_cell(row.get("detailed_addresses", ""))
    if detailed:
        return [part.strip() for part in detailed.split(",") if part.strip()]
    label = matrix_label(row).split(",")[0]
    return [label] if label else []


def estimate_drive_minutes_between(lat1: Any, lng1: Any, lat2: Any, lng2: Any) -> int:
    meters = calculate_haversine(lat1, lng1, lat2, lng2)
    if not np.isfinite(meters):
        return 999
    minutes = (meters * 1.35) / (42 * 1000 / 60)
    return max(1, int(round(minutes)))


def drive_duration_minutes(
    origin_lat: Any,
    origin_lng: Any,
    dest_lat: Any,
    dest_lng: Any,
    args: argparse.Namespace,
    api_key: str | None,
) -> int:
    try:
        cache_key = (round(float(origin_lat), 6), round(float(origin_lng), 6), round(float(dest_lat), 6), round(float(dest_lng), 6))
        cached_minutes = load_origin_duration_cache().get(cache_key)
        if cached_minutes is not None:
            return int(round(cached_minutes))
    except Exception:
        pass
    if args.mock_google or not api_key:
        return estimate_drive_minutes_between(origin_lat, origin_lng, dest_lat, dest_lng)
    response = requests.get(
        "https://maps.googleapis.com/maps/api/distancematrix/json",
        params={
            "origins": f"{origin_lat},{origin_lng}",
            "destinations": f"{dest_lat},{dest_lng}",
            "mode": "driving",
            "departure_time": "now",
            "key": api_key,
        },
        timeout=30,
    )
    response.raise_for_status()
    data = response.json()
    if data.get("status") != "OK":
        return 999
    element = data["rows"][0]["elements"][0]
    if element.get("status") != "OK":
        return 999
    seconds = element.get("duration_in_traffic", element.get("duration", {})).get("value")
    return max(0, int(round((seconds or 0) / 60)))


def load_group_matrix(matrix_folder: Path, group_value: Any) -> pd.DataFrame:
    path = matrix_folder / f"05_distance_matrix_group-{sanitize_filename(str(group_value))}.xlsx"
    if not path.exists():
        path = matrix_folder / f"05_distance_matrix_group-{str(group_value).strip()}.xlsx"
    if not path.exists():
        raise FileNotFoundError(f"Distance matrix for group {group_value} not found in {matrix_folder}")
    return pd.read_excel(path, index_col=0)


def add_current_origin_matrix(
    matrix_df: pd.DataFrame,
    group_rows: pd.DataFrame,
    current_lat: Any,
    current_lng: Any,
    args: argparse.Namespace,
    api_key: str | None,
) -> pd.DataFrame:
    labels = [matrix_label(row) for _, row in group_rows.iterrows()]
    matrix_df = matrix_df.reindex(index=labels, columns=labels)
    if matrix_df.isna().any().any():
        matrix_df = matrix_df.fillna(999)
    out = pd.DataFrame(index=["ORIGIN", *labels], columns=["ORIGIN", *labels], dtype=float)
    out.loc["ORIGIN", "ORIGIN"] = 0
    out.loc[labels, "ORIGIN"] = 999
    out.loc[labels, labels] = matrix_df.values
    for label, (_, row) in zip(labels, group_rows.iterrows()):
        out.loc["ORIGIN", label] = drive_duration_minutes(
            current_lat,
            current_lng,
            row.get("LAT"),
            row.get("LNG"),
            args,
            api_key,
        )
    return out


def cluster_key_for_row(row: pd.Series) -> tuple[str, str]:
    return (clean_cell(row.get("_matrix_folder", "")), clean_cell(row.get("cluster_group", "")))


def cluster_key_mask(df: pd.DataFrame, key: tuple[str, str]) -> pd.Series:
    return (
        df["_matrix_folder"].astype(str).map(clean_cell).eq(key[0])
        & df["cluster_group"].astype(str).map(clean_cell).eq(key[1])
    )


def closest_next_row(
    remaining: pd.DataFrame,
    current_lat: Any,
    current_lng: Any,
    active_city: Any,
    pending_cluster_key: tuple[str, str] | None,
) -> tuple[pd.Series | None, Any, tuple[str, str] | None]:
    if remaining.empty:
        return None, active_city, pending_cluster_key
    if pending_cluster_key is not None:
        pending_rows = remaining[cluster_key_mask(remaining, pending_cluster_key)]
        if not pending_rows.empty:
            row = pending_rows.iloc[0]
            return row, row.get("City", active_city), pending_cluster_key
        pending_cluster_key = None

    candidates = remaining.copy()
    candidates["_air_distance"] = candidates.apply(
        lambda row: calculate_haversine(current_lat, current_lng, row.get("LAT"), row.get("LNG")),
        axis=1,
    )
    if active_city not in (None, "") and "City" in candidates.columns:
        city_candidates = candidates[candidates["City"].astype(str).eq(str(active_city))]
        if not city_candidates.empty:
            row = city_candidates.sort_values("_air_distance", kind="stable").iloc[0]
            return row, active_city, None
    row = candidates.sort_values("_air_distance", kind="stable").iloc[0]
    return row, row.get("City", None), None


def remaining_city_has_rows(remaining: pd.DataFrame, city: Any) -> bool:
    if remaining.empty or city in (None, "") or "City" not in remaining.columns:
        return False
    return bool(remaining["City"].astype(str).eq(str(city)).any())


def tsp_route_for_rows(
    deliveries: Any,
    matrix_folder: Path,
    group_value: Any,
    group_rows: pd.DataFrame,
    current_lat: Any,
    current_lng: Any,
    args: argparse.Namespace,
    api_key: str | None,
) -> tuple[list[int], pd.DataFrame, str]:
    matrix_df = load_group_matrix(matrix_folder, group_value)
    modified_matrix = add_current_origin_matrix(
        matrix_df,
        group_rows,
        current_lat,
        current_lng,
        args,
        api_key,
    )
    stdout = io.StringIO()
    stderr = io.StringIO()
    with contextlib.redirect_stdout(stdout), contextlib.redirect_stderr(stderr):
        tsp_route, _ = deliveries.run_tsp_on_matrix(modified_matrix)
    return tsp_route, modified_matrix, stdout.getvalue() + stderr.getvalue()


def choose_load_rows(remaining: pd.DataFrame, max_packages: int) -> pd.DataFrame:
    load_indices = []
    load_packages = 0
    sort_columns = [col for col in ["_matrix_folder", "cluster_group", "City"] if col in remaining.columns]
    ordered = remaining.sort_values(by=sort_columns, kind="stable") if sort_columns else remaining
    for idx, row in ordered.iterrows():
        packages = package_count_from_row(row)
        if packages > max_packages and not load_indices:
            break
        if load_packages + packages > max_packages:
            continue
        load_indices.append(idx)
        load_packages += packages
        if load_packages >= max_packages:
            break
    return remaining.loc[load_indices].copy()


def pop_deferred_batch(deferred_pool: pd.DataFrame, max_packages: int) -> tuple[pd.DataFrame, pd.DataFrame]:
    if deferred_pool.empty:
        return deferred_pool.copy(), deferred_pool.copy()
    pool = deferred_pool.sort_values(
        by=["delivery_priority", "required_delivery_date"],
        kind="stable",
    )
    selected_indices = []
    packages = 0
    for idx, row in pool.iterrows():
        count = int(pd.to_numeric(row.get("package_count", 1), errors="coerce") or 1)
        if packages + count > max_packages and selected_indices:
            break
        if packages + count > max_packages:
            continue
        selected_indices.append(idx)
        packages += count
        if packages >= max_packages:
            break
    batch = deferred_pool.loc[selected_indices].copy()
    leftover = deferred_pool.drop(index=selected_indices).copy()
    return batch, leftover


def prepare_deferred_reload_batch(
    deferred_batch: pd.DataFrame,
    run_dir: Path,
    args: argparse.Namespace,
    driver_id: int,
    load_number: int,
    next_route_id: int,
) -> tuple[pd.DataFrame, int, Path, list[Path]]:
    if deferred_batch.empty:
        return deferred_batch.copy(), next_route_id, run_dir, []
    batch_dir = run_dir / f"06_deferred_driver-{driver_id}_load-{load_number}"
    batch_dir.mkdir(parents=True, exist_ok=True)
    batch_input = batch_dir / "06a_deferred_orders_for_reload.xlsx"
    deferred_batch.to_excel(batch_input, index=False)
    clustered_path, _ = stage_3_cluster(batch_input, batch_dir, args)
    grouped_path, _ = stage_4_group(clustered_path, batch_dir, args)
    stage_5_distance_matrices(grouped_path, batch_dir, args)
    prepared = pd.read_excel(grouped_path).fillna("")
    prepared["_matrix_folder"] = str(batch_dir)
    prepared["_route_id"] = range(next_route_id, next_route_id + len(prepared))
    return prepared, next_route_id + len(prepared), batch_dir, [
        batch_input,
        clustered_path,
        grouped_path,
        *sorted(batch_dir.glob("04_cluster_map_*.png")),
        *sorted(batch_dir.glob("05_distance_matrix_group-*.xlsx")),
        batch_dir / "05_distance_coords_cache.xlsx",
    ]


def stage_6_delivery_plan(
    input_path: Path,
    matrix_folder: Path,
    run_dir: Path,
    args: argparse.Namespace,
) -> tuple[Path, StageResult]:
    deliveries = import_script(ROOT / "deliveries.py", "routecraft_deliveries_const")
    output_path = run_dir / "06_delivery_plan.xlsx"
    api_key = None if args.mock_google else get_distance_key(args)

    df = pd.read_excel(input_path).fillna("")
    if df.empty:
        raise RuntimeError("Stage 6 has no selected orders to route.")
    if "delivery_priority" not in df.columns:
        df["delivery_priority"] = int(constraint_value(args, "no_due_date_priority", 999999))
    df["_route_id"] = range(len(df))
    df["_matrix_folder"] = str(matrix_folder)

    drivers = int(constraint_value(args, "drivers", 3))
    max_packages = int(constraint_value(args, "max_packages_per_driver", 100))
    ideal_shift = float(constraint_value(args, "ideal_shift_minutes", 420))
    max_shift = float(constraint_value(args, "max_shift_minutes", 720))
    service_minutes = float(constraint_value(args, "service_minutes_per_package", 4))

    remaining = df.copy()
    deferred_path = run_dir / "02c_deferred_orders_next_run.xlsx"
    if not deferred_path.exists():
        deferred_path = input_path.parent / "02c_deferred_orders_next_run.xlsx"
    deferred_pool = pd.read_excel(deferred_path).fillna("") if deferred_path.exists() else pd.DataFrame()
    if deferred_path.exists():
        try:
            (run_dir / "02c_deferred_orders_next_run.xlsx").write_bytes(deferred_path.read_bytes())
        except OSError:
            pass
    if not deferred_pool.empty and "delivery_priority" not in deferred_pool.columns:
        due_col = str(constraint_value(args, "due_date_column", "required_delivery_date"))
        deferred_pool["delivery_priority"] = deferred_pool.get(due_col, "").apply(lambda value: delivery_priority(value, args))
    if not deferred_pool.empty and "package_count" not in deferred_pool.columns:
        deferred_pool["package_count"] = 1
    next_route_id = len(df)
    all_driver_results: dict[int, dict[str, Any]] = {}
    log_lines = []
    extra_output_files: list[Path] = []
    pending_cluster_key: tuple[str, str] | None = None
    active_city: Any = None

    log_lines.extend(
        [
            "=" * 60,
            "LOADING DATA",
            "=" * 60,
            f"Loaded {len(df)} selected rows for constrained delivery planning",
            f"Drivers: {drivers}",
            f"Capacity per driver load: {max_packages}",
            f"Ideal shift minutes: {ideal_shift:.0f}",
            f"Max shift minutes: {max_shift:.0f}",
            "",
            "=" * 60,
            "STARTING MULTI-VEHICLE TSP JOURNEY (CONSTRAINT PRODUCTION)"
            if not args.mock_google
            else "STARTING MULTI-VEHICLE TSP JOURNEY (CONSTRAINT MOCK)",
            "=" * 60,
        ]
    )

    for driver_id in range(1, drivers + 1):
        driver_time = 0.0
        driver_packages = 0
        journey = []
        current_lat = WAREHOUSE_LAT
        current_lng = WAREHOUSE_LNG
        current_name = WAREHOUSE_NAME
        load_number = 1
        load_packages = 0

        log_lines.extend(
            [
                "",
                "=" * 60,
                f"STARTING SHIFT FOR DRIVER {driver_id}",
                "=" * 60,
            ]
        )

        while driver_time < max_shift:
            if remaining.empty:
                if driver_time >= ideal_shift or deferred_pool.empty:
                    break
                deferred_batch, deferred_pool = pop_deferred_batch(deferred_pool, max_packages)
                prepared, next_route_id, _batch_dir, batch_files = prepare_deferred_reload_batch(
                    deferred_batch,
                    run_dir,
                    args,
                    driver_id,
                    load_number,
                    next_route_id,
                )
                extra_output_files.extend(path for path in batch_files if path.exists())
                if prepared.empty:
                    break
                remaining = pd.concat([remaining, prepared], ignore_index=True, sort=False)
                log_lines.append(
                    f"Driver {driver_id} pulled {len(deferred_batch)} deferred orders for load {load_number}."
                )

            if load_packages >= max_packages:
                return_to_warehouse = drive_duration_minutes(
                    current_lat,
                    current_lng,
                    WAREHOUSE_LAT,
                    WAREHOUSE_LNG,
                    args,
                    api_key,
                )
                if driver_time >= ideal_shift or driver_time + return_to_warehouse >= max_shift:
                    break
                driver_time += return_to_warehouse
                journey.append(
                    {
                        "load": load_number,
                        "cluster": "RELOAD",
                        "city": "Warehouse",
                        "route_str": f"{current_name} → {WAREHOUSE_NAME}",
                        "travel_time": return_to_warehouse,
                        "delivery_time": 0,
                        "cost": return_to_warehouse,
                        "endpoint": WAREHOUSE_NAME,
                        "packages": 0,
                        "segment_type": "reload",
                    }
                )
                log_lines.append(
                    f"Driver {driver_id} returning to Eshtaol to reload: {return_to_warehouse:.2f} min."
                )
                current_lat = WAREHOUSE_LAT
                current_lng = WAREHOUSE_LNG
                current_name = WAREHOUSE_NAME
                load_number += 1
                load_packages = 0

            closest, active_city, pending_cluster_key = closest_next_row(
                remaining,
                current_lat,
                current_lng,
                active_city,
                pending_cluster_key,
            )
            if closest is None:
                break

            folder_value, group_value = cluster_key_for_row(closest)
            group_rows = remaining[cluster_key_mask(remaining, (folder_value, group_value))].copy()
            if group_rows.empty:
                pending_cluster_key = None
                continue

            log_lines.append("")
            log_lines.append(
                f"Next Target: {matrix_label(closest)} | City: {active_city or 'None'} | "
                f"Cluster: {group_value} | Load {load_number} packages: {load_packages}/{max_packages} | "
                f"Time: {driver_time:.2f}/{max_shift:.0f} min"
            )

            tsp_route, modified_matrix, tsp_log = tsp_route_for_rows(
                deliveries,
                Path(folder_value),
                group_value,
                group_rows,
                current_lat,
                current_lng,
                args,
                api_key,
            )
            if tsp_log:
                log_lines.append(tsp_log)

            prev_idx = tsp_route[0]
            route_names = [current_name]
            travel_time = 0.0
            delivery_time = 0.0
            segment_rows: list[pd.Series] = []
            delivered_ids: list[int] = []
            stop_reason = ""

            for curr_idx in tsp_route[1:]:
                label = modified_matrix.columns[curr_idx]
                match = group_rows[group_rows.apply(lambda row: matrix_label(row) == label, axis=1)]
                match = match[~match["_route_id"].isin(delivered_ids)]
                if match.empty:
                    prev_idx = curr_idx
                    continue
                row = match.iloc[0]
                packages = package_count_from_row(row)
                step_travel = float(modified_matrix.iloc[prev_idx, curr_idx])
                step_delivery = packages * service_minutes
                total_step = step_travel + step_delivery
                if load_packages + packages > max_packages:
                    if load_packages == 0 and packages > max_packages:
                        log_lines.append(
                            f"Warning: {matrix_label(row)} has {packages} packages, above load capacity {max_packages}. "
                            "Delivering it to avoid getting stuck."
                        )
                    else:
                        stop_reason = "capacity"
                        pending_cluster_key = (folder_value, group_value)
                        break
                if driver_time + travel_time + delivery_time + total_step > max_shift:
                    if not segment_rows and driver_time == 0:
                        log_lines.append(
                            f"Warning: {matrix_label(row)} exceeds max shift alone. Delivering it to avoid getting stuck."
                        )
                    else:
                        stop_reason = "max_shift"
                        pending_cluster_key = (folder_value, group_value)
                        break

                travel_time += step_travel
                delivery_time += step_delivery
                route_parts = route_parts_for_row(row)
                route_names.extend(route_parts)
                segment_rows.append(row)
                delivered_ids.append(int(row["_route_id"]))
                driver_packages += packages
                load_packages += packages
                current_lat = row.get("LAT")
                current_lng = row.get("LNG")
                current_name = route_parts[-1] if route_parts else label
                prev_idx = curr_idx

            if not segment_rows:
                if stop_reason == "capacity" and load_packages > 0:
                    load_packages = max_packages
                    continue
                break

            driver_time += travel_time + delivery_time
            journey.append(
                {
                    "load": load_number,
                    "cluster": group_value,
                    "city": segment_rows[-1].get("City", ""),
                    "route_str": " → ".join(route_names),
                    "travel_time": travel_time,
                    "delivery_time": delivery_time,
                    "cost": travel_time + delivery_time,
                    "endpoint": current_name,
                    "packages": sum(package_count_from_row(row) for row in segment_rows),
                    "segment_type": "delivery",
                }
            )
            remaining = remaining[~remaining["_route_id"].isin(delivered_ids)].copy()
            log_lines.append(f"Route: {' → '.join(route_names)}")
            log_lines.append(
                f"Travel: {travel_time:.2f} min | Delivery: {delivery_time:.2f} min | "
                f"Shift time: {driver_time:.2f} min | Load packages: {load_packages}/{max_packages}"
            )

            if not remaining[cluster_key_mask(remaining, (folder_value, group_value))].empty:
                pending_cluster_key = (folder_value, group_value)
            else:
                pending_cluster_key = None
                if not remaining_city_has_rows(remaining, active_city):
                    log_lines.append(f"City '{active_city}' is complete. Searching globally for next city.")
                    active_city = None

            if stop_reason == "max_shift":
                log_lines.append(
                    f"Driver {driver_id} reached max shift limit. Cluster {group_value} will continue with the next driver."
                )
                break
            if stop_reason == "capacity":
                load_packages = max_packages

            if driver_time >= ideal_shift:
                log_lines.append(
                    f"Driver {driver_id} reached/passed ideal shift ({driver_time:.2f} >= {ideal_shift:.0f}). Ending shift."
                )
                break

        all_driver_results[driver_id] = {
            "journey": journey,
            "time": driver_time,
            "addresses": driver_packages,
            "shift_type": "Regular" if driver_time <= ideal_shift else "Long",
        }
        log_lines.extend(
            [
                "",
                f"SHIFT COMPLETE - DRIVER {driver_id}",
                f"Time worked: {driver_time:.2f} mins ({driver_time / 60:.2f} hours)",
                f"Addresses delivered: {driver_packages}",
            ]
        )
        if remaining.empty and deferred_pool.empty:
            break

    leftover_path = run_dir / "06_leftover_orders_next_run.xlsx"
    leftover = pd.concat(
        [
            remaining.drop(columns=["_route_id", "_matrix_folder"], errors="ignore"),
            deferred_pool.drop(columns=["_route_id", "_matrix_folder"], errors="ignore"),
        ],
        ignore_index=True,
        sort=False,
    )
    leftover.to_excel(leftover_path, index=False)

    summary_rows = []
    detail_rows = []
    for driver_id, data in all_driver_results.items():
        summary_rows.append(
            {
                "Driver": f"Driver {driver_id}",
                "Configured Drivers": drivers,
                "Capacity Per Driver": max_packages,
                "Addresses Delivered": data["addresses"],
                "Clusters Visited": len([step for step in data["journey"] if step.get("segment_type") == "delivery"]),
                "Total Time (min)": f"{data['time']:.2f}",
                "Total Time (hours)": f"{data['time'] / 60:.2f}",
                "Shift Type": data["shift_type"],
                "Mode": "CONSTRAINT MOCK" if args.mock_google else "CONSTRAINT PRODUCTION",
            }
        )
        cumulative = 0.0
        for step_index, step in enumerate(data["journey"], 1):
            cumulative += float(step["cost"])
            detail_rows.append(
                {
                    "Driver": f"Driver {driver_id}",
                    "Step": step_index,
                    "Load": step.get("load", ""),
                    "Segment Type": step.get("segment_type", ""),
                    "City": step.get("city", ""),
                    "Cluster": step.get("cluster", ""),
                    "Packages": step.get("packages", 0),
                    "Total Time (min)": f"{step['cost']:.2f}",
                    "Travel Time": f"{step['travel_time']:.2f}",
                    "Delivery Time": f"{step['delivery_time']:.2f}",
                    "Route Path": step["route_str"],
                    "Endpoint": step["endpoint"],
                    "Shift Time So Far": f"{cumulative:.2f}",
                }
            )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        pd.DataFrame(summary_rows).to_excel(writer, sheet_name="Summary", index=False)
        pd.DataFrame(detail_rows or [{"Info": "No routes delivered"}]).to_excel(
            writer, sheet_name="Detailed Routes", index=False
        )
        pd.DataFrame(
            [
                {
                    "Leftover Orders": len(leftover),
                    "Drivers": drivers,
                    "Max Packages Per Driver": max_packages,
                    "Ideal Shift Minutes": ideal_shift,
                    "Max Shift Minutes": max_shift,
                }
            ]
        ).to_excel(writer, sheet_name="Run Constraints", index=False)

    log_path = run_dir / "06_delivery_plan.log"
    mode_label = "CONSTRAINT MOCK" if args.mock_google else "CONSTRAINT PRODUCTION"
    log_lines.extend(["", "=" * 60, f"ALL DELIVERIES COMPLETE ({mode_label})", "=" * 60, "", "SUMMARY:"])
    for driver_id, data in all_driver_results.items():
        log_lines.append(
            f"Driver {driver_id}: {data['addresses']} addresses | "
            f"{data['time']:.2f} min ({data['time'] / 60:.2f} hours)"
        )
    log_lines.extend(["", "=" * 60, "FULL ROUTES BY DRIVER", "=" * 60])
    for driver_id, data in all_driver_results.items():
        full_route_nodes = []
        for step in data["journey"]:
            parts = split_route_path(step.get("route_str", ""))
            if not full_route_nodes:
                full_route_nodes.extend(parts)
            elif full_route_nodes[-1] == parts[0]:
                full_route_nodes.extend(parts[1:])
            else:
                full_route_nodes.extend(parts)
        log_lines.append("")
        log_lines.append(f"Driver {driver_id} Route:")
        log_lines.append(" → ".join(full_route_nodes))
    log_lines.append(f"Leftover orders for next run: {len(leftover)}")
    log_path.write_text("\n".join(log_lines), encoding="utf-8")
    html_path = render_delivery_plan_html(output_path, run_dir / "06_delivery_plan.html")
    output_files = [output_path, html_path, log_path, leftover_path, *extra_output_files]
    return output_path, StageResult(
        6,
        "Plan constrained multi-load deliveries",
        "completed",
        output_files,
        [f"{len(leftover)} orders were left for the next run."],
    )


def render_delivery_plan_html(workbook_path: Path, output_path: Path) -> Path:
    summary_df = pd.read_excel(workbook_path, sheet_name="Summary")
    detail_df = pd.read_excel(workbook_path, sheet_name="Detailed Routes")
    order_details = load_order_details(output_path.parent / "01c_good_orders_original_format.xlsx")

    summary_cards = []
    if not summary_df.empty:
        totals = {
            "Drivers": summary_df["Configured Drivers"].iloc[0] if "Configured Drivers" in summary_df.columns else len(summary_df),
            "Capacity / Driver": summary_df["Capacity Per Driver"].iloc[0] if "Capacity Per Driver" in summary_df.columns else "",
            "Addresses": int(pd.to_numeric(summary_df["Addresses Delivered"], errors="coerce").fillna(0).sum()),
            "Clusters": int(pd.to_numeric(summary_df["Clusters Visited"], errors="coerce").fillna(0).sum()),
            "Max Shift": f"{pd.to_numeric(summary_df['Total Time (hours)'], errors='coerce').max():.2f} h",
        }
        for label, value in totals.items():
            summary_cards.append(
                f"<div class=\"metric\"><span>{html.escape(label)}</span><strong>{html.escape(str(value))}</strong></div>"
            )

    driver_sections = []
    for driver, driver_df in detail_df.groupby("Driver", sort=False):
        steps = []
        full_route_nodes: list[str] = []
        for _, row in driver_df.iterrows():
            route_nodes = split_route_path(row.get("Route Path", ""))
            for node in route_nodes:
                if not full_route_nodes or full_route_nodes[-1] != node:
                    full_route_nodes.append(node)
            route_html = "".join(
                render_route_node(node, str(row.get("City", "")), order_details)
                for node in route_nodes
            )
            steps.append(
                "<article class=\"step\">"
                "<div class=\"step-head\">"
                f"<span>Step {html.escape(str(row.get('Step', '')))}</span>"
                f"<strong>{html.escape(str(row.get('City', '')))} / Group {html.escape(str(row.get('Cluster', '')))}</strong>"
                "</div>"
                "<div class=\"times\">"
                f"<span>Total {html.escape(str(row.get('Total Time (min)', '')))} min</span>"
                f"<span>Travel {html.escape(str(row.get('Travel Time', '')))} min</span>"
                f"<span>Delivery {html.escape(str(row.get('Delivery Time', '')))} min</span>"
                f"<span>Shift {html.escape(str(row.get('Shift Time So Far', '')))} min</span>"
                "</div>"
                f"<ol class=\"route\">{route_html}</ol>"
                "</article>"
            )
        full_route_html = "".join(
            f"<li><span>{index}</span><strong>{html.escape(node)}</strong></li>"
            for index, node in enumerate(full_route_nodes, 1)
        )
        driver_sections.append(
            "<section class=\"driver\">"
            f"<h2>{html.escape(str(driver))}</h2>"
            "<article class=\"route-overview\">"
            "<div class=\"step-head\">"
            "<span>Full route</span>"
            "<strong>Start to finish</strong>"
            "</div>"
            f"<ol class=\"full-route\">{full_route_html}</ol>"
            "</article>"
            f"{''.join(steps)}"
            "</section>"
        )

    summary_table = summary_df.to_html(index=False, escape=True, classes="data-table")

    output_path.write_text(
        f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Routecraft Delivery Plan</title>
  <style>
    body {{ margin: 0; font-family: Arial, sans-serif; background: #f5f7f8; color: #1f2933; }}
    header {{ padding: 28px 36px; background: #17324d; color: white; }}
    main {{ padding: 28px 36px 48px; max-width: 1280px; margin: 0 auto; }}
    h1, h2 {{ margin: 0; }}
    .subtitle {{ margin-top: 8px; color: #d7e2ea; }}
    .metrics {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 12px; margin: 22px 0; }}
    .metric {{ background: white; border: 1px solid #d9e0e7; padding: 14px 16px; border-radius: 8px; }}
    .metric span {{ display: block; color: #65717d; font-size: 13px; }}
    .metric strong {{ display: block; margin-top: 6px; font-size: 26px; }}
    .data-table {{ width: 100%; border-collapse: collapse; background: white; border: 1px solid #d9e0e7; margin-bottom: 24px; }}
    th, td {{ padding: 10px 12px; border-bottom: 1px solid #e7ebef; text-align: left; vertical-align: top; }}
    th {{ background: #edf2f6; font-size: 12px; text-transform: uppercase; letter-spacing: 0; }}
    .driver {{ margin-top: 24px; }}
    .driver h2 {{ padding-bottom: 10px; border-bottom: 2px solid #17324d; }}
    .step {{ background: white; border: 1px solid #d9e0e7; border-radius: 8px; margin-top: 14px; padding: 16px; }}
    .route-overview {{ background: white; border: 1px solid #c7d2dd; border-radius: 8px; margin-top: 14px; padding: 16px; }}
    .step-head {{ display: flex; justify-content: space-between; gap: 12px; flex-wrap: wrap; }}
    .step-head span {{ color: #65717d; }}
    .times {{ display: flex; gap: 8px; flex-wrap: wrap; margin: 12px 0; }}
    .times span {{ background: #edf2f6; padding: 6px 9px; border-radius: 999px; font-size: 13px; }}
    .route {{ display: flex; flex-wrap: wrap; gap: 8px; padding: 0; margin: 0; list-style: none; }}
    .route li {{ border: 1px solid #cfd8e3; padding: 7px 9px; border-radius: 6px; background: #fbfcfd; max-width: 360px; }}
    .route li strong {{ display: block; }}
    .full-route {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 8px; padding: 0; margin: 12px 0 0; list-style: none; }}
    .full-route li {{ display: grid; grid-template-columns: 32px 1fr; align-items: start; gap: 8px; border: 1px solid #cfd8e3; border-radius: 6px; background: #fbfcfd; padding: 8px; }}
    .full-route span {{ display: inline-flex; align-items: center; justify-content: center; width: 24px; height: 24px; border-radius: 999px; background: #17324d; color: white; font-size: 12px; }}
    .full-route strong {{ overflow-wrap: anywhere; }}
    .order-details {{ margin-top: 8px; display: grid; gap: 6px; }}
    .order-detail {{ border-top: 1px solid #e2e8f0; padding-top: 6px; font-size: 12px; color: #52606d; }}
    .order-detail span {{ display: block; font-weight: 700; }}
    .order-detail p {{ margin: 3px 0 0; color: #1f2933; }}
    .route li:not(:last-child)::after {{ content: "→"; margin-left: 8px; color: #8a96a3; }}
  </style>
</head>
<body>
  <header>
    <h1>Routecraft Delivery Plan</h1>
    <div class="subtitle">Generated {html.escape(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))}</div>
  </header>
  <main>
    <div class="metrics">{''.join(summary_cards)}</div>
    {summary_table}
    {''.join(driver_sections)}
  </main>
</body>
</html>
""",
        encoding="utf-8",
    )
    return output_path


def split_route_path(route_path: Any) -> list[str]:
    return [
        str(part).strip()
        for part in str(route_path).split("→")
        if str(part).strip()
    ]


def normalize_detail_key(city: str, address: str) -> tuple[str, str, str]:
    address = clean_cell(address).replace(",", " ")
    city = clean_cell(city)
    if city:
        address = re.sub(re.escape(city), " ", address, flags=re.IGNORECASE)
    address = re.sub(r"\s+", " ", address).strip()
    match = re.search(r"\d+[\wא-ת/-]*", address)
    if not match:
        return city, address, ""
    number = match.group(0).replace(".0", "").strip()
    street = address[: match.start()].strip(" ,-/")
    if not street:
        street = address[match.end() :].strip(" ,-/")
    return city, re.sub(r"\s+", " ", street), number


def load_order_details(path: Path) -> dict[tuple[str, str, str], list[dict[str, str]]]:
    if not path.exists():
        return {}
    df = pd.read_excel(path).fillna("")
    normalized_cols = {str(col).strip().lower().replace("_", " "): col for col in df.columns}

    def col(*names: str) -> str | None:
        for name in names:
            found = normalized_cols.get(name)
            if found is not None:
                return found
        return None

    order_col = col("order id", "orderid")
    client_col = col("client id", "clientid")
    site_col = col("site id", "siteid")
    comments_col = col("comments", "comment", "notes")

    details: dict[tuple[str, str, str], list[dict[str, str]]] = {}
    for _, row in df.iterrows():
        key = normalize_detail_key(row.get("City", ""), f"{row.get('Street_Name', '')} {row.get('House_Number', '')}")
        item = {
            "order": clean_cell(row.get(order_col, "")) if order_col else "",
            "client": clean_cell(row.get(client_col, "")) if client_col else "",
            "site": clean_cell(row.get(site_col, "")) if site_col else "",
            "comments": clean_cell(row.get(comments_col, "")) if comments_col else "",
        }
        details.setdefault(key, []).append(item)
    return details


def render_route_node(node: str, city: str, order_details: dict[tuple[str, str, str], list[dict[str, str]]]) -> str:
    key = normalize_detail_key(city, node)
    details = order_details.get(key, [])
    detail_html = ""
    if details:
        rows = []
        for item in details:
            meta = " | ".join(
                part
                for part in [
                    f"Order {item['order']}" if item["order"] else "",
                    f"Client {item['client']}" if item["client"] else "",
                    f"Site {item['site']}" if item["site"] else "",
                ]
                if part
            )
            rows.append(
                "<div class=\"order-detail\">"
                f"{'<span>' + html.escape(meta) + '</span>' if meta else ''}"
                f"{'<p>' + html.escape(item['comments']) + '</p>' if item['comments'] else ''}"
                "</div>"
            )
        detail_html = f"<div class=\"order-details\">{''.join(rows)}</div>"
    return f"<li><strong>{html.escape(node)}</strong>{detail_html}</li>"


def run_tsp_script(matrix_path: Path) -> dict[str, Any]:
    stdout = io.StringIO()
    stderr = io.StringIO()
    install_bidi_fallback()
    module_suffix = re.sub(r"\W+", "_", matrix_path.stem)
    with patch("tkinter.Tk", return_value=HeadlessTk()):
        with patch("tkinter.filedialog.askopenfilename", return_value=str(matrix_path)):
            with contextlib.redirect_stdout(stdout), contextlib.redirect_stderr(stderr):
                import_script(ROOT / "TSP.py", f"routecraft_tsp_{module_suffix}")

    text = stdout.getvalue()
    err = stderr.getvalue()
    fallback = fallback_tsp_summary(matrix_path, text)
    return {
        "matrix": matrix_path,
        "stdout": text,
        "stderr": err,
        "best_ending": find_line_value(text, "Best ending:") or fallback.get("best_ending", ""),
        "best_cost": find_line_value(text, "Best cost:") or fallback.get("best_cost", ""),
        "best_route": find_line_value(text, "Best route:") or fallback.get("best_route", ""),
        "real_addresses": find_line_value(text, "Real addresses:") or fallback.get("real_addresses", ""),
    }


class HeadlessTk:
    def withdraw(self) -> None:
        return None

    def attributes(self, *_args: Any, **_kwargs: Any) -> None:
        return None

    def destroy(self) -> None:
        return None


def install_bidi_fallback() -> None:
    try:
        if importlib.util.find_spec("bidi") is not None:
            return
    except ValueError:
        return

    bidi_module = types.ModuleType("bidi")
    algorithm_module = types.ModuleType("bidi.algorithm")
    bidi_module.__spec__ = importlib.machinery.ModuleSpec("bidi", None)
    algorithm_module.__spec__ = importlib.machinery.ModuleSpec("bidi.algorithm", None)
    algorithm_module.get_display = lambda value: value
    bidi_module.algorithm = algorithm_module
    sys.modules.setdefault("bidi", bidi_module)
    sys.modules.setdefault("bidi.algorithm", algorithm_module)


def find_line_value(text: str, prefix: str) -> str:
    for line in text.splitlines():
        if line.strip().startswith(prefix):
            return line.split(prefix, 1)[1].strip()
    return ""


def fallback_tsp_summary(matrix_path: Path, text: str) -> dict[str, str]:
    matches = re.findall(
        r"OPTION:\s+END AT\s+(\d+).*?2-Opt Optimized:\s*Route:\s*([0-9\s\->]+)\s*Cost:\s*([0-9.]+)\s+minutes",
        text,
        flags=re.DOTALL,
    )
    if not matches:
        return {}

    best_end, best_route, best_cost = min(matches, key=lambda item: float(item[2]))
    labels = pd.read_excel(matrix_path, header=0, index_col=0).index.tolist()
    route_indices = [int(part.strip()) for part in best_route.split("->")]
    real_addresses = " -> ".join(str(labels[idx]) for idx in route_indices if idx < len(labels))
    return {
        "best_ending": best_end,
        "best_cost": f"{float(best_cost):.2f} minutes",
        "best_route": best_route.strip(),
        "real_addresses": real_addresses,
    }


def stage_6_tsp_html(matrix_paths: list[Path], run_dir: Path) -> tuple[Path, StageResult]:
    if not matrix_paths:
        raise RuntimeError("Stage 6 requires at least one matrix file.")

    tsp_results = [run_tsp_script(path) for path in matrix_paths]
    output_path = run_dir / "06_route_results.html"
    output_path.write_text(render_tsp_html(tsp_results), encoding="utf-8")
    return output_path, StageResult(6, "Solve routes and render HTML", "completed", [output_path])


def render_tsp_html(results: list[dict[str, Any]]) -> str:
    rows = []
    detail_sections = []
    for result in results:
        matrix_name = html.escape(result["matrix"].name)
        rows.append(
            "<tr>"
            f"<td>{matrix_name}</td>"
            f"<td>{html.escape(result['best_cost'] or 'n/a')}</td>"
            f"<td>{html.escape(result['best_route'] or 'n/a')}</td>"
            f"<td>{html.escape(result['real_addresses'] or 'n/a')}</td>"
            "</tr>"
        )
        detail_sections.append(
            f"<section><h2>{matrix_name}</h2>"
            f"<pre>{html.escape(result['stdout'])}</pre>"
            f"{'<h3>Diagnostics</h3><pre>' + html.escape(result['stderr']) + '</pre>' if result['stderr'] else ''}"
            "</section>"
        )

    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Routecraft Route Results</title>
  <style>
    body {{ margin: 0; font-family: Arial, sans-serif; background: #f6f7f9; color: #1d232a; }}
    header {{ background: #153243; color: white; padding: 28px 36px; }}
    main {{ padding: 28px 36px; }}
    table {{ width: 100%; border-collapse: collapse; background: white; border: 1px solid #d9dee5; }}
    th, td {{ padding: 12px 14px; border-bottom: 1px solid #e5e9ef; text-align: left; vertical-align: top; }}
    th {{ background: #eef2f6; font-size: 13px; text-transform: uppercase; letter-spacing: 0; }}
    section {{ margin-top: 24px; background: white; border: 1px solid #d9dee5; padding: 18px; }}
    h1, h2 {{ margin: 0 0 10px; }}
    pre {{ white-space: pre-wrap; overflow-wrap: anywhere; background: #111820; color: #edf3f7; padding: 14px; }}
  </style>
</head>
<body>
  <header>
    <h1>Routecraft Route Results</h1>
    <div>Generated {html.escape(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))}</div>
  </header>
  <main>
    <table>
      <thead><tr><th>Matrix</th><th>Best Cost</th><th>Best Route</th><th>Addresses</th></tr></thead>
      <tbody>{''.join(rows)}</tbody>
    </table>
    {''.join(detail_sections)}
  </main>
</body>
</html>
"""


def render_failure_html(run_dir: Path, stage_results: list[StageResult], error: BaseException) -> Path:
    output_path = run_dir / "failure_report.html"
    completed = "".join(
        f"<li>Stage {r.number}: {html.escape(r.name)} - {html.escape(r.status)}</li>"
        for r in stage_results
    )
    output_path.write_text(
        f"""<!doctype html>
<html lang="en">
<head><meta charset="utf-8"><title>Routecraft Flow Failure</title>
<style>body{{font-family:Arial,sans-serif;padding:28px;background:#f7f7f7;color:#222}}pre{{background:#191d24;color:#f5f5f5;padding:16px;white-space:pre-wrap}}</style></head>
<body>
<h1>Routecraft Flow Failed</h1>
<p>The pipeline stopped and wrote this report as requested.</p>
<h2>Completed Stages</h2>
<ul>{completed or '<li>None</li>'}</ul>
<h2>Error</h2>
<pre>{html.escape(str(error))}</pre>
<h2>Traceback</h2>
<pre>{html.escape(traceback.format_exc())}</pre>
</body></html>
""",
        encoding="utf-8",
    )
    return output_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run the Routecraft constraints pipeline without file pickers.")
    parser.add_argument("--input", required=True, help="Input Excel file for the selected start stage.")
    parser.add_argument("--start-stage", type=int, default=1, choices=range(1, 7))
    parser.add_argument("--runs-dir", default=str(RUNS_DIR))
    parser.add_argument("--google-api-key")
    parser.add_argument("--distance-api-key")
    parser.add_argument(
        "--mock-google",
        action="store_true",
        help="Use deterministic local mocks for Google geocoding, distance matrices, and delivery origin travel times.",
    )
    parser.add_argument("--max-workers", type=int, default=10)
    parser.add_argument("--cluster-threshold-meters", type=float, default=150)
    parser.add_argument("--constraints-file", default=str(CONSTRAINTS_PARAMETERS_PATH))
    args = parser.parse_args()
    args.constraints = load_constraints_parameters(Path(args.constraints_file))
    return args


def main() -> int:
    args = parse_args()
    input_path = require_file(Path(args.input), "Input file")
    load_env_file(ROOT / ".env")
    load_env_file(ROOT / "addresses" / ".env")
    load_env_file(input_path.parent / ".env")
    run_dir = make_run_dir(Path(args.runs_dir).resolve())
    stage_results: list[StageResult] = []

    try:
        current: Path | list[Path] = input_path
        if args.start_stage <= 1:
            current, result = stage_1_cleanup(Path(current), run_dir, args)
            stage_results.append(result)
        if args.start_stage <= 2:
            current, result = stage_2_geocode(Path(current), run_dir, args)
            stage_results.append(result)
        if args.start_stage <= 3:
            current, result = stage_3_cluster(Path(current), run_dir, args)
            stage_results.append(result)
        if args.start_stage <= 4:
            current, result = stage_4_group(Path(current), run_dir, args)
            stage_results.append(result)
        delivery_input_path: Path | None = None
        matrix_folder = run_dir
        if args.start_stage <= 5:
            delivery_input_path = Path(current)
            current, result = stage_5_distance_matrices(Path(current), run_dir, args)
            stage_results.append(result)
        else:
            matrix_folder = input_path.parent
        if args.start_stage <= 6:
            if delivery_input_path is None:
                delivery_input_path = Path(current)
            current, result = stage_6_delivery_plan(delivery_input_path, matrix_folder, run_dir, args)
            stage_results.append(result)

        if not args.mock_google:
            saved_mock_paths = save_real_mock_data(run_dir)
            if saved_mock_paths:
                stage_results.append(
                    StageResult(
                        7,
                        "Refresh mock Google data",
                        "completed",
                        saved_mock_paths,
                        ["Saved real Google results for future --mock-google test runs."],
                    )
                )

        print(f"Flow completed. Run folder: {run_dir}")
        print(f"Final report: {current}")
        return 0
    except BaseException as exc:
        report_path = render_failure_html(run_dir, stage_results, exc)
        print(f"Flow failed. Report: {report_path}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
