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
CONSTRAINTS_PARAMETERS_PATH = ROOT / "constraints_parameters_experimental.json"
WAREHOUSE_NAME = "WAREHOUSE ESHTAOL"
WAREHOUSE_LAT = 31.77927525
WAREHOUSE_LNG = 35.0105885
ROUTE_SEPARATOR = " -> "


@dataclass
class StageResult:
    number: int
    name: str
    status: str
    output_files: list[Path] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)


@dataclass
class SmartSelectionEstimate:
    selected_count: int
    city_count: int
    load_count: int
    estimated_driver_times: list[float]
    estimated_total_minutes: float
    fits: bool
    safety_limit: float


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
    sys.modules[module_name] = module
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


def address_cache_id(address: str) -> str:
    normalized = re.sub(r"\s+", " ", str(address or "").strip()).casefold()
    return hashlib.sha256(normalized.encode("utf-8")).hexdigest()[:16]


def distance_pair_cache_paths() -> list[Path]:
    return [
        CONST_MOCK_DATA_DIR / "distance_pair_cache.xlsx",
        MOCK_DATA_DIR / "distance_pair_cache.xlsx",
    ]


def load_distance_pair_cache() -> dict[tuple[str, str], dict[str, Any]]:
    cache: dict[tuple[str, str], dict[str, Any]] = {}
    for path in distance_pair_cache_paths():
        if not path.exists():
            continue
        df = pd.read_excel(path).fillna("")
        for _, row in df.iterrows():
            origin = str(row.get("origin_address", "")).strip()
            destination = str(row.get("destination_address", "")).strip()
            minutes = pd.to_numeric(row.get("duration_minutes"), errors="coerce")
            if origin and destination and pd.notna(minutes):
                cache[(origin, destination)] = row.to_dict()
    return cache


def pair_cache_row(
    origin: str,
    destination: str,
    duration_minutes: int,
    source: str,
    distance_meters: Any = "",
) -> dict[str, Any]:
    return {
        "origin_id": address_cache_id(origin),
        "destination_id": address_cache_id(destination),
        "origin_address": origin,
        "destination_address": destination,
        "duration_minutes": int(duration_minutes),
        "distance_meters": distance_meters,
        "source": source,
        "calculated_at": datetime.now().isoformat(timespec="seconds"),
    }


def merge_distance_pair_cache_rows(existing: pd.DataFrame, updates: pd.DataFrame) -> pd.DataFrame:
    if existing.empty:
        merged = updates.copy()
    elif updates.empty:
        merged = existing.copy()
    else:
        merged = pd.concat([existing, updates], ignore_index=True)
    if merged.empty:
        return merged
    return merged.drop_duplicates(
        subset=["origin_address", "destination_address"],
        keep="last",
    ).sort_values(["origin_address", "destination_address"], kind="stable")


def save_distance_pair_cache_updates(run_dir: Path, rows: list[dict[str, Any]]) -> Path | None:
    if not rows:
        return None
    path = run_dir / "05_distance_pair_cache_updates.xlsx"
    pd.DataFrame(rows).to_excel(path, index=False)
    return path


def seed_pair_cache_from_matrix(
    matrix_path: Path,
    full_addresses: list[str],
    pair_cache: dict[tuple[str, str], dict[str, Any]],
) -> list[dict[str, Any]]:
    if not matrix_path.exists():
        return []
    matrix_df = pd.read_excel(matrix_path, index_col=0)
    updates: list[dict[str, Any]] = []
    for origin in full_addresses:
        if origin not in matrix_df.index:
            continue
        for destination in full_addresses:
            if origin == destination or destination not in matrix_df.columns:
                continue
            if (origin, destination) in pair_cache:
                continue
            minutes = pd.to_numeric(matrix_df.loc[origin, destination], errors="coerce")
            if pd.isna(minutes):
                continue
            update = pair_cache_row(origin, destination, int(minutes), "group-matrix-cache")
            pair_cache[(origin, destination)] = update
            updates.append(update)
    return updates


def build_matrix_with_pair_cache(
    full_addresses: list[str],
    coords_dict: dict[str, Any],
    api_key: str,
    allow_api: bool,
    pair_cache: dict[tuple[str, str], dict[str, Any]],
) -> tuple[pd.DataFrame, list[dict[str, Any]]]:
    matrix_data = []
    updates: list[dict[str, Any]] = []
    destination_chunk_size = 25

    for i, origin in enumerate(full_addresses):
        row_minutes = [999] * len(full_addresses)
        row_minutes[i] = 0
        origin_coords = coords_dict.get(origin)

        for start in range(0, len(full_addresses), destination_chunk_size):
            dest_addresses = full_addresses[start : start + destination_chunk_size]
            api_destinations = []
            api_positions = []
            api_estimates: dict[int, int] = {}

            for offset, destination in enumerate(dest_addresses):
                j = start + offset
                if i == j:
                    continue
                cached = pair_cache.get((origin, destination))
                if cached:
                    row_minutes[j] = int(pd.to_numeric(cached.get("duration_minutes"), errors="coerce"))
                    continue

                dest_coords = coords_dict.get(destination)
                estimated = 999
                if origin_coords and dest_coords:
                    estimated = mock_distance_matrix([origin, destination], coords_dict).iloc[0, 1]
                    estimated = int(estimated)
                if not allow_api:
                    row_minutes[j] = estimated
                    update = pair_cache_row(origin, destination, estimated, "mock-estimated")
                    pair_cache[(origin, destination)] = update
                    updates.append(update)
                    continue

                if not origin_coords or not dest_coords:
                    row_minutes[j] = estimated
                    update = pair_cache_row(origin, destination, estimated, "missing-coordinates")
                    pair_cache[(origin, destination)] = update
                    updates.append(update)
                    continue

                row_minutes[j] = estimated
                api_estimates[j] = estimated
                api_destinations.append(f"{dest_coords['lat']},{dest_coords['lng']}")
                api_positions.append(j)

            if not allow_api or not api_destinations or not origin_coords:
                continue

            dm_params = {
                "origins": f"{origin_coords['lat']},{origin_coords['lng']}",
                "destinations": "|".join(api_destinations),
                "mode": "driving",
                "departure_time": "now",
                "key": api_key,
            }
            try:
                dm_res = requests.get(
                    "https://maps.googleapis.com/maps/api/distancematrix/json",
                    params=dm_params,
                    timeout=30,
                ).json()
                if dm_res.get("status") == "OK":
                    elements = dm_res["rows"][0]["elements"]
                    for dest_index, element in zip(api_positions, elements):
                        destination = full_addresses[dest_index]
                        minutes = api_estimates.get(dest_index, row_minutes[dest_index])
                        meters: Any = ""
                        source = "google"
                        if element.get("status") == "OK":
                            seconds = element.get("duration_in_traffic", element["duration"])["value"]
                            minutes = round(seconds / 60)
                            meters = element.get("distance", {}).get("value", "")
                        else:
                            source = "google-fallback-estimated"
                        row_minutes[dest_index] = minutes
                        update = pair_cache_row(origin, destination, minutes, source, meters)
                        pair_cache[(origin, destination)] = update
                        updates.append(update)
            except Exception:
                for dest_index in api_positions:
                    destination = full_addresses[dest_index]
                    minutes = api_estimates.get(dest_index, row_minutes[dest_index])
                    update = pair_cache_row(origin, destination, minutes, "api-error-estimated")
                    pair_cache[(origin, destination)] = update
                    updates.append(update)

        matrix_data.append(row_minutes)

    return pd.DataFrame(matrix_data, index=full_addresses, columns=full_addresses), updates


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

    pair_updates_path = run_dir / "05_distance_pair_cache_updates.xlsx"
    if pair_updates_path.exists():
        pair_target = target_dir / "distance_pair_cache.xlsx"
        existing = pd.read_excel(pair_target).fillna("") if pair_target.exists() else pd.DataFrame()
        updates = pd.read_excel(pair_updates_path).fillna("")
        merge_distance_pair_cache_rows(existing, updates).to_excel(pair_target, index=False)
        saved_paths.append(pair_target)

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


def geocode_warehouse_address(address: str, key: str) -> tuple[float, float, str]:
    response = requests.get(
        "https://maps.googleapis.com/maps/api/geocode/json",
        params={
            "address": address,
            "components": "country:IL",
            "language": "he",
            "key": key,
        },
        timeout=30,
    )
    response.raise_for_status()
    data = response.json()
    if data.get("status") != "OK" or not data.get("results"):
        raise RuntimeError(f"Could not geocode warehouse address '{address}': {data.get('status', 'no result')}")
    location = data["results"][0].get("geometry", {}).get("location", {})
    lat = location.get("lat")
    lng = location.get("lng")
    if lat is None or lng is None:
        raise RuntimeError(f"Could not geocode warehouse address '{address}': missing coordinates")
    return float(lat), float(lng), str(data["results"][0].get("formatted_address", address))


def resolve_warehouse(args: argparse.Namespace) -> dict[str, Any]:
    name = clean_cell(constraint_value(args, "warehouse_name", WAREHOUSE_NAME)) or WAREHOUSE_NAME
    address = clean_cell(constraint_value(args, "warehouse_address", "")) or name
    default_addresses = {WAREHOUSE_NAME.lower(), "eshtaol", "eshtaol, israel", "אשתאול"}
    if address.lower() in default_addresses:
        return {
            "name": name,
            "address": address,
            "lat": WAREHOUSE_LAT,
            "lng": WAREHOUSE_LNG,
            "source": "default_coordinates",
            "formatted_address": address,
        }
    if args.mock_google:
        lat, lng = mock_lat_lng(address)
        return {
            "name": name,
            "address": address,
            "lat": lat,
            "lng": lng,
            "source": "mock_coordinates",
            "formatted_address": address,
        }
    lat, lng, formatted = geocode_warehouse_address(address, get_google_key(args))
    return {
        "name": name,
        "address": address,
        "lat": lat,
        "lng": lng,
        "source": "google",
        "formatted_address": formatted,
    }


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
    capacity = initial_selection_capacity(args)
    selected, deferred = split_orders_by_priority_and_closeness(df, capacity, args)
    selected_path = run_dir / "02b_selected_orders_for_run.xlsx"
    deferred_path = run_dir / "02c_deferred_orders_next_run.xlsx"
    selected.to_excel(selected_path, index=False)
    deferred.to_excel(deferred_path, index=False)
    return selected, deferred, selected_path, deferred_path


def sort_orders_for_selection(df: pd.DataFrame, args: argparse.Namespace) -> pd.DataFrame:
    df = df.copy()
    due_col = str(constraint_value(args, "due_date_column", "required_delivery_date"))
    if due_col not in df.columns:
        df[due_col] = ""
    parsed_due_dates = df[due_col].apply(parse_required_delivery_date)
    df["required_delivery_date"] = df[due_col].apply(format_required_delivery_date)
    df["delivery_priority"] = parsed_due_dates.apply(lambda value: delivery_priority_for_date(value, args))
    df["package_count"] = 1
    df["_selection_order"] = range(len(df))
    df["_parsed_required_delivery_date"] = parsed_due_dates
    sorted_df = df.sort_values(
        by=["delivery_priority", "_parsed_required_delivery_date", "_selection_order"],
        kind="stable",
    ).reset_index(drop=True)
    return sorted_df.drop(columns=["_selection_order", "_parsed_required_delivery_date"])


def normalize_selection_text(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip()).casefold()


def location_key(row: pd.Series) -> tuple[str, str]:
    return (
        normalize_selection_text(row.get("City", "")),
        normalize_selection_text(row.get("Street_Name", "")),
    )


def city_key(row: pd.Series) -> str:
    return normalize_selection_text(row.get("City", ""))


def append_candidate_indices(
    chosen: list[Any],
    chosen_set: set[Any],
    candidate_indices: list[Any],
    remaining: int,
) -> int:
    for index in candidate_indices:
        if remaining <= 0:
            break
        if index in chosen_set:
            continue
        chosen.append(index)
        chosen_set.add(index)
        remaining -= 1
    return remaining


def choose_priority_zero_locality(priority_df: pd.DataFrame, remaining: int) -> list[Any]:
    chosen: list[Any] = []
    chosen_set: set[Any] = set()
    city_order: list[str] = []
    for _, row in priority_df.iterrows():
        city = city_key(row)
        if city not in city_order:
            city_order.append(city)

    for city in city_order:
        city_df = priority_df[priority_df.apply(city_key, axis=1) == city]
        street_order: list[str] = []
        for _, row in city_df.iterrows():
            street = location_key(row)[1]
            if street not in street_order:
                street_order.append(street)
        for street in street_order:
            block_indices = [
                index
                for index, row in city_df.iterrows()
                if location_key(row)[1] == street
            ]
            remaining = append_candidate_indices(chosen, chosen_set, block_indices, remaining)
            if remaining <= 0:
                return chosen
    return chosen


def distance_to_selected(row: pd.Series, selected_df: pd.DataFrame) -> float:
    lat = pd.to_numeric(row.get("LAT"), errors="coerce")
    lng = pd.to_numeric(row.get("LNG"), errors="coerce")
    if pd.isna(lat) or pd.isna(lng):
        return float("inf")
    best = float("inf")
    for _, selected_row in selected_df.iterrows():
        selected_lat = pd.to_numeric(selected_row.get("LAT"), errors="coerce")
        selected_lng = pd.to_numeric(selected_row.get("LNG"), errors="coerce")
        if pd.isna(selected_lat) or pd.isna(selected_lng):
            continue
        best = min(best, calculate_haversine(lat, lng, selected_lat, selected_lng))
    return best


def choose_local_to_selected(
    priority_df: pd.DataFrame,
    selected_df: pd.DataFrame,
    remaining: int,
) -> list[Any]:
    if selected_df.empty:
        return choose_priority_zero_locality(priority_df, remaining)

    chosen: list[Any] = []
    chosen_set: set[Any] = set()
    selected_location_order: list[tuple[str, str]] = []
    selected_city_order: list[str] = []
    for _, row in selected_df.iterrows():
        loc = location_key(row)
        city = loc[0]
        if loc not in selected_location_order:
            selected_location_order.append(loc)
        if city not in selected_city_order:
            selected_city_order.append(city)

    for loc in selected_location_order:
        indices = [
            index
            for index, row in priority_df.iterrows()
            if location_key(row) == loc
        ]
        remaining = append_candidate_indices(chosen, chosen_set, indices, remaining)
        if remaining <= 0:
            return chosen

    for city in selected_city_order:
        city_df = priority_df[priority_df.apply(city_key, axis=1) == city]
        street_order: list[str] = []
        for _, row in city_df.iterrows():
            street = location_key(row)[1]
            if street not in street_order:
                street_order.append(street)
        for street in street_order:
            indices = [
                index
                for index, row in city_df.iterrows()
                if location_key(row)[1] == street
            ]
            remaining = append_candidate_indices(chosen, chosen_set, indices, remaining)
            if remaining <= 0:
                return chosen

    nearest_indices = sorted(
        [index for index in priority_df.index if index not in chosen_set],
        key=lambda index: distance_to_selected(priority_df.loc[index], selected_df),
    )
    remaining = append_candidate_indices(chosen, chosen_set, nearest_indices, remaining)
    if remaining <= 0:
        return chosen

    remaining_indices = [index for index in priority_df.index if index not in chosen_set]
    append_candidate_indices(chosen, chosen_set, remaining_indices, remaining)
    return chosen


def split_orders_by_priority_and_closeness(
    df: pd.DataFrame, selected_count: int, args: argparse.Namespace
) -> tuple[pd.DataFrame, pd.DataFrame]:
    ordered_df = sort_orders_for_selection(df, args)
    selected_count = max(1, min(int(selected_count), len(ordered_df))) if len(ordered_df) else 0
    selected_indices: list[Any] = []

    for priority in ordered_df["delivery_priority"].drop_duplicates().tolist():
        if len(selected_indices) >= selected_count:
            break
        priority_df = ordered_df[ordered_df["delivery_priority"] == priority]
        remaining = selected_count - len(selected_indices)
        if len(priority_df) <= remaining:
            selected_indices.extend(priority_df.index.tolist())
            continue
        selected_so_far = ordered_df.loc[selected_indices] if selected_indices else ordered_df.iloc[0:0]
        selected_indices.extend(choose_local_to_selected(priority_df, selected_so_far, remaining))
        break

    selected_set = set(selected_indices)
    deferred_indices = [index for index in ordered_df.index if index not in selected_set]
    selected = ordered_df.loc[selected_indices].reset_index(drop=True)
    deferred = ordered_df.loc[deferred_indices].reset_index(drop=True)
    return selected, deferred


def initial_selection_capacity(args: argparse.Namespace) -> int:
    return int(
        int(constraint_value(args, "drivers", 3))
        * int(constraint_value(args, "max_packages_per_driver", 100))
        * float(constraint_value(args, "selection_capacity_multiplier", 2))
    )


def constraint_bool(args: argparse.Namespace, name: str, default: bool) -> bool:
    value = constraint_value(args, name, default)
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.strip().lower() in {"1", "true", "yes", "y", "on"}
    return bool(value)


def air_drive_minutes_between(
    lat1: Any,
    lng1: Any,
    lat2: Any,
    lng2: Any,
    args: argparse.Namespace,
) -> float:
    meters = calculate_haversine(lat1, lng1, lat2, lng2)
    drive_multiplier = float(constraint_value(args, "air_distance_drive_multiplier", 1.35))
    return (meters * drive_multiplier) / (42 * 1000 / 60)


def estimate_city_centroids(selected: pd.DataFrame) -> dict[str, dict[str, Any]]:
    cities: dict[str, dict[str, Any]] = {}
    for _, row in selected.iterrows():
        city = clean_cell(row.get("City", ""))
        lat = pd.to_numeric(row.get("LAT"), errors="coerce")
        lng = pd.to_numeric(row.get("LNG"), errors="coerce")
        if not city or pd.isna(lat) or pd.isna(lng):
            continue
        entry = cities.setdefault(city, {"lat_values": [], "lng_values": [], "count": 0})
        entry["lat_values"].append(float(lat))
        entry["lng_values"].append(float(lng))
        entry["count"] += 1
    for entry in cities.values():
        entry["lat"] = sum(entry["lat_values"]) / len(entry["lat_values"])
        entry["lng"] = sum(entry["lng_values"]) / len(entry["lng_values"])
    return cities


def city_sequence_for_rows(rows: pd.DataFrame) -> list[str]:
    sequence: list[str] = []
    for _, row in rows.iterrows():
        city = clean_cell(row.get("City", ""))
        if city and city not in sequence:
            sequence.append(city)
    return sequence


def estimate_load_minutes(
    load_rows: pd.DataFrame,
    city_centroids: dict[str, dict[str, Any]],
    warehouse: dict[str, Any],
    args: argparse.Namespace,
) -> float:
    service_minutes = float(constraint_value(args, "service_minutes_per_package", 4))
    city_stop_minutes = float(constraint_value(args, "estimated_city_stop_minutes", 3))
    service_time = len(load_rows) * service_minutes
    city_sequence = city_sequence_for_rows(load_rows)
    if not city_sequence:
        return service_time

    travel_time = 0.0
    previous_lat = float(warehouse["lat"])
    previous_lng = float(warehouse["lng"])
    for city in city_sequence:
        centroid = city_centroids.get(city)
        if not centroid:
            continue
        travel_time += air_drive_minutes_between(previous_lat, previous_lng, centroid["lat"], centroid["lng"], args)
        previous_lat = float(centroid["lat"])
        previous_lng = float(centroid["lng"])
        city_rows = load_rows[load_rows["City"].astype(str).eq(str(city))]
        travel_time += max(0, len(city_rows) - 1) * city_stop_minutes
    travel_time += air_drive_minutes_between(previous_lat, previous_lng, warehouse["lat"], warehouse["lng"], args)
    return service_time + travel_time


def estimate_selected_rows(
    ordered_orders: pd.DataFrame,
    selected_count: int,
    args: argparse.Namespace,
    warehouse: dict[str, Any],
) -> SmartSelectionEstimate:
    selected, _deferred = split_orders_by_priority_and_closeness(ordered_orders, selected_count, args)
    drivers = max(1, int(constraint_value(args, "drivers", 3)))
    max_packages = max(1, int(constraint_value(args, "max_packages_per_driver", 100)))
    max_shift = float(constraint_value(args, "max_shift_minutes", 720))
    safety = float(constraint_value(args, "smart_estimate_safety_buffer", 0.9))
    safety_limit = max_shift * safety
    city_centroids = estimate_city_centroids(selected)
    driver_times = [0.0 for _ in range(drivers)]
    load_count = 0

    for start in range(0, len(selected), max_packages):
        load_rows = selected.iloc[start : start + max_packages]
        if load_rows.empty:
            continue
        load_count += 1
        load_minutes = estimate_load_minutes(load_rows, city_centroids, warehouse, args)
        driver_index = min(range(drivers), key=lambda index: driver_times[index])
        driver_times[driver_index] += load_minutes

    fits = bool(driver_times) and max(driver_times) <= safety_limit
    return SmartSelectionEstimate(
        selected_count=len(selected),
        city_count=len(city_centroids),
        load_count=load_count,
        estimated_driver_times=driver_times,
        estimated_total_minutes=sum(driver_times),
        fits=fits,
        safety_limit=safety_limit,
    )


def find_smart_initial_selection_count(
    ordered_orders: pd.DataFrame,
    args: argparse.Namespace,
    warehouse: dict[str, Any],
) -> tuple[int, SmartSelectionEstimate]:
    if ordered_orders.empty:
        raise RuntimeError("Smart selection has no orders to estimate.")
    low = 1
    high = len(ordered_orders)
    best_count = 1
    best_estimate = estimate_selected_rows(ordered_orders, 1, args, warehouse)
    last_estimate = best_estimate

    while low <= high:
        mid = (low + high) // 2
        estimate = estimate_selected_rows(ordered_orders, mid, args, warehouse)
        last_estimate = estimate
        if estimate.fits:
            best_count = mid
            best_estimate = estimate
            low = mid + 1
        else:
            high = mid - 1

    if not best_estimate.fits:
        return min(initial_selection_capacity(args), len(ordered_orders)), last_estimate
    return best_count, best_estimate


def smart_estimate_notes(estimate: SmartSelectionEstimate, label: str = "Smart initial estimate") -> list[str]:
    times = ", ".join(f"{minutes:.0f}" for minutes in estimate.estimated_driver_times)
    return [
        f"{label}: selected {estimate.selected_count} rows across {estimate.city_count} cities.",
        f"Estimated loads: {estimate.load_count}; estimated driver minutes: [{times}]; safety limit per driver: {estimate.safety_limit:.0f}.",
    ]


def stage_1_cleanup(input_path: Path, run_dir: Path, args: argparse.Namespace) -> tuple[Path, StageResult]:
    cleanup = import_script(ROOT / "address_parser_experimental.py", "routecraft_address_parser_experimental")
    df = pd.read_excel(input_path, dtype=str).fillna("")
    columns = cleanup.resolve_columns(df)
    direct_city_col = find_raw_col(df, "City", "city", "site name", "site_name")
    direct_street_col = find_raw_col(df, "Street_Name", "street name")
    direct_number_col = find_raw_col(df, "House_Number", "house number", "building number")
    if "ship_to_city" not in columns and direct_city_col:
        columns["ship_to_city"] = direct_city_col
    if "ship_to_street1" not in columns and direct_street_col:
        columns["ship_to_street1"] = direct_street_col
    if "ship_to_street2" not in columns and direct_number_col:
        columns["ship_to_street2"] = direct_number_col
    missing = [name for name in ["ship_to_city", "ship_to_street1"] if name not in columns]
    if missing:
        raise RuntimeError(f"Stage 1 input is missing address columns: {', '.join(missing)}")

    good_rows: list[dict[str, Any]] = []
    original_rows: list[dict[str, Any]] = []
    failed_rows: list[dict[str, Any]] = []
    use_skill = not constraint_bool(args, "disable_address_skill", False)
    for idx, row in df.iterrows():
        source_row = idx + 2
        city = row.get(columns["ship_to_city"], "")
        street1 = row.get(columns["ship_to_street1"], "")
        street2 = row.get(columns.get("ship_to_street2", ""), "")
        if direct_city_col and direct_street_col and direct_number_col:
            city = row.get(direct_city_col, city)
            street1 = " ".join(
                part
                for part in [
                    clean_cell(row.get(direct_street_col, "")),
                    clean_cell(row.get(direct_number_col, "")),
                ]
                if part
            )
            street2 = ""
        cleaned = cleanup.clean_raw_address(city, street1, street2, use_skill=use_skill)
        required_delivery_date = row.get(columns.get("required_delivery_date", ""), "")
        comments = row.get(columns.get("comments", ""), "")
        base_result = row.to_dict()
        base_result.update(
            {
                "order_id": row.get(columns.get("order_id", ""), ""),
                "client_id": row.get(columns.get("client_id", ""), ""),
                "site_id": row.get(columns.get("site_id", ""), ""),
                "source_row": source_row,
                "City": cleaned.city,
                "Street_Name": cleaned.street,
                "House_Number": cleaned.house_number,
                "Apartment": cleaned.apartment,
                "Floor": cleaned.floor,
                "Entrance": cleaned.entrance,
                "Comments": comments,
                "Secondary_Notes": cleaned.secondary_notes,
                "Secondary_Classification": cleaned.secondary_classification,
                "Needs_Review": "yes" if cleaned.needs_review else "no",
                "original_city": clean_cell(city),
                "original_street1": clean_cell(street1),
                "original_street2": clean_cell(street2),
                "required_delivery_date": required_delivery_date,
                "merged_address": cleaned.merged_address,
                "cleaned_address": cleaned.cleaned_address,
                "cleanup_status": cleaned.status,
                "Confidence": cleaned.confidence,
                "ParserOutput": cleaned.parser_output,
            }
        )
        if cleaned.is_valid:
            good_rows.append(
                {
                    "City": cleaned.city,
                    "Street_Name": cleaned.street,
                    "House_Number": cleaned.house_number,
                    "source_row": source_row,
                    "order_id": base_result["order_id"],
                    "client_id": base_result["client_id"],
                    "site_id": base_result["site_id"],
                    "Apartment": cleaned.apartment,
                    "Floor": cleaned.floor,
                    "Entrance": cleaned.entrance,
                    "Comments": comments,
                    "original_city": clean_cell(city),
                    "original_street1": clean_cell(street1),
                    "original_street2": clean_cell(street2),
                    "required_delivery_date": required_delivery_date,
                    "merged_address": cleaned.merged_address,
                    "cleaned_address": cleaned.cleaned_address,
                    "cleanup_status": cleaned.status,
                    "Confidence": cleaned.confidence,
                    "Needs_Review": "yes" if cleaned.needs_review else "no",
                }
            )
            original_rows.append(base_result)
        else:
            failed_rows.append(base_result)

    failed_path = run_dir / "01a_failed_addresses.xlsx"
    shaped_path = run_dir / "01b_addresses_for_geocoding.xlsx"
    original_path = run_dir / "01c_good_orders_original_format.xlsx"
    stage_1_columns = [
        "City",
        "Street_Name",
        "House_Number",
        "source_row",
        "required_delivery_date",
    ]
    pd.DataFrame(failed_rows).to_excel(failed_path, index=False)
    pd.DataFrame(good_rows).reindex(columns=stage_1_columns).to_excel(shaped_path, index=False)
    pd.DataFrame(original_rows).to_excel(original_path, index=False)

    return shaped_path, StageResult(
        1,
        "Smart clean raw order addresses",
        "completed",
        [failed_path, shaped_path, original_path],
        [
            f"Smart Israeli parser cleaned {len(good_rows)} rows; {len(failed_rows)} rows require review."
        ],
    )


def stage_2_geocode(input_path: Path, run_dir: Path, args: argparse.Namespace) -> tuple[Path, StageResult]:
    smart_address = import_script(ROOT / "address_parser_experimental.py", "routecraft_address_parser_experimental_stage2")
    google_key = "" if args.mock_google else get_google_key(args)
    df = pd.read_excel(input_path)
    if df.shape[1] < 3:
        raise RuntimeError("Stage 2 input must have at least 3 columns: city, street, house number.")

    rows = []
    failed_rows = []
    geocode_cache_rows = []
    for _, row in df.iterrows():
        city = clean_cell(row.get("City", ""))
        street = clean_cell(row.get("Street_Name", ""))
        house_number = clean_cell(row.get("House_Number", ""))
        cleaned_address = clean_cell(row.get("cleaned_address", "")) or f"{street} {house_number}".strip()
        query = f"{cleaned_address}, {city}".strip(" ,")
        if args.mock_google:
            lat, lng = mock_lat_lng(query)
            geocode_result = {
                "LAT": lat,
                "LNG": lng,
                "Coordinates": f"{lat},{lng}",
                "Geocode_Query": query,
                "Geocode_Status": "mock precise match",
                "Geocode_Precise": "yes",
                "Geocode_Usable": "yes",
                "Geocode_Formatted": query,
                "Geocode_Attempt_Count": 0,
                "Geocode_Query_Used": query,
                "Geocode_Result_Types": "mock",
                "Geocode_Location_Type": "MOCK",
                "Geocode_Diagnostic_Coordinates": "",
                "Geocode_Source": "mock",
                "Geocode_Estimated": "no",
                "Geocode_Failure_Reason": "",
                "Google_Street": "",
                "Google_House_Number": "",
                "Google_City": "",
            }
        else:
            geocode_result = smart_address.geocode_address(
                cleaned_address,
                city,
                google_key,
                street=street,
                house_number=house_number,
            )

        geocode_cache_rows.append(
            {
                "query": query,
                "city": city,
                "street": street,
                "house_number": house_number,
                "lat": geocode_result.get("LAT", ""),
                "lng": geocode_result.get("LNG", ""),
                "is_precise": clean_cell(geocode_result.get("Geocode_Usable", "")).lower() in {"yes", "review"},
                "status": geocode_result.get("Geocode_Status", ""),
                **geocode_result,
            }
        )

        rows.append(
            {
                "City": city,
                "Street_Name": street,
                "House_Number": house_number,
                "source_row": row.get("source_row", ""),
                "required_delivery_date": row.get("required_delivery_date", ""),
                **geocode_result,
            }
        )

    if not args.mock_google:
        smart_address.estimate_failed_geocodes_from_neighbors(rows)

    route_ready_rows = []
    for row_data in rows:
        if not smart_address.is_route_ready(row_data):
            failed_rows.append(
                {
                    "City": row_data.get("City", ""),
                    "Street_Name": row_data.get("Street_Name", ""),
                    "House_Number": row_data.get("House_Number", ""),
                    "merged_address": row_data.get("Geocode_Query", ""),
                    "cleanup_status": row_data.get("Geocode_Failure_Reason", "") or row_data.get("Geocode_Status", ""),
                    **row_data,
                }
            )
            continue
        google_street = clean_cell(row_data.get("Google_Street", ""))
        google_house_number = clean_cell(row_data.get("Google_House_Number", ""))
        if clean_cell(row_data.get("Geocode_Usable", "")).lower() == "yes":
            if google_street:
                row_data["Street_Name"] = google_street
            if google_house_number:
                row_data["House_Number"] = google_house_number
        route_ready_rows.append(
            {
                "City": row_data.get("City", ""),
                "Street_Name": row_data.get("Street_Name", ""),
                "House_Number": row_data.get("House_Number", ""),
                "LAT": row_data.get("LAT", ""),
                "LNG": row_data.get("LNG", ""),
                "geocode_status": row_data.get("Geocode_Status", ""),
                "source_row": row_data.get("source_row", ""),
                "required_delivery_date": row_data.get("required_delivery_date", ""),
            }
        )

    output_path = run_dir / "02_geocoded_addresses.xlsx"
    if not route_ready_rows:
        record_geocode_failures(run_dir, failed_rows)
        raise RuntimeError("Stage 2 did not find any route-ready geocoded addresses.")

    geocoded_df = pd.DataFrame(route_ready_rows)
    geocoded_df.to_excel(output_path, index=False)
    if geocode_cache_rows:
        pd.DataFrame(geocode_cache_rows).to_excel(run_dir / "02_geocoding_cache.xlsx", index=False)
    failure_path = record_geocode_failures(run_dir, failed_rows)
    output_files = [output_path, run_dir / "02_geocoding_cache.xlsx"] + ([failure_path] if failure_path else [])
    notes = []
    if args.disable_iterative_delivery:
        selected_df, deferred_df, selected_path, deferred_path = select_orders_for_current_run(geocoded_df, run_dir, args)
        output_files.extend([selected_path, deferred_path])
        notes.append(f"Selected {len(selected_df)} orders for this run; deferred {len(deferred_df)} orders for next run.")
        current = selected_path
    else:
        notes.append(
            "Wrote geocoded addresses; iterative search will create selected/deferred files inside each attempt folder."
        )
        current = output_path
    if failed_rows:
        notes.append(
            f"{len(failed_rows)} addresses were moved to the error file because they were not route-ready."
        )
    return current, StageResult(2, "Smart geocode addresses", "completed", output_files, notes)


def stage_3_cluster(input_path: Path, run_dir: Path, args: argparse.Namespace) -> tuple[Path, StageResult]:
    df = pd.read_excel(input_path)
    required = {"City", "Street_Name", "House_Number", "LAT", "LNG"}
    missing = required.difference(df.columns)
    if missing:
        raise RuntimeError(f"Stage 3 input is missing columns: {', '.join(sorted(missing))}")

    df["House_Number"] = df["House_Number"].apply(clean_cell)
    df["_house_number_sort"] = pd.to_numeric(
        df["House_Number"].astype(str).str.extract(r"(\d+)", expand=False),
        errors="coerce",
    )
    if df.empty:
        raise RuntimeError("Stage 3 has no rows to cluster.")

    df = df.sort_values(by=["City", "Street_Name", "_house_number_sort", "House_Number"], na_position="last").reset_index(drop=True)

    clusters: list[list[dict[str, Any]]] = []
    for _city, city_df in df.groupby("City", sort=False):
        city_df = city_df.sort_values(by=["Street_Name", "_house_number_sort", "House_Number"], na_position="last").reset_index(drop=True)
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
            " ".join(part for part in [clean_cell(item.get("Street_Name", "")), clean_cell(item.get("House_Number", ""))] if part)
            for item in cluster_rows
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
        rep.pop("_house_number_sort", None)
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
    distance = import_script(ROOT / "distance.matrix.experimental.py", "routecraft_distance_matrix_experimental")
    if args.mock_google:
        distance.API_KEY = "mock-google"
        distance.get_coords = mock_coords_from_cache
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

    pair_cache = load_distance_pair_cache()
    pair_cache_updates: list[dict[str, Any]] = []
    matrix_paths = []
    for group_value, group_df in grouped_frames.items():
        full_addresses = distance.build_full_addresses(group_df)
        safe_group = distance.sanitize_group_value(group_value)
        matrix_path = run_dir / f"05_distance_matrix_group-{safe_group}.xlsx"
        cached_matrix_path = mock_matrix_cache_path(group_value)
        pair_cache_updates.extend(seed_pair_cache_from_matrix(cached_matrix_path, full_addresses, pair_cache))
        matrix_df, updates = build_matrix_with_pair_cache(
            full_addresses,
            coords_dict,
            distance.API_KEY,
            not args.mock_google,
            pair_cache,
        )
        pair_cache_updates.extend(updates)
        matrix_df.to_excel(matrix_path)
        matrix_paths.append(matrix_path)

    pair_cache_path = save_distance_pair_cache_updates(run_dir, pair_cache_updates)
    output_files = [coords_cache_path, *matrix_paths]
    if pair_cache_path:
        output_files.append(pair_cache_path)
    return matrix_paths, StageResult(5, "Build distance matrices", "completed", output_files)


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


def join_route(parts: list[Any]) -> str:
    return ROUTE_SEPARATOR.join(clean_cell(part) for part in parts if clean_cell(part))


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
    minutes, _source = drive_duration_minutes_with_source(
        origin_lat,
        origin_lng,
        dest_lat,
        dest_lng,
        args,
        api_key,
    )
    return minutes


def drive_duration_minutes_with_source(
    origin_lat: Any,
    origin_lng: Any,
    dest_lat: Any,
    dest_lng: Any,
    args: argparse.Namespace,
    api_key: str | None,
) -> tuple[int, str]:
    try:
        cache_key = (round(float(origin_lat), 6), round(float(origin_lng), 6), round(float(dest_lat), 6), round(float(dest_lng), 6))
        cached_minutes = load_origin_duration_cache().get(cache_key)
        if cached_minutes is not None:
            return int(round(cached_minutes)), "origin_cache"
    except Exception:
        pass
    if args.mock_google or not api_key:
        return estimate_drive_minutes_between(origin_lat, origin_lng, dest_lat, dest_lng), "estimate"
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
        return estimate_drive_minutes_between(origin_lat, origin_lng, dest_lat, dest_lng), "google_error_estimate"
    element = data["rows"][0]["elements"][0]
    if element.get("status") != "OK":
        return estimate_drive_minutes_between(origin_lat, origin_lng, dest_lat, dest_lng), "google_error_estimate"
    seconds = element.get("duration_in_traffic", element.get("duration", {})).get("value")
    if seconds is None:
        return estimate_drive_minutes_between(origin_lat, origin_lng, dest_lat, dest_lng), "google_error_estimate"
    return max(0, int(round(seconds / 60))), "google"


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
) -> tuple[pd.DataFrame, dict[tuple[str, str], str]]:
    labels = [matrix_label(row) for _, row in group_rows.iterrows()]
    matrix_df = matrix_df.reindex(index=labels, columns=labels)
    if matrix_df.isna().any().any():
        matrix_df = matrix_df.fillna(999)
    out = pd.DataFrame(index=["ORIGIN", *labels], columns=["ORIGIN", *labels], dtype=float)
    out.loc["ORIGIN", "ORIGIN"] = 0
    out.loc[labels, "ORIGIN"] = 999
    out.loc[labels, labels] = matrix_df.values
    edge_sources: dict[tuple[str, str], str] = {}
    for label, (_, row) in zip(labels, group_rows.iterrows()):
        minutes, source = drive_duration_minutes_with_source(
            current_lat,
            current_lng,
            row.get("LAT"),
            row.get("LNG"),
            args,
            api_key,
        )
        out.loc["ORIGIN", label] = minutes
        edge_sources[("ORIGIN", label)] = source
    return out, edge_sources


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


def greedy_tsp_route(matrix_df: pd.DataFrame) -> tuple[list[int], float]:
    matrix = matrix_df.to_numpy()
    unvisited = set(range(1, len(matrix_df)))
    route = [0]
    total_cost = 0.0
    current = 0
    while unvisited:
        next_idx = min(unvisited, key=lambda idx: float(matrix[current, idx]))
        total_cost += float(matrix[current, next_idx])
        route.append(next_idx)
        unvisited.remove(next_idx)
        current = next_idx
    return route, total_cost


def route_cost_for_indices(matrix_df: pd.DataFrame, route: list[int]) -> float:
    matrix = matrix_df.to_numpy(dtype=float)
    return sum(float(matrix[route[i], route[i + 1]]) for i in range(len(route) - 1))


def two_opt_open_route(matrix_df: pd.DataFrame, route: list[int]) -> tuple[list[int], float]:
    if len(route) <= 3:
        return route, route_cost_for_indices(matrix_df, route)
    best_route = route[:]
    best_cost = route_cost_for_indices(matrix_df, best_route)
    improved = True
    while improved:
        improved = False
        for i in range(1, len(best_route) - 1):
            for j in range(i + 1, len(best_route)):
                candidate = best_route[:i] + best_route[i : j + 1][::-1] + best_route[j + 1 :]
                candidate_cost = route_cost_for_indices(matrix_df, candidate)
                if candidate_cost < best_cost:
                    best_route = candidate
                    best_cost = candidate_cost
                    improved = True
                    break
            if improved:
                break
    return best_route, best_cost


def two_opt_fixed_end_route(matrix_df: pd.DataFrame, route: list[int]) -> tuple[list[int], float]:
    if len(route) <= 4:
        return route, route_cost_for_indices(matrix_df, route)
    best_route = route[:]
    best_cost = route_cost_for_indices(matrix_df, best_route)
    improved = True
    while improved:
        improved = False
        for i in range(1, len(best_route) - 2):
            for j in range(i + 1, len(best_route) - 1):
                candidate = best_route[:i] + best_route[i : j + 1][::-1] + best_route[j + 1 :]
                candidate_cost = route_cost_for_indices(matrix_df, candidate)
                if candidate_cost < best_cost:
                    best_route = candidate
                    best_cost = candidate_cost
                    improved = True
                    break
            if improved:
                break
    return best_route, best_cost


def open_nearest_neighbor_two_opt_route(matrix_df: pd.DataFrame) -> tuple[list[int], float, str]:
    nn_route, nn_cost = greedy_tsp_route(matrix_df)
    opt_route, opt_cost = two_opt_open_route(matrix_df, nn_route)
    log = (
        "Using open nearest-neighbor + 2-opt route. "
        f"NN cost {nn_cost:.2f} min; 2-opt cost {opt_cost:.2f} min; "
        f"ending at {matrix_df.columns[opt_route[-1]]}."
    )
    return opt_route, opt_cost, log


def closed_route_to_warehouse_for_rows(
    matrix_folder: Path,
    group_value: Any,
    segment_rows: list[pd.Series],
    start_lat: Any,
    start_lng: Any,
    start_name: str,
    warehouse_lat: float,
    warehouse_lng: float,
    warehouse_name: str,
    service_minutes: float,
    args: argparse.Namespace,
    api_key: str | None,
) -> dict[str, Any] | None:
    if not segment_rows:
        return None

    labels = [matrix_label(row) for row in segment_rows]
    matrix_df = load_group_matrix(matrix_folder, group_value).reindex(index=labels, columns=labels)
    if matrix_df.isna().any().any():
        matrix_df = matrix_df.fillna(999)

    node_names = ["ORIGIN", *labels, "WAREHOUSE"]
    warehouse_idx = len(node_names) - 1
    closed_matrix = pd.DataFrame(index=node_names, columns=node_names, dtype=float)
    closed_matrix.loc[:, :] = 999
    for name in node_names:
        closed_matrix.loc[name, name] = 0

    edge_sources: dict[tuple[str, str], str] = {}
    for label, row in zip(labels, segment_rows):
        minutes, source = drive_duration_minutes_with_source(
            start_lat,
            start_lng,
            row.get("LAT"),
            row.get("LNG"),
            args,
            api_key,
        )
        closed_matrix.loc["ORIGIN", label] = minutes
        edge_sources[("ORIGIN", label)] = source

        minutes, source = drive_duration_minutes_with_source(
            row.get("LAT"),
            row.get("LNG"),
            warehouse_lat,
            warehouse_lng,
            args,
            api_key,
        )
        closed_matrix.loc[label, "WAREHOUSE"] = minutes
        edge_sources[(label, "WAREHOUSE")] = source

    closed_matrix.loc[labels, labels] = matrix_df.values

    unvisited = set(range(1, warehouse_idx))
    route = [0]
    current = 0
    while unvisited:
        next_idx = min(unvisited, key=lambda idx: float(closed_matrix.iloc[current, idx]))
        route.append(next_idx)
        unvisited.remove(next_idx)
        current = next_idx
    route.append(warehouse_idx)
    route, travel_time = two_opt_fixed_end_route(closed_matrix, route)

    rows_by_label = {matrix_label(row): row for row in segment_rows}
    route_names = [start_name]
    ordered_rows: list[pd.Series] = []
    travel_sources: list[str] = []
    for prev_idx, curr_idx in zip(route, route[1:]):
        prev_label = closed_matrix.columns[prev_idx]
        curr_label = closed_matrix.columns[curr_idx]
        travel_sources.append(edge_sources.get((prev_label, curr_label), "matrix"))
        if curr_label == "WAREHOUSE":
            route_names.append(warehouse_name)
            continue
        row = rows_by_label.get(str(curr_label))
        if row is None:
            continue
        ordered_rows.append(row)
        route_names.extend(route_parts_for_row(row))

    delivery_time = sum(package_count_from_row(row) for row in ordered_rows) * service_minutes
    return {
        "route_names": route_names,
        "ordered_rows": ordered_rows,
        "travel_time": travel_time,
        "delivery_time": delivery_time,
        "cost": travel_time + delivery_time,
        "endpoint": warehouse_name,
        "travel_time_source": " + ".join(dict.fromkeys(travel_sources)) if travel_sources else "",
        "start_name": start_name,
    }


def tsp_route_for_rows(
    deliveries: Any,
    matrix_folder: Path,
    group_value: Any,
    group_rows: pd.DataFrame,
    current_lat: Any,
    current_lng: Any,
    args: argparse.Namespace,
    api_key: str | None,
) -> tuple[list[int], pd.DataFrame, str, dict[tuple[str, str], str]]:
    matrix_df = load_group_matrix(matrix_folder, group_value)
    modified_matrix, edge_sources = add_current_origin_matrix(
        matrix_df,
        group_rows,
        current_lat,
        current_lng,
        args,
        api_key,
    )
    tsp_route, _tsp_cost, tsp_log = open_nearest_neighbor_two_opt_route(modified_matrix)
    return tsp_route, modified_matrix, tsp_log, edge_sources


def can_deliver_after_reload(
    deliveries: Any,
    remaining: pd.DataFrame,
    active_city: Any,
    pending_cluster_key: tuple[str, str] | None,
    warehouse_lat: float,
    warehouse_lng: float,
    driver_time_after_return: float,
    max_shift: float,
    max_packages: int,
    service_minutes: float,
    args: argparse.Namespace,
    api_key: str | None,
) -> bool:
    closest, _next_city, _next_pending = closest_next_row(
        remaining,
        warehouse_lat,
        warehouse_lng,
        active_city,
        pending_cluster_key,
    )
    if closest is None:
        return False
    folder_value, group_value = cluster_key_for_row(closest)
    group_rows = remaining[cluster_key_mask(remaining, (folder_value, group_value))].copy()
    if group_rows.empty:
        return False
    try:
        tsp_route, modified_matrix, _tsp_log, _edge_sources = tsp_route_for_rows(
            deliveries,
            Path(folder_value),
            group_value,
            group_rows,
            warehouse_lat,
            warehouse_lng,
            args,
            api_key,
        )
    except Exception:
        return False

    prev_idx = tsp_route[0]
    for curr_idx in tsp_route[1:]:
        label = modified_matrix.columns[curr_idx]
        match = group_rows[group_rows.apply(lambda row: matrix_label(row) == label, axis=1)]
        if match.empty:
            prev_idx = curr_idx
            continue
        row = match.iloc[0]
        packages = package_count_from_row(row)
        step_travel = float(modified_matrix.iloc[prev_idx, curr_idx])
        step_delivery = packages * service_minutes
        return driver_time_after_return + step_travel + step_delivery <= max_shift
    return False


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
    deliveries = import_script(ROOT / "deliveries_experimental.py", "routecraft_deliveries_experimental")
    output_path = run_dir / "06_delivery_plan.xlsx"
    api_key = None if args.mock_google else get_distance_key(args)
    warehouse = resolve_warehouse(args)
    warehouse_name = str(warehouse["name"])
    warehouse_lat = float(warehouse["lat"])
    warehouse_lng = float(warehouse["lng"])

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
    enable_deferred_reload_batches = bool(constraint_value(args, "enable_deferred_reload_batches", False))

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
            f"Warehouse: {warehouse_name}",
            f"Warehouse address: {warehouse['address']}",
            f"Warehouse coordinates: {warehouse_lat:.6f}, {warehouse_lng:.6f}",
            f"Warehouse coordinate source: {warehouse['source']}",
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
        current_lat = warehouse_lat
        current_lng = warehouse_lng
        current_name = warehouse_name
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
                if driver_time >= ideal_shift or deferred_pool.empty or not enable_deferred_reload_batches:
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
                previous_step = journey[-1] if journey else None
                if driver_time >= ideal_shift:
                    break
                if previous_step and previous_step.get("segment_type") == "delivery":
                    closed_segment = closed_route_to_warehouse_for_rows(
                        Path(previous_step.get("_matrix_folder", run_dir)),
                        previous_step.get("_group_value", ""),
                        previous_step.get("_segment_rows", []),
                        previous_step.get("_segment_start_lat"),
                        previous_step.get("_segment_start_lng"),
                        str(previous_step.get("_segment_start_name", warehouse_name)),
                        warehouse_lat,
                        warehouse_lng,
                        warehouse_name,
                        service_minutes,
                        args,
                        api_key,
                    )
                    if closed_segment is not None:
                        open_segment_cost = float(previous_step["cost"])
                        driver_time_after_closed_return = driver_time - open_segment_cost + float(closed_segment["cost"])
                        if driver_time_after_closed_return >= max_shift:
                            break
                        if not can_deliver_after_reload(
                            deliveries,
                            remaining,
                            active_city,
                            pending_cluster_key,
                            warehouse_lat,
                            warehouse_lng,
                            driver_time_after_closed_return,
                            max_shift,
                            max_packages,
                            service_minutes,
                            args,
                            api_key,
                        ):
                            log_lines.append(
                                f"Driver {driver_id} could return to {warehouse_name}, but would not have enough "
                                "time to complete another delivery after reloading. Keeping the open route and ending "
                                "at the current stop."
                            )
                            break

                        previous_step["route_str"] = " ג†’ ".join(closed_segment["route_names"])
                        previous_step["travel_time"] = closed_segment["travel_time"]
                        previous_step["delivery_time"] = closed_segment["delivery_time"]
                        previous_step["cost"] = closed_segment["cost"]
                        previous_step["endpoint"] = warehouse_name
                        previous_step["segment_type"] = "delivery_closed_reload"
                        previous_step["travel_time_source"] = closed_segment["travel_time_source"]
                        driver_time = driver_time_after_closed_return
                        current_lat = warehouse_lat
                        current_lng = warehouse_lng
                        current_name = warehouse_name
                        load_number += 1
                        load_packages = 0
                        log_lines.append(
                            f"Driver {driver_id} capacity is full, so the last delivery segment was re-optimized "
                            f"as a closed route back to {warehouse_name}. Shift time: {driver_time:.2f} min."
                        )
                        continue

                return_to_warehouse, return_source = drive_duration_minutes_with_source(
                    current_lat,
                    current_lng,
                    warehouse_lat,
                    warehouse_lng,
                    args,
                    api_key,
                )
                if driver_time >= ideal_shift or driver_time + return_to_warehouse >= max_shift:
                    break
                if not can_deliver_after_reload(
                    deliveries,
                    remaining,
                    active_city,
                    pending_cluster_key,
                    warehouse_lat,
                    warehouse_lng,
                    driver_time + return_to_warehouse,
                    max_shift,
                    max_packages,
                    service_minutes,
                    args,
                    api_key,
                ):
                    log_lines.append(
                        f"Driver {driver_id} has enough time to return to {warehouse_name} "
                        f"({return_to_warehouse:.2f} min, source: {return_source}) but not enough "
                        "time to complete another delivery after reloading. Ending shift at current stop."
                    )
                    break
                driver_time += return_to_warehouse
                journey.append(
                    {
                        "load": load_number,
                        "cluster": "RELOAD",
                        "city": "Warehouse",
                        "route_str": f"{current_name} → {warehouse_name}",
                        "travel_time": return_to_warehouse,
                        "delivery_time": 0,
                        "cost": return_to_warehouse,
                        "endpoint": warehouse_name,
                        "packages": 0,
                        "segment_type": "reload",
                        "travel_time_source": return_source,
                    }
                )
                log_lines.append(
                    f"Driver {driver_id} returning to {warehouse_name} to reload: {return_to_warehouse:.2f} min. "
                    f"Source: {return_source}."
                )
                current_lat = warehouse_lat
                current_lng = warehouse_lng
                current_name = warehouse_name
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

            tsp_route, modified_matrix, tsp_log, edge_sources = tsp_route_for_rows(
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
            segment_start_lat = current_lat
            segment_start_lng = current_lng
            segment_start_name = current_name
            route_names = [current_name]
            travel_time = 0.0
            delivery_time = 0.0
            segment_rows: list[pd.Series] = []
            delivered_ids: list[int] = []
            travel_sources: list[str] = []
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
                prev_label = modified_matrix.columns[prev_idx]
                travel_sources.append(edge_sources.get((prev_label, label), "matrix"))
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
                    "travel_time_source": " + ".join(dict.fromkeys(travel_sources)) if travel_sources else "",
                    "_matrix_folder": folder_value,
                    "_group_value": group_value,
                    "_segment_rows": segment_rows,
                    "_segment_start_lat": segment_start_lat,
                    "_segment_start_lng": segment_start_lng,
                    "_segment_start_name": segment_start_name,
                }
            )
            remaining = remaining[~remaining["_route_id"].isin(delivered_ids)].copy()
            log_lines.append(f"Route: {' → '.join(route_names)}")
            log_lines.append(
                f"Travel: {travel_time:.2f} min | Delivery: {delivery_time:.2f} min | "
                f"Source: {' + '.join(dict.fromkeys(travel_sources)) if travel_sources else 'n/a'} | "
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
                "Clusters Visited": len(
                    [step for step in data["journey"] if str(step.get("segment_type", "")).startswith("delivery")]
                ),
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
                    "Travel Time Source": step.get("travel_time_source", ""),
                    "Delivery Time": f"{step['delivery_time']:.2f}",
                    "Route Path": join_route(split_route_path(step["route_str"])),
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
                    "Warehouse Name": warehouse_name,
                    "Warehouse Address": warehouse["address"],
                    "Warehouse Latitude": warehouse_lat,
                    "Warehouse Longitude": warehouse_lng,
                    "Warehouse Coordinate Source": warehouse["source"],
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


def delivery_iteration_delivered_selected(iteration_dir: Path, deferred_count: int) -> bool:
    leftover_path = iteration_dir / "06_leftover_orders_next_run.xlsx"
    if not leftover_path.exists():
        return False
    leftover_count = len(pd.read_excel(leftover_path).fillna(""))
    return leftover_count <= deferred_count


def delivery_iteration_metrics(
    iteration_dir: Path,
    selected_count: int,
    deferred_count: int,
    args: argparse.Namespace,
) -> dict[str, Any]:
    workbook_path = iteration_dir / "06_delivery_plan.xlsx"
    leftover_path = iteration_dir / "06_leftover_orders_next_run.xlsx"
    max_shift = float(constraint_value(args, "max_shift_minutes", 720))
    driver_times: list[float] = []
    if workbook_path.exists():
        try:
            summary_df = pd.read_excel(workbook_path, sheet_name="Summary").fillna("")
            if "Total Time (min)" in summary_df.columns:
                driver_times = [
                    float(value)
                    for value in pd.to_numeric(summary_df["Total Time (min)"], errors="coerce").dropna().tolist()
                ]
        except Exception:
            driver_times = []
    leftover_count = len(pd.read_excel(leftover_path).fillna("")) if leftover_path.exists() else selected_count + deferred_count
    selected_leftover = max(0, leftover_count - deferred_count)
    total_slack = sum(max(0.0, max_shift - minutes) for minutes in driver_times)
    return {
        "driver_times": driver_times,
        "leftover_count": leftover_count,
        "selected_leftover": selected_leftover,
        "total_slack": total_slack,
    }


def smart_success_step(
    ordered_orders: pd.DataFrame,
    selected_count: int,
    metrics: dict[str, Any],
    args: argparse.Namespace,
    warehouse: dict[str, Any],
) -> int:
    min_step = max(1, int(constraint_value(args, "smart_min_step", 5)))
    max_step = max(min_step, int(constraint_value(args, "smart_max_step", 30)))
    remaining_rows = max(0, len(ordered_orders) - selected_count)
    if remaining_rows <= 0:
        return 0
    current_estimate = estimate_selected_rows(ordered_orders, selected_count, args, warehouse)
    average_minutes = max(
        float(constraint_value(args, "service_minutes_per_package", 4))
        + float(constraint_value(args, "estimated_city_stop_minutes", 3)),
        current_estimate.estimated_total_minutes / max(1, selected_count),
    )
    slack_based = int(metrics.get("total_slack", 0) / average_minutes)
    if slack_based <= 0:
        return min(min_step, remaining_rows)
    return min(max(min_step, slack_based), max_step, remaining_rows)


def smart_failure_step(selected_count: int, metrics: dict[str, Any], args: argparse.Namespace) -> int:
    min_step = max(1, int(constraint_value(args, "smart_min_step", 5)))
    max_step = max(min_step, int(constraint_value(args, "smart_max_step", 30)))
    selected_leftover = int(metrics.get("selected_leftover", 0) or 0)
    if selected_leftover > 0:
        return min(max(selected_leftover, min_step), max_step, selected_count - 1)
    fallback = max(min_step, int(round(selected_count * 0.1)))
    return min(fallback, max_step, selected_count - 1)


def next_unattempted_count(
    preferred: int,
    attempted: set[int],
    lower_bound: int,
    upper_bound: int,
) -> int | None:
    preferred = max(lower_bound, min(preferred, upper_bound))
    if preferred not in attempted:
        return preferred
    for distance in range(1, upper_bound - lower_bound + 1):
        for candidate in (preferred + distance, preferred - distance):
            if lower_bound <= candidate <= upper_bound and candidate not in attempted:
                return candidate
    return None


def copy_final_successful_output(run_dir: Path, best_count: int, output_path: Path, summary_path: Path) -> Path:
    final_dir = run_dir / f"final_successful_output_{best_count:04d}"
    final_dir.mkdir(parents=True, exist_ok=True)
    source_dir = output_path.parent
    for filename in [
        "06_delivery_plan.xlsx",
        "06_delivery_plan.html",
        "02b_selected_orders_for_run.xlsx",
        "02c_deferred_orders_next_run.xlsx",
    ]:
        source = source_dir / filename
        if source.exists():
            (final_dir / filename).write_bytes(source.read_bytes())
    if summary_path.exists():
        (final_dir / summary_path.name).write_bytes(summary_path.read_bytes())
    return final_dir


def run_delivery_iteration(
    ordered_orders: pd.DataFrame,
    selected_count: int,
    run_dir: Path,
    args: argparse.Namespace,
    attempt_number: int | None = None,
) -> tuple[Path, StageResult, bool, int, Path]:
    selected_count = max(1, min(selected_count, len(ordered_orders)))
    if attempt_number is None:
        iteration_dir = run_dir / f"iterative_selected-{selected_count:04d}"
    else:
        iteration_dir = run_dir / f"attempt-{attempt_number:03d}_selected-{selected_count:04d}"
    iteration_dir.mkdir(parents=True, exist_ok=True)
    selected, deferred = split_orders_by_priority_and_closeness(ordered_orders, selected_count, args)
    selected_path = iteration_dir / "02b_selected_orders_for_run.xlsx"
    deferred_path = iteration_dir / "02c_deferred_orders_next_run.xlsx"
    selected.to_excel(selected_path, index=False)
    deferred.to_excel(deferred_path, index=False)

    clustered_path, cluster_result = stage_3_cluster(selected_path, iteration_dir, args)
    grouped_path, group_result = stage_4_group(clustered_path, iteration_dir, args)
    matrix_paths, matrix_result = stage_5_distance_matrices(grouped_path, iteration_dir, args)
    output_path, delivery_result = stage_6_delivery_plan(grouped_path, iteration_dir, iteration_dir, args)

    selected_delivered = delivery_iteration_delivered_selected(iteration_dir, len(deferred))
    output_files = [
        selected_path,
        deferred_path,
        *cluster_result.output_files,
        *group_result.output_files,
        *matrix_result.output_files,
        *delivery_result.output_files,
    ]
    status = "completed" if selected_delivered else "selected-leftover"
    result = StageResult(
        6,
        f"Iterative delivery attempt with {selected_count} selected orders",
        status,
        output_files,
        [
            f"Selected {selected_count}; deferred {len(deferred)}.",
            "All selected orders delivered." if selected_delivered else "Some selected orders were left over.",
        ],
    )
    return output_path, result, selected_delivered, len(deferred), iteration_dir


def run_iterative_delivery_search(
    geocoded_path: Path,
    run_dir: Path,
    args: argparse.Namespace,
) -> tuple[Path, list[StageResult]]:
    ordered_orders = sort_orders_for_selection(pd.read_excel(geocoded_path).fillna(""), args)
    if ordered_orders.empty:
        raise RuntimeError("Iterative delivery search has no geocoded orders to route.")

    smart_enabled = constraint_bool(args, "smart_iterative_selection", True)
    step = int(constraint_value(args, "iterative_address_step", 10))
    if step <= 0:
        step = 10
    max_iterations = int(constraint_value(args, "iterative_max_iterations", 25))
    warehouse = resolve_warehouse(args)
    initial_estimate: SmartSelectionEstimate | None = None
    if smart_enabled:
        selected_count, initial_estimate = find_smart_initial_selection_count(ordered_orders, args, warehouse)
    else:
        selected_count = min(initial_selection_capacity(args), len(ordered_orders))
    attempted: set[int] = set()
    results: list[StageResult] = []
    best_success: tuple[int, Path] | None = None
    failure_bound: int | None = None
    last_output_path: Path | None = None
    decision_lines: list[str] = []
    stop_gap = max(1, int(constraint_value(args, "smart_stop_gap", 1)))
    if smart_enabled and initial_estimate is not None:
        decision_lines.extend(smart_estimate_notes(initial_estimate))

    for _ in range(max_iterations):
        if selected_count in attempted:
            break
        attempted.add(selected_count)
        output_path, result, success, _deferred_count, _iteration_dir = run_delivery_iteration(
            ordered_orders,
            selected_count,
            run_dir,
            args,
            len(results) + 1,
        )
        last_output_path = output_path
        metrics = delivery_iteration_metrics(_iteration_dir, selected_count, _deferred_count, args)
        driver_times_text = ", ".join(f"{minutes:.0f}" for minutes in metrics["driver_times"]) or "unknown"
        result.notes.append(
            f"Driver minutes: [{driver_times_text}]; selected leftovers: {metrics['selected_leftover']}; slack minutes: {metrics['total_slack']:.0f}."
        )
        results.append(result)

        if not smart_enabled:
            if success:
                best_success = (selected_count, output_path)
                next_count = min(selected_count + step, len(ordered_orders))
                if next_count == selected_count:
                    break
                if failure_bound is not None and next_count >= failure_bound:
                    break
                selected_count = next_count
            else:
                failure_bound = selected_count if failure_bound is None else min(failure_bound, selected_count)
                next_count = selected_count - step
                if next_count < 1:
                    break
                selected_count = next_count
            continue

        if success:
            best_success = (selected_count, output_path)
            if failure_bound is not None:
                gap = failure_bound - selected_count
                decision_lines.append(f"{selected_count} succeeded; narrowing bracket {selected_count}-{failure_bound}.")
                if gap <= stop_gap:
                    break
                next_count = selected_count + max(1, gap // 2)
            else:
                smart_step = smart_success_step(ordered_orders, selected_count, metrics, args, warehouse)
                if smart_step <= 0:
                    break
                next_count = selected_count + smart_step
                decision_lines.append(f"{selected_count} succeeded; adding {smart_step} rows from driver slack.")
            next_attempt = next_unattempted_count(next_count, attempted, selected_count + 1, len(ordered_orders))
            if next_attempt is None:
                break
            if failure_bound is not None and next_attempt >= failure_bound:
                next_attempt = next_unattempted_count(
                    selected_count + max(1, (failure_bound - selected_count) // 2),
                    attempted,
                    selected_count + 1,
                    failure_bound - 1,
                )
                if next_attempt is None:
                    break
            selected_count = next_attempt
        else:
            failure_bound = selected_count if failure_bound is None else min(failure_bound, selected_count)
            if best_success is not None:
                gap = failure_bound - best_success[0]
                decision_lines.append(f"{selected_count} failed; narrowing bracket {best_success[0]}-{failure_bound}.")
                if gap <= stop_gap:
                    break
                next_count = best_success[0] + max(1, gap // 2)
                next_attempt = next_unattempted_count(next_count, attempted, best_success[0] + 1, failure_bound - 1)
            else:
                remove_step = smart_failure_step(selected_count, metrics, args)
                next_count = selected_count - remove_step
                decision_lines.append(f"{selected_count} failed; removing {remove_step} rows from selected leftovers.")
                next_attempt = next_unattempted_count(next_count, attempted, 1, selected_count - 1)
            if next_attempt is None:
                break
            selected_count = next_attempt

    summary_path = run_dir / "iterative_delivery_summary.txt"
    lines = [
        "Smart iterative delivery search" if smart_enabled else "Iterative delivery search",
        f"Mode: {'smart estimate + bracket search' if smart_enabled else 'fixed step'}",
        f"Fallback/fixed step size: {step}",
        f"Smart stop gap: {stop_gap}" if smart_enabled else "",
        f"Attempts: {len(results)}",
        f"Best successful selected count: {best_success[0] if best_success else 'none'}",
        "",
    ]
    lines.extend(line for line in decision_lines if line)
    if decision_lines:
        lines.append("")
    for result in results:
        lines.append(f"{result.name}: {result.status}")
        lines.extend(f"  - {note}" for note in result.notes)
    summary_path.write_text("\n".join(lines), encoding="utf-8")
    results.append(
        StageResult(
            6,
            "Iterative delivery search summary",
            "completed" if best_success else "no-successful-selection",
            [summary_path],
            [f"Best successful selected count: {best_success[0] if best_success else 'none'}"],
        )
    )
    if best_success is None:
        if last_output_path is None:
            raise RuntimeError("Iterative delivery search did not complete any attempts.")
        return last_output_path, results
    final_dir = copy_final_successful_output(run_dir, best_success[0], best_success[1], summary_path)
    results.append(
        StageResult(
            6,
            "Final successful output",
            "completed",
            [final_dir],
            [f"Copied best successful attempt ({best_success[0]} selected rows) into {final_dir.name}."],
        )
    )
    return best_success[1], results


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
            if full_route_nodes and route_nodes and full_route_nodes[-1] == route_nodes[0]:
                full_route_nodes.extend(route_nodes[1:])
            else:
                full_route_nodes.extend(route_nodes)
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
                f"<span>Source {html.escape(str(row.get('Travel Time Source', '')) or 'n/a')}</span>"
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
            f"<span>Full route ({len(full_route_nodes)} stops shown)</span>"
            "<strong>Start to finish, repeated deliveries preserved</strong>"
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
    .full-route {{ display: grid; grid-template-columns: 1fr; gap: 6px; padding: 0; margin: 12px 0 0; list-style: none; }}
    .full-route li {{ display: grid; grid-template-columns: 36px 1fr; align-items: start; gap: 10px; border: 1px solid #cfd8e3; border-radius: 6px; background: #fbfcfd; padding: 8px 10px; }}
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


def split_route_path(route_path: Any) -> list[str]:
    text = str(route_path)
    text = text.replace("׳’ג€ ג€™", ROUTE_SEPARATOR)
    text = text.replace("ג†’", ROUTE_SEPARATOR)
    text = text.replace("→", ROUTE_SEPARATOR)
    return [part.strip() for part in text.split(ROUTE_SEPARATOR) if part.strip()]


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
                import_script(ROOT / "TSP_experimental.py", f"routecraft_tsp_experimental_{module_suffix}")

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
    parser.add_argument(
        "--disable-iterative-delivery",
        action="store_true",
        help="Run stages 3-6 once instead of searching for the largest deliverable selected set.",
    )
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
        if args.start_stage <= 2 and not args.disable_iterative_delivery:
            current, iterative_results = run_iterative_delivery_search(
                run_dir / "02_geocoded_addresses.xlsx",
                run_dir,
                args,
            )
            stage_results.extend(iterative_results)
        else:
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
