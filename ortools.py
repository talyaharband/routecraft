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
from dataclasses import dataclass, field
from datetime import datetime, time
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
CONSTRAINTS_PARAMETERS_PATH = ROOT / "constraints_parameters.json"


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
    path = MOCK_DATA_DIR / "geocoding_cache.xlsx"
    if not path.exists():
        return {}
    df = pd.read_excel(path).fillna("")
    cache = {}
    for _, row in df.iterrows():
        cache[(str(row.get("query", "")), str(row.get("city", "")))] = row.to_dict()
    return cache


def load_mock_coords_cache() -> dict[str, dict[str, float]]:
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


def mock_coords_from_cache(address: str, _key: str = "") -> tuple[dict[str, float], str]:
    cached = load_mock_coords_cache().get(address)
    if cached:
        return cached, "mock cache"
    return mock_coords(address, _key)


def mock_matrix_cache_path(group_value: Any) -> Path:
    safe_group = re.sub(r'[<>:"/\\|?*]+', "_", str(group_value).strip())
    safe_group = re.sub(r"\s+", "_", safe_group) or "unknown"
    return MOCK_DATA_DIR / "distance_matrices" / f"05_distance_matrix_group-{safe_group}.xlsx"


def save_real_mock_data(run_dir: Path) -> list[Path]:
    MOCK_DATA_DIR.mkdir(parents=True, exist_ok=True)
    saved_paths: list[Path] = []

    geocode_path = run_dir / "02_geocoding_cache.xlsx"
    if geocode_path.exists():
        path = MOCK_DATA_DIR / "geocoding_cache.xlsx"
        pd.read_excel(geocode_path).to_excel(path, index=False)
        saved_paths.append(path)

    coords_path = run_dir / "05_distance_coords_cache.xlsx"
    if coords_path.exists():
        path = MOCK_DATA_DIR / "distance_coords_cache.xlsx"
        pd.read_excel(coords_path).to_excel(path, index=False)
        saved_paths.append(path)

    matrix_dir = MOCK_DATA_DIR / "distance_matrices"
    matrix_dir.mkdir(parents=True, exist_ok=True)
    for matrix_path in sorted(run_dir.glob("05_distance_matrix_group-*.xlsx")):
        target = matrix_dir / matrix_path.name
        pd.read_excel(matrix_path, index_col=0).to_excel(target)
        saved_paths.append(target)

    origin_source = run_dir / "06_origin_duration_cache.xlsx"
    if origin_source.exists():
        origin_target = MOCK_DATA_DIR / "origin_duration_cache.xlsx"
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


def stage_1_cleanup(input_path: Path, run_dir: Path, args: argparse.Namespace) -> tuple[Path, StageResult]:
    cleanup = import_script(ROOT / "data.cleanup.py", "routecraft_data_cleanup")
    paths = cleanup.clean_raw_orders(input_path, run_dir)
    failed_path = paths["failed"]
    shaped_path = paths["good_addresses"]
    original_path = paths["good_original"]

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
    geocode_cache = load_mock_geocode_cache() if args.mock_google else {}
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
        if args.mock_google:
            cached = geocode_cache.get((query, city))
            if cached:
                lat = cached.get("lat")
                lng = cached.get("lng")
                is_precise = bool(cached.get("is_precise", True))
                status = str(cached.get("status", "mock cached precise match"))
            else:
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

    pd.DataFrame(rows).to_excel(output_path, index=False)
    if geocode_cache_rows:
        pd.DataFrame(geocode_cache_rows).to_excel(run_dir / "02_geocoding_cache.xlsx", index=False)
    failure_path = record_geocode_failures(run_dir, failed_rows)
    output_files = [output_path, run_dir / "02_geocoding_cache.xlsx"] + ([failure_path] if failure_path else [])
    notes = []
    if failed_rows:
        notes.append(
            f"{len(failed_rows)} addresses were moved to the error file because Google could not accurately find them."
        )
    return output_path, StageResult(2, "Geocode addresses", "completed", output_files, notes)


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
            if row["Street_Name"] == anchor["Street_Name"] and distance <= args.cluster_threshold_meters:
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
    for address in unique_addresses:
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
        if args.mock_google and cached_matrix_path.exists():
            matrix_df = pd.read_excel(cached_matrix_path, index_col=0)
            matrix_df = matrix_df.reindex(index=full_addresses, columns=full_addresses)
            if matrix_df.isna().any().any():
                matrix_df = distance.build_matrix_for_addresses(full_addresses, coords_dict)
        else:
            matrix_df = distance.build_matrix_for_addresses(full_addresses, coords_dict)
        matrix_df.to_excel(matrix_path)
        matrix_paths.append(matrix_path)

    return matrix_paths, StageResult(5, "Build distance matrices", "completed", [coords_cache_path, *matrix_paths])


def truthy(value: Any) -> bool:
    text = clean_cell(value).strip().lower()
    return text in {"1", "true", "yes", "y", "priority", "prioritized", "high", "urgent"}


def parse_minutes_value(value: Any, shift_start: str = "08:00") -> int | None:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return int(round(float(value)))
    if isinstance(value, pd.Timestamp):
        target_time = value.time()
    elif isinstance(value, datetime):
        target_time = value.time()
    elif isinstance(value, time):
        target_time = value
    else:
        text = clean_cell(value)
        if not text:
            return None
        numeric = pd.to_numeric(text, errors="coerce")
        if pd.notna(numeric):
            return int(round(float(numeric)))
        parsed = pd.to_datetime(text, errors="coerce", dayfirst=True)
        if pd.isna(parsed):
            return None
        target_time = parsed.time()

    start = pd.to_datetime(shift_start, format="%H:%M", errors="coerce")
    if pd.isna(start):
        start = pd.to_datetime("08:00", format="%H:%M")
    start_minutes = int(start.hour * 60 + start.minute)
    target_minutes = int(target_time.hour * 60 + target_time.minute)
    if target_minutes < start_minutes:
        target_minutes += 24 * 60
    return target_minutes - start_minutes


def parse_planning_date(value: Any) -> pd.Timestamp:
    if value is None or clean_cell(value).lower() in {"", "today"}:
        return pd.Timestamp.today().normalize()
    text = clean_cell(value)
    if re.fullmatch(r"\d{4}-\d{1,2}-\d{1,2}", text):
        parsed = pd.to_datetime(text, errors="coerce", dayfirst=False)
    else:
        parsed = pd.to_datetime(value, errors="coerce", dayfirst=True)
    if pd.isna(parsed):
        raise RuntimeError(f"Could not parse planning_date from constraints file: {value}")
    return pd.Timestamp(parsed).normalize()


def parse_due_date(value: Any) -> pd.Timestamp | None:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return None
    text = clean_cell(value)
    if not text:
        return None
    if re.fullmatch(r"\d{4}-\d{1,2}-\d{1,2}", text):
        parsed = pd.to_datetime(text, errors="coerce", dayfirst=False)
    else:
        parsed = pd.to_datetime(value, errors="coerce", dayfirst=True)
    if pd.isna(parsed):
        return None
    return pd.Timestamp(parsed).normalize()


def looks_like_time_of_day(value: Any) -> bool:
    if isinstance(value, time):
        return True
    if isinstance(value, datetime):
        return False
    if isinstance(value, pd.Timestamp):
        return False
    text = clean_cell(value)
    return bool(re.fullmatch(r"\d{1,2}:\d{2}(:\d{2})?", text))


def due_date_penalty(due_date: pd.Timestamp | None, args: argparse.Namespace) -> tuple[int, str, bool]:
    if due_date is None:
        return int(args.no_due_date_penalty), "no_due_date", False
    planning_date = parse_planning_date(args.planning_date)
    days_until_due = int((due_date - planning_date).days)
    if days_until_due < 0:
        return int(args.overdue_penalty), "overdue", bool(args.due_today_hard)
    if days_until_due == 0:
        return int(args.due_today_penalty), "due_today", bool(args.due_today_hard)
    if days_until_due == 1:
        return int(args.due_tomorrow_penalty), "due_tomorrow", False
    return int(args.future_due_penalty), f"due_in_{days_until_due}_days", False


def delivery_count_from_row(row: pd.Series) -> int:
    value = pd.to_numeric(row.get("total_orders_in_cluster", ""), errors="coerce")
    if pd.notna(value) and int(value) > 0:
        return int(value)
    detailed = clean_cell(row.get("detailed_addresses", ""))
    if detailed:
        return max(1, len([part for part in detailed.split(",") if part.strip()]))
    return 1


def stop_label(row: pd.Series) -> str:
    street = clean_cell(row.get("Street_Name", ""))
    house = clean_cell(row.get("House_Number", "")).replace(".0", "")
    return f"{street} {house}".strip()


def stop_route_parts(row: pd.Series) -> list[str]:
    detailed = clean_cell(row.get("detailed_addresses", ""))
    if detailed:
        return [part.strip() for part in detailed.split(",") if part.strip()]
    label = stop_label(row)
    return [label] if label else []


def build_ortools_stops(input_path: Path, args: argparse.Namespace) -> list[dict[str, Any]]:
    df = pd.read_excel(input_path).fillna("")
    required = {"City", "Street_Name", "House_Number", "LAT", "LNG"}
    missing = required.difference(df.columns)
    if missing:
        raise RuntimeError(f"OR-Tools stage 6 input is missing columns: {', '.join(sorted(missing))}")

    stops = [
        {
            "label": "WAREHOUSE ESHTAOL",
            "city": "Warehouse",
            "cluster": "DEPOT",
            "lat": 31.77927525,
            "lng": 35.0105885,
            "packages": 0,
            "service_minutes": 0,
            "route_parts": ["WAREHOUSE ESHTAOL"],
            "deadline_minutes": None,
            "due_date": "",
            "urgency": "depot",
            "skip_penalty": 0,
            "force_delivery": True,
            "is_priority": False,
        }
    ]

    priority_deadline = parse_minutes_value(args.prioritized_arrival_time, args.shift_start_time)
    for _, row in df.iterrows():
        packages = delivery_count_from_row(row)
        is_priority = args.priority_column in df.columns and truthy(row.get(args.priority_column))
        due_date = None
        urgency = "no_due_date"
        skip_penalty = int(args.no_due_date_penalty)
        force_delivery = False
        deadline = None
        if args.arrival_deadline_column in df.columns:
            deadline_value = row.get(args.arrival_deadline_column)
            if looks_like_time_of_day(deadline_value):
                deadline = parse_minutes_value(deadline_value, args.shift_start_time)
            due_date = parse_due_date(deadline_value)
            skip_penalty, urgency, force_delivery = due_date_penalty(due_date, args)
        if deadline is None and is_priority:
            deadline = priority_deadline
            if skip_penalty == int(args.no_due_date_penalty):
                skip_penalty = int(args.due_today_penalty)
                urgency = "priority"

        stops.append(
            {
                "label": stop_label(row),
                "city": clean_cell(row.get("City", "")),
                "cluster": clean_cell(row.get("cluster_group", "")),
                "lat": float(row.get("LAT")),
                "lng": float(row.get("LNG")),
                "packages": packages,
                "service_minutes": int(round(packages * args.service_minutes_per_package)),
                "route_parts": stop_route_parts(row),
                "deadline_minutes": deadline,
                "due_date": due_date.strftime("%Y-%m-%d") if due_date is not None else "",
                "urgency": urgency,
                "skip_penalty": skip_penalty,
                "force_delivery": force_delivery,
                "is_priority": is_priority or deadline is not None or force_delivery,
            }
        )
    if len(stops) <= 1:
        raise RuntimeError("OR-Tools stage 6 found no delivery stops.")
    return stops


def estimate_drive_minutes(origin: dict[str, Any], destination: dict[str, Any]) -> int:
    if origin is destination:
        return 0
    meters = calculate_haversine(origin["lat"], origin["lng"], destination["lat"], destination["lng"])
    road_factor = 1.35
    average_kmh = 42
    minutes = (meters * road_factor) / (average_kmh * 1000 / 60)
    return max(1, int(round(minutes)))


def google_duration_matrix(stops: list[dict[str, Any]], api_key: str) -> list[list[int]]:
    matrix: list[list[int]] = []
    destination_chunk_size = 25
    destinations = [f"{stop['lat']},{stop['lng']}" for stop in stops]
    for origin in stops:
        row_minutes: list[int] = []
        origin_text = f"{origin['lat']},{origin['lng']}"
        for start in range(0, len(stops), destination_chunk_size):
            chunk = destinations[start : start + destination_chunk_size]
            response = requests.get(
                "https://maps.googleapis.com/maps/api/distancematrix/json",
                params={
                    "origins": origin_text,
                    "destinations": "|".join(chunk),
                    "mode": "driving",
                    "departure_time": "now",
                    "key": api_key,
                },
                timeout=30,
            )
            response.raise_for_status()
            data = response.json()
            if data.get("status") != "OK":
                raise RuntimeError(f"Google Distance Matrix status: {data.get('status')}")
            elements = data["rows"][0]["elements"]
            for element in elements:
                if element.get("status") != "OK":
                    row_minutes.append(999)
                    continue
                seconds = element.get("duration_in_traffic", element.get("duration", {})).get("value")
                row_minutes.append(max(0, int(round((seconds or 0) / 60))))
        matrix.append(row_minutes)
    for i in range(len(matrix)):
        matrix[i][i] = 0
    return matrix


def build_ortools_time_matrix(stops: list[dict[str, Any]], run_dir: Path, args: argparse.Namespace) -> list[list[int]]:
    if args.mock_google:
        matrix = [[estimate_drive_minutes(origin, destination) for destination in stops] for origin in stops]
    else:
        matrix = google_duration_matrix(stops, get_distance_key(args))
    pd.DataFrame(matrix, index=[s["label"] for s in stops], columns=[s["label"] for s in stops]).to_excel(
        run_dir / "06_ortools_time_matrix.xlsx"
    )
    return matrix


def load_google_ortools_modules() -> tuple[Any, Any]:
    """Import Google's ortools package even though this script is named ortools.py."""
    import importlib

    original_path = without_local_ortools_shadow()
    try:
        pywrapcp = importlib.import_module("ortools.constraint_solver.pywrapcp")
        routing_enums_pb2 = importlib.import_module("ortools.constraint_solver.routing_enums_pb2")
        return pywrapcp, routing_enums_pb2
    finally:
        sys.path = original_path


def solve_ortools_routes(
    stops: list[dict[str, Any]],
    time_matrix: list[list[int]],
    args: argparse.Namespace,
) -> tuple[dict[int, dict[str, Any]], list[int]]:
    pywrapcp, routing_enums_pb2 = load_google_ortools_modules()
    manager = pywrapcp.RoutingIndexManager(len(stops), args.drivers, 0)
    routing = pywrapcp.RoutingModel(manager)

    def transit_callback(from_index: int, to_index: int) -> int:
        from_node = manager.IndexToNode(from_index)
        to_node = manager.IndexToNode(to_index)
        return int(time_matrix[from_node][to_node] + stops[from_node]["service_minutes"])

    transit_index = routing.RegisterTransitCallback(transit_callback)
    routing.SetArcCostEvaluatorOfAllVehicles(transit_index)
    routing.AddDimension(
        transit_index,
        0,
        int(args.max_shift_minutes),
        True,
        "Time",
    )
    time_dimension = routing.GetDimensionOrDie("Time")
    for vehicle_id in range(args.drivers):
        end_index = routing.End(vehicle_id)
        time_dimension.SetCumulVarSoftUpperBound(end_index, int(args.ideal_shift_minutes), 25)

    for node, stop in enumerate(stops[1:], start=1):
        if stop["deadline_minutes"] is not None:
            index = manager.NodeToIndex(node)
            time_dimension.CumulVar(index).SetMax(int(stop["deadline_minutes"]))

    def demand_callback(from_index: int) -> int:
        return int(stops[manager.IndexToNode(from_index)]["packages"])

    demand_index = routing.RegisterUnaryTransitCallback(demand_callback)
    routing.AddDimensionWithVehicleCapacity(
        demand_index,
        0,
        [int(args.max_packages_per_driver)] * args.drivers,
        True,
        "Packages",
    )

    if not args.require_all_deliveries:
        for node, stop in enumerate(stops[1:], start=1):
            if stop.get("force_delivery"):
                continue
            penalty = int(stop.get("skip_penalty", 25_000)) + int(stop["packages"]) * 1_000
            routing.AddDisjunction([manager.NodeToIndex(node)], penalty)

    search_parameters = pywrapcp.DefaultRoutingSearchParameters()
    search_parameters.first_solution_strategy = routing_enums_pb2.FirstSolutionStrategy.PARALLEL_CHEAPEST_INSERTION
    search_parameters.local_search_metaheuristic = routing_enums_pb2.LocalSearchMetaheuristic.GUIDED_LOCAL_SEARCH
    search_parameters.time_limit.FromSeconds(int(args.ortools_time_limit_seconds))

    solution = routing.SolveWithParameters(search_parameters)
    if solution is None:
        raise RuntimeError(
            "OR-Tools could not find a feasible plan. Try more drivers, a higher max shift, "
            "a higher package limit, or omit --require-all-deliveries."
        )

    routes: dict[int, dict[str, Any]] = {}
    delivered_nodes: set[int] = set()
    for vehicle_id in range(args.drivers):
        index = routing.Start(vehicle_id)
        route_nodes: list[int] = []
        while not routing.IsEnd(index):
            node = manager.IndexToNode(index)
            route_nodes.append(node)
            if node != 0:
                delivered_nodes.add(node)
            index = solution.Value(routing.NextVar(index))
        route_nodes.append(manager.IndexToNode(index))
        route_time = solution.Value(time_dimension.CumulVar(index))
        packages = sum(stops[node]["packages"] for node in route_nodes if node != 0)
        routes[vehicle_id + 1] = {"nodes": route_nodes, "time": route_time, "packages": packages}

    unassigned = [node for node in range(1, len(stops)) if node not in delivered_nodes]
    return routes, unassigned


def write_ortools_delivery_plan(
    stops: list[dict[str, Any]],
    time_matrix: list[list[int]],
    routes: dict[int, dict[str, Any]],
    unassigned: list[int],
    output_path: Path,
    args: argparse.Namespace,
) -> Path:
    summary_rows = []
    detail_rows = []
    for driver_id, route in routes.items():
        node_sequence = [node for node in route["nodes"] if node != 0]
        if not node_sequence:
            continue

        summary_rows.append(
            {
                "Driver": f"Driver {driver_id}",
                "Addresses Delivered": sum(stops[node]["packages"] for node in node_sequence),
                "Clusters Visited": len({stops[node]["cluster"] for node in node_sequence}),
                "Total Time (min)": f"{route['time']:.2f}",
                "Total Time (hours)": f"{route['time'] / 60:.2f}",
                "Mode": "OR-Tools Mock" if args.mock_google else "OR-Tools Production",
            }
        )

        previous = 0
        cumulative = 0
        step = 1
        segment_nodes: list[int] = []
        current_key: tuple[str, str] | None = None
        segment_travel = 0
        segment_delivery = 0

        def flush_segment() -> None:
            nonlocal step, segment_nodes, segment_travel, segment_delivery, cumulative, current_key
            if not segment_nodes:
                return
            route_parts = ["WAREHOUSE ESHTAOL"] if step == 1 else []
            for node in segment_nodes:
                route_parts.extend(stops[node]["route_parts"])
            total = segment_travel + segment_delivery
            cumulative += total
            last_node = segment_nodes[-1]
            detail_rows.append(
                {
                    "Driver": f"Driver {driver_id}",
                    "Step": step,
                    "City": stops[last_node]["city"],
                    "Cluster": stops[last_node]["cluster"],
                    "Total Time (min)": f"{total:.2f}",
                    "Travel Time": f"{segment_travel:.2f}",
                    "Delivery Time": f"{segment_delivery:.2f}",
                    "Route Path": " → ".join(route_parts),
                    "Endpoint": stops[last_node]["route_parts"][-1] if stops[last_node]["route_parts"] else stops[last_node]["label"],
                    "Shift Time So Far": f"{cumulative:.2f}",
                }
            )
            step += 1
            segment_nodes = []
            current_key = None
            segment_travel = 0
            segment_delivery = 0

        for node in node_sequence:
            key = (stops[node]["city"], stops[node]["cluster"])
            travel = time_matrix[previous][node]
            if current_key is not None and key != current_key:
                flush_segment()
            current_key = key
            segment_nodes.append(node)
            segment_travel += travel
            segment_delivery += stops[node]["service_minutes"]
            previous = node
        flush_segment()

    unassigned_rows = [
        {
            "Stop": stops[node]["label"],
            "City": stops[node]["city"],
            "Cluster": stops[node]["cluster"],
            "Packages": stops[node]["packages"],
            "Priority": stops[node]["is_priority"],
            "Due Date": stops[node].get("due_date", ""),
            "Urgency": stops[node].get("urgency", ""),
            "Skip Penalty": stops[node].get("skip_penalty", ""),
            "Deadline Minutes": stops[node]["deadline_minutes"],
        }
        for node in unassigned
    ]

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        pd.DataFrame(summary_rows).to_excel(writer, sheet_name="Summary", index=False)
        pd.DataFrame(detail_rows).to_excel(writer, sheet_name="Detailed Routes", index=False)
        pd.DataFrame(unassigned_rows or [{"Info": "All deliveries assigned"}]).to_excel(
            writer, sheet_name="Unassigned", index=False
        )
    return output_path


def stage_6_delivery_plan(
    input_path: Path,
    matrix_folder: Path,
    run_dir: Path,
    args: argparse.Namespace,
) -> tuple[Path, StageResult]:
    output_path = run_dir / "06_delivery_plan.xlsx"
    stops = build_ortools_stops(input_path, args)
    time_matrix = build_ortools_time_matrix(stops, run_dir, args)
    routes, unassigned = solve_ortools_routes(stops, time_matrix, args)
    write_ortools_delivery_plan(stops, time_matrix, routes, unassigned, output_path, args)

    log_path = run_dir / "06_delivery_plan.log"
    log_path.write_text(
        "\n".join(
            [
                "OR-Tools delivery planning completed.",
                f"Drivers: {args.drivers}",
                f"Max packages per driver: {args.max_packages_per_driver}",
                f"Ideal shift minutes: {args.ideal_shift_minutes}",
                f"Max shift minutes: {args.max_shift_minutes}",
                f"Planning date: {parse_planning_date(args.planning_date).date()}",
                f"Due today hard: {args.due_today_hard}",
                f"Unassigned stops: {len(unassigned)}",
            ]
        ),
        encoding="utf-8",
    )
    html_path = render_delivery_plan_html(output_path, run_dir / "06_delivery_plan.html")
    output_files = [output_path, html_path, log_path, run_dir / "06_ortools_time_matrix.xlsx"]
    return output_path, StageResult(
        6,
        "Plan multi-driver deliveries with OR-Tools",
        "completed",
        output_files,
        [f"{len(unassigned)} stops were left unassigned."] if unassigned else [],
    )


def render_delivery_plan_html(workbook_path: Path, output_path: Path) -> Path:
    summary_df = pd.read_excel(workbook_path, sheet_name="Summary")
    detail_df = pd.read_excel(workbook_path, sheet_name="Detailed Routes")
    order_details = load_order_details(output_path.parent / "01c_good_orders_original_format.xlsx")

    summary_cards = []
    if not summary_df.empty:
        totals = {
            "Drivers": len(summary_df),
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
        for _, row in driver_df.iterrows():
            route_nodes = [
                str(part).strip()
                for part in str(row.get("Route Path", "")).split("→")
                if str(part).strip()
            ]
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
        driver_sections.append(
            "<section class=\"driver\">"
            f"<h2>{html.escape(str(driver))}</h2>"
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
    .step-head {{ display: flex; justify-content: space-between; gap: 12px; flex-wrap: wrap; }}
    .step-head span {{ color: #65717d; }}
    .times {{ display: flex; gap: 8px; flex-wrap: wrap; margin: 12px 0; }}
    .times span {{ background: #edf2f6; padding: 6px 9px; border-radius: 999px; font-size: 13px; }}
    .route {{ display: flex; flex-wrap: wrap; gap: 8px; padding: 0; margin: 0; list-style: none; }}
    .route li {{ border: 1px solid #cfd8e3; padding: 7px 9px; border-radius: 6px; background: #fbfcfd; max-width: 360px; }}
    .route li strong {{ display: block; }}
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
    constraints = load_constraints_parameters()
    parser = argparse.ArgumentParser(description="Run the Routecraft pipeline with OR-Tools delivery planning.")
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
    parser.add_argument(
        "--drivers",
        type=int,
        default=int(constraints.get("drivers", 2)),
        help="Maximum number of available drivers/vehicles.",
    )
    parser.add_argument(
        "--max-packages-per-driver",
        type=int,
        default=int(constraints.get("max_packages_per_driver", 10_000)),
        help="Maximum packages one driver can carry on a route.",
    )
    parser.add_argument(
        "--prioritized-arrival-time",
        default=constraints.get("prioritized_arrival_time"),
        help="Target latest arrival for prioritized deliveries, as minutes after shift start or HH:MM.",
    )
    parser.add_argument(
        "--planning-date",
        default=str(constraints.get("planning_date", "today")),
        help="Dispatch planning date for due-date urgency. Use YYYY-MM-DD or 'today'.",
    )
    parser.add_argument(
        "--due-today-hard",
        action=argparse.BooleanOptionalAction,
        default=bool(constraints.get("due_today_hard", True)),
        help="Force overdue and due-today deliveries to be assigned when feasible.",
    )
    parser.add_argument("--overdue-penalty", type=int, default=int(constraints.get("overdue_penalty", 2_000_000)))
    parser.add_argument("--due-today-penalty", type=int, default=int(constraints.get("due_today_penalty", 1_500_000)))
    parser.add_argument("--due-tomorrow-penalty", type=int, default=int(constraints.get("due_tomorrow_penalty", 300_000)))
    parser.add_argument("--future-due-penalty", type=int, default=int(constraints.get("future_due_penalty", 75_000)))
    parser.add_argument("--no-due-date-penalty", type=int, default=int(constraints.get("no_due_date_penalty", 25_000)))
    parser.add_argument(
        "--priority-column",
        default=str(constraints.get("priority_column", "priority")),
        help="Input column whose truthy values mark prioritized deliveries.",
    )
    parser.add_argument(
        "--arrival-deadline-column",
        default=str(constraints.get("arrival_deadline_column", "arrival_date")),
        help="Input column with per-delivery deadline minutes or time values.",
    )
    parser.add_argument(
        "--shift-start-time",
        default=str(constraints.get("shift_start_time", "08:00")),
        help="Shift start time used to interpret HH:MM deadlines.",
    )
    parser.add_argument(
        "--ideal-shift-minutes",
        type=float,
        default=float(constraints.get("ideal_shift_minutes", 420)),
        help="Preferred shift length. OR-Tools may exceed this up to --max-shift-minutes.",
    )
    parser.add_argument(
        "--max-shift-minutes",
        type=float,
        default=float(constraints.get("max_shift_minutes", 480)),
        help="Hard maximum shift length.",
    )
    parser.add_argument(
        "--service-minutes-per-package",
        type=float,
        default=float(constraints.get("service_minutes_per_package", 4)),
        help="On-site service time added per package.",
    )
    parser.add_argument(
        "--ortools-time-limit-seconds",
        type=int,
        default=int(constraints.get("ortools_time_limit_seconds", 60)),
        help="Solver search time limit.",
    )
    parser.add_argument(
        "--require-all-deliveries",
        action="store_true",
        default=bool(constraints.get("require_all_deliveries", False)),
        help="Fail if constraints cannot assign every delivery stop.",
    )
    return parser.parse_args()


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
            stage_results.append(
                StageResult(
                    5,
                    "Skip legacy per-cluster distance matrices",
                    "completed",
                    [],
                    ["OR-Tools builds one global time matrix during stage 6."],
                )
            )
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
