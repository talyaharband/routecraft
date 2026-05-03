from __future__ import annotations

import argparse
import contextlib
import hashlib
import html
import importlib.util
import importlib.machinery
import io
import os
import re
import sys
import types
import traceback
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
    df = pd.read_excel(input_path)
    if df.shape[1] < 3:
        raise RuntimeError("Stage 2 input must have at least 3 columns: city, street, house number.")

    city_col, street_col, number_col = df.columns[:3]
    rows = []
    failed_rows = []
    for _, row in df.iterrows():
        city = clean_cell(row[city_col])
        street = clean_cell(row[street_col])
        house_number = clean_cell(row[number_col])
        query = f"{street} {house_number}, {city}".strip()
        if args.mock_google:
            lat, lng = mock_lat_lng(query)
            is_precise = True
            status = "mock precise match"
        else:
            lat, lng, is_precise, status = geocode_precise_address(query, google_key, city)

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
            }
        )

    output_path = run_dir / "02_geocoded_addresses.xlsx"
    if not rows:
        record_geocode_failures(run_dir, failed_rows)
        raise RuntimeError("Stage 2 did not find any addresses with precise Google geocoding matches.")

    pd.DataFrame(rows).to_excel(output_path, index=False)
    failure_path = record_geocode_failures(run_dir, failed_rows)
    output_files = [output_path] + ([failure_path] if failure_path else [])
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

    from k_means_constrained import KMeansConstrained

    df["cluster_group"] = pd.NA
    next_group = 0
    notes = []
    for city, city_df in df.groupby("City", sort=False):
        city_indices = city_df.index
        coords = city_df[["LAT", "LNG"]].values
        num_addresses = len(city_df)
        if num_addresses <= args.max_group_size:
            df.loc[city_indices, "cluster_group"] = next_group
            notes.append(f"{city}: {num_addresses} rows fit in group {next_group}.")
            next_group += 1
            continue

        n_clusters = int(np.ceil(num_addresses / args.max_group_size))
        clf = KMeansConstrained(
            n_clusters=n_clusters,
            size_min=1,
            size_max=args.max_group_size,
            random_state=42,
        )
        local_labels = clf.fit_predict(coords)
        label_map = {label: next_group + offset for offset, label in enumerate(sorted(set(local_labels)))}
        df.loc[city_indices, "cluster_group"] = [label_map[label] for label in local_labels]
        notes.append(f"{city}: created groups {next_group}-{next_group + n_clusters - 1}.")
        next_group += n_clusters

    df["cluster_group"] = df["cluster_group"].astype(int)

    output_path = run_dir / "04_clustered_delivery_groups.xlsx"
    df.to_excel(output_path, index=False)
    return output_path, StageResult(4, "Create delivery groups", "completed", [output_path], notes)


def stage_5_distance_matrices(input_path: Path, run_dir: Path, args: argparse.Namespace) -> tuple[list[Path], StageResult]:
    distance = import_script(ROOT / "distance.matrix.py", "routecraft_distance_matrix")
    if args.mock_google:
        distance.API_KEY = "mock-google"
        distance.get_coords = mock_coords
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
    for address in unique_addresses:
        coords, _status = distance.get_coords(address, distance.API_KEY)
        coords_dict[address] = coords

    matrix_paths = []
    for group_value, group_df in grouped_frames.items():
        full_addresses = distance.build_full_addresses(group_df)
        matrix_df = distance.build_matrix_for_addresses(full_addresses, coords_dict)
        safe_group = distance.sanitize_group_value(group_value)
        matrix_path = run_dir / f"05_distance_matrix_group-{safe_group}.xlsx"
        matrix_df.to_excel(matrix_path)
        matrix_paths.append(matrix_path)

    return matrix_paths, StageResult(5, "Build distance matrices", "completed", matrix_paths)


def stage_6_delivery_plan(input_path: Path, matrix_folder: Path, run_dir: Path) -> tuple[Path, StageResult]:
    deliveries = import_script(ROOT / "deliveries.py", "routecraft_deliveries")
    output_path = run_dir / "06_delivery_plan.xlsx"
    stdout = io.StringIO()
    stderr = io.StringIO()
    with contextlib.redirect_stdout(stdout), contextlib.redirect_stderr(stderr):
        deliveries.run_multi_cluster_tsp(
            kmeans_path=input_path,
            matrix_folder=matrix_folder,
            output_path=output_path,
            test_mode=True,
        )
    log_path = run_dir / "06_delivery_plan.log"
    log_path.write_text(stdout.getvalue(), encoding="utf-8")
    if stderr.getvalue():
        (run_dir / "06_delivery_plan.err.log").write_text(stderr.getvalue(), encoding="utf-8")
    html_path = render_delivery_plan_html(output_path, run_dir / "06_delivery_plan.html")
    return output_path, StageResult(
        6,
        "Plan multi-driver deliveries",
        "completed",
        [output_path, html_path, log_path],
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
    parser = argparse.ArgumentParser(description="Run the Routecraft pipeline without file pickers.")
    parser.add_argument("--input", required=True, help="Input Excel file for the selected start stage.")
    parser.add_argument("--start-stage", type=int, default=1, choices=range(1, 7))
    parser.add_argument("--runs-dir", default=str(RUNS_DIR))
    parser.add_argument("--google-api-key")
    parser.add_argument("--distance-api-key")
    parser.add_argument(
        "--mock-google",
        action="store_true",
        help="Use deterministic local mocks for stage 2 geocoding and stage 5 distance matrices.",
    )
    parser.add_argument("--max-workers", type=int, default=10)
    parser.add_argument("--cluster-threshold-meters", type=float, default=150)
    parser.add_argument("--max-group-size", type=int, default=30)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    input_path = require_file(Path(args.input), "Input file")
    load_env_file(ROOT / ".env")
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
            current, result = stage_6_delivery_plan(delivery_input_path, matrix_folder, run_dir)
            stage_results.append(result)

        print(f"Flow completed. Run folder: {run_dir}")
        print(f"Final report: {current}")
        return 0
    except BaseException as exc:
        report_path = render_failure_html(run_dir, stage_results, exc)
        print(f"Flow failed. Report: {report_path}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
