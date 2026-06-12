from __future__ import annotations

import json
import math
import mimetypes
import os
import re
import shutil
import subprocess
import sys
import threading
import time
import uuid
import zipfile
from dataclasses import dataclass, field
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse

import pandas as pd


ROOT = Path(__file__).resolve().parent
WEB_DIR = ROOT / "web"
STATIC_DIR = WEB_DIR / "static"
if not STATIC_DIR.exists():
    STATIC_DIR = WEB_DIR / "data tor desktop screens"
MOBILE_DIR = ROOT / "mobile"
MOBILE_STATIC_DIR = MOBILE_DIR / "static"
JOBS_DIR = ROOT / ".tmp" / "web_jobs"
WEB_RUNS_DIR = ROOT / "runs" / "web_app"
DEFAULT_CONSTRAINTS_PATH = ROOT / "constraints_parameters_experimental.json"
PIPELINE_SCRIPT = ROOT / "const_experimental.py"
DEMO_RESULTS_ZIP = ROOT / "runs" / "results_for_web" / "final_successful_output_0510.zip"
DEMO_RESULTS_DIR = ROOT / ".tmp" / "results_for_web" / "final_successful_output_0510"
COMPARISON_RESULTS_ROOT = ROOT / "runs" / "results_for_comparison"
DISPLAY_DRIVER_NAMES = [
    "דביר לוי",
    "שמעיה סבן",
    "נהוראי מלצר",
    "יוסי אביטן",
    "איתיאל סופר",
    "יעל רבינוביץ",
]

JOBS_DIR.mkdir(parents=True, exist_ok=True)
WEB_RUNS_DIR.mkdir(parents=True, exist_ok=True)
DEMO_RESULTS_DIR.mkdir(parents=True, exist_ok=True)


@dataclass
class JobState:
    id: str
    status: str = "queued"
    created_at: float = field(default_factory=time.time)
    updated_at: float = field(default_factory=time.time)
    input_name: str = ""
    drivers: int = 0
    capacity: int = 0
    job_dir: Path = Path()
    run_dir: Path | None = None
    final_report: Path | None = None
    returncode: int | None = None
    log: list[str] = field(default_factory=list)
    error: str = ""
    result: dict | None = None

    def public(self) -> dict:
        return {
            "id": self.id,
            "status": self.status,
            "created_at": self.created_at,
            "updated_at": self.updated_at,
            "input_name": self.input_name,
            "drivers": self.drivers,
            "capacity": self.capacity,
            "run_dir": str(self.run_dir) if self.run_dir else "",
            "final_report": str(self.final_report) if self.final_report else "",
            "returncode": self.returncode,
            "log": self.log[-80:],
            "error": self.error,
            "result": self.result,
        }


jobs: dict[str, JobState] = {}
jobs_lock = threading.Lock()


def clean_cell(value) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def safe_int(value: str, default: int = 0) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def json_response(handler: BaseHTTPRequestHandler, payload: dict, status: int = 200) -> None:
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.end_headers()
    handler.wfile.write(body)


def read_bytes(path: Path) -> bytes:
    return path.read_bytes()


def parse_multipart(body: bytes, content_type: str) -> tuple[dict[str, str], dict[str, dict]]:
    match = re.search(r"boundary=(?P<boundary>[^;]+)", content_type)
    if not match:
        raise ValueError("Missing multipart boundary")
    boundary = match.group("boundary").strip().strip('"').encode("utf-8")
    fields: dict[str, str] = {}
    files: dict[str, dict] = {}
    for part in body.split(b"--" + boundary):
        part = part.strip()
        if not part or part == b"--":
            continue
        if part.endswith(b"--"):
            part = part[:-2].strip()
        if b"\r\n\r\n" not in part:
            continue
        raw_headers, content = part.split(b"\r\n\r\n", 1)
        if content.endswith(b"\r\n"):
            content = content[:-2]
        headers = raw_headers.decode("utf-8", errors="replace").split("\r\n")
        disposition = next((line for line in headers if line.lower().startswith("content-disposition:")), "")
        name_match = re.search(r'name="([^"]+)"', disposition)
        if not name_match:
            continue
        name = name_match.group(1)
        filename_match = re.search(r'filename="([^"]*)"', disposition)
        if filename_match:
            filename = Path(filename_match.group(1)).name or "upload.xlsx"
            files[name] = {"filename": filename, "content": content}
        else:
            fields[name] = content.decode("utf-8", errors="replace").strip()
    return fields, files


def append_log(job: JobState, line: str) -> None:
    with jobs_lock:
        job.log.append(line.rstrip())
        job.updated_at = time.time()


def load_base_constraints() -> dict:
    if not DEFAULT_CONSTRAINTS_PATH.exists():
        return {}
    return json.loads(DEFAULT_CONSTRAINTS_PATH.read_text(encoding="utf-8"))


def create_constraints_file(job: JobState) -> Path:
    constraints = load_base_constraints()
    constraints["drivers"] = job.drivers
    constraints["max_packages_per_driver"] = job.capacity
    path = job.job_dir / "constraints.json"
    path.write_text(json.dumps(constraints, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def extract_path_from_log(lines: list[str], label: str) -> Path | None:
    prefix = f"{label}:"
    for line in reversed(lines):
        if line.startswith(prefix):
            value = line.split(":", 1)[1].strip()
            if value:
                return Path(value)
    return None


def split_route_path(value: str) -> list[str]:
    parts = [clean_cell(part) for part in str(value).split(" -> ")]
    return [part for part in parts if part]


def normalize_lookup_key(value) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip()).casefold()


def parse_float(value) -> float | None:
    number = pd.to_numeric(value, errors="coerce")
    if pd.isna(number):
        return None
    return float(number)


def sheet_rows(path: Path, sheet_name: str) -> list[dict]:
    try:
        return pd.read_excel(path, sheet_name=sheet_name).fillna("").to_dict("records")
    except Exception:
        return []


def count_rows(path: Path) -> int:
    if not path.exists():
        return 0
    try:
        return len(pd.read_excel(path))
    except Exception:
        return 0


def build_coordinate_lookup(paths: list[Path]) -> dict[str, dict]:
    lookup: dict[str, dict] = {}
    for path in paths:
        if not path.exists():
            continue
        try:
            rows = pd.read_excel(path).fillna("").to_dict("records")
        except Exception:
            continue
        for row in rows:
            lat = parse_float(row.get("LAT"))
            lng = parse_float(row.get("LNG"))
            if lat is None or lng is None:
                continue
            city = clean_cell(row.get("City", ""))
            street = clean_cell(row.get("Street_Name", ""))
            house = clean_cell(row.get("House_Number", "")).replace(".0", "")
            address = " ".join(part for part in [street, house] if part)
            full_label = ", ".join(part for part in [address, city] if part)
            point = {
                "label": address or full_label,
                "city": city,
                "lat": lat,
                "lng": lng,
            }
            for key in [
                normalize_lookup_key(address),
                normalize_lookup_key(full_label),
                normalize_lookup_key(f"{address}|{city}"),
            ]:
                if key and key not in lookup:
                    lookup[key] = point
    return lookup


def route_point_for_label(label: str, city: str, lookup: dict[str, dict], warehouse: dict) -> dict | None:
    clean_label = clean_cell(label)
    if not clean_label:
        return None
    if clean_label == clean_cell(warehouse.get("name", "")):
        return warehouse
    for key in [
        normalize_lookup_key(f"{clean_label}|{city}"),
        normalize_lookup_key(clean_label),
        normalize_lookup_key(f"{clean_label}, {city}"),
    ]:
        point = lookup.get(key)
        if point:
            return {
                "label": clean_label,
                "city": point.get("city", city),
                "lat": point["lat"],
                "lng": point["lng"],
            }
    return {"label": clean_label, "city": city, "lat": None, "lng": None}


def route_points_for_detail(detail: dict, lookup: dict[str, dict], warehouse: dict) -> list[dict]:
    city = clean_cell(detail.get("City", ""))
    points = []
    for label in split_route_path(clean_cell(detail.get("Route Path", ""))):
        point = route_point_for_label(label, city, lookup, warehouse)
        if point:
            points.append(point)
    return points


def address_label(street, house, city: str = "") -> str:
    house_text = clean_cell(house).replace(".0", "")
    street_text = clean_cell(street)
    city_text = clean_cell(city)
    address = " ".join(part for part in [street_text, house_text] if part)
    return ", ".join(part for part in [address, city_text] if part)


def boss_delivery_points(path: Path) -> dict[str, list[dict]]:
    if not path.exists():
        return {}
    try:
        rows = pd.read_excel(path).fillna("").to_dict("records")
    except Exception:
        return {}
    rows.sort(
        key=lambda row: (
            clean_cell(row.get("driver", "")),
            safe_int(clean_cell(row.get("round", ""))),
            safe_int(clean_cell(row.get("delivery_number_in_round", ""))),
        )
    )
    by_driver: dict[str, list[dict]] = {}
    for row in rows:
        driver = clean_cell(row.get("driver", ""))
        if not driver:
            continue
        lat = parse_float(row.get("lat"))
        lng = parse_float(row.get("lng"))
        by_driver.setdefault(driver, []).append(
            {
                "label": address_label(row.get("street_name", ""), row.get("house_number", ""), row.get("city", "")),
                "city": clean_cell(row.get("city", "")),
                "lat": lat,
                "lng": lng,
                "order_id": clean_cell(row.get("order_id", "")),
                "round": safe_int(clean_cell(row.get("round", ""))),
                "stop_number": safe_int(clean_cell(row.get("delivery_number_in_round", ""))),
            }
        )
    return by_driver


def delivery_point_lookup(points: list[dict]) -> dict[str, dict]:
    lookup: dict[str, dict] = {}
    for point in points:
        label = clean_cell(point.get("label", ""))
        city = clean_cell(point.get("city", ""))
        if not label:
            continue
        for key in [
            normalize_lookup_key(label),
            normalize_lookup_key(f"{label}|{city}"),
            normalize_lookup_key(f"{label}, {city}"),
        ]:
            if key and key not in lookup:
                lookup[key] = point
    return lookup


def demo_route_points_for_detail(
    detail: dict,
    lookup: dict[str, dict],
    warehouse: dict,
    has_started: bool,
) -> list[dict]:
    city = clean_cell(detail.get("City", ""))
    segment_type = clean_cell(detail.get("Segment Type", ""))
    points = []
    for label in split_route_path(clean_cell(detail.get("Route Path", ""))):
        if label == clean_cell(warehouse.get("name", "")):
            stop_type = "warehouse_start" if not has_started and not points else "warehouse_reload"
            points.append(
                {
                    "label": warehouse["label"],
                    "city": warehouse["city"],
                    "lat": warehouse["lat"],
                    "lng": warehouse["lng"],
                    "stop_type": stop_type,
                    "segment_type": segment_type,
                }
            )
            has_started = True
            continue
        point = None
        for key in [
            normalize_lookup_key(f"{label}|{city}"),
            normalize_lookup_key(label),
            normalize_lookup_key(f"{label}, {city}"),
        ]:
            point = lookup.get(key)
            if point:
                break
        if point:
            points.append(
                {
                    **point,
                    "stop_type": "delivery",
                    "segment_type": segment_type,
                }
            )
        else:
            points.append(
                {
                    "label": label,
                    "city": city,
                    "lat": None,
                    "lng": None,
                    "stop_type": "delivery",
                    "segment_type": segment_type,
                }
            )
        has_started = True
    return points


def demo_driver_route_points(delivered_points: list[dict], warehouse: dict) -> list[dict]:
    points = []
    current_round: int | None = None
    for point in delivered_points:
        round_number = safe_int(clean_cell(point.get("round", "")), 0)
        if round_number != current_round:
            current_round = round_number
            points.append(
                {
                    "label": warehouse["label"],
                    "city": warehouse["city"],
                    "lat": warehouse["lat"],
                    "lng": warehouse["lng"],
                    "stop_type": "warehouse_start" if len(points) == 0 else "warehouse_reload",
                    "round": round_number,
                }
            )
        points.append(
            {
                **point,
                "stop_type": "delivery",
            }
        )
    return points


def count_excel_rows(path: Path) -> int:
    try:
        return len(pd.read_excel(path).fillna(""))
    except Exception:
        return 0


def ensure_demo_results_dir() -> Path:
    if not DEMO_RESULTS_ZIP.exists():
        raise FileNotFoundError(f"Saved results ZIP not found: {DEMO_RESULTS_ZIP}")
    marker = DEMO_RESULTS_DIR / ".extracted_from_zip"
    needs_extract = not (DEMO_RESULTS_DIR / "06_delivery_plan.xlsx").exists()
    if marker.exists():
        try:
            needs_extract = needs_extract or marker.stat().st_mtime < DEMO_RESULTS_ZIP.stat().st_mtime
        except OSError:
            needs_extract = True
    if needs_extract:
        if DEMO_RESULTS_DIR.exists():
            shutil.rmtree(DEMO_RESULTS_DIR)
        DEMO_RESULTS_DIR.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(DEMO_RESULTS_ZIP) as archive:
            archive.extractall(DEMO_RESULTS_DIR)
        marker.write_text(str(time.time()), encoding="utf-8")
    return DEMO_RESULTS_DIR


def latest_comparison_results_dir() -> Path:
    if not COMPARISON_RESULTS_ROOT.exists():
        raise FileNotFoundError(f"Comparison results folder not found: {COMPARISON_RESULTS_ROOT}")
    candidates = [
        path
        for path in COMPARISON_RESULTS_ROOT.rglob("*")
        if path.is_dir() and (path / "06_delivery_plan.xlsx").exists()
    ]
    if not candidates and (COMPARISON_RESULTS_ROOT / "06_delivery_plan.xlsx").exists():
        candidates = [COMPARISON_RESULTS_ROOT]
    if not candidates:
        raise FileNotFoundError(f"No generated delivery plan found under: {COMPARISON_RESULTS_ROOT}")
    return max(candidates, key=lambda path: path.stat().st_mtime)


def parse_pipeline_result(job: JobState) -> dict:
    final_report = job.final_report
    if final_report is None or not final_report.exists():
        raise RuntimeError("Pipeline did not produce a final delivery workbook.")
    run_dir = job.run_dir or final_report.parent
    attempt_dir = final_report.parent

    summary_rows = sheet_rows(final_report, "Summary")
    detail_rows = sheet_rows(final_report, "Detailed Routes")
    constraint_rows = sheet_rows(final_report, "Run Constraints")
    constraints = constraint_rows[0] if constraint_rows else {}
    geocoded_path = run_dir / "02_geocoded_addresses.xlsx"
    failed_path = run_dir / "01a_failed_addresses.xlsx"
    original_path = run_dir / "01c_good_orders_original_format.xlsx"
    deferred_path = attempt_dir / "02c_deferred_orders_next_run.xlsx"
    leftover_path = attempt_dir / "06_leftover_orders_next_run.xlsx"
    selected_path = attempt_dir / "02b_selected_orders_for_run.xlsx"
    html_path = attempt_dir / "06_delivery_plan.html"
    warehouse = {
        "label": clean_cell(constraints.get("Warehouse Name", "WAREHOUSE ESHTAOL")) or "WAREHOUSE ESHTAOL",
        "name": clean_cell(constraints.get("Warehouse Name", "WAREHOUSE ESHTAOL")) or "WAREHOUSE ESHTAOL",
        "city": clean_cell(constraints.get("Warehouse Address", "")),
        "lat": parse_float(constraints.get("Warehouse Latitude")),
        "lng": parse_float(constraints.get("Warehouse Longitude")),
    }
    coordinate_lookup = build_coordinate_lookup([selected_path, geocoded_path])

    drivers = []
    detail_by_driver: dict[str, list[dict]] = {}
    for row in detail_rows:
        detail_by_driver.setdefault(clean_cell(row.get("Driver", "")), []).append(row)

    for row in summary_rows:
        driver_name = clean_cell(row.get("Driver", ""))
        details = detail_by_driver.get(driver_name, [])
        stops: list[str] = []
        segments = []
        driver_route_points = []
        driver_delivery_points = []
        for detail in details:
            route = split_route_path(clean_cell(detail.get("Route Path", "")))
            if not stops:
                stops.extend(route)
            elif route:
                stops.extend(route[1:] if stops[-1] == route[0] else route)
            route_points = route_points_for_detail(detail, coordinate_lookup, warehouse)
            for point in route_points:
                if not driver_route_points or (
                    driver_route_points[-1].get("label") != point.get("label")
                    or driver_route_points[-1].get("lat") != point.get("lat")
                    or driver_route_points[-1].get("lng") != point.get("lng")
                ):
                    driver_route_points.append(point)
                if point.get("label") != warehouse.get("label") and point.get("lat") is not None and point.get("lng") is not None:
                    driver_delivery_points.append(point)
            segments.append(
                {
                    "step": clean_cell(detail.get("Step", "")),
                    "city": clean_cell(detail.get("City", "")),
                    "packages": safe_int(clean_cell(detail.get("Packages", ""))),
                    "segment_type": clean_cell(detail.get("Segment Type", "")),
                    "total_time_min": clean_cell(detail.get("Total Time (min)", "")),
                    "route": route,
                    "route_points": route_points,
                }
            )
        unique_stops = [stop for stop in stops if stop and stop != "WAREHOUSE ESHTAOL"]
        city_names = []
        for detail in details:
            city = clean_cell(detail.get("City", ""))
            if city and city not in city_names:
                city_names.append(city)
        drivers.append(
            {
                "name": driver_name,
                "addresses": safe_int(clean_cell(row.get("Addresses Delivered", ""))),
                "clusters": safe_int(clean_cell(row.get("Clusters Visited", ""))),
                "total_minutes": float(pd.to_numeric(row.get("Total Time (min)", 0), errors="coerce") or 0),
                "total_hours": clean_cell(row.get("Total Time (hours)", "")),
                "shift_type": clean_cell(row.get("Shift Type", "")),
                "mode": clean_cell(row.get("Mode", "")),
                "cities": city_names,
                "stops": unique_stops,
                "route_points": driver_route_points,
                "delivery_points": driver_delivery_points,
                "segments": segments,
            }
        )

    total_delivered = sum(driver["addresses"] for driver in drivers)
    result = {
        "summary": {
            "drivers": safe_int(clean_cell(constraints.get("Drivers", job.drivers)), job.drivers),
            "capacity": safe_int(clean_cell(constraints.get("Max Packages Per Driver", job.capacity)), job.capacity),
            "uploaded_orders": count_rows(original_path) + count_rows(failed_path),
            "route_ready": count_rows(geocoded_path),
            "selected_orders": count_rows(selected_path),
            "deferred_orders": count_rows(deferred_path),
            "leftover_orders": safe_int(clean_cell(constraints.get("Leftover Orders", "")), count_rows(leftover_path)),
            "delivered_addresses": total_delivered,
            "regular_shifts": sum(1 for driver in drivers if driver["shift_type"].lower() == "regular"),
            "extended_shifts": sum(1 for driver in drivers if driver["shift_type"].lower() != "regular"),
            "max_shift_minutes": clean_cell(constraints.get("Max Shift Minutes", "")),
            "ideal_shift_minutes": clean_cell(constraints.get("Ideal Shift Minutes", "")),
            "warehouse": clean_cell(constraints.get("Warehouse Name", "")),
        },
        "drivers": drivers,
        "files": {
            "delivery_plan_xlsx": file_url(job, final_report),
            "delivery_plan_html": file_url(job, html_path) if html_path.exists() else "",
            "failed_addresses": file_url(job, failed_path) if failed_path.exists() else "",
            "leftover_orders": file_url(job, leftover_path) if leftover_path.exists() else "",
            "selected_orders": file_url(job, selected_path) if selected_path.exists() else "",
            "run_folder": str(run_dir),
        },
    }
    return result


def parse_saved_result(result_dir: Path, file_prefix: str, source: str) -> dict:
    final_report = result_dir / "06_delivery_plan.xlsx"
    if not final_report.exists():
        raise RuntimeError(f"Saved results folder does not contain 06_delivery_plan.xlsx: {result_dir}")

    summary_rows = sheet_rows(final_report, "Summary")
    detail_rows = sheet_rows(final_report, "Detailed Routes")
    constraint_rows = sheet_rows(final_report, "Run Constraints")
    constraints = constraint_rows[0] if constraint_rows else {}
    delivered_path = result_dir / "07_boss_delivered_addresses.xlsx"
    failed_path = result_dir / "07_boss_failed_addresses.xlsx"
    not_in_run_path = result_dir / "07_boss_not_in_run_addresses.xlsx"
    selected_path = result_dir / "04_clustered_delivery_groups.xlsx"
    html_path = result_dir / "06_delivery_plan.html"
    delivered_by_driver = boss_delivery_points(delivered_path)
    warehouse = {
        "label": clean_cell(constraints.get("Warehouse Name", "WAREHOUSE ESHTAOL")) or "WAREHOUSE ESHTAOL",
        "name": clean_cell(constraints.get("Warehouse Name", "WAREHOUSE ESHTAOL")) or "WAREHOUSE ESHTAOL",
        "city": clean_cell(constraints.get("Warehouse Address", "Eshtaol, Israel")) or "Eshtaol, Israel",
        "lat": parse_float(constraints.get("Warehouse Latitude")),
        "lng": parse_float(constraints.get("Warehouse Longitude")),
    }

    detail_by_driver: dict[str, list[dict]] = {}
    for row in detail_rows:
        detail_by_driver.setdefault(clean_cell(row.get("Driver", "")), []).append(row)

    drivers = []
    for row in summary_rows:
        driver_name = clean_cell(row.get("Driver", ""))
        details = detail_by_driver.get(driver_name, [])
        delivered_points = delivered_by_driver.get(driver_name, [])
        lookup = delivery_point_lookup(delivered_points)
        route_points = demo_driver_route_points(delivered_points, warehouse)
        delivery_points = route_points
        stops = [point["label"] for point in route_points if point.get("label")]
        city_names = []
        for point in delivered_points:
            city = clean_cell(point.get("city", ""))
            if city and city not in city_names:
                city_names.append(city)
        segments = []
        for detail in details:
            route = split_route_path(clean_cell(detail.get("Route Path", "")))
            segment_points = demo_route_points_for_detail(detail, lookup, warehouse, True)
            segments.append(
                {
                    "step": clean_cell(detail.get("Step", "")),
                    "city": clean_cell(detail.get("City", "")),
                    "packages": safe_int(clean_cell(detail.get("Packages", ""))),
                    "segment_type": clean_cell(detail.get("Segment Type", "")),
                    "total_time_min": clean_cell(detail.get("Total Time (min)", "")),
                    "route": route,
                    "route_points": segment_points,
                }
            )
        drivers.append(
            {
                "name": driver_name,
                "addresses": safe_int(clean_cell(row.get("Addresses Delivered", "")), len(stops)),
                "clusters": safe_int(clean_cell(row.get("Clusters Visited", ""))),
                "total_minutes": float(pd.to_numeric(row.get("Total Time (min)", 0), errors="coerce") or 0),
                "total_hours": clean_cell(row.get("Total Time (hours)", "")),
                "shift_type": clean_cell(row.get("Shift Type", "")),
                "mode": clean_cell(row.get("Mode", "CONSTRAINT PRODUCTION")),
                "cities": city_names,
                "stops": stops,
                "route_points": route_points,
                "delivery_points": delivery_points,
                "segments": segments,
            }
        )

    delivered_count = count_excel_rows(delivered_path)
    failed_count = count_excel_rows(failed_path)
    not_in_run_count = count_excel_rows(not_in_run_path)
    total_uploaded = delivered_count + failed_count + not_in_run_count
    return {
        "summary": {
            "drivers": safe_int(clean_cell(constraints.get("Drivers", "")), len(drivers)),
            "capacity": safe_int(clean_cell(constraints.get("Max Packages Per Driver", ""))),
            "uploaded_orders": total_uploaded,
            "route_ready": delivered_count + not_in_run_count,
            "selected_orders": delivered_count,
            "deferred_orders": not_in_run_count,
            "leftover_orders": safe_int(clean_cell(constraints.get("Leftover Orders", "")), not_in_run_count),
            "delivered_addresses": delivered_count,
            "regular_shifts": sum(1 for driver in drivers if driver["shift_type"].lower() == "regular"),
            "extended_shifts": sum(1 for driver in drivers if driver["shift_type"].lower() != "regular"),
            "max_shift_minutes": clean_cell(constraints.get("Max Shift Minutes", "")),
            "ideal_shift_minutes": clean_cell(constraints.get("Ideal Shift Minutes", "")),
            "warehouse": clean_cell(constraints.get("Warehouse Name", "")),
            "source": source,
        },
        "drivers": drivers,
        "files": {
            "delivery_plan_xlsx": f"{file_prefix}/files/06_delivery_plan.xlsx",
            "delivery_plan_html": f"{file_prefix}/files/06_delivery_plan.html" if html_path.exists() else "",
            "failed_addresses": f"{file_prefix}/files/07_boss_failed_addresses.xlsx" if failed_path.exists() else "",
            "leftover_orders": f"{file_prefix}/files/07_boss_not_in_run_addresses.xlsx" if not_in_run_path.exists() else "",
            "selected_orders": f"{file_prefix}/files/07_boss_delivered_addresses.xlsx" if delivered_path.exists() else "",
            "run_folder": str(result_dir),
        },
    }


def parse_demo_result() -> dict:
    return parse_saved_result(ensure_demo_results_dir(), "/api/demo-result", "Saved real Google API run")


def parse_comparison_result() -> dict:
    return parse_saved_result(latest_comparison_results_dir(), "/api/comparison-result", "Comparison results run")


def parse_source_result(source: str) -> dict:
    if source == "demo":
        return parse_demo_result()
    return parse_comparison_result()


def haversine_km(first: dict, second: dict) -> float:
    lat1, lng1 = float(first["lat"]), float(first["lng"])
    lat2, lng2 = float(second["lat"]), float(second["lng"])
    radius_km = 6371.0
    dlat = math.radians(lat2 - lat1)
    dlng = math.radians(lng2 - lng1)
    a = (
        math.sin(dlat / 2) ** 2
        + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dlng / 2) ** 2
    )
    return radius_km * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))


def route_distance_km(points: list[dict]) -> float:
    clean_points = [
        point
        for point in points
        if point.get("lat") is not None and point.get("lng") is not None
    ]
    return sum(haversine_km(a, b) for a, b in zip(clean_points, clean_points[1:]))


def minutes_to_duration(minutes: float) -> str:
    total = max(0, int(round(minutes)))
    hours = total // 60
    mins = total % 60
    if hours and mins:
        return f"{hours}h {mins}m"
    if hours:
        return f"{hours}h"
    return f"{mins}m"


def display_driver(driver: dict, index: int) -> dict:
    route_points = driver.get("route_points") or driver.get("delivery_points") or []
    stops = route_points or [{"label": stop} for stop in driver.get("stops", []) if stop]
    source_name = clean_cell(driver.get("name", "")) or f"Driver {index + 1}"
    display_name = DISPLAY_DRIVER_NAMES[index] if index < len(DISPLAY_DRIVER_NAMES) else source_name
    addresses = int(driver.get("addresses") or max(0, len(stops) - 1))
    distance = route_distance_km(route_points)
    return {
        **driver,
        "name": display_name,
        "source_name": source_name,
        "driver_index": index,
        "addresses": addresses,
        "duration_label": minutes_to_duration(float(driver.get("total_minutes") or 0)),
        "distance_km": round(distance, 1),
        "stops": stops,
        "route_points": route_points,
    }


def mobile_payload(source: str, requested_driver: str) -> dict:
    result = parse_source_result(source)
    drivers = result.get("drivers") or []
    if not drivers:
        raise RuntimeError("No driver routes were found in the selected Routecraft result.")

    selected_index = 0
    normalized_request = clean_cell(requested_driver)
    if normalized_request:
        for index, driver in enumerate(drivers):
            display_name = DISPLAY_DRIVER_NAMES[index] if index < len(DISPLAY_DRIVER_NAMES) else clean_cell(driver.get("name", ""))
            if normalized_request in {display_name, clean_cell(driver.get("name", ""))}:
                selected_index = index
                break

    driver = display_driver(drivers[selected_index], selected_index)
    summary = result.get("summary") or {}
    capacity = int(summary.get("capacity") or max(driver.get("addresses") or 0, 1))
    return {
        "summary": summary,
        "driver": driver,
        "capacity": capacity,
        "source": source,
    }


def file_url(job: JobState, path: Path) -> str:
    base = job.run_dir
    if base is None:
        return ""
    try:
        rel = path.resolve().relative_to(base.resolve())
    except ValueError:
        return ""
    return f"/api/jobs/{job.id}/files/{rel.as_posix()}"


def run_job(job: JobState, input_path: Path) -> None:
    with jobs_lock:
        job.status = "running"
        job.updated_at = time.time()
    constraints_path = create_constraints_file(job)
    cmd = [
        sys.executable,
        str(PIPELINE_SCRIPT),
        "--input",
        str(input_path),
        "--mock-google",
        "--constraints-file",
        str(constraints_path),
        "--runs-dir",
        str(WEB_RUNS_DIR),
    ]
    append_log(job, "Starting mock Routecraft pipeline...")
    append_log(job, " ".join(cmd))
    process = subprocess.Popen(
        cmd,
        cwd=ROOT,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        bufsize=1,
    )
    assert process.stdout is not None
    for line in process.stdout:
        append_log(job, line)
    returncode = process.wait()
    with jobs_lock:
        job.returncode = returncode
    if returncode != 0:
        with jobs_lock:
            job.status = "failed"
            job.error = f"Pipeline exited with code {returncode}."
            job.updated_at = time.time()
        return

    run_dir = extract_path_from_log(job.log, "Flow completed. Run folder")
    final_report = extract_path_from_log(job.log, "Final report")
    if run_dir is None:
        candidates = sorted(WEB_RUNS_DIR.iterdir(), key=lambda p: p.stat().st_mtime, reverse=True)
        run_dir = candidates[0] if candidates else None
    with jobs_lock:
        job.run_dir = run_dir
        job.final_report = final_report
    try:
        result = parse_pipeline_result(job)
        with jobs_lock:
            job.result = result
            job.status = "completed"
            job.updated_at = time.time()
    except Exception as exc:
        with jobs_lock:
            job.status = "failed"
            job.error = str(exc)
            job.updated_at = time.time()


class RoutecraftHandler(BaseHTTPRequestHandler):
    server_version = "RoutecraftWeb/1.0"

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        path = parsed.path
        if path == "/":
            self.serve_file(WEB_DIR / "index.html")
            return
        if path in {"/mobile", "/mobile/"}:
            self.serve_file(MOBILE_DIR / "index.html")
            return
        if path.startswith("/mobile/static/"):
            self.serve_file(MOBILE_STATIC_DIR / unquote(path.removeprefix("/mobile/static/")))
            return
        if path.startswith("/static/"):
            rel_path = unquote(path.removeprefix("/static/"))
            desktop_asset = STATIC_DIR / rel_path
            mobile_asset = MOBILE_STATIC_DIR / rel_path
            self.serve_file(desktop_asset if desktop_asset.exists() else mobile_asset)
            return
        if path == "/api/driver-route":
            params = parse_qs(parsed.query)
            source = clean_cell((params.get("source") or ["comparison"])[0]) or "comparison"
            driver = clean_cell((params.get("driver") or ["דביר לוי"])[0])
            try:
                json_response(self, mobile_payload(source, driver))
            except Exception as exc:
                json_response(self, {"error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)
            return
        if path == "/api/demo-result":
            try:
                json_response(
                    self,
                    {
                        "id": "saved-real-google-results",
                        "status": "completed",
                        "input_name": "Saved real Google API run",
                        "result": parse_demo_result(),
                    },
                )
            except Exception as exc:
                json_response(self, {"error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)
            return
        demo_file_match = re.match(r"^/api/demo-result/files/(.+)$", path)
        if demo_file_match:
            self.serve_demo_file(unquote(demo_file_match.group(1)))
            return
        if path == "/api/comparison-result":
            try:
                json_response(
                    self,
                    {
                        "id": "comparison-results",
                        "status": "completed",
                        "input_name": "Latest comparison run",
                        "result": parse_comparison_result(),
                    },
                )
            except Exception as exc:
                json_response(self, {"error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)
            return
        comparison_file_match = re.match(r"^/api/comparison-result/files/(.+)$", path)
        if comparison_file_match:
            self.serve_comparison_file(unquote(comparison_file_match.group(1)))
            return
        if re.fullmatch(r"/api/jobs/[^/]+", path):
            job_id = path.rsplit("/", 1)[1]
            with jobs_lock:
                job = jobs.get(job_id)
                payload = job.public() if job else None
            if payload is None:
                json_response(self, {"error": "Job not found"}, HTTPStatus.NOT_FOUND)
            else:
                json_response(self, payload)
            return
        file_match = re.match(r"^/api/jobs/([^/]+)/files/(.+)$", path)
        if file_match:
            self.serve_job_file(file_match.group(1), unquote(file_match.group(2)))
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path != "/api/jobs":
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        try:
            content_length = int(self.headers.get("Content-Length", "0"))
            body = self.rfile.read(content_length)
            fields, files = parse_multipart(body, self.headers.get("Content-Type", ""))
            drivers = safe_int(fields.get("drivers", ""), 0)
            capacity = safe_int(fields.get("capacity", ""), 0)
            upload = files.get("orders")
            if drivers <= 0 or capacity <= 0:
                raise ValueError("Drivers and vehicle capacity must be positive numbers.")
            if not upload:
                raise ValueError("Missing Excel upload.")
            filename = upload["filename"]
            if not filename.lower().endswith((".xlsx", ".xls")):
                raise ValueError("Upload must be an Excel file.")

            job_id = uuid.uuid4().hex[:12]
            job_dir = JOBS_DIR / job_id
            job_dir.mkdir(parents=True, exist_ok=False)
            input_path = job_dir / filename
            input_path.write_bytes(upload["content"])
            job = JobState(
                id=job_id,
                input_name=filename,
                drivers=drivers,
                capacity=capacity,
                job_dir=job_dir,
            )
            with jobs_lock:
                jobs[job_id] = job
            thread = threading.Thread(target=run_job, args=(job, input_path), daemon=True)
            thread.start()
            json_response(self, {"id": job_id, "status": job.status}, HTTPStatus.ACCEPTED)
        except Exception as exc:
            json_response(self, {"error": str(exc)}, HTTPStatus.BAD_REQUEST)

    def serve_file(self, path: Path) -> None:
        resolved = path.resolve()
        try:
            if not resolved.exists() or not resolved.is_file():
                self.send_error(HTTPStatus.NOT_FOUND)
                return
            body = read_bytes(resolved)
            content_type = mimetypes.guess_type(resolved.name)[0] or "application/octet-stream"
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        except OSError:
            self.send_error(HTTPStatus.INTERNAL_SERVER_ERROR)

    def serve_job_file(self, job_id: str, rel_path: str) -> None:
        with jobs_lock:
            job = jobs.get(job_id)
        if not job or not job.run_dir:
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        base = job.run_dir.resolve()
        target = (base / rel_path).resolve()
        try:
            target.relative_to(base)
        except ValueError:
            self.send_error(HTTPStatus.FORBIDDEN)
            return
        self.serve_file(target)

    def serve_demo_file(self, rel_path: str) -> None:
        try:
            base = ensure_demo_results_dir().resolve()
        except Exception:
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        target = (base / rel_path).resolve()
        try:
            target.relative_to(base)
        except ValueError:
            self.send_error(HTTPStatus.FORBIDDEN)
            return
        self.serve_file(target)

    def serve_comparison_file(self, rel_path: str) -> None:
        try:
            base = latest_comparison_results_dir().resolve()
        except Exception:
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        target = (base / rel_path).resolve()
        try:
            target.relative_to(base)
        except ValueError:
            self.send_error(HTTPStatus.FORBIDDEN)
            return
        self.serve_file(target)

    def log_message(self, format: str, *args) -> None:
        return


def main() -> int:
    host = os.getenv("ROUTECRAFT_HOST", "0.0.0.0")
    port = int(os.getenv("PORT", os.getenv("ROUTECRAFT_PORT", "8080")))
    server = ThreadingHTTPServer((host, port), RoutecraftHandler)
    try:
        print(f"Routecraft web app running at http://{host}:{port}", flush=True)
    except (AttributeError, OSError):
        pass
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        try:
            print("\nStopping Routecraft web app.")
        except (AttributeError, OSError):
            pass
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
