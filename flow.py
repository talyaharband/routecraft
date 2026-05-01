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
from concurrent.futures import ThreadPoolExecutor, as_completed
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


def get_openai_key(args: argparse.Namespace) -> str:
    key = args.openai_api_key or os.getenv("OPENAI_API_KEY")
    if not key:
        raise RuntimeError(
            "Missing OpenAI API key. Set OPENAI_API_KEY in .env or pass --openai-api-key."
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
    from openai import OpenAI

    google_key = get_google_key(args)
    openai_key = get_openai_key(args)
    client = OpenAI(api_key=openai_key)
    df = pd.read_excel(input_path).fillna("")

    for col in ["ship_to_street1", "ship_to_street2", "site_name"]:
        if col not in df.columns:
            df[col] = ""

    addresses = [""] * len(df)
    statuses = [""] * len(df)
    coords = [""] * len(df)

    with ThreadPoolExecutor(max_workers=args.max_workers) as executor:
        futures = {
            executor.submit(
                cleanup.process_address_clean,
                i,
                row.get("ship_to_street1", ""),
                row.get("ship_to_street2", ""),
                row.get("site_name", ""),
                client,
                google_key,
            ): i
            for i, row in df.iterrows()
        }
        for future in as_completed(futures):
            idx, address, status, coord = future.result()
            addresses[idx] = address
            statuses[idx] = status
            coords[idx] = coord

    df["ship_to_street_final"] = addresses
    df["is_real_address"] = statuses
    df["coordinates"] = coords

    cleanup_path = run_dir / "01_cleanup_raw_orders.xlsx"
    df.to_excel(cleanup_path, index=False)

    shaped_rows = []
    for _, row in df.iterrows():
        street, number = split_street_and_number(row.get("ship_to_street_final", ""))
        shaped_rows.append(
            {
                "City": clean_cell(row.get("site_name", "")),
                "Street_Name": street,
                "House_Number": number,
            }
        )
    shaped_df = pd.DataFrame(shaped_rows)
    shaped_path = run_dir / "01b_addresses_for_geocoding.xlsx"
    shaped_df.to_excel(shaped_path, index=False)

    return shaped_path, StageResult(
        1,
        "Clean raw order addresses",
        "completed",
        [cleanup_path, shaped_path],
        [
            "Automated the manual handoff noted in PIPELINE.md by reshaping "
            "ship_to_street_final into City, Street_Name, House_Number."
        ],
    )


def stage_2_geocode(input_path: Path, run_dir: Path, args: argparse.Namespace) -> tuple[Path, StageResult]:
    google_key = "" if args.mock_google else get_google_key(args)
    df = pd.read_excel(input_path)
    if df.shape[1] < 3:
        raise RuntimeError("Stage 2 input must have at least 3 columns: city, street, house number.")

    city_col, street_col, number_col = df.columns[:3]
    rows = []
    for _, row in df.iterrows():
        city = clean_cell(row[city_col])
        street = clean_cell(row[street_col])
        house_number = clean_cell(row[number_col])
        query = f"{street} {house_number}, {city}".strip()
        lat, lng = mock_lat_lng(query) if args.mock_google else get_lat_lng(query, google_key)
        rows.append(
            {
                "City": city,
                "Street_Name": street,
                "House_Number": house_number,
                "LAT": lat,
                "LNG": lng,
            }
        )

    output_path = run_dir / "02_geocoded_addresses.xlsx"
    pd.DataFrame(rows).to_excel(output_path, index=False)
    return output_path, StageResult(2, "Geocode addresses", "completed", [output_path])


def stage_3_cluster(input_path: Path, run_dir: Path, args: argparse.Namespace) -> tuple[Path, StageResult]:
    df = pd.read_excel(input_path)
    required = {"Street_Name", "House_Number", "LAT", "LNG"}
    missing = required.difference(df.columns)
    if missing:
        raise RuntimeError(f"Stage 3 input is missing columns: {', '.join(sorted(missing))}")

    df["House_Number"] = pd.to_numeric(df["House_Number"], errors="coerce")
    df = df.dropna(subset=["House_Number"]).reset_index(drop=True)
    if df.empty:
        raise RuntimeError("Stage 3 has no rows with numeric House_Number.")

    df["House_Number"] = df["House_Number"].astype(int)
    df = df.sort_values(by=["Street_Name", "House_Number"]).reset_index(drop=True)

    clusters: list[list[dict[str, Any]]] = []
    anchor = df.iloc[0]
    current = [anchor.to_dict()]
    for i in range(1, len(df)):
        row = df.iloc[i]
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
    if df.shape[1] < 5:
        raise RuntimeError("Stage 4 input must include LAT/LNG in columns D and E.")

    coords = df.iloc[:, [3, 4]].values
    num_addresses = len(df)
    if num_addresses <= args.max_group_size:
        df["cluster_group"] = 0
        note = f"{num_addresses} rows fit in one group."
    else:
        from k_means_constrained import KMeansConstrained

        n_clusters = int(np.ceil(num_addresses / args.max_group_size))
        clf = KMeansConstrained(
            n_clusters=n_clusters,
            size_min=1,
            size_max=args.max_group_size,
            random_state=42,
        )
        df["cluster_group"] = clf.fit_predict(coords)
        note = f"Created {n_clusters} constrained groups."

    output_path = run_dir / "04_clustered_delivery_groups.xlsx"
    df.to_excel(output_path, index=False)
    return output_path, StageResult(4, "Create delivery groups", "completed", [output_path], [note])


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
    if df.shape[1] < 9:
        raise RuntimeError("Stage 5 input must include column I with the group number.")

    group_col = df.columns[8]
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
    parser.add_argument("--openai-api-key")
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
        if args.start_stage <= 5:
            current, result = stage_5_distance_matrices(Path(current), run_dir, args)
            stage_results.append(result)
        if args.start_stage <= 6:
            matrix_paths = current if isinstance(current, list) else [Path(current)]
            current, result = stage_6_tsp_html(matrix_paths, run_dir)
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
