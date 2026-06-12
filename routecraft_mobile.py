from __future__ import annotations

import json
import mimetypes
import os
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from math import atan2, cos, radians, sin, sqrt
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse

from routecraft_web import clean_cell, parse_comparison_result, parse_demo_result


ROOT = Path(__file__).resolve().parent
MOBILE_DIR = ROOT / "mobile"
MOBILE_STATIC_DIR = MOBILE_DIR / "static"
DISPLAY_DRIVER_NAMES = [
    "דביר לוי",
    "שמעיה סבן",
    "נהוראי מלצר",
    "יוסי אביטן",
    "איתיאל סופר",
    "יעל רבינוביץ",
]


def json_response(handler: BaseHTTPRequestHandler, payload: dict, status: int = 200) -> None:
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.end_headers()
    handler.wfile.write(body)


def read_bytes(path: Path) -> bytes:
    return path.read_bytes()


def parse_source_result(source: str) -> dict:
    if source == "demo":
        return parse_demo_result()
    return parse_comparison_result()


def haversine_km(first: dict, second: dict) -> float:
    lat1, lng1 = float(first["lat"]), float(first["lng"])
    lat2, lng2 = float(second["lat"]), float(second["lng"])
    radius_km = 6371.0
    dlat = radians(lat2 - lat1)
    dlng = radians(lng2 - lng1)
    a = sin(dlat / 2) ** 2 + cos(radians(lat1)) * cos(radians(lat2)) * sin(dlng / 2) ** 2
    return radius_km * 2 * atan2(sqrt(a), sqrt(1 - a))


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


class RoutecraftMobileHandler(BaseHTTPRequestHandler):
    server_version = "RoutecraftMobile/1.0"

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        path = parsed.path
        if path in {"/", "/mobile", "/mobile/"}:
            self.serve_file(MOBILE_DIR / "index.html")
            return
        if path.startswith("/static/"):
            self.serve_file(MOBILE_STATIC_DIR / unquote(path.removeprefix("/static/")))
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
        self.send_error(HTTPStatus.NOT_FOUND)

    def serve_file(self, path: Path) -> None:
        resolved = path.resolve()
        try:
            resolved.relative_to(MOBILE_DIR.resolve())
        except ValueError:
            self.send_error(HTTPStatus.FORBIDDEN)
            return
        if not resolved.exists() or not resolved.is_file():
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        body = read_bytes(resolved)
        content_type = mimetypes.guess_type(resolved.name)[0] or "application/octet-stream"
        if resolved.suffix == ".js":
            content_type = "text/javascript; charset=utf-8"
        elif resolved.suffix in {".html", ".css"}:
            content_type = f"{content_type}; charset=utf-8"
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, format: str, *args) -> None:
        return


def main() -> int:
    host = os.getenv("ROUTECRAFT_MOBILE_HOST", "127.0.0.1")
    port = int(os.getenv("ROUTECRAFT_MOBILE_PORT", "8081"))
    server = ThreadingHTTPServer((host, port), RoutecraftMobileHandler)
    print(f"Routecraft mobile app running at http://{host}:{port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping Routecraft mobile app.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
