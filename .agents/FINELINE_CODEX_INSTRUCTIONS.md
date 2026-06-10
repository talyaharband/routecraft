# FINELINE Route Planner - Codex Instructions

This project connects a FINELINE route planning UI designed in Google Stitch to an existing Python routing algorithm.

## Main Goal

Build an interactive web application that allows a warehouse shift manager to:

1. Enter daily route planning parameters.
2. Upload or select an orders Excel file.
3. Run the existing Python routing algorithm.
4. View assigned driver routes, failed addresses, and unassigned deliveries in the FINELINE dashboard UI.

## Current Assets

* `docs/PROJECT_BRIEF.md` contains the Stitch project/product brief.
* `frontend-stitch-code/` contains static HTML exports from Google Stitch.
* `frontend-stitch-export/` contains the original ZIP export files from Stitch, including Markdown, PNG, and HTML.
* `backend/CONST_EXPERIMENTAL.py` contains the existing Python routing algorithm.

## Important Rules

* Preserve the visual design exported from Google Stitch as closely as possible.
* Do not redesign the UI unless explicitly asked.
* Do not change the routing logic inside `CONST_EXPERIMENTAL.py` unless absolutely necessary.
* Prefer wrapping the existing algorithm with a clean function/API instead of rewriting it.
* The manager dashboard interface must remain in English.
* Hebrew delivery addresses must display correctly right-to-left inside the English interface.
* Use the files in `docs/` as product and UI specifications.
* Use the Stitch HTML files as the visual reference for the frontend.

## Recommended Architecture

* Frontend: React app.
* Backend: FastAPI.
* The frontend should call backend API endpoints and render dynamic data.
* The backend should call the existing Python routing algorithm.

## First Working Goal

Build the basic flow only:

Daily Setup -> Processing -> Route Results

The flow should work as follows:

1. Manager enters number of drivers and vehicle capacity.
2. Manager uploads or selects an orders Excel file.
3. Manager clicks "Generate Daily Routes".
4. Frontend sends the file and parameters to the backend.
5. Backend runs the existing routing algorithm.
6. Backend returns structured JSON.
7. Frontend displays the Route Results Dashboard using the real returned data.

## Expected Backend JSON Shape

The backend response should include:

* `run_summary`

  * `uploaded_deliveries`
  * `assigned_deliveries`
  * `assigned_drivers`
  * `regular_shifts`
  * `extended_shifts`

* `drivers`

  * `driver_id`
  * `driver_name`
  * `stops_count`
  * `route_duration`
  * `route_distance_km`
  * `shift_type`
  * `capacity_used`
  * `capacity_total`
  * `stops`

* `failed_addresses`

* `unassigned_deliveries`

## Do Not Implement Yet

Do not implement these until the basic flow works:

* Live delivery monitoring
* Proof of Delivery
* Driver mobile app functionality
* User authentication
* Database
* Deployment
* Real-time driver tracking
