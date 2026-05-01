# Routecraft Pipeline

Routecraft is a set of small Python scripts that turn messy delivery addresses into route-ready driving-time matrices, then solve each matrix as a route. The scripts are run one at a time and pass Excel files between stages.

For a clickable visual version of this flow, open [pipeline.html](pipeline.html) in a browser.

## Flow At A Glance

```text
Raw order Excel
  -> data.cleanup.py
  -> address Excel with City, Street_Name, House_Number
  -> geocoding.py
  -> cluster.py
  -> k-mean.py
  -> distance.matrix.py
  -> TSP.py
```

`closesToEshtaol.py` is an optional helper. It does not create a file; it tells you which address and `cluster_group` are closest to the Eshtaol warehouse coordinates.

## Important Handoff Note

The current `data.cleanup.py` output is not directly compatible with `geocoding.py`.

`data.cleanup.py` preserves the raw order file and adds cleaned-address columns such as `ship_to_street_final`. `geocoding.py`, however, reads the first three columns of its input file as:

1. `City`
2. `Street_Name`
3. `House_Number`

So after cleanup, there is currently a manual or external preparation step that must produce an Excel file in that three-column shape. The checked-in sample file `addresses/good_addresses.xlsx` is an example of the format that `geocoding.py` expects.

## Stage Details

| Step | Script | Input | Output | Notes |
| --- | --- | --- | --- | --- |
| 1 | `data.cleanup.py` | Raw order Excel selected with a file chooser. It uses `ship_to_street1`, `ship_to_street2`, and `site_name` when present. | User-selected Excel file that keeps the input columns and adds `ship_to_street_final`, `is_real_address`, and `coordinates`. | Prompts for an OpenAI key and a Google Maps key. Uses OpenAI to parse/normalize Israeli addresses, then validates precise addresses with Google Geocoding. |
| 2 | `geocoding.py` | Excel selected with a file chooser. The first three columns must be city, street, and house number. | User-selected Excel file with `City`, `Street_Name`, `House_Number`, `LAT`, and `LNG`. | Reads the Google key from `.env` as `GOOGLE_MAPS_API_KEY` or `GEOCODING_API_KEY`. It does not preserve all original columns; it creates a fresh five-column file. |
| 3 | `cluster.py` | Geocoded Excel with `Street_Name`, `House_Number`, `LAT`, and `LNG`. | User-selected Excel file with one representative row per cluster, plus `cluster_id`, `total_orders_in_cluster`, and `detailed_addresses`. | Clusters nearby addresses on the same street using a 150-meter haversine threshold. This collapses multiple orders into a representative row, rather than keeping every original row. |
| 4 | `k-mean.py` | Clustered Excel. In the current file shape, columns D and E are `LAT` and `LNG`. | User-selected Excel file with a new `cluster_group` column. | If there are 30 or fewer rows, all rows get group `0`. Otherwise it uses constrained k-means with a maximum of 30 rows per group and shows a scatter plot before saving. |
| 5 | `distance.matrix.py` | Grouped Excel from `k-mean.py`. The 9th column, column I, must be `cluster_group`. | One Excel matrix per group, saved next to the input file as `matrix_<group>_<timestamp>.xlsx`. | Uses a hard-coded `API_KEY` value in the script. It builds labels from column B and column A, currently `Street_Name, City`, and calculates driving minutes with Google Distance Matrix. |
| 6 | `TSP.py` | One matrix Excel file from `distance.matrix.py`. | Console output only: route options, best route, total minutes, and address labels. | Run it once per matrix file. It treats row/column `0` as the start point and searches over possible ending points. |

## Optional Helper: Closest To Eshtaol

`closesToEshtaol.py` reads an Excel file with `LAT`, `LNG`, `Street_Name`, `House_Number`, and `cluster_group`. It compares every row to the hard-coded warehouse coordinates:

```text
31.77927525, 35.0105885
```

It prints:

- the closest address to the warehouse
- the `cluster_group` that address belongs to

It does not write an output file. In practice, run it after `k-mean.py` when you want to identify which group contains the warehouse-nearest stop.

## Example File Chain

The repository includes sample outputs that match the main pipeline:

```text
addresses/good_addresses.xlsx
  -> addresses/good_addresses_with_geocoding.xlsx
  -> addresses/outcluster.xlsx
  -> addresses/outk-mean.xlsx
  -> addresses/matrix_*.xlsx
```

`addresses/datacleanup_out.xlsx` is an example of the cleanup output, but it still needs to be reshaped into the `City`, `Street_Name`, `House_Number` format before `geocoding.py`.

## Running The Scripts

From an activated virtual environment:

```powershell
python .\data.cleanup.py
python .\geocoding.py
python .\cluster.py
python .\k-mean.py
python .\distance.matrix.py
python .\closesToEshtaol.py
python .\TSP.py
```

The scripts use desktop file pickers, so run them from a local Windows session where Tkinter windows can open.

## API Keys

- `data.cleanup.py` asks for the OpenAI key and Google Maps key in popup windows.
- `geocoding.py` reads `GOOGLE_MAPS_API_KEY` or `GEOCODING_API_KEY` from `.env`.
- `distance.matrix.py` currently uses the hard-coded `API_KEY` variable inside the file.
