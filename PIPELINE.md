# Routecraft Pipeline Index

Routecraft now has three main pipeline entry points. Each has its own matching pipeline notes:

- Original flow: `flow.py`, [PIPELINE_FLOW.md](PIPELINE_FLOW.md), [pipeline_flow.html](pipeline_flow.html)
- OR-Tools flow: `ortools.py`, [PIPELINE_ORTOOLS.md](PIPELINE_ORTOOLS.md), [pipeline_ortools.html](pipeline_ortools.html)
- Constraints flow: `const.py`, [PIPELINE_CONST.md](PIPELINE_CONST.md), [pipeline_const.html](pipeline_const.html)

The old clickable visual remains at [pipeline.html](pipeline.html) for the original `flow.py` view.

## Flow At A Glance

```text
Raw order Excel
  -> flow.py stage 1 using data.cleanup.py
  -> flow.py stages 2-4
  -> flow.py stage 5 using distance.matrix.py
  -> flow.py stage 6 using deliveries.py
```

## Stage Details

| Stage | Code Used | Input | Output |
| --- | --- | --- | --- |
| 1. Clean raw orders | `flow.py`, `data.cleanup.py` | Raw order Excel with order/client/site/city/address/comment columns. | `01a_failed_addresses.xlsx`, `01b_addresses_for_geocoding.xlsx`, `01c_good_orders_original_format.xlsx` |
| 2. Geocode | `flow.py` | `City`, `Street_Name`, `House_Number` rows. | `02_geocoded_addresses.xlsx`, plus `02a_failed_geocoding_addresses.xlsx` when Google cannot precisely match rows. |
| 3. Cluster nearby addresses | `flow.py` | Geocoded rows with `LAT` and `LNG`. | `03_nearby_address_clusters.xlsx` |
| 4. Group stops | `flow.py` | Clustered rows. | `04_clustered_delivery_groups.xlsx` |
| 5. Build matrices | `flow.py`, `distance.matrix.py` | Grouped rows with `cluster_group`. | `05_distance_matrix_group-<group>.xlsx` files. |
| 6. Delivery plan | `flow.py`, `deliveries.py` | Grouped rows, matrix files, and optional stage-1 original order metadata. | `06_delivery_plan.xlsx`, `06_delivery_plan.html`, `06_delivery_plan.log` |

## Running

Mock Google for safe testing:

```powershell
.\.venv\Scripts\python.exe flow.py --start-stage 2 --input addresses\good_addresses.xlsx --mock-google
```

Full run from raw orders:

```powershell
.\.venv\Scripts\python.exe flow.py --start-stage 1 --input addresses\base_addresses.xlsx
```

Use `--mock-google` when testing without Google API quota. It mocks stage 2 geocoding, stage 5 distance matrices, and stage 6 current-location travel times.
