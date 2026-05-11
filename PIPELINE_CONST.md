# Routecraft Constraints Pipeline

This pipeline belongs to `const.py`.

```text
Raw/order address Excel
  -> Stage 1 clean addresses while preserving order due dates
  -> Stage 2 geocode and assign delivery priority
  -> Stage 2 select the loadable orders for this run
  -> Stage 3 cluster nearby selected orders
  -> Stage 4 create k-means delivery groups and maps
  -> Stage 5 build per-group distance matrices
  -> Stage 6 plan each driver sequentially with reload segments
```

The constraints defaults are read from `constraints_parameters_const.json`.

Important outputs:

- `01a_failed_addresses.xlsx`
- `01b_addresses_for_geocoding.xlsx`
- `01c_good_orders_original_format.xlsx`
- `02_geocoded_addresses.xlsx`
- `02b_selected_orders_for_run.xlsx`
- `02c_deferred_orders_next_run.xlsx`
- `03_nearby_address_clusters.xlsx`
- `04_clustered_delivery_groups.xlsx`
- `04_cluster_map_<city>.png`
- `05_distance_matrix_group-<group>.xlsx`
- `06_delivery_plan.xlsx`
- `06_delivery_plan.html`
- `06_delivery_plan.log`
- `06_leftover_orders_next_run.xlsx`

