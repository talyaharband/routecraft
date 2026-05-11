# Routecraft OR-Tools Pipeline

This pipeline belongs to `ortools.py`.

```text
Raw/order address Excel
  -> Stages 1-4 mostly matching flow.py
  -> Stage 5 skipped for legacy matrices
  -> Stage 6 builds one global OR-Tools time matrix
  -> Stage 6 solves constrained multi-driver routing
```

The OR-Tools defaults are read from `constraints_parameters.json`.

Primary outputs:

- `04_clustered_delivery_groups.xlsx`
- `06_ortools_time_matrix.xlsx`
- `06_delivery_plan.xlsx`
- `06_delivery_plan.html`
- `06_delivery_plan.log`

