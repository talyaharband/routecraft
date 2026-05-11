# Routecraft Flow Pipeline

This is the original automated pipeline for `flow.py`.

```text
Raw order Excel
  -> Stage 1 clean addresses
  -> Stage 2 geocode
  -> Stage 3 cluster nearby addresses
  -> Stage 4 create delivery groups
  -> Stage 5 build per-group distance matrices
  -> Stage 6 plan deliveries with deliveries.py
```

Outputs are written into a timestamped `runs/<timestamp>` folder, including cleaned data, geocoding caches, clusters, maps, matrices, and final delivery plan files.

