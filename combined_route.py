import pandas as pd
import numpy as np
import math
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from bidi.algorithm import get_display
import time

# Warehouse coordinates
WAREHOUSE_LAT = 31.77927525
WAREHOUSE_LNG = 35.0105885
DELIVERY_TIME_MINUTES = 4


def calculate_haversine(lat1, lon1, lat2, lon2):
    """Calculate air distance in meters between two coordinates"""
    if pd.isna(lat1) or pd.isna(lon1) or pd.isna(lat2) or pd.isna(lon2):
        return float('inf')
    R = 6371000
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlambda = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2) ** 2 + math.cos(phi1) * math.cos(phi2) * math.sin(dlambda / 2) ** 2
    return 2 * R * math.atan2(math.sqrt(a), math.sqrt(1 - a))


def find_closest_unvisited_address(df, current_lat, current_lng, visited_clusters):
    """Find closest unvisited address to current location"""
    unvisited_df = df[~df['cluster_group'].isin(visited_clusters)].copy()
    
    if unvisited_df.empty:
        return None
    
    unvisited_df['air_distance'] = unvisited_df.apply(
        lambda row: calculate_haversine(current_lat, current_lng, row['LAT'], row['LNG']),
        axis=1
    )
    
    closest_idx = unvisited_df['air_distance'].idxmin()
    return df.loc[closest_idx]


def load_distance_matrix(cluster_group, matrix_folder):
    """Load distance matrix for a specific cluster"""
    folder_path = Path(matrix_folder)
    
    # Search recursively in the folder, and intelligently parse the number mathematically
    all_matrices = list(folder_path.rglob("matrix_*.xlsx"))
    matrix_files = []
    
    for file in all_matrices:
        parts = file.name.split('_')
        if len(parts) >= 2:
            try:
                # Compares the number mathematically (e.g., 0 == 0.0)
                if float(parts[1]) == float(cluster_group):
                    matrix_files.append(file)
            except ValueError:
                pass
    
    if not matrix_files:
        print(f"❌ No matrix found for cluster {cluster_group}")
        return None
    
    # Sort by the actual last modified date, so it ALWAYS picks the file you manually edited
    matrix_files.sort(key=lambda f: f.stat().st_mtime)
    latest_matrix = matrix_files[-1]
    print(f"📂 Found {len(matrix_files)} matching files. Loading the most recently saved: {latest_matrix.name}")
    
    try:
        df = pd.read_excel(latest_matrix, index_col=0)
        return df
    except Exception as e:
        print(f"❌ Error loading matrix: {e}")
        return None


def add_origin_to_matrix_test(matrix_df, origin_lat, origin_lng, addresses_data):
    """Add origin location as row 0 with sequential numbers (TEST MODE)"""
    new_matrix = pd.DataFrame(index=["ORIGIN"] + list(matrix_df.index), 
                             columns=["ORIGIN"] + list(matrix_df.columns))
    
    # Copy existing matrix to lower-right
    new_matrix.iloc[1:, 1:] = matrix_df.values
    
    # Column 0: all 999 (never return to origin)
    new_matrix.iloc[:, 0] = 999
    new_matrix.loc["ORIGIN", "ORIGIN"] = 0
    
    # Row 0: SEQUENTIAL numbers (0, 1, 2, 3...) (TEST MODE)
    print("🧪 Filling row 0 with sequential times (0, 1, 2...)...")
    for i, addr in enumerate(matrix_df.columns, 1):
        # The cell (ORIGIN, ORIGIN) is already 0. This loop fills 1, 2, 3...
        new_matrix.loc["ORIGIN", addr] = i
        time.sleep(0.01)
    
    # Convert to numeric
    new_matrix = new_matrix.apply(pd.to_numeric, errors='coerce')
    
    # 🚨 Enforce square matrix (removes phantom Excel columns/rows)
    if new_matrix.shape[0] != new_matrix.shape[1]:
        min_dim = min(new_matrix.shape[0], new_matrix.shape[1])
        new_matrix = new_matrix.iloc[:min_dim, :min_dim]

    
    return new_matrix


class ClusterTSPSolver:
    """Encapsulates the Branch and Bound + 2-Opt logic natively from tsp.py"""
    def __init__(self, distance_matrix):
        self.distance_matrix = distance_matrix
        self.n = distance_matrix.shape[0]

    def route_cost(self, route):
        return sum(self.distance_matrix[route[i]][route[i + 1]] for i in range(len(route) - 1))

    def nearest_neighbor_UB(self, end_address):
        visited = set([0])
        path = [0]
        current = 0
        nodes_to_visit = [i for i in range(1, self.n) if i != end_address]
        while nodes_to_visit:
            next_node = min(nodes_to_visit, key=lambda j: self.distance_matrix[current][j])
            path.append(next_node)
            visited.add(next_node)
            nodes_to_visit.remove(next_node)
            current = next_node
        path.append(end_address)
        return path, self.route_cost(path)

    def two_opt(self, route, end_address):
        improved = True
        best_route = route
        best_cost = self.route_cost(route)
        while improved:
            improved = False
            for i in range(1, len(best_route) - 2):
                for j in range(i + 1, len(best_route) - 1):
                    new_route = best_route[:i] + best_route[i:j + 1][::-1] + best_route[j + 1:]
                    new_cost = self.route_cost(new_route)
                    if new_cost < best_cost:
                        best_route, best_cost, improved = new_route, new_cost, True
                        break
                if improved: break
        return best_route, best_cost

    def create_numeric_case_matrix(self, end_address):
        mat = self.distance_matrix.copy().astype(float)
        mat[:, 0] = np.inf
        mat[end_address, :] = np.inf
        for i in range(self.n):
            mat[i, i] = np.inf
        return mat

    def compute_LB_and_reduced_matrix(self, end_address):
        mat = self.create_numeric_case_matrix(end_address)
        lb = 0.0
        for i in range(self.n):
            row = mat[i, :]
            finite_vals = row[np.isfinite(row)]
            if finite_vals.size == 0: continue
            m = finite_vals.min()
            if m > 0:
                mat[i, :] = np.where(np.isfinite(mat[i, :]), mat[i, :] - m, mat[i, :])
                lb += m
        for j in range(self.n):
            col = mat[:, j]
            finite_vals = col[np.isfinite(col)]
            if finite_vals.size == 0: continue
            m = finite_vals.min()
            if m > 0:
                mat[:, j] = np.where(np.isfinite(mat[:, j]), mat[:, j] - m, mat[:, j])
                lb += m
        return lb, mat

    def reduce_matrix(self, mat):
        mat = mat.copy()
        added_lb = 0.0
        for i in range(mat.shape[0]):
            row = mat[i, :]
            finite_vals = row[np.isfinite(row)]
            if finite_vals.size:
                m = finite_vals.min()
                if m > 0: 
                    mat[i, :] = np.where(np.isfinite(mat[i, :]), mat[i, :] - m, mat[i, :])
                    added_lb += m
        for j in range(mat.shape[1]):
            col = mat[:, j]
            finite_vals = col[np.isfinite(col)]
            if finite_vals.size:
                m = finite_vals.min()
                if m > 0: 
                    mat[:, j] = np.where(np.isfinite(mat[:, j]), mat[:, j] - m, mat[:, j])
                    added_lb += m
        return mat, added_lb

    def choose_branch_zero(self, mat, rows, cols, depth):
        best_penalty = -1
        best_pair = (None, None)
        for r_idx in range(mat.shape[0]):
            for c_idx in range(mat.shape[1]):
                if mat[r_idx, c_idx] != 0: continue
                row = mat[r_idx, :]
                row_mask = np.isfinite(row)
                row_mask[c_idx] = False
                row_vals = row[row_mask]
                row_min = row_vals.min() if row_vals.size else 0
                col = mat[:, c_idx]
                col_mask = np.isfinite(col)
                col_mask[r_idx] = False
                col_vals = col[col_mask]
                col_min = col_vals.min() if col_vals.size else 0
                penalty = row_min + col_min
                i, j = rows[r_idx], cols[c_idx]
                if penalty > best_penalty or (penalty == best_penalty and best_pair[0] is not None and i < best_pair[0]):
                    best_penalty, best_pair = penalty, (i, j)
        return best_pair

    def build_paths_from_forced(self, forced_edges):
        succ, pred = {}, {}
        for i, j in forced_edges: 
            succ[i], pred[j] = j, i
        return succ, pred

    def block_cycle_edges(self, mat, rows, cols, forced_edges, end_node):
        succ, pred = self.build_paths_from_forced(forced_edges)
        def block_edge(u, v):
            if u in rows and v in cols: mat[rows.index(u), cols.index(v)] = np.inf
        for i, j in forced_edges: block_edge(j, i)
        starts = [node for node in succ.keys() if node not in pred]
        for start in starts:
            current, visited = start, {start}
            while current in succ:
                nxt = succ[current]
                if nxt in visited: break
                visited.add(nxt)
                current = nxt
            block_edge(current, start)
            if start == 0 and current != end_node and len(visited) < self.n - 1: 
                block_edge(current, end_node)

    def reconstruct_route_from_forced(self, forced_edges, end_node):
        if len(forced_edges) != self.n - 1: return None
        succ, pred = self.build_paths_from_forced(forced_edges)
        if 0 not in succ: return None
        route, current, visited = [0], 0, {0}
        while current in succ:
            nxt = succ[current]
            if nxt in visited: return None
            route.append(nxt)
            visited.add(nxt)
            current = nxt
        return route if route[-1] == end_node and len(route) == self.n else None

    def apply_forced_moves(self, mat, rows, cols, forced_edges, end_node, lb, depth, current_ub):
        changed = True
        while changed:
            changed = False
            for r_idx in range(len(rows)):
                if np.isfinite(mat[r_idx, :]).sum() == 1:
                    c_idx = np.where(np.isfinite(mat[r_idx, :]))[0][0]
                    i, j = rows[r_idx], cols[c_idx]
                    forced_edges.append((i, j))
                    mat = np.delete(np.delete(mat, r_idx, axis=0), c_idx, axis=1)
                    rows.pop(r_idx)
                    cols.pop(c_idx)
                    self.block_cycle_edges(mat, rows, cols, forced_edges, end_node)
                    mat, add_lb = self.reduce_matrix(mat)
                    lb += add_lb
                    if lb >= current_ub: return mat, rows, cols, forced_edges, lb, True
                    changed = True
                    break
            if not changed:
                for c_idx in range(len(cols)):
                    if np.isfinite(mat[:, c_idx]).sum() == 1:
                        r_idx = np.where(np.isfinite(mat[:, c_idx]))[0][0]
                        i, j = rows[r_idx], cols[c_idx]
                        forced_edges.append((i, j))
                        mat = np.delete(np.delete(mat, r_idx, axis=0), c_idx, axis=1)
                        rows.pop(r_idx)
                        cols.pop(c_idx)
                        self.block_cycle_edges(mat, rows, cols, forced_edges, end_node)
                        mat, add_lb = self.reduce_matrix(mat)
                        lb += add_lb
                        if lb >= current_ub: return mat, rows, cols, forced_edges, lb, True
                        changed = True
                        break
        return mat, rows, cols, forced_edges, lb, False

    def explore_node(self, node, current_ub, best_route, depth):
        lb, mat, rows, cols = node["lb"], node["matrix"], node["rows"], node["cols"]
        forced_edges, end_node = node["forced_edges"], node["end"]
        if lb >= current_ub: return current_ub, best_route
        mat, rows, cols, forced_edges, lb, pruned = self.apply_forced_moves(
            mat, rows, cols, forced_edges, end_node, lb, depth, current_ub)
        if pruned: return current_ub, best_route
        route = self.reconstruct_route_from_forced(forced_edges, end_node)
        if route is not None:
            cost = self.route_cost(route)
            if cost < current_ub: return cost, route
            return current_ub, best_route
        if mat.size == 0: return current_ub, best_route
        i, j = self.choose_branch_zero(mat, rows, cols, depth)
        if i is None: return current_ub, best_route

        r_idx, c_idx = rows.index(i), cols.index(j)
        
        # WITH edge
        mat_with = np.delete(np.delete(mat.copy(), r_idx, axis=0), c_idx, axis=1)
        rows_w = rows[:r_idx] + rows[r_idx + 1:]
        cols_w = cols[:c_idx] + cols[c_idx + 1:]
        forced_w = forced_edges + [(i, j)]
        self.block_cycle_edges(mat_with, rows_w, cols_w, forced_w, end_node)
        mat_w_red, add_w = self.reduce_matrix(mat_with)

        # WITHOUT edge
        mat_wo = mat.copy()
        mat_wo[r_idx, c_idx] = np.inf
        mat_wo_red, add_wo = self.reduce_matrix(mat_wo)

        children = []
        if lb + add_w < current_ub:
            children.append({
                "lb": lb + add_w, "matrix": mat_w_red, "rows": rows_w, 
                "cols": cols_w, "forced_edges": forced_w, "end": end_node
            })
        if lb + add_wo < current_ub:
            children.append({
                "lb": lb + add_wo, "matrix": mat_wo_red, "rows": rows.copy(), 
                "cols": cols.copy(), "forced_edges": forced_edges.copy(), "end": end_node
            })

        children.sort(key=lambda x: x["lb"])
        for child in children:
            current_ub, best_route = self.explore_node(child, current_ub, best_route, depth + 1)
        return current_ub, best_route

    def solve(self):
        cases = []
        for end in range(1, self.n):
            lb, reduced_numeric = self.compute_LB_and_reduced_matrix(end)
            nn_path, nn_cost = self.nearest_neighbor_UB(end)
            opt2_path, opt2_cost = self.two_opt(nn_path, end)
            cases.append({
                "end": end, "initial_lb": lb, "initial_matrix": reduced_numeric,
                "ub": opt2_cost, "ub_route": opt2_path
            })

        sorted_cases = sorted(cases, key=lambda c: c["ub"])
        global_ub = sorted_cases[0]["ub"]
        global_best_route = sorted_cases[0]["ub_route"]
        results = []

        for case in sorted_cases:
            end, initial_lb = case["end"], case["initial_lb"]
            if initial_lb >= global_ub: continue
            root_node = {
                "lb": initial_lb, "matrix": case["initial_matrix"].copy(),
                "rows": list(range(self.n)), "cols": list(range(self.n)),
                "forced_edges": [], "end": end
            }
            best_ub_for_end, best_route_for_end = self.explore_node(root_node, global_ub, case["ub_route"], 0)
            if best_route_for_end is not None and best_ub_for_end < global_ub:
                global_ub, global_best_route = best_ub_for_end, best_route_for_end
            if best_route_for_end is not None:
                results.append({"end": end, "best_cost": best_ub_for_end, "best_route": best_route_for_end})

        if results:
            best_overall = min(results, key=lambda r: r["best_cost"])
            return best_overall["best_route"], best_overall["best_cost"]
        return global_best_route, global_ub


def run_tsp_on_matrix(matrix_df):
    """Run advanced TSP algorithm (B&B + 2-Opt) inherited from tsp.py"""
    print("\n🔄 Running advanced TSP optimization (Branch & Bound + 2-Opt)...")
    distance_matrix = matrix_df.to_numpy()
    
    solver = ClusterTSPSolver(distance_matrix)
    best_route, best_cost = solver.solve()
    
    return best_route, best_cost


def run_multi_cluster_tsp_test():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    # 1. Select k-means output Excel
    print("Step 1: Select k-means output file")
    kmeans_path = filedialog.askopenfilename(
        title="Select k-means output Excel (with cluster_group)",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not kmeans_path:
        print("❌ Cancelled")
        return

    # 2. Select distance matrices folder
    print("\nStep 2: Select folder with distance matrices")
    matrix_folder = filedialog.askdirectory(title="Select folder with matrix_*.xlsx files")
    if not matrix_folder:
        print("❌ Cancelled")
        return

    # 3. Load data
    print("\n" + "="*60)
    print("📂 LOADING DATA")
    print("="*60)
    
    df_addresses = pd.read_excel(kmeans_path)
    print(f"✓ Loaded {len(df_addresses)} addresses")
    print(f"✓ Clusters found: {sorted(df_addresses['cluster_group'].unique())}")

    print("\n🔍 Checking for detailed addresses (aftercluster)...")
    aftercluster_path = None
    for f in Path(matrix_folder).glob("*.xlsx"):
        if "aftercluster" in f.name.lower() or "after_cluster" in f.name.lower():
            aftercluster_path = f
            break
            
    if aftercluster_path:
        print(f"✓ Loading detailed addresses from: {aftercluster_path.name}")
        df_details = pd.read_excel(aftercluster_path)
        details_map = {}
        for _, row in df_details.iterrows():
            st = str(row.get('Street_Name', '')).strip()
            hn = str(row.get('House_Number', '')).replace('.0', '').strip()
            det = str(row.get('detailed_addresses', f"{st} {hn}"))
            details_map[f"{st} {hn}"] = det
            
        def get_detail(r):
            st = str(r.get('Street_Name', '')).strip()
            hn = str(r.get('House_Number', '')).replace('.0', '').strip()
            return details_map.get(f"{st} {hn}", f"{st} {hn}")
            
        df_addresses['detailed_addresses'] = df_addresses.apply(get_detail, axis=1)
    else:
        if 'detailed_addresses' not in df_addresses.columns:
            print("⚠️ Could not find aftercluster file or detailed_addresses column. Using standard addresses.")
            df_addresses['detailed_addresses'] = df_addresses.apply(
                lambda r: f"{str(r.get('Street_Name', '')).strip()} {str(r.get('House_Number', '')).replace('.0', '').strip()}", 
                axis=1
            )
        else:
            print("✓ detailed_addresses column already exists in the loaded file.")

    # Create address lookup
    address_lookup = {}
    for idx, row in df_addresses.iterrows():
        addr_name = f"{row['Street_Name']} {row['House_Number']}"
        address_lookup[addr_name] = {
            'LAT': row['LAT'],
            'LNG': row['LNG'],
            'cluster_group': row['cluster_group'],
            'street': row['Street_Name'],
            'house': row['House_Number']
        }

    # 4. Start journey
    print("\n" + "="*60)
    print("🚀 STARTING MULTI-CLUSTER TSP JOURNEY (TEST MODE)")
    print("="*60)

    current_lat = WAREHOUSE_LAT
    current_lng = WAREHOUSE_LNG
    current_location_name = "WAREHOUSE ESHTAOL"
    full_journey_path = [get_display(current_location_name)]
    
    visited_clusters = set()
    journey_results = []
    total_journey_time = 0

    cluster_count = 0
    total_clusters = len(df_addresses['cluster_group'].unique())

    # 5. Main loop
    while len(visited_clusters) < total_clusters:
        cluster_count += 1
        print(f"\n{'='*60}")
        print(f"STEP {cluster_count}: Finding next closest address")
        print(f"{'='*60}")
        print(f"Current location: {get_display(current_location_name)}")
        print(f"Clusters remaining: {total_clusters - len(visited_clusters)}")

        # Find closest unvisited
        closest = find_closest_unvisited_address(df_addresses, current_lat, current_lng, visited_clusters)
        if closest is None:
            break

        cluster_group = int(closest['cluster_group'])
        street = closest['Street_Name']
        house = closest['House_Number']
        closest_addr_name = f"{street} {house}"

        print(f"\n✓ Found: {get_display(closest_addr_name)}")
        print(f"✓ Cluster: {cluster_group}")

        # Load matrix for this cluster
        print(f"\nLoading matrix for cluster {cluster_group}...")
        matrix_df = load_distance_matrix(cluster_group, matrix_folder)
        if matrix_df is None:
            print(f"⚠️ Skipping cluster {cluster_group} because its matrix file is missing.")
            visited_clusters.add(cluster_group)
            continue

        # Get all addresses in this cluster for lookup
        cluster_addresses = df_addresses[df_addresses['cluster_group'] == cluster_group]
        cluster_lookup = {}
        
        # Robustly match the exact names found in the Excel matrix to their coordinates
        for matrix_addr in matrix_df.columns:
            if matrix_addr == "ORIGIN":
                continue
                
            matched_row = None
            for idx, row in cluster_addresses.iterrows():
                street = str(row.get('Street_Name', '')).strip()
                house = str(row.get('House_Number', '')).replace('.0', '').strip()
                
                # Check if both street and house number exist in the matrix label
                if street and street in str(matrix_addr) and house in str(matrix_addr):
                    matched_row = row
                    break
            
            # Fallback to the first address in the cluster if no exact string match is found
            if matched_row is None:
                matched_row = cluster_addresses.iloc[0]
                
            cluster_lookup[matrix_addr] = {
                'LAT': matched_row['LAT'], 
                'LNG': matched_row['LNG'],
                'detailed_addresses': matched_row['detailed_addresses']
            }

        # Check if single address in cluster
        if len(matrix_df) == 1:
            print("\n⚠️  Single address in cluster - skipping TSP")
            single_addr = list(matrix_df.index)[0]
            
            detailed_str = cluster_lookup.get(single_addr, {}).get('detailed_addresses', single_addr)
            if pd.isna(detailed_str) or not detailed_str:
                detailed_str = single_addr
            sub_addresses = [s.strip() for s in str(detailed_str).split(',') if s.strip()]
            if not sub_addresses:
                sub_addresses = [single_addr]
                
            single_time = np.random.randint(1, 8)  # Random time in test mode
            
            extra_travel = (len(sub_addresses) - 1) * 1
            total_travel = single_time + extra_travel
            delivery_time = len(sub_addresses) * DELIVERY_TIME_MINUTES
            cluster_total_time = total_travel + delivery_time
            
            route_names = [get_display(current_location_name)] + [get_display(s) for s in sub_addresses]
            hebrew_route_str = " → ".join(route_names)
            
            journey_results.append({
                'cluster': cluster_group,
                'route': [0, 1],
                'route_str': hebrew_route_str,
                'travel_time': total_travel,
                'delivery_time': delivery_time,
                'cost': cluster_total_time,
                'endpoint': sub_addresses[-1]
            })
            total_journey_time += cluster_total_time
            visited_clusters.add(cluster_group)
            full_journey_path.extend(route_names[1:])
            
            if single_addr in cluster_lookup:
                current_lat = cluster_lookup[single_addr]['LAT']
                current_lng = cluster_lookup[single_addr]['LNG']
            else: # Fallback
                current_lat = cluster_addresses.iloc[0]['LAT']
                current_lng = cluster_addresses.iloc[0]['LNG']
            current_location_name = sub_addresses[-1]
            print(f"\n✓ Route: {hebrew_route_str}")
            print(f"✓ Travel time: {total_travel:.2f} min ({single_time} base + {extra_travel} extra) | Delivery time: {delivery_time:.2f} min")
            print(f"✓ Total time for cluster: {cluster_total_time:.2f} minutes")
            continue

        # Add origin to matrix (TEST MODE - random numbers)
        modified_matrix = add_origin_to_matrix_test(matrix_df, current_lat, current_lng, cluster_lookup)

        # Run TSP
        tsp_route, tsp_cost = run_tsp_on_matrix(modified_matrix)
        
        # Extract endpoint index and matrix name
        endpoint_idx = tsp_route[-1]
        endpoint_matrix_name = list(modified_matrix.columns)[endpoint_idx]
        
        # Build real address names for printing and track expanded deliveries
        route_names = []
        total_deliveries = 0
        extra_travel_time = 0
        
        for idx in tsp_route:
            name = list(modified_matrix.columns)[idx]
            if name == "ORIGIN":
                route_names.append(get_display(current_location_name))
            else:
                detailed_str = cluster_lookup.get(name, {}).get('detailed_addresses', name)
                if pd.isna(detailed_str) or not detailed_str:
                    detailed_str = name
                    
                sub_addresses = [s.strip() for s in str(detailed_str).split(',') if s.strip()]
                if not sub_addresses:
                    sub_addresses = [name]
                
                for sub in sub_addresses:
                    route_names.append(get_display(sub))
                
                total_deliveries += len(sub_addresses)
                extra_travel_time += (len(sub_addresses) - 1) * 1
                
        # Calculate delivery times
        delivery_time = total_deliveries * DELIVERY_TIME_MINUTES
        final_travel_time = tsp_cost + extra_travel_time
        cluster_total_time = final_travel_time + delivery_time
        
        # Determine final endpoint display name
        detailed_end_str = cluster_lookup.get(endpoint_matrix_name, {}).get('detailed_addresses', endpoint_matrix_name)
        if pd.isna(detailed_end_str) or not detailed_end_str:
            detailed_end_str = endpoint_matrix_name
        end_sub_addresses = [s.strip() for s in str(detailed_end_str).split(',') if s.strip()]
        display_endpoint = end_sub_addresses[-1] if end_sub_addresses else endpoint_matrix_name
                
        hebrew_route_str = " → ".join(route_names)
        print(f"\n✓ TSP Route: {hebrew_route_str}")
        print(f"✓ Travel time: {final_travel_time:.2f} min ({tsp_cost:.2f} base + {extra_travel_time:.2f} extra) | Delivery time: {delivery_time:.2f} min")
        print(f"✓ Total time for cluster: {cluster_total_time:.2f} minutes")

        # Add to the full journey path, skipping the first element (which is the start)
        full_journey_path.extend(route_names[1:])

        # Store results
        journey_results.append({
            'cluster': cluster_group,
            'route': tsp_route,
            'route_str': hebrew_route_str,
            'travel_time': final_travel_time,
            'delivery_time': delivery_time,
            'cost': cluster_total_time,
            'endpoint': display_endpoint
        })
        
        total_journey_time += cluster_total_time
        visited_clusters.add(cluster_group)

        # Update current location
        if endpoint_matrix_name in cluster_lookup:
            current_lat = cluster_lookup[endpoint_matrix_name]['LAT']
            current_lng = cluster_lookup[endpoint_matrix_name]['LNG']
        else: # Fallback
            print(f"⚠️ Warning: '{endpoint_matrix_name}' not found in lookup. Using fallback.")
            current_lat = cluster_addresses.iloc[-1]['LAT']
            current_lng = cluster_addresses.iloc[-1]['LNG']
        current_location_name = display_endpoint

    # 6. Print final results
    print("\n" + "="*60)
    print("✅ JOURNEY COMPLETE (TEST MODE)")
    print("="*60)
    
    print("\n" + "="*60)
    print("🗺️  FULL JOURNEY ROUTE")
    print("="*60)
    print(" → ".join(full_journey_path))
    
    total_addresses_delivered = len(full_journey_path) - 1
    
    print(f"\n📊 SUMMARY:")
    print(f"  Total addresses delivered: {total_addresses_delivered}")
    print(f"  Clusters visited: {len(visited_clusters)}/{total_clusters}")
    print(f"  Total journey time: {total_journey_time:.2f} minutes ({total_journey_time/60:.2f} hours)")
    
    print(f"\n📋 CLUSTER DETAILS:")
    for i, result in enumerate(journey_results, 1):
        print(f"  {i}. Cluster {result['cluster']}: {result['cost']:.2f} min (Travel: {result['travel_time']:.2f}, Drop-off: {result['delivery_time']:.2f}) → {get_display(result['endpoint'])}")

    # 7. Export to Excel
    print("\n💾 Exporting to Excel...")
    export_journey_to_excel(journey_results, total_journey_time, df_addresses, total_addresses_delivered, test_mode=True)


def export_journey_to_excel(journey_results, total_time, df_addresses, total_addresses_delivered=0, test_mode=False):
    """Export journey to Excel file"""
    filename = "multi_cluster_journey_TEST.xlsx" if test_mode else "multi_cluster_journey.xlsx"
    excel_path = Path.home() / "Desktop" / filename
    
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        # Summary sheet
        summary_data = {
            'Metric': ['Total Addresses Delivered', 'Total Clusters', 'Total Time (min)', 'Total Time (hours)', 'Starting Point', 'Final Endpoint', 'Mode'],
            'Value': [
                total_addresses_delivered,
                len(journey_results),
                f"{total_time:.2f}",
                f"{total_time/60:.2f}",
                "WAREHOUSE ESHTAOL",
                journey_results[-1]['endpoint'] if journey_results else "N/A",
                "TEST (Random times)" if test_mode else "Production (API)"
            ]
        }
        pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False)
        
        # Detailed journey
        journey_data = []
        for i, result in enumerate(journey_results, 1):
            journey_data.append({
                'Step': i,
                'Cluster': result['cluster'],
                'Total Time (min)': f"{result['cost']:.2f}",
                'Travel Time': f"{result['travel_time']:.2f}",
                'Delivery Time': f"{result['delivery_time']:.2f}",
                'Full Route Path': result['route_str'],
                'Endpoint': result['endpoint'],
                'Cumulative Time': f"{sum(r['cost'] for r in journey_results[:i]):.2f}"
            })
        pd.DataFrame(journey_data).to_excel(writer, sheet_name='Journey', index=False)
    
    print(f"✅ Saved to: {excel_path}")


if __name__ == "__main__":
    print("🧪 RUNNING IN TEST MODE - Using random times instead of API")
    run_multi_cluster_tsp_test()
