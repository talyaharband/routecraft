import pandas as pd
import numpy as np
import math
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import time

try:
    from bidi.algorithm import get_display
except ModuleNotFoundError:
    get_display = lambda value: value

# Warehouse coordinates
WAREHOUSE_LAT = 31.77927525
WAREHOUSE_LNG = 35.0105885
DELIVERY_TIME_MINUTES = 4
MAX_SHIFT_MINUTES = 420  # 7 hours in minutes


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
    unvisited_df = df[~df['cluster_group'].astype(str).isin(visited_clusters)].copy()
    
    if unvisited_df.empty:
        return None
    
    unvisited_df['air_distance'] = unvisited_df.apply(
        lambda row: calculate_haversine(current_lat, current_lng, row['LAT'], row['LNG']),
        axis=1
    )
    
    closest_idx = unvisited_df['air_distance'].idxmin()
    return df.loc[closest_idx]


def find_closest_unvisited_address_city_aware(df, current_lat, current_lng, visited_clusters, active_city):
    """Find closest unvisited address to current location, prioritizing the active city"""
    unvisited_df = df[~df['cluster_group'].astype(str).isin(visited_clusters)].copy()
    
    if unvisited_df.empty:
        return None
    
    unvisited_df['air_distance'] = unvisited_df.apply(
        lambda row: calculate_haversine(current_lat, current_lng, row['LAT'], row['LNG']),
        axis=1
    )
    
    if active_city is not None and 'City' in unvisited_df.columns:
        city_df = unvisited_df[unvisited_df['City'] == active_city]
        if not city_df.empty:
            closest_idx = city_df['air_distance'].idxmin()
            return df.loc[closest_idx], active_city
            
    closest_idx = unvisited_df['air_distance'].idxmin()
    next_city = df.loc[closest_idx].get('City', None)
    return df.loc[closest_idx], next_city


def load_distance_matrix(cluster_group, matrix_folder):
    """Load distance matrix for a specific cluster"""
    folder_path = Path(matrix_folder)
    
    # Search recursively in the folder, and intelligently parse the number mathematically
    all_matrices = list(folder_path.rglob("matrix_*.xlsx"))
    all_matrices.extend(folder_path.rglob("05_distance_matrix_group-*.xlsx"))
    matrix_files = []
    
    target_prefix = f"matrix_{cluster_group}_"
    for file in all_matrices:
        if file.name.startswith(target_prefix) or file.name == f"05_distance_matrix_group-{cluster_group}.xlsx":
            matrix_files.append(file)
            
    if not matrix_files:
        for file in all_matrices:
            parts = file.name.split('_')
            if len(parts) >= 2:
                try:
                    # Compares the number mathematically if they are standard integers
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
        
        # Clean phantom columns that pandas might add if Excel was manually edited
        unnamed_cols = [c for c in df.columns if str(c).startswith('Unnamed:')]
        if unnamed_cols:
            df = df.drop(columns=unnamed_cols)
            
        # Ensure matrix is perfectly square
        min_dim = min(df.shape[0], df.shape[1])
        df = df.iloc[:min_dim, :min_dim]
        
        # Force exact string match between index and columns to prevent KeyError in .loc
        df.index = df.columns
        
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


def run_multi_cluster_tsp(kmeans_path, matrix_folder, output_path=None, test_mode=True):
    kmeans_path = Path(kmeans_path)
    matrix_folder = Path(matrix_folder)

    # 1. Load data
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
    print("🚀 STARTING MULTI-VEHICLE TSP JOURNEY (TEST MODE - Sequential)")
    print("="*60)

    completed_clusters = set()
    delivered_matrix_nodes = {} # cluster_group -> set(nodes)
    all_drivers_results = {}
    driver_id = 1
    
    total_clusters = len(df_addresses['cluster_group'].astype(str).unique())
    active_city = None
    sequential_single_time = 1  # For single-address clusters

    while len(completed_clusters) < total_clusters:
        print(f"\n{'='*60}")
        print(f"🚚 STARTING SHIFT FOR DRIVER {driver_id}")
        print(f"{'='*60}")
        
        current_lat = WAREHOUSE_LAT
        current_lng = WAREHOUSE_LNG
        current_location_name = "WAREHOUSE ESHTAOL"
        
        driver_journey = []
        driver_time = 0
        driver_addresses_delivered = 0
        
        while driver_time < MAX_SHIFT_MINUTES and len(completed_clusters) < total_clusters:
            # 1. Find closest unvisited
            closest_info = find_closest_unvisited_address_city_aware(
                df_addresses, current_lat, current_lng, completed_clusters, active_city
            )
            
            if closest_info is None:
                if active_city is not None:
                    print(f"\n🏙️  City '{get_display(str(active_city))}' is complete. Searching globally for next city...")
                    active_city = None
                    continue
                else:
                    break
                    
            closest, next_city = closest_info
            
            if active_city is not None and next_city != active_city:
                print(f"\n🏙️  City '{get_display(str(active_city))}' is complete! Moving to new city: '{get_display(str(next_city))}'")
                
            active_city = next_city
            cluster_group = str(closest['cluster_group'])
            
            print(f"\n📍 Next Target: {get_display(f'{closest['Street_Name']} {closest['House_Number']}')}")
            print(f"   City: {get_display(str(active_city)) if active_city else 'None'} | Cluster: {cluster_group} | Time so far: {driver_time:.2f}/420 mins")
            
            # Load matrix
            matrix_df = load_distance_matrix(cluster_group, matrix_folder)
            if matrix_df is None:
                completed_clusters.add(cluster_group)
                continue
                
            # Filter out already delivered nodes
            delivered_nodes_here = delivered_matrix_nodes.get(cluster_group, set())
            remaining_cols = [c for c in matrix_df.columns if c not in delivered_nodes_here]
            
            if len(remaining_cols) == 0:
                completed_clusters.add(cluster_group)
                continue
                
            matrix_df = matrix_df.loc[remaining_cols, remaining_cols]
            
            # Build cluster lookup
            cluster_addresses = df_addresses[df_addresses['cluster_group'].astype(str) == cluster_group]
            cluster_lookup = {}
            
            for matrix_addr in matrix_df.columns:
                if matrix_addr == "ORIGIN": continue
                matched_row = None
                for idx, row in cluster_addresses.iterrows():
                    street = str(row.get('Street_Name', '')).strip()
                    house = str(row.get('House_Number', '')).replace('.0', '').strip()
                    if street and street in str(matrix_addr) and house in str(matrix_addr):
                        matched_row = row
                        break
                
                if matched_row is None:
                    matched_row = cluster_addresses.iloc[0]
                    
                cluster_lookup[matrix_addr] = {
                    'LAT': matched_row['LAT'],
                    'LNG': matched_row['LNG'],
                    'detailed_addresses': matched_row.get('detailed_addresses', matrix_addr)
                }
                
            shift_ended = False
            
            if len(matrix_df.columns) == 1:
                node_name = list(matrix_df.columns)[0]
                detailed_str = cluster_lookup.get(node_name, {}).get('detailed_addresses', node_name)
                if pd.isna(detailed_str) or not detailed_str: detailed_str = node_name
                sub_addresses = [s.strip() for s in str(detailed_str).split(',') if s.strip()]
                if not sub_addresses: sub_addresses = [node_name]
                
                step_travel_time = sequential_single_time
                sequential_single_time += 1
                step_extra_travel = (len(sub_addresses) - 1) * 1
                step_delivery_time = len(sub_addresses) * DELIVERY_TIME_MINUTES
                total_step_time = step_travel_time + step_extra_travel + step_delivery_time
                
                if driver_time + total_step_time > MAX_SHIFT_MINUTES and driver_time > 0:
                    print(f"⚠️  Driver {driver_id} reached 7-hour limit ({driver_time + total_step_time:.2f} > 420)! Ending shift.")
                    shift_ended = True
                else:
                    if driver_time == 0 and total_step_time > MAX_SHIFT_MINUTES:
                        print(f"⚠️  Warning: Single location takes {total_step_time:.2f} mins (> 7 hours). Forcing delivery to prevent getting stuck.")
                    driver_time += total_step_time
                    driver_addresses_delivered += len(sub_addresses)
                    
                    route_names = [get_display(current_location_name)] + [get_display(s) for s in sub_addresses]
                    hebrew_route_str = " → ".join(route_names)
                    
                    driver_journey.append({
                        'cluster': cluster_group,
                        'city': active_city,
                        'route_str': hebrew_route_str,
                        'travel_time': step_travel_time + step_extra_travel,
                        'delivery_time': step_delivery_time,
                        'cost': total_step_time,
                        'endpoint': sub_addresses[-1]
                    })
                    
                    if cluster_group not in delivered_matrix_nodes:
                        delivered_matrix_nodes[cluster_group] = set()
                    delivered_matrix_nodes[cluster_group].add(node_name)
                    completed_clusters.add(cluster_group)
                    
                    current_lat = cluster_lookup[node_name]['LAT']
                    current_lng = cluster_lookup[node_name]['LNG']
                    current_location_name = sub_addresses[-1]
                    
                    print(f"✓ Route: {hebrew_route_str}")
                    print(f"✓ Time for node: {total_step_time:.2f} minutes")
                    
            else:
                # TSP Logic
                modified_matrix = add_origin_to_matrix_test(matrix_df, current_lat, current_lng, cluster_lookup)
                tsp_route, _ = run_tsp_on_matrix(modified_matrix)
                
                route_names_segment = [get_display(current_location_name)]
                cluster_travel = 0
                cluster_delivery = 0
                nodes_delivered_this_shift = []
                
                prev_idx = tsp_route[0]
                
                for curr_idx in tsp_route[1:]:
                    step_travel_time = modified_matrix.iloc[prev_idx, curr_idx]
                    node_name = modified_matrix.columns[curr_idx]
                    
                    detailed_str = cluster_lookup.get(node_name, {}).get('detailed_addresses', node_name)
                    if pd.isna(detailed_str) or not detailed_str: detailed_str = node_name
                    sub_addresses = [s.strip() for s in str(detailed_str).split(',') if s.strip()]
                    if not sub_addresses: sub_addresses = [node_name]
                    
                    step_extra_travel = (len(sub_addresses) - 1) * 1
                    step_delivery_time = len(sub_addresses) * DELIVERY_TIME_MINUTES
                    total_step_time = step_travel_time + step_extra_travel + step_delivery_time
                    
                    if driver_time + total_step_time > MAX_SHIFT_MINUTES:
                        if driver_time == 0:
                            print(f"⚠️  Warning: Location takes {total_step_time:.2f} mins (> 7 hours). Forcing delivery to prevent getting stuck.")
                        else:
                            print(f"⚠️  Driver {driver_id} reached 7-hour limit ({driver_time + total_step_time:.2f} > 420)! Ending shift.")
                            shift_ended = True
                            break
                        
                    driver_time += total_step_time
                    driver_addresses_delivered += len(sub_addresses)
                    cluster_travel += (step_travel_time + step_extra_travel)
                    cluster_delivery += step_delivery_time
                    
                    for sub in sub_addresses:
                        route_names_segment.append(get_display(sub))
                        
                    nodes_delivered_this_shift.append(node_name)
                    
                    current_lat = cluster_lookup[node_name]['LAT']
                    current_lng = cluster_lookup[node_name]['LNG']
                    current_location_name = sub_addresses[-1]
                    
                    prev_idx = curr_idx
                    
                if len(nodes_delivered_this_shift) > 0:
                    hebrew_route_str = " → ".join(route_names_segment)
                    driver_journey.append({
                        'cluster': cluster_group,
                        'city': active_city,
                        'route_str': hebrew_route_str,
                        'travel_time': cluster_travel,
                        'delivery_time': cluster_delivery,
                        'cost': cluster_travel + cluster_delivery,
                        'endpoint': current_location_name
                    })
                    
                    print(f"✓ Route: {hebrew_route_str}")
                    print(f"✓ Travel: {cluster_travel:.2f} min | Delivery: {cluster_delivery:.2f} min")
                    
                    if cluster_group not in delivered_matrix_nodes:
                        delivered_matrix_nodes[cluster_group] = set()
                    for n in nodes_delivered_this_shift:
                        delivered_matrix_nodes[cluster_group].add(n)
                        
                if not shift_ended:
                    completed_clusters.add(cluster_group)
                    
            if shift_ended:
                break
                
        # End of shift tracking
        all_drivers_results[driver_id] = {
            'journey': driver_journey,
            'time': driver_time,
            'addresses': driver_addresses_delivered
        }
        
        print(f"\n✅ SHIFT COMPLETE - DRIVER {driver_id}")
        print(f"   Time worked: {driver_time:.2f} mins ({(driver_time/60):.2f} hours)")
        print(f"   Addresses delivered: {driver_addresses_delivered}")
        
        driver_id += 1

    # 6. Print final results
    print("\n" + "="*60)
    print("✅ ALL DELIVERIES COMPLETE (TEST MODE)")
    print("="*60)
    
    print(f"\n📊 SUMMARY:")
    for d_id, data in all_drivers_results.items():
        print(f"  Driver {d_id}: {data['addresses']} addresses | {data['time']:.2f} min ({(data['time']/60):.2f} hours)")

    print("\n" + "="*60)
    print("🗺️ FULL ROUTES BY DRIVER")
    print("="*60)
    for d_id, data in all_drivers_results.items():
        full_route_nodes = []
        for step in data['journey']:
            parts = step['route_str'].split(" → ")
            if not full_route_nodes:
                full_route_nodes.extend(parts)
            else:
                if full_route_nodes[-1] == parts[0]:
                    full_route_nodes.extend(parts[1:])
                else:
                    full_route_nodes.extend(parts)
        full_route_str = " → ".join(full_route_nodes)
        print(f"\n🚚 Driver {d_id} Route:")
        print(f"   {full_route_str}")

    # 7. Export to Excel
    print("\n💾 Exporting to Excel...")
    return export_journey_to_excel(
        all_drivers_results,
        df_addresses,
        test_mode=test_mode,
        output_path=output_path,
    )


def run_multi_cluster_tsp_test():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    print("Step 1: Select k-means output file")
    kmeans_path = filedialog.askopenfilename(
        title="Select k-means output Excel (with cluster_group)",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not kmeans_path:
        print("Cancelled")
        return

    print("\nStep 2: Select folder with distance matrices")
    matrix_folder = filedialog.askdirectory(title="Select folder with matrix files")
    if not matrix_folder:
        print("Cancelled")
        return

    run_multi_cluster_tsp(kmeans_path, matrix_folder, test_mode=True)


def export_journey_to_excel(all_drivers_results, df_addresses, test_mode=False, output_path=None):
    """Export journey to Excel file"""
    if output_path is None:
        filename = "multi_vehicle_journey_TEST.xlsx" if test_mode else "multi_vehicle_journey.xlsx"
        excel_path = Path.home() / "Desktop" / filename
    else:
        excel_path = Path(output_path)
        excel_path.parent.mkdir(parents=True, exist_ok=True)
    
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        # Summary sheet
        summary_data = []
        for d_id, data in all_drivers_results.items():
            summary_data.append({
                'Driver': f"Driver {d_id}",
                'Addresses Delivered': data['addresses'],
                'Clusters Visited': len(data['journey']),
                'Total Time (min)': f"{data['time']:.2f}",
                'Total Time (hours)': f"{(data['time']/60):.2f}",
                'Mode': "TEST (Sequential times)" if test_mode else "Production (API)"
            })
            
        pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False)
        
        # Detailed journey
        journey_data = []
        for d_id, data in all_drivers_results.items():
            cumulative = 0
            for i, result in enumerate(data['journey'], 1):
                cumulative += result['cost']
                journey_data.append({
                    'Driver': f"Driver {d_id}",
                    'Step': i,
                    'City': result.get('city', 'N/A'),
                    'Cluster': result['cluster'],
                    'Total Time (min)': f"{result['cost']:.2f}",
                    'Travel Time': f"{result['travel_time']:.2f}",
                    'Delivery Time': f"{result['delivery_time']:.2f}",
                    'Route Path': result['route_str'],
                    'Endpoint': result['endpoint'],
                    'Shift Time So Far': f"{cumulative:.2f}"
                })
                
        if journey_data:
            pd.DataFrame(journey_data).to_excel(writer, sheet_name='Detailed Routes', index=False)
        else:
            pd.DataFrame([{'Info': 'No routes delivered'}]).to_excel(writer, sheet_name='Detailed Routes', index=False)
    
    print(f"✅ Saved to: {excel_path}")
    return excel_path


if __name__ == "__main__":
    print("🧪 RUNNING IN TEST MODE - Using sequential times instead of API")
    run_multi_cluster_tsp_test()
