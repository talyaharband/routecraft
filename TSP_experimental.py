import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog
from tqdm import tqdm
import os
from bidi.algorithm import get_display

# Debug flag — turn on to see detailed prints
DEBUG = False


# ============================
# 1. SELECT THE EXCEL FILE (PyCharm/Local Version)
# ============================
def get_file_path():
    root = tk.Tk()
    root.withdraw()  # Hide the main tkinter window
    print("Opening file selector... Please select your Excel matrix.")
    file_path = filedialog.askopenfilename(
        title="Select Excel File (Minutes Matrix)",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not file_path:
        print("❌ No file selected. Exiting.")
        exit()
    return file_path


filename = get_file_path()

# ============================
# 2. LOAD THE DISTANCE MATRIX
# ============================
df = pd.read_excel(filename, header=0, index_col=0)

if DEBUG:
    print("Raw dataframe:")
    print(df)  # Changed from display() to print() for PyCharm

# Store address labels for later use
address_labels = df.index.tolist()

distance_matrix = df.to_numpy()
n = distance_matrix.shape[0]

print(f"✅ Loaded file: {os.path.basename(filename)}")
print("Distance matrix shape:", distance_matrix.shape)

base_df = pd.DataFrame(distance_matrix,
                       index=range(n),
                       columns=range(n))

base_df = base_df.astype("object")


# ============================
# 3. FUNCTION TO CREATE CASE MATRIX (WITH X)
# ============================
def create_case_matrix(end_address):
    df_case = base_df.copy()
    df_case[0] = "X"
    df_case.loc[end_address, :] = "X"
    for i in range(n):
        df_case.loc[i, i] = "X"
    return df_case


# ============================
# 4. NUMERIC CASE MATRIX FOR LB (WITHOUT X)
# ============================
def create_numeric_case_matrix(end_address):
    mat = distance_matrix.copy().astype(float)
    mat[:, 0] = np.inf
    mat[end_address, :] = np.inf
    for i in range(n):
        mat[i, i] = np.inf
    return mat


# ============================================================
# 5. INITIAL LOWER BOUND (ROW + COLUMN REDUCTION)
# ============================================================
def compute_LB_and_reduced_matrix(end_address):
    mat = create_numeric_case_matrix(end_address)
    lb = 0.0
    for i in range(n):
        row = mat[i, :]
        finite_vals = row[np.isfinite(row)]
        if finite_vals.size == 0: continue
        m = finite_vals.min()
        if m > 0:
            mat[i, :] = np.where(np.isfinite(mat[i, :]), mat[i, :] - m, mat[i, :])
            lb += m
    for j in range(n):
        col = mat[:, j]
        finite_vals = col[np.isfinite(col)]
        if finite_vals.size == 0: continue
        m = finite_vals.min()
        if m > 0:
            mat[:, j] = np.where(np.isfinite(mat[:, j]), mat[:, j] - m, mat[:, j])
            lb += m
    reduced_numeric = mat.copy()
    display_mat = pd.DataFrame(mat.copy(), index=range(n), columns=range(n))
    display_mat = display_mat.replace(np.inf, "X")
    return lb, reduced_numeric, display_mat


# ============================================================
# 6. NEAREST NEIGHBOR UB
# ============================================================
def nearest_neighbor_UB(end_address):
    visited = set([0]);
    path = [0];
    current = 0
    nodes_to_visit = [i for i in range(1, n) if i != end_address]
    while nodes_to_visit:
        next_node = min(nodes_to_visit, key=lambda j: distance_matrix[current][j])
        path.append(next_node);
        visited.add(next_node);
        nodes_to_visit.remove(next_node);
        current = next_node
    path.append(end_address)
    total_cost = sum(distance_matrix[path[i]][path[i + 1]] for i in range(len(path) - 1))
    return path, total_cost


# ============================================================
# 7. 2-OPT
# ============================================================
def route_cost(route):
    return sum(distance_matrix[route[i]][route[i + 1]] for i in range(len(route) - 1))


def two_opt(route, end_address):
    improved = True;
    best_route = route;
    best_cost = route_cost(route)
    while improved:
        improved = False
        for i in range(1, len(best_route) - 2):
            for j in range(i + 1, len(best_route) - 1):
                new_route = best_route[:i] + best_route[i:j + 1][::-1] + best_route[j + 1:]
                new_cost = route_cost(new_route)
                if new_cost < best_cost:
                    best_route, best_cost, improved = new_route, new_cost, True
                    break
            if improved: break
    return best_route, best_cost


# ============================================================
# 8. COLLECT CASE DATA (With Progress Bar)
# ============================================================
cases = []
print("\n--- Starting Heuristic Calculations ---")
for end in tqdm(range(1, n), desc="Analyzing Ends"):
    lb, reduced_numeric, reduced_display = compute_LB_and_reduced_matrix(end)
    nn_path, nn_cost = nearest_neighbor_UB(end)
    opt2_path, opt2_cost = two_opt(nn_path, end)

    cases.append({
        "end": end, "initial_lb": lb, "initial_matrix": reduced_numeric,
        "nn_cost": nn_cost, "nn_route": nn_path,
        "opt2_cost": opt2_cost, "opt2_route": opt2_path,
        "ub": opt2_cost, "ub_route": opt2_path
    })

# Print headers after progress bar to maintain your original output format
for case in cases:
    end = case["end"]
    print("\n====================================================")
    print(f"                 OPTION: END AT {end}")
    print("====================================================\n")
    print(f"Initial Lower Bound (LB): {case['initial_lb']}")
    
    print("\nNearest Neighbor:")
    print("Route:", " -> ".join(map(str, case['nn_route'])))
    print(f"Cost: {case['nn_cost']:.2f} minutes")
    
    print("\n2-Opt Optimized:")
    print("Route:", " -> ".join(map(str, case['opt2_route'])))
    print(f"Cost: {case['opt2_cost']:.2f} minutes")
    
    improvement = case['nn_cost'] - case['opt2_cost']
    if improvement > 0:
        print(f"\n✅ Improvement: {improvement:.2f} minutes ({(improvement/case['nn_cost']*100):.1f}%)")
    else:
        print("\n➡️  No improvement from 2-Opt")
    print("\n----------------------------------------------------\n")


# ============================================================
# 9-10. B&B HELPERS & EXPLORATION (Logic Unchanged)
# ============================================================

def reduce_matrix(mat):
    mat = mat.copy();
    added_lb = 0.0
    for i in range(mat.shape[0]):
        row = mat[i, :];
        finite_vals = row[np.isfinite(row)]
        if finite_vals.size:
            m = finite_vals.min()
            if m > 0: mat[i, :] = np.where(np.isfinite(mat[i, :]), mat[i, :] - m, mat[i, :]); added_lb += m
    for j in range(mat.shape[1]):
        col = mat[:, j];
        finite_vals = col[np.isfinite(col)]
        if finite_vals.size:
            m = finite_vals.min()
            if m > 0: mat[:, j] = np.where(np.isfinite(mat[:, j]), mat[:, j] - m, mat[:, j]); added_lb += m
    return mat, added_lb


def choose_branch_zero(mat, rows, cols, depth):
    best_penalty = -1;
    best_pair = (None, None)
    for r_idx in range(mat.shape[0]):
        for c_idx in range(mat.shape[1]):
            if mat[r_idx, c_idx] != 0: continue
            row = mat[r_idx, :];
            row_mask = np.isfinite(row);
            row_mask[c_idx] = False
            row_vals = row[row_mask];
            row_min = row_vals.min() if row_vals.size else 0
            col = mat[:, c_idx];
            col_mask = np.isfinite(col);
            col_mask[r_idx] = False
            col_vals = col[col_mask];
            col_min = col_vals.min() if col_vals.size else 0
            penalty = row_min + col_min;
            i, j = rows[r_idx], cols[c_idx]
            if penalty > best_penalty or (penalty == best_penalty and best_pair[0] is not None and i < best_pair[0]):
                best_penalty, best_pair = penalty, (i, j)
    return best_pair


def build_paths_from_forced(forced_edges):
    succ, pred = {}, {}
    for i, j in forced_edges: succ[i], pred[j] = j, i
    return succ, pred


def block_cycle_edges(mat, rows, cols, forced_edges, end_node):
    succ, pred = build_paths_from_forced(forced_edges)

    def block_edge(u, v):
        if u in rows and v in cols: mat[rows.index(u), cols.index(v)] = np.inf

    for i, j in forced_edges: block_edge(j, i)
    starts = [node for node in succ.keys() if node not in pred]
    for start in starts:
        current, visited = start, {start}
        while current in succ:
            nxt = succ[current]
            if nxt in visited: break
            visited.add(nxt);
            current = nxt
        block_edge(current, start)
        if start == 0 and current != end_node and len(visited) < n - 1: block_edge(current, end_node)


def reconstruct_route_from_forced(forced_edges, end_node):
    if len(forced_edges) != n - 1: return None
    succ, pred = build_paths_from_forced(forced_edges)
    if 0 not in succ: return None
    route, current, visited = [0], 0, {0}
    while current in succ:
        nxt = succ[current]
        if nxt in visited: return None
        route.append(nxt);
        visited.add(nxt);
        current = nxt
    return route if route[-1] == end_node and len(route) == n else None


def apply_forced_moves(mat, rows, cols, forced_edges, end_node, lb, depth, current_ub):
    changed = True
    while changed:

        changed = False
        for r_idx in range(len(rows)):
            if np.isfinite(mat[r_idx, :]).sum() == 1:
                c_idx = np.where(np.isfinite(mat[r_idx, :]))[0][0];
                i, j = rows[r_idx], cols[c_idx]
                forced_edges.append((i, j))


                mat = np.delete(np.delete(mat, r_idx, axis=0), c_idx, axis=1)
                rows.pop(r_idx);
                cols.pop(c_idx);
                block_cycle_edges(mat, rows, cols, forced_edges, end_node)
                mat, add_lb = reduce_matrix(mat);
                lb += add_lb
                if lb >= current_ub: return mat, rows, cols, forced_edges, lb, True
                changed = True;
                break
        if not changed:
            for c_idx in range(len(cols)):
                if np.isfinite(mat[:, c_idx]).sum() == 1:
                    r_idx = np.where(np.isfinite(mat[:, c_idx]))[0][0];
                    i, j = rows[r_idx], cols[c_idx]
                    forced_edges.append((i, j))
                    mat = np.delete(np.delete(mat, r_idx, axis=0), c_idx, axis=1)
                    rows.pop(r_idx);
                    cols.pop(c_idx);
                    block_cycle_edges(mat, rows, cols, forced_edges, end_node)
                    mat, add_lb = reduce_matrix(mat);
                    lb += add_lb
                    if lb >= current_ub: return mat, rows, cols, forced_edges, lb, True
                    changed = True;
                    break
    return mat, rows, cols, forced_edges, lb, False


def explore_node(node, current_ub, best_route, depth):
    lb, mat, rows, cols, forced_edges, end_node = node["lb"], node["matrix"], node["rows"], node["cols"], node[
        "forced_edges"], node["end"]
    if lb >= current_ub: return current_ub, best_route
    mat, rows, cols, forced_edges, lb, pruned = apply_forced_moves(mat, rows, cols, forced_edges, end_node, lb, depth,
                                                                   current_ub)
    if pruned: return current_ub, best_route
    route = reconstruct_route_from_forced(forced_edges, end_node)
    if route is not None:
        cost = route_cost(route)
        print(f"Complete route found: {' → '.join(map(str, route))}  |  Cost = {cost:.2f} min")
        if cost < current_ub:
            print(f"UB improved: {current_ub:.2f} → {cost:.2f}")
            return cost, route
        return current_ub, best_route
    if mat.size == 0: return current_ub, best_route
    i, j = choose_branch_zero(mat, rows, cols, depth)
    if i is None: return current_ub, best_route

    # Branching
    r_idx, c_idx = rows.index(i), cols.index(j)
    # WITH
    mat_with = np.delete(np.delete(mat.copy(), r_idx, axis=0), c_idx, axis=1)
    rows_w, cols_w, forced_w = rows[:r_idx] + rows[r_idx + 1:], cols[:c_idx] + cols[c_idx + 1:], forced_edges + [(i, j)]
    block_cycle_edges(mat_with, rows_w, cols_w, forced_w, end_node)
    mat_w_red, add_w = reduce_matrix(mat_with)

    # WITHOUT
    mat_wo = mat.copy();
    mat_wo[r_idx, c_idx] = np.inf
    mat_wo_red, add_wo = reduce_matrix(mat_wo)

    children = []
    if lb + add_w < current_ub:
        children.append(
            {"lb": lb + add_w, "matrix": mat_w_red, "rows": rows_w, "cols": cols_w, "forced_edges": forced_w,
             "end": end_node, "forbidden_edges": []})
    if lb + add_wo < current_ub:
        children.append({"lb": lb + add_wo, "matrix": mat_wo_red, "rows": rows.copy(), "cols": cols.copy(),
                         "forced_edges": forced_edges.copy(), "end": end_node, "forbidden_edges": []})

    children.sort(key=lambda x: x["lb"])
    for child in children:
        current_ub, best_route = explore_node(child, current_ub, best_route, depth + 1)
    return current_ub, best_route


# ============================================================
# 11. RUN GLOBAL B&B (With Progress Bar)
# ============================================================
results = []
sorted_cases = sorted(cases, key=lambda c: c["ub"])
global_ub = sorted_cases[0]["ub"]
global_best_route = sorted_cases[0]["ub_route"]

print("\n================ GLOBAL BRANCH-AND-BOUND ================\n")
for case in tqdm(sorted_cases, desc="Global Optimization"):
    end, initial_lb = case["end"], case["initial_lb"]
    if initial_lb >= global_ub: continue

    root_node = {
        "lb": initial_lb, "matrix": case["initial_matrix"].copy(),
        "rows": list(range(n)), "cols": list(range(n)),
        "forced_edges": [], "forbidden_edges": [], "end": end
    }

    best_ub_for_end, best_route_for_end = explore_node(root_node, global_ub, case["ub_route"], 0)

    if best_route_for_end is not None and best_ub_for_end < global_ub:
        global_ub, global_best_route = best_ub_for_end, best_route_for_end

    if best_route_for_end is not None:
        results.append({"end": end, "best_cost": best_ub_for_end, "best_route": best_route_for_end})

# ============================================================
# 12. FINAL RESULTS
# ============================================================
if results:
    best_overall = min(results, key=lambda r: r["best_cost"])
    print("\n================ GLOBAL BEST ROUTE ================\n")
    print(f"Best ending: {best_overall['end']}")
    print(f"Best cost: {best_overall['best_cost']:.2f} minutes")
    print("Best route:", " → ".join(map(str, best_overall["best_route"])))    
    # Print real addresses with proper RTL support for Hebrew
    real_addresses = " → ".join([get_display(address_labels[idx]) for idx in best_overall["best_route"]])
    print("Real addresses:", real_addresses)