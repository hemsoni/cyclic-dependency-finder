"""
Cyclic Dependency Finder
========================
Reads an Excel file (all sheets, all columns) and detects circular dependencies.

Usage:
    python find_cycles.py <excel_file> --source <column> --target <column>

Example:
    python find_cycles.py dependencies.xlsx --source "Task" --target "Depends On"
"""

import argparse
import sys
from collections import defaultdict

import openpyxl


# ---------------------------------------------------------------------------
# Graph helpers
# ---------------------------------------------------------------------------

def build_graph(workbook, source_col, target_col, separator):
    """
    Read every sheet in the workbook and build a directed graph.

    Each row creates edges:  source_value  -->  target_value

    If a target cell contains multiple values separated by `separator`,
    each value becomes a separate edge.

    Returns
    -------
    graph : dict[str, set[str]]
        Adjacency list  {node: {neighbours ...}}
    edge_origins : dict[tuple, list[str]]
        Maps (source, target) to the sheet(s) where that edge was found.
    """
    graph = defaultdict(set)
    edge_origins = defaultdict(list)

    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]

        # Read the header row to find column positions
        headers = [
            str(cell.value).strip() if cell.value is not None else ""
            for cell in next(sheet.iter_rows(min_row=1, max_row=1))
        ]

        # Case-insensitive lookup
        headers_lower = [h.lower() for h in headers]
        src_lower = source_col.lower()
        tgt_lower = target_col.lower()

        if src_lower not in headers_lower or tgt_lower not in headers_lower:
            # This sheet doesn't have the required columns -- skip it
            continue

        src_idx = headers_lower.index(src_lower)
        tgt_idx = headers_lower.index(tgt_lower)

        for row in sheet.iter_rows(min_row=2, values_only=True):
            if row[src_idx] is None or row[tgt_idx] is None:
                continue

            source_val = str(row[src_idx]).strip()
            if not source_val:
                continue

            raw_targets = str(row[tgt_idx]).strip()
            if not raw_targets:
                continue

            targets = [t.strip() for t in raw_targets.split(separator) if t.strip()]

            for target_val in targets:
                graph[source_val].add(target_val)
                edge_origins[(source_val, target_val)].append(sheet_name)
                # Make sure target nodes exist in the graph even if they have
                # no outgoing edges, so the cycle detection visits them.
                if target_val not in graph:
                    graph[target_val] = set()

    return graph, edge_origins


def find_all_cycles(graph):
    """
    Find ALL elementary cycles in a directed graph using Johnson's approach
    (simplified DFS-based).

    Returns a list of cycles, where each cycle is a list of node names.
    Example: ["A", "B", "C"] means A -> B -> C -> A
    """
    WHITE, GRAY, BLACK = 0, 1, 2
    cycles = []

    def dfs(node, color, path, path_set):
        color[node] = GRAY
        path.append(node)
        path_set.add(node)

        for neighbour in graph.get(node, []):
            if color[neighbour] == GRAY and neighbour in path_set:
                # Found a cycle -- extract it
                cycle_start = path.index(neighbour)
                cycle = path[cycle_start:]
                cycles.append(cycle)
            elif color[neighbour] == WHITE:
                dfs(neighbour, color, path, path_set)

        path.pop()
        path_set.discard(node)
        color[node] = BLACK

    color = {node: WHITE for node in graph}
    for node in graph:
        if color[node] == WHITE:
            dfs(node, color, [], set())

    return cycles


def deduplicate_cycles(cycles):
    """Remove duplicate cycles that are just rotations of each other."""
    seen = set()
    unique = []
    for cycle in cycles:
        # Normalize: rotate so the smallest element is first
        min_idx = cycle.index(min(cycle))
        rotated = tuple(cycle[min_idx:] + cycle[:min_idx])
        if rotated not in seen:
            seen.add(rotated)
            unique.append(cycle)
    return unique


# ---------------------------------------------------------------------------
# Display helpers
# ---------------------------------------------------------------------------

def print_cycles(cycles, edge_origins):
    """Pretty-print each cycle with sheet origin information."""
    if not cycles:
        print()
        print("=== RESULT: No cyclic dependencies found! ===")
        print()
        return

    print()
    print(f"=== RESULT: Found {len(cycles)} cyclic dependency(ies)! ===")
    print()

    for i, cycle in enumerate(cycles, 1):
        chain = " -> ".join(cycle) + " -> " + cycle[0]
        print(f"  Cycle {i}:  {chain}")

        # Show which sheet each edge came from
        for j in range(len(cycle)):
            src = cycle[j]
            tgt = cycle[(j + 1) % len(cycle)]
            sheets = edge_origins.get((src, tgt), ["unknown"])
            print(f"            {src} -> {tgt}  (found in sheet: {', '.join(sheets)})")
        print()


def print_summary(workbook, graph, source_col, target_col):
    """Print a summary of what was read."""
    print()
    print("=" * 60)
    print("  Cyclic Dependency Finder")
    print("=" * 60)
    print(f"  Sheets scanned  : {len(workbook.sheetnames)}")
    print(f"  Source column    : {source_col}")
    print(f"  Target column    : {target_col}")
    print(f"  Unique nodes     : {len(graph)}")
    total_edges = sum(len(v) for v in graph.values())
    print(f"  Total edges      : {total_edges}")
    print("=" * 60)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Find cyclic (circular) dependencies in an Excel file.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python find_cycles.py tasks.xlsx --source "Task" --target "Depends On"
  python find_cycles.py modules.xlsx --source "Module" --target "Requires" --separator ";"
        """,
    )
    parser.add_argument("excel_file", help="Path to the Excel file (.xlsx)")
    parser.add_argument(
        "--source",
        required=True,
        help='Column name that contains the item name (e.g. "Task")',
    )
    parser.add_argument(
        "--target",
        required=True,
        help='Column name that contains the dependency (e.g. "Depends On")',
    )
    parser.add_argument(
        "--separator",
        default=",",
        help='Separator when one cell lists multiple dependencies (default: ",")',
    )

    args = parser.parse_args()

    # --- Load workbook ---
    try:
        wb = openpyxl.load_workbook(args.excel_file, read_only=True, data_only=True)
    except FileNotFoundError:
        print(f"ERROR: File not found: {args.excel_file}")
        sys.exit(1)
    except Exception as e:
        print(f"ERROR: Could not open Excel file: {e}")
        sys.exit(1)

    # --- Build graph ---
    graph, edge_origins = build_graph(wb, args.source, args.target, args.separator)

    if not graph:
        print(f"ERROR: No data found. Check that columns '{args.source}' and "
              f"'{args.target}' exist in at least one sheet.")
        sys.exit(1)

    print_summary(wb, graph, args.source, args.target)

    # --- Detect cycles ---
    raw_cycles = find_all_cycles(graph)
    cycles = deduplicate_cycles(raw_cycles)

    print_cycles(cycles, edge_origins)

    wb.close()

    # Exit code: 1 if cycles found, 0 if clean
    sys.exit(1 if cycles else 0)


if __name__ == "__main__":
    main()
