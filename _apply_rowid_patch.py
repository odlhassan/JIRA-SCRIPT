"""
One-shot script to extract row-id refactored functions from the backup
and patch them into the HEAD-restored report_server.py.
"""
import re
import sys

HEAD_FILE = "report_server.py"
BACKUP_FILE = "report_server.py.rowid-backup"

# Functions to replace (name -> expected to exist in both files)
FUNCTIONS_TO_REPLACE = [
    "_init_epics_management_db",
    "_upsert_epics_management_story_sync_rows",
    "_load_epics_management_story_sync_rows",
    "_normalize_epics_management_plan",
    "_normalize_epics_management_payload",
    "_save_epics_management_row",
    "_update_epics_management_row",
    "_delete_epics_management_row",
    "_load_epics_management_rows",
    "_build_epics_management_snapshot_dict",
    "_seal_epics_management_epics",
    "_rebudget_epics_management_epic",
    "_load_epics_management_sealed_dates",
    "_load_epics_management_sealed_dates_for_epic",
    "_load_epics_management_snapshot_for_epic_date",
    "_delete_epics_management_approved_date",
    "_sync_epic_plan_from_jira",
]

# New functions that only exist in backup (insert BEFORE a known anchor function in HEAD)
# (function_name, insert_before_function_in_HEAD)
NEW_FUNCTIONS = [
    "_normalize_epics_management_row_id",
    "_generate_epics_management_row_id",
    "_resolve_epics_management_row",
    "_compute_epics_management_tk_budgeted_plans",
]


def extract_function(lines, func_name, start_search=0):
    """Extract a top-level function from lines. Returns (start_idx, end_idx, body)."""
    pattern = re.compile(rf'^def {re.escape(func_name)}\(')
    start_idx = None
    for i in range(start_search, len(lines)):
        if pattern.match(lines[i]):
            start_idx = i
            break
    if start_idx is None:
        return None, None, None

    # Find end: next top-level def/class or end of file
    end_idx = len(lines)
    for i in range(start_idx + 1, len(lines)):
        stripped = lines[i]
        if stripped and not stripped[0].isspace() and (stripped.startswith("def ") or stripped.startswith("class ")):
            # Go back to skip trailing blank lines
            end_idx = i
            while end_idx > start_idx + 1 and lines[end_idx - 1].strip() == "":
                end_idx -= 1
            end_idx += 1  # keep one blank line
            break

    return start_idx, end_idx, lines[start_idx:end_idx]


def find_all_functions(lines, func_names):
    """Find all occurrences of each function. Returns dict of name -> list of (start, end, body)."""
    result = {}
    for name in func_names:
        occurrences = []
        search_from = 0
        while True:
            s, e, body = extract_function(lines, name, search_from)
            if s is None:
                break
            occurrences.append((s, e, body))
            search_from = e
        result[name] = occurrences
    return result


def main():
    with open(HEAD_FILE, "r", encoding="utf-8") as f:
        head_lines = f.readlines()
    with open(BACKUP_FILE, "r", encoding="utf-8") as f:
        backup_lines = f.readlines()

    print(f"HEAD: {len(head_lines)} lines")
    print(f"BACKUP: {len(backup_lines)} lines")

    # Find all functions in HEAD
    head_funcs = find_all_functions(head_lines, FUNCTIONS_TO_REPLACE)
    # Find all functions in backup (use the LAST occurrence as canonical, since early ones are corrupted duplicates)
    all_backup_names = FUNCTIONS_TO_REPLACE + NEW_FUNCTIONS
    backup_funcs = find_all_functions(backup_lines, all_backup_names)

    # Report findings
    print("\n=== Functions to REPLACE ===")
    for name in FUNCTIONS_TO_REPLACE:
        h_occs = head_funcs.get(name, [])
        b_occs = backup_funcs.get(name, [])
        h_info = ", ".join(f"L{s+1}-{e}" for s, e, _ in h_occs) if h_occs else "NOT FOUND"
        b_info = ", ".join(f"L{s+1}-{e}" for s, e, _ in b_occs) if b_occs else "NOT FOUND"
        print(f"  {name}: HEAD=[{h_info}] BACKUP=[{b_info}]")
        if not h_occs:
            print(f"    WARNING: Not found in HEAD!")
        if not b_occs:
            print(f"    WARNING: Not found in BACKUP!")

    print("\n=== NEW functions (backup only) ===")
    for name in NEW_FUNCTIONS:
        b_occs = backup_funcs.get(name, [])
        b_info = ", ".join(f"L{s+1}-{e}" for s, e, _ in b_occs) if b_occs else "NOT FOUND"
        print(f"  {name}: BACKUP=[{b_info}]")

    # Now do the patching
    # Strategy: work backwards through HEAD to avoid index shifts
    replacements = []  # (head_start, head_end, new_body_lines)

    for name in FUNCTIONS_TO_REPLACE:
        h_occs = head_funcs.get(name, [])
        b_occs = backup_funcs.get(name, [])
        if not h_occs or not b_occs:
            print(f"\nSKIPPING {name} - missing in one file")
            continue
        # Use the single HEAD occurrence (should be exactly 1)
        if len(h_occs) != 1:
            print(f"\nWARNING: {name} has {len(h_occs)} occurrences in HEAD, using first")
        h_start, h_end, _ = h_occs[0]
        # Use the LAST backup occurrence (canonical, not the corrupted early duplicate)
        _, _, b_body = b_occs[-1]
        replacements.append((h_start, h_end, b_body, name))

    # For new functions, insert them before _normalize_epics_management_payload in HEAD
    # (this is a good anchor point - the new helpers should come before it)
    anchor_name = "_normalize_epics_management_payload"
    anchor_occs = head_funcs.get(anchor_name, [])
    if anchor_occs:
        anchor_start = anchor_occs[0][0]
        new_func_lines = []
        for name in NEW_FUNCTIONS:
            b_occs = backup_funcs.get(name, [])
            if b_occs:
                _, _, b_body = b_occs[-1]
                new_func_lines.extend(b_body)
                if not new_func_lines[-1].strip() == "":
                    new_func_lines.append("\n")
                new_func_lines.append("\n")
                print(f"\nWill INSERT new function: {name} ({len(b_body)} lines)")
            else:
                print(f"\nWARNING: New function {name} not found in backup!")
        if new_func_lines:
            replacements.append((anchor_start, anchor_start, new_func_lines, "NEW_FUNCTIONS_BLOCK"))

    # Sort replacements by start position, descending (to apply from bottom up)
    replacements.sort(key=lambda x: x[0], reverse=True)

    print(f"\n=== Applying {len(replacements)} patches (bottom-up) ===")
    for h_start, h_end, new_body, name in replacements:
        old_count = h_end - h_start
        new_count = len(new_body)
        print(f"  {name}: HEAD L{h_start+1}-{h_end} ({old_count} lines) -> {new_count} lines")
        head_lines[h_start:h_end] = new_body

    # Write result
    output_file = HEAD_FILE
    with open(output_file, "w", encoding="utf-8") as f:
        f.writelines(head_lines)

    print(f"\nDone! Wrote {len(head_lines)} lines to {output_file}")


if __name__ == "__main__":
    main()
