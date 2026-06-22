"""
GENERAL PURPOSE RESTRUCTURE SCRIPT
------------------------------------
Works for any customer folder in \\tri-mil-dev01\Jobs\5_JOB FOLDERS\

Set DRY_RUN = True to preview without copying anything.
Set DRY_RUN = False to actually copy files (source is preserved).

Usage:
  1. Set CUSTOMER_FOLDER to the customer folder name
  2. Set DRY_RUN to True or False
  3. Run: python restructure.py
"""

import os
import csv
import re
import shutil
import openpyxl
import xlrd

# ── CONFIGURATION ──────────────────────────────────────────────────────────────

CUSTOMER_FOLDER = "Colgate"  # <-- Change this to the customer folder name
DRY_RUN         = True          # <-- Set to False to actually move files

SOURCE_ROOT  = r"\\tri-mil-dev01\Jobs\5_JOB FOLDERS"
OUTPUT_ROOT  = r"\\tri-mil-dev01\Working Directory D"
DESKTOP      = os.path.join(os.path.expanduser("~"), "OneDrive - Tripack, LLC", "Desktop")

CUSTOMER_PATH = os.path.join(SOURCE_ROOT, CUSTOMER_FOLDER)

MODE     = "DRYRUN" if DRY_RUN else "COPY"
LOG_NAME = CUSTOMER_FOLDER.replace(" ", "_") + "_copylog_" + MODE + ".csv"
LOG_PATH = os.path.join(DESKTOP, LOG_NAME)

TEMPLATE_SUBFOLDERS = [
    "1_SERVICE",
    "2_CURP",
    "Accounting",
    "Art Files_Dielines",
    "Calibration_Chart",
    "Customer_Documentation_Package",
    "Electrical",
    "Manuals",
    "Mechanical",
    "Misc",
    "Punchlists & Customer Sign-Offs",
]

CALIBRATION_FOLDER_VARIANTS = [
    "manual & calibration",
    "manual, spare & calibration",
    "calibration",
]

# ── HELPERS ────────────────────────────────────────────────────────────────────

def extract_sn_from_name(name):
    """Extract SN digits from a name using SN or S prefix pattern."""
    # Try SN##### first (more specific)
    match = re.search(r'SN(\d+)', name, re.IGNORECASE)
    if match:
        return match.group(1)
    # Try S##### (e.g. S3325_Barbasol) — must be at start or after underscore/space
    match = re.search(r'(?:^|[_\s])S(\d{4,})', name, re.IGNORECASE)
    if match:
        return match.group(1)
    return None


def extract_all_long_numbers(text):
    """Extract all standalone numbers with 4+ digits from text."""
    return re.findall(r'\b(\d{4,})\b', text)


def extract_sn_from_excel(filepath, sn_map=None):
    """
    Open an Excel file and search for SN patterns.
    Also searches for any long number that matches a known SN.
    Returns SN digits string or None.
    """
    try:
        ext = os.path.splitext(filepath)[1].lower()
        if ext == ".xlsx":
            wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
            for sheet in wb.worksheets:
                for row in sheet.iter_rows(values_only=True):
                    for cell in row:
                        if cell is None:
                            continue
                        val = str(cell)
                        # Check for explicit SN pattern
                        match = re.search(r'SN(\d+)', val, re.IGNORECASE)
                        if match:
                            wb.close()
                            return match.group(1)
                        # Check if any long number matches a known SN
                        if sn_map:
                            for num in extract_all_long_numbers(val):
                                matched = match_sn_to_folder(num, sn_map, silent=True)
                                if matched:
                                    sn_key = [k for k, v in sn_map.items() if v == matched][0]
                                    wb.close()
                                    return sn_key
            wb.close()
        elif ext in (".xls",):
            wb = xlrd.open_workbook(filepath)
            for sheet in wb.sheets():
                for row in range(sheet.nrows):
                    for col in range(sheet.ncols):
                        val = str(sheet.cell_value(row, col))
                        # Check for explicit SN pattern
                        match = re.search(r'SN(\d+)', val, re.IGNORECASE)
                        if match:
                            return match.group(1)
                        # Check if any long number matches a known SN
                        if sn_map:
                            for num in extract_all_long_numbers(val):
                                matched = match_sn_to_folder(num, sn_map, silent=True)
                                if matched:
                                    return [k for k, v in sn_map.items() if v == matched][0]
    except Exception as e:
        if "not a zip file" not in str(e).lower():
            print("    [WARNING] Could not read " + os.path.basename(filepath) + ": " + str(e))
    return None


def match_sn_to_folder(sn_num, sn_map, silent=False):
    """
    Match a found SN number to an SN folder.
    Checks: exact match, prefix match, suffix match, substring match.
    Returns the folder name or None.
    """
    if not sn_num:
        return None

    # Exact match
    if sn_num in sn_map:
        return sn_map[sn_num]

    # Prefix, suffix, or substring match against known SN keys
    for key, folder in sn_map.items():
        if key.startswith(sn_num) or sn_num.startswith(key):
            if not silent:
                print("  [Partial SN match] " + sn_num + " -> " + key)
            return folder
        # Check if sn_num appears anywhere in the SN folder name
        if sn_num in folder:
            if not silent:
                print("  [Substring SN match] " + sn_num + " found in " + folder)
            return folder

    return None


def find_sn_for_folder(folder_path, label="", sn_map=None):
    """
    Universal SN finder for any folder (CURP, SR, etc).
    Strategy:
      1. Folder name (SN prefix pattern)
      2. File names inside folder (SN prefix pattern)
      3. File names inside folder (long number matching known SNs)
      4. Excel file contents (SN prefix pattern + long number matching)
    Returns SN digits string or None.
    """
    folder_name = os.path.basename(folder_path)

    # Step 1: folder name — SN prefix
    sn = extract_sn_from_name(folder_name)
    if sn:
        print("  [SN found in folder name] " + sn)
        return sn

    # Step 2: file names inside — SN prefix
    try:
        for item in os.listdir(folder_path):
            sn = extract_sn_from_name(item)
            if sn:
                print("  [SN found in file name '" + item + "'] " + sn)
                return sn
    except Exception:
        pass

    # Step 3: file names inside — long number matching known SNs
    if sn_map:
        try:
            for item in os.listdir(folder_path):
                for num in extract_all_long_numbers(item):
                    matched = match_sn_to_folder(num, sn_map, silent=True)
                    if matched:
                        sn_key = [k for k, v in sn_map.items() if v == matched][0]
                        print("  [SN found via number in file name '" + item + "'] " + sn_key)
                        return sn_key
        except Exception:
            pass

    # Step 4: open Excel files
    print("  [No SN in names - opening Excel files in " + folder_name + "]")
    for root, dirs, files in os.walk(folder_path):
        for fname in files:
            ext = os.path.splitext(fname)[1].lower()
            if ext in (".xlsx", ".xls") and not fname.startswith("~$"):
                fpath = os.path.join(root, fname)
                print("    Checking: " + fname)
                sn = extract_sn_from_excel(fpath, sn_map)
                if sn:
                    print("    [SN found in file contents] " + sn)
                    return sn

    print("  [SN could not be determined for " + (label or folder_name) + "]")
    return None


def is_calibration_folder(folder_name):
    return folder_name.lower() in CALIBRATION_FOLDER_VARIANTS


def count_files(folder):
    total = 0
    for root, dirs, files in os.walk(folder):
        total += len(files)
    return total


# ── PROCESS CURP FOLDER ────────────────────────────────────────────────────────

def process_curp_folder(curp_path, sn_map, log, errors):
    curp_folder = os.path.basename(curp_path)
    print("\n  CURP: " + curp_folder)
    sn_num = find_sn_for_folder(curp_path, "CURP", sn_map)
    target_sn = match_sn_to_folder(sn_num, sn_map)

    if target_sn:
        new_curp_path = os.path.join(OUTPUT_ROOT, target_sn, "2_CURP", curp_folder)
        print("  -> Routes to " + target_sn + "/2_CURP")
        if DRY_RUN:
            log(curp_folder, "folder", curp_path, new_curp_path, "moved")
            for root, dirs, files in os.walk(curp_path):
                rel = os.path.relpath(root, curp_path)
                rel = "" if rel == "." else rel
                for fname in files:
                    fpath = os.path.join(root, fname)
                    new_path = os.path.join(new_curp_path, rel, fname) if rel else os.path.join(new_curp_path, fname)
                    log(fname, "file", fpath, new_path, "moved")
        else:
            try:
                dest = os.path.join(OUTPUT_ROOT, target_sn, "2_CURP", curp_folder)
                os.makedirs(os.path.dirname(dest), exist_ok=True)
                shutil.copytree(curp_path, dest, dirs_exist_ok=True)
                log(curp_folder, "folder", curp_path, new_curp_path, "copied")
            except Exception as e:
                err = "ERROR moving CURP " + curp_folder + ": " + str(e)
                print("  " + err)
                errors.append(err)
                log(curp_folder, "folder", curp_path, new_curp_path, "ERROR: " + str(e))
    elif sn_num:
        print("  -> SN" + sn_num + " found but no matching SN folder")
        log(curp_folder, "folder", curp_path, None,
            "unchanged (SN" + sn_num + " identified but no matching SN folder)")
    else:
        log(curp_folder, "folder", curp_path, None, "unchanged (SN could not be determined)")


# ── PROCESS SR FOLDER ──────────────────────────────────────────────────────────

def process_sr_folder(sr_path, sn_map, log, errors):
    sr_folder = os.path.basename(sr_path)
    print("\n  SR: " + sr_folder)
    sn_num = find_sn_for_folder(sr_path, "SR", sn_map)
    target_sn = match_sn_to_folder(sn_num, sn_map)

    if target_sn:
        new_sr_path = os.path.join(OUTPUT_ROOT, target_sn, "1_SERVICE", sr_folder)
        print("  -> Routes to " + target_sn + "/1_SERVICE")
        if DRY_RUN:
            log(sr_folder, "folder", sr_path, new_sr_path, "moved")
            for root, dirs, files in os.walk(sr_path):
                rel = os.path.relpath(root, sr_path)
                rel = "" if rel == "." else rel
                for fname in files:
                    fpath = os.path.join(root, fname)
                    new_path = os.path.join(new_sr_path, rel, fname) if rel else os.path.join(new_sr_path, fname)
                    log(fname, "file", fpath, new_path, "moved")
        else:
            try:
                dest = os.path.join(OUTPUT_ROOT, target_sn, "1_SERVICE", sr_folder)
                os.makedirs(os.path.dirname(dest), exist_ok=True)
                shutil.copytree(sr_path, dest, dirs_exist_ok=True)
                log(sr_folder, "folder", sr_path, new_sr_path, "copied")
            except Exception as e:
                err = "ERROR moving SR " + sr_folder + ": " + str(e)
                print("  " + err)
                errors.append(err)
                log(sr_folder, "folder", sr_path, new_sr_path, "ERROR: " + str(e))
    elif sn_num:
        print("  -> SN" + sn_num + " found but no matching SN folder")
        log(sr_folder, "folder", sr_path, None,
            "unchanged (SN" + sn_num + " identified but no matching SN folder)")
    else:
        log(sr_folder, "folder", sr_path, None, "unchanged (SN could not be determined)")


# ── PROCESS SN FOLDER ──────────────────────────────────────────────────────────

def process_sn_folder(sn_folder, log, errors):
    sn_path = os.path.join(CUSTOMER_PATH, sn_folder)
    new_sn_path = os.path.join(OUTPUT_ROOT, sn_folder)
    print("\n-- Processing SN folder: " + sn_folder)

    # Reference files for duplicate detection
    reference_files = set()
    for root, dirs, files in os.walk(sn_path):
        if "THUMB DRIVE_TEMPLATE" in root:
            continue
        for f in files:
            reference_files.add(f)

    # Handle THUMB DRIVE_TEMPLATE
    thumb_path = os.path.join(sn_path, "THUMB DRIVE_TEMPLATE")
    if os.path.isdir(thumb_path):
        print("  Handling THUMB DRIVE_TEMPLATE...")
        for root, dirs, files in os.walk(thumb_path):
            for fname in files:
                fpath = os.path.join(root, fname)
                if fname in reference_files:
                    # In copy mode, don't delete duplicates — just log
                    log(fname, "file", fpath, None, "would-be-deleted (duplicate in THUMB DRIVE_TEMPLATE — skipped in copy mode)")
                else:
                    misc_dest = os.path.join(sn_path, "Misc", fname)
                    misc_dest_log = os.path.join(new_sn_path, "Misc", fname)
                    if not DRY_RUN:
                        try:
                            os.makedirs(os.path.dirname(misc_dest), exist_ok=True)
                            os.makedirs(os.path.dirname(misc_dest), exist_ok=True); shutil.copy2(fpath, misc_dest)
                        except Exception as e:
                            errors.append("ERROR moving " + fpath + ": " + str(e))
                    log(fname, "file", fpath, misc_dest_log,
                        "moved to Misc (unique file from THUMB DRIVE_TEMPLATE)")
        # In copy mode, don't delete THUMB DRIVE folder
        log("THUMB DRIVE_TEMPLATE", "folder", thumb_path, None, "processed (preserved in copy mode)")

    # Handle Manual/Calibration folder variants
    for item in os.listdir(sn_path):
        item_path = os.path.join(sn_path, item)
        if os.path.isdir(item_path) and is_calibration_folder(item):
            print("  Splitting '" + item + "' folder...")
            for root, dirs, files in os.walk(item_path):
                for fname in files:
                    fpath = os.path.join(root, fname)
                    parent_lower = os.path.basename(root).lower()
                    fname_lower = fname.lower()
                    if "calibration" in parent_lower or fname_lower.startswith("cc"):
                        dest_sub = "Calibration_Chart"
                    elif "spare" in parent_lower:
                        dest_sub = "Misc"
                    else:
                        dest_sub = "Manuals"
                    dest = os.path.join(sn_path, dest_sub, fname)
                    dest_log = os.path.join(new_sn_path, dest_sub, fname)
                    if not DRY_RUN:
                        try:
                            os.makedirs(os.path.dirname(dest), exist_ok=True)
                            os.makedirs(os.path.dirname(dest), exist_ok=True); shutil.copy2(fpath, dest)
                        except Exception as e:
                            errors.append("ERROR moving " + fpath + ": " + str(e))
                    log(fname, "file", fpath, dest_log,
                        "moved (" + item + " -> " + dest_sub + ")")
            # In copy mode, don't delete the original folder
            log(item, "folder", item_path, None, "split into template subfolders (original preserved in copy mode)")

    # Create missing template subfolders
    existing_lower = [e.lower() for e in os.listdir(sn_path)]
    for subfolder in TEMPLATE_SUBFOLDERS:
        if subfolder.lower() not in existing_lower:
            subfolder_path = os.path.join(sn_path, subfolder)
            if not DRY_RUN:
                try:
                    os.makedirs(subfolder_path, exist_ok=True)
                except Exception as e:
                    errors.append("ERROR creating " + subfolder + ": " + str(e))
            log(subfolder, "folder", "N/A (did not exist)",
                os.path.join(new_sn_path, subfolder),
                "created (missing template subfolder)")

    # Log all files in dry run
    if DRY_RUN:
        for root, dirs, files in os.walk(sn_path):
            rel = os.path.relpath(root, sn_path)
            rel = "" if rel == "." else rel
            for fname in files:
                fpath = os.path.join(root, fname)
                new_path = os.path.join(new_sn_path, rel, fname) if rel else os.path.join(new_sn_path, fname)
                log(fname, "file", fpath, new_path, "moved")

    # Move SN folder to output
    if not DRY_RUN:
        try:
            shutil.copytree(sn_path, new_sn_path, dirs_exist_ok=True)
            print("  Copied to output: " + new_sn_path)
        except Exception as e:
            err = "ERROR moving SN folder " + sn_folder + ": " + str(e)
            print("  " + err)
            errors.append(err)
            log(sn_folder, "folder", sn_path, new_sn_path, "ERROR copying: " + str(e))
            return

    log(sn_folder, "folder", sn_path, new_sn_path, "copied")


# ── MAIN ───────────────────────────────────────────────────────────────────────

def main():
    log_entries = []
    errors = []

    def log(name, item_type, original, new_path, result):
        log_entries.append({
            "Name": name,
            "Type": item_type,
            "Original Location Path": original,
            "New Location Path": new_path if new_path else "N/A",
            "Result": result,
        })

    print("=" * 60)
    print(MODE + " - " + CUSTOMER_FOLDER + " Restructure")
    if DRY_RUN:
        print("No files will be moved.")
    else:
        print("Files will be COPIED. Source (E drive) will NOT be modified.")
    print("=" * 60)

    if not os.path.isdir(CUSTOMER_PATH):
        print("ERROR: Customer folder not found: " + CUSTOMER_PATH)
        return

    print("\nCounting source files...")
    source_count = count_files(CUSTOMER_PATH)
    print("Source file count: " + str(source_count))

    # ── STEP 1: Classify top-level items ──
    top_level = os.listdir(CUSTOMER_PATH)
    sn_folders, curp_folders, sr_folders = [], [], []
    curp_container_path = None
    sr_container_path = None
    other_items = []

    for item in top_level:
        item_path = os.path.join(CUSTOMER_PATH, item)
        item_lower = item.lower().strip()
        if re.match(r'^SN|^S[0-9]', item, re.IGNORECASE) and os.path.isdir(item_path):
            sn_folders.append(item)
        elif re.match(r'^CURP', item, re.IGNORECASE) and os.path.isdir(item_path):
            try:
                sub_items = os.listdir(item_path)
                has_curp_children = any(re.match(r'^CURP', s, re.IGNORECASE) for s in sub_items)
                if has_curp_children or item_lower in ("curp", "curps"):
                    curp_container_path = item_path
                else:
                    curp_folders.append(item)
            except Exception:
                curp_folders.append(item)
        elif re.match(r'^SR', item, re.IGNORECASE) and os.path.isdir(item_path):
            sr_folders.append(item)
        elif item_lower in ("service", "1_service", "service reports") and os.path.isdir(item_path):
            sr_container_path = item_path
        else:
            other_items.append(item)

    if curp_container_path:
        print("\n  Found CURP container: " + os.path.basename(curp_container_path))
        for item in os.listdir(curp_container_path):
            sub_path = os.path.join(curp_container_path, item)
            if re.match(r'^CURP', item, re.IGNORECASE) and os.path.isdir(sub_path):
                curp_folders.append(os.path.join(os.path.basename(curp_container_path), item))

    if sr_container_path:
        print("  Found SERVICE container: " + os.path.basename(sr_container_path))
        for item in os.listdir(sr_container_path):
            sub_path = os.path.join(sr_container_path, item)
            if re.match(r'^SR', item, re.IGNORECASE) and os.path.isdir(sub_path):
                sr_folders.append(os.path.join(os.path.basename(sr_container_path), item))

    print("\nFound " + str(len(sn_folders)) + " SN folder(s)")
    print("Found " + str(len(curp_folders)) + " CURP folder(s)")
    print("Found " + str(len(sr_folders)) + " SR folder(s)")
    print("Found " + str(len(other_items)) + " other item(s): " + str(other_items))

    # Build SN map
    sn_map = {}
    for sn_folder in sn_folders:
        sn_num = extract_sn_from_name(sn_folder)
        if sn_num:
            sn_map[sn_num] = sn_folder
            print("  Mapped SN" + sn_num + " -> " + sn_folder)
        else:
            print("  [WARNING] Could not extract SN from: " + sn_folder)

    # ── STEP 2: Process CURPs ──
    if curp_folders:
        print("\n-- Processing CURP folders")
    for curp_item in curp_folders:
        curp_path = os.path.join(CUSTOMER_PATH, curp_item)
        process_curp_folder(curp_path, sn_map, log, errors)

    # ── STEP 3: Process SRs ──
    if sr_folders:
        print("\n-- Processing SERVICE folders")
    for sr_item in sr_folders:
        sr_path = os.path.join(CUSTOMER_PATH, sr_item)
        process_sr_folder(sr_path, sn_map, log, errors)

    # ── STEP 4: Process SN folders ──
    for sn_folder in sn_folders:
        process_sn_folder(sn_folder, log, errors)

    # ── STEP 5: Handle remaining loose items ──
    if os.path.isdir(CUSTOMER_PATH):
        for item in os.listdir(CUSTOMER_PATH):
            item_path = os.path.join(CUSTOMER_PATH, item)
            if item in sn_folders:
                continue
            item_lower = item.lower().strip()
            if item_lower in ("curp", "curps", "service", "1_service", "service reports") \
                    or (curp_container_path and item_path == curp_container_path) \
                    or (sr_container_path and item_path == sr_container_path):
                if os.path.isdir(item_path):
                    total_files = sum(len(files) for _, _, files in os.walk(item_path))
                    log(item, "folder", item_path, None,
                        "unchanged (container preserved in copy mode, " + str(total_files) + " files)")
                continue
            item_type = "folder" if os.path.isdir(item_path) else "file"
            log(item, item_type, item_path, None, "unchanged (could not assign to SN)")

    # ── STEP 6: Delete customer folder if empty ──
    if os.path.isdir(CUSTOMER_PATH):
        remaining = os.listdir(CUSTOMER_PATH)
        log(CUSTOMER_FOLDER, "folder", CUSTOMER_PATH, None,
                "unchanged (copy mode - source preserved, " + str(len(remaining)) + " item(s) remain)")

    # ── VERIFICATION ──
    print("\n-- Verification")
    moved     = sum(1 for r in log_entries if r["Result"] == "moved" and r["Type"] == "file")
    deleted   = sum(1 for r in log_entries if "deleted" in r["Result"] and r["Type"] == "file")
    misc      = sum(1 for r in log_entries if "moved to Misc" in r["Result"] and r["Type"] == "file")
    split     = sum(1 for r in log_entries if "->" in r["Result"] and r["Type"] == "file")
    unchanged = sum(1 for r in log_entries if "unchanged" in r["Result"] and r["Type"] == "file")
    errs      = sum(1 for r in log_entries if "ERROR" in r["Result"] and r["Type"] == "file")
    accounted = moved + deleted + misc + split + unchanged + errs

    print("Source file count (before): " + str(source_count))
    print("Files moved:                " + str(moved + misc + split))
    print("Files deleted (duplicates): " + str(deleted))
    print("Files unchanged:            " + str(unchanged))
    print("Files with errors:          " + str(errs))
    print("Total accounted for:        " + str(accounted))

    if accounted >= source_count:
        print("\n[PASS] All files accounted for.")
    else:
        print("\n[WARNING] " + str(source_count - accounted) + " file(s) unaccounted for. Review the log.")

    # ── WRITE LOG ──
    os.makedirs(DESKTOP, exist_ok=True)
    with open(LOG_PATH, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=[
            "Name", "Type", "Original Location Path", "New Location Path", "Result"
        ])
        writer.writeheader()
        writer.writerows(log_entries)

    print("\n" + "=" * 60)
    print(MODE + " COMPLETE - " + CUSTOMER_FOLDER + " (source preserved)")
    print("Total items logged: " + str(len(log_entries)))
    if errors:
        print("Errors: " + str(len(errors)))
        for e in errors:
            print("  " + e)
    print("Log saved to: " + LOG_PATH)
    print("=" * 60)


if __name__ == "__main__":
    main()
