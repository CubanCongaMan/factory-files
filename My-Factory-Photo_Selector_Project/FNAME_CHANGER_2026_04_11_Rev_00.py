import csv
import os
import re
import shutil
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox


SOURCE_HEADERS = ["PHOTO'S CAMERA NAME", "PHOTO'S CAMERAS NAME"]
DATE_HEADER = "DATE"
TARGET_HEADER = "NEW DESCRIPTIVE FILE NAME"
SOURCE_NAME_PATTERN = re.compile(r"^IMG_(\d{4})$", re.IGNORECASE)
WINDOWS_INVALID_CHARS = set('<>:"/\\|?*')


def normalize_header(header_name):
    return str(header_name or "").strip()


def canonical_stem(file_name):
    return Path(file_name).stem.strip().upper()


def has_jpg_extension(file_name):
    return Path(file_name).suffix.lower() == ".jpg"


def validate_source_name(file_name):
    stem = canonical_stem(file_name)
    return bool(SOURCE_NAME_PATTERN.fullmatch(stem))


def validate_new_name(raw_name):
    candidate = str(raw_name or "").strip()
    if not candidate:
        return False, "blank"
    if any(char in WINDOWS_INVALID_CHARS for char in candidate):
        return False, "contains invalid Windows filename characters"
    if candidate.endswith(".") or candidate.endswith(" "):
        return False, "ends with a space or period"
    return True, candidate


def stop_with_warning(message):
    messagebox.showwarning("WARNING!", message)
    raise SystemExit(1)


def read_csv_rows(csv_path):
    with csv_path.open("r", newline="", encoding="utf-8-sig") as handle:
        reader = csv.DictReader(handle)
        if reader.fieldnames is None:
            stop_with_warning("The CSV file is empty or does not contain a header row.")

        normalized_fields = {normalize_header(name): name for name in reader.fieldnames}
        source_key = None
        for candidate in SOURCE_HEADERS:
            if candidate in normalized_fields:
                source_key = normalized_fields[candidate]
                break

        missing = []
        if source_key is None:
            missing.append(" or ".join(SOURCE_HEADERS))
        for header in (DATE_HEADER, TARGET_HEADER):
            if header not in normalized_fields:
                missing.append(header)
        if missing:
            stop_with_warning(
                "The CSV file is missing required column header(s):\n- " + "\n- ".join(missing)
            )

        date_key = normalized_fields[DATE_HEADER]
        target_key = normalized_fields[TARGET_HEADER]

        rows = []
        for row_number, row in enumerate(reader, start=2):
            source_name = str(row.get(source_key, "") or "").strip()
            date_value = str(row.get(date_key, "") or "").strip()
            target_name = str(row.get(target_key, "") or "").strip()
            if not source_name and not date_value and not target_name:
                continue
            rows.append(
                {
                    "row_number": row_number,
                    "source_name": source_name,
                    "date_value": date_value,
                    "target_name": target_name,
                }
            )
        return rows


def collect_input_images(input_folder):
    files = [item for item in input_folder.iterdir() if item.is_file()]
    if not files:
        stop_with_warning("The input folder does not contain any files.")

    non_jpg_files = [item.name for item in files if not has_jpg_extension(item.name)]
    if non_jpg_files:
        details = "\n".join(
            f"- File named '{Path(name).stem}' has a different extension." for name in sorted(non_jpg_files)
        )
        stop_with_warning(
            "All files in the input folder must have a .jpg extension.\n\n" + details
        )

    invalid_pattern_files = [item.name for item in files if not validate_source_name(item.name)]
    if invalid_pattern_files:
        details = "\n".join(f"- {name}" for name in sorted(invalid_pattern_files))
        stop_with_warning(
            "All input JPG filenames must match the IMG_####.jpg pattern.\n\n"
            "Files with unexpected names:\n" + details
        )

    return sorted(files, key=lambda item: int(SOURCE_NAME_PATTERN.fullmatch(canonical_stem(item.name)).group(1)))


def validate_csv_against_inputs(csv_rows, input_images):
    if len(csv_rows) != len(input_images):
        stop_with_warning(
            "File-count discrepancy detected.\n\n"
            f"Input folder JPG files: {len(input_images)}\n"
            f"CSV rows with new names: {len(csv_rows)}\n\n"
            "Please review the input folder and the CSV file before trying again."
        )

    input_stems = [canonical_stem(item.name) for item in input_images]
    csv_stems = []
    invalid_csv_sources = []
    invalid_targets = []
    duplicate_targets = {}
    seen_targets = {}

    for row in csv_rows:
        source_name = row["source_name"]
        target_name = row["target_name"]
        source_stem = canonical_stem(source_name)
        csv_stems.append(source_stem)

        if not validate_source_name(source_name):
            invalid_csv_sources.append(
                f"- Row {row['row_number']}: source name '{source_name}' is not in IMG_#### format"
            )

        is_valid_target, normalized_target = validate_new_name(target_name)
        if not is_valid_target:
            invalid_targets.append(
                f"- Row {row['row_number']}: new descriptive name '{target_name}' {normalized_target}"
            )
            continue

        target_key = normalized_target.casefold()
        if target_key in seen_targets:
            duplicate_targets.setdefault(target_key, []).append(row["row_number"])
        else:
            seen_targets[target_key] = row["row_number"]

    if invalid_csv_sources:
        stop_with_warning(
            "The CSV contains invalid original filenames in column 1.\n\n" + "\n".join(invalid_csv_sources)
        )

    if invalid_targets:
        stop_with_warning(
            "The CSV contains invalid new descriptive filenames in column 3.\n\n" + "\n".join(invalid_targets)
        )

    if duplicate_targets:
        details = []
        for key, extra_rows in sorted(duplicate_targets.items()):
            first_row = seen_targets[key]
            all_rows = ", ".join(str(row_num) for row_num in [first_row] + extra_rows)
            details.append(f"- '{key}' appears more than once in rows: {all_rows}")
        stop_with_warning(
            "The CSV contains duplicate new descriptive filenames.\n\n" + "\n".join(details)
        )

    mismatches = []
    for index, input_stem in enumerate(input_stems):
        csv_stem = csv_stems[index]
        if input_stem != csv_stem:
            mismatches.append(
                f"- Position {index + 1}: input folder has '{input_stem}.jpg' but CSV column 1 has '{csv_rows[index]['source_name']}'"
            )

    input_set = set(input_stems)
    csv_set = set(csv_stems)
    missing_from_csv = sorted(input_set - csv_set)
    missing_from_folder = sorted(csv_set - input_set)
    if mismatches or missing_from_csv or missing_from_folder:
        parts = []
        if mismatches:
            parts.append("Order mismatch detected:\n" + "\n".join(mismatches))
        if missing_from_csv:
            parts.append(
                "Input JPG names missing from CSV column 1:\n" +
                "\n".join(f"- {name}.jpg" for name in missing_from_csv)
            )
        if missing_from_folder:
            parts.append(
                "CSV column 1 names missing from input folder:\n" +
                "\n".join(f"- {name}" for name in missing_from_folder)
            )
        stop_with_warning("CSV/input filename comparison failed.\n\n" + "\n\n".join(parts))


def build_copy_plan(csv_rows, input_images, output_folder):
    plan = []
    collisions = []
    for row, source_path in zip(csv_rows, input_images):
        target_name = validate_new_name(row["target_name"])[1] + ".jpg"
        destination = output_folder / target_name
        if destination.exists():
            collisions.append(str(destination))
        plan.append((source_path, destination))

    if collisions:
        details = "\n".join(f"- {path}" for path in collisions)
        stop_with_warning(
            "One or more output files already exist. No files were copied.\n\n" + details
        )
    return plan


def execute_copy_plan(copy_plan):
    for source_path, destination in copy_plan:
        shutil.copy2(source_path, destination)


def choose_input_folder(root):
    folder = filedialog.askdirectory(title="Select Input Folder With IMG_####.jpg Files", parent=root)
    if not folder:
        raise SystemExit(0)
    return Path(folder).resolve()


def choose_csv_file(root):
    csv_file = filedialog.askopenfilename(
        title="Select CSV File With New Descriptive Names",
        filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")],
        parent=root,
    )
    if not csv_file:
        raise SystemExit(0)
    return Path(csv_file).resolve()


def choose_output_folder(root):
    folder = filedialog.askdirectory(title="Select Output Folder For Renamed Copies", parent=root)
    if not folder:
        raise SystemExit(0)
    return Path(folder).resolve()


def main():
    root = tk.Tk()
    root.withdraw()

    try:
        input_folder = choose_input_folder(root)
        csv_path = choose_csv_file(root)
        output_folder = choose_output_folder(root)

        if input_folder == output_folder:
            stop_with_warning(
                "The output folder cannot be the same as the input folder.\n"
                "Please choose a different output folder."
            )

        csv_rows = read_csv_rows(csv_path)
        input_images = collect_input_images(input_folder)
        validate_csv_against_inputs(csv_rows, input_images)
        copy_plan = build_copy_plan(csv_rows, input_images, output_folder)
        execute_copy_plan(copy_plan)

        success_message = (
            "Finished successfully.\n\n"
            f"Input JPG files: {len(input_images)}\n"
            f"CSV rows used: {len(csv_rows)}\n"
            f"Output folder: {output_folder}"
        )
        messagebox.showinfo("Success", success_message)
        print(success_message)
    finally:
        root.destroy()


if __name__ == "__main__":
    try:
        main()
    except SystemExit as exc:
        raise
    except Exception as exc:
        tk.Tk().withdraw()
        messagebox.showerror("Unexpected Error", str(exc))
        print(f"Unexpected Error: {exc}", file=sys.stderr)
        raise