# Duplicate File Finder

A lightweight Python desktop app for finding duplicate files in a selected folder.

It scans a directory tree, groups files that share the same case-insensitive filename and exact byte size, and shows the results in a clear table. The app can also export a detailed Excel report for review or sharing.

## Features

- Scan a single folder or include all subfolders
- Detect duplicates using the matching rule: case-insensitive file name + exact file size
- Display scan status with duplicate and unique entries
- Group duplicate files into numbered duplicate sets
- Export results to an Excel workbook (.xlsx)
- Include a summary sheet with scan metadata and counts

## Requirements

- Python 3
- Tkinter (included with most Python installs on Windows/macOS)
- Optional for Excel export: `openpyxl`

Install the Excel dependency if needed:

```bash
python -m pip install openpyxl
```

## Running the app

From the project directory, run:

```bash
python duplicate_file_finder.py
```

## How to use

1. Click "Browse…" next to "FOLDER TO SCAN" and select the directory to inspect.
2. Optionally choose a destination folder for the generated Excel report.
3. Tick "Include subfolders" if you want the scan to recurse into nested folders.
4. Click "Scan for Duplicates".
5. Review the table of files and duplicate groups.
6. Click "Export Excel (.xlsx)" to save the current results.

## Notes

- The app intentionally matches files by name and size rather than content hash, so it highlights near-duplicate candidates based on the same filename and file length.
- This is useful for quickly spotting repeated files such as backups, copies, or renamed duplicates.
- Exporting requires an Excel-compatible destination folder and the `openpyxl` package.
