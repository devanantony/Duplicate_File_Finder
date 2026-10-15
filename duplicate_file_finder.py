import os
import threading
from collections import defaultdict
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    Workbook = None


APP_TITLE = "Duplicate File Finder"


class DuplicateFileFinder:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1120x740")
        self.root.minsize(900, 600)

        self.bg = "#111827"
        self.panel = "#1f2937"
        self.panel2 = "#263244"
        self.text = "#f3f4f6"
        self.muted = "#9ca3af"
        self.accent = "#60a5fa"
        self.green = "#34d399"
        self.orange = "#fbbf24"
        self.red = "#f87171"

        self.root.configure(bg=self.bg)
        self.folder_var = tk.StringVar()
        self.dest_var = tk.StringVar()
        self.recursive_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="Choose a folder to begin.")
        self.files_scanned = 0
        self.duplicate_count = 0
        self.results = []
        self.scan_running = False

        self._setup_style()
        self._build_ui()

    def _setup_style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(
            "Treeview", background=self.panel, foreground=self.text,
            fieldbackground=self.panel, rowheight=29, borderwidth=0,
            font=("Segoe UI", 10)
        )
        style.configure(
            "Treeview.Heading", background=self.panel2, foreground=self.text,
            font=("Segoe UI", 10, "bold"), relief="flat"
        )
        style.map("Treeview", background=[("selected", "#1d4ed8")],
                  foreground=[("selected", "#ffffff")])
        style.configure(
            "Horizontal.TProgressbar", troughcolor=self.panel2,
            background=self.accent, bordercolor=self.panel2, lightcolor=self.accent,
            darkcolor=self.accent
        )

    def _label(self, parent, text, **kwargs):
        return tk.Label(parent, text=text, bg=kwargs.get("bg", self.panel),
                        fg=kwargs.get("fg", self.text),
                        font=kwargs.get("font", ("Segoe UI", 10)),
                        anchor=kwargs.get("anchor", "w"))

    def _button(self, parent, text, command, primary=False):
        color = self.accent if primary else self.panel2
        fg = "#0b1220" if primary else self.text
        return tk.Button(
            parent, text=text, command=command, bg=color, fg=fg,
            activebackground="#93c5fd" if primary else "#374151",
            activeforeground=fg, relief="flat", bd=0,
            padx=14, pady=9, cursor="hand2",
            font=("Segoe UI", 10, "bold")
        )

    def _build_ui(self):
        header = tk.Frame(self.root, bg=self.bg, padx=24, pady=20)
        header.pack(fill="x")
        self._label(header, "Duplicate File Finder", bg=self.bg,
                    font=("Segoe UI", 22, "bold")).pack(anchor="w")
        self._label(
            header,
            "Find files with matching names and sizes • Scan safely • Export to Excel",
            bg=self.bg, fg=self.muted, font=("Segoe UI", 10)
        ).pack(anchor="w", pady=(4, 0))

        paths = tk.Frame(self.root, bg=self.panel, padx=18, pady=16)
        paths.pack(fill="x", padx=24, pady=(0, 14))

        self._label(paths, "FOLDER TO SCAN", font=("Segoe UI", 9, "bold"),
                    fg=self.muted).grid(row=0, column=0, sticky="w", pady=(0, 7))
        self.folder_entry = tk.Entry(
            paths, textvariable=self.folder_var, bg="#111827", fg=self.text,
            insertbackground=self.text, relief="flat", font=("Segoe UI", 10)
        )
        self.folder_entry.grid(row=1, column=0, sticky="ew", ipady=9, padx=(0, 10))
        self._button(paths, "Browse…", self.browse_folder).grid(row=1, column=1)

        self._label(paths, "REPORT DESTINATION", font=("Segoe UI", 9, "bold"),
                    fg=self.muted).grid(row=2, column=0, sticky="w", pady=(15, 7))
        self.dest_entry = tk.Entry(
            paths, textvariable=self.dest_var, bg="#111827", fg=self.text,
            insertbackground=self.text, relief="flat", font=("Segoe UI", 10)
        )
        self.dest_entry.grid(row=3, column=0, sticky="ew", ipady=9, padx=(0, 10))
        self._button(paths, "Browse…", self.browse_destination).grid(row=3, column=1)

        options = tk.Frame(paths, bg=self.panel)
        options.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        tk.Checkbutton(
            options, text="Include subfolders", variable=self.recursive_var,
            bg=self.panel, fg=self.text, selectcolor=self.panel2,
            activebackground=self.panel, activeforeground=self.text,
            font=("Segoe UI", 10)
        ).pack(side="left")
        self.scan_btn = self._button(options, "▶  Scan for Duplicates",
                                     self.start_scan, primary=True)
        self.scan_btn.pack(side="right")
        paths.columnconfigure(0, weight=1)

        stats = tk.Frame(self.root, bg=self.bg)
        stats.pack(fill="x", padx=24, pady=(0, 14))
        self.total_card, self.total_value = self._stat_card(stats, "FILES SCANNED", "0", self.accent)
        self.dupe_card, self.dupe_value = self._stat_card(stats, "DUPLICATE FILES", "0", self.orange)
        self.group_card, self.group_value = self._stat_card(stats, "DUPLICATE GROUPS", "0", self.green)

        results_panel = tk.Frame(self.root, bg=self.panel, padx=14, pady=12)
        results_panel.pack(fill="both", expand=True, padx=24, pady=(0, 12))
        top = tk.Frame(results_panel, bg=self.panel)
        top.pack(fill="x", pady=(0, 10))
        self._label(top, "Scan Results", font=("Segoe UI", 13, "bold")).pack(side="left")
        self._label(top, "Matching rule: exact filename + exact file size",
                    fg=self.muted).pack(side="right")

        columns = ("status", "name", "size", "group", "path")
        self.tree = ttk.Treeview(results_panel, columns=columns, show="headings")
        headings = {
            "status": "TAG", "name": "FILE NAME", "size": "SIZE",
            "group": "GROUP", "path": "FULL PATH"
        }
        widths = {"status": 120, "name": 220, "size": 110, "group": 90, "path": 500}
        for col in columns:
            self.tree.heading(col, text=headings[col])
            self.tree.column(col, width=widths[col],
                             anchor="w" if col in ("name", "path") else "center",
                             minwidth=70)
        self.tree.tag_configure("duplicate", foreground=self.orange)
        self.tree.tag_configure("unique", foreground=self.muted)
        yscroll = ttk.Scrollbar(results_panel, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(results_panel, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.tree.pack(fill="both", expand=True, side="top")
        xscroll.pack(fill="x", side="bottom")
        yscroll.place(relx=1.0, rely=0.15, relheight=0.78, anchor="ne")

        footer = tk.Frame(self.root, bg=self.bg, padx=24, pady=0)
        footer.pack(fill="x", pady=(0, 16))
        self.progress = ttk.Progressbar(footer, mode="indeterminate",
                                        style="Horizontal.TProgressbar")
        self.progress.pack(fill="x", pady=(0, 8))
        self._label(footer, "", bg=self.bg).pack_forget()
        self.status_label = self._label(footer, self.status_var.get(),
                                        bg=self.bg, fg=self.muted)
        self.status_label.pack(side="left")
        self._button(footer, "Export Excel (.xlsx)", self.export_report,
                     primary=True).pack(side="right")

    def _stat_card(self, parent, title, value, color):
        card = tk.Frame(parent, bg=self.panel, padx=16, pady=12)
        card.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self._label(card, title, fg=self.muted, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        val = self._label(card, value, fg=color, font=("Segoe UI", 21, "bold"))
        val.pack(anchor="w", pady=(4, 0))
        return card, val

    def browse_folder(self):
        folder = filedialog.askdirectory(title="Select folder to scan")
        if folder:
            self.folder_var.set(folder)
            if not self.dest_var.get().strip():
                self.dest_var.set(folder)

    def browse_destination(self):
        folder = filedialog.askdirectory(title="Select report destination folder")
        if folder:
            self.dest_var.set(folder)

    def start_scan(self):
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("Folder required", "Please select a valid folder to scan.")
            return
        if self.scan_running:
            return
        self.scan_running = True
        self.scan_btn.configure(state="disabled")
        self.progress.start(12)
        self.status_var.set("Scanning files…")
        self.status_label.configure(textvariable=self.status_var)
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.results = []
        self.files_scanned = 0
        self.duplicate_count = 0
        threading.Thread(target=self._scan_worker, args=(folder, self.recursive_var.get()),
                         daemon=True).start()

    def _scan_worker(self, folder, recursive):
        groups = defaultdict(list)
        scanned = 0
        errors = 0
        try:
            if recursive:
                walker = os.walk(folder, onerror=lambda _e: None)
                for root, _dirs, files in walker:
                    for name in files:
                        path = os.path.join(root, name)
                        try:
                            size = os.path.getsize(path)
                            groups[(name.casefold(), size)].append({
                                "name": name, "size": size, "path": path
                            })
                            scanned += 1
                        except (OSError, PermissionError):
                            errors += 1
            else:
                try:
                    entries = os.scandir(folder)
                except OSError:
                    entries = []
                with entries if hasattr(entries, "__enter__") else _NullContext(entries) as it:
                    for entry in it:
                        if not entry.is_file(follow_symlinks=False):
                            continue
                        try:
                            size = entry.stat(follow_symlinks=False).st_size
                            groups[(entry.name.casefold(), size)].append({
                                "name": entry.name, "size": size, "path": entry.path
                            })
                            scanned += 1
                        except OSError:
                            errors += 1

            results = []
            duplicate_groups = 0
            duplicate_files = 0
            for _key, items in groups.items():
                is_duplicate = len(items) > 1
                if is_duplicate:
                    duplicate_groups += 1
                    duplicate_files += len(items)
                for item in items:
                    row = dict(item)
                    row["status"] = "DUPLICATE" if is_duplicate else "UNIQUE"
                    row["group"] = ""
                    results.append(row)
            results.sort(key=lambda r: (r["status"] != "DUPLICATE",
                                        r["name"].casefold(), r["size"], r["path"].casefold()))
            group_id = 0
            last_key = None
            for row in results:
                key = (row["name"].casefold(), row["size"])
                if row["status"] == "DUPLICATE":
                    if key != last_key:
                        group_id += 1
                    row["group"] = f"DUP-{group_id:03d}"
                    last_key = key
            self.root.after(0, self._scan_complete, results, scanned,
                            duplicate_files, duplicate_groups, errors)
        except Exception as exc:
            self.root.after(0, self._scan_failed, str(exc))

    def _scan_complete(self, results, scanned, duplicate_files, duplicate_groups, errors):
        self.results = results
        self.files_scanned = scanned
        self.duplicate_count = duplicate_files
        self.total_value.configure(text=str(scanned))
        self.dupe_value.configure(text=str(duplicate_files))
        self.group_value.configure(text=str(duplicate_groups))
        for row in results:
            size_text = self._format_size(row["size"])
            tag = "duplicate" if row["status"] == "DUPLICATE" else "unique"
            self.tree.insert("", "end", values=(
                row["status"], row["name"], size_text, row["group"], row["path"]
            ), tags=(tag,))
        self.progress.stop()
        self.scan_running = False
        self.scan_btn.configure(state="normal")
        message = f"Scan complete. {scanned:,} files checked."
        if errors:
            message += f" {errors:,} file(s) could not be read."
        self.status_var.set(message)

    def _scan_failed(self, error):
        self.progress.stop()
        self.scan_running = False
        self.scan_btn.configure(state="normal")
        self.status_var.set("Scan failed.")
        messagebox.showerror("Scan failed", error)

    @staticmethod
    def _format_size(size):
        units = ("B", "KB", "MB", "GB", "TB")
        value = float(size)
        for unit in units:
            if value < 1024 or unit == units[-1]:
                return f"{value:,.0f} {unit}" if unit == "B" else f"{value:,.2f} {unit}"
            value /= 1024

    def export_report(self):
        if not self.results:
            messagebox.showinfo("No results", "Run a scan before exporting a report.")
            return
        if Workbook is None:
            messagebox.showerror(
                "Missing dependency",
                "Excel export requires openpyxl.\n\nInstall it with:\npython -m pip install openpyxl"
            )
            return
        destination = self.dest_var.get().strip()
        if not destination or not os.path.isdir(destination):
            messagebox.showwarning("Destination required",
                                   "Please choose a valid report destination folder.")
            return

        filename = f"duplicate_file_report_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        output_path = os.path.join(destination, filename)
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Duplicate Scan"
            headers = ["Tag", "Group ID", "File Name", "Size (bytes)",
                       "Size (readable)", "Full Path"]
            ws.append(headers)
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F2937")
                cell.alignment = Alignment(vertical="center")
            for row in self.results:
                ws.append([
                    row["status"], row["group"], row["name"], row["size"],
                    self._format_size(row["size"]), row["path"]
                ])
                fill = "FEF3C7" if row["status"] == "DUPLICATE" else "F3F4F6"
                for cell in ws[ws.max_row]:
                    cell.fill = PatternFill("solid", fgColor=fill)
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            widths = [14, 14, 32, 16, 18, 90]
            for i, width in enumerate(widths, start=1):
                ws.column_dimensions[get_column_letter(i)].width = width
            summary = wb.create_sheet("Summary")
            summary.append(["Metric", "Value"])
            summary.append(["Scan date", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
            summary.append(["Folder scanned", self.folder_var.get().strip()])
            summary.append(["Files scanned", self.files_scanned])
            summary.append(["Duplicate files (all members of matching groups)", self.duplicate_count])
            summary.append(["Duplicate groups", self._duplicate_group_count()])
            summary.append(["Matching rule", "Case-insensitive file name + exact file size"])
            for cell in summary[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F2937")
            summary.column_dimensions["A"].width = 52
            summary.column_dimensions["B"].width = 90
            summary.freeze_panes = "A2"
            wb.save(output_path)
            self.status_var.set(f"Report exported: {output_path}")
            if messagebox.askyesno("Export complete",
                                   f"Excel report saved successfully:\n\n{output_path}\n\nOpen destination folder?"):
                try:
                    os.startfile(destination)  # Windows
                except AttributeError:
                    import subprocess, sys
                    if sys.platform == "darwin":
                        subprocess.Popen(["open", destination])
                    else:
                        subprocess.Popen(["xdg-open", destination])
                except OSError as exc:
                    messagebox.showwarning("Cannot open folder", str(exc))
        except Exception as exc:
            messagebox.showerror("Export failed", f"Could not save the Excel report:\n{exc}")

    def _duplicate_group_count(self):
        return len({(r["name"].casefold(), r["size"]) for r in self.results
                    if r["status"] == "DUPLICATE"})


class _NullContext:
    def __init__(self, value):
        self.value = value

    def __enter__(self):
        return self.value

    def __exit__(self, *_args):
        return False


def main():
    root = tk.Tk()
    DuplicateFileFinder(root)
    root.mainloop()


if __name__ == "__main__":
    main()
