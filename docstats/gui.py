"""Tkinter front end for the Word / PDF table and figure statistics."""

from __future__ import annotations

import os
import queue
import threading
import tkinter as tk
import webbrowser
from tkinter import filedialog, messagebox, ttk

from .analyzer import (
    DEFAULT_COMPLEX_THRESHOLD,
    DEFAULT_FIGURE_MIN_AREA,
    DocResult,
    ScanResult,
    TableInfo,
    analyze_folder,
    open_document,
    recompute_complexity,
    recompute_figures,
)
from .export import export_csv, export_xlsx, page_text

APP_TITLE = "Table & figure statistics for Word / PDF documents"

# row background per issue type (first match wins)
TAG_STYLES = [
    ("split", "#ffe0e0"),  # split across pages
    ("nested", "#ece0ff"),  # nested table
    ("complex", "#fff2cc"),  # complex
    ("merge", "#e2f0ff"),  # merged cells
    ("error", "#ffd6d6"),
    ("small", "#f2f2f2"),  # small figure, not counted
]

FILTERS = [
    ("all", "All tables"),
    ("complex", "Complex tables"),
    ("split", "Split across pages"),
    ("nested_parent", "Tables containing nested tables"),
    ("nested_child", "Nested tables"),
    ("merge", "Tables with merged cells"),
    ("split_no_header", "Split without repeating the header"),
]
FILTER_LABELS = {key: label for key, label in FILTERS}
FILTER_KEYS = {label: key for key, label in FILTERS}


def _match_filter(t: TableInfo, key: str) -> bool:
    if key == "all":
        return True
    if key == "complex":
        return t.is_complex
    if key == "split":
        return t.is_split
    if key == "nested_parent":
        return t.nested_direct > 0
    if key == "nested_child":
        return t.is_nested
    if key == "merge":
        return t.has_merge
    if key == "split_no_header":
        return t.split_without_header
    return True


class Card(ttk.Frame):
    """A clickable metric tile in the summary bar."""

    def __init__(self, master, title: str, color: str, command=None):
        super().__init__(master, style="Card.TFrame", padding=(12, 8))
        self.command = command
        self.value = tk.StringVar(value="0")
        self.number = ttk.Label(
            self, textvariable=self.value, style="CardValue.TLabel", foreground=color
        )
        self.number.pack(anchor="w")
        self.caption = ttk.Label(self, text=title, style="CardTitle.TLabel")
        self.caption.pack(anchor="w")
        for widget in (self, self.number, self.caption):
            widget.bind("<Button-1>", self._click)
            if command:
                widget.configure(cursor="hand2")

    def _click(self, _event=None):
        if self.command:
            self.command()

    def set(self, value) -> None:
        self.value.set(str(value))


class App(ttk.Frame):
    def __init__(self, master: tk.Tk):
        super().__init__(master, padding=8)
        self.master = master
        self.pack(fill="both", expand=True)

        self.result: ScanResult | None = None
        self.table_rows: dict[str, TableInfo] = {}
        self.file_rows: dict[str, DocResult] = {}
        self.figure_rows: dict[str, object] = {}
        self.sort_state: dict[str, tuple[str, bool]] = {}

        self.queue: queue.Queue = queue.Queue()
        self.stop_event = threading.Event()
        self.worker: threading.Thread | None = None

        self.var_folder = tk.StringVar()
        self.var_recursive = tk.BooleanVar(value=True)
        self.var_word = tk.BooleanVar(value=True)
        self.var_borderless = tk.BooleanVar(value=False)
        self.var_threshold = tk.IntVar(value=DEFAULT_COMPLEX_THRESHOLD)
        self.var_fig_limit = tk.BooleanVar(value=DEFAULT_FIGURE_MIN_AREA > 0)
        self.var_fig_area = tk.IntVar(value=round(DEFAULT_FIGURE_MIN_AREA * 100))
        self.var_fig_small = tk.BooleanVar(value=False)
        self.var_filter = tk.StringVar(value=FILTER_LABELS["all"])
        self.var_search = tk.StringVar()
        self.var_file_filter = tk.StringVar(value="")
        self.var_status = tk.StringVar(value="Pick a folder with documents, then press Scan.")

        self._build_styles()
        self._build_toolbar()
        self._build_cards()
        self._build_notebook()
        self._build_statusbar()

        self.after(120, self._drain_queue)

    # ------------------------------------------------------------------
    # Building the UI
    # ------------------------------------------------------------------
    def _build_styles(self) -> None:
        import tkinter.font as tkfont

        # DPI factor, so the tables are not squashed on 125% / 150% displays
        self.scale = max(1.0, self.master.winfo_fpixels("1i") / 96.0)

        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
        base = tkfont.nametofont("TkDefaultFont")
        style.configure("Card.TFrame", relief="solid", borderwidth=1)
        style.configure("CardValue.TLabel", font=("Segoe UI Semibold", 16))
        style.configure("CardTitle.TLabel", font=("Segoe UI", 9), foreground="#555555")
        style.configure("Treeview", rowheight=base.metrics("linespace") + 8)
        style.configure("Treeview.Heading", font=("Segoe UI Semibold", 9))
        style.configure("Hint.TLabel", foreground="#666666")

    def _build_toolbar(self) -> None:
        bar = ttk.Frame(self)
        bar.pack(fill="x")

        ttk.Label(bar, text="Folder:").pack(side="left")
        self.btn_scan = ttk.Button(bar, text="Scan", command=self.start_scan, width=10)
        self.btn_scan.pack(side="right")
        self.btn_stop = ttk.Button(
            bar, text="Stop", command=self.stop_scan, width=8, state="disabled"
        )
        self.btn_stop.pack(side="right", padx=(0, 6))
        self.btn_export = ttk.Button(
            bar, text="Export Excel", command=self.export, width=13, state="disabled"
        )
        self.btn_export.pack(side="right", padx=(0, 6))
        ttk.Button(bar, text="Browse…", command=self.pick_folder, width=9).pack(
            side="right", padx=(6, 12)
        )
        entry = ttk.Entry(bar, textvariable=self.var_folder)
        entry.pack(side="left", fill="x", expand=True, padx=(6, 0))
        entry.bind("<Return>", lambda _e: self.start_scan())

        opts = ttk.Frame(self)
        opts.pack(fill="x", pady=(6, 8))
        ttk.Label(opts, text="Accepts .docx / .doc / .pdf", style="Hint.TLabel").pack(
            side="left", padx=(0, 14)
        )
        ttk.Checkbutton(opts, text="Include subfolders", variable=self.var_recursive).pack(
            side="left"
        )
        ttk.Checkbutton(
            opts, text="Word: use MS Word for page numbers", variable=self.var_word
        ).pack(side="left", padx=(14, 0))
        ttk.Checkbutton(
            opts, text="PDF: detect borderless tables", variable=self.var_borderless
        ).pack(side="left", padx=(14, 0))

        ttk.Label(opts, text="cells").pack(side="right")
        spin = ttk.Spinbox(
            opts,
            from_=2,
            to=5000,
            increment=5,
            width=6,
            textvariable=self.var_threshold,
            command=self.apply_threshold,
        )
        spin.pack(side="right", padx=(4, 3))
        spin.bind("<Return>", lambda _e: self.apply_threshold())
        ttk.Label(opts, text="Table is complex at ≥").pack(side="right", padx=(14, 0))

        figs = ttk.Frame(self)
        figs.pack(fill="x", pady=(0, 8))
        ttk.Checkbutton(
            figs,
            text="Only count a figure when its area is ≥",
            variable=self.var_fig_limit,
            command=self.apply_figure_area,
        ).pack(side="left")
        fspin = ttk.Spinbox(
            figs,
            from_=1,
            to=100,
            increment=5,
            width=5,
            textvariable=self.var_fig_area,
            command=self.apply_figure_area,
        )
        fspin.pack(side="left", padx=(4, 3))
        fspin.bind("<Return>", lambda _e: self.apply_figure_area())
        ttk.Label(
            figs,
            text="% of the page (25% = a quarter — drops logos, stamps, icons)",
            style="Hint.TLabel",
        ).pack(side="left")

    def _build_cards(self) -> None:
        wrap = ttk.Frame(self)
        wrap.pack(fill="x", pady=(0, 8))
        specs = [
            ("Files", "#333333", lambda: self.focus_files()),
            ("Tables", "#1f6feb", lambda: self.focus_filter("all")),
            ("Figures", "#0f8a5f", lambda: self.focus_figures()),
            ("Complex tables", "#b7791f", lambda: self.focus_filter("complex")),
            ("Split across pages", "#c53030", lambda: self.focus_filter("split")),
            ("Nested tables", "#6b46c1", lambda: self.focus_filter("nested_parent")),
            ("Merged cells", "#2b6cb0", lambda: self.focus_filter("merge")),
            ("Split w/o header", "#9c4221", lambda: self.focus_filter("split_no_header")),
        ]
        self.cards: dict[str, Card] = {}
        for i, (title, color, command) in enumerate(specs):
            card = Card(wrap, title, color, command)
            card.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 6, 0))
            wrap.columnconfigure(i, weight=1)
            self.cards[title] = card

    def _build_notebook(self) -> None:
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True)
        self._build_tab_files()
        self._build_tab_tables()
        self._build_tab_figures()
        self._build_tab_log()

    # -- tab 1: by file ------------------------------------------------
    def _build_tab_files(self) -> None:
        tab = ttk.Frame(self.nb, padding=6)
        self.nb.add(tab, text="  By file  ")
        ttk.Label(
            tab,
            text="Double-click a file to see the details of its tables.",
            style="Hint.TLabel",
        ).pack(anchor="w", pady=(0, 4))

        cols = [
            ("name", "File", 280, "w"),
            ("type", "Type", 55, "center"),
            ("pages", "Pages", 60, "e"),
            ("tables", "Tables", 60, "e"),
            ("figures", "Figures", 65, "e"),
            ("figsmall", "Small figs", 80, "e"),
            ("complex", "Complex", 75, "e"),
            ("split", "Split", 60, "e"),
            ("nested", "Nested", 70, "e"),
            ("merge", "Merged", 70, "e"),
            ("nohdr", "Split w/o hdr", 100, "e"),
            ("cells", "Cells", 70, "e"),
            ("mcells", "Merged away", 95, "e"),
            ("max", "Largest table", 100, "e"),
            ("folder", "Folder", 260, "w"),
        ]
        self.tree_files = self._make_tree(tab, cols, "files")
        self.tree_files.bind("<Double-1>", self._on_file_double)

    # -- tab 2: tables ---------------------------------------------------
    def _build_tab_tables(self) -> None:
        tab = ttk.Frame(self.nb, padding=6)
        self.nb.add(tab, text="  Tables  ")

        bar = ttk.Frame(tab)
        bar.pack(fill="x", pady=(0, 6))
        ttk.Label(bar, text="Filter:").pack(side="left")
        combo = ttk.Combobox(
            bar,
            textvariable=self.var_filter,
            values=[label for _k, label in FILTERS],
            state="readonly",
            width=34,
        )
        combo.pack(side="left", padx=(6, 12))
        combo.bind("<<ComboboxSelected>>", lambda _e: self.refresh_tables())

        ttk.Label(bar, text="Search:").pack(side="left")
        search = ttk.Entry(bar, textvariable=self.var_search, width=28)
        search.pack(side="left", padx=(6, 12))
        search.bind("<KeyRelease>", lambda _e: self.refresh_tables())

        self.lbl_file_filter = ttk.Label(bar, textvariable=self.var_file_filter, style="Hint.TLabel")
        self.lbl_file_filter.pack(side="left")
        ttk.Button(bar, text="Clear file filter", command=self.clear_file_filter, width=15).pack(
            side="left", padx=(6, 0)
        )
        ttk.Button(
            bar, text="Open document", command=self.open_selected, width=15
        ).pack(side="right")

        cols = [
            ("file", "File", 230, "w"),
            ("table", "Table", 75, "w"),
            ("page", "Page", 70, "center"),
            ("rows", "Rows", 55, "e"),
            ("cols", "Cols", 50, "e"),
            ("cells", "Cells (score)", 95, "e"),
            ("mcells", "Merged away", 95, "e"),
            ("mh", "Horiz. merges", 100, "e"),
            ("mv", "Vert. merges", 95, "e"),
            ("nested", "Nested", 70, "e"),
            ("hdr", "Repeat header", 100, "center"),
            ("brk", "Break after row", 110, "center"),
            ("issues", "Issues", 200, "w"),
            ("preview", "First row", 320, "w"),
        ]
        self.tree_tables = self._make_tree(tab, cols, "tables")
        self.tree_tables.bind("<Double-1>", lambda _e: self.open_selected())
        self.tree_tables.bind("<<TreeviewSelect>>", self._on_table_select)

        self.detail = tk.Text(
            tab,
            height=7,
            wrap="word",
            relief="solid",
            borderwidth=1,
            font=("Segoe UI", 9),
            padx=6,
            pady=4,
        )
        self.detail.pack(fill="x", pady=(6, 0))
        self.detail.configure(state="disabled", background="#fbfbfb")

    # -- tab 3: figures --------------------------------------------------
    def _build_tab_figures(self) -> None:
        tab = ttk.Frame(self.nb, padding=6)
        self.nb.add(tab, text="  Figures  ")
        ttk.Label(
            tab,
            text="A figure is an image, chart, SmartArt or drawing (text boxes, "
            "table ruling lines and full-page scans do not count).",
            style="Hint.TLabel",
        ).pack(anchor="w", pady=(0, 4))

        bar = ttk.Frame(tab)
        bar.pack(fill="x", pady=(0, 6))
        ttk.Checkbutton(
            bar,
            text="Show the small figures that were dropped",
            variable=self.var_fig_small,
            command=self.refresh_figures,
        ).pack(side="left", padx=(0, 14))
        ttk.Label(bar, textvariable=self.var_file_filter, style="Hint.TLabel").pack(side="left")
        ttk.Button(bar, text="Clear file filter", command=self.clear_file_filter, width=15).pack(
            side="left", padx=(6, 0)
        )

        cols = [
            ("file", "File", 300, "w"),
            ("idx", "Figure #", 70, "e"),
            ("kind", "Kind", 130, "w"),
            ("page", "Page", 60, "center"),
            ("size", "Width×Height (cm)", 125, "e"),
            ("area", "% of page", 80, "e"),
            ("counted", "Counted", 80, "center"),
            ("intbl", "Inside a table", 100, "center"),
            ("folder", "Folder", 300, "w"),
        ]
        self.tree_figs = self._make_tree(tab, cols, "figures")
        self.tree_figs.bind("<Double-1>", self._on_figure_double)

    # -- tab 4: log ------------------------------------------------------
    def _build_tab_log(self) -> None:
        tab = ttk.Frame(self.nb, padding=6)
        self.nb.add(tab, text="  Log / Errors  ")
        self.log = tk.Text(tab, wrap="word", relief="solid", borderwidth=1)
        scroll = ttk.Scrollbar(tab, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.log.pack(fill="both", expand=True)

    def _build_statusbar(self) -> None:
        bar = ttk.Frame(self)
        bar.pack(fill="x", pady=(6, 0))
        self.progress = ttk.Progressbar(bar, mode="determinate", length=220)
        self.progress.pack(side="right")
        ttk.Label(bar, textvariable=self.var_status).pack(side="left")

    def _make_tree(self, parent, cols, name: str) -> ttk.Treeview:
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)
        tree = ttk.Treeview(
            frame, columns=[c[0] for c in cols], show="headings", selectmode="browse"
        )
        for key, title, width, anchor in cols:
            tree.heading(
                key,
                text=title,
                command=lambda k=key, n=name: self.sort_by(n, k),
            )
            tree.column(
                key,
                width=int(width * self.scale),
                minwidth=40,
                anchor=anchor,
                stretch=(anchor == "w"),
            )
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        for tag, color in TAG_STYLES:
            tree.tag_configure(tag, background=color)
        return tree

    # ------------------------------------------------------------------
    # Scanning
    # ------------------------------------------------------------------
    def pick_folder(self) -> None:
        folder = filedialog.askdirectory(title="Choose a folder of documents")
        if folder:
            self.var_folder.set(os.path.normpath(folder))

    def start_scan(self) -> None:
        folder = self.var_folder.get().strip('" ')
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning(APP_TITLE, "Please choose a valid folder.")
            return
        if self.worker and self.worker.is_alive():
            return

        self.stop_event.clear()
        self.btn_scan.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        self.btn_export.configure(state="disabled")
        self.progress.configure(value=0, maximum=100)
        self._log_clear()
        self._log(f"Scanning: {folder}")
        if self.var_word.get():
            self._log(
                "Word files are opened read-only in MS Word to get page numbers. "
                "PDFs are read directly and do not need Word."
            )

        threshold = max(2, int(self.var_threshold.get() or DEFAULT_COMPLEX_THRESHOLD))
        recursive = self.var_recursive.get()
        use_word = self.var_word.get()
        borderless = self.var_borderless.get()
        fig_min_area = self.figure_min_area()

        def progress(done: int, total: int, path: str) -> None:
            self.queue.put(("progress", (done, total, path)))

        def run() -> None:
            try:
                result = analyze_folder(
                    folder,
                    recursive=recursive,
                    threshold=threshold,
                    figure_min_area=fig_min_area,
                    use_word=use_word,
                    pdf_borderless=borderless,
                    progress=progress,
                    should_stop=self.stop_event.is_set,
                )
                self.queue.put(("done", result))
            except Exception as exc:  # pragma: no cover
                import traceback

                self.queue.put(("error", traceback.format_exc()))
                del exc

        self.worker = threading.Thread(target=run, daemon=True)
        self.worker.start()

    def stop_scan(self) -> None:
        self.stop_event.set()
        self.var_status.set("Stopping…")

    def _drain_queue(self) -> None:
        try:
            while True:
                kind, payload = self.queue.get_nowait()
                if kind == "progress":
                    done, total, path = payload
                    self.progress.configure(maximum=max(total, 1), value=done)
                    self.var_status.set(
                        f"Processing {done}/{total}: {os.path.basename(path)}"
                    )
                elif kind == "done":
                    self._on_scan_done(payload)
                elif kind == "error":
                    self._finish_scan()
                    self._log(payload)
                    messagebox.showerror(APP_TITLE, "The scan failed, see the Log tab.")
        except queue.Empty:
            pass
        self.after(120, self._drain_queue)

    def _finish_scan(self) -> None:
        self.btn_scan.configure(state="normal")
        self.btn_stop.configure(state="disabled")

    def _on_scan_done(self, result: ScanResult) -> None:
        self._finish_scan()
        self.result = result
        self.btn_export.configure(state="normal" if result.docs else "disabled")

        totals = result.totals()
        skipped = (
            f" ({totals['figures_small']} small ones dropped)"
            if totals["figures_small"]
            else ""
        )
        self.var_status.set(
            f"Done: {totals['files']} files, {totals['tables']} tables, "
            f"{totals['figures']} figures{skipped} — {result.elapsed:.1f}s"
        )
        self.progress.configure(value=self.progress["maximum"])

        self._log(
            f"Scanned {totals['files_word']} Word files and {totals['files_pdf']} PDF files."
        )
        if totals["figures_small"]:
            self._log(
                f"Dropped {totals['figures_small']}/{totals['figures_all']} figures smaller "
                f"than {result.figure_min_area * 100:.0f}% of the page (logos, stamps, icons…). "
                "See them on the Figures tab → tick 'Show the small figures that were dropped'."
            )
        if totals["files_word"] and not result.word_engine and self.var_word.get():
            self._log(
                "! MS Word could not be started: for Word files the Page column and "
                "'split across pages' stay empty. (PDFs do not need Word.)"
            )
        for doc in result.docs:
            if doc.error:
                self._log(f"[ERROR] {doc.path}\n        {doc.error}")
            for warn in doc.warnings:
                if not warn.startswith("Traceback"):
                    self._log(f"[!] {doc.name}: {warn}")
        self._log(f"Finished in {result.elapsed:.1f}s.")

        self.clear_file_filter(refresh=False)
        self.refresh_all()

    # ------------------------------------------------------------------
    # Filling the views
    # ------------------------------------------------------------------
    def apply_threshold(self) -> None:
        if not self.result:
            return
        try:
            threshold = max(2, int(self.var_threshold.get()))
        except (tk.TclError, ValueError):
            return
        recompute_complexity(self.result, threshold)
        self.refresh_all()

    def figure_min_area(self) -> float:
        """The figure area threshold currently selected, as a 0..1 ratio."""
        if not self.var_fig_limit.get():
            return 0.0
        try:
            percent = int(self.var_fig_area.get())
        except (tk.TclError, ValueError):
            percent = round(DEFAULT_FIGURE_MIN_AREA * 100)
        return min(100, max(1, percent)) / 100

    def apply_figure_area(self) -> None:
        """Change the figure area threshold - recomputed at once, no rescan."""
        if not self.result:
            return
        recompute_figures(self.result, self.figure_min_area())
        self.refresh_all()
        self.var_status.set(self._figure_status())

    def refresh_all(self) -> None:
        self.refresh_cards()
        self.refresh_files()
        self.refresh_tables()
        self.refresh_figures()

    def refresh_cards(self) -> None:
        totals = self.result.totals() if self.result else {}
        mapping = {
            "Files": totals.get("files", 0),
            "Tables": totals.get("tables", 0),
            "Figures": totals.get("figures", 0),
            "Complex tables": totals.get("complex", 0),
            "Split across pages": totals.get("split", 0),
            "Nested tables": totals.get("nested_parent", 0),
            "Merged cells": totals.get("merged", 0),
            "Split w/o header": totals.get("split_no_header", 0),
        }
        for title, value in mapping.items():
            self.cards[title].set(value)

    def refresh_files(self) -> None:
        tree = self.tree_files
        tree.delete(*tree.get_children())
        self.file_rows.clear()
        if not self.result:
            return
        for doc in self.result.docs:
            tags = []
            if doc.error:
                tags.append("error")
            elif doc.n_split:
                tags.append("split")
            elif doc.n_nested_parent:
                tags.append("nested")
            elif doc.n_complex:
                tags.append("complex")
            iid = tree.insert(
                "",
                "end",
                values=(
                    doc.name,
                    doc.file_type,
                    doc.pages or "",
                    doc.n_tables_all,
                    doc.n_figures,
                    doc.n_figures_small or "",
                    doc.n_complex,
                    doc.n_split,
                    doc.n_nested_parent,
                    doc.n_merged,
                    doc.n_split_no_header,
                    doc.total_cells,
                    doc.total_merged_cells,
                    doc.max_score,
                    os.path.dirname(doc.path),
                ),
                tags=tags,
            )
            self.file_rows[iid] = doc

    def refresh_tables(self) -> None:
        tree = self.tree_tables
        tree.delete(*tree.get_children())
        self.table_rows.clear()
        if not self.result:
            return
        key = FILTER_KEYS.get(self.var_filter.get(), "all")
        needle = self.var_search.get().strip().lower()
        only_file = getattr(self, "_file_filter_path", None)

        shown = 0
        for t in self.result.all_tables():
            if only_file and t.file != only_file:
                continue
            if not _match_filter(t, key):
                continue
            if needle and needle not in " ".join(
                (t.file_name, t.label, t.caption, t.preview)
            ).lower():
                continue
            tags = []
            if t.is_split:
                tags.append("split")
            elif t.nested_direct or t.is_nested:
                tags.append("nested")
            elif t.is_complex:
                tags.append("complex")
            elif t.has_merge:
                tags.append("merge")
            iid = tree.insert(
                "",
                "end",
                values=(
                    t.file_name,
                    t.label,
                    page_text(t),
                    t.rows,
                    t.grid_cols,
                    t.cells,
                    t.merged_cells or "",
                    t.merged_h or "",
                    t.merged_v or "",
                    t.nested_direct,
                    "✓" if t.repeat_header else "",
                    ", ".join(str(r) for r in t.split_rows),
                    "; ".join(t.issue_tags()),
                    t.caption or t.preview,
                ),
                tags=tags,
            )
            self.table_rows[iid] = t
            shown += 1

        label = FILTER_LABELS.get(key, "")
        scope = (
            f"in {os.path.basename(only_file)} only"
            if only_file
            else f"across all {len(self.result.docs)} files"
        )
        self.var_status.set(f"{label}: {shown} tables {scope}.")

    def refresh_figures(self) -> None:
        tree = self.tree_figs
        tree.delete(*tree.get_children())
        self.figure_rows.clear()
        if not self.result:
            return
        only_file = getattr(self, "_file_filter_path", None)
        show_small = self.var_fig_small.get()
        for fig in self.result.all_figures():
            if only_file and fig.file != only_file:
                continue
            if not fig.counted and not show_small:
                continue
            iid = tree.insert(
                "",
                "end",
                values=(
                    fig.file_name,
                    fig.index,
                    fig.kind,
                    fig.page or "",
                    fig.size_text,
                    f"{fig.area_percent:.1f}" if fig.size_known else "?",
                    "✓" if fig.counted else "– small",
                    "✓" if fig.in_table else "",
                    os.path.dirname(fig.file),
                ),
                tags=() if fig.counted else ("small",),
            )
            self.figure_rows[iid] = fig

    # ------------------------------------------------------------------
    # Interaction
    # ------------------------------------------------------------------
    def focus_filter(self, key: str) -> None:
        # A tile shows a folder-wide total, so clicking it must drop everything
        # that narrows the list (file filter, search box) to match that number.
        self.var_filter.set(FILTER_LABELS[key])
        self.var_search.set("")
        self.clear_file_filter()
        self.nb.select(1)

    def focus_files(self) -> None:
        """The 'Files' tile: back to every file in the folder."""
        self.clear_file_filter()
        self.nb.select(0)
        n = len(self.result.docs) if self.result else 0
        self.var_status.set(f"All {n} files in the folder.")

    def focus_figures(self) -> None:
        """The 'Figures' tile: back to the figures of every file."""
        self.clear_file_filter()
        self.nb.select(2)
        self.var_status.set(self._figure_status())

    def _figure_status(self) -> str:
        totals = self.result.totals() if self.result else {}
        n = totals.get("figures", 0)
        small = totals.get("figures_small", 0)
        min_area = self.figure_min_area()
        if not min_area:
            return f"All {n} figures in the folder (counting every figure, logos included)."
        return (
            f"All {n} figures covering ≥ {min_area * 100:.0f}% of the page "
            f"— {small} smaller ones dropped."
        )

    def clear_file_filter(self, refresh: bool = True) -> None:
        self._file_filter_path = None
        self.var_file_filter.set("")
        if refresh:
            self.refresh_tables()
            self.refresh_figures()

    def _on_file_double(self, _event=None) -> None:
        sel = self.tree_files.selection()
        if not sel:
            return
        doc = self.file_rows.get(sel[0])
        if not doc:
            return
        self._file_filter_path = doc.path
        self.var_file_filter.set(f"Filtering by file: {doc.name}")
        self.var_filter.set(FILTER_LABELS["all"])
        self.nb.select(1)
        self.refresh_tables()
        self.refresh_figures()

    def _on_figure_double(self, _event=None) -> None:
        sel = self.tree_figs.selection()
        if not sel:
            return
        fig = self.figure_rows.get(sel[0])
        if fig:
            self._open_in_word(fig.file, 0, fig.page)

    def _on_table_select(self, _event=None) -> None:
        sel = self.tree_tables.selection()
        text = ""
        if sel:
            t = self.table_rows.get(sel[0])
            if t:
                lines = [
                    f"{t.file_name}  —  {t.label}"
                    + (f"  (nested in {t.parent_label})" if t.parent_label else ""),
                    f"Path: {t.file}",
                    f"Page: {page_text(t)}"
                    + (f"  (printed page: {t.page_label})" if t.page_label else "")
                    + (f"  • spans {t.pages_spanned} pages" if t.pages_spanned > 1 else "")
                    + ("  • approximate page (nested table)" if t.page_is_approx else ""),
                    f"Size: {t.rows} rows × {t.grid_cols} columns = {t.grid_cells} cells "
                    f"unmerged; actually {t.cells} cells → complexity score {t.score}",
                    (
                        f"Merges: {t.grid_cells} − {t.cells} = {t.merged_cells} cells merged away "
                        f"({t.merged_h} horizontal, {t.merged_v} vertical merge regions)"
                        if t.merged_cells
                        else "Merges: no merged cells"
                    ),
                    f"Direct nested tables: {t.nested_direct} (all levels: {t.nested_total})",
                ]
                if t.is_split:
                    where = (
                        ", ".join(f"after row {r}" for r in t.split_rows)
                        if t.split_rows
                        else "position unknown (table too large, or vertical merges)"
                    )
                    lines.append(
                        f"SPLIT ACROSS PAGES: page {t.page_start} to {t.page_end} — breaks {where}. "
                        + ("Header row is repeated." if t.repeat_header else "Header row is NOT repeated!")
                    )
                if t.caption:
                    lines.append(f"Caption: {t.caption}")
                if t.preview:
                    lines.append(f"First row: {t.preview}")
                text = "\n".join(lines)
        self.detail.configure(state="normal")
        self.detail.delete("1.0", "end")
        self.detail.insert("1.0", text)
        self.detail.configure(state="disabled")

    def open_selected(self) -> None:
        sel = self.tree_tables.selection()
        if not sel:
            messagebox.showinfo(APP_TITLE, "Please select a table in the list first.")
            return
        t = self.table_rows.get(sel[0])
        if t:
            self._open_in_word(t.file, t.word_index, t.page_start)

    def _open_in_word(self, path: str, word_index: int, page: int | None) -> None:
        self.var_status.set(f"Opening {os.path.basename(path)}…")

        def run() -> None:
            ok = open_document(path, word_index, page)
            if not ok:
                try:
                    os.startfile(path)  # noqa: S606
                except Exception:
                    webbrowser.open(path)

        threading.Thread(target=run, daemon=True).start()

    # ------------------------------------------------------------------
    def sort_by(self, tree_name: str, column: str) -> None:
        tree = {
            "files": self.tree_files,
            "tables": self.tree_tables,
            "figures": self.tree_figs,
        }[tree_name]
        prev_col, prev_desc = self.sort_state.get(tree_name, ("", False))
        desc = not prev_desc if prev_col == column else False
        self.sort_state[tree_name] = (column, desc)

        def key(item):
            value = tree.set(item, column)
            try:
                return (0, float(str(value).replace(",", "").split("-")[0].lstrip("~")))
            except ValueError:
                return (1, str(value).lower())

        items = sorted(tree.get_children(""), key=key, reverse=desc)
        for pos, item in enumerate(items):
            tree.move(item, "", pos)

    def export(self) -> None:
        if not self.result:
            return
        path = filedialog.asksaveasfilename(
            title="Save report",
            defaultextension=".xlsx",
            initialfile="document_stats.xlsx",
            filetypes=[("Excel", "*.xlsx"), ("CSV (3 files)", "*.csv")],
        )
        if not path:
            return
        try:
            if path.lower().endswith(".csv"):
                written = export_csv(self.result, os.path.dirname(path))
                msg = "Saved:\n" + "\n".join(written)
            else:
                export_xlsx(self.result, path)
                msg = f"Saved: {path}"
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Could not save: {exc}")
            return
        self._log(msg)
        if messagebox.askyesno(APP_TITLE, msg + "\n\nOpen it now?"):
            try:
                os.startfile(path if not path.lower().endswith(".csv") else os.path.dirname(path))
            except Exception:
                pass

    # ------------------------------------------------------------------
    def _log(self, text: str) -> None:
        self.log.insert("end", text.rstrip() + "\n")
        self.log.see("end")

    def _log_clear(self) -> None:
        self.log.delete("1.0", "end")


def enable_dpi_awareness() -> None:
    """Tell Windows the app handles DPI itself, so it is not blurred or oversized."""
    try:
        import ctypes

        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)  # system DPI aware
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def main(argv: list[str] | None = None) -> int:
    enable_dpi_awareness()
    root = tk.Tk()
    root.title(APP_TITLE)
    try:
        root.iconbitmap(os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.ico"))
    except tk.TclError:
        pass
    screen_w, screen_h = root.winfo_screenwidth(), root.winfo_screenheight()
    width = min(1460, screen_w - 60)
    height = min(880, screen_h - 110)
    root.geometry(
        f"{width}x{height}+{max(0, (screen_w - width) // 2)}"
        f"+{max(0, (screen_h - height) // 2 - 20)}"
    )
    root.minsize(960, 600)
    app = App(root)
    if argv:
        app.var_folder.set(argv[0])
        root.after(300, app.start_scan)
    root.mainloop()
    return 0
