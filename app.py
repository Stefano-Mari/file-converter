"""File Converter: a small desktop UI around converter.py."""
from __future__ import annotations

import os
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import converter

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    BaseTk = TkinterDnD.Tk
    HAS_DND = True
except Exception:  # drag-and-drop is optional
    BaseTk = tk.Tk
    HAS_DND = False

# ---------------------------------------------------------------- palette
BG = "#f5f6f8"
CARD = "#ffffff"
BORDER = "#dfe2e7"
TEXT = "#1f2328"
MUTED = "#6b7280"
ACCENT = "#2563eb"
ACCENT_HOVER = "#1d4ed8"
OK = "#15803d"
ERR = "#b91c1c"

LABELS = {
    "pdf": "PDF document", "docx": "Word document", "doc": "Word 97-2003", "rtf": "Rich Text",
    "odt": "OpenDocument Text", "txt": "Plain text", "html": "Web page (HTML)", "md": "Markdown",
    "pptx": "PowerPoint", "odp": "OpenDocument Slides", "xlsx": "Excel workbook", "xls": "Excel 97-2003",
    "ods": "OpenDocument Sheet", "csv": "CSV", "tsv": "TSV", "json": "JSON",
    "png": "PNG image", "jpg": "JPEG image", "webp": "WebP image", "bmp": "Bitmap image",
    "gif": "GIF (animated)", "tiff": "TIFF image", "ico": "Icon (.ico)",
    "mp3": "MP3 audio", "wav": "WAV audio", "flac": "FLAC audio", "ogg": "OGG audio",
    "m4a": "M4A audio", "aac": "AAC audio", "opus": "Opus audio",
    "mp4": "MP4 video", "mkv": "MKV video", "avi": "AVI video", "mov": "MOV video", "webm": "WebM video",
}


PAGE_MODES = [
    ("Combine into one image", "combine"),
    ("One file per page / slide", "separate"),
    ("One file per page, zipped", "zip"),
]


def label_for(ext: str) -> str:
    return f"{ext.upper()}  —  {LABELS.get(ext, ext.upper())}"


class App(BaseTk):
    def __init__(self):
        super().__init__()
        self.title("File Converter")
        self.geometry("900x720")
        self.minsize(760, 640)
        self.configure(bg=BG)
        icon = Path(__file__).with_name("icon.ico")
        if icon.exists():
            self.iconbitmap(default=str(icon))

        self.files: dict[str, Path] = {}          # tree item id -> path
        self.outputs: dict[str, list[Path]] = {}  # tree item id -> converted files
        self.events: queue.Queue = queue.Queue()
        self.busy = False
        self.out_dir: Path | None = None
        self.target_map: dict[str, str] = {}

        self._style()
        self._build()
        self._refresh_targets()
        self.after(100, self._poll)

    # ------------------------------------------------------------ layout
    def _style(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        base_font = ("Segoe UI", 10)
        s.configure(".", background=BG, foreground=TEXT, font=base_font)
        s.configure("Card.TFrame", background=CARD, relief="solid", borderwidth=1, bordercolor=BORDER)
        s.configure("Card.TLabel", background=CARD)
        s.configure("Muted.TLabel", foreground=MUTED, background=BG)
        s.configure("CardMuted.TLabel", foreground=MUTED, background=CARD)
        s.configure("Title.TLabel", font=("Segoe UI Semibold", 17), background=BG)
        s.configure("TButton", padding=(12, 6), background=CARD, bordercolor=BORDER, focuscolor=BG)
        s.map("TButton", background=[("active", "#eef0f3")])
        s.configure("Accent.TButton", background=ACCENT, foreground="white", bordercolor=ACCENT,
                    font=("Segoe UI Semibold", 11), padding=(22, 9))
        s.map("Accent.TButton", background=[("disabled", "#9db7f0"), ("active", ACCENT_HOVER)],
              foreground=[("disabled", "white")])
        s.configure("Treeview", background=CARD, fieldbackground=CARD, rowheight=28, bordercolor=BORDER)
        s.configure("Treeview.Heading", background="#eef0f3", font=("Segoe UI Semibold", 10), relief="flat")
        s.map("Treeview", background=[("selected", "#dbe7fe")], foreground=[("selected", TEXT)])
        s.configure("TCombobox", padding=5)
        s.configure("TRadiobutton", background=CARD)
        s.configure("Horizontal.TProgressbar", troughcolor="#e5e7eb", background=ACCENT, thickness=8)

    def _build(self):
        pad = {"padx": 18}

        header = ttk.Frame(self)
        header.pack(fill="x", pady=(16, 4), **pad)
        ttk.Label(header, text="File Converter", style="Title.TLabel").pack(side="left")
        hint = "Drag files onto the list, or use Add files." if HAS_DND else "Use Add files to choose files."
        ttk.Label(header, text=hint, style="Muted.TLabel").pack(side="left", padx=14, pady=(6, 0))

        # file list card
        card = ttk.Frame(self, style="Card.TFrame", padding=10)
        card.pack(fill="both", expand=True, pady=8, **pad)

        bar = ttk.Frame(card, style="Card.TFrame", padding=0)
        bar.configure(borderwidth=0)
        bar.pack(fill="x", pady=(0, 8))
        ttk.Button(bar, text="+  Add files", command=self.add_files).pack(side="left")
        ttk.Button(bar, text="Add folder", command=self.add_folder).pack(side="left", padx=6)
        ttk.Button(bar, text="Remove", command=self.remove_selected).pack(side="left")
        ttk.Button(bar, text="Clear", command=self.clear).pack(side="left", padx=6)
        self.count_lbl = ttk.Label(bar, text="", style="CardMuted.TLabel")
        self.count_lbl.pack(side="right")

        tree_wrap = ttk.Frame(card)
        tree_wrap.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(tree_wrap, columns=("type", "status"), show="tree headings", selectmode="extended")
        self.tree.heading("#0", text="File", anchor="w")
        self.tree.heading("type", text="Type", anchor="w")
        self.tree.heading("status", text="Status", anchor="w")
        self.tree.column("#0", width=360, stretch=True)
        self.tree.column("type", width=140, stretch=False)
        self.tree.column("status", width=260, stretch=True)
        self.tree.tag_configure("ok", foreground=OK)
        self.tree.tag_configure("err", foreground=ERR)
        self.tree.tag_configure("skip", foreground=MUTED)
        sb = ttk.Scrollbar(tree_wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", self._open_item)
        self.tree.bind("<Delete>", lambda e: self.remove_selected())

        self.empty_lbl = tk.Label(self.tree, text="Drop files here\n\nPDF · Word · PowerPoint · Excel · Markdown · "
                                  "images · audio · video · CSV/JSON",
                                  bg=CARD, fg=MUTED, font=("Segoe UI", 11), justify="center")
        self.empty_lbl.place(relx=0.5, rely=0.5, anchor="center")

        if HAS_DND:
            for w in (self.tree, self.empty_lbl, card):
                w.drop_target_register(DND_FILES)
                w.dnd_bind("<<Drop>>", self._on_drop)

        # options card
        opts = ttk.Frame(self, style="Card.TFrame", padding=12)
        opts.pack(fill="x", pady=(0, 8), **pad)
        opts.columnconfigure(1, weight=1)

        ttk.Label(opts, text="Convert to", style="Card.TLabel", font=("Segoe UI Semibold", 10)).grid(row=0, column=0, sticky="w")
        self.target_var = tk.StringVar()
        self.target_box = ttk.Combobox(opts, textvariable=self.target_var, state="readonly", width=34)
        self.target_box.grid(row=0, column=1, sticky="w", padx=10)
        self.target_hint = ttk.Label(opts, text="", style="CardMuted.TLabel")
        self.target_hint.grid(row=0, column=2, columnspan=3, sticky="w")

        ttk.Label(opts, text="Save to", style="Card.TLabel", font=("Segoe UI Semibold", 10)).grid(row=1, column=0, sticky="w", pady=(10, 0))
        out_row = ttk.Frame(opts, style="Card.TFrame")
        out_row.configure(borderwidth=0)
        out_row.grid(row=1, column=1, columnspan=4, sticky="w", padx=10, pady=(10, 0))
        self.out_mode = tk.StringVar(value="same")
        ttk.Radiobutton(out_row, text="Same folder as each file", variable=self.out_mode, value="same").pack(side="left")
        ttk.Radiobutton(out_row, text="Other folder…", variable=self.out_mode, value="custom",
                        command=self._choose_out_dir).pack(side="left", padx=(14, 6))
        self.out_lbl = ttk.Label(out_row, text="", style="CardMuted.TLabel")
        self.out_lbl.pack(side="left")

        ttk.Label(opts, text="Quality", style="Card.TLabel", font=("Segoe UI Semibold", 10)).grid(row=2, column=0, sticky="w", pady=(10, 0))
        q_row = ttk.Frame(opts, style="Card.TFrame")
        q_row.configure(borderwidth=0)
        q_row.grid(row=2, column=1, columnspan=4, sticky="w", padx=10, pady=(10, 0))
        ttk.Label(q_row, text="Image DPI (PDF/slides → image)", style="CardMuted.TLabel").pack(side="left")
        self.dpi_var = tk.IntVar(value=150)
        ttk.Spinbox(q_row, from_=72, to=600, increment=25, textvariable=self.dpi_var, width=6).pack(side="left", padx=(6, 18))
        ttk.Label(q_row, text="JPEG/WebP quality", style="CardMuted.TLabel").pack(side="left")
        self.quality_var = tk.IntVar(value=90)
        ttk.Spinbox(q_row, from_=10, to=100, increment=5, textvariable=self.quality_var, width=5).pack(side="left", padx=6)

        ttk.Label(opts, text="Pages", style="Card.TLabel", font=("Segoe UI Semibold", 10)).grid(row=3, column=0, sticky="w", pady=(10, 0))
        p_row = ttk.Frame(opts, style="Card.TFrame")
        p_row.configure(borderwidth=0)
        p_row.grid(row=3, column=1, columnspan=4, sticky="w", padx=10, pady=(10, 0))
        self.pages_var = tk.StringVar(value=PAGE_MODES[0][0])
        ttk.Combobox(p_row, textvariable=self.pages_var, state="readonly", width=34,
                     values=[label for label, _ in PAGE_MODES]).pack(side="left")
        ttk.Label(p_row, text="when a PDF or slideshow becomes images", style="CardMuted.TLabel").pack(side="left", padx=8)

        # action bar
        actions = ttk.Frame(self)
        actions.pack(fill="x", pady=(4, 16), **pad)
        self.convert_btn = ttk.Button(actions, text="Convert", style="Accent.TButton", command=self.start)
        self.convert_btn.pack(side="right")
        self.open_btn = ttk.Button(actions, text="Open output folder", command=self.open_output, state="disabled")
        self.open_btn.pack(side="right", padx=8)
        self.progress = ttk.Progressbar(actions, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=(0, 12))
        self.status_lbl = ttk.Label(actions, text="", style="Muted.TLabel", width=24)
        self.status_lbl.pack(side="left", padx=(0, 12))

    # ------------------------------------------------------------ file list
    def _on_drop(self, event):
        self.add_paths(self.tk.splitlist(event.data))

    def add_files(self):
        exts = sorted(converter.CONVERSIONS)
        paths = filedialog.askopenfilenames(
            title="Choose files to convert",
            filetypes=[("Supported files", " ".join(f"*.{e}" for e in exts)), ("All files", "*.*")],
        )
        self.add_paths(paths)

    def add_folder(self):
        folder = filedialog.askdirectory(title="Choose a folder")
        if folder:
            self.add_paths([str(p) for p in sorted(Path(folder).iterdir()) if p.is_file()])

    def add_paths(self, paths):
        existing = {str(p).lower() for p in self.files.values()}
        skipped = 0
        for raw in paths:
            p = Path(raw)
            if p.is_dir():
                self.add_paths([str(c) for c in sorted(p.iterdir()) if c.is_file()])
                continue
            if str(p).lower() in existing:
                continue
            ext = converter.ext_of(p)
            if ext not in converter.CONVERSIONS:
                skipped += 1
                continue
            iid = self.tree.insert("", "end", text=f"  {p.name}",
                                   values=(converter.CATEGORY.get(ext, ext.upper()) + f" (.{ext})", "Ready"))
            self.files[iid] = p
            existing.add(str(p).lower())
        if skipped:
            self.status_lbl.configure(text=f"Skipped {skipped} unsupported file(s)")
        self._refresh_targets()

    def remove_selected(self):
        if self.busy:
            return
        for iid in self.tree.selection():
            self.tree.delete(iid)
            self.files.pop(iid, None)
            self.outputs.pop(iid, None)
        self._refresh_targets()

    def clear(self):
        if self.busy:
            return
        self.tree.delete(*self.tree.get_children())
        self.files.clear()
        self.outputs.clear()
        self.progress["value"] = 0
        self.status_lbl.configure(text="")
        self.open_btn.configure(state="disabled")
        self._refresh_targets()

    def _refresh_targets(self):
        n = len(self.files)
        self.count_lbl.configure(text=f"{n} file{'s' if n != 1 else ''}" if n else "")
        if n:
            self.empty_lbl.place_forget()
        else:
            self.empty_lbl.place(relx=0.5, rely=0.5, anchor="center")

        # Offer every target at least one file supports, most widely supported first.
        counts: dict[str, int] = {}
        for p in self.files.values():
            for t in converter.targets_for(p):
                counts[t] = counts.get(t, 0) + 1
        ordered = sorted(counts, key=lambda t: (-counts[t], t))
        self.target_map = {label_for(t): t for t in ordered}
        labels = list(self.target_map)
        self.target_box.configure(values=labels)
        if self.target_var.get() not in self.target_map:
            self.target_var.set(labels[0] if labels else "")
        self.target_box.bind("<<ComboboxSelected>>", lambda e: self._update_hint(counts, n))
        self._update_hint(counts, n)

    def _update_hint(self, counts, n):
        t = self.target_map.get(self.target_var.get())
        if not n:
            self.target_hint.configure(text="Add files to see available formats")
        elif t and counts.get(t, 0) < n:
            self.target_hint.configure(text=f"{n - counts[t]} file(s) can't become .{t} and will be skipped")
        else:
            self.target_hint.configure(text="")

    def _choose_out_dir(self):
        folder = filedialog.askdirectory(title="Save converted files to…")
        if folder:
            self.out_dir = Path(folder)
            shown = str(self.out_dir)
            if len(shown) > 45:
                shown = shown[:15] + "…" + shown[-28:]
            self.out_lbl.configure(text=shown)
        elif not self.out_dir:
            self.out_mode.set("same")

    # ------------------------------------------------------------ conversion
    def start(self):
        if self.busy:
            return
        if not self.files:
            messagebox.showinfo("File Converter", "Add some files first.")
            return
        target = self.target_map.get(self.target_var.get())
        if not target:
            return
        out_dir = self.out_dir if self.out_mode.get() == "custom" else None
        try:
            opts = {"dpi": int(self.dpi_var.get()), "quality": int(self.quality_var.get())}
        except (tk.TclError, ValueError):
            opts = {"dpi": 150, "quality": 90}
        opts["pages"] = dict(PAGE_MODES).get(self.pages_var.get(), "combine")

        jobs = list(self.files.items())
        self.busy = True
        self.convert_btn.configure(state="disabled", text="Converting…")
        self.progress.configure(maximum=len(jobs), value=0)
        for iid, _ in jobs:
            self.tree.item(iid, values=(self.tree.item(iid, "values")[0], "Waiting…"), tags=())
        threading.Thread(target=self._worker, args=(jobs, target, out_dir, opts), daemon=True).start()

    def _worker(self, jobs, target, out_dir, opts):
        for iid, path in jobs:
            if target not in converter.targets_for(path):
                self.events.put(("skip", iid, f"Skipped — can't make .{target}"))
                continue
            self.events.put(("running", iid, "Converting…"))
            try:
                outs = converter.convert(path, target, out_dir, opts)
                self.events.put(("ok", iid, outs))
            except Exception as e:  # report and keep going
                msg = str(e).strip().splitlines()[-1] if str(e).strip() else type(e).__name__
                self.events.put(("err", iid, msg))
        self.events.put(("done", None, None))

    def _poll(self):
        try:
            while True:
                kind, iid, data = self.events.get_nowait()
                if kind == "done":
                    self._finish()
                    continue
                type_col = self.tree.item(iid, "values")[0]
                if kind == "running":
                    self.tree.item(iid, values=(type_col, data))
                    self.tree.see(iid)
                    continue
                self.progress["value"] = self.progress["value"] + 1  # step() wraps to 0 at max
                if kind == "ok":
                    self.outputs[iid] = data
                    text = f"✓  {data[0].name}" if len(data) == 1 else f"✓  {len(data)} files ({data[0].name}, …)"
                    self.tree.item(iid, values=(type_col, text), tags=("ok",))
                elif kind == "err":
                    self.tree.item(iid, values=(type_col, f"✗  {data}"), tags=("err",))
                else:
                    self.tree.item(iid, values=(type_col, data), tags=("skip",))
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def _finish(self):
        self.busy = False
        self.convert_btn.configure(state="normal", text="Convert")
        ok = sum(1 for iid in self.files if "ok" in self.tree.item(iid, "tags"))
        bad = sum(1 for iid in self.files if "err" in self.tree.item(iid, "tags"))
        text = f"Done: {ok} converted" + (f", {bad} failed" if bad else "")
        self.status_lbl.configure(text=text)
        if self.outputs:
            self.open_btn.configure(state="normal")

    # ------------------------------------------------------------ opening results
    def _open_item(self, event):
        iid = self.tree.identify_row(event.y)
        if iid in self.outputs:
            os.startfile(self.outputs[iid][0])
        elif iid in self.files:
            os.startfile(self.files[iid])

    def open_output(self):
        if self.out_mode.get() == "custom" and self.out_dir:
            os.startfile(self.out_dir)
            return
        last = next(reversed(self.outputs.values()), None)
        if last:
            # Open Explorer with the converted file selected.
            import subprocess

            subprocess.Popen(["explorer", "/select,", str(last[0])])


if __name__ == "__main__":
    App().mainloop()
