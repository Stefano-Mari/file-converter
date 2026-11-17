"""Conversion engine.

Every converter is a function ``fn(src: Path, dst: Path, opts: dict) -> list[Path]``
that writes the converted file(s) and returns the paths it created.
``CONVERSIONS[src_ext][dst_ext]`` maps a pair of extensions to its converter.
"""
from __future__ import annotations

import csv
import html as html_lib
import json
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

# --------------------------------------------------------------------------- helpers


def unique_path(path: Path) -> Path:
    """Return ``path`` or ``path (1)``, ``path (2)``... so nothing is overwritten."""
    if not path.exists():
        return path
    i = 1
    while True:
        candidate = path.with_name(f"{path.stem} ({i}){path.suffix}")
        if not candidate.exists():
            return candidate
        i += 1


def _win(path: Path) -> str:
    """Absolute Windows path, as Office's COM interfaces require."""
    return str(Path(path).resolve())


CREATE_NO_WINDOW = 0x08000000 if os.name == "nt" else 0


# --------------------------------------------------------------------------- images

IMAGE_EXTS = ["png", "jpg", "jpeg", "webp", "bmp", "gif", "tiff", "tif", "ico", "heic", "heif", "avif"]
IMAGE_OUT = ["png", "jpg", "webp", "bmp", "gif", "tiff", "ico", "pdf"]

_PIL_FORMAT = {
    "png": "PNG", "jpg": "JPEG", "jpeg": "JPEG", "webp": "WEBP", "bmp": "BMP",
    "gif": "GIF", "tiff": "TIFF", "tif": "TIFF", "ico": "ICO", "pdf": "PDF",
}


def _open_image(src: Path):
    from PIL import Image

    if src.suffix.lower() in (".heic", ".heif", ".avif"):
        import pillow_heif

        pillow_heif.register_heif_opener()
    img = Image.open(src)
    # iPhone JPEGs are "MPO" files: the photo plus a hidden HDR gain map as a
    # second frame. They are single pictures, not animations.
    if getattr(img, "n_frames", 1) == 1 or img.format == "MPO":
        # Cameras store photos sideways plus an EXIF "orientation" tag; apply it
        # so the pixels are upright (and the tag isn't carried into the output).
        # The result is a plain single-frame image.
        from PIL import ImageOps

        img.seek(0)
        img = ImageOps.exif_transpose(img)
    return img


def image_to_image(src: Path, dst: Path, opts: dict) -> list[Path]:
    img = _open_image(src)
    fmt = _PIL_FORMAT[dst.suffix.lower().lstrip(".")]
    quality = int(opts.get("quality", 90))
    kwargs = {}
    if fmt in ("JPEG", "BMP", "PDF"):
        # These formats have no alpha channel: flatten onto white.
        if img.mode in ("RGBA", "LA", "P"):
            img = img.convert("RGBA")
            from PIL import Image

            bg = Image.new("RGB", img.size, (255, 255, 255))
            bg.paste(img, mask=img.split()[-1])
            img = bg
        elif img.mode != "RGB":
            img = img.convert("RGB")
    if fmt in ("JPEG", "WEBP"):
        kwargs["quality"] = quality
    if fmt == "PDF":
        kwargs["resolution"] = 100.0
    if fmt == "ICO":
        sizes = [(s, s) for s in (16, 24, 32, 48, 64, 128, 256) if s <= max(img.size)]
        kwargs["sizes"] = sizes or [(16, 16)]
    img.save(dst, fmt, **kwargs)
    return [dst]


# --------------------------------------------------------------------------- PDF


PAGE_MODES = ("combine", "separate", "zip")

# Largest single-image height each format allows.
_MAX_HEIGHT = {".jpg": 65500, ".jpeg": 65500, ".webp": 16383}


def _finish_pages(pages: list[Path], dst: Path, opts: dict, label: str) -> list[Path]:
    """Turn per-page images (in a temp dir) into the output the user asked for.

    ``combine`` stacks every page into one tall image, ``zip`` packs them into
    one archive, ``separate`` keeps one file per page next to ``dst``.
    """
    mode = opts.get("pages", "combine")
    if len(pages) == 1:
        shutil.move(pages[0], dst)
        return [dst]

    if mode == "zip":
        import zipfile

        out = unique_path(dst.with_suffix(".zip"))
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
            for p in pages:
                zf.write(p, p.name)
        return [out]

    if mode == "combine":
        return _stack_images(pages, dst, opts)

    outputs = []
    for i, p in enumerate(pages, start=1):
        out = unique_path(dst.with_name(f"{dst.stem}_{label}{i}{dst.suffix}"))
        shutil.move(p, out)
        outputs.append(out)
    return outputs


def _frames(paths: list[Path]):
    """Every frame of every image (animated GIFs and multi-page TIFFs included), as RGB."""
    from PIL import Image, ImageSequence

    for p in paths:
        img = _open_image(p)
        for frame in ImageSequence.Iterator(img):
            rgba = frame.convert("RGBA")
            bg = Image.new("RGB", rgba.size, (255, 255, 255))
            bg.paste(rgba, mask=rgba.split()[-1])
            yield bg


def _stack_images(paths: list[Path], dst: Path, opts: dict) -> list[Path]:
    """Stack images top to bottom into one tall image."""
    from PIL import Image

    images = list(_frames(paths))
    width = max(im.width for im in images)
    gap = max(8, width // 120)
    height = sum(im.height for im in images) + gap * (len(images) - 1)
    limit = _MAX_HEIGHT.get(dst.suffix.lower())
    if limit and height > limit:  # shrink everything so it fits the format
        gap = max(1, int(gap * limit / height))
        room = limit - gap * (len(images) - 1)
        scale = room / sum(im.height for im in images)
        images = [im.resize((max(1, int(im.width * scale)), max(1, int(im.height * scale))), Image.LANCZOS)
                  for im in images]
        width = max(im.width for im in images)
        height = sum(im.height for im in images) + gap * (len(images) - 1)
    sheet = Image.new("RGB", (width, height), (225, 228, 233))
    y = 0
    for im in images:
        sheet.paste(im, ((width - im.width) // 2, y))
        y += im.height + gap
    return image_to_image_obj(sheet, dst, opts)


def image_to_image_obj(img, dst: Path, opts: dict) -> list[Path]:
    """Save an already-open PIL image in the format implied by ``dst``."""
    with tempfile.TemporaryDirectory() as tmp:
        tmp_png = Path(tmp) / "sheet.png"
        img.save(tmp_png)
        return image_to_image(tmp_png, dst, opts)


def pdf_to_images(src: Path, dst: Path, opts: dict) -> list[Path]:
    import pymupdf

    dpi = int(opts.get("dpi", 150))
    ext = dst.suffix.lower()
    with tempfile.TemporaryDirectory() as tmp, pymupdf.open(src) as doc:
        pages = []
        for i, page in enumerate(doc, start=1):
            pix = page.get_pixmap(dpi=dpi, alpha=False)
            out = Path(tmp) / f"{dst.stem}_page{i}{ext}"
            if ext in (".jpg", ".jpeg"):
                pix.save(out, jpg_quality=int(opts.get("quality", 90)))
            elif ext == ".png":
                pix.save(out)
            else:  # webp, tiff etc. via Pillow
                pix.pil_save(out)
            pages.append(out)
        return _finish_pages(pages, dst, opts, "page")


def pdf_to_text(src: Path, dst: Path, opts: dict) -> list[Path]:
    import pymupdf

    with pymupdf.open(src) as doc:
        text = "\n\n".join(page.get_text() for page in doc)
    dst.write_text(text, encoding="utf-8")
    return [dst]


def pdf_to_markdown(src: Path, dst: Path, opts: dict) -> list[Path]:
    import pymupdf4llm

    md = pymupdf4llm.to_markdown(str(src), show_progress=False)
    dst.write_text(md, encoding="utf-8")
    return [dst]


def pdf_to_html(src: Path, dst: Path, opts: dict) -> list[Path]:
    import pymupdf

    with pymupdf.open(src) as doc:
        body = "\n".join(page.get_text("xhtml") for page in doc)
    dst.write_text(_wrap_html(body, src.stem), encoding="utf-8")
    return [dst]


# --------------------------------------------------------------------------- HTML / Markdown / text

_CSS = """
body { font-family: Helvetica, Arial, sans-serif; font-size: 11pt; line-height: 1.5; color: #222; }
h1, h2, h3, h4 { color: #111; margin-top: 1em; }
h1 { font-size: 22pt; } h2 { font-size: 17pt; } h3 { font-size: 14pt; }
code { font-family: Courier, monospace; font-size: 9.5pt; background: #f2f2f2; }
pre { font-family: Courier, monospace; font-size: 9.5pt; background: #f2f2f2; padding: 6pt; }
blockquote { color: #555; margin-left: 12pt; padding-left: 8pt; border-left: 3px solid #ccc; }
table { border-collapse: collapse; }
th, td { border: 1px solid #999; padding: 3pt 6pt; }
th { background: #eee; }
img { max-width: 100%; }
"""


def _wrap_html(body: str, title: str) -> str:
    return (
        "<!DOCTYPE html><html><head><meta charset='utf-8'>"
        f"<title>{html_lib.escape(title)}</title><style>{_CSS}</style></head>"
        f"<body>{body}</body></html>"
    )


def _markdown_html(src: Path) -> str:
    import markdown

    text = src.read_text(encoding="utf-8", errors="replace")
    body = markdown.markdown(
        text, extensions=["tables", "fenced_code", "sane_lists", "toc", "nl2br"]
    )
    return _wrap_html(body, src.stem)


def _text_html(src: Path) -> str:
    text = src.read_text(encoding="utf-8", errors="replace")
    return _wrap_html(f"<pre style='white-space: pre-wrap'>{html_lib.escape(text)}</pre>", src.stem)


def _html_string_to_pdf(html: str, dst: Path, base_dir: Path) -> None:
    import pymupdf

    story = pymupdf.Story(html=html, user_css=_CSS, archive=str(base_dir))
    writer = pymupdf.DocumentWriter(str(dst))
    mediabox = pymupdf.paper_rect("letter")
    where = mediabox + (54, 54, -54, -54)  # 0.75in margins
    more = True
    while more:
        device = writer.begin_page(mediabox)
        more, _ = story.place(where)
        story.draw(device)
        writer.end_page()
    writer.close()


def _html_string_via_word(html: str, dst: Path, word_format: int) -> None:
    with tempfile.TemporaryDirectory() as tmp:
        tmp_html = Path(tmp) / "doc.html"
        tmp_html.write_text(html, encoding="utf-8")
        _word_save(tmp_html, dst, word_format)


def markdown_to(src: Path, dst: Path, opts: dict) -> list[Path]:
    return _render_html_doc(_markdown_html(src), src, dst)


def text_to(src: Path, dst: Path, opts: dict) -> list[Path]:
    return _render_html_doc(_text_html(src), src, dst)


def html_to(src: Path, dst: Path, opts: dict) -> list[Path]:
    html = src.read_text(encoding="utf-8", errors="replace")
    ext = dst.suffix.lower()
    if ext in (".pdf", ".docx"):
        # Word renders real-world HTML (CSS, layout) much better than MuPDF.
        try:
            _word_save(src, dst, WORD_FORMATS[ext.lstrip(".")])
            return [dst]
        except OfficeMissing:
            pass
    if ext == ".txt":
        import re

        text = re.sub(r"<(script|style)[^>]*>.*?</\1>", "", html, flags=re.S | re.I)
        text = re.sub(r"<br\s*/?>|</p>|</div>|</h\d>|</li>", "\n", text, flags=re.I)
        text = html_lib.unescape(re.sub(r"<[^>]+>", "", text))
        dst.write_text(re.sub(r"\n{3,}", "\n\n", text).strip() + "\n", encoding="utf-8")
        return [dst]
    return _render_html_doc(html, src, dst)


def _render_html_doc(html: str, src: Path, dst: Path) -> list[Path]:
    ext = dst.suffix.lower()
    if ext == ".html":
        dst.write_text(html, encoding="utf-8")
    elif ext == ".pdf":
        _html_string_to_pdf(html, dst, src.parent)
    elif ext in (".docx", ".odt", ".rtf"):
        _html_string_via_word(html, dst, WORD_FORMATS[ext.lstrip(".")])
    elif ext == ".txt":
        shutil.copyfile(src, dst)
    else:
        raise ValueError(f"Unsupported target {ext}")
    return [dst]


# --------------------------------------------------------------------------- Microsoft Office (COM)


class OfficeMissing(RuntimeError):
    pass


def _dispatch(prog_id: str):
    try:
        import pythoncom
        import win32com.client
    except ImportError as e:  # pragma: no cover
        raise OfficeMissing("pywin32 is not installed") from e
    pythoncom.CoInitialize()
    try:
        return win32com.client.DispatchEx(prog_id)
    except Exception as e:
        raise OfficeMissing(f"Microsoft Office ({prog_id}) is not available") from e


WORD_FORMATS = {"pdf": 17, "docx": 16, "doc": 0, "rtf": 6, "txt": 7, "html": 10, "odt": 23}


def _word_save(src: Path, dst: Path, fmt: int) -> None:
    word = _dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    doc = None
    try:
        doc = word.Documents.Open(
            _win(src), ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False
        )
        if fmt == 7:  # Unicode text
            doc.SaveAs2(_win(dst), FileFormat=7, Encoding=65001)
        else:
            doc.SaveAs2(_win(dst), FileFormat=fmt)
    finally:
        if doc is not None:
            doc.Close(SaveChanges=False)
        word.Quit()


def word_convert(src: Path, dst: Path, opts: dict) -> list[Path]:
    _word_save(src, dst, WORD_FORMATS[dst.suffix.lower().lstrip(".")])
    return [dst]


def pdf_to_word(src: Path, dst: Path, opts: dict) -> list[Path]:
    """PDF -> DOCX with pdf2docx (Word's own PDF import blocks on a hidden prompt)."""
    from pdf2docx import Converter

    if dst.suffix.lower() == ".docx":
        cv = Converter(str(src))
        try:
            cv.convert(str(dst))
        finally:
            cv.close()
        return [dst]
    # rtf / odt: go through a temporary docx, then let Word re-save it.
    with tempfile.TemporaryDirectory() as tmp:
        tmp_docx = Path(tmp) / f"{src.stem}.docx"
        pdf_to_word(src, tmp_docx, opts)
        _word_save(tmp_docx, dst, WORD_FORMATS[dst.suffix.lower().lstrip(".")])
    return [dst]


PPT_FORMATS = {"pdf": 32, "pptx": 24, "ppt": 1, "odp": 35}


def powerpoint_convert(src: Path, dst: Path, opts: dict) -> list[Path]:
    ext = dst.suffix.lower().lstrip(".")
    ppt = _dispatch("PowerPoint.Application")
    pres = None
    try:
        pres = ppt.Presentations.Open(_win(src), ReadOnly=True, Untitled=False, WithWindow=False)
        if ext in ("png", "jpg"):
            dpi = int(opts.get("dpi", 150))
            width = int(pres.PageSetup.SlideWidth * dpi / 72)
            height = int(pres.PageSetup.SlideHeight * dpi / 72)
            with tempfile.TemporaryDirectory() as tmp:
                pages = []
                for i, slide in enumerate(pres.Slides, start=1):
                    out = Path(tmp) / f"{dst.stem}_slide{i}.{ext}"
                    slide.Export(_win(out), ext.upper(), width, height)
                    pages.append(out)
                return _finish_pages(pages, dst, opts, "slide")
        pres.SaveAs(_win(dst), PPT_FORMATS[ext])
        return [dst]
    finally:
        if pres is not None:
            pres.Close()
        _release_powerpoint(ppt)


def _release_powerpoint(ppt) -> None:
    """Quit PowerPoint unless the user has presentations open, and wait for it to exit.

    PowerPoint runs as one shared instance, so automation attaches to whatever
    is already running. Quitting it would close the user's own work, and
    starting the next conversion while it is still shutting down can hang.
    """
    import win32api
    import win32event
    import win32process

    try:
        if ppt.Presentations.Count > 0:
            return  # someone else's presentations are open: leave it running
        pid = win32process.GetWindowThreadProcessId(ppt.HWND)[1]
    except Exception:
        pid = None
    try:
        ppt.Quit()
    except Exception:
        return
    if pid:
        try:
            handle = win32api.OpenProcess(0x00100000, False, pid)  # SYNCHRONIZE
            win32event.WaitForSingleObject(handle, 20000)
            win32api.CloseHandle(handle)
        except Exception:
            pass


EXCEL_FORMATS = {"xlsx": 51, "xls": 56, "csv": 62, "ods": 60, "html": 44}


def excel_convert(src: Path, dst: Path, opts: dict) -> list[Path]:
    ext = dst.suffix.lower().lstrip(".")
    xl = _dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    wb = None
    try:
        wb = xl.Workbooks.Open(_win(src), ReadOnly=True)
        if ext == "pdf":
            wb.ExportAsFixedFormat(0, _win(dst))
        else:
            wb.SaveAs(_win(dst), EXCEL_FORMATS[ext])
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        xl.Quit()
    return [dst]


# --------------------------------------------------------------------------- spreadsheets / data


def _read_table(src: Path) -> list[list]:
    ext = src.suffix.lower()
    if ext in (".csv", ".tsv"):
        with open(src, newline="", encoding="utf-8-sig", errors="replace") as f:
            sample = f.read(4096)
            f.seek(0)
            if ext == ".tsv":
                dialect = csv.excel_tab
            else:
                try:
                    dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
                except csv.Error:
                    dialect = csv.excel
            return [row for row in csv.reader(f, dialect)]
    if ext == ".xlsx":
        import openpyxl

        wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
        ws = wb.worksheets[0]
        rows = [["" if v is None else v for v in row] for row in ws.iter_rows(values_only=True)]
        wb.close()
        return rows
    if ext == ".json":
        data = json.loads(src.read_text(encoding="utf-8"))
        if isinstance(data, dict):
            # {"key": [...]} -> first list value; otherwise a single record
            lists = [v for v in data.values() if isinstance(v, list)]
            data = lists[0] if lists else [data]
        if not isinstance(data, list):
            raise ValueError("JSON must be a list of objects or lists")
        if data and all(isinstance(r, dict) for r in data):
            headers: list[str] = []
            for r in data:
                for k in r:
                    if k not in headers:
                        headers.append(k)
            return [headers] + [[_cell(r.get(h, "")) for h in headers] for r in data]
        return [[_cell(c) for c in (r if isinstance(r, list) else [r])] for r in data]
    raise ValueError(f"Cannot read table from {ext}")


def _cell(v):
    return json.dumps(v) if isinstance(v, (dict, list)) else v


def _write_table(rows: list[list], dst: Path, title: str) -> None:
    ext = dst.suffix.lower()
    if ext in (".csv", ".tsv"):
        with open(dst, "w", newline="", encoding="utf-8-sig") as f:
            csv.writer(f, delimiter="\t" if ext == ".tsv" else ",").writerows(rows)
    elif ext == ".xlsx":
        import openpyxl

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = title[:31] or "Sheet1"
        for row in rows:
            ws.append([_typed(v) for v in row])
        wb.save(dst)
    elif ext == ".json":
        if rows:
            headers = [str(h) for h in rows[0]]
            records = [dict(zip(headers, (_typed(v) for v in r))) for r in rows[1:]]
        else:
            records = []
        dst.write_text(json.dumps(records, indent=2, ensure_ascii=False, default=str), encoding="utf-8")
    elif ext in (".html", ".pdf", ".md"):
        if ext == ".md":
            dst.write_text(_rows_markdown(rows), encoding="utf-8")
            return
        cells = lambda r, tag: "".join(f"<{tag}>{html_lib.escape(str(v))}</{tag}>" for v in r)
        body = "<table>"
        if rows:
            body += f"<tr>{cells(rows[0], 'th')}</tr>"
            body += "".join(f"<tr>{cells(r, 'td')}</tr>" for r in rows[1:])
        body += "</table>"
        html = _wrap_html(body, title)
        if ext == ".html":
            dst.write_text(html, encoding="utf-8")
        else:
            _html_string_to_pdf(html, dst, dst.parent)
    else:
        raise ValueError(f"Cannot write table to {ext}")


def _rows_markdown(rows: list[list]) -> str:
    if not rows:
        return ""
    esc = lambda v: str(v).replace("|", "\\|").replace("\n", " ")
    width = max(len(r) for r in rows)
    rows = [list(r) + [""] * (width - len(r)) for r in rows]
    lines = ["| " + " | ".join(esc(v) for v in rows[0]) + " |", "|" + " --- |" * width]
    lines += ["| " + " | ".join(esc(v) for v in r) + " |" for r in rows[1:]]
    return "\n".join(lines) + "\n"


def _typed(v):
    """Turn numeric-looking strings back into numbers for xlsx/json."""
    if not isinstance(v, str):
        return v
    s = v.strip()
    digits = s.lstrip("-")
    if digits.startswith("0") and len(digits) > 1 and not digits.startswith("0."):
        return v  # keep leading zeros (zip codes, IDs)
    try:
        return int(s)
    except ValueError:
        pass
    try:
        return float(s) if s.lower() not in ("nan", "inf", "-inf", "infinity") else v
    except ValueError:
        return v


def table_convert(src: Path, dst: Path, opts: dict) -> list[Path]:
    if src.suffix.lower() == ".xlsx" and dst.suffix.lower() == ".pdf":
        try:  # Excel keeps formatting, column widths, all sheets
            return excel_convert(src, dst, opts)
        except OfficeMissing:
            pass
    _write_table(_read_table(src), dst, src.stem)
    return [dst]


# --------------------------------------------------------------------------- audio / video (ffmpeg)

AUDIO_EXTS = ["mp3", "wav", "flac", "ogg", "m4a", "aac", "wma", "opus", "aiff"]
VIDEO_EXTS = ["mp4", "mkv", "avi", "mov", "webm", "wmv", "flv", "m4v", "mpeg", "mpg", "3gp"]
AUDIO_OUT = ["mp3", "wav", "flac", "ogg", "m4a", "aac", "opus"]
VIDEO_OUT = ["mp4", "mkv", "avi", "mov", "webm", "gif"]


def _ffmpeg() -> str:
    found = shutil.which("ffmpeg")
    if found:
        return found
    import imageio_ffmpeg

    return imageio_ffmpeg.get_ffmpeg_exe()


_AUDIO_ARGS = {
    "mp3": ["-c:a", "libmp3lame", "-q:a", "2"],
    "m4a": ["-c:a", "aac", "-b:a", "192k"],
    "aac": ["-c:a", "aac", "-b:a", "192k"],
    "ogg": ["-c:a", "libvorbis", "-q:a", "5"],
    "opus": ["-c:a", "libopus", "-b:a", "128k"],
}


def _run_ffmpeg(args: list[str]) -> None:
    result = subprocess.run(args, capture_output=True, text=True, creationflags=CREATE_NO_WINDOW)
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip().splitlines()[-1] if result.stderr.strip() else "ffmpeg failed")


def _concat_audio(paths: list[Path], dst: Path) -> None:
    """Join the audio of several audio/video files end to end."""
    ext = dst.suffix.lower().lstrip(".")
    args = [_ffmpeg(), "-y", "-hide_banner", "-loglevel", "error"]
    for p in paths:
        args += ["-i", str(p)]
    # Bring every input to the same sample rate / layout so they can be joined.
    rate = "48000" if ext == "opus" else "44100"
    chains = [f"[{i}:a:0]aresample={rate},aformat=sample_fmts=fltp:channel_layouts=stereo[a{i}]" for i in range(len(paths))]
    joined = "".join(f"[a{i}]" for i in range(len(paths)))
    graph = ";".join(chains) + f";{joined}concat=n={len(paths)}:v=0:a=1[out]"
    args += ["-filter_complex", graph, "-map", "[out]"] + _AUDIO_ARGS.get(ext, []) + [str(dst)]
    _run_ffmpeg(args)


def ffmpeg_convert(src: Path, dst: Path, opts: dict) -> list[Path]:
    ext = dst.suffix.lower().lstrip(".")
    args = [_ffmpeg(), "-y", "-hide_banner", "-loglevel", "error", "-i", str(src)]
    if ext in AUDIO_OUT:
        args += ["-vn"] + _AUDIO_ARGS.get(ext, [])
    elif ext == "gif":
        args += ["-vf", "fps=12,scale='min(640,iw)':-2:flags=lanczos,split[a][b];[a]palettegen[p];[b][p]paletteuse", "-loop", "0"]
    elif ext == "webm":
        args += ["-c:v", "libvpx-vp9", "-crf", "32", "-b:v", "0", "-c:a", "libopus"]
    elif ext in ("mp4", "mov", "mkv"):
        args += ["-c:v", "libx264", "-preset", "medium", "-crf", "23", "-pix_fmt", "yuv420p", "-c:a", "aac", "-b:a", "160k"]
        if ext == "mp4":
            args += ["-movflags", "+faststart"]
    args.append(str(dst))
    _run_ffmpeg(args)
    return [dst]


# --------------------------------------------------------------------------- registry

CONVERSIONS: dict[str, dict[str, callable]] = {}


def _register(sources, targets, fn):
    for s in sources:
        for t in targets:
            if s != t and not (s in ("jpg", "jpeg") and t in ("jpg", "jpeg")) and not (s in ("tif", "tiff") and t == "tiff"):
                CONVERSIONS.setdefault(s, {})[t] = fn


_register(IMAGE_EXTS, IMAGE_OUT, image_to_image)
_register(["gif"], ["mp4", "webm"], ffmpeg_convert)

_register(["pdf"], ["png", "jpg", "webp", "tiff"], pdf_to_images)
_register(["pdf"], ["txt"], pdf_to_text)
_register(["pdf"], ["md"], pdf_to_markdown)
_register(["pdf"], ["html"], pdf_to_html)
_register(["pdf"], ["docx", "rtf", "odt"], pdf_to_word)

_register(["md", "markdown"], ["pdf", "html", "docx", "odt", "rtf"], markdown_to)
_register(["txt"], ["pdf", "html", "docx", "odt", "rtf"], text_to)
_register(["html", "htm"], ["pdf", "docx", "txt"], html_to)

_register(["docx", "doc", "rtf", "odt", "wpd"], ["pdf", "docx", "doc", "rtf", "txt", "html", "odt"], word_convert)
_register(["pptx", "ppt", "odp", "ppsx", "pps"], ["pdf", "pptx", "png", "jpg", "odp"], powerpoint_convert)
_register(["xls", "ods", "xlsm", "xlsb"], ["pdf", "xlsx", "csv", "ods", "html"], excel_convert)
_register(["xlsx"], ["pdf", "csv", "tsv", "json", "html", "md"], table_convert)
_register(["xlsx"], ["xls", "ods"], excel_convert)
_register(["csv", "tsv", "json"], ["xlsx", "csv", "tsv", "json", "html", "pdf", "md"], table_convert)

_register(AUDIO_EXTS, AUDIO_OUT, ffmpeg_convert)
_register(VIDEO_EXTS, VIDEO_OUT + AUDIO_OUT, ffmpeg_convert)

CATEGORY = {}
for _e in IMAGE_EXTS:
    CATEGORY[_e] = "Image"
for _e in AUDIO_EXTS:
    CATEGORY[_e] = "Audio"
for _e in VIDEO_EXTS:
    CATEGORY[_e] = "Video"
for _e in ["pdf"]:
    CATEGORY[_e] = "PDF"
for _e in ["docx", "doc", "rtf", "odt", "wpd", "txt", "md", "markdown", "html", "htm"]:
    CATEGORY[_e] = "Document"
for _e in ["pptx", "ppt", "odp", "ppsx", "pps"]:
    CATEGORY[_e] = "Presentation"
for _e in ["xlsx", "xls", "ods", "xlsm", "xlsb", "csv", "tsv", "json"]:
    CATEGORY[_e] = "Spreadsheet"


def ext_of(path: Path | str) -> str:
    return Path(path).suffix.lower().lstrip(".")


def targets_for(path: Path | str) -> list[str]:
    return sorted(CONVERSIONS.get(ext_of(path), {}))


def convert(src: Path | str, target_ext: str, out_dir: Path | str | None = None, opts: dict | None = None) -> list[Path]:
    src = Path(src)
    target_ext = target_ext.lower().lstrip(".")
    fn = CONVERSIONS.get(ext_of(src), {}).get(target_ext)
    if fn is None:
        raise ValueError(f"Can't convert .{ext_of(src)} to .{target_ext}")
    out_dir = Path(out_dir) if out_dir else src.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    dst = unique_path(out_dir / f"{src.stem}.{target_ext}")
    return fn(src, dst, opts or {})


# --------------------------------------------------------------------------- combining files

_SAME = {"jpeg": "jpg", "tif": "tiff", "markdown": "md", "htm": "html"}
_STACKED = {"png", "jpg", "webp", "bmp"}
MERGE_TARGETS = {"pdf", "tiff", "gif", "txt", "md"} | _STACKED | set(AUDIO_OUT)


def _norm(ext: str) -> str:
    return _SAME.get(ext, ext)


def merge_targets_for(path: Path | str) -> list[str]:
    """Formats this file can be combined into (together with other files)."""
    return sorted(t for t in MERGE_TARGETS if can_merge(path, t))


def can_merge(path: Path | str, target: str) -> bool:
    ext = ext_of(path)
    if target in AUDIO_OUT:
        return ext in AUDIO_EXTS or ext in VIDEO_EXTS
    return _norm(ext) == target or target in CONVERSIONS.get(ext, {})


def _merge_parts(parts: list[Path], dst: Path, opts: dict) -> None:
    target = dst.suffix.lower().lstrip(".")
    if target == "pdf":
        import pymupdf

        out = pymupdf.open()
        for p in parts:
            with pymupdf.open(p) as d:
                out.insert_pdf(d)
        out.save(dst, garbage=3, deflate=True)
        out.close()
    elif target in _STACKED:
        _stack_images(parts, dst, opts)
    elif target == "tiff":  # one page per image
        frames = list(_frames(parts))
        frames[0].save(dst, "TIFF", save_all=True, append_images=frames[1:], compression="tiff_deflate")
    elif target == "gif":  # one animation frame per image
        from PIL import Image

        frames = list(_frames(parts))
        for f in frames:
            f.thumbnail((1600, 1600))
        size = (max(f.width for f in frames), max(f.height for f in frames))
        padded = []
        for f in frames:
            canvas = Image.new("RGB", size, (255, 255, 255))
            canvas.paste(f, ((size[0] - f.width) // 2, (size[1] - f.height) // 2))
            padded.append(canvas)
        padded[0].save(dst, "GIF", save_all=True, append_images=padded[1:],
                       duration=int(opts.get("frame_ms", 1000)), loop=0)
    elif target in ("txt", "md"):
        texts = [p.read_text(encoding="utf-8-sig", errors="replace").rstrip() for p in parts]
        dst.write_text("\n\n".join(texts) + "\n", encoding="utf-8")
    else:
        raise ValueError(f"Can't combine files into .{target}")


def merge(paths, target_ext: str, out_dir: Path | str | None = None, opts: dict | None = None,
          progress=None) -> list[Path]:
    """Combine several files, in order, into one ``target_ext`` file.

    Each file is first converted to the target format if needed, so e.g. a
    Word document, two JPEGs and a PDF can all go into one PDF.
    ``progress(index, state)`` is called with state "start"/"done" per file and
    ``(None, "merging")`` before the final join.
    """
    paths = [Path(p) for p in paths]
    target = _norm(target_ext.lower().lstrip("."))
    if target not in MERGE_TARGETS:
        raise ValueError(f"Can't combine files into .{target}")
    if not paths:
        raise ValueError("No files to combine")
    notify = progress or (lambda i, state: None)
    out_dir = Path(out_dir) if out_dir else paths[0].parent
    out_dir.mkdir(parents=True, exist_ok=True)
    dst = unique_path(out_dir / f"{paths[0].stem}_combined.{target}")
    opts = dict(opts or {}, pages="separate")  # every page/slide becomes its own part

    if target in AUDIO_OUT:
        for i in range(len(paths)):
            notify(i, "start")
            notify(i, "done")
        notify(None, "merging")
        _concat_audio(paths, dst)
        return [dst]

    with tempfile.TemporaryDirectory() as tmp:
        parts: list[Path] = []
        for i, p in enumerate(paths):
            notify(i, "start")
            if _norm(ext_of(p)) == target:
                parts.append(p)
            else:
                sub = Path(tmp) / str(i)
                sub.mkdir()
                parts += convert(p, target, sub, opts)
            notify(i, "done")
        notify(None, "merging")
        _merge_parts(parts, dst, opts)
    return [dst]
