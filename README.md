# File Converter

A simple desktop app for converting files between formats. Drag files in, pick a format, click **Convert**.

## Run it

Double-click **`File Converter.bat`**. On the first run it sets up its own Python environment, which takes a minute or two.

Or from a terminal:

```
.venv\Scripts\python.exe app.py
```

## Using it

1. Drag files or folders onto the list, or use **Add files** / **Add folder**. Mixed file types are fine.
2. Choose a format in **Convert to**. The list only shows formats your files can actually become. Files that can't become the chosen format are skipped.
3. Choose where to save: next to each original file, or in one folder.
4. Click **Convert**. Double-click a finished row to open its result.

When a PDF or slideshow becomes images, **Pages** controls what you get:
- **Combine into one image** (default): all pages stacked into one tall image.
- **One file per page / slide**: e.g. `deck_slide1.png`, `deck_slide2.png`, …
- **One file per page, zipped**: the separate images packed into one `.zip`.

JPEG and WebP have maximum image heights (about 65,000 and 16,000 pixels). Very long documents are scaled down to fit, so for those, PNG or the zipped option keeps full detail.

## Combining files

Tick **Combine all files into one file** to merge everything in the list into a single file, in list order. Use the **↑** and **↓** buttons to reorder. The result is named after the first file, e.g. `photo_combined.pdf`.

| Combine into | What you get | Works with |
|---|---|---|
| **PDF** | One PDF, each file's pages in order | Anything that can become a PDF: images, PDFs, Word, PowerPoint, Excel, Markdown, text, HTML, CSV |
| **PNG, JPG, WebP, BMP** | One tall image, stacked top to bottom | Images, PDFs, slideshows |
| **TIFF** | One multi-page TIFF | Images, PDFs |
| **GIF** | One animated GIF, one frame per image (1 second each) | Images |
| **MP3, WAV, FLAC, OGG, M4A, AAC, Opus** | One audio track, played back to back | Audio files, and the audio of video files |
| **TXT, Markdown** | One text file | Text files, PDFs, Word documents |

Files that can't go into the chosen format are skipped. If a file fails, nothing is produced, so you never get a combined file with a piece silently missing.

Existing files are never overwritten. A new copy gets a numbered name, like `report (1).pdf`.

## Supported conversions

| From | To |
|---|---|
| **PDF** | PNG, JPG, WebP, TIFF (combined, per page, or zipped), DOCX, RTF, ODT, TXT, Markdown, HTML |
| **Word** (docx, doc, rtf, odt) | PDF, DOCX, DOC, RTF, ODT, TXT, HTML |
| **PowerPoint** (pptx, ppt, odp, ppsx) | PDF, PNG/JPG (combined, per slide, or zipped), PPTX, ODP |
| **Excel** (xlsx, xls, xlsm, ods) | PDF, CSV, TSV, JSON, HTML, Markdown table, XLSX, XLS, ODS |
| **Markdown** | PDF, HTML, DOCX, ODT, RTF |
| **Text** (txt) | PDF, HTML, DOCX, ODT, RTF |
| **HTML** | PDF, DOCX, TXT |
| **CSV / TSV / JSON** | XLSX, CSV, TSV, JSON, HTML, PDF, Markdown table |
| **Images** (png, jpg, webp, bmp, gif, tiff, ico, heic, avif) | PNG, JPG, WebP, BMP, GIF, TIFF, ICO, PDF |
| **Video** (mp4, mkv, avi, mov, webm, wmv, flv, …) | MP4, MKV, AVI, MOV, WebM, animated GIF, or audio only (MP3, WAV, FLAC, …) |
| **Audio** (mp3, wav, flac, ogg, m4a, aac, wma, opus, aiff) | MP3, WAV, FLAC, OGG, M4A, AAC, Opus |

## How it works

- **Office formats** (Word, PowerPoint, Excel, and HTML → PDF/DOCX) are converted by your installed Microsoft Office running in the background, so the output looks the same as Office's own "Save as PDF".
- **PDF** handling uses PyMuPDF. PDF → Word uses `pdf2docx`.
- **Markdown/text → PDF** renders styled HTML with PyMuPDF.
- **Audio/video** uses ffmpeg, which `imageio-ffmpeg` installs automatically.

## Files

- `app.py`: the user interface (tkinter)
- `converter.py`: the conversion engine. It can also be used from code:
  ```python
  import converter
  converter.convert("slides.pptx", "pdf")
  ```
