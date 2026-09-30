# DXL2HTML Converter in Powershell

Convert **Lotus Notes / HCL Notes mail** exported as **DXL** (Domino XML) into clean, self-contained **HTML** – including **attachments**, which are extracted and made available to open or save.

Two equivalent implementations are provided:

| File | Runtime | Notes |
|---|---|---|
| `dxl2html.ps1` | Windows PowerShell 5.1 / PowerShell 7 | Native on Windows, nothing to install |
| `dxl2html.cmd` | Windows | Drag & drop launcher for the PowerShell script |
| `dxl2html.py` | Python 3.8+ | Standard library only, cross-platform |

## Features

- **Mail header**: subject, sender, recipients (To / Cc / Bcc), date, priority. Notes canonical names such as `CN=Jane Doe/O=ACME` are shown as `Jane Doe` (full name in the tooltip).
- **Rich-text body**: fonts (bold, italic, underline, strikethrough, color, size), paragraph alignment and indentation, bullet/numbered lists, tables (incl. col/row spans), collapsible sections, hyperlinks, embedded images (GIF/JPEG/PNG).
- **Attachments** (`$FILE` items):
  - extracted and linked, with **Open** and **Save** actions
  - inline **preview** for images, PDFs and text-like files (`.txt`, `.csv`, `.json`, `.xml`, `.log`, …)
  - attachments referenced inside the body are linked at their original position
- **Multiple mails per DXL file** and **folders as input**; an `index.html` overview is generated when more than one mail is converted.
- **Large exports**: the PowerShell version streams the DXL file mail by mail instead of loading it entirely into memory.
- **Safe by design**: no JavaScript, no external resources or tracking pixels, a `script-src 'none'` Content-Security-Policy in every page, and only `http(s)`, `ftp` and `mailto` links are kept.

## Quick start (Windows)

### Drag & drop
Drag a `.dxl` file (or a folder containing DXL files) onto **`dxl2html.cmd`**. The result opens in your default browser.

### PowerShell
```powershell
.\dxl2html.ps1 mail.dxl -Open                    # convert one file and open it
.\dxl2html.ps1 C:\Export -OutDir C:\Temp\mails   # convert a whole folder
.\dxl2html.ps1 mail.dxl -Embed                   # single self-contained HTML file
.\dxl2html.ps1                                   # no argument: file picker dialog
```

If Windows refuses to run the script (*"running scripts is disabled on this system"*), either start it through `dxl2html.cmd`, or unblock it once:

```powershell
Unblock-File .\dxl2html.ps1
```

### Parameters (PowerShell)

| Parameter | Description |
|---|---|
| `-Path` (positional) | One or more DXL files or folders. Folders are searched recursively for `*.dxl` / `*.xml`. Omit to get a file picker. |
| `-OutDir` | Output folder. Default: `.\dxl_html` (next to the selected file when the picker is used). |
| `-Embed` | Embed attachments as `data:` URIs so the result is one single HTML file. See [Embed mode](#embed-mode). |
| `-Open` | Open the result in the default browser. |

## Quick start (Python)

```bash
python dxl2html.py mail.dxl                 # -> dxl_html/mail.html + dxl_html/mail_files/
python dxl2html.py export_folder/ -o out --open
python dxl2html.py mail.dxl --embed
python dxl2html.py                          # no argument: file picker (tkinter)
```

| Option | Description |
|---|---|
| `inputs` | One or more DXL files or folders. |
| `-o`, `--out` | Output folder (default `./dxl_html`). |
| `--embed` | Embed attachments as `data:` URIs (single HTML file). |
| `--open` | Open the result in the default browser. |

## Output layout

```
dxl_html/
├── index.html              # only when more than one mail was converted
├── mail.html               # first mail of mail.dxl
├── mail_files/             # its attachments
│   ├── report.pdf
│   └── image.png
├── mail_2.html             # second mail contained in the same DXL file
└── mail_2_files/
```

Keep each `*.html` together with its `*_files` folder when moving or sharing the results.

## Embed mode

`-Embed` / `--embed` produces a single portable HTML file. Because browsers do not open `data:` URLs in a new tab, attachments are **download-only** in this mode (image, text and PDF previews still work in most browsers). Use the default mode if you want attachments to open directly from the page.

## Getting DXL files

DXL is the XML export format of Notes/Domino. Files can be produced by any DXL exporter, for example the `NotesDXLExporter` class in LotusScript/Java, or a Notes/Domino tool of your choice. Attachments must be included in the export (they are stored as base64 `<filedata>` inside `$FILE` items) – otherwise the page shows the attachment name with the note *"file not contained in the DXL"*.

## Limitations

- Notes-specific objects that have no web equivalent are shown as placeholders: proprietary bitmaps (`notesbitmap`), shared images, doclinks, embedded OLE objects.
- Rich-text layout is an approximation of the Notes rendering (no exact page margins, fixed fonts are not reproduced).
- Mails without a rich-text body: the PowerShell version shows plain text only; the Python version additionally tries to decode MIME content (best effort).
- Time zone information in Notes date/time values is ignored; dates are displayed as stored (`dd.mm.yyyy hh:mm:ss`).
- HTML attachments are offered for download only and are never rendered inline.

## Troubleshooting

| Symptom | Cause / fix |
|---|---|
| *Running scripts is disabled* | Use `dxl2html.cmd`, or run `Unblock-File .\dxl2html.ps1`, or start with `powershell -ExecutionPolicy Bypass -File .\dxl2html.ps1 …`. |
| *"contains no `<document>`"* | The file is not a mail/document DXL export (e.g. only design elements). |
| Parse error / invalid XML | The DXL file is truncated or contains invalid characters. Re-export it. |
| Attachment shows *"file not contained in the DXL"* | The export was created without attachment data. |
| Umlauts look wrong | Output is always UTF-8. Open the HTML in a browser rather than a legacy text editor. |

## Privacy

Everything runs locally. The generated pages load nothing from the internet and contain no scripts. Be aware that the output folder contains the extracted attachments in plain form – treat it with the same care as the original mail data.
