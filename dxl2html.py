#!/usr/bin/env python3
"""
dxl2html – Lotus / HCL Notes Mails (DXL) als HTML anzeigen.

Funktionen
  * Mail-Kopf (Von, An, Kopie, Betreff, Datum, ...) und Rich-Text-Body
    (Formatierung, Listen, Tabellen, Sektionen, Links, eingebettete Bilder)
  * Anhänge ($FILE) werden extrahiert und verlinkt (Öffnen / Speichern);
    Vorschau für Bilder, PDF und Textdateien
  * Mehrere Mails pro DXL-Datei, Ordner als Eingabe, Index-Seite
  * Kein JavaScript, keine externen Ressourcen (Datenschutz / Sicherheit)

Nutzung
  python dxl2html.py mail.dxl                  -> dxl_html/mail.html + mail_files/
  python dxl2html.py ordner/ -o ausgabe --open
  python dxl2html.py mail.dxl --embed          -> eine einzige, selbstständige HTML
  python dxl2html.py                           -> Dateiauswahl-Dialog (tkinter)

Nur Python >= 3.8, Standardbibliothek.
"""
import argparse
import base64
import email
import email.policy
import html as htmllib
import mimetypes
import os
import re
import sys
import webbrowser
import xml.etree.ElementTree as ET
from pathlib import Path
from urllib.parse import quote

# --------------------------------------------------------------------------- #
# Helfer
# --------------------------------------------------------------------------- #

def esc(s):
    return htmllib.escape(s or "", quote=True)


def ln(el):
    """Lokaler Tag-Name ohne Namespace."""
    return el.tag.rsplit("}", 1)[-1] if isinstance(el.tag, str) else ""


def find_desc(el, name):
    for e in el.iter():
        if ln(e) == name:
            return e
    return None


def safe_name(n):
    n = os.path.basename((n or "attachment").replace("\\", "/"))
    n = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", n).strip(" .") or "attachment"
    return n[:150]


def human(n):
    for u in ("B", "KB", "MB", "GB"):
        if n < 1024 or u == "GB":
            return f"{n:.0f} {u}" if u == "B" else f"{n:.1f} {u}"
        n /= 1024


def fmt_dt(s):
    m = re.match(r"(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2}))?", (s or "").strip())
    if not m:
        return (s or "").strip()
    y, mo, d, h, mi, se = m.groups()
    return f"{d}.{mo}.{y}" if h is None else f"{d}.{mo}.{y} {h}:{mi}:{se}"


def pretty_name(n):
    m = re.match(r"(?i)^CN=([^/]+)", n or "")
    return m.group(1) if m else (n or "")


def names_html(lst):
    return ", ".join(f'<span title="{esc(n)}">{esc(pretty_name(n))}</span>' for n in lst)


COL = {"dkred": "#800000", "dkgreen": "#008000", "dkblue": "#000080",
       "dkyellow": "#808000", "dkmagenta": "#800080", "dkcyan": "#008080",
       "gray": "#808080", "ltgray": "#c0c0c0", "lightgray": "#c0c0c0",
       "dkgray": "#404040"}


def css_color(c):
    if not c:
        return None
    c = c.strip().lower()
    if c in COL:
        return COL[c]
    if re.fullmatch(r"#[0-9a-f]{3,8}", c) or re.fullmatch(r"[a-z]+", c):
        return c
    return None


IMG_EXT = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".svg"}
TXT_EXT = {".txt", ".log", ".csv", ".json", ".xml", ".md", ".ini", ".cfg",
           ".yaml", ".yml", ".sql", ".tsv", ".properties"}
ICONS = {".pdf": "📕", ".doc": "📝", ".docx": "📝", ".xls": "📊", ".xlsx": "📊",
         ".csv": "📊", ".ppt": "📽️", ".pptx": "📽️", ".zip": "🗜️", ".7z": "🗜️",
         ".rar": "🗜️", ".msg": "✉️", ".eml": "✉️", ".txt": "📄", ".ics": "📅"}

# --------------------------------------------------------------------------- #
# Datenzugriff auf das Notes-Dokument
# --------------------------------------------------------------------------- #

def item_values(item):
    out = []
    for c in item:
        t = ln(c)
        if t in ("text", "number", "datetime"):
            v = "".join(c.itertext()).strip() if t != "text" else (c.text or "")
            out.append(fmt_dt(v) if t == "datetime" else v)
        elif t in ("textlist", "numberlist", "datetimelist"):
            for x in c:
                v = "".join(x.itertext()).strip() if ln(x) != "text" else (x.text or "")
                out.append(fmt_dt(v) if ln(x) == "datetime" else v)
    return out


def index_items(doc):
    items = {}
    for it in doc:
        if ln(it) == "item":
            items.setdefault((it.get("name") or "").lower(), []).append(it)
    return items


def vals(items, name):
    out = []
    for it in items.get(name.lower(), []):
        out += item_values(it)
    return [v for v in out if v != ""]


class Att:
    def __init__(self, name, data, modified=""):
        self.name, self.data, self.modified = name, data, modified
        self.href = None
        self.dl_name = safe_name(name)


def collect_files(doc):
    files = {}
    for f in doc.iter():
        if ln(f) != "file":
            continue
        name = f.get("name") or ""
        fd = next((c for c in f if ln(c) == "filedata"), None)
        if not name or fd is None:
            continue
        try:
            data = base64.b64decode(fd.text or "")
        except Exception:
            continue
        mod = next((c for c in f if ln(c) == "modified"), None)
        modified = fmt_dt("".join(mod.itertext()).strip()) if mod is not None else ""
        files.setdefault(name, Att(name, data, modified))
    return files

# --------------------------------------------------------------------------- #
# Rich-Text-Renderer
# --------------------------------------------------------------------------- #

class Renderer:
    def __init__(self, files, filesdir, embed):
        self.files, self.filesdir, self.embed = files, filesdir, embed
        self.pardefs = {}
        self.taken = set()

    # ---- Anhänge: Speichern / URL -----------------------------------------
    def href(self, att):
        if att.href:
            return att.href
        if self.embed:
            mime = mimetypes.guess_type(att.name)[0] or "application/octet-stream"
            att.href = f"data:{mime};base64," + base64.b64encode(att.data).decode()
        else:
            self.filesdir.mkdir(parents=True, exist_ok=True)
            fn = safe_name(att.name)
            base, ext = os.path.splitext(fn)
            k = 1
            while fn in self.taken:
                k += 1
                fn = f"{base}_{k}{ext}"
            self.taken.add(fn)
            (self.filesdir / fn).write_bytes(att.data)
            att.dl_name = fn
            att.href = quote(f"{self.filesdir.name}/{fn}")
        return att.href

    def link_attrs(self, att):
        # data:-URLs dürfen Browser nicht in neuem Tab öffnen -> nur Download
        return f'download="{esc(att.dl_name)}"' if self.embed else 'target="_blank" rel="noopener"'

    # ---- Blöcke -------------------------------------------------------------
    def blocks(self, container):
        out, lst = [], None

        def close():
            nonlocal lst
            if lst:
                out.append(f"</{lst}>")
                lst = None

        for c in container:
            t = ln(c)
            if t == "pardef":
                self.pardefs[c.get("id")] = c
            elif t == "par":
                pd = self.pardefs.get(c.get("def"))
                if pd is not None and pd.get("hide"):
                    continue
                content = self.inline(c)
                kind = pd.get("list") if pd is not None else None
                if kind:
                    want = "ol" if kind in ("number", "uppernumber", "lowernumber",
                                            "alphaupper", "alphalower") else "ul"
                    if lst != want:
                        close()
                        out.append(f"<{want}>")
                        lst = want
                    out.append(f"<li>{content or '&nbsp;'}</li>")
                else:
                    close()
                    out.append(f"<p{self.par_style(pd)}>{content or '<br>'}</p>")
            elif t == "table":
                close()
                out.append(self.table(c))
            elif t == "section":
                close()
                out.append(self.section(c))
            elif t in ("attachmentref", "picture"):
                close()
                out.append(f"<p>{self.inline_el(c)}</p>")
        close()
        return "".join(out)

    @staticmethod
    def par_style(pd):
        if pd is None:
            return ""
        css = []
        al = pd.get("align")
        if al in ("center", "right"):
            css.append(f"text-align:{al}")
        elif al == "full":
            css.append("text-align:justify")
        lm = pd.get("leftmargin")
        if lm:
            m = re.match(r"([\d.]+)in", lm)
            if m and float(m.group(1)) - 1.0 > 0.05:
                css.append(f"margin-left:{float(m.group(1)) - 1.0:.2f}in")
        return f' style="{";".join(css)}"' if css else ""

    def table(self, t):
        rows = []
        for r in t:
            if ln(r) != "tablerow":
                continue
            cells = []
            for c in r:
                if ln(c) != "tablecell":
                    continue
                attrs = ""
                if c.get("columnspan"):
                    attrs += f' colspan="{esc(c.get("columnspan"))}"'
                if c.get("rowspan"):
                    attrs += f' rowspan="{esc(c.get("rowspan"))}"'
                bg = css_color(c.get("bgcolor"))
                if bg:
                    attrs += f' style="background:{bg}"'
                cells.append(f"<td{attrs}>{self.blocks(c)}</td>")
            rows.append("<tr>" + "".join(cells) + "</tr>")
        return f'<div class="tw"><table class="nt">{"".join(rows)}</table></div>'

    def section(self, s):
        title = next((c for c in s if ln(c) == "title"), None)
        head = self.inline(title) if title is not None else "Abschnitt"
        body = self.blocks([c for c in s if ln(c) != "title"])
        return f'<details open class="sec"><summary>{head}</summary>{body}</details>'

    # ---- Inline -------------------------------------------------------------
    @staticmethod
    def font_css(f):
        st = (f.get("style") or "").lower().split()
        css = []
        if "bold" in st:
            css.append("font-weight:bold")
        if "italic" in st:
            css.append("font-style:italic")
        deco = [d for d, k in (("underline", "underline"), ("line-through", "strikethrough")) if k in st]
        if deco:
            css.append("text-decoration:" + " ".join(deco))
        if "superscript" in st:
            css.append("vertical-align:super;font-size:smaller")
        if "subscript" in st:
            css.append("vertical-align:sub;font-size:smaller")
        c = css_color(f.get("color"))
        if c:
            css.append(f"color:{c}")
        sz = f.get("size") or ""
        if re.fullmatch(r"\d+(\.\d+)?pt", sz):
            css.append(f"font-size:{sz}")
        return ";".join(css)

    def inline(self, el, style=""):
        out = []

        def emit(txt, st):
            if txt:
                s = esc(txt)
                out.append(f'<span style="{st}">{s}</span>' if st else s)

        emit(el.text, style)
        cur = style
        for c in el:
            t = ln(c)
            if t == "font":
                cur = self.font_css(c)
            elif t == "run":
                out.append(self.inline(c, cur))
            elif t == "urllink":
                href = (c.get("href") or "").strip()
                inner = self.inline(c, cur)
                if re.match(r"(?i)^(https?:|mailto:|ftp:)", href):
                    out.append(f'<a href="{esc(href)}" target="_blank" rel="noopener noreferrer">{inner}</a>')
                else:
                    out.append(inner)
            else:
                out.append(self.inline_el(c, cur))
            emit(c.tail, cur)
        return "".join(out)

    def inline_el(self, c, cur=""):
        t = ln(c)
        if t == "break":
            return "<br>"
        if t == "tab":
            return "&emsp;"
        if t == "attachmentref":
            name = c.get("name") or ""
            disp = c.get("displayname") or name
            att = self.files.get(name)
            if att:
                return (f'<a class="att" href="{self.href(att)}" {self.link_attrs(att)}>'
                        f'📎 {esc(disp)}</a>')
            return f'<span class="att miss">📎 {esc(disp)} (Datei nicht in der DXL enthalten)</span>'
        if t == "picture":
            return self.picture(c)
        if t == "sharedimageref":
            return '<span class="miss">[Gemeinsam genutztes Bild]</span>'
        if t == "doclink":
            return f'<span class="miss">[Doclink: {esc(c.get("description") or c.get("hint") or "")}]</span>'
        if t in ("pardef", "title"):
            return ""
        return self.inline(c, cur)

    def picture(self, el):
        for c in el:
            t = ln(c)
            if t in ("gif", "jpeg", "png") and (c.text or "").strip():
                data = re.sub(r"\s+", "", c.text)
                return f'<img class="pic" alt="" src="data:image/{t};base64,{data}">'
            if t == "imageref":
                att = self.files.get(c.get("name") or "")
                if att:
                    return f'<img class="pic" alt="{esc(att.name)}" src="{self.href(att)}">'
            if t == "notesbitmap":
                return '<span class="miss">[Notes-Bitmap – nicht darstellbar]</span>'
        return ""

# --------------------------------------------------------------------------- #
# Mail -> HTML
# --------------------------------------------------------------------------- #

CSS = """
:root{color-scheme:light}
*{box-sizing:border-box}
body{margin:0;background:#eef0f3;font:14px/1.45 -apple-system,Segoe UI,Roboto,Arial,sans-serif;color:#1b1f24}
.wrap{max-width:960px;margin:24px auto;background:#fff;border:1px solid #d5d9df;border-radius:8px;overflow:hidden}
.hdr{padding:18px 22px;background:#f7f8fa;border-bottom:1px solid #e1e4e8}
.hdr h1{margin:0 0 10px;font-size:20px}
.hdr table{border-collapse:collapse}
.hdr th{color:#5b6570;font-weight:600;text-align:left;padding:2px 14px 2px 0;vertical-align:top;white-space:nowrap}
.hdr td{padding:2px 0;word-break:break-word}
.body{padding:20px 22px;overflow-wrap:anywhere}
.body p{margin:0;min-height:1.35em}
.body ul,.body ol{margin:.2em 0 .2em 1.4em;padding:0}
.pic{max-width:100%;height:auto;vertical-align:middle}
.tw{overflow-x:auto;margin:.4em 0}
.nt{border-collapse:collapse}
.nt td{border:1px solid #cfd4da;padding:4px 8px;vertical-align:top}
.sec>summary{cursor:pointer;font-weight:600}
.miss{color:#9a5b00}
.att{white-space:nowrap}
.atts{padding:16px 22px 20px;border-top:1px solid #e1e4e8;background:#fbfbfc}
.atts h2{margin:0 0 10px;font-size:15px}
.a{border:1px solid #dde1e6;border-radius:6px;background:#fff;padding:8px 12px;margin-bottom:8px}
.a small{color:#6b7580;margin-left:6px}
.a .dl{margin-left:10px;font-size:12px}
.prev{max-width:100%;max-height:520px;display:block;margin-top:8px;border:1px solid #e1e4e8}
.pdf{width:100%;height:640px;border:1px solid #e1e4e8;margin-top:6px}
pre.txt{max-height:360px;overflow:auto;background:#f4f5f7;padding:8px;margin:8px 0 0;font-size:12px;white-space:pre-wrap}
a{color:#0b5cad}
@media print{body{background:#fff}.wrap{border:0;margin:0}.pdf{display:none}}
"""


def first(lst, default=""):
    return lst[0] if lst else default


def fallback_body(items):
    """Body ohne Rich-Text: reiner Text oder MIME (best effort)."""
    for it in items.get("body", []):
        for c in it:
            t = ln(c)
            if t in ("text", "textlist"):
                txt = "\n".join(item_values(it))
                if txt.strip():
                    return f'<pre class="txt" style="max-height:none">{esc(txt)}</pre>'
            if t == "rawitemdata" and (c.text or "").strip():
                try:
                    msg = email.message_from_bytes(base64.b64decode(c.text),
                                                   policy=email.policy.default)
                    part = msg.get_body(preferencelist=("plain", "html"))
                    if part is not None:
                        txt = part.get_content()
                        return f'<pre class="txt" style="max-height:none">{esc(txt)}</pre>'
                except Exception:
                    pass
    return '<p class="miss">Kein darstellbarer Nachrichtentext gefunden.</p>'


def attachments_html(r):
    if not r.files:
        return ""
    rows = []
    for att in r.files.values():
        href = r.href(att)
        ext = Path(att.name).suffix.lower()
        meta = human(len(att.data)) + (f", geändert {att.modified}" if att.modified else "")
        row = (f'<div class="a"><span>{ICONS.get(ext, "🖼️" if ext in IMG_EXT else "📎")}</span> '
               f'<a href="{href}" {r.link_attrs(att)}><b>{esc(att.name)}</b></a>'
               f'<small>{meta}</small>'
               f'<a class="dl" href="{href}" download="{esc(att.dl_name)}">Speichern</a>')
        if ext in IMG_EXT:
            row += f'<img class="prev" src="{href}" alt="{esc(att.name)}">'
        elif ext == ".pdf":
            row += f'<details open><summary>PDF-Vorschau</summary><iframe class="pdf" src="{href}"></iframe></details>'
        elif ext in TXT_EXT:
            raw = att.data[:150_000]
            try:
                txt = raw.decode("utf-8")
            except UnicodeDecodeError:
                txt = raw.decode("cp1252", errors="replace")
            more = "\n[… gekürzt …]" if len(att.data) > 150_000 else ""
            row += f'<details><summary>Textvorschau</summary><pre class="txt">{esc(txt + more)}</pre></details>'
        rows.append(row + "</div>")
    return f'<div class="atts"><h2>Anhänge ({len(r.files)})</h2>{"".join(rows)}</div>'


def convert_doc(doc, outdir, stem, embed):
    items = index_items(doc)
    files = collect_files(doc)
    r = Renderer(files, outdir / f"{stem}_files", embed)

    subject = first(vals(items, "Subject")) or "(kein Betreff)"
    sender = first(vals(items, "From")) or first(vals(items, "Principal")) or first(vals(items, "$AltFrom"))
    date = (first(vals(items, "PostedDate")) or first(vals(items, "DeliveredDate"))
            or first(vals(items, "$Created")))
    if not date:
        ni = find_desc(doc, "noteinfo")
        cr = find_desc(ni, "created") if ni is not None else None
        date = fmt_dt("".join(cr.itertext()).strip()) if cr is not None else ""

    head_rows = []

    def row(label, content):
        if content:
            head_rows.append(f"<tr><th>{label}</th><td>{content}</td></tr>")

    row("Von", names_html([sender]) if sender else "")
    row("An", names_html(vals(items, "SendTo")))
    row("Kopie", names_html(vals(items, "CopyTo")))
    row("Blindkopie", names_html(vals(items, "BlindCopyTo")))
    row("Datum", esc(date))
    imp = first(vals(items, "Importance"))
    row("Priorität", {"1": "Hoch", "3": "Niedrig"}.get(imp, ""))
    if files:
        row("Anhänge", esc(f"{len(files)} Datei(en)"))

    body = ""
    for it in items.get("body", []):
        rt = next((c for c in it if ln(c) == "richtext"), None)
        if rt is not None:
            body += r.blocks(rt)
    if not body.strip():
        body = fallback_body(items)

    atts = attachments_html(r)
    page = f"""<!DOCTYPE html>
<html lang="de"><head><meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="script-src 'none'">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{esc(subject)}</title><style>{CSS}</style></head>
<body><div class="wrap">
<div class="hdr"><h1>{esc(subject)}</h1><table>{"".join(head_rows)}</table></div>
<div class="body">{body}</div>
{atts}
</div></body></html>"""
    (outdir / f"{stem}.html").write_text(page, encoding="utf-8")
    return {"file": f"{stem}.html", "subject": subject, "from": pretty_name(sender),
            "date": date, "atts": len(files)}


def write_index(outdir, entries):
    rows = "".join(
        f'<tr><td><a href="{quote(e["file"])}">{esc(e["subject"])}</a></td>'
        f'<td>{esc(e["from"])}</td><td>{esc(e["date"])}</td>'
        f'<td>{"📎 " + str(e["atts"]) if e["atts"] else ""}</td></tr>'
        for e in entries)
    page = f"""<!DOCTYPE html><html lang="de"><head><meta charset="utf-8">
<title>DXL-Mails</title><style>{CSS}
.idx{{width:100%;border-collapse:collapse}}
.idx th,.idx td{{padding:8px 12px;border-bottom:1px solid #e1e4e8;text-align:left}}
</style></head><body><div class="wrap"><div class="hdr"><h1>{len(entries)} Mails</h1></div>
<table class="idx"><tr><th>Betreff</th><th>Von</th><th>Datum</th><th>Anhänge</th></tr>{rows}</table>
</div></body></html>"""
    p = outdir / "index.html"
    p.write_text(page, encoding="utf-8")
    return p


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #

def gather_inputs(paths):
    out = []
    for p in paths:
        p = Path(p)
        if p.is_dir():
            out += sorted(x for x in p.rglob("*") if x.suffix.lower() in (".dxl", ".xml"))
        elif p.exists():
            out.append(p)
        else:
            print(f"Nicht gefunden: {p}", file=sys.stderr)
    return out


def main():
    ap = argparse.ArgumentParser(description="Lotus/HCL Notes DXL-Mails nach HTML konvertieren")
    ap.add_argument("inputs", nargs="*", help="DXL-Datei(en) oder Ordner")
    ap.add_argument("-o", "--out", help="Ausgabeordner (Standard: ./dxl_html)")
    ap.add_argument("--embed", action="store_true",
                    help="Anhänge als data:-URI einbetten (eine einzige HTML-Datei)")
    ap.add_argument("--open", action="store_true", help="Ergebnis im Browser öffnen")
    a = ap.parse_args()

    inputs = a.inputs
    dialog = False
    if not inputs:
        try:
            import tkinter
            from tkinter import filedialog
            root = tkinter.Tk()
            root.withdraw()
            inputs = list(filedialog.askopenfilenames(
                title="DXL-Datei(en) wählen",
                filetypes=[("DXL", "*.dxl *.xml"), ("Alle Dateien", "*.*")]))
            root.destroy()
            dialog = True
        except Exception:
            ap.error("Keine Eingabe angegeben.")
        if not inputs:
            return 1

    files = gather_inputs(inputs)
    if not files:
        print("Keine DXL-Dateien gefunden.", file=sys.stderr)
        return 1

    outdir = Path(a.out) if a.out else (files[0].parent / "dxl_html" if dialog else Path("dxl_html"))
    outdir.mkdir(parents=True, exist_ok=True)

    entries, used = [], set()
    for f in files:
        n = 0
        try:
            for _, el in ET.iterparse(str(f), events=("end",)):
                if ln(el) != "document":
                    continue
                n += 1
                stem = safe_name(f.stem) + ("" if n == 1 else f"_{n}")
                base, k = stem, 1
                while stem in used:
                    k += 1
                    stem = f"{base}-{k}"
                used.add(stem)
                entries.append(convert_doc(el, outdir, stem, a.embed))
                print(f"OK  {f.name} -> {entries[-1]['file']}  ({entries[-1]['atts']} Anhang/Anhänge)")
                el.clear()
        except ET.ParseError as e:
            print(f"FEHLER {f}: kein gültiges XML ({e})", file=sys.stderr)
        if n == 0:
            print(f"Hinweis: {f.name} enthält kein <document>.", file=sys.stderr)

    if not entries:
        return 1
    target = outdir / entries[0]["file"]
    if len(entries) > 1:
        target = write_index(outdir, entries)
    print(f"\nFertig: {target.resolve()}")
    if a.open or dialog:
        webbrowser.open(target.resolve().as_uri())
    return 0


if __name__ == "__main__":
    sys.exit(main())
