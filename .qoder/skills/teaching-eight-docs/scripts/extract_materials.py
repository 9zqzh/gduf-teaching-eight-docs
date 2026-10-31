"""Extract text and tables from course materials into one JSON file.

Usage:
    python extract_materials.py --out materials.json --dir "D:/materials"
    python extract_materials.py --out materials.json a.docx b.ppt c.pdf

Supported: .docx/.doc (Word), .pptx/.ppt (PowerPoint), .pdf. Legacy .doc/.ppt
files are converted through Office COM to a temporary .docx/.pptx first.

Output JSON:
    {"files": [{"path", "kind", "title", "paragraphs", "tables", "slides"}],
     "digest": "===== <file> =====\n..."}

The digest is what an agent should read first; per-file fields keep the full
content for targeted lookups.
"""

import argparse
import glob
import json
import os
import shutil
import sys
import tempfile

WORD_FORMATS = (".doc", ".docx")
PPT_FORMATS = (".ppt", ".pptx")
PDF_FORMATS = (".pdf",)

# ---------------------------------------------------------------- office COM
def _word_convert(src, dst, fmt):
    import win32com.client as win32

    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(os.path.abspath(src), ReadOnly=True)
        doc.SaveAs2(os.path.abspath(dst), FileFormat=fmt)
        doc.Close(False)
    finally:
        word.Quit()


def _ppt_convert(src, dst):
    import win32com.client as win32

    ppt = win32.gencache.EnsureDispatch("PowerPoint.Application")
    try:
        pres = ppt.Presentations.Open(os.path.abspath(src), WithWindow=False)
        pres.SaveAs(os.path.abspath(dst), 24)  # ppSaveAsOpenXMLPresentation
        pres.Close()
    finally:
        ppt.Quit()


# ---------------------------------------------------------------- parsers
def parse_docx(path):
    from docx import Document

    doc = Document(path)
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    tables = []
    for table in doc.tables:
        rows = []
        for row in table.rows:
            cells = []
            seen = set()
            for cell in row.cells:
                if id(cell._tc) in seen:
                    continue
                seen.add(id(cell._tc))
                cells.append(cell.text.strip())
            rows.append(cells)
        tables.append(rows)
    title = paragraphs[0] if paragraphs else ""
    return {"title": title, "paragraphs": paragraphs, "tables": tables}


def parse_pptx(path):
    from pptx import Presentation

    prs = Presentation(path)
    slides = []
    for i, slide in enumerate(prs.slides, start=1):
        lines = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                text = "\n".join(p.text.strip() for p in shape.text_frame.paragraphs if p.text.strip())
                if text:
                    lines.append(text)
            if getattr(shape, "has_table", False) and shape.has_table:
                for row in shape.table.rows:
                    lines.append(" | ".join(c.text.strip() for c in row.cells))
        slides.append({"index": i, "text": "\n".join(lines)})
    title = ""
    if prs.slides:
        first = slides[0]["text"].splitlines() if slides else []
        title = first[0] if first else ""
    return {"title": title, "slides": slides}


def parse_pdf(path):
    import fitz

    doc = fitz.open(path)
    pages = [{"index": i, "text": page.get_text("text").strip()} for i, page in enumerate(doc, start=1)]
    first = pages[0]["text"].splitlines() if pages else []
    return {"title": first[0] if first else "", "pages": pages}


def parse(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        return "docx", parse_docx(path)
    if ext == ".pptx":
        return "pptx", parse_pptx(path)
    if ext == ".pdf":
        return "pdf", parse_pdf(path)

    tmpdir = tempfile.mkdtemp(prefix="materials_")
    try:
        if ext == ".doc":
            dst = os.path.join(tmpdir, "converted.docx")
            _word_convert(path, dst, 12)  # wdFormatXMLDocument
            kind, data = "docx", parse_docx(dst)
        elif ext == ".ppt":
            dst = os.path.join(tmpdir, "converted.pptx")
            _ppt_convert(path, dst)
            kind, data = "pptx", parse_pptx(dst)
        else:
            raise ValueError("unsupported file type: %s" % ext)
        return kind, data
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


# ---------------------------------------------------------------- output
def to_digest(item, limit=40000):
    lines = ["===== %s =====" % item["path"]]
    if item.get("title"):
        lines.append("标题: %s" % item["title"])
    if item["kind"] in ("docx", "pdf"):
        body = item.get("paragraphs") or []
        if item["kind"] == "pdf":
            body = []
            for page in item.get("pages", []):
                body.append("--- 第%d页 ---" % page["index"])
                body.append(page["text"])
        if body:
            lines.append("【正文】")
            lines.extend(body)
        for ti, table in enumerate(item.get("tables", []), start=1):
            lines.append("【表格%d】" % ti)
            for row in table:
                lines.append(" | ".join(row))
    else:
        for slide in item.get("slides", []):
            lines.append("--- 第%d页 ---" % slide["index"])
            lines.append(slide["text"])
    text = "\n".join(lines)
    if len(text) > limit:
        text = text[:limit] + "\n...（已截断，完整内容见 files 字段）"
    return text


def collect(paths):
    results = []
    for path in paths:
        path = os.path.abspath(path)
        if not os.path.exists(path):
            results.append({"path": path, "kind": "missing", "error": "file not found"})
            continue
        try:
            kind, data = parse(path)
            item = {"path": path, "kind": kind}
            item.update(data)
            results.append(item)
            print("ok   %s (%s)" % (os.path.basename(path), kind))
        except Exception as exc:  # noqa: BLE001 - report per file, keep going
            results.append({"path": path, "kind": "error", "error": str(exc)})
            print("fail %s: %s" % (os.path.basename(path), exc))
    return results


def expand_inputs(args):
    paths = list(args.files)
    if args.dir:
        for ext in (".docx", ".doc", ".pptx", ".ppt", ".pdf"):
            paths.extend(glob.glob(os.path.join(args.dir, "**", "*" + ext), recursive=True))
    seen, unique = set(), []
    for path in paths:
        if os.path.basename(path).startswith("~$"):
            continue
        key = os.path.abspath(path).lower()
        if key not in seen:
            seen.add(key)
            unique.append(path)
    return unique


def main():
    parser = argparse.ArgumentParser(description="Extract teaching materials to JSON")
    parser.add_argument("files", nargs="*", help="files to extract")
    parser.add_argument("--dir", help="also scan this directory recursively")
    parser.add_argument("--out", required=True, help="output JSON path")
    args = parser.parse_args()

    paths = expand_inputs(args)
    if not paths:
        raise SystemExit("no input files; pass paths or --dir")

    files = collect(paths)
    digest = "\n\n".join(to_digest(item) for item in files if item.get("kind") not in ("missing", "error"))
    payload = {"files": files, "digest": digest}
    with open(args.out, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=2)
    print("written: %s (%d file(s), %d chars)" % (args.out, len(files), len(digest)))


if __name__ == "__main__":
    sys.exit(main())
