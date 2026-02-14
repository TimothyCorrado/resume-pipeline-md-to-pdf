#!/usr/bin/env python3
"""
build_resume_onepage.py
Auto-fit a resume Markdown to ONE PAGE by iteratively:
- generating tight DOCX
- converting to PDF via LibreOffice
- counting PDF pages
- trimming low-priority content if needed

Usage:
  python build_resume_onepage.py resume_helpdesk.md out.docx out.pdf

Requires:
  pip install python-docx pypdf
  LibreOffice installed (soffice.exe)
"""

import re
import sys
import shutil
import subprocess
from pathlib import Path
from typing import List, Tuple

from docx import Document
from docx.shared import Pt, Inches
from pypdf import PdfReader

# ----------------------------
# Config: how to trim if >1 page
# ----------------------------
TRIM_RULES = [
    # Highest impact, lowest downside:
    ("DROP_SECTION", "Other Professional Experience"),
    # Then reduce projects:
    ("KEEP_ONLY_PROJECTS", 2),  # keep first N projects under TECHNICAL PROJECTS
    # Then trim bullets inside sections (keep top bullets)
    ("TRIM_BULLETS_UNDER", ("Technical & IT Support Roles", 3)),
    ("TRIM_BULLETS_UNDER", ("Windows Event Monitoring & Mini SOC Lab", 2)),
    ("TRIM_BULLETS_UNDER", ("pfSense Firewall & Network Segmentation Lab", 2)),
    ("TRIM_BULLETS_UNDER", ("Small Business IT & Security Assessments", 1)),
]

# ----------------------------
# Tight formatting knobs
# ----------------------------
FMT_PRESETS = [
    # Try a normal tight layout first:
    dict(margin=0.50, body_pt=10.0, name_pt=13.0, h2_pt=10.5, h3_pt=10.0,
         h2_before=4, h2_after=1, para_after=0, bullet_after=0),
    # If still 2 pages, tighten slightly:
    dict(margin=0.45, body_pt=9.8, name_pt=12.5, h2_pt=10.2, h3_pt=9.8,
         h2_before=3, h2_after=1, para_after=0, bullet_after=0),
]

BOLD_RE = re.compile(r"\*\*(.+?)\*\*")
H1_RE = re.compile(r"^#\s+(.+)$")
H2_RE = re.compile(r"^##\s+(.+)$")
H3_RE = re.compile(r"^###\s+(.+)$")
BULLET_RE = re.compile(r"^\s*-\s+(.+)$")


def find_soffice() -> str:
    # Common path
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files\LibreOffice\program\soffice.com",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.com",
    ]
    for c in candidates:
        if Path(c).exists():
            return c
    # PATH fallback
    lo = shutil.which("soffice") or shutil.which("soffice.exe") or shutil.which("soffice.com")
    if lo:
        return lo
    raise RuntimeError("LibreOffice not found. Install LibreOffice or add soffice to PATH.")


def pdf_pages(pdf_path: Path) -> int:
    reader = PdfReader(str(pdf_path))
    return len(reader.pages)


def md_read(md_path: Path) -> List[str]:
    return md_path.read_text(encoding="utf-8").splitlines()


def md_write(md_path: Path, lines: List[str]) -> None:
    md_path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")


def drop_section(lines: List[str], section_title: str) -> List[str]:
    out = []
    i = 0
    in_target = False
    while i < len(lines):
        line = lines[i]
        m2 = H3_RE.match(line)
        if m2 and m2.group(1).strip() == section_title:
            in_target = True
            # skip this header
            i += 1
            # skip until next ### or ## header
            while i < len(lines) and not H3_RE.match(lines[i]) and not H2_RE.match(lines[i]):
                i += 1
            continue
        if in_target:
            # once we hit next header, we stop skipping
            in_target = False
        out.append(line)
        i += 1
    return out


def keep_only_projects(lines: List[str], keep_n: int) -> List[str]:
    """
    Under '## TECHNICAL PROJECTS ...' keep only first N '### ' blocks.
    """
    out = []
    i = 0
    in_projects = False
    project_count = 0
    skipping = False

    while i < len(lines):
        line = lines[i]
        if H2_RE.match(line) and "TECHNICAL PROJECTS" in line:
            in_projects = True
            skipping = False
            out.append(line)
            i += 1
            continue

        if in_projects and H2_RE.match(line):
            # next major section ends projects
            in_projects = False
            skipping = False
            out.append(line)
            i += 1
            continue

        if in_projects and H3_RE.match(line):
            project_count += 1
            skipping = project_count > keep_n
            if not skipping:
                out.append(line)
            i += 1
            continue

        if in_projects and skipping:
            i += 1
            continue

        out.append(line)
        i += 1

    return out


def trim_bullets_under(lines: List[str], header_title: str, keep_k: int) -> List[str]:
    """
    Within a specific ### header section, keep only first K bullets.
    """
    out = []
    i = 0
    in_target = False
    kept = 0

    while i < len(lines):
        line = lines[i]
        if H3_RE.match(line):
            title = H3_RE.match(line).group(1).strip()
            in_target = (title == header_title)
            kept = 0
            out.append(line)
            i += 1
            continue

        if in_target:
            bm = BULLET_RE.match(line)
            if bm:
                kept += 1
                if kept <= keep_k:
                    out.append(line)
                i += 1
                continue
            # allow non-bullet lines (tools line etc.)
            out.append(line)
            i += 1
            continue

        out.append(line)
        i += 1

    return out


def normalize_whitespace(lines: List[str]) -> List[str]:
    """
    Remove repeated blank lines; keep at most one blank line between blocks.
    Also remove trailing double-spaces Markdown line breaks that become extra paragraphs.
    """
    cleaned = []
    blank = 0
    for ln in lines:
        ln = ln.rstrip()
        # remove forced md line break markers (two spaces at end)
        ln = re.sub(r"\s{2,}$", "", ln)

        if ln.strip() == "":
            blank += 1
            if blank <= 1:
                cleaned.append("")
        else:
            blank = 0
            cleaned.append(ln)
    return cleaned


def add_text_with_bold(p, text: str):
    last = 0
    for m in BOLD_RE.finditer(text):
        if m.start() > last:
            p.add_run(text[last:m.start()])
        r = p.add_run(m.group(1))
        r.bold = True
        last = m.end()
    if last < len(text):
        p.add_run(text[last:])


def md_to_docx(lines: List[str], docx_path: Path, fmt: dict):
    doc = Document()
    sec = doc.sections[0]

    sec.top_margin = Inches(fmt["margin"])
    sec.bottom_margin = Inches(fmt["margin"])
    sec.left_margin = Inches(fmt["margin"])
    sec.right_margin = Inches(fmt["margin"])

    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(fmt["body_pt"])

    for line in lines:
        line = line.rstrip()
        if not line.strip():
            continue

        if line.startswith("# "):
            p = doc.add_paragraph()
            r = p.add_run(line[2:].strip())
            r.bold = True
            r.font.size = Pt(fmt["name_pt"])
            p.paragraph_format.space_after = Pt(2)
            continue

        if line.startswith("## "):
            p = doc.add_paragraph()
            r = p.add_run(line[3:].strip())
            r.bold = True
            r.font.size = Pt(fmt["h2_pt"])
            p.paragraph_format.space_before = Pt(fmt["h2_before"])
            p.paragraph_format.space_after = Pt(fmt["h2_after"])
            continue

        if line.startswith("### "):
            p = doc.add_paragraph()
            r = p.add_run(line[4:].strip())
            r.bold = True
            r.font.size = Pt(fmt["h3_pt"])
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(0)
            continue

        if line.startswith("- "):
            p = doc.add_paragraph(style="List Bullet")
            add_text_with_bold(p, line[2:].strip())
            p.paragraph_format.space_after = Pt(fmt["bullet_after"])
            continue

        p = doc.add_paragraph()
        add_text_with_bold(p, line.strip())
        p.paragraph_format.space_after = Pt(fmt["para_after"])

    doc.save(str(docx_path))


def docx_to_pdf(docx_path: Path, pdf_path: Path):
    soffice = find_soffice()
    outdir = pdf_path.parent
    outdir.mkdir(parents=True, exist_ok=True)

    cmd = [
        soffice,
        "--headless",
        "--nologo",
        "--nofirststartwizard",
        "--convert-to", "pdf",
        "--outdir", str(outdir),
        str(docx_path),
    ]
    subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    generated = outdir / (docx_path.stem + ".pdf")
    if not generated.exists():
        raise RuntimeError("LibreOffice conversion ran, but PDF was not created.")
    if generated.resolve() != pdf_path.resolve():
        generated.replace(pdf_path)


def apply_trim_rule(lines: List[str], rule: Tuple) -> List[str]:
    kind = rule[0]
    if kind == "DROP_SECTION":
        return drop_section(lines, rule[1])
    if kind == "KEEP_ONLY_PROJECTS":
        return keep_only_projects(lines, rule[1])
    if kind == "TRIM_BULLETS_UNDER":
        title, keep_k = rule[1]
        return trim_bullets_under(lines, title, keep_k)
    return lines


def main():
    if len(sys.argv) != 4:
        print("Usage: python build_resume_onepage.py resume.md out.docx out.pdf")
        sys.exit(2)

    md_path = Path(sys.argv[1]).expanduser().resolve()
    docx_path = Path(sys.argv[2]).expanduser().resolve()
    pdf_path = Path(sys.argv[3]).expanduser().resolve()

    if not md_path.exists():
        print(f"Markdown not found: {md_path}")
        sys.exit(1)

    original_lines = md_read(md_path)
    lines = normalize_whitespace(original_lines)

    # Iteration: try formatting presets first, then trimming rules.
    # We keep the MD file untouched; trimming happens in-memory.
    attempts = 0
    best = None

    for fmt in FMT_PRESETS:
        # Start from normalized base each time
        cur = list(lines)

        # Try without trimming first
        md_to_docx(cur, docx_path, fmt)
        docx_to_pdf(docx_path, pdf_path)
        pages = pdf_pages(pdf_path)
        attempts += 1

        if pages == 1:
            print(f"OK: 1 page (no trimming). Attempts: {attempts}")
            return

        # Apply trim rules one-by-one until 1 page or rules exhausted
        for rule in TRIM_RULES:
            cur = apply_trim_rule(cur, rule)
            cur = normalize_whitespace(cur)

            md_to_docx(cur, docx_path, fmt)
            docx_to_pdf(docx_path, pdf_path)
            pages = pdf_pages(pdf_path)
            attempts += 1

            if pages == 1:
                print(f"OK: 1 page after trimming rule: {rule}. Attempts: {attempts}")
                return

        best = (fmt, cur)

    print("Could not reach 1 page with current rules.")
    print("Last output was generated anyway; consider trimming one more bullet under IT Support Roles.")
    sys.exit(3)


if __name__ == "__main__":
    main()
