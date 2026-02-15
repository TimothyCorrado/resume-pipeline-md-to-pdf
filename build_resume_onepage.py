#!/usr/bin/env python3
"""
build_resume_onepage.py

Auto-fit a resume Markdown to ONE PAGE by iteratively:
- generating tight DOCX
- converting to PDF via LibreOffice
- counting PDF pages
- trimming low-priority content if needed

Adds:
- Email auto-linking (mailto:) for FULL email address
- URL auto-linking (LinkedIn/GitHub/etc.)
- Proper parsing of bold inside H3 (###) lines (no literal **)
- Ignores Markdown horizontal rules (---)
- Trim rules resilient to **bold** inside H3 titles

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
from typing import List, Tuple, Optional

from docx import Document
from docx.shared import Pt, Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from pypdf import PdfReader

# ----------------------------
# Config: how to trim if >1 page
# ----------------------------
TRIM_RULES = [
    ("DROP_SECTION", "Other Professional Experience"),
    ("KEEP_ONLY_PROJECTS", 2),
    ("TRIM_BULLETS_UNDER", ("Technical & IT Support Roles", 3)),
    ("TRIM_BULLETS_UNDER", ("Windows Event Monitoring & Mini SOC Lab", 2)),
    ("TRIM_BULLETS_UNDER", ("pfSense Firewall & Network Segmentation Lab", 2)),
    ("TRIM_BULLETS_UNDER", ("Small Business IT & Security Assessments", 1)),
]

# ----------------------------
# Tight formatting knobs
# ----------------------------
FMT_PRESETS = [
    dict(
        margin=0.50,
        body_pt=10.0,
        name_pt=13.0,
        h2_pt=10.5,
        h3_pt=10.0,
        h2_before=4,
        h2_after=1,
        para_after=0,
        bullet_after=0,
    ),
    dict(
        margin=0.45,
        body_pt=9.8,
        name_pt=12.5,
        h2_pt=10.2,
        h3_pt=9.8,
        h2_before=3,
        h2_after=1,
        para_after=0,
        bullet_after=0,
    ),
]

# Markdown / parsing helpers
BOLD_RE = re.compile(r"\*\*(.+?)\*\*")
H2_RE = re.compile(r"^##\s+(.+)$")
H3_RE = re.compile(r"^###\s+(.+)$")
BULLET_RE = re.compile(r"^\s*-\s+(.+)$")
HR_RE = re.compile(r"^\s*-{3,}\s*$")  # --- horizontal rules

# EMAIL: match whole email token (bounded, not partial)
EMAIL_RE = re.compile(r"(?P<email>[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})")

# URL: match whole URL-like token; stop before common separators
# - handles https://..., http://..., and bare domains like linkedin.com/in/...
URL_RE = re.compile(
    r"(?P<url>(?:https?://)?(?:www\.)?[A-Za-z0-9.-]+\.[A-Za-z]{2,}(?:/[^\s|,)]+)?)"
)

def strip_md_bold(s: str) -> str:
    return re.sub(r"\*\*(.+?)\*\*", r"\1", s).strip()

def normalize_url(url: str) -> str:
    u = url.strip()
    if u.startswith("http://") or u.startswith("https://"):
        return u
    return "https://" + u

def find_soffice() -> str:
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files\LibreOffice\program\soffice.com",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.com",
    ]
    for c in candidates:
        if Path(c).exists():
            return c
    lo = shutil.which("soffice") or shutil.which("soffice.exe") or shutil.which("soffice.com")
    if lo:
        return lo
    raise RuntimeError("LibreOffice not found. Install LibreOffice or add soffice to PATH.")

def pdf_pages(pdf_path: Path) -> int:
    reader = PdfReader(str(pdf_path))
    return len(reader.pages)

def md_read(md_path: Path) -> List[str]:
    return md_path.read_text(encoding="utf-8").splitlines()

def drop_section(lines: List[str], section_title: str) -> List[str]:
    out: List[str] = []
    i = 0
    while i < len(lines):
        line = lines[i]
        m3 = H3_RE.match(line)
        if m3 and strip_md_bold(m3.group(1)) == section_title:
            i += 1
            while i < len(lines) and not H3_RE.match(lines[i]) and not H2_RE.match(lines[i]):
                i += 1
            continue
        out.append(line)
        i += 1
    return out

def keep_only_projects(lines: List[str], keep_n: int) -> List[str]:
    out: List[str] = []
    i = 0
    in_projects = False
    project_count = 0
    skipping = False

    while i < len(lines):
        line = lines[i]
        if H2_RE.match(line) and "TECHNICAL PROJECT" in line.upper():
            in_projects = True
            skipping = False
            out.append(line)
            i += 1
            continue

        if in_projects and H2_RE.match(line):
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
    out: List[str] = []
    i = 0
    in_target = False
    kept = 0

    while i < len(lines):
        line = lines[i]
        if H3_RE.match(line):
            title = strip_md_bold(H3_RE.match(line).group(1))
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
            out.append(line)
            i += 1
            continue

        out.append(line)
        i += 1

    return out

def normalize_whitespace(lines: List[str]) -> List[str]:
    cleaned: List[str] = []
    blank = 0
    for ln in lines:
        ln = ln.rstrip()
        ln = re.sub(r"\s{2,}$", "", ln)
        if HR_RE.match(ln):
            continue
        if ln.strip() == "":
            blank += 1
            if blank <= 1:
                cleaned.append("")
        else:
            blank = 0
            cleaned.append(ln)
    return cleaned

def add_hyperlink(paragraph, url: str, text: str, *, bold: bool = False, font_size: Optional[Pt] = None):
    part = paragraph.part
    r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")

    if bold:
        b = OxmlElement("w:b")
        rPr.append(b)

    if font_size is not None:
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), str(int(font_size.pt * 2)))
        rPr.append(sz)

    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    rPr.append(u)

    new_run.append(rPr)

    t = OxmlElement("w:t")
    t.text = text
    new_run.append(t)

    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)

def add_text_with_bold_links_emails(p, text: str, *, default_font_size: Optional[Pt] = None, force_bold: bool = False):
    """
    Adds runs to paragraph p supporting:
    - **bold** segments
    - FULL email linking (mailto:)
    - URL linking
    Priority: EMAIL matches first, then URL matches.
    """
    # Split into bold vs non-bold segments first
    segments = []
    last = 0
    for m in BOLD_RE.finditer(text):
        if m.start() > last:
            segments.append((False, text[last:m.start()]))
        segments.append((True, m.group(1)))
        last = m.end()
    if last < len(text):
        segments.append((False, text[last:]))

    def add_plain(s: str, is_bold: bool):
        if not s:
            return
        run = p.add_run(s)
        run.bold = force_bold or is_bold
        if default_font_size is not None:
            run.font.size = default_font_size

    for is_bold_seg, seg in segments:
        seg = seg or ""
        i = 0
        while i < len(seg):
            # Find next email or url
            em = EMAIL_RE.search(seg, i)
            um = URL_RE.search(seg, i)

            # Pick the earliest match (email wins ties)
            next_m = None
            kind = None
            if em and um:
                if em.start() <= um.start():
                    next_m, kind = em, "email"
                else:
                    next_m, kind = um, "url"
            elif em:
                next_m, kind = em, "email"
            elif um:
                next_m, kind = um, "url"

            if not next_m:
                add_plain(seg[i:], is_bold_seg)
                break

            # Add leading text before match
            if next_m.start() > i:
                add_plain(seg[i:next_m.start()], is_bold_seg)

            token = next_m.group(0)

            if kind == "email":
                email = next_m.group("email")
                add_hyperlink(
                    p,
                    f"mailto:{email}",
                    email,
                    bold=(force_bold or is_bold_seg),
                    font_size=default_font_size,
                )
            else:
                raw_url = next_m.group("url")
                # Avoid linking emails as URLs (already handled), and ensure it's a plausible url token
                if "@" in raw_url:
                    add_plain(raw_url, is_bold_seg)
                else:
                    url = normalize_url(raw_url)
                    add_hyperlink(
                        p,
                        url,
                        raw_url,
                        bold=(force_bold or is_bold_seg),
                        font_size=default_font_size,
                    )

            i = next_m.end()

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
        if HR_RE.match(line):
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
            title = line[4:].strip()
            add_text_with_bold_links_emails(
                p, title, default_font_size=Pt(fmt["h3_pt"]), force_bold=True
            )
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(0)
            continue

        if line.startswith("- "):
            p = doc.add_paragraph(style="List Bullet")
            add_text_with_bold_links_emails(p, line[2:].strip())
            p.paragraph_format.space_after = Pt(fmt["bullet_after"])
            continue

        p = doc.add_paragraph()
        add_text_with_bold_links_emails(p, line.strip())
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

    lines = normalize_whitespace(md_read(md_path))

    attempts = 0

    for fmt in FMT_PRESETS:
        cur = list(lines)

        md_to_docx(cur, docx_path, fmt)
        docx_to_pdf(docx_path, pdf_path)
        pages = pdf_pages(pdf_path)
        attempts += 1

        if pages == 1:
            print(f"OK: 1 page (no trimming). Attempts: {attempts}")
            return

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

    print("Could not reach 1 page with current rules.")
    print("Last output was generated anyway; consider trimming one more bullet under IT Support Roles.")
    sys.exit(3)

if __name__ == "__main__":
    main()
