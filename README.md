# Resume Pipeline: Markdown → 1-Page DOCX → PDF

A lightweight Python automation tool that converts a Markdown resume into:

- ATS-friendly DOCX
- Locked-layout PDF
- Enforced single-page formatting

## Why

Maintaining resumes manually in Word leads to:
- Layout drift
- Inconsistent spacing
- Multi-page overflow
- Versioning issues

This tool uses Markdown as the source of truth and generates a clean, repeatable output.

## Features

- Tight margin enforcement
- Font size normalization
- Bullet spacing control
- Automatic trimming if PDF exceeds one page
- LibreOffice headless PDF conversion

## Requirements

- Python 3.10+
- python-docx
- pypdf
- LibreOffice

Install dependencies:

```bash
pip install python-docx pypdf

Usage:

python build_resume_onepage.py resume.md output.docx output.pdf

Outputs:

output.docx

output.pdf