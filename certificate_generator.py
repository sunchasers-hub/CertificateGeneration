"""
certificate_generator.py
Fills the certificate PPTX template with a participant's data and renders
a per-person PDF.

Placeholder replacement works at the *run* level inside each paragraph,
which correctly handles the common case where PowerPoint splits a single
placeholder like "{{Name}}" across multiple runs due to formatting.
"""

import re
import shutil
import subprocess
import platform
from pathlib import Path

from pptx import Presentation

import config
from excel_reader import Participant


def _safe_filename(name: str) -> str:
    cleaned = re.sub(r"[^\w\-. ]", "_", name).strip()
    return cleaned or "participant"


def _build_replacement_pattern(tokens: dict):
    """Builds a single regex that matches either the bare word (e.g. "Name")
    as a whole word, or the double-brace form (e.g. "{{Name}}") — so the
    same code works whether your template uses plain words or {{tokens}}.
    """
    alternatives = []
    for word in tokens:
        alternatives.append(r"\{\{" + re.escape(word) + r"\}\}")
        alternatives.append(r"\b" + re.escape(word) + r"\b")
    return re.compile("|".join(alternatives))


def _resolve_match(matched_text: str, tokens: dict) -> str:
    inner = matched_text
    if inner.startswith("{{") and inner.endswith("}}"):
        inner = inner[2:-2]
    return tokens.get(inner, matched_text)


def _replace_placeholders_in_text_frame(text_frame, replacements: dict, pattern):
    for paragraph in text_frame.paragraphs:
        full_text = "".join(run.text for run in paragraph.runs)
        if not full_text or not pattern.search(full_text):
            continue

        new_text = pattern.sub(lambda m: _resolve_match(m.group(0), replacements), full_text)

        if paragraph.runs:
            # Push the fully-substituted text into the first run (keeps its
            # formatting) and clear the rest — this correctly handles the
            # common case where PowerPoint/Slides split one placeholder
            # word across multiple runs.
            paragraph.runs[0].text = new_text
            for run in paragraph.runs[1:]:
                run.text = ""


def _replace_placeholders_in_slide(slide, replacements: dict, pattern):
    _replace_placeholders_in_shapes(slide.shapes, replacements, pattern)


def _replace_placeholders_in_shapes(shapes, replacements: dict, pattern):
    for shape in shapes:
        if shape.shape_type == 6:  # MSO_SHAPE_TYPE.GROUP
            _replace_placeholders_in_shapes(shape.shapes, replacements, pattern)
            continue
        if shape.has_text_frame:
            _replace_placeholders_in_text_frame(shape.text_frame, replacements, pattern)
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    _replace_placeholders_in_text_frame(cell.text_frame, replacements, pattern)


def fill_certificate_pptx(participant: Participant) -> Path:
    """Creates a filled-in copy of the template PPTX for one participant."""
    config.PPTX_WORK_DIR.mkdir(parents=True, exist_ok=True)

    prs = Presentation(config.PPTX_TEMPLATE_PATH)
    replacements = {
        "Name": participant.name,
        "Distance": participant.distance,
        "Time": participant.time_display,
        "Date": participant.date_display,
    }
    pattern = _build_replacement_pattern(replacements)

    for slide in prs.slides:
        _replace_placeholders_in_slide(slide, replacements, pattern)

    out_path = config.PPTX_WORK_DIR / f"{_safe_filename(participant.name)}_{participant.row_number}.pptx"
    prs.save(out_path)
    return out_path


def convert_pptx_to_pdf_with_powerpoint(pptx_path: Path) -> Path:
    """Converts a .pptx to .pdf using Microsoft PowerPoint via COM automation.

    Windows-only. Requires PowerPoint to be installed and the `pywin32`
    package (`pip install pywin32`).
    """
    import win32com.client  # imported lazily so non-Windows machines don't need pywin32

    config.PDF_DIR.mkdir(parents=True, exist_ok=True)

    # PowerPoint's COM API requires absolute paths.
    pptx_path = pptx_path.resolve()
    pdf_path = (config.PDF_DIR / (pptx_path.stem + ".pdf")).resolve()

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    presentation = None
    try:
        # WithWindow=False keeps PowerPoint from popping a visible window
        # for every conversion.
        presentation = powerpoint.Presentations.Open(
            str(pptx_path), WithWindow=False
        )
        # 32 == ppSaveAsPDF
        presentation.SaveAs(str(pdf_path), 32)
    finally:
        if presentation is not None:
            presentation.Close()
        # Only quit if we have no other presentations open, to avoid
        # closing a PowerPoint instance the user has open for other work.
        try:
            if powerpoint.Presentations.Count == 0:
                powerpoint.Quit()
        except Exception:
            pass

    if not pdf_path.exists():
        raise RuntimeError(f"PowerPoint did not produce a PDF for {pptx_path.name}")
    return pdf_path


def convert_pptx_to_pdf_with_libreoffice(pptx_path: Path) -> Path:
    """Converts a .pptx to .pdf using headless LibreOffice.

    Requires LibreOffice to be installed and the `soffice` binary to be on
    PATH (or config.SOFFICE_BIN pointing to it).
    """
    config.PDF_DIR.mkdir(parents=True, exist_ok=True)

    result = subprocess.run(
        [
            config.SOFFICE_BIN,
            "--headless",
            "--norestore",
            "--convert-to", "pdf",
            "--outdir", str(config.PDF_DIR),
            str(pptx_path),
        ],
        capture_output=True,
        text=True,
        timeout=120,
    )

    expected_pdf = config.PDF_DIR / (pptx_path.stem + ".pdf")
    if result.returncode != 0 or not expected_pdf.exists():
        raise RuntimeError(
            f"LibreOffice conversion failed for {pptx_path.name}\n"
            f"stdout: {result.stdout}\nstderr: {result.stderr}"
        )
    return expected_pdf


def convert_pptx_to_pdf(pptx_path: Path) -> Path:
    """Converts a filled .pptx into a PDF, choosing the conversion backend
    based on config.CONVERSION_METHOD."""
    method = config.CONVERSION_METHOD

    if method == "powerpoint":
        return convert_pptx_to_pdf_with_powerpoint(pptx_path)

    if method == "libreoffice":
        return convert_pptx_to_pdf_with_libreoffice(pptx_path)

    # "auto": prefer PowerPoint on Windows, fall back to LibreOffice
    if platform.system() == "Windows":
        try:
            return convert_pptx_to_pdf_with_powerpoint(pptx_path)
        except Exception:
            pass
    return convert_pptx_to_pdf_with_libreoffice(pptx_path)


def generate_certificate_pdf(participant: Participant) -> Path:
    """End-to-end: fill template -> convert to PDF -> return PDF path."""
    filled_pptx = fill_certificate_pptx(participant)
    pdf_path = convert_pptx_to_pdf(filled_pptx)
    return pdf_path
