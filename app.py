#!/usr/bin/env python3
"""Kites Robo Challenge certificate desk.

Serves a small dashboard and builds an A4 PDF on the official Kites
template. Body text uses the same font, size and line positions as
certificate.pdf.
"""

import json
import sys
import threading
import webbrowser
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

from fpdf import FPDF
from fpdf.errors import FPDFUnicodeEncodingException
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

ROOT = Path(__file__).resolve().parent
TEMPLATES = {
    "participation": ROOT / "public" / "template.jpg",
    "winner": ROOT / "public" / "winner.jpg",
}
KIND_LABELS = {
    "participation": "Participation",
    "winner": "Winner",
}
INDEX = ROOT / "public" / "index.html"
EXCEL = ROOT / "Kites-Robo-Challenge.xlsx"
EXCEL_LOCK = threading.Lock()
EXCEL_HEADERS = (
    "Participant name",
    "Institution name",
    "Grade",
    "Phone number",
    "Added on",
    "Certificate",
)

# Page geometry copied from certificate.pdf (FPDF, A4, points).
PAGE_W = 595.28
PAGE_H = 841.89
IMG_X = 28.35
IMG_Y_BOTTOM = 51.76
IMG_W = 538.58
IMG_H = 761.79

# The body sits in the open area under the title and above the signatures.
# Positions are distances from the top of the page.
BLOCK_TOP = 330.0
BLOCK_BOTTOM = 545.0
FONT_SIZE = 13.0
MAX_WIDTH = 460.0
MIN_SIZE = 9.0
LEADING = 1.55
PARAGRAPH_GAP = 12.0


def tidy(value, label, limit):
    if not isinstance(value, str):
        raise ValueError(f"Enter the {label}.")
    text = " ".join(value.split())
    if not text:
        raise ValueError(f"Enter the {label}.")
    if len(text) > limit:
        raise ValueError(f"The {label} is too long. Keep it under {limit} characters.")
    if any(ord(ch) < 32 for ch in text):
        raise ValueError(f"The {label} has characters that cannot be printed.")
    return text


def certificate_kind(value):
    if value in (None, "", "participation"):
        return "participation"
    if value == "winner":
        return "winner"
    raise ValueError("Choose a participation certificate or a winner certificate.")


def tidy_phone(value):
    if value is None:
        return ""
    if not isinstance(value, str):
        raise ValueError("Enter the phone number using digits.")
    text = " ".join(value.split())
    if not text:
        return ""
    if len(text) > 20:
        raise ValueError("The phone number is too long. Keep it under 20 characters.")
    if any(ch not in "0123456789+ -()" for ch in text):
        raise ValueError("Enter the phone number using digits.")
    return text


def append_participant(name, institution, grade, phone, certificate):
    """Create the workbook on the first entry, then append one row."""
    with EXCEL_LOCK:
        if EXCEL.exists():
            workbook = load_workbook(EXCEL)
            sheet = workbook.active
            if sheet.cell(1, 6).value != "Certificate":
                cell = sheet.cell(1, 6, "Certificate")
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="B43224")
                cell.alignment = Alignment(vertical="center")
                sheet.column_dimensions["F"].width = 18
        else:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Participants"
            sheet.append(list(EXCEL_HEADERS))
            for cell in sheet[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="B43224")
                cell.alignment = Alignment(vertical="center")
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = "A1:F1"
            for column, width in zip("ABCDEF", (28, 42, 40, 18, 22, 18)):
                sheet.column_dimensions[column].width = width

        sheet.append([
            name,
            institution,
            grade,
            phone,
            datetime.now().strftime("%d %b %Y, %I:%M %p"),
            certificate,
        ])
        sheet.cell(sheet.max_row, 4).number_format = "@"
        sheet.auto_filter.ref = f"A1:F{sheet.max_row}"
        temporary = EXCEL.with_suffix(".tmp.xlsx")
        workbook.save(temporary)
        temporary.replace(EXCEL)
        return sheet.max_row - 1


def certificate_paragraphs(name, institution, grade, kind):
    """Fixed wording. Only the name, institution and grade are filled in."""
    if kind == "winner":
        achievement = "is a winner of"
        closing = "outstanding performance"
    else:
        achievement = "has successfully participated in"
        closing = "active participation"
    return (
        (
            f"This is to certify that {name}, a student of {institution}, "
            f"studying in {grade}, {achievement} the "
            "KITES ROBO CHALLENGE held on 17th November 2026."
        ),
        (
            "We appreciate his enthusiasm, "
            f"{closing}, and interest in robotics and innovation."
        ),
    )


def wrap_line(pdf, text, width):
    words = text.split()
    lines = []
    current = ""
    for word in words:
        trial = word if not current else f"{current} {word}"
        if not current or pdf.get_string_width(trial) <= width:
            current = trial
        else:
            lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines


def layout_paragraphs(pdf, paragraphs, size, force=False):
    pdf.set_font("Times", "BI", size)
    blocks = [wrap_line(pdf, text, MAX_WIDTH) for text in paragraphs]
    too_wide = any(
        pdf.get_string_width(line) > MAX_WIDTH + 0.5
        for block in blocks
        for line in block
    )
    line_gap = size * LEADING
    line_count = sum(len(block) for block in blocks)
    span = (line_count - 1) * line_gap + PARAGRAPH_GAP * (len(blocks) - 1)
    if not force and (too_wide or BLOCK_TOP + span > BLOCK_BOTTOM):
        return None
    spare = max(0.0, BLOCK_BOTTOM - BLOCK_TOP - span)
    y = BLOCK_TOP + spare / 2.0
    placed = []
    for index, block in enumerate(blocks):
        if index:
            y += PARAGRAPH_GAP
        for line in block:
            placed.append((line, y))
            y += line_gap
    return placed


def pdf_filename(name, kind):
    slug = "".join(ch if ch.isalnum() else "-" for ch in name)
    slug = "-".join(part for part in slug.split("-") if part)
    role = "Winner" if kind == "winner" else "Participation"
    return f"Kites-{role}-{slug or 'Participant'}.pdf"


def build_pdf(name, institution, grade, kind="participation"):
    pdf = FPDF(unit="pt", format=(PAGE_W, PAGE_H))
    pdf.set_auto_page_break(False)
    pdf.add_page()
    img_top = PAGE_H - (IMG_Y_BOTTOM + IMG_H)
    pdf.image(str(TEMPLATES[kind]), x=IMG_X, y=img_top, w=IMG_W, h=IMG_H)
    pdf.set_text_color(0, 0, 0)

    paragraphs = certificate_paragraphs(name, institution, grade, kind)
    size = FONT_SIZE
    placed = None
    while size >= MIN_SIZE:
        placed = layout_paragraphs(pdf, paragraphs, size)
        if placed is not None:
            break
        size -= 0.25
    if placed is None:
        size = MIN_SIZE
        placed = layout_paragraphs(pdf, paragraphs, size, force=True)
    pdf.set_font("Times", "BI", size)
    for text, top in placed:
        width = pdf.get_string_width(text)
        pdf.text((PAGE_W - width) / 2.0, top, text)

    return bytes(pdf.output())


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        sys.stderr.write("[%s] %s\n" % (self.log_date_time_string(), fmt % args))

    def _send(self, code, body, content_type, extra=None):
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        if extra:
            for key, value in extra.items():
                self.send_header(key, value)
        self.end_headers()
        self.wfile.write(body)

    def _json(self, code, payload):
        self._send(code, json.dumps(payload).encode("utf-8"), "application/json")

    def do_GET(self):
        path = urlparse(self.path).path
        if path in ("/", "/index.html"):
            self._send(200, INDEX.read_bytes(), "text/html; charset=utf-8")
            return
        if path == "/template.jpg":
            self._send(200, TEMPLATES["participation"].read_bytes(), "image/jpeg")
            return
        if path == "/winner.jpg":
            self._send(200, TEMPLATES["winner"].read_bytes(), "image/jpeg")
            return
        if path == "/participants.xlsx":
            if not EXCEL.is_file():
                self._json(404, {"error": "The participant list is created when the first certificate is downloaded."})
                return
            self._send(
                200,
                EXCEL.read_bytes(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                {"Content-Disposition": 'attachment; filename="Kites-Robo-Challenge.xlsx"'},
            )
            return
        self._send(404, b"Not found", "text/plain; charset=utf-8")

    def do_POST(self):
        if urlparse(self.path).path != "/api/certificate":
            self._send(404, b"Not found", "text/plain; charset=utf-8")
            return

        length = int(self.headers.get("Content-Length", "0") or 0)
        if length <= 0 or length > 20000:
            self._json(400, {"error": "Could not read the form."})
            return

        try:
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError):
            self._json(400, {"error": "Could not read the form."})
            return

        try:
            name = tidy(payload.get("name"), "participant name", 60)
            institution = tidy(payload.get("institution"), "institution name", 80)
            grade = tidy(payload.get("grade"), "grade", 90)
            phone = tidy_phone(payload.get("phone"))
            kind = certificate_kind(payload.get("kind"))
        except ValueError as exc:
            self._json(400, {"error": str(exc)})
            return

        try:
            pdf = build_pdf(name, institution, grade, kind)
        except FPDFUnicodeEncodingException:
            self._json(
                400,
                {
                    "error": "Use English letters, numbers and basic punctuation. This certificate font cannot print other scripts."
                },
            )
            return

        try:
            count = append_participant(name, institution, grade, phone, KIND_LABELS[kind])
        except Exception:
            self._json(
                500,
                {"error": "The participant list could not be saved. Close Kites-Robo-Challenge.xlsx if it is open, then try again."},
            )
            return

        self._send(
            200,
            pdf,
            "application/pdf",
            {
                "Content-Disposition": f'attachment; filename="{pdf_filename(name, kind)}"',
                "X-Participant-Count": str(count),
            },
        )


def main():
    if not INDEX.is_file() or any(not path.is_file() for path in TEMPLATES.values()):
        sys.exit("Missing a certificate template or the dashboard page.")

    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8765
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    url = f"http://127.0.0.1:{port}"
    print(f"Certificate desk is running at {url}", flush=True)
    print("Leave this window open while you generate certificates.", flush=True)
    webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")


if __name__ == "__main__":
    main()
