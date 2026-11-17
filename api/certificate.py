"""Vercel entry point for certificate PDFs.

The phone number is accepted and ignored here. The browser keeps the
participant list, so phone numbers are not stored on the public site.
"""

import json
import sys
from http.server import BaseHTTPRequestHandler
from pathlib import Path

from fpdf.errors import FPDFUnicodeEncodingException

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from app import build_pdf, pdf_filename, tidy, tidy_phone


class handler(BaseHTTPRequestHandler):
    def _json(self, code, payload):
        body = json.dumps(payload).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self):
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
            tidy_phone(payload.get("phone"))
        except ValueError as exc:
            self._json(400, {"error": str(exc)})
            return
        try:
            pdf = build_pdf(name.upper(), institution.upper(), grade.upper())
        except FPDFUnicodeEncodingException:
            self._json(
                400,
                {
                    "error": "Use English letters, numbers and basic punctuation. This certificate font cannot print other scripts."
                },
            )
            return
        self.send_response(200)
        self.send_header("Content-Type", "application/pdf")
        self.send_header("Content-Length", str(len(pdf)))
        self.send_header(
            "Content-Disposition",
            f'attachment; filename="{pdf_filename(name.upper())}"',
        )
        self.end_headers()
        self.wfile.write(pdf)
