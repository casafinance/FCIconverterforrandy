"""
Vercel serverless function: POST /api/convert
Body: raw PDF bytes
Header X-Filename: original filename (used for the output download name)
Response: 200 + .xlsx bytes, or 4xx/5xx + plain-text error message
"""
import os
import shutil
import sys
import tempfile
import traceback
from http.server import BaseHTTPRequestHandler

# Make the local deposit_to_excel module importable
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from deposit_to_excel import convert

# Free-tier Vercel allows a 4.5 MB request body. Cap slightly under that.
MAX_BODY_BYTES = 4 * 1024 * 1024  # 4 MB


class handler(BaseHTTPRequestHandler):
    # ----- helpers -----
    def _send_text(self, code, message):
        body = message.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_xlsx(self, xlsx_bytes, download_name):
        self.send_response(200)
        self.send_header(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        self.send_header(
            "Content-Disposition",
            f'attachment; filename="{download_name}"',
        )
        self.send_header("Content-Length", str(len(xlsx_bytes)))
        self.end_headers()
        self.wfile.write(xlsx_bytes)

    # ----- routes -----
    def do_GET(self):
        # Friendly response so visiting /api/convert in a browser
        # doesn't look broken.
        self._send_text(
            200,
            "FCI Deposit Converter API. POST a PDF as the request body to convert it.",
        )

    def do_POST(self):
        try:
            content_length = int(self.headers.get("Content-Length") or 0)
            if content_length <= 0:
                self._send_text(400, "No file in request body.")
                return
            if content_length > MAX_BODY_BYTES:
                self._send_text(
                    413,
                    f"File too large ({content_length} bytes). "
                    f"Max is {MAX_BODY_BYTES} bytes on the free Vercel tier.",
                )
                return

            body = self.rfile.read(content_length)
            if not body.startswith(b"%PDF-"):
                self._send_text(400, "Uploaded file does not look like a PDF.")
                return

            # Build a friendly download name from the X-Filename header,
            # stripping any path and the .pdf extension.
            original = self.headers.get("X-Filename") or "deposit.pdf"
            stem = os.path.basename(original).rsplit(".", 1)[0] or "deposit"
            download_name = f"{stem}.xlsx"

            tmpdir = tempfile.mkdtemp(prefix="fci_")
            try:
                pdf_path = os.path.join(tmpdir, "input.pdf")
                with open(pdf_path, "wb") as f:
                    f.write(body)

                out_path = convert(pdf_path, output_dir=tmpdir)

                with open(out_path, "rb") as f:
                    xlsx_bytes = f.read()
            finally:
                shutil.rmtree(tmpdir, ignore_errors=True)

            self._send_xlsx(xlsx_bytes, download_name)

        except Exception as exc:
            # Surface a short error message to the client; print the full
            # traceback to Vercel's logs for debugging.
            traceback.print_exc()
            self._send_text(500, f"Conversion failed: {type(exc).__name__}: {exc}")
