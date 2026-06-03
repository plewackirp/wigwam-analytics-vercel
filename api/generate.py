from __future__ import annotations

import cgi
import html
import tempfile
import traceback
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler
from pathlib import Path

from lib.report_generator import generate_report


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        try:
            form = cgi.FieldStorage(
                fp=self.rfile,
                headers=self.headers,
                environ={
                    "REQUEST_METHOD": "POST",
                    "CONTENT_TYPE": self.headers.get("Content-Type", ""),
                },
            )

            with tempfile.TemporaryDirectory() as tmp_name:
                tmp = Path(tmp_name)
                business_path = _save_upload(form, "business_report", tmp)
                inventory_path = _save_upload(form, "inventory_report", tmp)
                gobros_path = _save_upload(form, "gobros_sales", tmp, required=False)
                month_ending = _field_value(form, "month_ending")
                dealer_name = _field_value(form, "dealer_name") or "GoBros"

                filename_label = f"Wigwam Analytics {month_ending or 'generated'}"
                filename = f"{_safe_stem(filename_label)}.xlsx"
                output_path = tmp / filename

                generate_report(
                    business_report=business_path,
                    inventory_report=inventory_path,
                    gobros_sales=gobros_path,
                    output_path=output_path,
                    month_ending=month_ending,
                    dealer_name=dealer_name,
                )

                payload = output_path.read_bytes()

            self.send_response(HTTPStatus.OK)
            self.send_header(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
            self.send_header("Content-Length", str(len(payload)))
            self.end_headers()
            self.wfile.write(payload)
        except Exception as exc:
            body = _error_page(str(exc), traceback.format_exc()).encode("utf-8")
            self.send_response(HTTPStatus.BAD_REQUEST)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)


def _save_upload(form, name: str, tmp: Path, required: bool = True) -> Path | None:
    item = form[name] if name in form else None
    if item is None or not getattr(item, "filename", ""):
        if required:
            raise ValueError(f"Missing required upload: {name.replace('_', ' ')}")
        return None

    target = tmp / Path(item.filename).name
    target.write_bytes(item.file.read())
    return target


def _field_value(form, name: str) -> str:
    if name not in form:
        return ""
    return str(form[name].value).strip()


def _safe_stem(value: str) -> str:
    safe = "".join(ch if ch.isalnum() or ch in {" ", "-", "_", "."} else "-" for ch in value)
    return "-".join(safe.split())[:120] or "wigwam-report"


def _error_page(message: str, detail: str) -> str:
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <title>Could not generate workbook</title>
  <style>
    body {{ font-family: system-ui, sans-serif; margin: 40px; color: #141b18; }}
    pre {{ white-space: pre-wrap; background: #f6f7f3; padding: 16px; border: 1px solid #d7ded8; }}
    a {{ color: #143b2b; font-weight: 700; }}
  </style>
</head>
<body>
  <h1>Could not generate workbook</h1>
  <p>{html.escape(message)}</p>
  <pre>{html.escape(detail)}</pre>
  <p><a href="/">Back to uploader</a></p>
</body>
</html>"""
