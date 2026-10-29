"""Vercel Python function: POST /api/deck {action: read|learn|turn, ...payload} -> JSON (see deck/service.py)."""
import json
import os
import sys
from http.server import BaseHTTPRequestHandler

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from deck import service  # noqa: E402


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        try:
            body = json.loads(self.rfile.read(int(self.headers.get("Content-Length", 0))) or b"{}")
            out, code = service.run(body.pop("action", ""), body), 200
        except Exception as e:  # surface the real error to the panel instead of a bare 500
            out, code = {"error": f"{type(e).__name__}: {e}"}, 500
        data = json.dumps(out, ensure_ascii=False).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)
