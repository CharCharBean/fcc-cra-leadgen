"""One-shot localhost receiver: saves a browser-POSTed body to disk."""
import sys
from http.server import BaseHTTPRequestHandler, HTTPServer

OUT = sys.argv[1]

class H(BaseHTTPRequestHandler):
    def _cors(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "*")
        self.send_header("Access-Control-Allow-Private-Network", "true")

    def do_OPTIONS(self):
        self.send_response(204)
        self._cors()
        self.end_headers()

    def do_POST(self):
        body = self.rfile.read(int(self.headers.get("Content-Length", 0)))
        with open(OUT, "wb") as f:
            f.write(body)
        self.send_response(200)
        self._cors()
        self.end_headers()
        self.wfile.write(b"saved")
        print(f"received {len(body)} bytes -> {OUT}", flush=True)
        raise KeyboardInterrupt  # stop after one successful POST

    def log_message(self, *a):
        pass

try:
    HTTPServer(("127.0.0.1", 8765), H).serve_forever()
except KeyboardInterrupt:
    pass
