"""Local helper for the Job Tracker Viewer: lets you edit the tracker from the viewer, so you never open Excel.

Usage:  python tracker_server.py <tracker.xlsx> [--port 8765] [--no-browser]

It serves the viewer (Job Tracker Viewer.html, next to this file) at http://127.0.0.1:<port>/ and saves your edits into the
workbook through tracker_tool.py (the same checks the modules use: it refuses while the workbook is open in Excel, keeps all
formatting, and never overwrites a save made by someone else in the meantime).

Only this computer can reach it (127.0.0.1), and every request needs the one-time token in the link it prints. Stop it with
Ctrl+C. Without this helper the viewer still opens on its own as a read-only page.
"""
import argparse, contextlib, io, json, os, re, secrets, sys, threading, time, webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import tracker_tool  # noqa: E402

VIEWER = os.path.join(HERE, "Job Tracker Viewer.html")
WRITE_LOCK = threading.Lock()
SEEN_LOCK = threading.Lock()


def seen_path():
    """Job IDs you have opened or acted on, kept next to the workbook (survives a different browser, port or cleared site data)."""
    return os.path.splitext(STATE["path"])[0] + ".seen.json"


def read_seen():
    try:
        with open(seen_path(), encoding="utf-8") as f:
            return [str(x) for x in json.load(f).get("seen", [])]
    except (OSError, ValueError, AttributeError):
        return None


def read_base():
    try:
        with open(seen_path(), encoding="utf-8") as f:
            return {str(k): str(v) for k, v in (json.load(f).get("base") or {}).items()}
    except (OSError, ValueError, AttributeError):
        return {}


def add_seen(ids, base=None):
    with SEEN_LOCK:
        cur_base = read_base()
        for k, v in (base or {}).items():
            cur_base.setdefault(k, v)  # the status at first sight is never overwritten
        merged = sorted(set(read_seen() or []) | set(ids))
        tmp = seen_path() + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump({"note": "Job IDs opened or acted on in the tracker viewer. Safe to delete: every job then shows as unseen again.", "seen": merged, "base": cur_base}, f)
        os.replace(tmp, seen_path())
        return len(merged)
STATE = {}


def run_tool(argv):
    """Run a tracker_tool command; returns (ok, message)."""
    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf):
            code = tracker_tool.main(argv)
    except SystemExit:
        return False, "That request was not understood."
    except Exception as e:  # never leave the page hanging
        return False, f"Unexpected error: {e}"
    msg = buf.getvalue().strip()
    return code == 0, msg[len("REFUSED: "):] if msg.startswith("REFUSED: ") else msg


class Handler(BaseHTTPRequestHandler):
    server_version = "TrackerViewer"

    def log_message(self, fmt, *args):  # keep the terminal quiet
        pass

    # ---- helpers
    def _host_ok(self):
        port = self.server.server_address[1]
        return self.headers.get("Host", "") in (f"127.0.0.1:{port}", f"localhost:{port}")

    def _token_ok(self, query=None):
        tok = STATE["token"]
        got = self.headers.get("X-Token") or ((query or {}).get("t") or [""])[0]
        return secrets.compare_digest(got, tok)

    def _send(self, code, body=b"", ctype="text/plain; charset=utf-8", extra=None):
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def _json(self, code, obj):
        self._send(code, json.dumps(obj), "application/json; charset=utf-8")

    def _deny(self):
        self._send(403, "Forbidden. Open the link printed in the terminal window where the helper is running.")

    # ---- routes
    def do_GET(self):
        if not self._host_ok():
            return self._deny()
        u = urlparse(self.path)
        q = parse_qs(u.query)
        if u.path == "/":
            if not self._token_ok(q):
                return self._deny()
            html = open(VIEWER, encoding="utf-8").read()
            boot = "<script>window.TRACKER_API=" + json.dumps({"token": STATE["token"], "name": os.path.basename(STATE["path"])}) + ";</script>"
            html = html.replace("</head>", boot + "</head>", 1)
            return self._send(200, html, "text/html; charset=utf-8")
        if u.path == "/api/seen":
            if not self._token_ok():
                return self._deny()
            return self._json(200, {"seen": read_seen(), "base": read_base() if os.path.exists(seen_path()) else None})
        if u.path in ("/api/stat", "/api/file"):
            if not self._token_ok():
                return self._deny()
            path = STATE["path"]
            try:
                if u.path == "/api/stat":
                    st = os.stat(path)
                    return self._json(200, {"mtime": st.st_mtime, "size": st.st_size, "name": os.path.basename(path)})
                for attempt in range(4):  # the file can be mid-save for a moment
                    try:
                        with open(path, "rb") as f:
                            data = f.read()
                        mt = os.path.getmtime(path)
                        break
                    except OSError:
                        if attempt == 3:
                            raise
                        time.sleep(0.3)
                return self._send(200, data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", {"X-Mtime": repr(mt)})
            except OSError as e:
                return self._json(503, {"ok": False, "message": f"Could not read the workbook: {e}"})
        self._send(404, "Not found")

    def do_POST(self):
        if not self._host_ok() or not self._token_ok():
            return self._deny()
        origin = self.headers.get("Origin")
        port = self.server.server_address[1]
        if origin and origin not in (f"http://127.0.0.1:{port}", f"http://localhost:{port}"):
            return self._deny()
        route = urlparse(self.path).path
        if route == "/api/seen":
            try:
                n = int(self.headers.get("Content-Length", "0"))
                if n <= 0 or n > 1_000_000:
                    raise ValueError
                body = json.loads(self.rfile.read(n).decode("utf-8"))
                ids = body.get("add") or []
                base_in = body.get("base") or {}
                if not isinstance(base_in, dict) or len(base_in) > 5000 or not all(re.match(r"^JOB-\d+$", str(k)) and isinstance(v, str) and len(v) <= 40 for k, v in base_in.items()):
                    raise ValueError
                if not isinstance(ids, list) or len(ids) > 5000 or not all(isinstance(i, str) and re.match(r"^JOB-\d+$", i) for i in ids):
                    raise ValueError
            except Exception:
                return self._json(400, {"ok": False, "message": "That request was not understood."})
            try:
                return self._json(200, {"ok": True, "count": add_seen(ids, base_in)})
            except OSError as e:
                return self._json(500, {"ok": False, "message": f"Could not save the seen list: {e}"})
        if route != "/api/edit":
            return self._send(404, "Not found")
        try:
            n = int(self.headers.get("Content-Length", "0"))
            if n <= 0 or n > 1_000_000:
                raise ValueError
            req = json.loads(self.rfile.read(n).decode("utf-8"))
            job = str(req["job"]).strip().upper()
            if not re.match(r"^JOB-\d+$", job):
                raise ValueError
            sets = req.get("set") or {}
            if not isinstance(sets, dict):
                raise ValueError
        except Exception:
            return self._json(400, {"ok": False, "message": "That request was not understood."})
        argv = ["edit", STATE["path"], job]
        for k, v in sets.items():
            argv += ["--set", f"{k}={v}"]
        if str(req.get("note") or "").strip():
            argv += ["--note", str(req["note"])]
        if str(req.get("reason") or "").strip():
            argv += ["--reason", str(req["reason"])]
        if req.get("mtime"):
            argv += ["--if-unchanged-since", str(float(req["mtime"]))]
        with WRITE_LOCK:
            ok, msg = run_tool(argv)
            mt = os.path.getmtime(STATE["path"]) if os.path.exists(STATE["path"]) else 0
        self._json(200, {"ok": ok, "message": msg, "mtime": mt})


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("tracker")
    ap.add_argument("--port", type=int, default=8765)
    ap.add_argument("--no-browser", action="store_true")
    a = ap.parse_args()
    path = os.path.abspath(a.tracker.strip().strip('"'))
    if not os.path.isfile(path):
        sys.exit(f"File not found: {path}")
    if not os.path.isfile(VIEWER):
        sys.exit(f"Could not find the viewer next to this script: {VIEWER}")
    STATE.update(path=path, token=secrets.token_urlsafe(16))
    srv = None
    for port in range(a.port, a.port + 20):
        try:
            srv = ThreadingHTTPServer(("127.0.0.1", port), Handler)
            break
        except OSError:
            continue
    if srv is None:
        sys.exit("Could not open a local port. Close other copies of the helper and try again.")
    url = f"http://127.0.0.1:{srv.server_address[1]}/?t={STATE['token']}"
    print("Tracker viewer is running for:", path)
    print("Open this link (it opens automatically):", url)
    print("Keep this window open while you work. Press Ctrl+C to stop.")
    if not a.no_browser and not os.environ.get("TRACKER_NO_BROWSER"):
        threading.Timer(0.5, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")


if __name__ == "__main__":
    main()
