#!/usr/bin/env python3
"""
Native messaging host for the "Outlook Web Helper" add-on.

The host has two faces:

  1. stdin/stdout: Mozilla native messaging, talking to the Firefox add-on.
     Each message is preceded by a 4-byte unsigned int (native byte order)
     that gives the byte length of the following UTF-8 JSON payload.

  2. A small HTTP server on 127.0.0.1 only, talking to scripts of the user.

HTTP API
--------
Every request needs the header "Authorization: Bearer <adminToken>" and must
come from the loopback interface. The token is created on the first start and
kept in the config file. The DOM of a logged-in Outlook page may hold personal
data, it must never leave this computer.

  GET  /status     -> {"ok": true, "addon": true, "hostVersion": ..., "build": <12 hex digits>}
                   build is the start of the SHA-256 of this file
  GET  /tabs       the open Outlook tabs
  The calls below all need an integer "tabId" from GET /tabs. Without it the
  answer is 400 with "code": "TAB_REQUIRED", and the add-on is not asked.
  POST /dom        body {"tabId", "selector"?, "mode"?, "maxBytes"?, "limit"?,
                         "waitFor"?, "timeout"?, "strip"?}
                   the DOM of an Outlook tab (see content script)
  POST /act        body {"tabId", "action": "click"|"type"|"press", "selector",
                         "index"?, "text"?, "key"?, "timeout"?}
                   click, type into or press a key on an element of the tab
  POST /navigate   body {"tabId", "url"}  load an Outlook page
                   in the tab and wait until it has loaded
  A tabId that is not an open outlook.live.com tab (any more) gives 409 with
  "code": "INVALID_TAB" for /dom, /act, /navigate, /back and /reload.
  POST /back       body {"tabId"}  go one step back in the history of the tab
                   -> {"ok": true, "url": <new address of the tab>}
  POST /focus      body {"tabId"}  bring the tab and its window to the front. The page
                   does not react to clicks while its tab is hidden.
  POST /reload     body {"tabId"}  reload the add-on (it re-reads its files, the helper
                   process restarts with it)

Messages add-on -> host:
  { "requestId", "cmd": "getInfo" }
  { "requestId", "cmd": "setConfig", "regenerateAdminToken"?: bool }
  { "requestId", "cmd": "result", "result": { ... } }   answer to a forwarded request

Messages host -> add-on:
  { "requestId", "ok": true, "result": any }            reply to a command
  { "requestId", "ok": false, "error": str }
  { "requestId", "type": "dom" | "tabs" | "navigate" | "back" | "act" | "reload", ... } forwarded request
"""

import hashlib
import hmac
import json
import os
import platform
import re
import secrets
import struct
import sys
import threading
import time
import uuid
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

HOST_VERSION = "1.0.0"


def file_build():
    """First 12 hex digits of the SHA-256 of this file, so a client can tell which file runs."""
    try:
        with open(os.path.abspath(__file__), "rb") as f:
            return hashlib.sha256(f.read()).hexdigest()[:12]
    except OSError:
        return "unknown"


BUILD = file_build()  # computed once at the start, it does not change in a running process
APP_ID = "outlook-web-helper"
DEFAULT_PORT = 8767
MAX_BODY = 4096
MAX_FIELD = 300
MAX_RECENT = 30
# seconds the HTTP side waits for the add-on, per request type
TIMEOUTS = {"tabs": 40.0, "dom": 55.0, "act": 45.0, "navigate": 40.0, "back": 40.0, "focus": 10.0, "reload": 10.0}


def config_dir():
    override = os.environ.get("OUTLOOK_WEB_HELPER_CONFIG_DIR")
    if override:
        return override
    if os.name == "nt":
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
    elif sys.platform == "darwin":
        base = os.path.expanduser("~/Library/Application Support")
    else:
        base = os.environ.get("XDG_CONFIG_HOME") or os.path.expanduser("~/.config")
    return os.path.join(base, APP_ID)


CONFIG_DIR = config_dir()
CONFIG_PATH = os.path.join(CONFIG_DIR, "config.json")


def load_config():
    cfg = {}
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            cfg = json.load(f)
    except (OSError, ValueError):
        pass
    if not isinstance(cfg, dict):
        cfg = {}
    port = cfg.get("port")
    if not isinstance(port, int) or isinstance(port, bool) or not 1024 <= port <= 65535:
        cfg["port"] = DEFAULT_PORT
    if not isinstance(cfg.get("adminToken"), str) or len(cfg["adminToken"]) < 16:
        cfg["adminToken"] = secrets.token_urlsafe(18)
    save_config(cfg)
    return cfg


def save_config(cfg):
    os.makedirs(CONFIG_DIR, exist_ok=True)
    tmp = CONFIG_PATH + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2)
    try:
        os.chmod(tmp, 0o600)
    except OSError:
        pass
    os.replace(tmp, CONFIG_PATH)


# ── Shared state ───────────────────────────────────────────────────────────────

STARTED = time.time()
STATE_LOCK = threading.Lock()
CONFIG = load_config()
RECENT = []  # the last requests, newest last, shown on the options page
PENDING = {}  # requestId -> {"event": Event, "result": dict|None}
SERVER = {"httpd": None, "thread": None, "error": None}
STDOUT_LOCK = threading.Lock()


def note_request(endpoint, result, started):
    """Remember a request for the options page. Only the endpoint, never the content."""
    entry = {"time": time.time(), "endpoint": endpoint, "ok": bool(result.get("ok")),
             "ms": int((time.time() - started) * 1000),
             "error": "" if result.get("ok") else str(result.get("error", ""))[:200]}
    with STATE_LOCK:
        RECENT.append(entry)
        del RECENT[:-MAX_RECENT]


# ── Native messaging I/O ───────────────────────────────────────────────────────


def read_message():
    raw_len = sys.stdin.buffer.read(4)
    if len(raw_len) < 4:
        return None
    (length,) = struct.unpack("=I", raw_len)
    data = sys.stdin.buffer.read(length)
    if len(data) < length:
        return None
    return json.loads(data.decode("utf-8"))


def send_message(obj):
    data = json.dumps(obj, ensure_ascii=False).encode("utf-8")
    with STDOUT_LOCK:
        sys.stdout.buffer.write(struct.pack("=I", len(data)))
        sys.stdout.buffer.write(data)
        sys.stdout.buffer.flush()


def forward(request_type, payload):
    """Send a request to the add-on and wait for its answer (a dict)."""
    timeout = TIMEOUTS[request_type]
    scale = os.environ.get("OUTLOOK_WEB_HELPER_MAX_WAIT")  # lets the tests run a timeout quickly
    if scale:
        timeout = min(timeout, float(scale))
    request_id = uuid.uuid4().hex
    entry = {"event": threading.Event(), "result": None}
    with STATE_LOCK:
        PENDING[request_id] = entry
    try:
        send_message({"requestId": request_id, "type": request_type, **payload})
    except (OSError, ValueError) as e:
        with STATE_LOCK:
            PENDING.pop(request_id, None)
        return {"ok": False, "error": "Could not reach the add-on: %s" % e}
    if not entry["event"].wait(timeout):
        with STATE_LOCK:
            PENDING.pop(request_id, None)
        return {"ok": False, "error": "The add-on did not answer in time."}
    with STATE_LOCK:
        PENDING.pop(request_id, None)
    return entry["result"] or {"ok": False, "error": "Empty answer from the add-on."}


# ── HTTP server ────────────────────────────────────────────────────────────────


def is_loopback(address):
    return address in ("127.0.0.1", "::1") or address.startswith("::ffff:127.")


class Handler(BaseHTTPRequestHandler):
    server_version = "OutlookWebHelper/" + HOST_VERSION
    protocol_version = "HTTP/1.1"

    def log_message(self, fmt, *args):
        pass  # stdout belongs to native messaging, never print to it

    # -- replies --------------------------------------------------------------

    def _reply(self, status, obj):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(body)

    def _require(self):
        """Return True if the request may go on. Otherwise the reply is sent."""
        header = self.headers.get("Authorization", "")
        given = header[7:].strip().encode() if header.startswith("Bearer ") else None
        with STATE_LOCK:
            admin = CONFIG["adminToken"].encode()
        if given is None or not hmac.compare_digest(given, admin):
            self._reply(401, {"ok": False, "error": "Invalid token."})
            return False
        if not is_loopback(self.client_address[0]):
            self._reply(403, {"ok": False, "error": "Only for the local computer."})
            return False
        return True

    def _forward(self, endpoint, request_type, payload):
        started = time.time()
        result = forward(request_type, payload)
        note_request(endpoint, result, started)
        # the add-on may answer 404 (no tab) or 409 (a bad tab), every other failure is 502
        self._reply(200 if result.get("ok") else (result["status"] if result.get("status") in (404, 409) else 502), result)

    # -- GET ------------------------------------------------------------------

    def do_HEAD(self):
        self.do_GET()

    def do_GET(self):
        path = urlparse(self.path).path
        if path == "/status":
            if not self._require():
                return
            return self._reply(200, {"ok": True, "addon": True, "hostVersion": HOST_VERSION, "build": BUILD})
        if path == "/tabs":
            if not self._require():
                return
            return self._forward("GET /tabs", "tabs", {})
        self._reply(404, {"ok": False, "error": "Not found."})

    # -- POST -----------------------------------------------------------------

    def _read_json(self):
        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            length = -1
        if length < 0 or length > MAX_BODY:
            self._reply(413, {"ok": False, "error": "Bad request size."})
            return None
        try:
            data = json.loads(self.rfile.read(length).decode("utf-8") or "{}")
        except (ValueError, UnicodeDecodeError):
            self._reply(400, {"ok": False, "error": "Body is not valid JSON."})
            return None
        if not isinstance(data, dict):
            self._reply(400, {"ok": False, "error": "Body must be a JSON object."})
            return None
        return data

    def _tab_id(self, data):
        """The tabId of a request. Without a whole number the reply is sent and None comes back."""
        tab_id = data.get("tabId")
        if not isinstance(tab_id, int) or isinstance(tab_id, bool) or tab_id < 0:
            self._reply(400, {"ok": False, "code": "TAB_REQUIRED",
                              "error": "tabId is required, list the tabs with GET /tabs"})
            return None
        return tab_id

    def do_POST(self):
        path = urlparse(self.path).path
        handlers = {"/dom": self._post_dom, "/navigate": self._post_navigate,
                    "/act": self._post_act, "/back": self._post_back,
                    "/focus": self._post_focus, "/reload": self._post_reload}
        if path not in handlers:
            return self._reply(404, {"ok": False, "error": "Not found."})
        if not self._require():
            return
        handlers[path]()

    def _post_dom(self):
        data = self._read_json()
        if data is None:
            return
        tab_id = self._tab_id(data)
        if tab_id is None:
            return
        payload = {"tabId": tab_id}
        selector = data.get("selector")
        if isinstance(selector, str) and selector.strip():
            payload["selector"] = selector.strip()[:MAX_FIELD]
        mode = data.get("mode")
        if mode is not None:
            if mode not in ("html", "text", "summary"):
                return self._reply(400, {"ok": False, "error": "mode must be html, text or summary."})
            payload["mode"] = mode
        wait_for = data.get("waitFor")
        if isinstance(wait_for, str) and wait_for.strip():
            payload["waitFor"] = wait_for.strip()[:MAX_FIELD]
        for key in ("maxBytes", "limit", "timeout"):
            value = data.get(key)
            if isinstance(value, int) and value >= 0:
                payload[key] = value
        if isinstance(data.get("strip"), bool):
            payload["strip"] = data["strip"]
        self._forward("POST /dom", "dom", payload)

    def _post_act(self):
        data = self._read_json()
        if data is None:
            return
        tab_id = self._tab_id(data)
        if tab_id is None:
            return
        action = data.get("action")
        selector = data.get("selector")
        if action not in ("click", "type", "press"):
            return self._reply(400, {"ok": False, "error": "action must be click, type or press."})
        if not isinstance(selector, str) or not selector.strip():
            return self._reply(400, {"ok": False, "error": "selector is required."})
        payload = {"tabId": tab_id, "action": action, "selector": selector.strip()[:MAX_FIELD]}
        text = data.get("text")
        if text is not None:
            if not isinstance(text, str) or len(text) > MAX_FIELD:
                return self._reply(400, {"ok": False, "error": "text must be a short string."})
            payload["text"] = text
        key = data.get("key")
        if key is not None:
            if not isinstance(key, str) or len(key) > 20:
                return self._reply(400, {"ok": False, "error": "key must be a short string."})
            payload["key"] = key
        for name in ("index", "timeout"):
            if isinstance(data.get(name), int) and data[name] >= 0:
                payload[name] = data[name]
        if isinstance(data.get("slow"), bool):
            payload["slow"] = data["slow"]  # click: let the press last a moment
        self._forward("POST /act", "act", payload)

    def _post_reload(self):
        data = self._read_json()
        if data is None:
            return
        tab_id = self._tab_id(data)
        if tab_id is None:
            return
        self._forward("POST /reload", "reload", {"tabId": tab_id})  # the tab is only checked

    def _post_focus(self):
        data = self._read_json()
        if data is None:
            return
        tab_id = self._tab_id(data)
        if tab_id is None:
            return
        self._forward("POST /focus", "focus", {"tabId": tab_id})

    def _post_back(self):
        data = self._read_json()
        if data is None:
            return
        tab_id = self._tab_id(data)
        if tab_id is None:
            return
        self._forward("POST /back", "back", {"tabId": tab_id})

    def _post_navigate(self):
        data = self._read_json()
        if data is None:
            return
        tab_id = self._tab_id(data)
        if tab_id is None:
            return
        url = data.get("url")
        if not isinstance(url, str) or not re.match(r"^https://outlook\.live\.com/", url):
            return self._reply(400, {"ok": False, "error": "url must be an https://outlook.live.com/ page."})
        payload = {"tabId": tab_id, "url": url[:2000]}
        self._forward("POST /navigate", "navigate", payload)


def stop_server():
    httpd = SERVER["httpd"]
    if httpd is not None:
        try:
            httpd.shutdown()
            httpd.server_close()
        except OSError:
            pass
    SERVER["httpd"] = None
    SERVER["thread"] = None


def start_server():
    stop_server()
    SERVER["error"] = None
    with STATE_LOCK:
        port = CONFIG["port"]
    try:
        httpd = ThreadingHTTPServer(("127.0.0.1", port), Handler)  # loopback only
    except OSError as e:
        SERVER["error"] = "Cannot listen on port %d: %s" % (port, e)
        return
    httpd.daemon_threads = True
    thread = threading.Thread(target=httpd.serve_forever, name="http", daemon=True)
    thread.start()
    SERVER["httpd"] = httpd
    SERVER["thread"] = thread


# ── Commands from the add-on ───────────────────────────────────────────────────


def get_info():
    """What the options page shows. Only the add-on can ask for this, it is no HTTP endpoint."""
    with STATE_LOCK:
        cfg = dict(CONFIG)
        recent = [dict(e) for e in RECENT]
    return {
        "hostVersion": HOST_VERSION,
        "build": BUILD,
        "pythonVersion": platform.python_version(),
        "pid": os.getpid(),
        "uptime": int(time.time() - STARTED),
        "configPath": CONFIG_PATH,
        "port": cfg["port"],
        "adminToken": cfg["adminToken"],
        "serverRunning": SERVER["httpd"] is not None,
        "serverError": SERVER["error"],
        "recent": recent,
    }


def set_config(msg):
    with STATE_LOCK:
        if msg.get("regenerateAdminToken"):
            CONFIG["adminToken"] = secrets.token_urlsafe(18)
            save_config(CONFIG)
    return get_info()


def handle_command(msg):
    request_id = msg.get("requestId")
    cmd = msg.get("cmd")
    try:
        if cmd == "getInfo":
            result = get_info()
        elif cmd == "setConfig":
            result = set_config(msg)
        elif cmd == "result":
            with STATE_LOCK:
                entry = PENDING.get(request_id)
            if entry is None:
                return  # the HTTP side gave up already
            result = msg.get("result")
            entry["result"] = result if isinstance(result, dict) else {"ok": False, "error": "Bad result."}
            entry["event"].set()
            return
        else:
            raise ValueError("Unknown command: %s" % cmd)
        send_message({"requestId": request_id, "ok": True, "result": result})
    except Exception as e:  # report every failure to the add-on, never die silently
        send_message({"requestId": request_id, "ok": False, "error": str(e)})


def main():
    start_server()
    while True:
        try:
            msg = read_message()
        except (ValueError, OSError):
            break
        if msg is None:
            break  # Firefox closed the pipe
        if isinstance(msg, dict):
            threading.Thread(target=handle_command, args=(msg,), daemon=True).start()
    stop_server()


if __name__ == "__main__":
    main()
