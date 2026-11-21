#!/usr/bin/env python3
"""
Tests for the native helper app. They play the role of Firefox (stdin/stdout)
and of a script (HTTP). Every run uses its own port and its own config folder.
Run with:  python3 test_host.py
"""
import json
import os
import socket
import struct
import subprocess
import sys
import tempfile
import threading
import time
import unittest
import urllib.error
import urllib.request

HOST = os.path.join(os.path.dirname(os.path.abspath(__file__)), "native-messaging-app",
                    "outlook_web_helper_host.py")


def free_port():
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    port = s.getsockname()[1]
    s.close()
    return port


class HostTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.dir = tempfile.TemporaryDirectory()
        cls.port = free_port()
        with open(os.path.join(cls.dir.name, "config.json"), "w") as f:
            json.dump({"port": cls.port}, f)
        env = dict(os.environ, OUTLOOK_WEB_HELPER_CONFIG_DIR=cls.dir.name, OUTLOOK_WEB_HELPER_MAX_WAIT="2")
        cls.proc = subprocess.Popen([sys.executable, "-I", HOST], stdin=subprocess.PIPE,
                                    stdout=subprocess.PIPE, env=env)
        cls.inbox = []
        cls.cond = threading.Condition()
        threading.Thread(target=cls._reader, daemon=True).start()
        cls.info = cls.command("getInfo")["result"]
        with open(os.path.join(cls.dir.name, "config.json")) as f:
            cls.admin = json.load(f)["adminToken"]

    @classmethod
    def tearDownClass(cls):
        cls.proc.stdin.close()
        cls.proc.wait(5)
        cls.dir.cleanup()

    # -- the Firefox side --------------------------------------------------

    @classmethod
    def _reader(cls):
        while True:
            raw = cls.proc.stdout.read(4)
            if len(raw) < 4:
                return
            msg = json.loads(cls.proc.stdout.read(struct.unpack("=I", raw)[0]))
            with cls.cond:
                cls.inbox.append(msg)
                cls.cond.notify_all()

    @classmethod
    def send(cls, obj):
        data = json.dumps(obj).encode()
        cls.proc.stdin.write(struct.pack("=I", len(data)) + data)
        cls.proc.stdin.flush()

    @classmethod
    def wait_for(cls, predicate, timeout=5):
        end = time.time() + timeout
        with cls.cond:
            while True:
                for msg in cls.inbox:
                    if predicate(msg):
                        cls.inbox.remove(msg)
                        return msg
                left = end - time.time()
                if left <= 0:
                    raise TimeoutError("no matching message")
                cls.cond.wait(left)

    @classmethod
    def command(cls, cmd, **args):
        rid = "t%d" % time.time_ns()
        cls.send({"requestId": rid, "cmd": cmd, **args})
        return cls.wait_for(lambda m: m.get("requestId") == rid)

    def addon_answers(self, type_, result):
        """Answer the next forwarded request of the given type."""
        def run():
            msg = self.wait_for(lambda m: m.get("type") == type_)
            self.seen = msg
            self.send({"requestId": msg["requestId"], "cmd": "result", "result": result})
        t = threading.Thread(target=run)
        t.start()
        return t

    # -- the HTTP side -----------------------------------------------------

    def http(self, method, path, body=None, token=None):
        headers = {"Authorization": "Bearer " + token} if token else {}
        req = urllib.request.Request("http://127.0.0.1:%d%s" % (self.port, path), method=method,
                                     data=json.dumps(body).encode() if body is not None else None,
                                     headers=headers)
        try:
            with urllib.request.urlopen(req, timeout=30) as r:
                return r.status, json.loads(r.read())
        except urllib.error.HTTPError as e:
            with e:
                return e.code, json.loads(e.read())

    # -- tests -------------------------------------------------------------

    def test_config_is_created(self):
        with open(os.path.join(self.dir.name, "config.json")) as f:
            cfg = json.load(f)
        self.assertEqual(cfg["port"], self.port)
        self.assertGreaterEqual(len(cfg["adminToken"]), 16)
        self.assertNotIn("token", cfg)
        self.assertTrue(self.info["serverRunning"])

    def test_default_config_in_a_fresh_folder(self):
        with tempfile.TemporaryDirectory() as d:
            code = ("import sys; sys.path.insert(0, %r); import outlook_web_helper_host as h; "
                    "print(h.CONFIG['port'], len(h.CONFIG['adminToken']))" % os.path.dirname(HOST))
            out = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True,
                                 env=dict(os.environ, OUTLOOK_WEB_HELPER_CONFIG_DIR=d)).stdout.split()
            self.assertEqual(out[0], "8767")
            self.assertGreaterEqual(int(out[1]), 16)
            self.assertTrue(os.path.exists(os.path.join(d, "config.json")))
            if os.name == "posix":
                self.assertEqual(os.stat(os.path.join(d, "config.json")).st_mode & 0o777, 0o600)

    def test_auth(self):
        self.assertEqual(self.http("GET", "/status")[0], 401)
        self.assertEqual(self.http("GET", "/status", token="bad")[0], 401)
        status, body = self.http("GET", "/status", token=self.admin)
        self.assertEqual(status, 200)
        self.assertEqual((body["ok"], body["addon"]), (True, True))
        for method, path in (("GET", "/tabs"), ("POST", "/dom"), ("POST", "/act"),
                             ("POST", "/navigate"), ("POST", "/back"), ("POST", "/reload")):
            self.assertEqual(self.http(method, path, {} if method == "POST" else None)[0], 401, path)
        self.assertEqual(self.http("GET", "/ping")[0], 404)
        self.assertEqual(self.http("POST", "/play", {}, token=self.admin)[0], 404)
        self.assertEqual(self.http("GET", "/info", token=self.admin)[0], 404)

    def test_status_tells_version_and_build(self):
        import hashlib
        with open(HOST, "rb") as f:
            expected = hashlib.sha256(f.read()).hexdigest()[:12]
        first = self.http("GET", "/status", token=self.admin)[1]
        self.assertEqual(first["hostVersion"], self.info["hostVersion"])
        self.assertRegex(first["build"], r"^[0-9a-f]{12}$")
        self.assertEqual(first["build"], expected)
        self.assertEqual(self.http("GET", "/status", token=self.admin)[1]["build"], first["build"])
        info = self.command("getInfo")["result"]
        self.assertEqual((info["hostVersion"], info["build"]), (first["hostVersion"], first["build"]))

    def test_server_is_loopback_only(self):
        # the server socket is bound to 127.0.0.1, so no LAN address reaches it
        out = subprocess.run(["ss", "-ltn"], capture_output=True, text=True).stdout
        if out:
            lines = [l for l in out.splitlines() if ":%d " % self.port in l]
            self.assertTrue(lines)
            self.assertTrue(all("127.0.0.1:%d" % self.port in l for l in lines), lines)

    def test_tabs_is_forwarded(self):
        t = self.addon_answers("tabs", {"ok": True, "tabs": []})
        self.assertEqual(self.http("GET", "/tabs", token=self.admin)[1]["tabs"], [])
        t.join()

    def test_dom_is_forwarded(self):
        t = self.addon_answers("dom", {"ok": True, "url": "https://outlook.live.com/", "data": "<html></html>"})
        status, body = self.http("POST", "/dom", {"tabId": 3, "selector": "my-app >>> input", "mode": "html",
                                                  "maxBytes": 1000, "bogus": 1}, token=self.admin)
        t.join()
        self.assertEqual((status, body["data"]), (200, "<html></html>"))
        self.assertEqual(self.seen["selector"], "my-app >>> input")
        self.assertEqual(self.seen["maxBytes"], 1000)
        self.assertNotIn("bogus", self.seen)
        self.assertEqual(self.http("POST", "/dom", {"tabId": 3, "mode": "evil"}, token=self.admin)[0], 400)
        t = self.addon_answers("dom", {"ok": True, "data": []})
        self.http("POST", "/dom", {"tabId": 3, "mode": "summary", "waitFor": "[data-testid=x]", "timeout": 5000,
                                   "limit": 7, "tabId": 3}, token=self.admin)
        t.join()
        self.assertEqual((self.seen["waitFor"], self.seen["timeout"], self.seen["limit"], self.seen["tabId"]),
                         ("[data-testid=x]", 5000, 7, 3))

    def test_act_is_forwarded_and_validated(self):
        t = self.addon_answers("act", {"ok": True, "matched": 1})
        status, body = self.http("POST", "/act", {"tabId": 3, "action": "type", "selector": "input", "text": "Abba",
                                                  "index": 0, "junk": 1}, token=self.admin)
        t.join()
        self.assertEqual((status, body["matched"]), (200, 1))
        self.assertEqual(self.seen["text"], "Abba")
        self.assertNotIn("junk", self.seen)
        for bad in ({"action": "eval", "selector": "a"}, {"action": "click"}, {"action": "click", "selector": " "},
                    {"action": "type", "selector": "a", "text": "x" * 400}):
            self.assertEqual(self.http("POST", "/act", dict(bad, tabId=3), token=self.admin)[0], 400)

    def test_slow_click_is_forwarded_only_as_a_boolean(self):
        t = self.addon_answers("act", {"ok": True, "matched": 1})
        self.http("POST", "/act", {"tabId": 3, "action": "click", "selector": "a", "slow": True}, token=self.admin)
        t.join()
        self.assertIs(self.seen["slow"], True)
        t = self.addon_answers("act", {"ok": True, "matched": 1})
        self.http("POST", "/act", {"tabId": 3, "action": "click", "selector": "a", "slow": "yes"}, token=self.admin)
        t.join()
        self.assertNotIn("slow", self.seen)

    def test_focus_is_forwarded_and_needs_auth(self):
        t = self.addon_answers("focus", {"ok": True, "tabId": 4, "active": True, "windowId": 1})
        status, body = self.http("POST", "/focus", {"tabId": 4, "junk": 1}, token=self.admin)
        t.join()
        self.assertEqual((status, body["active"], self.seen["tabId"]), (200, True, 4))
        self.assertNotIn("junk", self.seen)
        self.assertEqual(self.http("POST", "/focus", {"tabId": 4})[0], 401)
        self.assertEqual(self.http("POST", "/focus", {}, token=self.admin)[0], 400)

    def test_navigate_is_forwarded_and_validated(self):
        t = self.addon_answers("navigate", {"ok": True, "url": "https://outlook.live.com/search/x"})
        status, body = self.http("POST", "/navigate", {"url": "https://outlook.live.com/search/x", "tabId": 4},
                                 token=self.admin)
        t.join()
        self.assertEqual((status, body["ok"], self.seen["tabId"]), (200, True, 4))
        for bad in ("https://evil.example/", "http://outlook.live.com/", "javascript:alert(1)", 5):
            self.assertEqual(self.http("POST", "/navigate", {"tabId": 4, "url": bad}, token=self.admin)[0], 400)

    def test_back_is_forwarded_and_validated(self):
        t = self.addon_answers("back", {"ok": True, "tabId": 4, "url": "https://outlook.live.com/"})
        status, body = self.http("POST", "/back", {"tabId": 4, "junk": 1}, token=self.admin)
        t.join()
        self.assertEqual((status, body["url"], self.seen["tabId"]), (200, "https://outlook.live.com/", 4))
        self.assertNotIn("junk", self.seen)
        t = self.addon_answers("back", {"ok": True, "url": "https://outlook.live.com/"})
        self.assertEqual(self.http("POST", "/back", {"tabId": 4}, token=self.admin)[0], 200)
        t.join()
        self.assertEqual(self.seen["tabId"], 4)

    def test_back_needs_auth(self):
        self.assertEqual(self.http("POST", "/back", {})[0], 401)
        self.assertEqual(self.http("POST", "/back", {}, token="bad")[0], 401)

    def test_back_errors_give_502(self):
        err = "The tab has no earlier page in its history."
        t = self.addon_answers("back", {"ok": False, "error": err})
        status, body = self.http("POST", "/back", {"tabId": 4}, token=self.admin)
        t.join()
        self.assertEqual((status, body["error"]), (502, err))

    def test_back_times_out_when_the_add_on_is_silent(self):
        status, body = self.http("POST", "/back", {"tabId": 4}, token=self.admin)
        self.assertEqual(status, 502)
        self.assertIn("did not answer in time", body["error"])
        self.wait_for(lambda m: m.get("type") == "back")

    def test_tabs_answer_is_passed_through(self):
        tabs = [{"tabId": 1, "windowId": 2, "contentScript": True}]
        t = self.addon_answers("tabs", {"ok": True, "tabs": tabs})
        self.assertEqual(self.http("GET", "/tabs", token=self.admin)[1]["tabs"], tabs)
        t.join()

    def test_invalid_tab_gives_409_for_all_tab_calls(self):
        calls = [("dom", "/dom", {"tabId": 3, "mode": "text"}),
                 ("act", "/act", {"tabId": 3, "action": "click", "selector": "a"}),
                 ("navigate", "/navigate", {"tabId": 3, "url": "https://outlook.live.com/"}),
                 ("back", "/back", {"tabId": 3}),
                 ("reload", "/reload", {"tabId": 3})]
        for kind, path, body in calls:
            t = self.addon_answers(kind, {"ok": False, "status": 409, "code": "INVALID_TAB", "error": "gone"})
            status, answer = self.http("POST", path, body, token=self.admin)
            t.join()
            self.assertEqual((status, answer["code"]), (409, "INVALID_TAB"), path)
            self.assertEqual(self.seen["tabId"], 3, path)
            self.assertEqual(self.http("POST", path, body)[0], 401)

    def test_tab_id_is_required_for_all_tab_calls(self):
        bodies = {"/dom": {"mode": "text"}, "/act": {"action": "click", "selector": "a"},
                  "/navigate": {"url": "https://outlook.live.com/"}, "/back": {}, "/reload": {}}
        for path, body in bodies.items():
            self.assertEqual(self.http("POST", path, body)[0], 401, path)  # auth comes before validation
            for bad in ({}, {"tabId": "3"}, {"tabId": 1.5}, {"tabId": True}, {"tabId": None}, {"tabId": -1}):
                status, answer = self.http("POST", path, dict(body, **bad), token=self.admin)
                self.assertEqual(status, 400, (path, bad))
                self.assertEqual(answer, {"ok": False, "code": "TAB_REQUIRED",
                                          "error": "tabId is required, list the tabs with GET /tabs"})
        # none of these requests reached the add-on
        with self.assertRaises(TimeoutError):
            self.wait_for(lambda m: "type" in m, timeout=0.5)

    def test_reload_is_forwarded(self):
        t = self.addon_answers("reload", {"ok": True, "message": "Reloading the add-on."})
        self.assertEqual(self.http("POST", "/reload", {"tabId": 3}, token=self.admin)[0], 200)
        t.join()

    def test_add_on_error_gives_502(self):
        t = self.addon_answers("dom", {"ok": False, "error": "Nothing matches \"x\"."})
        status, body = self.http("POST", "/dom", {"tabId": 3, "selector": "x"}, token=self.admin)
        t.join()
        self.assertEqual((status, body["ok"]), (502, False))

    def test_timeout_when_the_add_on_is_silent(self):
        started = time.time()
        status, body = self.http("GET", "/tabs", token=self.admin)
        self.assertEqual(status, 502)
        self.assertIn("did not answer in time", body["error"])
        self.assertLess(time.time() - started, 10)
        self.wait_for(lambda m: m.get("type") == "tabs")  # drop the unanswered request

    def test_bad_bodies(self):
        self.assertEqual(self.http("POST", "/dom", [1], token=self.admin)[0], 400)
        req = urllib.request.Request("http://127.0.0.1:%d/dom" % self.port, method="POST", data=b"{nope",
                                     headers={"Authorization": "Bearer " + self.admin})
        with self.assertRaises(urllib.error.HTTPError) as cm:
            urllib.request.urlopen(req, timeout=5)
        self.assertEqual(cm.exception.code, 400)
        cm.exception.close()
        big = {"selector": "x" * 5000}
        self.assertEqual(self.http("POST", "/dom", dict(big, tabId=3), token=self.admin)[0], 413)

    def test_info_for_the_options_page(self):
        info = self.command("getInfo")["result"]
        self.assertEqual((info["port"], info["adminToken"]), (self.port, self.admin))
        self.assertIn("recent", info)
        self.assertFalse(self.command("nonsense")["ok"])

    def test_zz_new_token(self):
        reply = self.command("setConfig", regenerateAdminToken=True)
        self.assertTrue(reply["ok"])
        new = reply["result"]["adminToken"]
        self.assertNotEqual(new, self.admin)
        with open(os.path.join(self.dir.name, "config.json")) as f:
            self.assertEqual(json.load(f)["adminToken"], new)
        self.assertEqual(self.http("GET", "/status", token=self.admin)[0], 401)
        self.assertEqual(self.http("GET", "/status", token=new)[0], 200)


if __name__ == "__main__":
    unittest.main(verbosity=2)
