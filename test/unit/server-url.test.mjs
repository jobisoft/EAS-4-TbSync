import { test } from "node:test";
import assert from "node:assert/strict";
import {
  normalizeCustomServerUrl as norm,
  upgradeInsecureUrl as up,
} from "../../src/modules/eas/server-url.mjs";

const EP = "/Microsoft-Server-ActiveSync";

test("plain http without a port is upgraded", () => {
  assert.equal(norm("http://example.com"), `https://example.com${EP}`);
  assert.equal(norm("http://example.com:80"), `https://example.com${EP}`);
  assert.equal(up("http://example.com/x"), "https://example.com/x");
  assert.equal(up(new URL("http://example.com")).protocol, "https:");
});

test("http with an explicit port is kept", () => {
  assert.equal(norm("http://example.com:8080"), `http://example.com:8080${EP}`);
});

test("http to loopback is kept", () => {
  assert.equal(norm("http://localhost"), `http://localhost${EP}`);
  assert.equal(norm("http://127.0.0.1"), `http://127.0.0.1${EP}`);
  assert.equal(up("http://[::1]/x"), "http://[::1]/x");
});

test("https, bare hosts and other schemes", () => {
  assert.equal(norm("example.com"), `https://example.com${EP}`);
  assert.equal(norm("https://example.com"), `https://example.com${EP}`);
  assert.equal(norm("ftp://example.com"), null);
});
