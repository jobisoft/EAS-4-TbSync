/**
 * Every CSS custom property a dialog uses must be one it can see.
 *
 * A dialog's colours come from `dialogs/shared/dialog.css`, and a shared
 * stylesheet like `dropdown.css` names tokens without declaring any. That
 * only works while every page loading such a sheet also loads the one that
 * declares what it names.
 *
 * It did not always. `config.html` linked `dropdown.css` while declaring its
 * own palette inline, and that palette was missing four of the tokens the
 * dropdown asks for - so its border, hover fill and open state resolved to
 * nothing and fell back to `currentColor`. Nothing failed, nothing was
 * logged, and it shipped.
 *
 * This is the check that catches the next one. It reads what each page
 * loads, not what it ought to load.
 *
 * Run with `npm run test:unit` (node --test).
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";

const DIALOGS = path.join(import.meta.dirname, "..", "..", "src", "dialogs");

/** The stylesheets a page carries: its own inline blocks, then every one it
 *  links, resolved relative to the page. A missing href is itself a fault
 *  and is reported as one rather than skipped. */
function stylesheetsOf(page) {
  const html = fs.readFileSync(page, "utf8");
  const sources = [...html.matchAll(/<style>([\s\S]*?)<\/style>/g)].map((m) => ({
    from: `${path.basename(page)} inline`,
    css: m[1],
  }));
  for (const m of html.matchAll(
    /<link[^>]*rel="stylesheet"[^>]*href="([^"]+)"/g,
  )) {
    const href = path.resolve(path.dirname(page), m[1]);
    assert.ok(fs.existsSync(href), `${page} links a missing ${m[1]}`);
    sources.push({ from: m[1], css: fs.readFileSync(href, "utf8") });
  }
  return sources;
}

const declared = (css) =>
  new Set([...css.matchAll(/(--[a-z0-9-]+)\s*:/g)].map((m) => m[1]));

/** Every `var(--x)` reference, with the sheet it appears in, so a failure
 *  names the file to go and look at. */
function usages(sources) {
  const out = [];
  for (const { from, css } of sources) {
    for (const m of css.matchAll(/var\(\s*(--[a-z0-9-]+)/g)) {
      out.push({ token: m[1], from });
    }
  }
  return out;
}

const pages = fs
  .readdirSync(DIALOGS, { withFileTypes: true })
  .filter((e) => e.isDirectory())
  .flatMap((e) =>
    fs
      .readdirSync(path.join(DIALOGS, e.name))
      .filter((f) => f.endsWith(".html"))
      .map((f) => path.join(DIALOGS, e.name, f)),
  );

test("there are dialogs to check", () => {
  assert.ok(pages.length >= 5, `found only ${pages.length} dialog pages`);
});

for (const page of pages) {
  const name = path.basename(page);
  test(`${name} declares every token it uses`, () => {
    const sources = stylesheetsOf(page);
    const have = new Set();
    for (const { css } of sources) for (const t of declared(css)) have.add(t);

    const missing = usages(sources).filter((u) => !have.has(u.token));
    assert.deepEqual(
      [...new Set(missing.map((u) => `${u.token} (used in ${u.from})`))].sort(),
      [],
      `${name} uses tokens nothing it loads declares`,
    );
  });
}

test("a dark value exists for every colour the light block sets", () => {
  const shared = path.join(DIALOGS, "shared", "dialog.css");
  const css = fs.readFileSync(shared, "utf8");
  const dark = css.slice(css.indexOf("prefers-color-scheme: dark"));
  const light = css.slice(0, css.indexOf("prefers-color-scheme: dark"));
  const missing = [...declared(light)].filter((t) => !declared(dark).has(t));
  assert.deepEqual(missing, [], "tokens with no dark counterpart");
});
