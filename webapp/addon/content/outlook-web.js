/**
 * Content script of the Outlook Web Helper add-on for Outlook on the web (outlook.live.com).
 *
 * It answers messages of the background script:
 *
 *   { cmd: "ping" }
 *       -> { ok, url, title, readyState, visibility, hasFocus }
 *
 *   { cmd: "dom", selector?, mode?, maxBytes?, limit?, strip?, waitFor?, timeout? }
 *       -> { ok, url, title, mode, selector, matched, truncated, data }
 *
 * The player is built from web components with shadow roots, which a plain
 * document.querySelector() or outerHTML cannot see into. The functions below
 * walk through open and closed shadow roots (Firefox hands content scripts
 * the closed ones as element.openOrClosedShadowRoot).
 *
 * Selectors may cross shadow boundaries with ">>>":
 *   "my-app >>> my-button"
 * means: find my-app in the document, then my-button inside the
 * shadow root of each match.
 *
 *   { cmd: "act", action: "click" | "type" | "press", selector, index?, text?, key?, timeout? }
 *       -> { ok, action, matched, tag, text, url }
 *       Acts on the element the selector finds (the first one, or the one at
 *       "index"), waiting for it to appear like "waitFor" does (default 5000 ms).
 *         click  pointer and mouse events, then click
 *         type   focus, set the value (text may be empty to clear), input and change events
 *         press  keydown, keypress and keyup of "key" (Enter, Escape, ArrowDown, ...)
 *       The events are synthetic (isTrusted is false), the page may treat them
 *       differently from real input.
 *
 * "waitFor" is a selector (same syntax). The answer comes as soon as it matches,
 * or fails after "timeout" milliseconds (default 10000, at most 30000). The
 * player is a single page app, so results appear after the page has loaded.
 *
 * Modes of "dom":
 *   html     serialized markup with <#shadow-root> markers (default)
 *   text     the visible text, one line per text node
 *   summary  a list of interactive elements with tag, id, classes, role,
 *            label, text and link, to find what to click
 */

(() => {
  if (window.__outlookWebHelper) return;
  window.__outlookWebHelper = true;

  const DEFAULT_MAX_BYTES = 400000;
  const DEFAULT_LIMIT = 400;
  const VOID_TAGS = new Set(['area', 'base', 'br', 'col', 'embed', 'hr', 'img', 'input',
    'link', 'meta', 'source', 'track', 'wbr']);
  const DROP_TAGS = new Set(['script', 'style', 'noscript', 'link', 'meta']);
  const INTERACTIVE_TAGS = new Set(['a', 'button', 'input', 'select', 'textarea', 'summary']);
  const INTERACTIVE_ROLES = new Set(['button', 'link', 'tab', 'menuitem', 'option', 'checkbox',
    'switch', 'slider', 'searchbox', 'textbox', 'combobox', 'listitem', 'row']);

  const shadowOf = el => el.openOrClosedShadowRoot || el.shadowRoot || null;

  // ── selection ────────────────────────────────────────────────────────────────

  function deepQueryAll(selector) {
    const parts = selector.split('>>>').map(s => s.trim()).filter(Boolean);
    let current = [document];
    parts.forEach((part, index) => {
      const next = [];
      for (const root of current) {
        const scope = index === 0 ? root : shadowOf(root);
        if (scope) next.push(...scope.querySelectorAll(part));
      }
      current = next;
    });
    return parts.length === 0 ? [] : current;
  }

  /** Visit every element below root, going through shadow roots as well. */
  function* walk(root, depth = 0, hosts = []) {
    for (const child of root.children || []) {
      yield { el: child, depth, hosts };
      const shadow = shadowOf(child);
      if (shadow) yield* walk(shadow, depth + 1, [...hosts, describe(child)]);
      yield* walk(child, depth + 1, hosts);
    }
  }

  function describe(el) {
    let s = el.localName;
    if (el.id) s += `#${el.id}`;
    const classes = (el.getAttribute('class') || '').split(/\s+/).filter(Boolean).slice(0, 2);
    if (classes.length) s += '.' + classes.join('.');
    return s;
  }

  // ── html ─────────────────────────────────────────────────────────────────────

  function attributes(el, strip) {
    let out = '';
    for (const attr of el.attributes) {
      if (strip && attr.name === 'style') continue;
      let value = attr.value;
      if (value.length > 200) value = value.slice(0, 200) + '...';
      out += ` ${attr.name}="${value.replace(/"/g, '&quot;')}"`;
    }
    return out;
  }

  function serialize(node, state, depth) {
    if (state.bytes > state.maxBytes) {
      state.truncated = true;
      return '';
    }
    const pad = ' '.repeat(Math.min(depth, 40));
    let out;
    if (node.nodeType === Node.TEXT_NODE) {
      const text = node.nodeValue.replace(/\s+/g, ' ').trim();
      out = text ? `${pad}${text}\n` : '';
    } else if (node.nodeType === Node.ELEMENT_NODE) {
      const tag = node.localName;
      if (state.strip && DROP_TAGS.has(tag)) return '';
      out = `${pad}<${tag}${attributes(node, state.strip)}>`;
      if (VOID_TAGS.has(tag)) {
        out += '\n';
      } else if (state.strip && tag === 'svg') {
        out += `</svg>\n`;
      } else {
        out += '\n';
        const shadow = shadowOf(node);
        if (shadow) {
          out += `${pad} <#shadow-root>\n`;
          for (const child of shadow.childNodes) out += serialize(child, state, depth + 2);
          out += `${pad} </#shadow-root>\n`;
        }
        for (const child of node.childNodes) out += serialize(child, state, depth + 1);
        out += `${pad}</${tag}>\n`;
      }
    } else {
      return '';
    }
    state.bytes += out.length;
    return out;
  }

  // ── text and summary ─────────────────────────────────────────────────────────

  function collectText(root, lines, state) {
    for (const node of root.childNodes) {
      if (state.bytes > state.maxBytes) {
        state.truncated = true;
        return;
      }
      if (node.nodeType === Node.TEXT_NODE) {
        const text = node.nodeValue.replace(/\s+/g, ' ').trim();
        if (text) {
          lines.push(text);
          state.bytes += text.length;
        }
      } else if (node.nodeType === Node.ELEMENT_NODE && !DROP_TAGS.has(node.localName)) {
        const shadow = shadowOf(node);
        if (shadow) collectText(shadow, lines, state);
        collectText(node, lines, state);
      }
    }
  }

  function summarize(roots, limit) {
    const items = [];
    for (const root of roots) {
      for (const { el, hosts } of walk(root)) {
        const role = el.getAttribute('role');
        const tag = el.localName;
        const testid = el.getAttribute('data-testid');
        const interactive = INTERACTIVE_TAGS.has(tag) || INTERACTIVE_ROLES.has(role);
        if (tag === 'svg' || testid === 'Icon') continue;
        const textLength = (el.textContent || '').trim().length;
        const label = el.getAttribute('aria-label') || el.getAttribute('title');
        // Containers with a lot of text only repeat what their children say.
        const interesting = interactive || (textLength <= 60 && (testid || label));
        if (!interesting) continue;
        const text = (el.textContent || '').replace(/\s+/g, ' ').trim().slice(0, 80);
        items.push({
          selector: describe(el),
          inside: hosts.length ? hosts.join(' >>> ') : undefined,
          role: role || undefined,
          label: el.getAttribute('aria-label') || el.getAttribute('title') || undefined,
          testid: el.getAttribute('data-testid') || undefined,
          href: el.getAttribute('href') || undefined,
          type: el.getAttribute('type') || undefined,
          placeholder: el.getAttribute('placeholder') || undefined,
          text: text || undefined,
        });
        if (items.length >= limit) return items;
      }
    }
    return items;
  }

  // ── messages ─────────────────────────────────────────────────────────────────

  function renderDom(message) {
    const mode = message.mode || 'html';
    const selector = message.selector || '';
    const maxBytes = message.maxBytes || DEFAULT_MAX_BYTES;
    const strip = message.strip !== false;
    const roots = selector ? deepQueryAll(selector) : [document.documentElement];
    const result = {
      ok: true,
      url: location.href,
      title: document.title,
      mode,
      selector,
      matched: roots.length,
      truncated: false,
    };
    if (selector && roots.length === 0) {
      return { ...result, ok: false, error: `Nothing matches "${selector}".` };
    }
    const state = { bytes: 0, maxBytes, truncated: false, strip };
    const shown = roots.slice(0, 20);
    if (mode === 'summary') {
      result.data = summarize(shown, message.limit || DEFAULT_LIMIT);
    } else if (mode === 'text') {
      const lines = [];
      for (const root of shown) collectText(root, lines, state);
      result.data = lines.join('\n');
    } else {
      result.data = shown.map(root => serialize(root, state, 0)).join('\n');
    }
    result.truncated = state.truncated;
    return result;
  }

  async function waitForMatches(selector, timeout) {
    const end = Date.now() + Math.min(timeout, 30000);
    let found = deepQueryAll(selector);
    while (found.length === 0 && Date.now() < end) {
      await new Promise(resolve => setTimeout(resolve, 200));
      found = deepQueryAll(selector);
    }
    return found;
  }

  const KEY_CODES = { Enter: 13, Escape: 27, ArrowDown: 40, ArrowUp: 38, ArrowLeft: 37,
    ArrowRight: 39, Tab: 9, ' ': 32 };

  function fire(el, Type, name, init) {
    el.dispatchEvent(new Type(name, { bubbles: true, cancelable: true, composed: true, ...init }));
  }

  async function clickElement(el, slow = false) {
    el.scrollIntoView({ block: 'center' });
    const rect = el.getBoundingClientRect();
    const at = { clientX: rect.left + rect.width / 2, clientY: rect.top + rect.height / 2, button: 0,
      pointerType: 'mouse', isPrimary: true };
    // The pointer arrives before it presses: some pages only react to a press
    // on an element they have seen the pointer enter.
    for (const name of ['pointerover', 'pointerenter', 'mouseover', 'mouseenter', 'pointermove', 'mousemove']) {
      fire(el, name.startsWith('pointer') ? PointerEvent : MouseEvent, name, at);
    }
    for (const name of ['pointerdown', 'mousedown']) {
      fire(el, name.startsWith('pointer') ? PointerEvent : MouseEvent, name, { ...at, buttons: 1 });
    }
    if (slow) {
      // A real press lasts a moment, and a page that decides on press whether
      // the element was already selected needs that moment to settle.
      await new Promise(resolve => setTimeout(resolve, 120));
    } else {
      el.focus?.();
    }
    for (const name of ['pointerup', 'mouseup']) {
      fire(el, name.startsWith('pointer') ? PointerEvent : MouseEvent, name, at);
    }
    if (slow) fire(el, MouseEvent, 'click', { ...at, detail: 1 });
    else el.click();
  }

  function setInputValue(el, text) {
    el.focus();
    if (el.isContentEditable) {
      // A rich-text box (the attendee field, the body): it has no value to set.
      // Select what is in it and insert the text the way the browser does for
      // typing, which is what the editor listens for.
      const selection = getSelection();
      const range = document.createRange();
      range.selectNodeContents(el);
      selection.removeAllRanges();
      selection.addRange(range);
      if (text) document.execCommand('insertText', false, text);
      else document.execCommand('delete');
      return;
    }
    const proto = el instanceof HTMLTextAreaElement ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
    // React tracks the value itself, so go through the native setter
    Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, text);
    fire(el, InputEvent, 'input', { inputType: 'insertText', data: text });
    fire(el, Event, 'change', {});
  }

  function pressKey(el, key) {
    const init = { key, code: key.length === 1 ? `Key${key.toUpperCase()}` : key,
      keyCode: KEY_CODES[key] || 0, which: KEY_CODES[key] || 0 };
    el.focus?.();
    for (const name of ['keydown', 'keypress', 'keyup']) fire(el, KeyboardEvent, name, init);
  }

  async function handleAct(message) {
    const matches = await waitForMatches(message.selector, message.timeout ?? 5000);
    const el = matches[message.index || 0];
    if (!el) {
      return { ok: false, url: location.href, matched: matches.length,
        error: `Nothing matches "${message.selector}" at index ${message.index || 0}.` };
    }
    el.scrollIntoView({ block: 'center' });
    switch (message.action) {
      case 'click':
        await clickElement(el, message.slow === true);
        break;
      case 'type':
        setInputValue(el, message.text ?? '');
        break;
      case 'press':
        pressKey(el, message.key || 'Enter');
        break;
      default:
        return { ok: false, error: `Unknown action: ${message.action}` };
    }
    return {
      ok: true,
      action: message.action,
      matched: matches.length,
      tag: el.localName,
      text: (el.textContent || el.value || '').replace(/\s+/g, ' ').trim().slice(0, 80),
      url: location.href,
    };
  }

  /** Wait until the selector matches, then answer like renderDom(). */
  async function handleDom(message) {
    if (message.waitFor) {
      const timeout = Math.min(message.timeout || 10000, 30000);
      const end = Date.now() + timeout;
      while (deepQueryAll(message.waitFor).length === 0) {
        if (Date.now() > end) {
          return {
            ok: false,
            url: location.href,
            error: `Timed out after ${timeout} ms waiting for "${message.waitFor}".`,
          };
        }
        await new Promise(resolve => setTimeout(resolve, 200));
      }
    }
    return renderDom(message);
  }

  browser.runtime.onMessage.addListener(message => {
    try {
      switch (message?.cmd) {
        case 'ping':
          return Promise.resolve({
            ok: true,
            url: location.href,
            title: document.title,
            readyState: document.readyState,
            // A page that is hidden, or whose window has no focus, may ignore
            // what a script does to it.
            visibility: document.visibilityState,
            hasFocus: document.hasFocus(),
          });
        case 'dom':
          return handleDom(message);
        case 'act':
          return handleAct(message).catch(ex => ({ ok: false, error: `${ex.name}: ${ex.message}` }));
      }
    } catch (ex) {
      return Promise.resolve({ ok: false, error: `${ex.name}: ${ex.message}` });
    }
    return undefined;
  });
})();
