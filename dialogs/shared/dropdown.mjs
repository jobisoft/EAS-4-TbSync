/**
 * Reusable rich dropdown widget - used by setup.html and config.html.
 *
 * Each option is a {value, label, hint} triple. The trigger shows the
 * currently-selected option's label + hint; clicking opens a panel
 * with the same row format. Keyboard support: ArrowDown/Up navigates,
 * Enter/Space selects, Escape closes. Click-outside closes.
 *
 * Locked mode renders the trigger as `disabled` and never opens - used
 * by the config popup to surface the account type read-only.
 */

/**
 * Replace the contents of `rootEl` with a fresh dropdown.
 *
 * @param {HTMLElement} rootEl  container that will own the markup
 * @param {{
 *   options: Array<{value: string, label: string, hint?: string}>,
 *   value?: string,                // initial selection (defaults to options[0].value)
 *   locked?: boolean,              // render disabled, never open
 *   onChange?: (newValue: string) => void,
 * }} config
 * @returns {{
 *   getValue: () => string,
 *   setValue: (v: string) => void,
 *   setLocked: (locked: boolean) => void,
 * }}
 */
export function createDropdown(rootEl, { options, value, locked = false, onChange }) {
  if (!options?.length) throw new Error("createDropdown: options is required");
  let currentValue = options.some(o => o.value === value) ? value : options[0].value;
  let isLocked = !!locked;
  let isOpen = false;

  rootEl.classList.add("dropdown");
  rootEl.dataset.state = "closed";

  const trigger = document.createElement("button");
  trigger.type = "button";
  trigger.className = "dropdown-trigger";
  trigger.setAttribute("aria-haspopup", "listbox");
  trigger.setAttribute("aria-expanded", "false");

  const currentBlock = document.createElement("span");
  currentBlock.className = "dropdown-current";
  const currentLabel = document.createElement("span");
  currentLabel.className = "dropdown-label";
  const currentHint = document.createElement("span");
  currentHint.className = "dropdown-hint";
  currentBlock.append(currentLabel, currentHint);

  const caret = document.createElement("span");
  caret.className = "dropdown-caret";
  caret.textContent = "▾";
  caret.setAttribute("aria-hidden", "true");

  trigger.append(currentBlock, caret);

  const panel = document.createElement("ul");
  panel.className = "dropdown-panel";
  panel.setAttribute("role", "listbox");
  panel.hidden = true;

  const optionEls = options.map((opt, i) => {
    const li = document.createElement("li");
    li.setAttribute("role", "option");
    li.dataset.value = opt.value;
    li.tabIndex = -1;
    const lbl = document.createElement("span");
    lbl.className = "label";
    lbl.textContent = opt.label;
    li.appendChild(lbl);
    if (opt.hint) {
      const h = document.createElement("span");
      h.className = "hint";
      h.textContent = opt.hint;
      li.appendChild(h);
    }
    li.addEventListener("click", () => {
      selectValue(opt.value);
      close();
      trigger.focus();
    });
    li.addEventListener("mouseenter", () => focusOption(i));
    return li;
  });
  panel.append(...optionEls);

  rootEl.replaceChildren(trigger, panel);

  function syncTrigger() {
    const opt = options.find(o => o.value === currentValue) ?? options[0];
    currentLabel.textContent = opt.label;
    currentHint.textContent = opt.hint ?? "";
    for (const li of optionEls) {
      li.classList.toggle("selected", li.dataset.value === currentValue);
      li.setAttribute("aria-selected", li.dataset.value === currentValue ? "true" : "false");
    }
  }

  function selectValue(v) {
    if (v === currentValue) return;
    currentValue = v;
    syncTrigger();
    onChange?.(v);
  }

  let focusedIdx = -1;
  function focusOption(i) {
    if (i < 0 || i >= optionEls.length) return;
    focusedIdx = i;
    for (const [idx, li] of optionEls.entries()) {
      li.classList.toggle("focused", idx === i);
    }
    optionEls[i].focus();
  }

  function open() {
    if (isOpen || isLocked) return;
    isOpen = true;
    panel.hidden = false;
    rootEl.dataset.state = "open";
    trigger.setAttribute("aria-expanded", "true");
    const initialIdx = Math.max(0, options.findIndex(o => o.value === currentValue));
    focusOption(initialIdx);
    document.addEventListener("click", onDocClick, true);
    document.addEventListener("keydown", onDocKey, true);
  }

  function close() {
    if (!isOpen) return;
    isOpen = false;
    panel.hidden = true;
    rootEl.dataset.state = "closed";
    trigger.setAttribute("aria-expanded", "false");
    for (const li of optionEls) li.classList.remove("focused");
    document.removeEventListener("click", onDocClick, true);
    document.removeEventListener("keydown", onDocKey, true);
  }

  function onDocClick(e) {
    if (!rootEl.contains(e.target)) close();
  }

  function onDocKey(e) {
    if (e.key === "Escape") {
      e.preventDefault();
      close();
      trigger.focus();
      return;
    }
    if (e.key === "ArrowDown") {
      e.preventDefault();
      focusOption(Math.min(optionEls.length - 1, focusedIdx + 1));
      return;
    }
    if (e.key === "ArrowUp") {
      e.preventDefault();
      focusOption(Math.max(0, focusedIdx - 1));
      return;
    }
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      const opt = options[focusedIdx];
      if (opt) {
        selectValue(opt.value);
        close();
        trigger.focus();
      }
    }
  }

  trigger.addEventListener("click", () => {
    if (isLocked) return;
    isOpen ? close() : open();
  });
  trigger.addEventListener("keydown", e => {
    if (isLocked) return;
    if (e.key === "ArrowDown" || e.key === "ArrowUp") {
      e.preventDefault();
      open();
    }
  });

  function applyLocked() {
    trigger.disabled = isLocked;
    rootEl.classList.toggle("locked", isLocked);
  }

  syncTrigger();
  applyLocked();

  return {
    getValue: () => currentValue,
    setValue: (v) => selectValue(v),
    setLocked: (next) => { isLocked = !!next; if (isLocked) close(); applyLocked(); },
  };
}
