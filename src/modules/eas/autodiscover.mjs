/**
 * EAS Autodiscover client. Given an email + password, races up to four
 * standard HTTPS endpoints derived from the email's domain in parallel
 * and returns the first MobileSync server URL the response advertises.
 * Mirrors the legacy `getServerConnectionViaAutodiscover` algorithm in
 * `EAS-4-TbSync/content/includes/network.js`, ported to fetch + AbortSignal.
 */

const ENDPOINT_TEMPLATES = [
  "https://autodiscover.{domain}/autodiscover/autodiscover.xml",
  "https://{domain}/autodiscover/autodiscover.xml",
  "https://autodiscover.{domain}/Autodiscover/Autodiscover.xml",
  "https://{domain}/Autodiscover/Autodiscover.xml",
];

const STAGGER_MS = 200;
const REQUEST_SCHEMA =
  "http://schemas.microsoft.com/exchange/autodiscover/mobilesync/requestschema/2006";
const RESPONSE_SCHEMA =
  "http://schemas.microsoft.com/exchange/autodiscover/mobilesync/responseschema/2006";
const REDIRECT_LIMIT = 5;

export const DISCOVER_ERR = {
  AUTH: "E:AUTH",
  NO_SERVER: "E:NO_SERVER",
  NETWORK: "E:NETWORK",
  CANCELLED: "E:CANCELLED",
};

export class DiscoverError extends Error {
  constructor(code, message, details = {}) {
    super(message);
    this.name = "DiscoverError";
    this.code = code;
    this.details = details;
  }
}

export async function discoverEasServer({ email, password, signal }) {
  const at = email?.indexOf?.("@") ?? -1;
  if (at < 1 || at === email.length - 1) {
    throw new DiscoverError(DISCOVER_ERR.NO_SERVER, "Invalid email address", {
      tried: [],
    });
  }
  const domain = email.slice(at + 1);
  const urls = ENDPOINT_TEMPLATES.map((t) => t.replace("{domain}", domain));

  const tried = [];
  const attempts = urls.map((url, i) =>
    (async () => {
      await sleep(i * STAGGER_MS, signal);
      if (signal?.aborted) throw makeCancel();
      return tryUrl(url, email, password, signal, tried, REDIRECT_LIMIT);
    })(),
  );

  const settled = await Promise.allSettled(attempts);
  for (const r of settled) {
    if (r.status === "fulfilled" && r.value?.server) {
      return { server: r.value.server, user: r.value.user };
    }
  }
  const cancelled = settled.find(
    (r) => r.status === "rejected" && r.reason?.code === DISCOVER_ERR.CANCELLED,
  );
  if (cancelled) throw cancelled.reason;

  const allAuth = settled.every(
    (r) => r.status === "fulfilled" && r.value?.error === "auth",
  );
  const anyAuth = settled.some(
    (r) => r.status === "fulfilled" && r.value?.error === "auth",
  );
  const everyNetwork = settled.every(
    (r) => r.status === "fulfilled" && r.value?.error === "network",
  );
  if (allAuth || anyAuth) {
    throw new DiscoverError(DISCOVER_ERR.AUTH, "Credentials rejected", {
      tried,
    });
  }
  if (everyNetwork) {
    throw new DiscoverError(
      DISCOVER_ERR.NETWORK,
      "Could not reach any discovery endpoint",
      { tried },
    );
  }
  throw new DiscoverError(
    DISCOVER_ERR.NO_SERVER,
    "No MobileSync server found",
    { tried },
  );
}

async function tryUrl(url, user, password, signal, tried, redirectsLeft) {
  if (redirectsLeft <= 0) {
    tried.push({ url, status: "redirect-limit" });
    return { error: "no-server" };
  }
  let resp;
  try {
    resp = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "text/xml; charset=utf-8",
        "User-Agent": "Thunderbird ActiveSync",
        Authorization: "Basic " + btoa(`${user}:${password}`),
      },
      body: buildRequestBody(user),
      redirect: "follow",
      signal,
    });
  } catch (err) {
    if (err?.name === "AbortError") throw makeCancel();
    tried.push({
      url,
      status: "network",
      message: err?.message ?? String(err),
    });
    return { error: "network" };
  }

  if (resp.status === 401 || resp.status === 403) {
    tried.push({ url, status: resp.status });
    return { error: "auth" };
  }
  if (!resp.ok) {
    tried.push({ url, status: resp.status });
    return { error: "http" };
  }

  const text = await resp.text();
  const doc = new DOMParser().parseFromString(text, "application/xml");
  if (doc.querySelector("parsererror")) {
    tried.push({ url, status: "bad-xml" });
    return { error: "no-server" };
  }

  const action = doc.querySelector("Autodiscover > Response > Action");
  if (!action) {
    tried.push({ url, status: "no-action" });
    return { error: "no-server" };
  }

  const redirectEl = action.querySelector(":scope > Redirect");
  if (redirectEl?.textContent?.trim()) {
    return tryUrl(
      url,
      redirectEl.textContent.trim(),
      password,
      signal,
      tried,
      redirectsLeft - 1,
    );
  }

  const servers = action.querySelectorAll(":scope > Settings > Server");
  for (const s of servers) {
    const type = s.querySelector(":scope > Type")?.textContent?.trim();
    const serverUrl = s.querySelector(":scope > Url")?.textContent?.trim();
    if (type === "MobileSync" && serverUrl) {
      tried.push({ url, status: 200 });
      return { server: serverUrl, user };
    }
  }
  tried.push({ url, status: "no-mobilesync" });
  return { error: "no-server" };
}

function buildRequestBody(email) {
  return [
    `<?xml version="1.0" encoding="utf-8"?>`,
    `<Autodiscover xmlns="${REQUEST_SCHEMA}">`,
    `<Request>`,
    `<EMailAddress>${escapeXml(email)}</EMailAddress>`,
    `<AcceptableResponseSchema>${RESPONSE_SCHEMA}</AcceptableResponseSchema>`,
    `</Request>`,
    `</Autodiscover>`,
  ].join("\r\n");
}

function escapeXml(s) {
  return String(s).replace(
    /[<>&'"]/g,
    (c) =>
      ({
        "<": "&lt;",
        ">": "&gt;",
        "&": "&amp;",
        "'": "&apos;",
        '"': "&quot;",
      })[c],
  );
}

function sleep(ms, signal) {
  return new Promise((resolve, reject) => {
    if (!ms) return resolve();
    const timer = setTimeout(resolve, ms);
    if (signal) {
      const onAbort = () => {
        clearTimeout(timer);
        reject(makeCancel());
      };
      if (signal.aborted) {
        clearTimeout(timer);
        reject(makeCancel());
        return;
      }
      signal.addEventListener("abort", onAbort, { once: true });
    }
  });
}

function makeCancel() {
  return new DiscoverError(DISCOVER_ERR.CANCELLED, "Discovery cancelled");
}
