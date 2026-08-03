/**
 * Wraps a node built by `parseAdNode` (or any `{tagName, textContent,
 * children}` tree) into the shape `network.mjs`'s real decoded `doc`
 * has: a `documentElement` plus a `getElementsByTagName` that does a
 * recursive descendant search. Used to build canned `easRequest`
 * responses for integration tests that mock the network seam.
 */
export function toMockDoc(rootNode) {
  return {
    documentElement: rootNode,
    getElementsByTagName(tag) {
      const out = [];
      (function walk(node) {
        if (!node) return;
        if (node.tagName === tag) out.push(node);
        for (const c of node.children ?? []) walk(c);
      })(rootNode);
      return out;
    },
  };
}
