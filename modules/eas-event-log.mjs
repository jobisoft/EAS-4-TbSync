/**
 * Tiny indirection so wire-level code (`network.mjs`, etc.) can emit
 * host event-log entries without importing the provider - which would
 * close a circular import.
 *
 * The provider points the sink at its inherited `reportEventLog`
 * during construction; everything fired before that point silently
 * no-ops, which is the safe default (early-boot misfires shouldn't
 * crash the wire path, and there's no useful sync work happening
 * before the provider's constructor returns anyway).
 *
 * Sink errors are swallowed: a logging failure must never break a
 * live network request.
 */

let sink = null;

export function setEventLogSink(fn) {
  sink = typeof fn === "function" ? fn : null;
}

export function reportEventLog(args) {
  if (!sink) return;
  try {
    sink(args);
  } catch {
    /* never break the wire path */
  }
}
