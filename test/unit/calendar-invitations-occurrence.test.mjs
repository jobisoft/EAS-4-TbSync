/**
 * Answering one occurrence of a recurring invitation.
 *
 * This is the ordinary way of answering, not an edge case: whenever the
 * answer is given from the calendar rather than from the message,
 * Thunderbird writes an override for that occurrence and leaves the master
 * alone. Reading only the master therefore misses it entirely - measured on
 * a live calendar, where a single-occurrence Accept was dropped with
 * "carries no answer to send" and nothing reached the server.
 *
 * Two things have to hold for such an answer to survive the round trip:
 * it must be found on the override, and it must still be there after the
 * pull adopts the server's copy - which knows nothing of it, and runs
 * before the answer is sent.
 */

import { test } from "node:test";
import assert from "node:assert/strict";

import {
  USERRESPONSE_TO_PARTSTAT,
  pinEasStamps,
  preserveSelfPartstat,
  selfUserResponses,
  serverKnownPartstat,
} from "../../src/modules/eas/calendar-codec.mjs";

const ME = "john.bieling@cvjmbonn.de";
const ORG = "john.bieling@outlook.de";

/** A recurring invitation. `master` is the series answer, `overrides` a map
 *  of RECURRENCE-ID to that occurrence's answer. `stamps` gives a component
 *  the server's own ResponseType. */
function series({ master = "NEEDS-ACTION", overrides = {}, stamps = {} } = {}) {
  const lines = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//eas-test//EN",
    "BEGIN:VEVENT",
    "UID:series-1",
    "DTSTART:20260901T090000Z",
    "DTEND:20260901T093000Z",
    "RRULE:FREQ=WEEKLY;COUNT=4",
    `ORGANIZER;CN=Someone:mailto:${ORG}`,
    `ATTENDEE;PARTSTAT=${master};CN=${ME}:mailto:${ME}`,
    `X-MOZ-INVITED-ATTENDEE:mailto:${ME}`,
  ];
  if (stamps.master != null) lines.push(`X-EAS-RESPONSETYPE:${stamps.master}`);
  lines.push("END:VEVENT");
  for (const [rid, partstat] of Object.entries(overrides)) {
    lines.push(
      "BEGIN:VEVENT",
      "UID:series-1",
      `RECURRENCE-ID:${rid}`,
      `DTSTART:${rid}`,
      "DTEND:20260908T093000Z",
      `ORGANIZER;CN=Someone:mailto:${ORG}`,
      `ATTENDEE;PARTSTAT=${partstat};CN=${ME}:mailto:${ME}`,
    );
    if (stamps[rid] != null) lines.push(`X-EAS-RESPONSETYPE:${stamps[rid]}`);
    lines.push("END:VEVENT");
  }
  lines.push("END:VCALENDAR");
  return lines.join("\r\n");
}

test("an answer on one occurrence is found, and names that occurrence", () => {
  const answers = selfUserResponses(
    series({ overrides: { "20260908T090000Z": "ACCEPTED" } }),
    ME,
  );
  assert.equal(answers.length, 1, "the unanswered series is not an answer");
  assert.equal(answers[0].userResponse, 1);
  assert.equal(
    answers[0].instanceId,
    "2026-09-08T09:00:00.000Z",
    "MeetingResponse wants the 24-character form, not the 16-character one",
  );
  assert.equal(answers[0].instanceId.length, 24);
});

test("the master-only reading is what missed it", () => {
  // The record of the bug: the response phase used to ask the master alone,
  // which says NEEDS-ACTION here and maps to no answer at all. The answer
  // is on the override, and only walking every component finds it.
  const ical = series({ overrides: { "20260908T090000Z": "ACCEPTED" } });
  const answers = selfUserResponses(ical, ME);
  assert.equal(answers.length, 1);
  assert.equal(answers[0].instanceId, "2026-09-08T09:00:00.000Z");
  assert.equal(answers.filter((a) => a.instanceId === null).length, 0);
});

test("series and occurrences are answered separately", () => {
  const answers = selfUserResponses(
    series({
      master: "ACCEPTED",
      overrides: {
        "20260908T090000Z": "DECLINED",
        "20260915T090000Z": "TENTATIVE",
      },
    }),
    ME,
  );
  assert.deepEqual(
    answers.map((a) => [a.instanceId, a.userResponse]),
    [
      [null, 1],
      ["2026-09-08T09:00:00.000Z", 3],
      ["2026-09-15T09:00:00.000Z", 2],
    ],
    "the series first, then each occurrence in order",
  );
});

test("an occurrence nobody answered is not an answer", () => {
  assert.deepEqual(
    selfUserResponses(
      series({ overrides: { "20260908T090000Z": "NEEDS-ACTION" } }),
      ME,
    ),
    [],
  );
  assert.deepEqual(selfUserResponses("", ME), []);
  assert.deepEqual(selfUserResponses("not a calendar", ME), []);
});

test("what the server already knows is what stops a second reply", () => {
  // ResponseType comes back on every pull. Without this, an unrelated edit
  // to an answered meeting mails the organizer the same reply again.
  assert.equal(serverKnownPartstat("3"), "ACCEPTED");
  assert.equal(serverKnownPartstat("4"), "DECLINED");
  assert.equal(serverKnownPartstat("2"), "TENTATIVE");
  assert.equal(serverKnownPartstat(null), null, "the server has not said");
  assert.equal(serverKnownPartstat(""), null);
  // 5 is "not responded". It once read as an acceptance, so an item stamped
  // by an older build can disagree with a live ACCEPTED - which costs a
  // duplicate reply, never a lost answer.
  assert.equal(serverKnownPartstat("5"), "NEEDS-ACTION");
  assert.notEqual(serverKnownPartstat("5"), USERRESPONSE_TO_PARTSTAT[1]);

  const answers = selfUserResponses(
    series({
      overrides: { "20260908T090000Z": "ACCEPTED" },
      stamps: { "20260908T090000Z": "3" },
    }),
    ME,
  );
  assert.equal(answers[0].responseType, "3");
  assert.equal(
    serverKnownPartstat(answers[0].responseType),
    USERRESPONSE_TO_PARTSTAT[answers[0].userResponse],
    "server and calendar agree, so this one is not sent again",
  );
});

test("the answer on an occurrence survives adopting the server's copy", () => {
  // The pull runs before the answer is sent. Carrying only the master here
  // is what would lose a per-occurrence answer for good, silently.
  const adopted = preserveSelfPartstat({
    builtIcal: series({ overrides: { "20260908T090000Z": "NEEDS-ACTION" } }),
    priorIcal: series({ overrides: { "20260908T090000Z": "ACCEPTED" } }),
    userEmail: ME,
  });
  const answers = selfUserResponses(adopted, ME);
  assert.equal(answers.length, 1);
  assert.equal(answers[0].userResponse, 1);
  assert.equal(answers[0].instanceId, "2026-09-08T09:00:00.000Z");
});

test("each occurrence keeps its own answer across the adopt", () => {
  const adopted = preserveSelfPartstat({
    builtIcal: series({
      master: "NEEDS-ACTION",
      overrides: {
        "20260908T090000Z": "NEEDS-ACTION",
        "20260915T090000Z": "NEEDS-ACTION",
      },
    }),
    priorIcal: series({
      master: "ACCEPTED",
      overrides: {
        "20260908T090000Z": "DECLINED",
        "20260915T090000Z": "TENTATIVE",
      },
    }),
    userEmail: ME,
  });
  assert.deepEqual(
    selfUserResponses(adopted, ME).map((a) => [a.instanceId, a.userResponse]),
    [
      [null, 1],
      ["2026-09-08T09:00:00.000Z", 3],
      ["2026-09-15T09:00:00.000Z", 2],
    ],
    "answers are matched by RECURRENCE-ID, not merged onto the series",
  );
});

test("an answer the server has and we do not is left alone", () => {
  const adopted = preserveSelfPartstat({
    builtIcal: series({ overrides: { "20260908T090000Z": "ACCEPTED" } }),
    priorIcal: series({ overrides: { "20260908T090000Z": "NEEDS-ACTION" } }),
    userEmail: ME,
  });
  assert.equal(selfUserResponses(adopted, ME)[0].userResponse, 1);
});

test("an occurrence keeps its own server stamps across a local edit", () => {
  // Measured on 16.1: answering one occurrence is recorded as a
  // ResponseType on that exception, while the series carries its own. The
  // guard strips every component before restoring, so restoring only the
  // master dropped the occurrence's stamp on every edit - and the answer
  // then looked unheard, and was sent to the organizer a second time.
  const prior = series({
    master: "TENTATIVE",
    overrides: { "20260908T090000Z": "ACCEPTED" },
    stamps: { master: "2", "20260908T090000Z": "3" },
  });
  const edited = series({
    master: "TENTATIVE",
    overrides: { "20260908T090000Z": "ACCEPTED" },
  });
  const guarded = pinEasStamps({ builtIcal: edited, priorIcal: prior });
  const answers = selfUserResponses(guarded, ME);
  assert.deepEqual(
    answers.map((a) => [a.instanceId, a.responseType]),
    [
      [null, "2"],
      ["2026-09-08T09:00:00.000Z", "3"],
    ],
    "each component got its own stamps back, not the master's",
  );
  for (const a of answers) {
    assert.equal(
      serverKnownPartstat(a.responseType),
      USERRESPONSE_TO_PARTSTAT[a.userResponse],
      "so nothing is answered twice",
    );
  }
});

test("an occurrence with no stamp of its own falls back to the series", () => {
  // The other dialect, measured on 14.1: the answer is recorded on the
  // master and the server's own exception carries nothing.
  const answers = selfUserResponses(
    series({
      master: "ACCEPTED",
      overrides: { "20260908T090000Z": "ACCEPTED" },
      stamps: { master: "3" },
    }),
    ME,
  );
  assert.equal(answers[1].responseType, "3", "inherited from the series");
  assert.equal(
    serverKnownPartstat(answers[1].responseType),
    USERRESPONSE_TO_PARTSTAT[answers[1].userResponse],
  );
});

test("but a disagreement still travels - declining one of an accepted series", () => {
  const answers = selfUserResponses(
    series({
      master: "ACCEPTED",
      overrides: { "20260908T090000Z": "DECLINED" },
      stamps: { master: "3" },
    }),
    ME,
  );
  const occurrence = answers.find((a) => a.instanceId);
  assert.equal(occurrence.userResponse, 3);
  assert.notEqual(
    serverKnownPartstat(occurrence.responseType),
    USERRESPONSE_TO_PARTSTAT[occurrence.userResponse],
    "the server thinks it is accepted, so the decline is sent",
  );
});
