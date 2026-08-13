/**
 * `Settings/Oof` - reading the out-of-office reply back.
 *
 * The first fixture is a verbatim capture from a live Exchange mailbox,
 * kept as the server sent it. An earlier version of the parser started
 * its walk at the `Document` rather than at `documentElement`, found no
 * `Oof` block in a reply that plainly had one, and told the user their
 * server offered no out-of-office at all. Every server looked unsupported.
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
installWebextEnv();
import { buildSetBody, readOofFromDoc } from "../../src/modules/eas/oof.mjs";
import { decodeWBXML } from "../../src/modules/wbxml.mjs";
import { parseAdNode } from "./support/ad-node.mjs";

const domLike = (node) => {
  node.children = node.children ?? [];
  for (const child of node.children) domLike(child);
  return node;
};

const doc = (xml) => ({ documentElement: domLike(parseAdNode(xml)) });

test("a live Exchange reply is parsed, not mistaken for an unsupported server", () => {
  const oof = readOofFromDoc(
    doc(`<Settings>
      <Status>1</Status>
      <Oof>
        <Status>1</Status>
        <Get>
          <OofState>0</OofState>
          <StartTime>2026-08-13T18:00:00.000Z</StartTime>
          <EndTime>2026-08-14T18:00:00.000Z</EndTime>
          <OofMessage><AppliesToInternal/><Enabled>0</Enabled></OofMessage>
          <OofMessage><AppliesToExternalKnown/><Enabled>0</Enabled></OofMessage>
          <OofMessage><AppliesToExternalUnknown/><Enabled>0</Enabled></OofMessage>
        </Get>
      </Oof>
    </Settings>`),
  );
  assert.ok(oof, "the block was found");
  assert.equal(oof.state, "0");
  assert.equal(oof.startTime, "2026-08-13T18:00:00.000Z");
  assert.equal(oof.endTime, "2026-08-14T18:00:00.000Z");
  for (const name of ["internal", "externalKnown", "externalUnknown"]) {
    assert.equal(oof.messages[name].enabled, false, name);
    assert.equal(oof.messages[name].reply, "", name);
  }
});

test("each audience keeps its own switch, text and format", () => {
  const oof = readOofFromDoc(
    doc(`<Settings>
      <Oof>
        <Get>
          <OofState>2</OofState>
          <OofMessage>
            <AppliesToInternal/><Enabled>1</Enabled>
            <ReplyMessage>Back Monday.</ReplyMessage><BodyType>Text</BodyType>
          </OofMessage>
          <OofMessage>
            <AppliesToExternalKnown/><Enabled>1</Enabled>
            <ReplyMessage>&lt;p&gt;Away.&lt;/p&gt;</ReplyMessage><BodyType>HTML</BodyType>
          </OofMessage>
          <OofMessage><AppliesToExternalUnknown/><Enabled>0</Enabled></OofMessage>
        </Get>
      </Oof>
    </Settings>`),
  );
  assert.equal(oof.state, "2");
  assert.deepEqual(oof.messages.internal, {
    enabled: true,
    reply: "Back Monday.",
    bodyType: "Text",
  });
  assert.equal(oof.messages.externalKnown.enabled, true);
  assert.equal(
    oof.messages.externalKnown.bodyType,
    "HTML",
    "reported so the dialog can warn before flattening it",
  );
  assert.equal(oof.messages.externalUnknown.enabled, false);
  assert.equal(oof.messages.externalUnknown.reply, "");
});

test("an audience the server does not name is absent, not switched off", () => {
  // Verbatim from Exchange 14.1, which offers the internal audience only.
  // The dialog reads this as "no external reply here" and drops the
  // control; reporting a disabled one would invite the user to write a
  // message the server accepts and never stores.
  const oof = readOofFromDoc(
    doc(`<Settings>
      <Status>1</Status>
      <Oof>
        <Status>1</Status>
        <Get>
          <OofState>1</OofState>
          <OofMessage>
            <AppliesToInternal/><Enabled>1</Enabled>
            <ReplyMessage>Ich bin nicht da.</ReplyMessage><BodyType>Text</BodyType>
          </OofMessage>
        </Get>
      </Oof>
    </Settings>`),
  );
  assert.deepEqual(Object.keys(oof.messages), ["internal"]);
  assert.equal(oof.messages.internal.enabled, true);
  assert.equal(oof.messages.externalKnown, undefined);
  assert.equal(oof.messages.externalUnknown, undefined);
});

test("a reply with no Oof block is the one case that means unsupported", () => {
  assert.equal(
    readOofFromDoc(doc(`<Settings><Status>1</Status></Settings>`)),
    null,
  );
  assert.equal(
    readOofFromDoc(doc(`<Settings><Oof><Status>1</Status></Oof></Settings>`)),
    null,
    "an Oof without Get",
  );
  assert.equal(readOofFromDoc(null), null, "no document at all");
});

/** The `<Set>` body as XML, so the assertions read like the wire. */
const setBody = (settings) =>
  decodeWBXML(buildSetBody(settings)).replace(/^<\?xml[^?]*\?>/, "");

test("switching an audience off keeps its text on the wire", () => {
  // Measured on a 16.1 mailbox: disable the internal audience without
  // sending its ReplyMessage and the server throws the message away, while
  // a 14.1 one keeps it. The user only asked to stop replying.
  const xml = setBody({
    state: "0",
    messages: {
      internal: { enabled: false, reply: "Ich bin nicht da." },
      externalKnown: { enabled: false, reply: "Away." },
      externalUnknown: { enabled: false, reply: "Away." },
    },
  });
  assert.match(xml, /<OofState>0<\/OofState>/);
  assert.equal(xml.match(/<Enabled>0<\/Enabled>/g).length, 3);
  assert.match(
    xml,
    /<AppliesToInternal\/><Enabled>0<\/Enabled><ReplyMessage>Ich%20bin%20nicht%20da.<\/ReplyMessage>/,
  );
  assert.equal(xml.match(/<ReplyMessage>/g).length, 3, "all three carried");
});

test("an audience we hold no text for is sent bare, not emptied", () => {
  // An empty body erases; absence leaves whatever the mailbox has.
  const xml = setBody({
    state: "0",
    messages: { internal: { enabled: false, reply: "" } },
  });
  assert.match(xml, /<AppliesToInternal\/><Enabled>0<\/Enabled><\/OofMessage>/);
  assert.equal(xml.includes("ReplyMessage"), false);
});

test("an empty box never travels, not even for an enabled audience", () => {
  // Measured on 16.1: a stored internal reply is not reported while out of
  // office is off, so the box comes up empty over a message that is still
  // there. Sending that emptiness on the next save would erase it.
  for (const enabled of [true, false]) {
    const xml = setBody({
      state: enabled ? "1" : "0",
      messages: { internal: { enabled, reply: "" } },
    });
    assert.equal(xml.includes("ReplyMessage"), false, `enabled=${enabled}`);
  }
});

test("the window is sent only in state 2, where it means anything", () => {
  const args = {
    startTime: "2026-08-13T18:00:00.000Z",
    endTime: "2026-08-14T18:00:00.000Z",
    messages: {},
  };
  assert.equal(setBody({ ...args, state: "1" }).includes("<StartTime>"), false);
  assert.match(setBody({ ...args, state: "2" }), /<StartTime>.+<\/StartTime>/);
});
