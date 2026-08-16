# Patches for Thunderbird itself

Fixes this add-on needs that live in Thunderbird's own code, not in ours or
in the vendored experiment. They cannot be shipped from an add-on; they are
kept here so the diagnosis is not lost and so the change can be offered
upstream.

Apply with `patch -p1` from the root of a `comm` checkout.

---

## thunderbird-autocomplete-empty-async-result.patch

**File:** `mailnews/addrbook/src/AbAutoCompleteSearch.sys.mjs`
**Against:** comm-central `4acf91d731b7` (2026-08-16)
**Symptom:** [EAS-4-TbSync#344] — the Global Address List returns nothing
while typing in a compose window, however many characters are typed.

### What goes wrong

A provider address book is an `ASYNC_DIRECTORY_TYPE`, so the autocomplete
searches it through `dir.search(...)` and takes its answer in
`onSearchFinished(status, isCompleteResult)`.

`isCompleteResult: false` means "this answer is not the whole match set, ask
me again for a longer string". The search records that by putting the
directory on `result.asyncDirectories`, which is the list the *next* search
re-queries. But that line sits inside `if (cards.length)`:

    if (cards.length) {
      ...
      if (!isCompleteResult) {
        result.asyncDirectories.push(dir);
      }
    }

So a directory that answers **empty and incomplete** is never recorded. And
the next keystroke cannot recover, because the reuse path does not rebuild
the list from the address book manager - it *replaces* it:

    asyncDirectories = aPreviousResult.asyncDirectories;

With that list empty, the search returns at

    if (!asyncDirectories.length) {
      // We're done. Just return our result immediately.

before any directory is searched. The provider is never asked again.

The reuse path is only taken when the previous result was
`RESULT_SUCCESS` - which requires at least one result from somewhere. A
*local* card matching the same prefix is enough. With no local match the
result is `RESULT_NOMATCH`, reuse is skipped, and the same typing works.
That is why the bug looks intermittent and why it is easy to miss in
testing.

### Why an add-on hits this

An Exchange GAL declines queries shorter than four characters: it answers
three with a bare `<Result/>` and no `<Total>`. The provider cannot claim
that is the whole match set, so it correctly answers empty-and-incomplete -
and that is exactly the answer whose flag is discarded.

### Reproduction

Measured against Exchange Online, same account and same typing, with the
only difference being one local contact:

| local contact matching the prefix | Search requests | result |
| --- | --- | --- |
| no  | 2 (`cvj`, `cvjm`) | 3 hits |
| yes | 1 (`cvj` only)    | no hits |

Steps:

1. An account whose GAL needs four characters (Exchange Online does).
2. Type `cvj` in a compose address field - the provider is asked, the server
   declines, the provider answers empty and incomplete.
3. Extend to `cvjm`. The provider is not asked. No GAL results appear.
4. Add a contact to a local address book whose name matches `cvj`, and the
   behaviour flips: without it the extension works, with it it does not.

Holding a breakpoint in the provider's callback also makes it work, for the
same reason from the other side: the search never finishes, so the previous
result is `RESULT_SUCCESS_ONGOING` rather than `RESULT_SUCCESS`, the reuse
condition fails, and the fresh path runs.

### The fix

Hoist the `if (!isCompleteResult)` out of the `cards.length` test, so an
incomplete answer is recorded whether or not it carried any cards.

### Status

Not upstreamed yet. Until it is, the add-on works around it in
`src/modules/gal.mjs` by withholding an empty answer instead of returning
it - the search then stays `RESULT_SUCCESS_ONGOING` and the reuse path is
not taken. That workaround exploits an implementation detail rather than a
contract and should be removed once this lands.

[EAS-4-TbSync#344]: https://github.com/jobisoft/EAS-4-TbSync/issues/344
