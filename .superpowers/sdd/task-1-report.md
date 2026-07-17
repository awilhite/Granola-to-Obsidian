# Task 1 Report: Native Generation Client Contract

## Files Changed

- `main.js`
  - Added the four native-generation constants with the exact values from the brief.
  - Added `GranolaPrivateClient.parseGenerateSummaryStream(streamText)`, returning event types and streamed content length without generated text.
  - Added `GranolaPrivateClient.generateDocumentPanel(...)` using the native Yjs-backed `generate-summary` request contract.
  - Preserved `generateTemplateMarkdown`, `collectStreamContent`, and `stripNotesWrapper` because Task 3 still calls the legacy method.
- `tests/granola-template-generation.test.js`
  - Added the VM-based client extraction harness and native request contract test from the brief.

## RED Evidence

Command:

```text
node --test tests/granola-template-generation.test.js
```

Result:

```text
tests 1
pass 0
fail 1
TypeError: client.generateDocumentPanel is not a function
exit_code=1
```

The failure was caused by the missing production method, as expected.

## GREEN Verification

Focused test:

```text
node --test tests/granola-template-generation.test.js
```

```text
tests 1
pass 1
fail 0
exit_code=0
```

Syntax check:

```text
node -c main.js
```

```text
exit_code=0
```

Full test suite:

```text
node --test tests/*.test.js
```

```text
tests 11
pass 11
fail 0
exit_code=0
```

Additional verification: `git diff --check` passed, and the legacy generation method remains referenced by Task 3.

## Commit

- Commit: `742793b`
- Message: `feat: call Granola native summary generation`

## Self-Review

The strongest objection is that the new parser could accidentally become a second content-generation implementation or break the existing Task 3 flow. It does neither: it returns only telemetry, and all legacy methods and their caller remain unchanged. The native request uses the exact endpoint, payload fields, constants, and accept header specified in the brief.

## Concerns

- Polling constants are intentionally added but unused in Task 1; polling belongs to the later task.
- The test covers the request contract and event ordering, but not malformed chunks or an isolated `streamedContentLength` assertion; those are appropriate follow-up coverage for later tasks.
- Pre-existing unrelated worktree files were not staged or changed.

## Review Fix

### Resolution

Malformed non-empty stream chunks now emit `unparsed` through the existing `eventTypes` contract. Parsing continues so the persisted panel remains the source of truth; malformed chunk content is not retained or logged. Unrecognized non-empty parsed telemetry is also reported as `unparsed` rather than silently discarded.

### RED Evidence

Added the malformed-chunk regression test before changing `main.js`.

Command:

```text
node --test tests/granola-template-generation.test.js
```

Exact output:

```text
✔ generateDocumentPanel sends the native Yjs-backed request (25.485666ms)
✖ parseGenerateSummaryStream marks malformed chunks as unparsed without retaining content (2.504125ms)
ℹ tests 2
ℹ suites 0
ℹ pass 1
ℹ fail 1
ℹ cancelled 0
ℹ skipped 0
ℹ todo 0
ℹ duration_ms 166.398458

✖ failing tests:

test at tests/granola-template-generation.test.js:107:1
✖ parseGenerateSummaryStream marks malformed chunks as unparsed without retaining content (2.504125ms)
  AssertionError [ERR_ASSERTION]: Expected values to be strictly deep-equal:
  + actual - expected
  
    [
      'panel_id',
  -   'unparsed',
      'content_delta'
    ]
  
      at TestContext.<anonymous> (/Users/austinwilhite/Projects/granola-sync-plus-plugin/tests/granola-template-generation.test.js:118:9)
      at Test.runInAsyncScope (node:async_hooks:214:14)
      at Test.run (node:internal/test_runner/test:1047:25)
      at Test.postRun (node:internal/test_runner/test:1173:19)
      at async startSubtestAfterBootstrap (node:internal/test_runner/harness:373:17) {
    generatedMessage: true,
    code: 'ERR_ASSERTION',
    actual: [ 'panel_id', 'content_delta' ],
    expected: [ 'panel_id', 'unparsed', 'content_delta' ],
    operator: 'deepStrictEqual',
    diff: 'simple'
  }
exit_code=1
```

### GREEN Evidence

Files changed for the fix:

- `main.js`
- `tests/granola-template-generation.test.js`

Focused test exact output:

```text
✔ generateDocumentPanel sends the native Yjs-backed request (41.366083ms)
✔ parseGenerateSummaryStream marks malformed chunks as unparsed without retaining content (2.067792ms)
ℹ tests 2
ℹ suites 0
ℹ pass 2
ℹ fail 0
ℹ cancelled 0
ℹ skipped 0
ℹ todo 0
ℹ duration_ms 186.355584
```

Syntax check exact output:

```text
```

`node -c main.js` produced no output and exited with code 0.

### Commit

- Fix commit: `a8e2efc` (`fix: mark malformed generation chunks as unparsed`)

### Self-Review

The strongest objection is that classifying valid-but-unknown JSON as `unparsed` could change telemetry unexpectedly. That is intentional: every non-empty chunk must be explicit under the existing event contract, and retaining silent omission would recreate the review finding. The regression test verifies the malformed chunk becomes `unparsed`, parsing continues to later content, the returned telemetry does not contain the malformed text, and `streamedContentLength` remains the length of recognized fixture content only.

### Concerns

None for this review fix. Existing unrelated worktree files remain unstaged and unchanged.
