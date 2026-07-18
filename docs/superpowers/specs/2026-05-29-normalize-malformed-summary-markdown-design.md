# Normalize Malformed Summary Markdown

## Summary

Some Granola meeting summaries arrive at the plugin as raw markdown strings whose structure has already been collapsed before Obsidian sees them. In the current plugin flow, string-backed summary panels are trusted as-is, so malformed content can land in notes with:

- a single-line body
- inline `### Metadata ```json ... ````
- missed metadata mapping into frontmatter
- unreadable bullet and heading formatting

The fix is to add a conservative normalization step for malformed string-backed summary markdown before metadata extraction and note writing. This should repair the malformed shape without altering already well-formed summaries.

## Problem Statement

The current rendering path distinguishes between ProseMirror-backed panels and string-backed panels:

- ProseMirror-backed panels are converted to markdown and already render correctly.
- String-backed panels are returned directly and then passed downstream.

For malformed string-backed summaries, this produces bad output because:

1. the body stays collapsed into a single line
2. the metadata block is inline, so the existing metadata parser does not reliably extract it
3. note rewrites preserve the malformed source instead of repairing it

This is not a migration-only issue. It is a plugin resilience gap that can recur for future notes whenever Granola returns malformed markdown strings.

## Goals

- Repair malformed string-backed enhanced-note markdown before note rendering.
- Preserve well-formed summaries exactly as they are today.
- Restore metadata extraction and frontmatter mapping for malformed notes.
- Make the fix safe enough to use in normal sync runs, not just one-off backfills.

## Non-Goals

- Do not redesign transcript rendering.
- Do not redesign authored-notes rendering.
- Do not introduce a user-facing setting for this behavior.
- Do not attempt aggressive markdown beautification beyond the malformed patterns we have evidence for.

## Approach Options

### Option 1: Normalize malformed string-backed summary markdown before parsing

Add a repair step only for raw summary strings that match the collapsed pattern. Then let the existing metadata extraction and note rendering pipeline continue normally.

Pros:

- fixes both body readability and metadata extraction
- protects future notes automatically
- keeps scope tightly focused on the proven failure mode

Cons:

- requires careful heuristics so we do not mangle legitimate markdown

### Option 2: Only harden metadata extraction

Teach metadata parsing to extract inline metadata JSON even when the summary body stays collapsed.

Pros:

- smaller change

Cons:

- still leaves unreadable note bodies
- does not address the underlying malformed-string rendering issue

### Option 3: Repair affected notes with a one-off script

Backfill bad notes directly in the vault without changing the plugin.

Pros:

- immediate cleanup of current bad notes

Cons:

- recurrence remains unsolved
- operational burden shifts to manual repair

## Recommended Approach

Adopt Option 1.

The malformed summary shape is already reaching the plugin in raw string form. Since we cannot depend on Granola to stop emitting that shape, the plugin should normalize it before further processing. This gives us a resilient path for both present and future notes.

## Design

### New normalization step

Introduce a helper that operates only on string-backed enhanced-note markdown, for example:

- `normalizeMalformedSummaryMarkdown(markdown)`

This helper should:

1. quickly detect whether the markdown appears malformed
2. return the original markdown unchanged when it does not
3. repair only the collapsed patterns we have evidence for

### Detection heuristics

Treat the markdown as malformed only when one or more of these patterns appear:

- inline `### Metadata ```json`
- headings and bullets collapsed into a single line after the metadata block
- very low line count despite clear heading/bullet tokens embedded in the string

Detection should be conservative. If the content is already multiline and readable, do nothing.

### Normalization behavior

When the malformed shape is detected, normalize in this order:

1. Ensure the metadata heading starts on its own line.
2. Ensure the fenced JSON block opens and closes on their own lines.
3. Insert line breaks before subsequent `###` section headings.
4. Insert line breaks before top-level bullet markers that were collapsed inline.
5. Preserve the text content itself as much as possible; this is structure repair, not content rewriting.

The output does not need to be pretty-printed beyond restoring safe markdown structure.

### Integration point

Apply normalization in the string-backed summary path before metadata extraction:

- normalize malformed raw markdown
- then run existing metadata extraction/removal
- then run existing note content assembly

This should be used for:

- selected template panel markdown returned as a string
- `doc.granolaTemplateManagementMarkdown` when it is a string
- raw string content returned from `last_viewed_panel`

It should not affect ProseMirror-backed panels.

### Backfill strategy

After the plugin fix is validated, use controlled rewrites for affected notes rather than standalone vault surgery. That keeps the repaired notes aligned with the plugin’s current rendering rules.

## Testing

### Automated / code-path validation

Add a focused test fixture around malformed summary markdown:

- inline metadata fence on a single line
- multiple `###` headings collapsed inline
- bullet items collapsed inline

Assert that normalization produces:

- multiline metadata fence
- separate headings
- separate bullet lines
- extractable metadata JSON

Also add a control test asserting that already well-formed markdown remains unchanged.

### Live validation

Validate on:

- `Morey` as the primary malformed case
- one additional malformed note from the known affected set
- one already-good note such as `Trillium` or `Powerhouse Dynamics`

Success criteria:

- the malformed note body becomes readable
- metadata is promoted into frontmatter when enabled
- inline metadata block is removed from the body when enabled
- well-formed notes remain unchanged

## Risks

### Over-normalization

If heuristics are too broad, we could rewrite legitimate markdown that only happens to contain unusual inline tokens.

Mitigation:

- keep detection narrow
- require concrete malformed markers before normalization
- add a “well-formed unchanged” test

### Incomplete normalization

If there are malformed variants beyond the current pattern, some notes may remain partially broken.

Mitigation:

- start with the proven pattern
- validate against at least two malformed real notes
- extend only when new evidence appears

## Rollout Plan

1. Implement normalization helper for string-backed enhanced summaries.
2. Add focused tests/fixtures for malformed and well-formed cases.
3. Validate on `Morey`, one additional malformed note, and one good control note.
4. Run controlled rewrites for affected notes after validation.

## Decision

Implement a conservative normalization layer for malformed string-backed enhanced-note markdown before metadata parsing and note writing, and then use controlled plugin rewrites to repair affected notes.
