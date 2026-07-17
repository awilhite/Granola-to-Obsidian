# Native Granola Template Generation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace malformed Markdown panel writeback with Granola's native `generate-summary` flow, including persisted-structure verification and retryable failure cleanup.

**Architecture:** Keep the private Granola client and sync orchestration in the shipped `main.js`. The client will create a panel, submit the native Yjs-backed stream request, parse event telemetry, poll Granola for persisted structured content, and soft-delete an empty failed panel. The sync layer will consume only the refreshed persisted panel and will continue normal Obsidian sync on template failure.

**Tech Stack:** Obsidian CommonJS plugin, JavaScript, Granola private HTTP APIs, base64 Yjs v1 seed, Node.js built-in test runner, VM-based source extraction.

## Global Constraints

- Use `generate-summary` as the only automatic template-generation endpoint.
- Never fall back to `llm-proxy-stream` or write generated Markdown directly to a Granola panel.
- Preserve malformed-summary normalization for legacy panels only.
- Continue normal sync when template generation fails.
- Delete a newly created empty panel after generation or verification failure so a later sync can retry.
- Persisted panel state, not stream events, is the source of truth.
- Do not log credentials, transcript text, generated prose, or full private API payloads.
- Do not add a runtime dependency or a new user-facing setting.

---

### Task 1: Native Generation Client Contract

**Files:**
- Modify: `main.js`
- Create: `tests/granola-template-generation.test.js`

**Interfaces:**
- Produces: `GranolaPrivateClient.generateDocumentPanel(document, metadata, transcriptEntries, template, panelId, options)` returning `{ eventTypes, streamedContentLength }`.
- Produces: `GranolaPrivateClient.parseGenerateSummaryStream(streamText)` returning telemetry without generated text.
- Consumes: Existing `buildPromptVariables`, `postText`, template objects, and authenticated private-client headers.

- [ ] **Step 1: Add a VM test harness and failing native-request test**

Create a test helper that extracts `GranolaPrivateClient` from `main.js`, evaluates it with mocked `obsidian.requestUrl`, `os`, `process`, and native constants, then assert the request:

```js
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const mainJs = fs.readFileSync(path.resolve(__dirname, '..', 'main.js'), 'utf8');

function extractBlock(source, marker) {
	const start = source.indexOf(marker);
	assert.notEqual(start, -1, `Missing ${marker}`);
	const bodyStart = source.indexOf('{', start);
	let depth = 0;
	for (let index = bodyStart; index < source.length; index += 1) {
		if (source[index] === '{') depth += 1;
		if (source[index] === '}') depth -= 1;
		if (depth === 0) return source.slice(start, index + 1);
	}
	throw new Error(`Unclosed ${marker}`);
}

function loadPrivateClient(responses) {
	const requests = [];
	const responseQueue = [...responses];
	const context = vm.createContext({
		Intl,
		Date,
		JSON,
		String,
		Array,
		Object,
		RegExp,
		Promise,
		setTimeout,
		process: { platform: 'darwin' },
		os: { release: () => 'test-os' },
		obsidian: {
			requestUrl: async (request) => {
				requests.push(request);
				const response = responseQueue.shift();
				if (response instanceof Error) throw response;
				return response;
			},
		},
	});
	const source = `
		const GRANOLA_TEMPLATE_CLIENT_VERSION = 'test-client';
		const GRANOLA_TEMPLATE_YDOC_STATE = 'AQLW64i1DgAHAQtwcm9zZW1pcnJvcgMJcGFyYWdyYXBoKAEEbWV0YQdoYXNTZWVkAXgA';
		const GRANOLA_TEMPLATE_YDOC_VERSION = 1;
		const GRANOLA_GENERATION_POLL_ATTEMPTS = 5;
		const GRANOLA_GENERATION_POLL_DELAY_MS = 1500;
		${extractBlock(mainJs, 'class GranolaPrivateClient')}
		GranolaPrivateClient;
	`;
	return { Client: vm.runInContext(source, context), requests };
}

function authContext() {
	return {
		token: 'access-token',
		clientVersion: '7.427.3',
		platform: 'darwin',
		osVersion: 'test-os',
		workspaceId: 'workspace-1',
		deviceId: 'device-1',
	};
}

function nativeStreamFixture() {
	return [
		JSON.stringify({ panel_id: 'panel-1' }),
		JSON.stringify({ choices: [{ delta: { content: '<h3>Metadata</h3>' } }] }),
		JSON.stringify({ generated_lines: [1] }),
		JSON.stringify({ ydoc_state: 'state', ydoc_version: 1 }),
	].join('-----CHUNK_BOUNDARY-----');
}

test('generateDocumentPanel sends the native Yjs-backed request', async () => {
	const { Client, requests } = loadPrivateClient([
		{ text: nativeStreamFixture() },
	]);
	const client = new Client(authContext());

	const result = await client.generateDocumentPanel(
		{ id: 'doc-1', title: 'Testing', created_at: '2026-07-17T20:04:20.170Z' },
		{},
		[{ source: 'microphone', text: 'Test transcript' }],
		{ id: 'template-1', title: 'Default', sections: [] },
		'panel-1',
		{ auto: true }
	);

	assert.equal(requests[0].url, 'https://stream.api.granola.ai/v1/generate-summary');
	const body = JSON.parse(requests[0].body);
	assert.equal(body.document_id, 'doc-1');
	assert.equal(body.panel_id, 'panel-1');
	assert.equal(body.panel_title, 'Default');
	assert.equal(body.prompt_slug, 'template-summary-consolidated');
	assert.equal(body.template_slug, 'template-1');
	assert.equal(body.ydoc_version, 1);
	assert.match(body.ydoc_state, /^[A-Za-z0-9+/]+=*$/);
	assert.equal(body.auto, true);
	assert.deepEqual(result.eventTypes, ['panel_id', 'content_delta', 'generated_lines', 'ydoc_state']);
});
```

- [ ] **Step 2: Run the focused test and verify failure**

Run: `node --test tests/granola-template-generation.test.js`

Expected: FAIL because `generateDocumentPanel` and `parseGenerateSummaryStream` do not exist.

- [ ] **Step 3: Add native constants and stream telemetry parsing**

Add constants beside the existing template constants:

```js
const GRANOLA_TEMPLATE_YDOC_STATE = 'AQLW64i1DgAHAQtwcm9zZW1pcnJvcgMJcGFyYWdyYXBoKAEEbWV0YQdoYXNTZWVkAXgA';
const GRANOLA_TEMPLATE_YDOC_VERSION = 1;
const GRANOLA_GENERATION_POLL_ATTEMPTS = 5;
const GRANOLA_GENERATION_POLL_DELAY_MS = 1500;
```

Replace legacy stream-content collection with telemetry parsing that recognizes `panel_id`, `ydoc_state`, `generated_lines`, and content-delta chunks while retaining only content length.

- [ ] **Step 4: Implement the native request**

Implement `generateDocumentPanel` using:

```js
const streamText = await this.postText(
	'https://stream.api.granola.ai/v1/generate-summary',
	{
		document_id: document.id,
		panel_id: panelId,
		panel_title: template.title || 'Summary',
		prompt_slug: 'template-summary-consolidated',
		prompt_variables: this.buildPromptVariables(document, metadata, transcriptEntries, template),
		chat_history: [],
		template_slug: template.id,
		ydoc_state: GRANOLA_TEMPLATE_YDOC_STATE,
		ydoc_version: GRANOLA_TEMPLATE_YDOC_VERSION,
		auto: options.auto === true,
	},
	{ accept: '*/*' }
);
return this.parseGenerateSummaryStream(streamText);
```

Remove `generateTemplateMarkdown`, `collectStreamContent`, and `stripNotesWrapper` only after confirming they have no remaining callers.

- [ ] **Step 5: Run the focused test and syntax check**

Run: `node --test tests/granola-template-generation.test.js`

Expected: PASS.

Run: `node -c main.js`

Expected: exit 0 with no output.

- [ ] **Step 6: Commit the client contract**

```bash
git add main.js tests/granola-template-generation.test.js
git commit -m "feat: call Granola native summary generation"
```

---

### Task 2: Persisted Panel Verification And Cleanup

**Files:**
- Modify: `main.js`
- Modify: `tests/granola-template-generation.test.js`

**Interfaces:**
- Produces: `GranolaPrivateClient.waitForGeneratedPanel(documentId, panelId, templateId, options)` returning the persisted active panel or throwing after bounded polling.
- Produces: `GranolaPrivateClient.deleteDocumentPanel(panelId)` performing a soft delete.
- Changes: `getDocumentPanels(documentId, options)` accepts `{ includeYdocState }`.
- Consumes: Task 1 native request and existing authenticated JSON API methods.

- [ ] **Step 1: Write failing persisted-state tests**

Add tests for these exact cases:

```js
test('waitForGeneratedPanel returns only a structured persisted matching panel', async () => {
	const client = clientWithPanelResponses([
		[{ id: 'panel-1', template_slug: 'template-1', content_updated_at: null, content: '' }],
		[{ id: 'panel-1', template_slug: 'template-1', content_updated_at: '2026-07-17T20:07:05.566Z', ydoc_state: 'state', content: { type: 'doc', content: [{ type: 'heading' }] } }],
	]);
	const panel = await client.waitForGeneratedPanel('doc-1', 'panel-1', 'template-1', { attempts: 2, delayMs: 0 });
	assert.equal(panel.id, 'panel-1');
});

test('waitForGeneratedPanel rejects deleted, wrong-template, and unstructured panels', async () => {
	const client = clientWithPanelResponses([[]]);
	await assert.rejects(
		client.waitForGeneratedPanel('doc-1', 'panel-1', 'template-1', { attempts: 1, delayMs: 0 }),
		/Persisted Granola template panel was not ready/
	);
});

test('deleteDocumentPanel soft deletes without writing content', async () => {
	const { client, request } = clientForDelete();
	await client.deleteDocumentPanel('panel-1');
	const body = JSON.parse(request.body);
	assert.equal(body.id, 'panel-1');
	assert.equal(body.was_trashed, true);
	assert.ok(body.deleted_at);
	assert.equal(Object.hasOwn(body, 'content'), false);
});
```

- [ ] **Step 2: Run the focused tests and verify failure**

Run: `node --test tests/granola-template-generation.test.js`

Expected: FAIL because persisted verification and cleanup methods do not exist.

- [ ] **Step 3: Implement option-aware panel fetch and readiness validation**

Send `include_ydoc_state: true` only when requested. Treat a panel as ready only when it is active, matches both IDs, has `content_updated_at`, has a ProseMirror document with at least one content node, and has non-empty string `ydoc_state`.

Poll at most `GRANOLA_GENERATION_POLL_ATTEMPTS`, sleeping only between attempts. Tests pass `{ delayMs: 0 }` to remain deterministic and fast.

- [ ] **Step 4: Implement soft-delete cleanup**

Add `deleteDocumentPanel(panelId)` that calls `update-document-panel` with:

```js
{
	id: panelId,
	deleted_at: now,
	updated_at: now,
	was_trashed: true,
}
```

Remove the old content-writing `updateDocumentPanel(panelId, content)` method after its caller is removed in Task 3.

- [ ] **Step 5: Run focused tests and syntax check**

Run: `node --test tests/granola-template-generation.test.js`

Expected: PASS.

Run: `node -c main.js`

Expected: exit 0.

- [ ] **Step 6: Commit verification and cleanup**

```bash
git add main.js tests/granola-template-generation.test.js
git commit -m "feat: verify and clean up generated panels"
```

---

### Task 3: Sync Orchestration And Retry Semantics

**Files:**
- Modify: `main.js`
- Modify: `tests/granola-template-generation.test.js`

**Interfaces:**
- Changes: `GranolaSyncPlugin.ensureGranolaTemplateForDocument(doc, authContext)` creates the panel before generation and consumes persisted panel state.
- Consumes: `createDocumentPanel`, `generateDocumentPanel`, `waitForGeneratedPanel`, and `deleteDocumentPanel` from Tasks 1-2.
- Preserves: Existing `attempted`, `applied`, `failed`, and `skipped` counters and non-blocking return of `doc`.

- [ ] **Step 1: Write failing orchestration tests**

Extract `ensureGranolaTemplateForDocument` into a VM object and invoke it with a fake plugin context. Cover:

```js
function extractMethod(methodName, consoleObject = console) {
	const marker = `async ${methodName}(`;
	return vm.runInNewContext(`({ ${extractBlock(mainJs, marker)} })`, {
		console: consoleObject,
		Date,
		Promise,
	});
}

function templatePluginFixture({ outcome }) {
	const calls = [];
	const errors = [];
	const sourceDoc = { id: 'doc-1', title: 'Testing', updated_at: 'before' };
	const persistedPanel = {
		id: 'panel-1',
		template_slug: 'template-1',
		content_updated_at: '2026-07-17T20:07:05.566Z',
		ydoc_state: 'state',
		content: { type: 'doc', content: [{ type: 'heading' }] },
	};
	const client = {
		getDocumentPanels: async () => { calls.push('getPanels'); return []; },
		getDocumentBatch: async () => {
			if (!calls.includes('createPanel')) { calls.push('getContext'); return sourceDoc; }
			calls.push('refreshDocument');
			return { ...sourceDoc, updated_at: 'after' };
		},
		getDocumentMetadata: async () => ({}),
		getDocumentTranscript: async () => [],
		createDocumentPanel: async () => { calls.push('createPanel'); return { id: 'panel-1' }; },
		generateDocumentPanel: async () => {
			calls.push('generate');
			if (outcome !== 'success') throw new Error('generation failed');
		},
		waitForGeneratedPanel: async () => { calls.push('waitForPanel'); return persistedPanel; },
		deleteDocumentPanel: async () => {
			calls.push('deletePanel');
			if (outcome === 'generation-and-cleanup-failure') throw new Error('cleanup failed');
		},
	};
	const testConsole = {
		...console,
		error: (...args) => errors.push(args.map((value) => value instanceof Error ? value.message : String(value)).join(' ')),
	};
	const method = extractMethod('ensureGranolaTemplateForDocument', testConsole);
	const plugin = {
		...method,
		settings: { granolaTemplateId: 'template-1' },
		templateManagementStats: { attempted: 0, applied: 0, failed: 0, skipped: 0 },
		activeSyncDiagnostics: { source: 'manual', templateStats: {} },
		shouldUseGranolaTemplateManagement: () => true,
		getGranolaPrivateClient: () => client,
		getGranolaTemplatePanel: () => null,
		getPanelMarkdownContent: () => '',
		fetchGranolaTemplates: async () => [{ id: 'template-1', title: 'Default', sections: [] }],
		updateActiveSyncProgress: () => {},
	};
	return { plugin, calls, errors, sourceDoc, persistedPanel };
}

test('template orchestration uses persisted structured panel content', async () => {
	const { plugin, calls, persistedPanel } = templatePluginFixture({ outcome: 'success' });
	const result = await plugin.ensureGranolaTemplateForDocument({ id: 'doc-1', title: 'Testing' }, authContext());
	assert.deepEqual(calls, ['getPanels', 'getContext', 'createPanel', 'generate', 'waitForPanel', 'refreshDocument']);
	assert.equal(result.privatePanels[0], persistedPanel);
	assert.equal(result.granolaTemplateManagementMarkdown, undefined);
	assert.equal(plugin.templateManagementStats.applied, 1);
});

test('generation failure deletes the new panel and returns the source document', async () => {
	const { plugin, calls, sourceDoc } = templatePluginFixture({ outcome: 'generation-failure' });
	const result = await plugin.ensureGranolaTemplateForDocument(sourceDoc, authContext());
	assert.ok(calls.includes('deletePanel'));
	assert.equal(result, sourceDoc);
	assert.equal(plugin.templateManagementStats.failed, 1);
});

test('cleanup failure does not replace the original generation error', async () => {
	const { plugin, errors } = templatePluginFixture({ outcome: 'generation-and-cleanup-failure' });
	await plugin.ensureGranolaTemplateForDocument({ id: 'doc-1', title: 'Testing' }, authContext());
	assert.match(errors.join('\n'), /generation failed/);
	assert.match(errors.join('\n'), /cleanup failed/);
});
```

Retain an existing-panel test asserting no create or generate call occurs.

- [ ] **Step 2: Run focused tests and verify failure**

Run: `node --test tests/granola-template-generation.test.js`

Expected: FAIL because orchestration still generates Markdown before panel creation and directly writes it.

- [ ] **Step 3: Replace the orchestration sequence**

Change the missing-template path to:

```js
let createdPanel = null;
try {
	createdPanel = await client.createDocumentPanel(doc.id, selectedTemplate.id, selectedTemplate.title || 'Summary');
	await client.generateDocumentPanel(
		batchDoc || doc,
		metadata || {},
		transcriptEntries || [],
		selectedTemplate,
		createdPanel.id,
		{ auto: this.activeSyncDiagnostics?.source === 'auto' }
	);
	const persistedPanel = await client.waitForGeneratedPanel(doc.id, createdPanel.id, selectedTemplate.id);
	const refreshedDoc = await client.getDocumentBatch(doc.id);
	doc.privatePanels = [persistedPanel, ...doc.privatePanels.filter((panel) => panel.id !== persistedPanel.id)];
	if (refreshedDoc?.updated_at) doc.updated_at = refreshedDoc.updated_at;
} catch (error) {
	if (createdPanel?.id) {
		try {
			await client.deleteDocumentPanel(createdPanel.id);
		} catch (cleanupError) {
			console.error('Granola Template Management cleanup failed for "' + (doc.title || doc.id) + '":', cleanupError);
		}
	}
	throw error;
}
```

Keep the outer failure handler non-blocking. Update progress phases to `Creating template panel`, `Generating`, and `Verifying generated template`. Log stage and durations, never payload content.

- [ ] **Step 4: Remove the legacy write path and verify no callers remain**

Run: `rg -n "llm-proxy-stream|generateTemplateMarkdown|updateDocumentPanel|granolaTemplateManagementMarkdown" main.js`

Expected: no legacy generation/write caller. `granolaTemplateManagementMarkdown` may remain only if another legacy read path still requires it; otherwise remove it.

- [ ] **Step 5: Run focused and full regression suites**

Run: `node --test tests/granola-template-generation.test.js`

Expected: PASS.

Run: `node --test tests/*.test.js`

Expected: all tests pass, including auth and malformed-summary normalization.

Run: `node -c main.js`

Expected: exit 0.

- [ ] **Step 6: Commit sync orchestration**

```bash
git add main.js tests/granola-template-generation.test.js
git commit -m "fix: use native Granola template generation"
```

---

### Task 4: Documentation And Static Validation

**Files:**
- Modify: `readme.md`
- Modify: `CHANGELOG.md`

**Interfaces:**
- Documents: Native generation behavior, retry semantics, private-API caveat, and legacy normalization scope.

- [ ] **Step 1: Update feature documentation**

In the Granola Template Management section, state that the plugin:

- creates a missing selected-template panel
- asks Granola's native `generate-summary` flow to populate it
- verifies persisted structured content before import
- deletes an empty failed panel so later sync can retry
- continues syncing current content after failure

Retain the experimental/fork-specific private-API warning.

- [ ] **Step 2: Update the changelog**

Add an Unreleased Changed entry:

```markdown
- **Native template generation**: Replaced direct Markdown panel writeback with Granola's Yjs-backed `generate-summary` flow, preventing automatically selected templates from being stored as collapsed ProseMirror paragraphs and making failed generation retryable.
```

- [ ] **Step 3: Run static and regression validation**

Run: `git diff --check`

Expected: no output.

Run: `node -c main.js`

Expected: exit 0.

Run: `node --test tests/*.test.js`

Expected: all tests pass.

- [ ] **Step 4: Commit documentation**

```bash
git add readme.md CHANGELOG.md
git commit -m "docs: explain native template generation"
```

---

### Task 5: Controlled Live Validation

**Files:**
- No source changes expected.
- Inspect: Obsidian developer console and vault output.
- Inspect: Granola panels through authenticated read-only API calls.

**Interfaces:**
- Validates: The shipped `main.js` through the live symlinked Obsidian plugin installation.

- [ ] **Step 1: Confirm live plugin linkage and settings**

Verify the live plugin resolves to this repository, Template Management is enabled, `Default` is selected by ID, and auto-sync is five minutes. Do not rewrite settings.

- [ ] **Step 2: Run controlled manual generation**

Use a newly created meeting with a finalized transcript and no active Default panel. Run manual sync once.

Expected:

- exactly one active Default panel exists
- the panel contains structured ProseMirror nodes
- `original_content` is HTML
- Yjs state is non-empty
- document `updated_at` advances
- the Obsidian note contains readable headings/lists and mapped metadata

- [ ] **Step 3: Validate idempotency**

Run manual sync again.

Expected: the existing Default panel is skipped; no second panel or generation request is created.

- [ ] **Step 4: Validate scheduled execution**

Create or identify another eligible meeting, allow one five-minute auto-sync cycle, and inspect diagnostics.

Expected: automatic generation completes without user interaction and normal sync remains responsive.

- [ ] **Step 5: Final verification report**

Record:

- test command results
- live document and panel IDs
- persisted node types and timestamps, without content
- idempotency result
- auto-sync result
- any remaining private-API risks

Do not merge into `main`; report the branch as ready for merge review.
