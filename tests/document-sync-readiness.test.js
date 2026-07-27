const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const mainJs = fs.readFileSync(path.resolve(__dirname, '..', 'main.js'), 'utf8');
const STALE_AGE_MS = 6 * 60 * 60 * 1000;

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

function extractMethod(methodName, context = {}) {
	const marker = `async ${methodName}(`;
	return vm.runInNewContext(`({ ${extractBlock(mainJs, marker)} })`, {
		console,
		Date,
		Promise,
		POST_MEETING_SYNC_DELAY_MS: 2 * 60 * 1000,
		STALE_IN_PROGRESS_MEETING_AGE_MS: STALE_AGE_MS,
		...context,
	});
}

function staleDocument(overrides = {}) {
	return {
		id: 'doc-1',
		title: 'Crosspoint',
		meeting_end_count: 0,
		transcribe: false,
		updated_at: new Date(Date.now() - STALE_AGE_MS - 1000).toISOString(),
		...overrides,
	};
}

function readinessFixture({
	panels = [],
	transcript = [],
	panelError = null,
} = {}) {
	const calls = [];
	let panelOptions = null;
	const client = {
		getDocumentPanels: async (_documentId, options) => {
			calls.push('panels');
			panelOptions = options;
			if (panelError) throw panelError;
			return panels;
		},
		getDocumentTranscript: async () => {
			calls.push('transcript');
			return transcript;
		},
	};
	const method = extractMethod('resolveDocumentSyncReadiness');
	const plugin = {
		...method,
		getDocumentSyncReadiness: (doc) => (
			doc.meeting_end_count === 0
				? { ready: false, reason: 'meeting is still in progress' }
				: { ready: true, reason: '' }
		),
		getGranolaPrivateClient: () => client,
	};
	return { plugin, calls, getPanelOptions: () => panelOptions };
}

test('recent in-progress meetings remain blocked without private verification', async () => {
	const { plugin, calls } = readinessFixture();
	const doc = staleDocument({
		updated_at: new Date(Date.now() - 60 * 1000).toISOString(),
	});

	const result = await plugin.resolveDocumentSyncReadiness(doc, {});

	assert.equal(result.ready, false);
	assert.deepEqual(calls, []);
});

test('stale in-progress meetings recover when a populated panel exists', async () => {
	const panel = {
		id: 'panel-1',
		template_slug: 'meeting-summary',
		content_updated_at: new Date().toISOString(),
		content: { type: 'doc', content: [{ type: 'heading' }] },
	};
	const { plugin, calls, getPanelOptions } = readinessFixture({ panels: [panel] });
	const doc = staleDocument({
		updated_at: new Date(Date.now() - 10 * 60 * 1000).toISOString(),
	});

	const result = await plugin.resolveDocumentSyncReadiness(doc, {});

	assert.equal(result.ready, true);
	assert.equal(result.recoveredStaleMeeting, true);
	assert.equal(result.completionEvidence, 'panel');
	assert.equal(doc.privatePanels[0], panel);
	assert.deepEqual(calls, ['panels']);
	assert.equal(getPanelOptions()?.includeYdocState, true);
});

test('authored notes alone do not recover an actively transcribing stale meeting', async () => {
	const authoredNotesPanel = {
		id: 'panel-my-notes',
		type: 'my_notes',
		content_updated_at: new Date().toISOString(),
		content: { type: 'doc', content: [{ type: 'paragraph' }] },
	};
	const { plugin, calls } = readinessFixture({ panels: [authoredNotesPanel] });

	const result = await plugin.resolveDocumentSyncReadiness(staleDocument({ transcribe: true }), {});

	assert.equal(result.ready, false);
	assert.deepEqual(calls, ['panels']);
});

test('stale stopped meetings recover when a transcript exists without a panel', async () => {
	const { plugin, calls } = readinessFixture({
		transcript: [{ id: 'segment-1', text: 'Completed meeting' }],
	});

	const result = await plugin.resolveDocumentSyncReadiness(staleDocument(), {});

	assert.equal(result.ready, true);
	assert.equal(result.recoveredStaleMeeting, true);
	assert.equal(result.completionEvidence, 'transcript');
	assert.deepEqual(calls, ['panels', 'transcript']);
});

test('stale meetings without completed evidence remain blocked', async () => {
	const { plugin, calls } = readinessFixture();

	const result = await plugin.resolveDocumentSyncReadiness(staleDocument(), {});

	assert.equal(result.ready, false);
	assert.deepEqual(calls, ['panels', 'transcript']);
});

test('private verification failure leaves a stale meeting blocked', async () => {
	const error = new Error('private response content');
	error.status = 503;
	const { plugin, calls } = readinessFixture({ panelError: error });

	const result = await plugin.resolveDocumentSyncReadiness(staleDocument(), {});

	assert.equal(result.ready, false);
	assert.deepEqual(calls, ['panels']);
});
