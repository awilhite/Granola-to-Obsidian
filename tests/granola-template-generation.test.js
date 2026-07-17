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

function extractMethod(methodName, consoleObject = console) {
	const marker = `async ${methodName}(`;
	return vm.runInNewContext(`({ ${extractBlock(mainJs, marker)} })`, {
		console: consoleObject,
		Date,
		Promise,
		GRANOLA_GENERATION_STALE_PANEL_AGE_MS: 10 * 60 * 1000,
	});
}

function extractPluginMethod(methodName, context = {}) {
	const marker = `\t${methodName}(`;
	return vm.runInNewContext(`({ ${extractBlock(mainJs, marker)} })`, context);
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

function clientWithPanelResponses(panelResponses) {
	const loaded = loadPrivateClient(panelResponses.map((panels) => ({ json: panels })));
	const client = new loaded.Client(authContext());
	client.requests = loaded.requests;
	return client;
}

function clientForDelete() {
	const loaded = loadPrivateClient([{ json: {} }]);
	const client = new loaded.Client(authContext());
	return {
		client,
		request: {
			get body() {
				return loaded.requests[0] && loaded.requests[0].body;
			},
		},
	};
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

function templatePluginFixture({
	outcome = 'success',
	existingPanel = null,
	panelSnapshots = null,
	source = 'manual',
	createResult = { id: 'panel-1' },
} = {}) {
	const calls = [];
	const logs = [];
	let generateOptions = null;
	const sourceDoc = { id: 'doc-1', title: 'PRIVATE-MEETING-TITLE', updated_at: 'before' };
	const snapshots = panelSnapshots || [existingPanel ? [existingPanel] : []];
	const persistedPanel = {
		id: 'panel-1',
		template_slug: 'template-1',
		content_updated_at: '2026-07-17T20:07:05.566Z',
		ydoc_state: 'state',
		content: { type: 'doc', content: [{ type: 'heading' }] },
	};
	const client = {
		getDocumentPanels: async () => {
			calls.push('getPanels');
			return snapshots[Math.min(calls.filter((call) => call === 'getPanels').length - 1, snapshots.length - 1)];
		},
		getDocumentBatch: async () => {
			if (!calls.includes('createPanel')) { calls.push('getContext'); return sourceDoc; }
			calls.push('refreshDocument');
			if (outcome === 'refresh-failure') throw new Error('refresh failed');
			return { ...sourceDoc, updated_at: 'after' };
		},
		getDocumentMetadata: async () => ({}),
		getDocumentTranscript: async () => [],
		createDocumentPanel: async () => {
			calls.push('createPanel');
			if (outcome === 'create-throws') throw new Error('create response lost');
			return createResult;
		},
		generateDocumentPanel: async (...args) => {
			calls.push('generate');
			generateOptions = args[5];
			if (outcome === 'generation-failure') throw new Error('generation failed');
			if (outcome === 'generation-and-cleanup-failure') throw new Error('generation failed');
		},
		waitForGeneratedPanel: async () => {
			calls.push('waitForPanel');
			if (outcome === 'verification-failure') throw new Error('verification failed');
			return persistedPanel;
		},
		deleteDocumentPanel: async () => {
			calls.push('deletePanel');
			if (outcome === 'generation-and-cleanup-failure') throw new Error('cleanup failed');
			if (outcome === 'stale-cleanup-failure') throw new Error('stale cleanup failed');
		},
	};
	const testConsole = {
		...console,
		log: (...args) => logs.push(args.map((value) => value instanceof Error ? value.message : String(value)).join(' ')),
		error: (...args) => logs.push(args.map((value) => value instanceof Error ? value.message : String(value)).join(' ')),
	};
	const method = extractMethod('ensureGranolaTemplateForDocument', testConsole);
	const plugin = {
		...method,
		settings: { granolaTemplateId: 'template-1' },
		templateManagementStats: { attempted: 0, applied: 0, failed: 0, skipped: 0, deferred: 0 },
		activeSyncDiagnostics: { source, templateStats: {} },
		shouldUseGranolaTemplateManagement: () => true,
		getGranolaPrivateClient: () => client,
		getGranolaTemplatePanel: (panels, templateId) => panels.find((panel) => panel.template_slug === templateId) || null,
		getPanelMarkdownContent: () => '',
		fetchGranolaTemplates: async () => [{ id: 'template-1', title: 'PRIVATE-TEMPLATE-TITLE', sections: [] }],
		updateActiveSyncProgress: () => {},
	};
	return { plugin, calls, logs, sourceDoc, persistedPanel, getGenerateOptions: () => generateOptions };
}

test('template orchestration uses persisted structured panel content', async () => {
	const { plugin, calls, persistedPanel, getGenerateOptions } = templatePluginFixture({ outcome: 'success', source: 'auto' });
	const result = await plugin.ensureGranolaTemplateForDocument({ id: 'doc-1', title: 'Testing' }, authContext());

	assert.deepEqual(calls, ['getPanels', 'getContext', 'createPanel', 'generate', 'waitForPanel', 'refreshDocument']);
	assert.equal(result.privatePanels[0], persistedPanel);
	assert.equal(result.granolaTemplateManagementMarkdown, undefined);
	assert.equal(getGenerateOptions().auto, true);
	assert.equal(result.updated_at, 'after');
	assert.equal(plugin.templateManagementStats.applied, 1);
});

test('generation failure deletes the new panel and returns the source document', async () => {
	const { plugin, calls, sourceDoc } = templatePluginFixture({ outcome: 'generation-failure' });
	const result = await plugin.ensureGranolaTemplateForDocument(sourceDoc, authContext());

	assert.ok(calls.includes('deletePanel'));
	assert.equal(result, sourceDoc);
	assert.equal(plugin.templateManagementStats.failed, 1);
});

test('cleanup failure preserves the non-blocking generation failure outcome', async () => {
	const { plugin, logs } = templatePluginFixture({ outcome: 'generation-and-cleanup-failure' });
	await plugin.ensureGranolaTemplateForDocument({ id: 'doc-1', title: 'Testing' }, authContext());

	assert.equal(plugin.templateManagementStats.failed, 1);
	assert.match(logs.join('\n'), /stage=cleanup status=failed/);
	assert.match(logs.join('\n'), /stage=orchestration status=failed/);
});

test('existing template panel skips native generation', async () => {
	const existingPanel = {
		id: 'panel-existing',
		template_slug: 'template-1',
		content: { type: 'doc', content: [{ type: 'heading' }] },
	};
	const { plugin, calls } = templatePluginFixture({ outcome: 'success', existingPanel });
	const result = await plugin.ensureGranolaTemplateForDocument({ id: 'doc-1', title: 'Testing' }, authContext());

	assert.equal(result.privatePanels[0], existingPanel);
	assert.deepEqual(calls, ['getPanels']);
	assert.equal(plugin.templateManagementStats.skipped, 1);
});

test('manual template generation sends auto false', async () => {
	const { plugin, getGenerateOptions } = templatePluginFixture({ source: 'manual' });
	await plugin.ensureGranolaTemplateForDocument({ id: 'doc-1', title: 'Testing' }, authContext());
	assert.equal(getGenerateOptions().auto, false);
});

test('verification failure deletes the new panel', async () => {
	const { plugin, calls } = templatePluginFixture({ outcome: 'verification-failure' });
	await plugin.ensureGranolaTemplateForDocument({ id: 'doc-1', title: 'Testing' }, authContext());
	assert.deepEqual(calls.slice(-2), ['waitForPanel', 'deletePanel']);
});

test('refresh failure deletes the new panel', async () => {
	const { plugin, calls } = templatePluginFixture({ outcome: 'refresh-failure' });
	await plugin.ensureGranolaTemplateForDocument({ id: 'doc-1', title: 'Testing' }, authContext());
	assert.deepEqual(calls.slice(-2), ['refreshDocument', 'deletePanel']);
});

test('ambiguous create cleanup deletes one newly appearing empty panel', async () => {
	const recoveredPanel = { id: 'panel-recovered', template_slug: 'template-1', content: '' };
	const { plugin, calls } = templatePluginFixture({
		outcome: 'create-throws',
		panelSnapshots: [[], [recoveredPanel]],
	});
	await plugin.ensureGranolaTemplateForDocument({ id: 'doc-1', title: 'Testing' }, authContext());
	assert.deepEqual(calls, ['getPanels', 'getContext', 'createPanel', 'getPanels', 'deletePanel']);
});

test('ambiguous create leaves multiple new panels untouched', async () => {
	const { plugin, calls } = templatePluginFixture({
		createResult: {},
		panelSnapshots: [[], [
			{ id: 'panel-a', template_slug: 'template-1', content: '' },
			{ id: 'panel-b', template_slug: 'template-1', content: '' },
		]],
	});
	await plugin.ensureGranolaTemplateForDocument({ id: 'doc-1', title: 'Testing' }, authContext());
	assert.equal(calls.includes('deletePanel'), false);
});

test('ambiguous create leaves a newly appearing populated panel untouched', async () => {
	const populatedPanel = {
		id: 'panel-concurrent',
		template_slug: 'template-1',
		content: { type: 'doc', content: [{ type: 'heading' }] },
	};
	const { plugin, calls } = templatePluginFixture({
		outcome: 'create-throws',
		panelSnapshots: [[], [populatedPanel]],
	});
	await plugin.ensureGranolaTemplateForDocument({ id: 'doc-1', title: 'Testing' }, authContext());
	assert.deepEqual(calls, ['getPanels', 'getContext', 'createPanel', 'getPanels']);
	assert.equal(calls.includes('deletePanel'), false);
});

test('ambiguous create does not delete an empty panel beside a concurrent populated panel', async () => {
	const { plugin, calls } = templatePluginFixture({
		outcome: 'create-throws',
		panelSnapshots: [[], [
			{ id: 'panel-empty', template_slug: 'template-1', content: '' },
			{
				id: 'panel-concurrent',
				template_slug: 'template-1',
				content: { type: 'doc', content: [{ type: 'heading' }] },
			},
		]],
	});
	await plugin.ensureGranolaTemplateForDocument({ id: 'doc-1', title: 'Testing' }, authContext());
	assert.equal(calls.includes('deletePanel'), false);
});

test('a recent empty matching panel is treated as in progress without creating a duplicate', async () => {
	const emptyPanel = {
		id: 'panel-recent',
		template_slug: 'template-1',
		content: '',
		updated_at: new Date().toISOString(),
	};
	const { plugin, calls, logs } = templatePluginFixture({ panelSnapshots: [[emptyPanel]] });
	await plugin.ensureGranolaTemplateForDocument({ id: 'doc-1', title: 'Testing' }, authContext());
	assert.deepEqual(calls, ['getPanels']);
	assert.equal(plugin.templateManagementStats.deferred, 1);
	assert.equal(plugin.templateManagementStats.skipped, 0);
	assert.match(logs.join('\n'), /stage=stale-panel status=in-progress/);
});

test('a null timestamp empty matching panel is deferred without creating a duplicate', async () => {
	const emptyPanel = {
		id: 'panel-null',
		template_slug: 'template-1',
		content: '',
		content_updated_at: null,
		updated_at: null,
		created_at: null,
	};
	const { plugin, calls } = templatePluginFixture({ panelSnapshots: [[emptyPanel]] });
	await plugin.ensureGranolaTemplateForDocument({ id: 'doc-1', title: 'Testing' }, authContext());
	assert.deepEqual(calls, ['getPanels']);
	assert.equal(plugin.templateManagementStats.deferred, 1);
	assert.equal(plugin.templateManagementStats.skipped, 0);
});

test('a future-dated empty matching panel is deferred without creating a duplicate', async () => {
	const emptyPanel = {
		id: 'panel-future',
		template_slug: 'template-1',
		content: '',
		updated_at: new Date(Date.now() + (60 * 1000)).toISOString(),
	};
	const { plugin, calls } = templatePluginFixture({ panelSnapshots: [[emptyPanel]] });
	await plugin.ensureGranolaTemplateForDocument({ id: 'doc-1', title: 'Testing' }, authContext());
	assert.deepEqual(calls, ['getPanels']);
	assert.equal(plugin.templateManagementStats.deferred, 1);
	assert.equal(plugin.templateManagementStats.skipped, 0);
});

test('a timestamp-less empty matching panel is deferred instead of ready', async () => {
	const emptyPanel = { id: 'panel-unknown', template_slug: 'template-1', content: '' };
	const { plugin, calls } = templatePluginFixture({ panelSnapshots: [[emptyPanel]] });
	await plugin.ensureGranolaTemplateForDocument({ id: 'doc-1', title: 'Testing' }, authContext());
	assert.deepEqual(calls, ['getPanels']);
	assert.equal(plugin.templateManagementStats.deferred, 1);
	assert.equal(plugin.templateManagementStats.skipped, 0);
});

test('template summaries distinguish deferred panels from ready panels', () => {
	const { getTemplateManagementStatsSummary } = extractPluginMethod('getTemplateManagementStatsSummary');
	const { formatSyncDiagnosticsSummary } = extractPluginMethod('formatSyncDiagnosticsSummary');
	const deferredStats = { attempted: 0, applied: 0, failed: 0, skipped: 0, deferred: 2 };

	assert.equal(
		getTemplateManagementStatsSummary.call({ templateManagementStats: deferredStats }),
		'2 template-deferred'
	);
	assert.equal(
		getTemplateManagementStatsSummary.call({
			templateManagementStats: { attempted: 0, applied: 0, failed: 0, skipped: 1, deferred: 0 },
		}),
		'1 template-ready'
	);
	assert.match(
		formatSyncDiagnosticsSummary.call(
			{ truncateSyncLabel: (value) => value },
			{
				runId: 'run-1',
				source: 'manual',
				docsProcessed: 0,
				docsFetched: 0,
				syncedCount: 0,
				docsSkippedNotReady: 0,
				transcriptFetches: 0,
				myNotesHydrations: 0,
				indexGranolaNotes: 0,
				indexTranscriptNotes: 0,
				templateStats: deferredStats,
			}
		),
		/templates 0\/0\/0, deferred 2/
	);
});

test('a stale empty matching panel is removed before a retry and absent from the returned document', async () => {
	const emptyPanel = {
		id: 'panel-empty',
		template_slug: 'template-1',
		content: '',
		created_at: new Date(Date.now() - (11 * 60 * 1000)).toISOString(),
	};
	const { plugin, calls, persistedPanel } = templatePluginFixture({ panelSnapshots: [[emptyPanel]] });
	const result = await plugin.ensureGranolaTemplateForDocument({ id: 'doc-1', title: 'Testing' }, authContext());
	assert.deepEqual(calls, ['getPanels', 'deletePanel', 'getContext', 'createPanel', 'generate', 'waitForPanel', 'refreshDocument']);
	assert.equal(result.privatePanels.length, 1);
	assert.equal(result.privatePanels[0].id, persistedPanel.id);
	assert.equal(result.privatePanels.some((panel) => panel.id === emptyPanel.id), false);
});

test('stale cleanup failure is separately reported and aborts the generation attempt', async () => {
	const emptyPanel = {
		id: 'panel-stale',
		template_slug: 'template-1',
		content: '',
		updated_at: new Date(Date.now() - (11 * 60 * 1000)).toISOString(),
	};
	const { plugin, calls, logs } = templatePluginFixture({
		outcome: 'stale-cleanup-failure',
		panelSnapshots: [[emptyPanel]],
	});
	await plugin.ensureGranolaTemplateForDocument({ id: 'doc-1', title: 'Testing' }, authContext());
	assert.deepEqual(calls, ['getPanels', 'deletePanel']);
	assert.equal(plugin.templateManagementStats.failed, 1);
	assert.match(logs.join('\n'), /stage=cleanup status=failed/);
	assert.match(logs.join('\n'), /stage=stale-panel status=failed/);
});

test('template orchestration logs never include meeting or template titles', async () => {
	const { plugin, logs } = templatePluginFixture({ outcome: 'generation-and-cleanup-failure' });
	await plugin.ensureGranolaTemplateForDocument({ id: 'doc-1', title: 'PRIVATE-MEETING-TITLE' }, authContext());
	const capturedLogs = logs.join('\n');
	assert.equal(capturedLogs.includes('PRIVATE-MEETING-TITLE'), false);
	assert.equal(capturedLogs.includes('PRIVATE-TEMPLATE-TITLE'), false);
});

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
	assert.equal(result.streamedContentLength, '<h3>Metadata</h3>'.length);
});

test('parseGenerateSummaryStream marks malformed chunks as unparsed without retaining content', () => {
	const { Client } = loadPrivateClient([]);
	const client = new Client(authContext());
	const malformedChunk = 'PRIVATE-MALFORMED-CHUNK';

	const result = client.parseGenerateSummaryStream([
		JSON.stringify({ panel_id: 'panel-1' }),
		malformedChunk,
		JSON.stringify({ choices: [{ delta: { content: 'safe content' } }] }),
	].join('-----CHUNK_BOUNDARY-----'));

	assert.deepEqual(result.eventTypes, ['panel_id', 'unparsed', 'content_delta']);
	assert.equal(result.streamedContentLength, 'safe content'.length);
	assert.equal(JSON.stringify(result).includes(malformedChunk), false);
});

test('getDocumentPanels only requests Ydoc state when explicitly requested', async () => {
	const client = clientWithPanelResponses([[], []]);

	await client.getDocumentPanels('doc-1');
	await client.getDocumentPanels('doc-1', { includeYdocState: true });

	assert.deepEqual(JSON.parse(client.requests[0].body), { document_id: 'doc-1' });
	assert.deepEqual(JSON.parse(client.requests[1].body), {
		document_id: 'doc-1',
		include_ydoc_state: true,
	});
});

test('waitForGeneratedPanel returns only a structured persisted matching panel', async () => {
	const client = clientWithPanelResponses([
		[{ id: 'panel-1', template_slug: 'template-1', content_updated_at: null, content: '' }],
		[{ id: 'panel-1', template_slug: 'template-1', content_updated_at: '2026-07-17T20:07:05.566Z', ydoc_state: 'state', content: { type: 'doc', content: [{ type: 'heading' }] } }],
	]);
	const panel = await client.waitForGeneratedPanel('doc-1', 'panel-1', 'template-1', { attempts: 2, delayMs: 0 });
	assert.equal(panel.id, 'panel-1');
	assert.equal(client.requests.length, 2);
	assert.equal(JSON.parse(client.requests[0].body).include_ydoc_state, true);
});

test('waitForGeneratedPanel continues polling past malformed persisted content entries', async () => {
	const client = clientWithPanelResponses([
		[{ id: 'panel-1', template_slug: 'template-1', content_updated_at: '2026-07-17T20:07:05.566Z', ydoc_state: 'state', content: { type: 'doc', content: [null, {}] } }],
		[{ id: 'panel-1', template_slug: 'template-1', content_updated_at: '2026-07-17T20:07:06.566Z', ydoc_state: 'state-2', content: { type: 'doc', content: [{ type: 'heading' }] } }],
	]);

	const panel = await client.waitForGeneratedPanel('doc-1', 'panel-1', 'template-1', { attempts: 2, delayMs: 0 });

	assert.deepEqual(panel.content.content, [{ type: 'heading' }]);
	assert.equal(client.requests.length, 2);
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
	assert.ok(body.updated_at);
	assert.equal(Object.hasOwn(body, 'content'), false);
});
