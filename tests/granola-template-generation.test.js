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
