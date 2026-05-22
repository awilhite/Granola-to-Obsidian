const test = require('node:test');
const assert = require('node:assert/strict');
const { GranolaApiClient } = require('../lib/granola-api-client');

test('fetchDocuments uses the auth resolver output and the current document endpoint', async () => {
	let capturedRequest = null;
	const client = new GranolaApiClient({
		requestUrl: async (request) => {
			capturedRequest = request;
			return {
				json: {
					docs: [{ id: 'doc-1', title: 'Test Meeting' }],
				},
			};
		},
	});

	const docs = await client.fetchDocuments({
		accessToken: 'access-token',
		clientVersion: '7.255.6',
		platform: 'darwin',
		osVersion: '25.2.0',
		workspaceId: null,
		deviceId: null,
	}, {
		limit: 10,
		offset: 0,
	});

	assert.equal(capturedRequest.url, 'https://api.granola.ai/v2/get-documents');
	assert.equal(capturedRequest.headers.Authorization, 'Bearer access-token');
	assert.equal(capturedRequest.headers['X-Client-Version'], '7.255.6');
	assert.equal(docs.length, 1);
	assert.equal(docs[0].id, 'doc-1');
});

test('refreshWorkOsSession uses the WorkOS refresh endpoint and current client headers', async () => {
	let capturedRequest = null;
	const client = new GranolaApiClient({
		requestUrl: async (request) => {
			capturedRequest = request;
			return {
				json: {
					access_token: 'refreshed-access-token',
					refresh_token: 'refresh-token-next',
				},
			};
		},
		getTimeZone: () => 'America/Indiana/Indianapolis',
	});

	const refreshed = await client.refreshWorkOsSession({
		accessToken: 'stale-access-token',
		refreshToken: 'refresh-token-123',
		clientVersion: '7.255.6',
		platform: 'darwin',
		osVersion: '25.2.0',
		workspaceId: 'workspace-1',
		deviceId: 'device-1',
	});

	assert.equal(capturedRequest.url, 'https://api.granola.ai/v1/refresh-access-token');
	assert.equal(capturedRequest.method, 'POST');
	assert.equal(capturedRequest.headers.Authorization, 'Bearer stale-access-token');
	assert.equal(capturedRequest.headers['X-Granola-Time-Zone'], 'America/Indiana/Indianapolis');
	assert.equal(capturedRequest.headers['X-Granola-Workspace-Id'], 'workspace-1');
	assert.equal(capturedRequest.headers['X-Granola-Device-Id'], 'device-1');
	assert.deepEqual(JSON.parse(capturedRequest.body), {
		refresh_token: 'refresh-token-123',
	});
	assert.equal(refreshed.access_token, 'refreshed-access-token');
	assert.equal(refreshed.refresh_token, 'refresh-token-next');
});
