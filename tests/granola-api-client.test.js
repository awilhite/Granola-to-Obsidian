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
