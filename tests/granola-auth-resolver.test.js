const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { GranolaAuthResolver } = require('../lib/granola-auth-resolver');

const fixturesDir = path.join(__dirname, 'fixtures', 'granola-auth');

function loadFixture(name) {
	return JSON.parse(fs.readFileSync(path.join(fixturesDir, name), 'utf8'));
}

test('parses WorkOS stored-accounts payloads into normalized sessions', async () => {
	const storedAccountsPath = '/Users/tester/Library/Application Support/Granola/stored-accounts.json';
	const resolver = new GranolaAuthResolver({
		fs: {
			existsSync: (file) => file === storedAccountsPath,
			readFileSync: () => JSON.stringify(loadFixture('stored-accounts.json')),
		},
		os: {
			homedir: () => '/Users/tester',
			release: () => '25.2.0',
			userInfo: () => ({ username: 'tester' }),
		},
		platform: 'darwin',
		clientVersion: '7.255.6',
	});

	const candidates = resolver.loadCandidateSessions();
	assert.equal(candidates.length, 1);
	assert.equal(candidates[0].sourceKind, 'stored-accounts');
	assert.equal(candidates[0].refreshToken, 'refresh-token-123');
	assert.equal(candidates[0].sessionId, 'session-123');
	assert.equal(candidates[0].signInMethod, 'GoogleOAuth');
});

test('continues to fallback auth sources when stored-accounts.json is malformed', async () => {
	const storedAccountsPath = '/Users/tester/Library/Application Support/Granola/stored-accounts.json';
	const supabasePath = '/Users/tester/Library/Application Support/Granola/supabase.json';
	const files = {
		[storedAccountsPath]: '{"accounts":"not valid json"',
		[supabasePath]: JSON.stringify(loadFixture('supabase.json')),
	};

	const resolver = new GranolaAuthResolver({
		fs: {
			existsSync: (file) => Object.hasOwn(files, file),
			readFileSync: (file) => files[file],
		},
		os: {
			homedir: () => '/Users/tester',
			release: () => '25.2.0',
			userInfo: () => ({ username: 'tester' }),
		},
		platform: 'darwin',
		clientVersion: '7.255.6',
	});

	const candidates = resolver.loadCandidateSessions();
	assert.equal(candidates.length, 1);
	assert.equal(candidates[0].sourceKind, 'supabase');
	assert.equal(candidates[0].accessToken, 'supabase-access-token');
});

test('parses WorkOS supabase payloads into normalized sessions', async () => {
	const supabasePath = '/Users/tester/Library/Application Support/Granola/supabase.json';
	const resolver = new GranolaAuthResolver({
		fs: {
			existsSync: (file) => file === supabasePath,
			readFileSync: () => JSON.stringify(loadFixture('supabase.json')),
		},
		os: {
			homedir: () => '/Users/tester',
			release: () => '25.2.0',
			userInfo: () => ({ username: 'tester' }),
		},
		platform: 'darwin',
		clientVersion: '7.255.6',
	});

	const candidates = resolver.loadCandidateSessions();
	assert.equal(candidates.length, 1);
	assert.equal(candidates[0].sourceKind, 'supabase');
	assert.equal(candidates[0].accessToken, 'supabase-access-token');
	assert.equal(candidates[0].refreshToken, 'supabase-refresh-token');
});

test('prefers stored-accounts over older fallback sources when both are present', async () => {
	const files = {
		'/Users/tester/Library/Application Support/Granola/stored-accounts.json': JSON.stringify(loadFixture('stored-accounts.json')),
		'/Users/tester/Library/Application Support/Granola/supabase.json': JSON.stringify(loadFixture('supabase.json')),
	};

	const resolver = new GranolaAuthResolver({
		fs: {
			existsSync: (file) => Object.hasOwn(files, file),
			readFileSync: (file) => files[file],
		},
		os: {
			homedir: () => '/Users/tester',
			release: () => '25.2.0',
			userInfo: () => ({ username: 'tester' }),
		},
		platform: 'darwin',
		clientVersion: '7.255.6',
	});

	const candidates = resolver.loadCandidateSessions();
	assert.equal(candidates.length, 2);
	assert.equal(candidates[0].sourceKind, 'stored-accounts');
	assert.equal(candidates[1].sourceKind, 'supabase');

	const resolved = await resolver.resolveSession();
	assert.equal(resolved.sourceKind, 'stored-accounts');
	assert.equal(resolved.accessToken, 'stored-access-token');
});
