const path = require('node:path');
const fs = require('node:fs');
const os = require('node:os');

function safeJsonParse(value, fallback = null) {
	if (value === null || value === undefined) {
		return fallback;
	}

	if (typeof value !== 'string') {
		return value;
	}

	try {
		return JSON.parse(value);
	} catch (error) {
		return fallback;
	}
}

function uniquePaths(paths) {
	return [...new Set(paths)];
}

class GranolaAuthResolver {
	constructor({
		fs: fileSystem = fs,
		os: osModule = os,
		platform = process.platform,
		clientVersion,
		authKeyPath = 'Library/Application Support/Granola/supabase.json',
	} = {}) {
		this.fs = fileSystem;
		this.os = osModule;
		this.platform = platform;
		this.clientVersion = clientVersion;
		this.authKeyPath = authKeyPath;
	}

	getCandidatePaths() {
		const homedir = this.os.homedir();
		const username = this.os.userInfo().username;

		return {
			storedAccountsPath: path.resolve(homedir, 'Library/Application Support/Granola/stored-accounts.json'),
			supabasePaths: uniquePaths([
				path.resolve(homedir, 'Users', username, 'Library/Application Support/Granola/supabase.json'),
				path.resolve(homedir, this.authKeyPath),
				path.resolve(homedir, 'Library/Application Support/Granola/supabase.json'),
			]),
		};
	}

	normalizeWorkOsTokens(tokens, sourcePath, sourceKind) {
		if (!tokens || !tokens.access_token) {
			return null;
		}

		const accessToken = tokens.access_token;
		return {
			accessToken,
			token: accessToken,
			refreshToken: tokens.refresh_token || null,
			sessionId: tokens.session_id || null,
			signInMethod: tokens.sign_in_method || null,
			obtainedAt: tokens.obtained_at || null,
			expiresInSeconds: tokens.expires_in || null,
			clientVersion: this.clientVersion,
			platform: this.platform,
			osVersion: this.os.release(),
			source: sourcePath,
			sourceKind,
			workspaceId: null,
			deviceId: null,
		};
	}

	loadCandidateSessions() {
		const { storedAccountsPath, supabasePaths } = this.getCandidatePaths();

		if (this.fs.existsSync(storedAccountsPath)) {
			const storedAccounts = JSON.parse(this.fs.readFileSync(storedAccountsPath, 'utf8'));
			const accounts = safeJsonParse(storedAccounts.accounts, []);
			const sessions = [];

			for (const account of Array.isArray(accounts) ? accounts : []) {
				const session = this.normalizeWorkOsTokens(
					safeJsonParse(account.tokens, {}),
					storedAccountsPath,
					'stored-accounts'
				);
				if (session) {
					sessions.push(session);
				}
			}

			if (sessions.length > 0) {
				return sessions;
			}
		}

		for (const authPath of supabasePaths) {
			if (!this.fs.existsSync(authPath)) {
				continue;
			}

			const raw = JSON.parse(this.fs.readFileSync(authPath, 'utf8'));
			const workosTokens = safeJsonParse(raw.workos_tokens, raw.workos_tokens);
			const session = this.normalizeWorkOsTokens(workosTokens, authPath, 'supabase');
			if (session) {
				return [session];
			}
		}

		return [];
	}

	async resolveSession() {
		return this.loadCandidateSessions()[0] || null;
	}
}

module.exports = {
	GranolaAuthResolver,
	safeJsonParse,
};
