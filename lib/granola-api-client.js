class GranolaApiClient {
	constructor({ requestUrl, getTimeZone = () => Intl.DateTimeFormat().resolvedOptions().timeZone || null }) {
		this.requestUrl = requestUrl;
		this.getTimeZone = getTimeZone;
	}

	buildHeaders(authContext) {
		const headers = {
			'Content-Type': 'application/json',
			Accept: 'application/json',
			'User-Agent': `Granola/${authContext.clientVersion}`,
			'X-Client-Version': authContext.clientVersion,
			'X-Granola-Platform': authContext.platform,
			'X-Granola-Os-Version': authContext.osVersion,
		};

		if (authContext.accessToken) {
			headers.Authorization = `Bearer ${authContext.accessToken}`;
		}
		if (authContext.workspaceId) {
			headers['X-Granola-Workspace-Id'] = authContext.workspaceId;
		}
		if (authContext.deviceId) {
			headers['X-Granola-Device-Id'] = authContext.deviceId;
		}

		return headers;
	}

	async fetchDocuments(authContext, { limit, offset }) {
		const response = await this.requestUrl({
			url: 'https://api.granola.ai/v2/get-documents',
			method: 'POST',
			headers: this.buildHeaders(authContext),
			body: JSON.stringify({
				limit,
				offset,
				include_last_viewed_panel: true,
				include_panels: true,
			}),
		});

		return response?.json?.docs || [];
	}

	async refreshWorkOsSession(authContext) {
		const timeZone = this.getTimeZone();
		const headers = this.buildHeaders(authContext);
		if (timeZone) {
			headers['X-Granola-Time-Zone'] = timeZone;
		}

		const response = await this.requestUrl({
			url: 'https://api.granola.ai/v1/refresh-access-token',
			method: 'POST',
			headers,
			body: JSON.stringify({
				refresh_token: authContext.refreshToken,
			}),
		});

		return response?.json || null;
	}
}

module.exports = {
	GranolaApiClient,
};
