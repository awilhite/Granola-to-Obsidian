class GranolaApiClient {
	constructor({ requestUrl }) {
		this.requestUrl = requestUrl;
	}

	buildHeaders(authContext) {
		const headers = {
			Authorization: `Bearer ${authContext.accessToken}`,
			'Content-Type': 'application/json',
			Accept: 'application/json',
			'User-Agent': `Granola/${authContext.clientVersion}`,
			'X-Client-Version': authContext.clientVersion,
			'X-Granola-Platform': authContext.platform,
			'X-Granola-Os-Version': authContext.osVersion,
		};

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
}

module.exports = {
	GranolaApiClient,
};
