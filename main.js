const obsidian = require('obsidian');
const path = require('path');
const fs = require('fs');
const os = require('os');

function getDefaultAuthPath() {
	if (obsidian.Platform.isWin) {
		return 'AppData/Roaming/Granola/supabase.json';
	} else if (obsidian.Platform.isLinux) {
		return '.config/Granola/supabase.json';
	} else {
		// Default to macOS path
		return 'Library/Application Support/Granola/supabase.json';
	}
}

const DEFAULT_SETTINGS = {
	syncDirectory: 'Granola',
	notePrefix: '',
	authKeyPath: getDefaultAuthPath(),
	filenameTemplate: '{title}',
	dateFormat: 'YYYY-MM-DD',
	autoSyncFrequency: 300000,
	enableDailyNoteIntegration: false,
	dailyNoteSectionName: '## Granola Meetings',
	enablePeriodicNoteIntegration: false,
	periodicNoteSectionName: '## Granola Meetings',
	existingNoteBehavior: 'changed', // 'never', 'changed', 'always'
	includeAttendeeTags: false,
	excludeMyNameFromTags: true,
	myName: 'Danny McClelland',
	includeFolderTags: false,
	includeGranolaUrl: false,
	attendeeTagTemplate: 'person/{name}',
	existingNoteSearchScope: 'syncDirectory', // 'syncDirectory', 'entireVault', 'specificFolders'
	specificSearchFolders: [], // Array of folder paths to search in when existingNoteSearchScope is 'specificFolders'
	enableDateBasedFolders: false,
	dateFolderFormat: 'YYYY-MM-DD',
	enableGranolaFolders: false, // Enable folder-based organization
	folderTagTemplate: 'folder/{name}', // Template for folder tags
	filenameSeparator: '_', // Character to separate words in filenames ('_', '-', or '')
	existingFileAction: 'timestamp', // 'timestamp' - create timestamped version, 'skip' - ignore existing file
	syncAllHistoricalNotes: false, // Sync all historical notes from Granola, not just recent ones
	documentSyncLimit: 100, // Maximum number of documents to sync (used when syncAllHistoricalNotes is false)
	includeFullTranscript: false, // Include full meeting transcript in notes
	storeTranscriptInSeparateNote: false, // Store transcripts in separate notes instead of embedding inline
	transcriptDirectory: 'Granola Transcripts', // Folder to store separate transcript notes in
	includeMyNotes: true, // Include "My Notes" section from Granola
	includeEnhancedNotes: true, // Include "Enhanced Notes" (AI summary) from Granola
	selectedGranolaFolders: [], // Array of Granola folder IDs to sync (empty = sync all)
	enableFolderFilter: false, // Enable filtering by Granola folders
	enableAutoReorganize: false, // Automatically reorganize notes after sync
	// Frontmatter customization
	includeTitle: true, // Include title field in frontmatter
	includeDates: true, // Include created_at and updated_at in frontmatter
	frontmatterDateFormat: 'iso', // 'iso', 'date-only', 'custom'
	customDateFormat: 'YYYY-MM-DD', // Custom date format string
	additionalFrontmatter: '', // Additional frontmatter lines (key: value, one per line)
	mapMetadataToFrontmatter: false, // Map Granola metadata block fields into frontmatter
	removeMetadataSectionFromBody: false, // Remove the inline metadata section after mapping it
	metadataOrgTemplate: '{name}', // Template for mapped org frontmatter values
	metadataPersonTemplate: '{name}', // Template for mapped people frontmatter values
	includeReviewTask: false, // Add a review task near the top of synced notes
	enableGranolaTemplateManagement: false, // Automatically ensure a selected Granola template exists before sync
	granolaTemplateId: '', // Selected Granola template ID for template management
	granolaTemplateTitle: '', // Selected Granola template title for display
};

const REVIEW_TASK_TEXT = '- [ ] Review imported Granola note';
const GRANOLA_TEMPLATE_CLIENT_VERSION = '7.71.1';
const POST_MEETING_SYNC_DELAY_MS = 2 * 60 * 1000;

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

class GranolaPrivateClient {
	constructor(authContext) {
		this.authContext = authContext;
	}

	buildHeaders({ accept = 'application/json', contentType = 'application/json' } = {}) {
		const headers = {
			'Authorization': 'Bearer ' + this.authContext.token,
			'Accept': accept,
			'User-Agent': 'Granola/' + (this.authContext.clientVersion || GRANOLA_TEMPLATE_CLIENT_VERSION),
			'X-Client-Version': this.authContext.clientVersion || GRANOLA_TEMPLATE_CLIENT_VERSION,
			'X-Granola-Platform': this.authContext.platform || process.platform,
			'X-Granola-Os-Version': this.authContext.osVersion || os.release(),
		};

		if (contentType) {
			headers['Content-Type'] = contentType;
		}
		if (this.authContext.workspaceId) {
			headers['X-Granola-Workspace-Id'] = this.authContext.workspaceId;
		}
		if (this.authContext.deviceId) {
			headers['X-Granola-Device-Id'] = this.authContext.deviceId;
		}
		return headers;
	}

	async postJson(url, body, options = {}) {
		const response = await obsidian.requestUrl({
			url,
			method: 'POST',
			headers: this.buildHeaders(options),
			body: JSON.stringify(body || {}),
		});
		return response.json;
	}

	async postText(url, body, options = {}) {
		const response = await obsidian.requestUrl({
			url,
			method: 'POST',
			headers: this.buildHeaders(options),
			body: JSON.stringify(body || {}),
		});
		return response.text || '';
	}

	async getPanelTemplates() {
		return await this.postJson('https://api.granola.ai/v1/get-panel-templates', {});
	}

	async getDocumentPanels(documentId) {
		return await this.postJson('https://api.granola.ai/v1/get-document-panels', {
			document_id: documentId,
		});
	}

	async getDocumentBatch(documentId) {
		const response = await this.postJson('https://api.granola.ai/v1/get-documents-batch', {
			document_ids: [documentId],
		});
		return response && Array.isArray(response.docs) ? response.docs[0] : null;
	}

	async getDocumentMetadata(documentId) {
		return await this.postJson('https://api.granola.ai/v1/get-document-metadata', {
			document_id: documentId,
		});
	}

	async getDocumentTranscript(documentId) {
		return await this.postJson('https://api.granola.ai/v1/get-document-transcript', {
			document_id: documentId,
		});
	}

	async createDocumentPanel(documentId, templateId, title = 'Summary') {
		return await this.postJson('https://api.granola.ai/v1/create-document-panel', {
			document_id: documentId,
			title,
			content: '',
			template_slug: templateId,
			last_viewed_at: new Date().toISOString(),
		});
	}

	async updateDocumentPanel(panelId, content) {
		const now = new Date().toISOString();
		return await this.postJson('https://api.granola.ai/v1/update-document-panel', {
			id: panelId,
			content,
			original_content: content,
			last_viewed_at: now,
			content_updated_at: now,
		});
	}

	formatTranscript(entries) {
		if (!Array.isArray(entries) || entries.length === 0) {
			return '';
		}

		let output = '';
		let lastSpeaker = null;
		for (const entry of entries) {
			const speaker = entry && entry.source === 'microphone' ? 'Me' : 'Them';
			if (speaker !== lastSpeaker) {
				if (output) {
					output += ' ';
				}
				output += `${speaker}: `;
				lastSpeaker = speaker;
			}
			if (entry && entry.text) {
				output += `${entry.text} `;
			}
		}

		return output.trim();
	}

	stripNotesWrapper(text) {
		return String(text || '')
			.replace(/^\s*<notes>\s*/i, '')
			.replace(/\s*<\/notes>\s*$/i, '')
			.trim();
	}

	collectStreamContent(streamText) {
		let content = '';
		for (const chunk of String(streamText || '').split('-----CHUNK_BOUNDARY-----')) {
			const trimmed = chunk.trim();
			if (!trimmed) {
				continue;
			}

			let parsed = null;
			try {
				parsed = JSON.parse(trimmed);
			} catch (error) {
				continue;
			}

			const delta = parsed && parsed.choices && parsed.choices[0] && parsed.choices[0].delta
				? parsed.choices[0].delta.content
				: null;
			if (typeof delta === 'string') {
				content += delta;
			}
		}
		return content;
	}

	buildPromptVariables(doc, metadata, transcriptEntries, template) {
		const transcript = this.formatTranscript(transcriptEntries);
		const createdAt = doc && doc.created_at ? new Date(doc.created_at) : new Date();
		const notes = doc ? (doc.notes_markdown || doc.notes_plain || '') : '';
		const transcriptLength = transcript.length;

		let maxNumHeadings = 4;
		if (transcriptLength >= 80000) {
			maxNumHeadings = 8;
		} else if (transcriptLength >= 20000) {
			maxNumHeadings = Math.floor(transcriptLength / 10000) + 1;
		} else if (transcriptLength >= 8000) {
			maxNumHeadings = 5;
		}

		const creator = metadata && metadata.creator ? metadata.creator : {};
		const creatorName = creator && creator.name ? creator.name : '';
		const creatorEmail = creator && creator.email ? creator.email : '';
		const creatorCompany = creator && creator.details && creator.details.company
			? (creator.details.company.name || '')
			: '';
		const dateOnly = new Intl.DateTimeFormat('en-US', {
			year: 'numeric',
			month: 'long',
			day: 'numeric',
			timeZone: 'UTC',
		}).format(createdAt);
		const timeOnly = new Intl.DateTimeFormat('en-US', {
			hour: 'numeric',
			minute: '2-digit',
			timeZone: 'UTC',
		}).format(createdAt);
		const todaysDate = new Intl.DateTimeFormat('en-US', {
			year: 'numeric',
			month: 'long',
			day: 'numeric',
		}).format(new Date());

		return {
			transcript,
			notes,
			viewer_name: '',
			viewer_bio: '',
			viewer_company: '',
			my_name: creatorName,
			my_bio: '',
			their_name: '',
			their_bio: '',
			participants: creatorEmail ? `${creatorName} <${creatorEmail}>` : creatorName,
			calendar_event_title: doc && doc.title ? doc.title : '',
			headers: '',
			document_id: doc && doc.id ? doc.id : '',
			urls: '',
			my_company: creatorCompany,
			my_colleagues: '',
			external_attendees: '',
			external_companies: '',
			time: timeOnly,
			date: dateOnly,
			todays_date: todaysDate,
			is_multi_language: false,
			is_british_english: false,
			english_only_summary: true,
			user_dictionary: '',
			is_short_transcript: transcriptLength < 4000,
			has_long_user_notes: notes.length > 1000,
			is_ios: doc && doc.creation_source === 'iOS',
			is_collaborative: false,
			is_user_type_vc: false,
			latest_meeting_summary: 'No summary',
			tags: '',
			summary_headings: Array.isArray(template.sections)
				? template.sections.map((section) => `#${section.heading}: ${section.section_description}`).join('\n')
				: '',
			first_heading: Array.isArray(template.sections) && template.sections.length > 0
				? `#${template.sections[0].heading}`
				: '',
			meeting_description: template.description || '',
			max_num_headings: maxNumHeadings,
			template_title: template.title || '',
		};
	}

	async generateTemplateMarkdown(doc, metadata, transcriptEntries, template) {
		const streamText = await this.postText(
			'https://stream.api.granola.ai/v1/llm-proxy-stream',
			{
				prompt_slug: 'template-summary-consolidated',
				prompt_variables: this.buildPromptVariables(doc, metadata, transcriptEntries, template),
				chat_history: [],
			},
			{ accept: '*/*' }
		);

		const rawContent = this.collectStreamContent(streamText);
		return this.stripNotesWrapper(rawContent);
	}
}

class GranolaSyncPlugin extends obsidian.Plugin {
	async onload() {
		this.autoSyncInterval = null;
		this.settings = DEFAULT_SETTINGS;
		this.statusBarItem = null;
		this.ribbonIconEl = null;
		
		try {
			const data = await this.loadData();
			if (data) {
				this.settings = this.normalizeSettings(data);
				if (JSON.stringify(data) !== JSON.stringify(this.settings)) {
					await this.saveData(this.settings);
				}
			} else {
				this.settings = this.normalizeSettings();
			}
		} catch (error) {
			// Could not load settings, using defaults
		}

		this.statusBarItem = this.addStatusBarItem();
		this.updateStatusBar('Idle');

		// Add ribbon icon for syncing
		this.ribbonIconEl = this.addRibbonIcon('sync', 'Sync Granola notes', () => {
			this.syncNotes();
		});

		// Add ribbon icon for finding duplicates
		this.ribbonDuplicatesEl = this.addRibbonIcon('search', 'Find duplicate Granola notes', () => {
			this.findDuplicatesAndOpen();
		});

		this.addCommand({
			id: 'sync-granola-notes',
			name: 'Sync Granola Notes',
			callback: () => {
				this.syncNotes();
			}
		});

		this.addCommand({
			id: 'find-duplicate-granola-notes',
			name: 'Find Duplicate Granola Notes',
			callback: () => {
				this.findDuplicatesAndOpen();
			}
		});

		this.addCommand({
			id: 'reorganize-granola-notes',
			name: 'Reorganize Granola Notes into Folders',
			callback: () => {
				this.reorganizeExistingNotes();
			}
		});

		this.addSettingTab(new GranolaSyncSettingTab(this.app, this));

		window.setTimeout(() => {
			this.setupAutoSync();
		}, 1000);
	}

	onunload() {
		this.clearAutoSync();
	}

	normalizeSettings(data = {}) {
		const normalized = Object.assign({}, DEFAULT_SETTINGS, data);
		if (!['never', 'changed', 'always'].includes(normalized.existingNoteBehavior)) {
			if (typeof data.skipExistingNotes === 'boolean') {
				normalized.existingNoteBehavior = data.skipExistingNotes ? 'changed' : 'always';
			} else {
				normalized.existingNoteBehavior = DEFAULT_SETTINGS.existingNoteBehavior;
			}
		}

		delete normalized.skipExistingNotes;
		return normalized;
	}

	async saveSettings() {
		try {
			await this.saveData(this.settings);
			this.setupAutoSync();
		} catch (error) {
			console.error('Failed to save settings:', error);
		}
	}

	async saveSettingsWithoutSync() {
		try {
			await this.saveData(this.settings);
		} catch (error) {
			console.error('Failed to save settings:', error);
		}
	}

	updateStatusBar(status, count) {
		if (!this.statusBarItem) return;

		let text = 'Granola Sync: ';

		if (status === 'Idle') {
			text += 'Idle';
		} else if (status === 'Syncing') {
			// If count is a string, use it as a custom message
			if (typeof count === 'string') {
				text += count;
			} else {
				text += 'Syncing...';
			}
		} else if (status === 'Complete') {
			if (typeof count === 'string') {
				text += count;
			} else {
				text += count + ' notes synced';
			}
			window.setTimeout(() => {
				this.updateStatusBar('Idle');
			}, 3000);
		} else if (status === 'Error') {
			text += 'Error - ' + (count || 'sync failed');
			window.setTimeout(() => {
				this.updateStatusBar('Idle');
			}, 5000);
		}

		this.statusBarItem.setText(text);
	}

	setupAutoSync() {
		this.clearAutoSync();
		
		if (this.settings.autoSyncFrequency > 0) {
			this.autoSyncInterval = window.setInterval(() => {
				this.syncNotes();
			}, this.settings.autoSyncFrequency);
		}
	}

	clearAutoSync() {
		if (this.autoSyncInterval) {
			window.clearInterval(this.autoSyncInterval);
			this.autoSyncInterval = null;
		}
	}

	getFrequencyLabel(frequency) {
		const minutes = frequency / (1000 * 60);
		const hours = frequency / (1000 * 60 * 60);
		
		if (frequency === 0) return 'Disabled';
		if (frequency < 60000) return (frequency / 1000) + ' seconds';
		if (minutes < 60) return minutes + ' minutes';
		return hours + ' hours';
	}

	getGranolaPlatform() {
		if (obsidian.Platform.isWin) return 'win32';
		if (obsidian.Platform.isLinux) return 'linux';
		return 'darwin';
	}

	getPublicGranolaHeaders(token) {
		return {
			'Authorization': 'Bearer ' + token,
			'Content-Type': 'application/json',
			'Accept': '*/*',
			'User-Agent': 'Granola/' + GRANOLA_TEMPLATE_CLIENT_VERSION,
			'X-Client-Version': GRANOLA_TEMPLATE_CLIENT_VERSION,
		};
	}

	getTemplateManagementStatsSummary() {
		const stats = this.templateManagementStats;
		if (!stats || (!stats.attempted && !stats.skipped)) {
			return '';
		}

		const parts = [];
		if (stats.applied) {
			parts.push(`${stats.applied} template-updated`);
		}
		if (stats.failed) {
			parts.push(`${stats.failed} template-failed`);
		}
		if (stats.skipped) {
			parts.push(`${stats.skipped} template-ready`);
		}

		return parts.join(', ');
	}

	getExistingNoteBehavior() {
		return this.settings.existingNoteBehavior || DEFAULT_SETTINGS.existingNoteBehavior;
	}

	// Helper function to get a readable speaker label
	getSpeakerLabel(source) {
		switch (source) {
			case "microphone":
				return "Me";
			case "system":
			default:
				return "Them";
		}
	}	

	// Helper function to format timestamp for display
	formatTimestamp(timestamp) {
		const d = new Date(timestamp);
		return [d.getHours(), d.getMinutes(), d.getSeconds()]
			.map(v => String(v).padStart(2, '0'))
			.join(':');
	}

	transcriptToMarkdown(segments) {
		if (!segments || segments.length === 0) {
			return "*No transcript content available*";
		}

		const sortedSegments = segments.slice().sort((a, b) => {
			const timeA = new Date(a.start_timestamp || 0);
			const timeB = new Date(b.start_timestamp || 0);
			return timeA - timeB;
		});

		const lines = [];
		let currentSpeaker = null;
		let currentText = "";
		let currentTimestamp = null;

		const flushCurrentSegment = () => {
			const cleanText = currentText.trim().replace(/\s+/g, " ");
			if (cleanText && currentSpeaker) {
				const timeStr = this.formatTimestamp(currentTimestamp);
				const speakerLabel = this.getSpeakerLabel(currentSpeaker);
				lines.push(`**${speakerLabel}** *(${timeStr})*: ${cleanText}`)
			}
			currentText = "";
			currentSpeaker = null;
			currentTimestamp = null;
		};

		for (const segment of sortedSegments) {
			if (currentSpeaker && currentSpeaker !== segment.source) {
				flushCurrentSegment();
			}
			if (!currentSpeaker) {
				currentSpeaker = segment.source;
				currentTimestamp = segment.start_timestamp;
			}
			const segmentText = segment.text;
			if (segmentText && segmentText.trim()) {
				currentText += currentText ? ` ${segmentText}` : segmentText;
			}
		}
		flushCurrentSegment();

		return lines.length === 0 ? "*No transcript content available*" : lines.join("\n\n");

	}

	shouldFetchTranscript() {
		return this.settings.includeFullTranscript || this.settings.storeTranscriptInSeparateNote;
	}

	normalizeVaultPath(filePath) {
		return obsidian.normalizePath(filePath).replace(/\\/g, '/');
	}

	getTranscriptDirectoryPath(doc) {
		const baseDirectory = this.normalizeVaultPath(this.settings.transcriptDirectory);
		if (!this.settings.enableDateBasedFolders || !doc.created_at) {
			return baseDirectory;
		}

		const dateFolder = this.formatDate(doc.created_at, this.settings.dateFolderFormat);
		return this.normalizeVaultPath(path.join(baseDirectory, dateFolder));
	}

	generateTranscriptLink(filePath) {
		const normalizedPath = this.normalizeVaultPath(filePath).replace(/\.md$/i, '');
		return '## Transcript\n\n[[' + normalizedPath + '|Transcript]]';
	}

	buildTranscriptFrontmatter(doc, title) {
		let frontmatter = '---\n';
		frontmatter += 'granola_id: ' + (doc.id || 'unknown_id') + '\n';
		frontmatter += 'granola_transcript: true\n';
		frontmatter += 'source: "Granola"\n';
		const syncUpdatedAt = this.getDocumentSyncUpdatedAt(doc);
		const noteDate = this.formatObsidianDateProperty(doc.created_at);

		if (this.settings.includeTitle) {
			const escapedTitle = title.replace(/"/g, '\\"');
			frontmatter += 'title: "' + escapedTitle + ' Transcript"\n';
		}

		if (noteDate) {
			frontmatter += 'date: ' + noteDate + '\n';
		}

		if (this.settings.includeDates) {
			if (doc.created_at) {
				frontmatter += 'created_at: ' + this.formatFrontmatterDate(doc.created_at) + '\n';
			}
			if (syncUpdatedAt) {
				frontmatter += 'updated_at: ' + this.formatFrontmatterDate(syncUpdatedAt) + '\n';
			}
		}

		if (syncUpdatedAt) {
			frontmatter += 'granola_updated_at: ' + syncUpdatedAt + '\n';
		}

		frontmatter += '---\n\n';
		return frontmatter;
	}

	buildTranscriptNoteContent(doc, transcript) {
		const noteTitle = this.generateNoteTitle(doc);
		return this.buildTranscriptFrontmatter(doc, noteTitle)
			+ '# ' + noteTitle + ' Transcript\n\n'
			+ transcript.trim()
			+ '\n';
	}

	async findExistingTranscriptNoteByGranolaId(docId) {
		const transcriptDirectory = this.normalizeVaultPath(this.settings.transcriptDirectory);
		const files = this.app.vault.getMarkdownFiles().filter(file =>
			file.path === transcriptDirectory || file.path.startsWith(transcriptDirectory + '/')
		);

		for (const file of files) {
			try {
				const content = await this.app.vault.read(file);
				const frontmatterMatch = content.match(/^---\n([\s\S]*?)\n---/);
				if (!frontmatterMatch) {
					continue;
				}

				const frontmatter = frontmatterMatch[1];
				const granolaIdMatch = frontmatter.match(/granola_id:\s*(.+)$/m);
				const transcriptMatch = frontmatter.match(/granola_transcript:\s*true$/m);

				if (granolaIdMatch && transcriptMatch && granolaIdMatch[1].trim() === docId) {
					return file;
				}
			} catch (error) {
				console.error('Error reading transcript note for duplicate detection:', file.path, error);
			}
		}

		return null;
	}

	async writeSeparateTranscriptNote(doc, transcript) {
		if (!this.settings.storeTranscriptInSeparateNote || !transcript || transcript === 'no_transcript') {
			return null;
		}

		const transcriptContent = this.buildTranscriptNoteContent(doc, transcript);
		const existingTranscript = await this.findExistingTranscriptNoteByGranolaId(doc.id || 'unknown_id');
		if (existingTranscript) {
			await this.app.vault.modify(existingTranscript, transcriptContent);
			return existingTranscript.path;
		}

		const targetDirectory = this.getTranscriptDirectoryPath(doc);
		await this.ensureDateBasedDirectoryExists(targetDirectory);

		const filename = this.generateFilename(doc) + '.md';
		const filepath = this.normalizeVaultPath(path.join(targetDirectory, filename));
		const existingFileByName = this.app.vault.getAbstractFileByPath(filepath);
		if (existingFileByName) {
			const timestamp = this.formatDate(doc.created_at, 'HH-mm');
			const baseFilename = this.generateFilename(doc);
			const uniqueFilename = baseFilename + ' Transcript ' + timestamp + '.md';
			const finalFilepath = this.normalizeVaultPath(path.join(targetDirectory, uniqueFilename));
			await this.app.vault.create(finalFilepath, transcriptContent);
			return finalFilepath;
		}

		await this.app.vault.create(filepath, transcriptContent);
		return filepath;
	}

	async syncNotes() {
		try {
			this.updateStatusBar('Syncing');
			this.templateManagementStats = { attempted: 0, applied: 0, failed: 0, skipped: 0 };
			
			await this.ensureDirectoryExists();

			const authContext = await this.loadCredentials();
			if (!authContext) {
				this.updateStatusBar('Error', 'credentials failed');
				return;
			}

			const documents = await this.fetchGranolaDocuments(authContext);
			if (!documents) {
				this.updateStatusBar('Error', 'fetch failed');
				return;
			}

			// Fetch folders if folder support or folder filtering is enabled
			let folders = null;
			if (this.settings.enableGranolaFolders || this.settings.enableFolderFilter) {
				folders = await this.fetchGranolaFolders(authContext);
				if (folders) {
					// Create a mapping of document ID to folder for quick lookup
					this.documentToFolderMap = {};
					// Also store all available folders for the settings UI
					this.availableGranolaFolders = folders;
					for (const folder of folders) {
						if (folder.document_ids) {
							for (const docId of folder.document_ids) {
								this.documentToFolderMap[docId] = folder;
							}
						}
					}
				}
			}

			// Filter documents by selected folders if folder filtering is enabled
			let documentsToSync = documents;
			if (this.settings.enableFolderFilter && this.settings.selectedGranolaFolders.length > 0 && this.documentToFolderMap) {
				documentsToSync = documents.filter(doc => {
					const folder = this.documentToFolderMap[doc.id];
					if (!folder) {
						// Document is not in any folder - check if user wants to include "unfiled" docs
						return false;
					}
					// Check if the folder ID is in the selected folders list
					return this.settings.selectedGranolaFolders.includes(folder.id);
				});
				console.log(`Folder filter: syncing ${documentsToSync.length} of ${documents.length} documents`);
			}

			let syncedCount = 0;
			const todaysNotes = [];
			const today = new Date().toDateString();

			for (let i = 0; i < documentsToSync.length; i++) {
				const doc = documentsToSync[i];
				try {
					const readiness = this.getDocumentSyncReadiness(doc);
					if (!readiness.ready) {
						console.log('Skipping document "' + (doc.title || doc.id) + '" - ' + readiness.reason);
						continue;
					}

					// Fetch transcript if enabled
					if (this.shouldFetchTranscript()) {
						const transcriptData = await this.fetchTranscript(authContext, doc.id);
						doc.transcript = this.transcriptToMarkdown(transcriptData);
					}

					const success = await this.processDocument(doc, authContext);
					if (success) {
						syncedCount++;
					}
					
					// Check for note integration regardless of sync success
					// This ensures existing notes from today are still included
					if ((this.settings.enableDailyNoteIntegration || this.settings.enablePeriodicNoteIntegration) && doc.created_at) {
						const noteDate = new Date(doc.created_at).toDateString();
						if (noteDate === today) {
							// Find the actual file that was created or already exists
							const actualFile = await this.findExistingNoteByGranolaId(doc.id);
							
							if (actualFile) {
								const noteData = {};
								noteData.title = doc.title || 'Untitled Granola Note';
								noteData.actualFilePath = actualFile.path; // Use actual file path
								
								const createdDate = new Date(doc.created_at);
								const hours = String(createdDate.getHours()).padStart(2, '0');
								const minutes = String(createdDate.getMinutes()).padStart(2, '0');
								noteData.time = hours + ':' + minutes;
								
								todaysNotes.push(noteData);
							}
						}
					}
				} catch (error) {
					console.error('Error processing document ' + doc.title + ':', error);
				}
			}

			// Create a deep copy to prevent any reference issues
			const todaysNotesCopy = todaysNotes.map(note => ({...note}));
			
			if (this.settings.enableDailyNoteIntegration && todaysNotes.length > 0) {
				await this.updateDailyNote(todaysNotesCopy);
			}

			if (this.settings.enablePeriodicNoteIntegration && todaysNotes.length > 0) {
				await this.updatePeriodicNote(todaysNotesCopy);
			}

			const templateSummary = this.getTemplateManagementStatsSummary();
			if (templateSummary) {
				console.log('Granola Template Management summary:', this.templateManagementStats);
				this.updateStatusBar('Complete', `${syncedCount} notes synced, ${templateSummary}`);
			} else {
				this.updateStatusBar('Complete', syncedCount);
			}

			// Auto-reorganize if enabled
			if (this.settings.enableAutoReorganize &&
				(this.settings.enableGranolaFolders || this.settings.enableDateBasedFolders)) {
				try {
					await this.reorganizeExistingNotes(true); // quiet mode
				} catch (error) {
					console.error('Auto-reorganization failed:', error);
				}
			}

		} catch (error) {
			console.error('Granola sync failed:', error);
			this.updateStatusBar('Error', 'sync failed');
		}
	}

	async loadCredentials() {
		const homedir = os.homedir();
		const storedAccountsPath = path.resolve(homedir, 'Library/Application Support/Granola/stored-accounts.json');
		const authPaths = [
			// New location (with Users in path)
			path.resolve(homedir, 'Users', os.userInfo().username, 'Library/Application Support/Granola/supabase.json'),
			// Current configured path
			path.resolve(homedir, this.settings.authKeyPath),
			// Fallback to old default location
			path.resolve(homedir, 'Library/Application Support/Granola/supabase.json')
		];

		try {
			if (fs.existsSync(storedAccountsPath)) {
				const storedAccountsFile = fs.readFileSync(storedAccountsPath, 'utf8');
				const storedAccounts = JSON.parse(storedAccountsFile);
				const accounts = safeJsonParse(storedAccounts.accounts, []);
				const primaryAccount = Array.isArray(accounts) && accounts.length > 0 ? accounts[0] : null;
				const tokenData = primaryAccount ? safeJsonParse(primaryAccount.tokens, {}) : null;
				if (tokenData && tokenData.access_token) {
					console.log('Successfully loaded Granola credentials from:', storedAccountsPath);
					return {
						token: tokenData.access_token,
						sessionId: tokenData.session_id || null,
						clientVersion: GRANOLA_TEMPLATE_CLIENT_VERSION,
						platform: this.getGranolaPlatform(),
						osVersion: os.release(),
						source: storedAccountsPath,
					};
				}
			}
		} catch (error) {
			console.error('Error reading credentials from', storedAccountsPath, ':', error);
		}

		for (const authPath of authPaths) {
			try {
				if (!fs.existsSync(authPath)) {
					continue;
				}

				const credentialsFile = fs.readFileSync(authPath, 'utf8');
				const data = JSON.parse(credentialsFile);
				
				let accessToken = null;
				let sessionId = data.session_id || null;
				
				// Try new token structure (workos_tokens)
				if (data.workos_tokens) {
					try {
						const workosTokens = JSON.parse(data.workos_tokens);
						accessToken = workosTokens.access_token;
						sessionId = sessionId || workosTokens.session_id || null;
					} catch (e) {
						// workos_tokens might already be an object
						accessToken = data.workos_tokens.access_token;
						sessionId = sessionId || data.workos_tokens.session_id || null;
					}
				}
				
				// Fallback to old token structure (cognito_tokens)
				if (!accessToken && data.cognito_tokens) {
					try {
						const cognitoTokens = JSON.parse(data.cognito_tokens);
						accessToken = cognitoTokens.access_token;
					} catch (e) {
						// cognito_tokens might already be an object
						accessToken = data.cognito_tokens.access_token;
					}
				}
				
				if (accessToken) {
					console.log('Successfully loaded credentials from:', authPath);
					return {
						token: accessToken,
						sessionId: sessionId || null,
						clientVersion: GRANOLA_TEMPLATE_CLIENT_VERSION,
						platform: this.getGranolaPlatform(),
						osVersion: os.release(),
						source: authPath,
					};
				}
			} catch (error) {
				console.error('Error reading credentials from', authPath, ':', error);
				continue;
			}
		}

		console.error('No valid credentials found in any of the expected locations');
		return null;
	}

	async fetchGranolaDocuments(authContext) {
		try {
			const token = authContext.token;
			const allDocs = [];
			let offset = 0;
			const batchSize = 100;
			let hasMore = true;

			// Determine the maximum number of documents to fetch
			const maxDocuments = this.settings.syncAllHistoricalNotes
				? Number.MAX_SAFE_INTEGER
				: this.settings.documentSyncLimit;

			while (hasMore && allDocs.length < maxDocuments) {
				const response = await obsidian.requestUrl({
					url: 'https://api.granola.ai/v2/get-documents',
					method: 'POST',
					headers: this.getPublicGranolaHeaders(token),
					body: JSON.stringify({
						limit: batchSize,
						offset: offset,
						include_last_viewed_panel: true,
						include_panels: true
					})
				});

				const apiResponse = response.json;

				if (!apiResponse || !apiResponse.docs) {
					console.error('API response format is unexpected');
					return allDocs.length > 0 ? allDocs : null;
				}

				const docs = apiResponse.docs;
				allDocs.push(...docs);

				// Check if there are more documents to fetch
				if (docs.length < batchSize) {
					// Received fewer docs than requested, so we've reached the end
					hasMore = false;
				} else if (!this.settings.syncAllHistoricalNotes && allDocs.length >= maxDocuments) {
					// Reached the user-specified limit
					hasMore = false;
				} else {
					// More documents may be available, increment offset
					offset += batchSize;
				}

				// Show progress for large syncs
				if (this.settings.syncAllHistoricalNotes && allDocs.length > 100) {
					this.updateStatusBar('Syncing', `${allDocs.length} docs fetched`);
				}
			}

			// Trim to max documents if we went over
			if (allDocs.length > maxDocuments) {
				allDocs.length = maxDocuments;
			}

			console.log(`Fetched ${allDocs.length} documents from Granola`);
			return allDocs;
		} catch (error) {
			console.error('Error fetching documents:', error);
			return null;
		}
	}

	async fetchGranolaFolders(authContext) {
		try {
			const response = await obsidian.requestUrl({
				url: 'https://api.granola.ai/v1/get-document-lists-metadata',
				method: 'POST',
				headers: this.getPublicGranolaHeaders(authContext.token),
				body: JSON.stringify({
					include_document_ids: true,
					include_only_joined_lists: false
				})
			});

			const apiResponse = response.json;
			
			if (!apiResponse || !apiResponse.lists) {
				console.error('Folders API response format is unexpected');
				return null;
			}

			// Convert the lists object to an array of folders
			const folders = Object.values(apiResponse.lists);
			return folders;
		} catch (error) {
			console.error('Error fetching folders:', error);
			return null;
		}
	}

	async fetchTranscript(authContext, docId) {
		try {
			const response = await obsidian.requestUrl({
				url: `https://api.granola.ai/v1/get-document-transcript`,
				method: 'POST',
				headers: this.getPublicGranolaHeaders(authContext.token),
				body: JSON.stringify({
					'document_id': docId
				})
			});

			return response.json;

		} catch (error) {
			console.error('Error fetching transcript for document ' + docId + ':' + error);
			return null;
		}
	}

	getGranolaPrivateClient(authContext) {
		return new GranolaPrivateClient(authContext);
	}

	async fetchGranolaTemplates(authContext, forceRefresh = false) {
		if (!forceRefresh && Array.isArray(this.availableGranolaTemplates) && this.availableGranolaTemplates.length > 0) {
			return this.availableGranolaTemplates;
		}

		const client = this.getGranolaPrivateClient(authContext);
		const templates = await client.getPanelTemplates();
		this.availableGranolaTemplates = Array.isArray(templates) ? templates : [];
		return this.availableGranolaTemplates;
	}

	convertProseMirrorToMarkdown(content) {
		if (!content || typeof content !== 'object' || !content.content) {
			return '';
		}

		const processNode = (node, indentLevel = 0) => {
			if (!node || typeof node !== 'object') {
				return '';
			}

			const nodeType = node.type || '';
			const nodeContent = node.content || [];
			const text = node.text || '';

			if (nodeType === 'heading') {
				const level = node.attrs && node.attrs.level ? node.attrs.level : 1;
				const headingText = nodeContent.map(child => processNode(child, indentLevel)).join('');
				return '#'.repeat(level) + ' ' + headingText + '\n\n';
			} else if (nodeType === 'paragraph') {
				const paraText = nodeContent.map(child => processNode(child, indentLevel)).join('');
				return paraText + '\n\n';
			} else if (nodeType === 'bulletList') {
				const items = [];
				for (let i = 0; i < nodeContent.length; i++) {
					const item = nodeContent[i];
					if (item.type === 'listItem') {
						const processedItem = this.processListItem(item, indentLevel);
						if (processedItem) {
							items.push(processedItem);
						}
					}
				}
				return items.join('\n') + '\n\n';
			} else if (nodeType === 'text') {
				return text;
			} else {
				return nodeContent.map(child => processNode(child, indentLevel)).join('');
			}
		};

		return processNode(content);
	}

	processListItem(listItem, indentLevel = 0) {
		if (!listItem || !listItem.content) {
			return '';
		}

		const indent = '  '.repeat(indentLevel); // 2 spaces per indent level
		let itemText = '';
		let hasNestedLists = false;

		for (const child of listItem.content) {
			if (child.type === 'paragraph') {
				// Process paragraph content for the main bullet text
				const paraText = (child.content || []).map(node => {
					if (node.type === 'text') {
						return node.text || '';
					}
					return '';
				}).join('').trim();
				if (paraText) {
					itemText += paraText;
				}
			} else if (child.type === 'bulletList') {
				// Handle nested bullet lists
				hasNestedLists = true;
				const nestedItems = [];
				for (const nestedItem of child.content || []) {
					if (nestedItem.type === 'listItem') {
						const nestedProcessed = this.processListItem(nestedItem, indentLevel + 1);
						if (nestedProcessed) {
							nestedItems.push(nestedProcessed);
						}
					}
				}
				if (nestedItems.length > 0) {
					itemText += '\n' + nestedItems.join('\n');
				}
			}
		}

		if (!itemText.trim()) {
			return '';
		}

		// Format the main bullet point
		const mainBullet = indent + '- ' + itemText.split('\n')[0];
		
		// If there are nested items, append them
		if (hasNestedLists) {
			const lines = itemText.split('\n');
			if (lines.length > 1) {
				const nestedLines = lines.slice(1).join('\n');
				return mainBullet + '\n' + nestedLines;
			}
		}

		return mainBullet;
	}

	formatDate(date, format) {
		if (!date) return '';
		
		const d = new Date(date);
		const year = d.getFullYear();
		const month = String(d.getMonth() + 1).padStart(2, '0');
		const day = String(d.getDate()).padStart(2, '0');
		const hours = String(d.getHours()).padStart(2, '0');
		const minutes = String(d.getMinutes()).padStart(2, '0');
		const seconds = String(d.getSeconds()).padStart(2, '0');
		
		return format
			.replace(/YYYY/g, year)
			.replace(/YY/g, String(year).slice(-2))
			.replace(/MM/g, month)
			.replace(/DD/g, day)
			.replace(/HH/g, hours)
			.replace(/mm/g, minutes)
			.replace(/ss/g, seconds);
	}

	generateNoteTitle(doc) {
		const title = doc.title || 'Untitled Granola Note';
		// Clean the title for use as a heading - remove invalid characters but keep spaces
		return title.replace(/[<>:"/\\|?*]/g, '').trim();
	}

	escapeYamlString(value) {
		return `"${String(value).replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"`;
	}

	slugifyTemplateValue(value) {
		return String(value)
			.replace(/[^\w\s-]/g, '')
			.trim()
			.replace(/\s+/g, '-')
			.toLowerCase();
	}

	applyMetadataValueTemplate(template, value) {
		if (value === null || value === undefined || value === '') {
			return '';
		}

		const stringValue = String(value).trim();
		if (!stringValue) {
			return '';
		}

		const resolvedTemplate = template && template.trim() ? template : '{name}';
		return resolvedTemplate
			.replace(/{name}/g, stringValue)
			.replace(/{value}/g, stringValue)
			.replace(/{slug}/g, this.slugifyTemplateValue(stringValue));
	}

	formatFrontmatterList(key, values) {
		if (!Array.isArray(values) || values.length === 0) {
			return `${key}: []\n`;
		}

		const lines = [`${key}:\n`];
		for (const value of values) {
			if (value === null || value === undefined || value === '') {
				continue;
			}
			lines.push(`  - ${this.escapeYamlString(value)}\n`);
		}

		return lines.length === 1 ? `${key}: []\n` : lines.join('');
	}

	extractMetadataSection(markdown) {
		if (!markdown || typeof markdown !== 'string') {
			return { metadata: null, content: markdown || '' };
		}

		const match = markdown.match(/^(?:#{1,6}\s*Metadata\s*\n+)?(?:```json\s*\n)?\s*({[\s\S]*?})\s*(?:\n```)?\s*(?=\n---\s*(?:\n|$)|\n#{1,6}\s|\nChat with meeting transcript:|$)/i);
		if (!match) {
			return { metadata: null, content: markdown };
		}

		const rawJson = match[1];
		try {
			const metadata = JSON.parse(rawJson);
			const content = (markdown.slice(0, match.index) + markdown.slice(match.index + match[0].length))
				.replace(/^\s+/, '')
				.replace(/\n{3,}/g, '\n\n');
			return { metadata, content };
		} catch (error) {
			console.error('Error parsing metadata section:', error);
			return { metadata: null, content: markdown };
		}
	}

	generateFilename(doc) {
		const title = doc.title || 'Untitled Granola Note';
		const docId = doc.id || 'unknown_id';

		let createdDate = '';
		let updatedDate = '';
		let createdTime = '';
		let updatedTime = '';
		let createdDateTime = '';
		let updatedDateTime = '';

		if (doc.created_at) {
			createdDate = this.formatDate(doc.created_at, this.settings.dateFormat);
			createdTime = this.formatDate(doc.created_at, 'HH-mm-ss');
			createdDateTime = this.formatDate(doc.created_at, this.settings.dateFormat + '_HH-mm-ss');
		}

		if (doc.updated_at) {
			updatedDate = this.formatDate(doc.updated_at, this.settings.dateFormat);
			updatedTime = this.formatDate(doc.updated_at, 'HH-mm-ss');
			updatedDateTime = this.formatDate(doc.updated_at, this.settings.dateFormat + '_HH-mm-ss');
		}

		let filename = this.settings.filenameTemplate
			.replace(/{title}/g, title)
			.replace(/{id}/g, docId)
			.replace(/{created_date}/g, createdDate)
			.replace(/{updated_date}/g, updatedDate)
			.replace(/{created_time}/g, createdTime)
			.replace(/{updated_time}/g, updatedTime)
			.replace(/{created_datetime}/g, createdDateTime)
			.replace(/{updated_datetime}/g, updatedDateTime);

		if (this.settings.notePrefix) {
			filename = this.settings.notePrefix + filename;
		}

		const invalidChars = /[<>:"/\\|?*]/g;
		filename = filename.replace(invalidChars, '');
		filename = filename.replace(/\s+/g, this.settings.filenameSeparator);

		return filename;
	}

	generateDateBasedPath(doc) {
		if (!this.settings.enableDateBasedFolders || !doc.created_at) {
			return this.settings.syncDirectory;
		}

		const dateFolder = this.formatDate(doc.created_at, this.settings.dateFolderFormat);
		return path.join(this.settings.syncDirectory, dateFolder);
	}

	generateFolderBasedPath(doc) {
		if (!this.settings.enableGranolaFolders || !this.documentToFolderMap) {
			return this.settings.syncDirectory;
		}

		const folder = this.documentToFolderMap[doc.id];
		if (!folder || !folder.title) {
			return this.settings.syncDirectory;
		}

		// Clean folder name for filesystem use
		const cleanFolderName = folder.title
			.replace(/[<>:"/\\|?*]/g, '') // Remove invalid filesystem characters
			.replace(/\s+/g, this.settings.filenameSeparator) // Replace spaces with configured separator
			.trim();

		return path.join(this.settings.syncDirectory, cleanFolderName);
	}

	async ensureDateBasedDirectoryExists(datePath) {
		try {
			const folder = this.app.vault.getFolderByPath(datePath);
			if (!folder) {
				await this.app.vault.createFolder(datePath);
			}
		} catch (error) {
			console.error('Error creating date-based directory:', datePath, error);
		}
	}

	async ensureFolderBasedDirectoryExists(folderPath) {
		try {
			const folder = this.app.vault.getFolderByPath(folderPath);
			if (!folder) {
				await this.app.vault.createFolder(folderPath);
			}
		} catch (error) {
			console.error('Error creating folder-based directory:', folderPath, error);
		}
	}

	/**
	 * Finds an existing note by its Granola ID based on the configured search scope.
	 * 
	 * Search scope options:
	 * - 'syncDirectory' (default): Only searches within the configured sync directory
	 * - 'entireVault': Searches all markdown files in the vault
	 * - 'specificFolders': Searches within user-specified folders (including subfolders)
	 * 
	 * This allows users to move their Granola notes to different folders while still
	 * avoiding duplicates when existing notes are matched by Granola ID.
	 * 
	 * @param {string} granolaId - The Granola ID to search for
	 * @param {object} options - Search options
	 * @param {boolean} options.includeTranscriptNotes - Whether transcript notes should be considered matches
	 * @returns {TFile|null} The found file or null if not found
	 */
		async findExistingNoteByGranolaId(granolaId, options = {}) {
		const includeTranscriptNotes = options.includeTranscriptNotes === true;
		let filesToSearch = [];

		if (this.settings.existingNoteSearchScope === 'entireVault') {
			// Search all markdown files in the vault
			filesToSearch = this.app.vault.getMarkdownFiles();
		} else if (this.settings.existingNoteSearchScope === 'specificFolders') {
			// Search in specific folders
			if (this.settings.specificSearchFolders.length === 0) {
			return null;
		}

			for (const folderPath of this.settings.specificSearchFolders) {
				const folder = this.app.vault.getFolderByPath(folderPath);
				if (folder) {
					const folderFiles = this.getAllMarkdownFilesInFolder(folder);
					filesToSearch = filesToSearch.concat(folderFiles);
				}
			}
		} else {
			// Default: search only in sync directory (including subfolders)
			const folder = this.app.vault.getFolderByPath(this.settings.syncDirectory);
			if (!folder) {
				return null;
			}
			// Use recursive search to find notes in subfolders (important for date-based or Granola folder organization)
			filesToSearch = this.getAllMarkdownFilesInFolder(folder);
		}
		
		for (const file of filesToSearch) {
			try {
				const content = await this.app.vault.read(file);
				const frontmatterMatch = content.match(/^---\n([\s\S]*?)\n---/);
				
				if (frontmatterMatch) {
					const frontmatter = frontmatterMatch[1];
					const granolaIdMatch = frontmatter.match(/granola_id:\s*(.+)$/m);
					const transcriptMatch = frontmatter.match(/granola_transcript:\s*true$/m);
					
					if (
						granolaIdMatch &&
						granolaIdMatch[1].trim() === granolaId &&
						(includeTranscriptNotes || !transcriptMatch)
					) {
						return file;
					}
				}
			} catch (error) {
				console.error('Error reading file for Granola ID check:', file.path, error);
			}
		}
		
		return null;
	}

	getAllMarkdownFilesInFolder(folder) {
		if (!folder) {
			return [];
		}
		const folderPath = folder.path;
		return this.app.vault.getMarkdownFiles().filter(
			(file) => file.path.startsWith(folderPath + '/')
		);
	}

	/**
	 * Gets all folder paths in the vault (useful for future UI improvements)
	 * @returns {string[]} Array of folder paths
	 */
	getAllFolderPaths() {
		const allFolders = this.app.vault.getAllFolders();
		return allFolders.map(folder => folder.path).sort();
	}

	async findDuplicateNotes(suppressNotice = false) {
		try {
			// Get all markdown files in the vault
			const allFiles = this.app.vault.getMarkdownFiles();
			const granolaFiles = {};
			const duplicates = [];

			// Check each file for granola-id
			for (const file of allFiles) {
				try {
					const content = await this.app.vault.read(file);
					const frontmatterMatch = content.match(/^---\n([\s\S]*?)\n---/);

					if (frontmatterMatch) {
						const frontmatter = frontmatterMatch[1];
						const granolaIdMatch = frontmatter.match(/granola_id:\s*(.+)$/m);

						if (granolaIdMatch) {
							const granolaId = granolaIdMatch[1].trim();

							if (granolaFiles[granolaId]) {
								// Found a duplicate
								if (!duplicates.some(d => d.granolaId === granolaId)) {
									duplicates.push({
										granolaId: granolaId,
										files: [granolaFiles[granolaId], file]
									});
								} else {
									// Add to existing duplicate group
									const duplicate = duplicates.find(d => d.granolaId === granolaId);
									duplicate.files.push(file);
								}
							} else {
								granolaFiles[granolaId] = file;
							}
						}
					}
				} catch (error) {
					console.error('Error reading file:', file.path, error);
				}
			}

			// Only show notices if not suppressed
			if (!suppressNotice) {
				if (duplicates.length === 0) {
					new obsidian.Notice('No duplicate Granola notes found! 🎉');
				} else {
					// Create a summary message
					let message = `Found ${duplicates.length} set(s) of duplicate Granola notes:\n\n`;

					for (const duplicate of duplicates) {
						message += `Granola ID: ${duplicate.granolaId}\n`;
						for (const file of duplicate.files) {
							message += `  • ${file.path}\n`;
						}
						message += '\n';
					}

					message += 'Check the console for full details. You can manually delete the duplicates you don\'t want to keep.';

					new obsidian.Notice(message, 10000); // Show for 10 seconds
				}
			}

			return duplicates;

		} catch (error) {
			console.error('Error finding duplicate notes:', error);
			if (!suppressNotice) {
				new obsidian.Notice('Error finding duplicate notes. Check console for details.');
			}
			return [];
		}
	}

	generateDuplicatesReport(duplicates) {
		let report = '---\n';
		report += 'title: "Granola Duplicates Report"\n';
		report += 'date: ' + new Date().toISOString() + '\n';
		report += '---\n\n';
		report += '# Duplicate Granola Notes\n\n';
		report += `Found ${duplicates.length} set(s) of duplicate notes.\n\n`;

		for (let i = 0; i < duplicates.length; i++) {
			const duplicate = duplicates[i];
			report += `## Duplicate Set ${i + 1}: ${duplicate.granolaId}\n\n`;
			report += '| Filename | Link |\n';
			report += '|----------|------|\n';

			for (const file of duplicate.files) {
				const fileName = file.path;
				const baseName = fileName.split('/').pop();
				report += `| ${baseName} | [[${fileName}]] |\n`;
			}

			report += '\n';
		}

		report += '## Instructions\n\n';
		report += '1. **Review**: Click the links above to open and compare each duplicate\n';
		report += '2. **Decide**: Choose which version you want to keep\n';
		report += '3. **Delete**: Use your file explorer to delete the files you don\'t need\n';
		report += '4. **Cleanup**: Delete this report when done\n\n';
		report += '> **Note**: All listed files have the same Granola ID but are stored as separate files in your vault.\n';

		return report;
	}

	async createOrUpdateDuplicatesFile(duplicates) {
		try {
			const report = this.generateDuplicatesReport(duplicates);
			const duplicatesPath = 'Granola/Duplicates Report.md';

			// Ensure Granola folder exists
			let folder = this.app.vault.getFolderByPath('Granola');
			if (!folder) {
				try {
					folder = await this.app.vault.createFolder('Granola');
				} catch (error) {
					// Folder might already exist or error creating, try to get it
					folder = this.app.vault.getFolderByPath('Granola');
				}
			}

			// Check if file exists
			let reportFile = this.app.vault.getAbstractFileByPath(duplicatesPath);

			if (reportFile && reportFile instanceof obsidian.TFile) {
				// File exists, modify it with new content
				await this.app.vault.modify(reportFile, report);
			} else {
				// File doesn't exist, try to create it
				try {
					reportFile = await this.app.vault.create(duplicatesPath, report);
				} catch (createError) {
					// File might have been created between the check and create, try to modify it
					if (createError.message && createError.message.includes('already exists')) {
						const existingFile = this.app.vault.getAbstractFileByPath(duplicatesPath);
						if (existingFile && existingFile instanceof obsidian.TFile) {
							await this.app.vault.modify(existingFile, report);
							reportFile = existingFile;
						} else {
							throw createError;
						}
					} else {
						throw createError;
					}
				}
			}

			// Open the report file
			if (reportFile instanceof obsidian.TFile) {
				await this.app.workspace.getLeaf().openFile(reportFile);
			}

			new obsidian.Notice('Duplicates report created and opened');

		} catch (error) {
			console.error('Error creating duplicates report:', error);
			new obsidian.Notice('Error creating duplicates report. Check console for details.');
		}
	}

	async findDuplicatesAndOpen() {
		try {
			// Find duplicates without showing popup notice
			const duplicates = await this.findDuplicateNotes(true);

			if (duplicates.length === 0) {
				new obsidian.Notice('No duplicate Granola notes found! 🎉');
				return;
			}

			// Create and open duplicates report file
			await this.createOrUpdateDuplicatesFile(duplicates);

		} catch (error) {
			console.error('Error processing duplicates:', error);
			new obsidian.Notice('Error processing duplicates. Check console for details.');
		}
	}

	/**
	 * Extracts content from document panels by type.
	 * Granola stores different content types in panels: 'my_notes', 'enhanced_notes', etc.
	 * @param {Object} doc - The document object
	 * @param {string} panelType - The type of panel to extract ('my_notes', 'enhanced_notes')
	 * @returns {Object|null} The panel content or null if not found
	 */
	extractPanelContent(doc, panelType) {
		// First check the panels array if available
		if (doc.panels && Array.isArray(doc.panels)) {
			for (const panel of doc.panels) {
				if (panel.type === panelType && panel.content && panel.content.type === 'doc') {
					return panel.content;
				}
			}
		}

		// Fallback: check last_viewed_panel for enhanced notes
		if (panelType === 'enhanced_notes' && doc.last_viewed_panel &&
			doc.last_viewed_panel.content && doc.last_viewed_panel.content.type === 'doc') {
			return doc.last_viewed_panel.content;
		}

		// Fallback: for my_notes, check doc.content directly (user's own notes)
		if (panelType === 'my_notes' && doc.content && doc.content.type === 'doc') {
			return doc.content;
		}

		return null;
	}

	shouldUseGranolaTemplateManagement() {
		return Boolean(this.settings.enableGranolaTemplateManagement && this.settings.granolaTemplateId);
	}

	getReviewTaskLine() {
		return this.settings.includeReviewTask ? REVIEW_TASK_TEXT : '';
	}

	getGranolaTemplatePanel(panels, templateId) {
		if (!templateId || !Array.isArray(panels)) {
			return null;
		}

		const matches = panels.filter((panel) =>
			panel &&
			!panel.deleted_at &&
			panel.template_slug === templateId
		);

		if (matches.length === 0) {
			return null;
		}

		matches.sort((a, b) => {
			const timeA = new Date(a.content_updated_at || a.updated_at || 0).getTime();
			const timeB = new Date(b.content_updated_at || b.updated_at || 0).getTime();
			return timeB - timeA;
		});

		return matches[0];
	}

	getPanelMarkdownContent(panel) {
		if (!panel || panel.deleted_at) {
			return '';
		}

		if (typeof panel.content === 'string') {
			return panel.content.trim();
		}

		if (panel.content && panel.content.type === 'doc') {
			return this.convertProseMirrorToMarkdown(panel.content).trim();
		}

		return '';
	}

	getEnhancedNotesMarkdown(doc) {
		const selectedTemplatePanel = this.getGranolaTemplatePanel(doc.privatePanels, this.settings.granolaTemplateId);
		const templateMarkdown = this.getPanelMarkdownContent(selectedTemplatePanel);
		if (templateMarkdown) {
			return templateMarkdown;
		}

		if (typeof doc.granolaTemplateManagementMarkdown === 'string' && doc.granolaTemplateManagementMarkdown.trim()) {
			return doc.granolaTemplateManagementMarkdown.trim();
		}

		const enhancedNotesContent = this.extractPanelContent(doc, 'enhanced_notes');
		if (!enhancedNotesContent) {
			// Some docs return the visible summary panel as raw markdown in
			// last_viewed_panel.content instead of a ProseMirror doc.
			return this.getPanelMarkdownContent(doc.last_viewed_panel);
		}

		return this.convertProseMirrorToMarkdown(enhancedNotesContent).trim();
	}

	getAllDocumentPanels(doc) {
		const panelsById = new Map();

		const addPanel = (panel) => {
			if (!panel || typeof panel !== 'object') {
				return;
			}

			if (panel.id) {
				panelsById.set(panel.id, panel);
				return;
			}

			const syntheticId = JSON.stringify([
				panel.template_slug || '',
				panel.type || '',
				panel.title || '',
				panel.updated_at || '',
				panel.content_updated_at || ''
			]);
			panelsById.set(syntheticId, panel);
		};

		for (const panel of doc.privatePanels || []) {
			addPanel(panel);
		}

		for (const panel of doc.panels || []) {
			addPanel(panel);
		}

		if (doc.last_viewed_panel) {
			addPanel(doc.last_viewed_panel);
		}

		return Array.from(panelsById.values());
	}

	getPanelSyncUpdatedAt(panel) {
		if (!panel || panel.deleted_at) {
			return '';
		}

		return panel.content_updated_at || panel.updated_at || '';
	}

	isSyncRelevantPanel(panel) {
		return Boolean(
			panel &&
			!panel.deleted_at &&
			(
				panel.template_slug ||
				panel.type === 'enhanced_notes' ||
				panel.title === 'Summary'
			)
		);
	}

	getDocumentSyncUpdatedAt(doc) {
		const candidates = [];
		const addCandidate = (timestamp) => {
			if (!timestamp) {
				return;
			}

			const time = new Date(timestamp).getTime();
			if (Number.isNaN(time)) {
				return;
			}

			candidates.push({ timestamp, time });
		};

		addCandidate(doc.updated_at);

		for (const panel of this.getAllDocumentPanels(doc)) {
			if (this.isSyncRelevantPanel(panel)) {
				addCandidate(this.getPanelSyncUpdatedAt(panel));
			}
		}

		if (candidates.length === 0) {
			return doc.updated_at || doc.created_at || '';
		}

		candidates.sort((a, b) => b.time - a.time);
		return candidates[0].timestamp;
	}

	getDocumentSyncReadiness(doc) {
		if (typeof doc.meeting_end_count === 'number' && doc.meeting_end_count === 0) {
			return { ready: false, reason: 'meeting is still in progress' };
		}

		const lastActivityAt = doc.updated_at || doc.created_at;
		if (!lastActivityAt) {
			return { ready: true, reason: '' };
		}

		const lastActivityTime = new Date(lastActivityAt).getTime();
		if (Number.isNaN(lastActivityTime)) {
			return { ready: true, reason: '' };
		}

		const ageMs = Date.now() - lastActivityTime;
		if (ageMs < POST_MEETING_SYNC_DELAY_MS) {
			const remainingSeconds = Math.ceil((POST_MEETING_SYNC_DELAY_MS - ageMs) / 1000);
			return {
				ready: false,
				reason: `waiting for Granola processing to settle (${remainingSeconds}s remaining)`
			};
		}

		return { ready: true, reason: '' };
	}

	async ensureGranolaTemplateForDocument(doc, authContext) {
		if (!this.shouldUseGranolaTemplateManagement()) {
			return doc;
		}

		if (!this.templateManagementStats) {
			this.templateManagementStats = { attempted: 0, applied: 0, failed: 0, skipped: 0 };
		}

		try {
			const client = this.getGranolaPrivateClient(authContext);
			const existingPanels = await client.getDocumentPanels(doc.id);
			doc.privatePanels = Array.isArray(existingPanels) ? existingPanels : [];

			const existingTemplatePanel = this.getGranolaTemplatePanel(doc.privatePanels, this.settings.granolaTemplateId);
			if (existingTemplatePanel) {
				this.templateManagementStats.skipped++;
				const existingMarkdown = this.getPanelMarkdownContent(existingTemplatePanel);
				if (existingMarkdown) {
					doc.granolaTemplateManagementMarkdown = existingMarkdown;
				}
				return doc;
			}

			this.templateManagementStats.attempted++;

			const templates = await this.fetchGranolaTemplates(authContext);
			const selectedTemplate = templates.find((template) => template.id === this.settings.granolaTemplateId);
			if (!selectedTemplate) {
				throw new Error('Selected Granola template could not be found');
			}

			const [batchDoc, metadata, transcriptEntries] = await Promise.all([
				client.getDocumentBatch(doc.id),
				client.getDocumentMetadata(doc.id),
				client.getDocumentTranscript(doc.id),
			]);

			const generatedMarkdown = await client.generateTemplateMarkdown(batchDoc || doc, metadata || {}, transcriptEntries || [], selectedTemplate);
			if (!generatedMarkdown) {
				throw new Error('Granola template generation returned empty content');
			}

			const createdPanel = await client.createDocumentPanel(doc.id, selectedTemplate.id);
			if (!createdPanel || !createdPanel.id) {
				throw new Error('Granola template panel could not be created');
			}

			await client.updateDocumentPanel(createdPanel.id, generatedMarkdown);

			const [refreshedPanels, refreshedDoc] = await Promise.all([
				client.getDocumentPanels(doc.id),
				client.getDocumentBatch(doc.id),
			]);

			doc.privatePanels = Array.isArray(refreshedPanels) ? refreshedPanels : doc.privatePanels;
			doc.granolaTemplateManagementMarkdown = generatedMarkdown;
			if (refreshedDoc && refreshedDoc.updated_at) {
				doc.updated_at = refreshedDoc.updated_at;
			} else {
				doc.updated_at = new Date().toISOString();
			}

			this.templateManagementStats.applied++;
			console.log(`Granola Template Management applied "${selectedTemplate.title}" to "${doc.title || doc.id}"`);
		} catch (error) {
			this.templateManagementStats.failed++;
			console.error('Granola Template Management failed for "' + (doc.title || doc.id) + '":', error);
		}

		return doc;
	}

	/**
	 * Builds the note content from available sections.
	 * Includes My Notes, Enhanced Notes, and Transcript based on settings and availability.
	 * @param {Object} doc - The document object
	 * @param {string} transcript - The transcript markdown (if fetched)
	 * @returns {string} The combined markdown content
	 */
	buildNoteContent(doc, transcript, transcriptFilePath = null) {
		const sections = [];
		const noteTitle = this.generateNoteTitle(doc);

		// Add main title
		sections.push('# ' + noteTitle);

		const reviewTaskLine = this.getReviewTaskLine();
		if (reviewTaskLine) {
			sections.push('\n' + reviewTaskLine);
		}

		// Extract My Notes content
		const myNotesContent = this.extractPanelContent(doc, 'my_notes');
		if (myNotesContent && this.settings.includeMyNotes) {
			const myNotesMarkdown = this.convertProseMirrorToMarkdown(myNotesContent);
			if (myNotesMarkdown && myNotesMarkdown.trim()) {
				sections.push('\n## My Notes\n\n' + myNotesMarkdown.trim());
			}
		}

		// Extract Enhanced Notes content
		const enhancedNotesMarkdown = this.getEnhancedNotesMarkdown(doc);
		if (enhancedNotesMarkdown && this.settings.includeEnhancedNotes) {
			const metadataSection = this.extractMetadataSection(enhancedNotesMarkdown);
				const enhancedBodyMarkdown =
					this.settings.mapMetadataToFrontmatter &&
					this.settings.removeMetadataSectionFromBody &&
					metadataSection.metadata
						? metadataSection.content
						: enhancedNotesMarkdown;

				// If we have My Notes, add Enhanced Notes as a separate section
			if (myNotesContent && this.settings.includeMyNotes) {
				sections.push('\n## Enhanced Notes\n\n' + enhancedBodyMarkdown.trim());
			} else {
				// If no My Notes, just add the enhanced notes content directly
				sections.push('\n' + enhancedBodyMarkdown.trim());
			}
		}

		if (transcriptFilePath) {
			sections.push('\n' + this.generateTranscriptLink(transcriptFilePath));
		} else if (this.settings.includeFullTranscript && transcript && transcript !== 'no_transcript') {
			// Add transcript section if enabled and available
			sections.push('\n## Transcript\n\n' + transcript);
		}

		return sections.join('\n');
	}

	async processDocument(doc, authContext) {
	try {
		const title = doc.title || 'Untitled Granola Note';
		const docId = doc.id || 'unknown_id';
		const transcript = doc.transcript || 'no_transcript';

		doc = await this.ensureGranolaTemplateForDocument(doc, authContext);

		// Extract all available content
		const myNotesContent = this.extractPanelContent(doc, 'my_notes');
		const enhancedNotesMarkdown = this.getEnhancedNotesMarkdown(doc);
		const hasTranscript = this.shouldFetchTranscript() && transcript && transcript !== 'no_transcript';

		// Check if there's any content to process
		const hasMyNotes = myNotesContent && this.settings.includeMyNotes;
		const hasEnhancedNotes = enhancedNotesMarkdown && this.settings.includeEnhancedNotes;

		// If no content is available at all, skip this document
		if (!hasMyNotes && !hasEnhancedNotes && !hasTranscript) {
			console.log('Skipping document "' + title + '" - no content available (no enhanced notes, my notes, or transcript)');
			return false;
		}

		// Check if note already exists by Granola ID
		const existingFile = await this.findExistingNoteByGranolaId(docId);

		if (existingFile) {
			const existingNoteBehavior = this.getExistingNoteBehavior();
			if (existingNoteBehavior === 'never') {
				return true;
			}

			if (existingNoteBehavior === 'changed') {
				const outdated = await this.isNoteOutdated(existingFile, doc);
				if (!outdated) {
					return true;
				}

				console.log('Note "' + title + '" has been updated in Granola, re-syncing...');
			}

			// Update existing note (full update)
			try {
				// Extract attendee information
				const attendeeNames = this.extractAttendeeNames(doc);
				const attendeeTags = this.generateAttendeeTags(attendeeNames);

				// Extract folder information
				const folderNames = this.extractFolderNames(doc);
				const folderTags = this.generateFolderTags(folderNames);

				// Generate Granola URL
				const granolaUrl = this.generateGranolaUrl(docId);

				// Combine all tags
				const allTags = [...attendeeTags, ...folderTags];

				// Build frontmatter using centralized helper
				const frontmatter = this.buildFrontmatter(doc, title, allTags, granolaUrl);

				// Build the note content with all sections
				const noteContent = this.buildNoteContent(doc, transcript);
				const finalMarkdown = frontmatter + noteContent;

				await this.app.vault.process(existingFile, () => finalMarkdown);

				if (this.settings.storeTranscriptInSeparateNote && hasTranscript) {
					const transcriptFilePath = await this.writeSeparateTranscriptNote(doc, transcript);
					if (transcriptFilePath) {
						const linkedMarkdown = frontmatter + this.buildNoteContent(doc, transcript, transcriptFilePath);
						await this.app.vault.process(existingFile, () => linkedMarkdown);
					}
				}
				return true;
			} catch (updateError) {
				console.error('Error updating existing note:', updateError);
				return false;
			}
		}

		// Create new note
		// Extract attendee information
		const attendeeNames = this.extractAttendeeNames(doc);
		const attendeeTags = this.generateAttendeeTags(attendeeNames);

		// Extract folder information
		const folderNames = this.extractFolderNames(doc);
		const folderTags = this.generateFolderTags(folderNames);

		// Generate Granola URL
		const granolaUrl = this.generateGranolaUrl(docId);

		// Combine all tags
		const allTags = [...attendeeTags, ...folderTags];

		// Build frontmatter using centralized helper
		const frontmatter = this.buildFrontmatter(doc, title, allTags, granolaUrl);

		// Build the note content with all sections
		const noteContent = this.buildNoteContent(doc, transcript);
		const finalMarkdown = frontmatter + noteContent;

		const filename = this.generateFilename(doc) + '.md';
		// Use folder-based path if enabled, otherwise date-based, otherwise sync directory
		let targetDirectory;
		if (this.settings.enableGranolaFolders) {
			targetDirectory = this.generateFolderBasedPath(doc);
			await this.ensureFolderBasedDirectoryExists(targetDirectory);
		} else {
			targetDirectory = this.generateDateBasedPath(doc);
			await this.ensureDateBasedDirectoryExists(targetDirectory);
		}
		const filepath = path.join(targetDirectory, filename);

		// Check if file with same name already exists
		let finalFilepath = filepath;
		const existingFileByName = this.app.vault.getAbstractFileByPath(filepath);
		if (existingFileByName) {
			// Check if this file has the same granola_id - if so, it's the same note
			// This catches cases where findExistingNoteByGranolaId missed it due to search scope
			try {
				const existingContent = await this.app.vault.read(existingFileByName);
				const frontmatterMatch = existingContent.match(/^---\n([\s\S]*?)\n---/);
				if (frontmatterMatch) {
					const granolaIdMatch = frontmatterMatch[1].match(/granola_id:\s*(.+)$/m);
					if (granolaIdMatch && granolaIdMatch[1].trim() === docId) {
						// Same granola_id - update the existing file instead of creating duplicate
						await this.app.vault.modify(existingFileByName, finalMarkdown);

						if (this.settings.storeTranscriptInSeparateNote && hasTranscript) {
							const transcriptFilePath = await this.writeSeparateTranscriptNote(doc, transcript);
							if (transcriptFilePath) {
								const linkedMarkdown = frontmatter + this.buildNoteContent(doc, transcript, transcriptFilePath);
								await this.app.vault.modify(existingFileByName, linkedMarkdown);
							}
						}
						return true;
					}
				}
			} catch (error) {
				console.error('Error checking existing file for granola_id:', error);
			}

			// Different file with same name - handle based on user's preference
			if (this.settings.existingFileAction === 'skip') {
				// Skip creating a new file if one with the same name exists
				return false;
			} else if (this.settings.existingFileAction === 'timestamp') {
				// Create a unique filename by appending timestamp
				const timestamp = this.formatDate(doc.created_at, 'HH-mm');
				const baseFilename = this.generateFilename(doc);
				const uniqueFilename = baseFilename + '_' + timestamp + '.md';
				finalFilepath = path.join(targetDirectory, uniqueFilename);

				// Check if the unique filename also exists
				const existingUniqueFile = this.app.vault.getAbstractFileByPath(finalFilepath);
				if (existingUniqueFile) {
					return false;
				}
			}
		}

		await this.app.vault.create(finalFilepath, finalMarkdown);

		if (this.settings.storeTranscriptInSeparateNote && hasTranscript) {
			const transcriptFilePath = await this.writeSeparateTranscriptNote(doc, transcript);
			if (transcriptFilePath) {
				const summaryFile = this.app.vault.getAbstractFileByPath(finalFilepath);
				if (summaryFile) {
					const linkedMarkdown = frontmatter + this.buildNoteContent(doc, transcript, transcriptFilePath);
					await this.app.vault.modify(summaryFile, linkedMarkdown);
				}
			}
		}
		return true;

	} catch (error) {
		console.error('Error processing document:', error);
		return false;
	}
}

	async ensureDirectoryExists() {
		try {
			const folder = this.app.vault.getFolderByPath(this.settings.syncDirectory);
			if (!folder) {
				await this.app.vault.createFolder(this.settings.syncDirectory);
			}
		} catch (error) {
			console.error('Error creating directory:', error);
		}
	}

	async updateDailyNote(todaysNotes) {
		try {
			const dailyNote = await this.getDailyNote();
			if (!dailyNote) {
				return;
			}

			let content = await this.app.vault.read(dailyNote);
			const sectionHeader = this.settings.dailyNoteSectionName;
			
			const notesList = todaysNotes
				.sort((a, b) => a.time.localeCompare(b.time))
				.map(note => '- ' + note.time + ' [[' + note.actualFilePath + '|' + note.title + ']]')
				.join('\n');
			
			const granolaSection = sectionHeader + '\n' + notesList;

			// Use MetadataCache to find existing headings
			const fileCache = this.app.metadataCache.getFileCache(dailyNote);
			const headings = fileCache?.headings || [];
			
			// Look for existing section by heading text
			const existingHeading = headings.find(heading => 
				heading.heading.trim() === sectionHeader.replace(/^#+\s*/, '').trim()
			);
			
			if (existingHeading) {
				// Found existing section, replace content
				const lines = content.split('\n');
				const sectionLineNum = existingHeading.position.start.line;
				
				// Find the end of this section (next heading of same or higher level, or end of file)
				let endLineNum = lines.length;
				for (const heading of headings) {
					if (heading.position.start.line > sectionLineNum && heading.level <= existingHeading.level) {
						endLineNum = heading.position.start.line;
						break;
					}
				}
				
				const beforeSection = lines.slice(0, sectionLineNum).join('\n');
				const afterSection = lines.slice(endLineNum).join('\n');
				content = beforeSection + '\n' + granolaSection + '\n' + afterSection;
			} else {
				// Section not found, append to end
				content += '\n\n' + granolaSection;
			}

			await this.app.vault.process(dailyNote, () => content);
			
		} catch (error) {
			console.error('Error updating daily note:', error);
		}
	}

	async updatePeriodicNote(todaysNotes) {
		try {
			const periodicNote = await this.getPeriodicNote();
			if (!periodicNote) {
				return;
			}

			let content = await this.app.vault.read(periodicNote);
			const sectionHeader = this.settings.periodicNoteSectionName;
			
			const notesList = todaysNotes
				.sort((a, b) => a.time.localeCompare(b.time))
				.map(note => '- ' + note.time + ' [[' + note.actualFilePath + '|' + note.title + ']]')
				.join('\n');
			
			const granolaSection = sectionHeader + '\n' + notesList;

			// Use MetadataCache to find existing headings
			const fileCache = this.app.metadataCache.getFileCache(periodicNote);
			const headings = fileCache?.headings || [];
			
			// Look for existing section by heading text
			const existingHeading = headings.find(heading => 
				heading.heading.trim() === sectionHeader.replace(/^#+\s*/, '').trim()
			);
			
			if (existingHeading) {
				// Found existing section, replace content
				const lines = content.split('\n');
				const sectionLineNum = existingHeading.position.start.line;
				
				// Find the end of this section (next heading of same or higher level, or end of file)
				let endLineNum = lines.length;
				for (const heading of headings) {
					if (heading.position.start.line > sectionLineNum && heading.level <= existingHeading.level) {
						endLineNum = heading.position.start.line;
						break;
					}
				}
				
				const beforeSection = lines.slice(0, sectionLineNum).join('\n');
				const afterSection = lines.slice(endLineNum).join('\n');
				content = beforeSection + '\n' + granolaSection + '\n' + afterSection;
			} else {
				// Section not found, append to end
				content += '\n\n' + granolaSection;
			}

			await this.app.vault.process(periodicNote, () => content);
			
		} catch (error) {
			console.error('Error updating periodic note:', error);
		}
	}

	async getDailyNote() {
		try {
			const today = new Date();

			// Try to get Daily Notes plugin settings from Obsidian
			const dailyNotesPlugin = this.app.internalPlugins.getPluginById('daily-notes');
			if (dailyNotesPlugin?.enabled) {
				const dailyNotesSettings = dailyNotesPlugin.instance?.options || {};
				const dateFormat = dailyNotesSettings.format || 'YYYY-MM-DD';
				const folder = dailyNotesSettings.folder || '';

				// Format today's date using the configured format
				const todayFormatted = this.formatDateWithPattern(today, dateFormat);

				// Build the expected path
				const expectedPath = folder
					? `${folder}/${todayFormatted}.md`
					: `${todayFormatted}.md`;

				// Try to get the file directly by path
				const dailyNote = this.app.vault.getAbstractFileByPath(expectedPath);
				if (dailyNote) {
					return dailyNote;
				}

				// Fallback: search for file by exact basename match
				const files = this.app.vault.getMarkdownFiles();
				const matchedFile = files.find(f => f.basename === todayFormatted);
				if (matchedFile) {
					return matchedFile;
				}
			}

			// Fallback for when Daily Notes plugin is disabled: use legacy fuzzy matching
			const year = today.getFullYear();
			const month = String(today.getMonth() + 1).padStart(2, '0');
			const day = String(today.getDate()).padStart(2, '0');

			const searchFormats = [
				`${day}-${month}-${year}`, // DD-MM-YYYY
				`${year}-${month}-${day}`, // YYYY-MM-DD
				`${month}-${day}-${year}`, // MM-DD-YYYY
				`${day}.${month}.${year}`, // DD.MM.YYYY
				`${year}/${month}/${day}`, // YYYY/MM/DD
				`${day}/${month}/${year}`, // DD/MM/YYYY
			];

			const files = this.app.vault.getMarkdownFiles();

			for (const file of files) {
				if (file.path.includes('Daily')) {
					for (const format of searchFormats) {
						if (file.path.includes(format)) {
							return file;
						}
					}
				}
			}

			return null;
		} catch (error) {
			console.error('Error getting daily note:', error);
			return null;
		}
	}

	formatDateWithPattern(date, pattern) {
		const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
		const dayNamesFull = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
		const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
		const monthNamesFull = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

		const year = date.getFullYear();
		const month = date.getMonth();
		const day = date.getDate();
		const dayOfWeek = date.getDay();

		// Order matters: replace longer patterns first to avoid partial matches
		return pattern
			.replace(/YYYY/g, year)
			.replace(/YY/g, String(year).slice(-2))
			.replace(/MMMM/g, monthNamesFull[month])
			.replace(/MMM/g, monthNames[month])
			.replace(/MM/g, String(month + 1).padStart(2, '0'))
			.replace(/M(?![ao])/g, String(month + 1))
			.replace(/dddd/g, dayNamesFull[dayOfWeek])
			.replace(/ddd/g, dayNames[dayOfWeek])
			.replace(/DD/g, String(day).padStart(2, '0'))
			.replace(/D(?![ae])/g, String(day));
	}

	isPeriodicNotesPluginAvailable() {
		return this.app.plugins.enabledPlugins.has('periodic-notes');
	}

	async getPeriodicNote() {
		try {
			if (!this.isPeriodicNotesPluginAvailable()) {
				return null;
			}

			// Since the Periodic Notes API is not accessible, let's try a different approach
			// Let's try to find the daily note directly by looking for it in the vault
			
			// Get today's date
			const today = new Date();
			const todayFormatted = today.toISOString().split('T')[0]; // YYYY-MM-DD format
			
			// Search for today's daily note in the vault
			const files = this.app.vault.getMarkdownFiles();
			
			// Look for files that might be today's daily note
			// Priority order: exact date match, then files in Daily Notes folder, then any file with today's date
			const possibleDailyNotes = files.filter(file => {
				// First priority: exact date match in filename
				if (file.name === todayFormatted + '.md' || file.name === todayFormatted) {
					return true;
				}
				// Second priority: files in Daily Notes folder with today's date
				if (file.path.includes('Daily') && (file.name.includes(todayFormatted) || file.path.includes(todayFormatted))) {
					return true;
				}
				// Third priority: any file with today's date
				return file.name.includes(todayFormatted) || 
					   file.path.includes(todayFormatted) ||
					   file.name.includes(today.toDateString().split(' ')[2]) || // Day of month
					   file.name.includes(today.getDate().toString());
			});
			
			// Sort by priority: exact date match first, then Daily Notes folder, then others
			possibleDailyNotes.sort((a, b) => {
				// Exact date match gets highest priority
				if (a.name === todayFormatted + '.md' || a.name === todayFormatted) return -1;
				if (b.name === todayFormatted + '.md' || b.name === todayFormatted) return 1;
				
				// Daily Notes folder gets second priority
				if (a.path.includes('Daily') && !b.path.includes('Daily')) return -1;
				if (b.path.includes('Daily') && !a.path.includes('Daily')) return 1;
				
				// Otherwise maintain original order
				return 0;
			});
			
			// Return the first match
			if (possibleDailyNotes.length > 0) {
				return possibleDailyNotes[0];
			}
			
			return null;
		} catch (error) {
			console.error('Error getting periodic note:', error);
			return null;
		}
	}

	extractAttendeeNames(doc) {
		const attendees = [];
		const processedEmails = new Set(); // Track processed emails to avoid duplicates
		
		try {
			// Check the people field for attendee information (enhanced with detailed person data)
			if (doc.people && Array.isArray(doc.people)) {
				for (const person of doc.people) {
					let name = null;
					
					// Try to get name from various fields
					if (person.name) {
						name = person.name;
					} else if (person.display_name) {
						name = person.display_name;
					} else if (person.details && person.details.person && person.details.person.name) {
						// Use the detailed person information if available
						const personDetails = person.details.person.name;
						if (personDetails.fullName) {
							name = personDetails.fullName;
						} else if (personDetails.givenName && personDetails.familyName) {
							name = `${personDetails.givenName} ${personDetails.familyName}`;
						} else if (personDetails.givenName) {
							name = personDetails.givenName;
						}
					} else if (person.email) {
						// Extract name from email if no display name
						const emailName = person.email.split('@')[0].replace(/[._]/g, ' ');
						name = emailName;
					}
					
					if (name && !attendees.includes(name)) {
						attendees.push(name);
						if (person.email) {
							processedEmails.add(person.email);
						}
					}
				}
			}
			
			// Also check google_calendar_event for additional attendee info
			if (doc.google_calendar_event && doc.google_calendar_event.attendees) {
				for (const attendee of doc.google_calendar_event.attendees) {
					// Skip if we've already processed this email
					if (attendee.email && processedEmails.has(attendee.email)) {
						continue;
					}
					
					if (attendee.displayName && !attendees.includes(attendee.displayName)) {
						attendees.push(attendee.displayName);
						if (attendee.email) {
							processedEmails.add(attendee.email);
						}
					} else if (attendee.email && !attendees.some(name => name.includes(attendee.email.split('@')[0]))) {
						const emailName = attendee.email.split('@')[0].replace(/[._]/g, ' ');
						attendees.push(emailName);
						processedEmails.add(attendee.email);
					}
				}
			}
			
			return attendees;
		} catch (error) {
			console.error('Error extracting attendee names:', error);
			return [];
		}
	}

	generateAttendeeTags(attendees) {
		if (!this.settings.includeAttendeeTags || !attendees || attendees.length === 0) {
			return [];
		}
		
		const tags = [];
		
		for (const attendee of attendees) {
			// Skip if this is the user's own name (case-insensitive, exact match)
			if (this.settings.excludeMyNameFromTags && this.settings.myName && 
				attendee.toLowerCase().trim() === this.settings.myName.toLowerCase().trim()) {
				continue;
			}
			
			// Convert name to valid tag format
			// Remove special characters, replace spaces with hyphens, convert to lowercase
			let cleanName = attendee
				.replace(/[^\w\s-]/g, '') // Remove special chars except spaces and hyphens
				.trim()
				.replace(/\s+/g, '-') // Replace spaces with hyphens
				.toLowerCase();
			
			// Use the customizable tag template
			let tag = this.settings.attendeeTagTemplate.replace('{name}', cleanName);
			
			// Ensure the tag is valid (no double slashes, etc.)
			tag = tag.replace(/\/+/g, '/').replace(/^\/|\/$/g, '');
			
			if (tag && !tags.includes(tag)) {
				tags.push(tag);
			}
		}
		
		return tags;
	}

	extractFolderNames(doc) {
		const folderNames = [];
		
		try {
			// Check if folder support is enabled and we have folder mapping
			if (this.settings.enableGranolaFolders && this.documentToFolderMap) {
				const folder = this.documentToFolderMap[doc.id];
				if (folder && folder.title) {
					folderNames.push(folder.title);
				}
			}
			
			return folderNames;
		} catch (error) {
			console.error('Error extracting folder names:', error);
			return [];
		}
	}

	findWorkspaceName(workspaceId) {
		if (!this.workspaces || !workspaceId) {
			return null;
		}
		
		try {
			// Try different possible structures for workspaces response
			if (Array.isArray(this.workspaces)) {
				const workspace = this.workspaces.find(ws => ws.id === workspaceId);
				if (workspace && workspace.name) {
					return workspace.name;
				}
			} else if (this.workspaces.workspaces && Array.isArray(this.workspaces.workspaces)) {
				const workspace = this.workspaces.workspaces.find(ws => ws.id === workspaceId);
				if (workspace && workspace.name) {
					return workspace.name;
				}
			} else if (this.workspaces.lists && Array.isArray(this.workspaces.lists)) {
				const list = this.workspaces.lists.find(l => l.id === workspaceId);
				if (list && list.name) {
					return list.name;
				}
			}
			
			return null;
		} catch (error) {
			console.error('Error finding workspace name:', error);
			return null;
		}
	}

	generateFolderTags(folderNames) {
		if (!this.settings.includeFolderTags || !folderNames || folderNames.length === 0) {
			return [];
		}
		
		try {
			const tags = [];
			
			for (const folderName of folderNames) {
				if (!folderName) continue;
				
				// Convert folder name to valid tag format
				// Remove special characters, replace spaces with hyphens, convert to lowercase
				let cleanName = folderName
					.replace(/[^\w\s-]/g, '') // Remove special chars except spaces and hyphens
					.trim()
					.replace(/\s+/g, '-') // Replace spaces with hyphens
					.toLowerCase();
				
				// Use the customizable tag template
				let tag = this.settings.folderTagTemplate.replace('{name}', cleanName);
				
				// Ensure the tag is valid (no double slashes, etc.)
				tag = tag.replace(/\/+/g, '/').replace(/^\/|\/$/g, '');
				
				if (tag && !tags.includes(tag)) {
					tags.push(tag);
				}
			}
			
			return tags;
		} catch (error) {
			console.error('Error generating folder tags:', error);
			return [];
		}
	}

	generateGranolaUrl(docId) {
		if (!this.settings.includeGranolaUrl || !docId) {
			return null;
		}
		
		try {
			// Construct the Granola notes URL using the correct format
			return `https://notes.granola.ai/d/${docId}`;
		} catch (error) {
			console.error('Error generating Granola URL:', error);
			return null;
		}
	}

	/**
	 * Format a date string according to frontmatter date format settings
	 * @param {string} dateString - ISO date string from Granola
	 * @returns {string} - Formatted date string
	 */
	formatFrontmatterDate(dateString) {
		if (!dateString) return '';
		
		try {
			const date = new Date(dateString);
			if (isNaN(date.getTime())) return dateString;
			
			switch (this.settings.frontmatterDateFormat) {
				case 'date-only':
					// Return just the date portion: YYYY-MM-DD
					return date.toISOString().split('T')[0];
				case 'custom':
					// Apply custom format
					return this.applyCustomDateFormat(date, this.settings.customDateFormat);
				case 'iso':
				default:
					// Return full ISO string (original behaviour)
					return dateString;
			}
		} catch (error) {
			console.error('Error formatting date:', error);
			return dateString;
		}
	}

	formatObsidianDateProperty(dateString) {
		if (!dateString) return '';

		try {
			const date = new Date(dateString);
			if (isNaN(date.getTime())) return '';
			return date.toISOString().split('T')[0];
		} catch (error) {
			console.error('Error formatting Obsidian date property:', error);
			return '';
		}
	}

	/**
	 * Apply a custom date format string to a date
	 * Supports: YYYY, MM, DD, HH, mm, ss
	 * @param {Date} date - Date object
	 * @param {string} format - Format string
	 * @returns {string} - Formatted date string
	 */
	applyCustomDateFormat(date, format) {
		const pad = (n) => n.toString().padStart(2, '0');
		return format
			.replace('YYYY', date.getFullYear())
			.replace('MM', pad(date.getMonth() + 1))
			.replace('DD', pad(date.getDate()))
			.replace('HH', pad(date.getHours()))
			.replace('mm', pad(date.getMinutes()))
			.replace('ss', pad(date.getSeconds()));
	}

	/**
	 * Parse additional frontmatter from settings
	 * @returns {string} - Additional frontmatter lines
	 */
	parseAdditionalFrontmatter() {
		if (!this.settings.additionalFrontmatter || !this.settings.additionalFrontmatter.trim()) {
			return '';
		}
		
		try {
			// Split by newlines, filter empty lines, ensure each line ends with newline
			const lines = this.settings.additionalFrontmatter
				.split('\n')
				.map(line => line.trim())
				.filter(line => line.length > 0 && line.includes(':'));
			
			if (lines.length === 0) return '';
			return lines.join('\n') + '\n';
		} catch (error) {
			console.error('Error parsing additional frontmatter:', error);
			return '';
		}
	}

	getAdditionalFrontmatterKeys() {
		if (!this.settings.additionalFrontmatter || !this.settings.additionalFrontmatter.trim()) {
			return new Set();
		}

		const keys = new Set();
		for (const line of this.settings.additionalFrontmatter.split('\n')) {
			const trimmed = line.trim();
			if (!trimmed || !trimmed.includes(':')) {
				continue;
			}
			const key = trimmed.split(':')[0].trim();
			if (key) {
				keys.add(key);
			}
		}

		return keys;
	}

	/**
	 * Build frontmatter string for a document
	 * @param {object} doc - Granola document
	 * @param {string} title - Note title
	 * @param {array} allTags - Combined tags array
	 * @param {string} granolaUrl - Granola URL (or null)
	 * @returns {string} - Complete frontmatter string with delimiters
	 */
	buildFrontmatter(doc, title, allTags, granolaUrl) {
		const docId = doc.id;
		let frontmatter = '---\n';
		const additionalFrontmatterKeys = this.getAdditionalFrontmatterKeys();
		const syncUpdatedAt = this.getDocumentSyncUpdatedAt(doc);
		const noteDate = this.formatObsidianDateProperty(doc.created_at);
		
		// granola_id is always included (required for duplicate detection)
		frontmatter += 'granola_id: ' + docId + '\n';
		
		// Title (optional based on settings)
		if (this.settings.includeTitle) {
			const escapedTitle = title.replace(/"/g, '\\"');
			frontmatter += 'title: "' + escapedTitle + '"\n';
		}
		
		// Granola URL (optional based on existing setting)
		if (granolaUrl) {
			frontmatter += 'granola_url: "' + granolaUrl + '"\n';
		}

		if (noteDate && !additionalFrontmatterKeys.has('date')) {
			frontmatter += 'date: ' + noteDate + '\n';
		}
		
		// Dates (optional based on settings)
		if (this.settings.includeDates) {
			if (doc.created_at) {
				frontmatter += 'created_at: ' + this.formatFrontmatterDate(doc.created_at) + '\n';
			}
			if (syncUpdatedAt) {
				frontmatter += 'updated_at: ' + this.formatFrontmatterDate(syncUpdatedAt) + '\n';
			}
		}

		if (syncUpdatedAt && !additionalFrontmatterKeys.has('granola_updated_at')) {
			frontmatter += 'granola_updated_at: ' + syncUpdatedAt + '\n';
		}

		if (this.settings.mapMetadataToFrontmatter) {
			const enhancedNotesMarkdown = this.getEnhancedNotesMarkdown(doc);
			const { metadata } = this.extractMetadataSection(enhancedNotesMarkdown);

			if (metadata && typeof metadata === 'object') {
				if (metadata.org && !additionalFrontmatterKeys.has('org')) {
					const mappedOrg = this.applyMetadataValueTemplate(this.settings.metadataOrgTemplate, metadata.org);
					if (mappedOrg) {
						frontmatter += 'org: ' + this.escapeYamlString(mappedOrg) + '\n';
					}
				}
				if (!additionalFrontmatterKeys.has('people')) {
					const mappedPeople = Array.isArray(metadata.people)
						? [...new Set(metadata.people
							.map((person) => this.applyMetadataValueTemplate(this.settings.metadataPersonTemplate, person))
							.filter(Boolean))]
						: metadata.people;
					frontmatter += this.formatFrontmatterList('people', mappedPeople);
				}
				if (!additionalFrontmatterKeys.has('topics')) {
					frontmatter += this.formatFrontmatterList('topics', metadata.topics);
				}
				if (!additionalFrontmatterKeys.has('loc')) {
					frontmatter += this.formatFrontmatterList('loc', metadata.loc);
				}
				if (metadata.meeting_type && !additionalFrontmatterKeys.has('meeting_type')) {
					frontmatter += 'meeting_type: ' + this.escapeYamlString(metadata.meeting_type) + '\n';
				}
			}
		}
		
		// Tags
		if (allTags.length > 0) {
			frontmatter += 'tags:\n';
			for (const tag of allTags) {
				frontmatter += '  - ' + tag + '\n';
			}
		}
		
		// Additional custom frontmatter
		const additionalFm = this.parseAdditionalFrontmatter();
		if (additionalFm) {
			frontmatter += additionalFm;
		}
		
		frontmatter += '---\n\n';
		return frontmatter;
	}

	async isNoteOutdated(existingFile, doc) {
		const syncUpdatedAt = this.getDocumentSyncUpdatedAt(doc);
		if (!syncUpdatedAt) return false;
		try {
			const content = await this.app.vault.read(existingFile);
			const frontmatterMatch = content.match(/^---\n([\s\S]*?)\n---/);
			if (frontmatterMatch) {
				const granolaUpdatedAtMatch = frontmatterMatch[1].match(/granola_updated_at:\s*(.+)$/m);
				const updatedAtMatch = frontmatterMatch[1].match(/updated_at:\s*(.+)$/m);
				const storedUpdatedAt = granolaUpdatedAtMatch ? granolaUpdatedAtMatch[1].trim() : updatedAtMatch ? updatedAtMatch[1].trim() : '';
				if (storedUpdatedAt) {
					const existingDate = new Date(storedUpdatedAt);
					const granolaDate = new Date(syncUpdatedAt);
					return granolaDate > existingDate;
				}
			}
			// No updated_at in frontmatter - treat as outdated so it gets updated
			return true;
		} catch (error) {
			console.error('Error checking if note is outdated:', error);
			return false;
		}
	}

	async reorganizeExistingNotes(quiet = false) {
		try {
			this.updateStatusBar('Syncing');
			if (!quiet) {
				new obsidian.Notice('Starting reorganization of existing Granola notes...');
			}

			// Get all markdown files in the vault
			const allFiles = this.app.vault.getMarkdownFiles();
			const granolaFiles = [];

			// Find all files with granola_id in frontmatter
			for (const file of allFiles) {
				try {
					const content = await this.app.vault.read(file);
					const frontmatterMatch = content.match(/^---\n([\s\S]*?)\n---/);

					if (frontmatterMatch) {
						const frontmatter = frontmatterMatch[1];
						const granolaIdMatch = frontmatter.match(/granola_id:\s*(.+)$/m);
						const createdAtMatch = frontmatter.match(/created_at:\s*(.+)$/m);

						if (granolaIdMatch) {
							granolaFiles.push({
								file: file,
								granolaId: granolaIdMatch[1].trim(),
								createdAt: createdAtMatch ? createdAtMatch[1].trim() : null
							});
						}
					}
				} catch (error) {
					console.error('Error reading file for reorganization:', file.path, error);
				}
			}

			if (granolaFiles.length === 0) {
				if (!quiet) {
					new obsidian.Notice('No Granola notes found to reorganize');
				}
				this.updateStatusBar('Idle');
				return;
			}

			// Fetch folders if folder support is enabled
			let folders = null;
			if (this.settings.enableGranolaFolders) {
				const authContext = await this.loadCredentials();
				if (authContext) {
					folders = await this.fetchGranolaFolders(authContext);
					if (folders) {
						// Create a mapping of document ID to folder for quick lookup
						this.documentToFolderMap = {};
						for (const folder of folders) {
							if (folder.document_ids) {
								for (const docId of folder.document_ids) {
									this.documentToFolderMap[docId] = folder;
								}
							}
						}
					}
				}
			}

			let movedCount = 0;
			let errorCount = 0;

			// Process each Granola file
			for (const granolaFile of granolaFiles) {
				try {
					// Create a mock document object with the information we have
					const mockDoc = {
						id: granolaFile.granolaId,
						created_at: granolaFile.createdAt
					};

					// Determine the correct target directory based on settings
					let targetDirectory;
					if (this.settings.enableGranolaFolders) {
						targetDirectory = this.generateFolderBasedPath(mockDoc);
						await this.ensureFolderBasedDirectoryExists(targetDirectory);
					} else if (this.settings.enableDateBasedFolders) {
						targetDirectory = this.generateDateBasedPath(mockDoc);
						await this.ensureDateBasedDirectoryExists(targetDirectory);
					} else {
						targetDirectory = this.settings.syncDirectory;
					}

					// Get the current directory of the file
					const currentDirectory = granolaFile.file.parent.path;

					// Check if the file is already in the correct location
					if (currentDirectory === targetDirectory) {
						continue; // File is already in the right place
					}

					// Construct the new file path
					const newFilePath = path.join(targetDirectory, granolaFile.file.name);

					// Check if a file with the same name already exists in the target directory
					const existingFile = this.app.vault.getAbstractFileByPath(newFilePath);
					if (existingFile && existingFile.path !== granolaFile.file.path) {
						console.warn('Cannot move file - target already exists:', newFilePath);
						errorCount++;
						continue;
					}

					// Move the file to the new location
					await this.app.fileManager.renameFile(granolaFile.file, newFilePath);
					movedCount++;

				} catch (error) {
					console.error('Error reorganizing file:', granolaFile.file.path, error);
					errorCount++;
				}
			}

			if (!quiet) {
				const message = `Reorganization complete! Moved ${movedCount} note(s). ${errorCount > 0 ? errorCount + ' error(s) occurred.' : ''}`;
				new obsidian.Notice(message, 8000);
			}
			// Always log to console for debugging
			if (movedCount > 0 || errorCount > 0) {
				console.log(`Reorganization: Moved ${movedCount} note(s), ${errorCount} error(s)`);
			}
			this.updateStatusBar('Idle');

		} catch (error) {
			console.error('Error during reorganization:', error);
			if (!quiet) {
				new obsidian.Notice('Error during reorganization. Check console for details.');
			}
			this.updateStatusBar('Error', 'reorganization failed');
		}
	}
}

class GranolaSyncSettingTab extends obsidian.PluginSettingTab {
	constructor(app, plugin) {
		super(app, plugin);
		this.plugin = plugin;
	}

	display() {
		const containerEl = this.containerEl;
		containerEl.empty();

		// Debug: Log all settings
		console.log('All plugin settings:', this.plugin.settings);
		console.log('enableGranolaFolders value:', this.plugin.settings.enableGranolaFolders);
		console.log('enableGranolaFolders type:', typeof this.plugin.settings.enableGranolaFolders);

		new obsidian.Setting(containerEl)
			.setName('Note prefix')
			.setDesc('Optional prefix to add to all synced note titles')
			.addText(text => {
				text.setPlaceholder('granola-');
				text.setValue(this.plugin.settings.notePrefix);
				text.onChange(async (value) => {
					this.plugin.settings.notePrefix = value;
					await this.plugin.saveSettings();
				});
			});

		new obsidian.Setting(containerEl)
			.setName('Auth key path')
			.setDesc('Path to your Granola authentication key file')
			.addText(text => {
				text.setPlaceholder(getDefaultAuthPath());
				text.setValue(this.plugin.settings.authKeyPath);
				text.onChange(async (value) => {
					this.plugin.settings.authKeyPath = value;
					await this.plugin.saveSettings();
				});
			});

		new obsidian.Setting(containerEl)
			.setName('Date format')
			.setDesc('Format for dates in filenames. Use YYYY (year), MM (month), DD (day)')
			.addText(text => {
				text.setPlaceholder('YYYY-MM-DD');
				text.setValue(this.plugin.settings.dateFormat);
				text.onChange(async (value) => {
					this.plugin.settings.dateFormat = value;
					await this.plugin.saveSettings();
				});
			});


		// Create a heading for note content settings
		containerEl.createEl('h3', {text: 'Note content'});

		new obsidian.Setting(containerEl)
			.setName('Include My Notes')
			.setDesc('Include your personal notes from Granola in a "## My Notes" section. These are the notes you write yourself during meetings.')
			.addToggle(toggle => {
				toggle.setValue(this.plugin.settings.includeMyNotes);
				toggle.onChange(async (value) => {
					this.plugin.settings.includeMyNotes = value;
					await this.plugin.saveSettings();
				});
			});

		new obsidian.Setting(containerEl)
			.setName('Include Enhanced Notes')
			.setDesc('Include AI-generated enhanced notes from Granola in a "## Enhanced Notes" section. These are the AI summaries Granola creates.')
			.addToggle(toggle => {
				toggle.setValue(this.plugin.settings.includeEnhancedNotes);
				toggle.onChange(async (value) => {
					this.plugin.settings.includeEnhancedNotes = value;
					await this.plugin.saveSettings();
				});
			});

		new obsidian.Setting(containerEl)
			.setName('Add review task to synced notes')
			.setDesc('Insert "- [ ] Review imported Granola note" near the top of each synced summary note. If the note is later rewritten, the task is reset to unchecked.')
			.addToggle(toggle => {
				toggle.setValue(this.plugin.settings.includeReviewTask);
				toggle.onChange(async (value) => {
					this.plugin.settings.includeReviewTask = value;
					await this.plugin.saveSettings();
				});
			});

		new obsidian.Setting(containerEl)
			.setName('Include full transcript')
			.setDesc('Include the full meeting transcript in each note under a "## Transcript" section. This requires an additional API call per note and may slow down sync.')
			.addToggle(toggle => {
				toggle.setValue(this.plugin.settings.includeFullTranscript);
				toggle.onChange(async (value) => {
					this.plugin.settings.includeFullTranscript = value;
					await this.plugin.saveSettings();
					this.display();
				});
			});

		new obsidian.Setting(containerEl)
			.setName('Store transcript in separate note')
			.setDesc('Write the transcript to a separate note and add a transcript link to the main note instead of embedding the full transcript inline.')
			.addToggle(toggle => {
				toggle.setValue(this.plugin.settings.storeTranscriptInSeparateNote);
				toggle.onChange(async (value) => {
					this.plugin.settings.storeTranscriptInSeparateNote = value;
					await this.plugin.saveSettings();
					this.display();
				});
			});

		if (this.plugin.settings.storeTranscriptInSeparateNote) {
			new obsidian.Setting(containerEl)
				.setName('Transcript directory')
				.setDesc('Folder where separate transcript notes should be stored.')
				.addText(text => {
					text.setPlaceholder('Granola Transcripts');
					text.setValue(this.plugin.settings.transcriptDirectory);
					text.onChange(async (value) => {
						this.plugin.settings.transcriptDirectory = value || 'Granola Transcripts';
						await this.plugin.saveSettings();
					});
				});
		}

		// Create a heading for filename settings
		containerEl.createEl('h3', {text: 'Filename settings'});

		new obsidian.Setting(containerEl)
			.setName('Filename template')
			.setDesc('Template for filenames. Use {title}, {created_date}, {updated_date}, etc.')
			.addText(text => {
				text.setPlaceholder('{created_date}_{title}');
				text.setValue(this.plugin.settings.filenameTemplate);
				text.onChange(async (value) => {
					this.plugin.settings.filenameTemplate = value;
					await this.plugin.saveSettings();
				});
			});

		new obsidian.Setting(containerEl)
			.setName('Filename word separator')
			.setDesc('Character to separate words in filenames (underscore, hyphen, space, or none)')
			.addDropdown(dropdown => {
				dropdown.addOption('_', 'Underscore (_) - Team_Standup');
				dropdown.addOption('-', 'Hyphen (-) - Team-Standup');
				dropdown.addOption(' ', 'Space ( ) - Team Standup');
				dropdown.addOption('', 'None - TeamStandup');

				dropdown.setValue(this.plugin.settings.filenameSeparator);
				dropdown.onChange(async (value) => {
					this.plugin.settings.filenameSeparator = value;
					await this.plugin.saveSettings();
				});
			});

		new obsidian.Setting(containerEl)
			.setName('When file already exists by name')
			.setDesc('Choose what to do when syncing a note with a filename that already exists')
			.addDropdown(dropdown => {
				dropdown.addOption('timestamp', 'Create timestamped version (e.g., filename_13-40.md)');
				dropdown.addOption('skip', 'Skip the file and don\'t create a new version');

				dropdown.setValue(this.plugin.settings.existingFileAction);
				dropdown.onChange(async (value) => {
					this.plugin.settings.existingFileAction = value;
					await this.plugin.saveSettings();
				});
			});

		new obsidian.Setting(containerEl)
			.setName('Auto-sync frequency')
			.setDesc('How often to automatically sync notes')
			.addDropdown(dropdown => {
				dropdown.addOption('0', 'Never');
				dropdown.addOption('60000', 'Every 1 minute');
				dropdown.addOption('300000', 'Every 5 minutes');
				dropdown.addOption('600000', 'Every 10 minutes');
				dropdown.addOption('1800000', 'Every 30 minutes');
				dropdown.addOption('3600000', 'Every 1 hour');
				dropdown.addOption('21600000', 'Every 6 hours');
				dropdown.addOption('86400000', 'Every 24 hours');

				dropdown.setValue(String(this.plugin.settings.autoSyncFrequency));
				dropdown.onChange(async (value) => {
					this.plugin.settings.autoSyncFrequency = parseInt(value);
					await this.plugin.saveSettings();

					const label = this.plugin.getFrequencyLabel(parseInt(value));
					new obsidian.Notice('Auto-sync updated: ' + label);
				});
			});

		new obsidian.Setting(containerEl)
			.setName('Sync all historical notes')
			.setDesc('When enabled, sync will fetch ALL notes from Granola (not just the most recent 100). This may take longer on first sync but ensures all historical notes are included.')
			.addToggle(toggle => {
				toggle.setValue(this.plugin.settings.syncAllHistoricalNotes);
				toggle.onChange(async (value) => {
					this.plugin.settings.syncAllHistoricalNotes = value;
					await this.plugin.saveSettings();
					this.display(); // Refresh to show/hide document limit setting
				});
			});

		// Show document limit setting only when NOT syncing all historical notes
		if (!this.plugin.settings.syncAllHistoricalNotes) {
			new obsidian.Setting(containerEl)
				.setName('Document sync limit')
				.setDesc('Maximum number of documents to sync from Granola (most recent notes will be synced first)')
				.addText(text => {
					text.setPlaceholder('100');
					text.setValue(String(this.plugin.settings.documentSyncLimit));
					text.onChange(async (value) => {
						const limit = parseInt(value);
						if (!isNaN(limit) && limit > 0) {
							this.plugin.settings.documentSyncLimit = limit;
							await this.plugin.saveSettings();
						} else {
							new obsidian.Notice('Please enter a valid positive number');
						}
					});
				});
		}

		new obsidian.Setting(containerEl)
			.setName('Existing note behavior')
			.setDesc('Choose what to do when a note with the same Granola ID already exists in your vault.')
			.addDropdown(dropdown => {
				dropdown.addOption('never', 'Never update existing notes');
				dropdown.addOption('changed', 'Update when Granola changed');
				dropdown.addOption('always', 'Always rewrite existing notes');
				dropdown.setValue(this.plugin.getExistingNoteBehavior());
				dropdown.onChange(async (value) => {
					this.plugin.settings.existingNoteBehavior = value;
					await this.plugin.saveSettings();
				});
			});

		new obsidian.Setting(containerEl)
			.setName('Enable date-based folders')
			.setDesc('Organize notes into subfolders based on their creation date')
			.addToggle(toggle => {
				toggle.setValue(this.plugin.settings.enableDateBasedFolders);
				toggle.onChange(async (value) => {
					this.plugin.settings.enableDateBasedFolders = value;
					await this.plugin.saveSettings();
					this.display(); // Refresh to update auto-reorganize toggle state
				});
			});

		new obsidian.Setting(containerEl)
			.setName('Date folder format')
			.setDesc('Format for date-based folder structure. Examples: "YYYY-MM-DD" or "YYYY/MM/DD" subfolders')
			.addText(text => {
				text.setPlaceholder('YYYY/MM/DD');
				text.setValue(this.plugin.settings.dateFolderFormat);
				text.onChange(async (value) => {
					this.plugin.settings.dateFolderFormat = value || 'YYYY/MM/DD';
					await this.plugin.saveSettings();
				});
			});

		// Create experimental section header
		containerEl.createEl('h4', {text: '🧪 Experimental features'});
		
		const experimentalWarning = containerEl.createEl('div', { cls: 'setting-item' });
		experimentalWarning.createEl('div', { cls: 'setting-item-info' });
		const warningNameEl = experimentalWarning.createEl('div', { cls: 'setting-item-name' });
		warningNameEl.setText('⚠️ Please backup your vault');
		const warningDescEl = experimentalWarning.createEl('div', { cls: 'setting-item-description' });
		warningDescEl.setText('⚠️ The features below are experimental and may create duplicate notes if not used carefully. Please backup your vault before changing these settings.');

		new obsidian.Setting(containerEl)
			.setName('Enable Granola Template Management')
			.setDesc('Automatically ensure a selected Granola template exists before syncing a note. This uses private Granola APIs and is currently fork-specific.')
			.addToggle(toggle => {
				toggle.setValue(this.plugin.settings.enableGranolaTemplateManagement);
				toggle.onChange(async (value) => {
					this.plugin.settings.enableGranolaTemplateManagement = value;
					await this.plugin.saveSettings();
					this.display();
				});
			});

		if (this.plugin.settings.enableGranolaTemplateManagement) {
			new obsidian.Setting(containerEl)
				.setName('Granola template')
				.setDesc('Choose the Granola template the plugin should ensure exists before syncing. Template management only runs when this template is missing on a note.')
				.addDropdown(dropdown => {
					const templates = Array.isArray(this.plugin.availableGranolaTemplates)
						? [...this.plugin.availableGranolaTemplates].sort((a, b) => (a.title || '').localeCompare(b.title || ''))
						: [];

					dropdown.addOption('', this.plugin.settings.granolaTemplateTitle || 'Select a Granola template');
					for (const template of templates) {
						dropdown.addOption(template.id, template.title || template.id);
					}

					dropdown.setValue(this.plugin.settings.granolaTemplateId || '');
					dropdown.onChange(async (value) => {
						this.plugin.settings.granolaTemplateId = value;
						const selectedTemplate = templates.find((template) => template.id === value);
						this.plugin.settings.granolaTemplateTitle = selectedTemplate ? (selectedTemplate.title || '') : '';
						await this.plugin.saveSettings();
					});
				})
				.addButton(button => {
					button.setButtonText('Refresh templates');
					button.onClick(async () => {
						try {
							const authContext = await this.plugin.loadCredentials();
							if (!authContext) {
								new obsidian.Notice('Could not load Granola credentials. Please check your auth key path.');
								return;
							}

							const templates = await this.plugin.fetchGranolaTemplates(authContext, true);
							if (templates && templates.length > 0) {
								new obsidian.Notice(`Loaded ${templates.length} Granola templates`);
							} else {
								new obsidian.Notice('No Granola templates were returned');
							}
							this.display();
						} catch (error) {
							console.error('Error refreshing Granola templates:', error);
							new obsidian.Notice('Error refreshing Granola templates. Check console for details.');
						}
					});
				});

			if (!Array.isArray(this.plugin.availableGranolaTemplates) || this.plugin.availableGranolaTemplates.length === 0) {
				const templateInfoEl = containerEl.createEl('div', { cls: 'setting-item' });
				templateInfoEl.createEl('div', { cls: 'setting-item-info' });
				const templateInfoNameEl = templateInfoEl.createEl('div', { cls: 'setting-item-name' });
				templateInfoNameEl.setText('Granola templates not loaded yet');
				const templateInfoDescEl = templateInfoEl.createEl('div', { cls: 'setting-item-description' });
				templateInfoDescEl.setText('Use "Refresh templates" to load the live template list from Granola before selecting one.');
			}
		}

		new obsidian.Setting(containerEl)
			.setName('Search scope for existing notes')
			.setDesc('Choose where to search for existing notes when checking granola-id. "Sync directory only" (default) only checks the configured sync folder. "Entire vault" allows you to move notes anywhere in your vault. "Specific folders" lets you choose which folders to search.')
			.addDropdown(dropdown => {
				dropdown.addOption('syncDirectory', 'Sync directory only (default)');
				dropdown.addOption('entireVault', 'Entire vault');
				dropdown.addOption('specificFolders', 'Specific folders');
				
				dropdown.setValue(this.plugin.settings.existingNoteSearchScope);
				dropdown.onChange(async (value) => {
					const oldValue = this.plugin.settings.existingNoteSearchScope;
					this.plugin.settings.existingNoteSearchScope = value;
					
					// Save settings without triggering auto-sync to prevent duplicates
					await this.plugin.saveSettingsWithoutSync();
					
					// Show warning if search scope changed
					if (oldValue !== value) {
						new obsidian.Notice('Search scope changed. Consider running a manual sync to test the new settings before relying on auto-sync.');
					}
					
					this.display(); // Refresh the settings display
				});
			});

		// Show folder selection only when 'specificFolders' is selected
		if (this.plugin.settings.existingNoteSearchScope === 'specificFolders') {
			new obsidian.Setting(containerEl)
				.setName('Specific search folders')
				.setDesc('Enter folder paths to search (one per line). Leave empty to search all folders.')
				.addTextArea(text => {
					text.setPlaceholder('Examples:\nMeetings\nProjects/Work\nDaily Notes');
					text.setValue(this.plugin.settings.specificSearchFolders.join('\n'));
					
					// Save settings immediately on change (without validation and without auto-sync)
					text.onChange(async (value) => {
						const folders = value.split('\n').map(f => f.trim()).filter(f => f.length > 0);
						this.plugin.settings.specificSearchFolders = folders;
						await this.plugin.saveSettingsWithoutSync();
					});
					
					// Validate only when user finishes editing (on blur)
					text.inputEl.addEventListener('blur', () => {
						const value = text.getValue();
						const folders = value.split('\n').map(f => f.trim()).filter(f => f.length > 0);
						
						if (folders.length === 0) {
							return; // Don't validate if no folders specified
						}
						
						// Validate folder paths
						const invalidFolders = [];
						for (const folderPath of folders) {
							const folder = this.app.vault.getFolderByPath(folderPath);
							if (!folder) {
								invalidFolders.push(folderPath);
							}
						}
						
						if (invalidFolders.length > 0) {
							new obsidian.Notice('Warning: These folders do not exist: ' + invalidFolders.join(', '));
						}
					});
				});
		}

		// Add info section about avoiding duplicates
		const infoEl = containerEl.createEl('div', { cls: 'setting-item' });
		infoEl.createEl('div', { cls: 'setting-item-info' });
		const infoNameEl = infoEl.createEl('div', { cls: 'setting-item-name' });
		infoNameEl.setText('⚠️ Avoiding duplicates');
		const infoDescEl = infoEl.createEl('div', { cls: 'setting-item-description' });
		infoDescEl.setText('When changing search scope, existing notes in other locations won\'t be found and may be recreated. To avoid duplicates: 1) Move your existing notes to the new search location first, or 2) Use "Entire Vault" to search everywhere, or 3) Run a manual sync after changing settings to test before auto-sync runs.');

		// Create a heading for metadata settings
		containerEl.createEl('h3', {text: 'Note metadata & tags'});

		// Frontmatter customization settings
		new obsidian.Setting(containerEl)
			.setName('Include title in frontmatter')
			.setDesc('Add the meeting title as a "title" field in frontmatter. Disable if you want to use the filename as the single source of truth.')
			.addToggle(toggle => {
				toggle.setValue(this.plugin.settings.includeTitle);
				toggle.onChange(async (value) => {
					this.plugin.settings.includeTitle = value;
					await this.plugin.saveSettings();
				});
			});

		new obsidian.Setting(containerEl)
			.setName('Include dates in frontmatter')
			.setDesc('Add created_at and updated_at fields to frontmatter')
			.addToggle(toggle => {
				toggle.setValue(this.plugin.settings.includeDates);
				toggle.onChange(async (value) => {
					this.plugin.settings.includeDates = value;
					await this.plugin.saveSettings();
				});
			});

		new obsidian.Setting(containerEl)
			.setName('Frontmatter date format')
			.setDesc('Choose how dates are formatted in frontmatter')
			.addDropdown(dropdown => {
				dropdown
					.addOption('iso', 'ISO 8601 (2026-02-05T22:30:00.000Z)')
					.addOption('date-only', 'Date only (2026-02-05)')
					.addOption('custom', 'Custom format')
					.setValue(this.plugin.settings.frontmatterDateFormat)
					.onChange(async (value) => {
						this.plugin.settings.frontmatterDateFormat = value;
						await this.plugin.saveSettings();
						// Re-render to show/hide custom format field
						this.display();
					});
			});

		// Show custom date format field only when 'custom' is selected
		if (this.plugin.settings.frontmatterDateFormat === 'custom') {
			new obsidian.Setting(containerEl)
				.setName('Custom date format')
				.setDesc('Use YYYY for year, MM for month, DD for day, HH for hour, mm for minute, ss for second. Example: YYYY-MM-DD')
				.addText(text => {
					text.setPlaceholder('YYYY-MM-DD');
					text.setValue(this.plugin.settings.customDateFormat);
					text.onChange(async (value) => {
						this.plugin.settings.customDateFormat = value;
						await this.plugin.saveSettings();
					});
				});
		}

		new obsidian.Setting(containerEl)
			.setName('Additional frontmatter')
			.setDesc('Add custom fields to frontmatter. One per line in "key: value" format. Example: "type: meeting" or "status: draft"')
			.addTextArea(textArea => {
				textArea.setPlaceholder('type: meeting\nstatus: draft');
				textArea.setValue(this.plugin.settings.additionalFrontmatter);
				textArea.inputEl.rows = 4;
				textArea.inputEl.cols = 30;
				textArea.onChange(async (value) => {
					this.plugin.settings.additionalFrontmatter = value;
					await this.plugin.saveSettings();
				});
			});

		new obsidian.Setting(containerEl)
			.setName('Map metadata to frontmatter')
			.setDesc('Extract Granola’s inline metadata block and map fields like org, people, topics, loc, and meeting_type into frontmatter.')
			.addToggle(toggle => {
				toggle.setValue(this.plugin.settings.mapMetadataToFrontmatter);
				toggle.onChange(async (value) => {
					this.plugin.settings.mapMetadataToFrontmatter = value;
					await this.plugin.saveSettings();
					this.display();
				});
			});

		new obsidian.Setting(containerEl)
			.setName('Remove metadata section from body')
			.setDesc('When metadata is mapped into frontmatter, remove the inline "### Metadata" block from the note body.')
			.addToggle(toggle => {
				toggle.setValue(this.plugin.settings.removeMetadataSectionFromBody);
				toggle.setDisabled(!this.plugin.settings.mapMetadataToFrontmatter);
				toggle.onChange(async (value) => {
					this.plugin.settings.removeMetadataSectionFromBody = value;
					await this.plugin.saveSettings();
				});
			});

		new obsidian.Setting(containerEl)
			.setName('Mapped org template')
			.setDesc('Customize mapped org values. Use {name} for the original value and {slug} for a lowercase hyphenated version. Example: "[[Organizations/{name}]]"')
			.addText(text => {
				text.setPlaceholder('{name}');
				text.setValue(this.plugin.settings.metadataOrgTemplate);
				text.setDisabled(!this.plugin.settings.mapMetadataToFrontmatter);
				text.onChange(async (value) => {
					if (value && !value.includes('{name}') && !value.includes('{value}') && !value.includes('{slug}')) {
						new obsidian.Notice('Warning: Org template should include {name}, {value}, or {slug}');
					}
					this.plugin.settings.metadataOrgTemplate = value || '{name}';
					await this.plugin.saveSettings();
				});
			});

		new obsidian.Setting(containerEl)
			.setName('Mapped people template')
			.setDesc('Customize each mapped people value. Use {name} for the original value and {slug} for a lowercase hyphenated version. Example: "[[People/{name}]]"')
			.addText(text => {
				text.setPlaceholder('{name}');
				text.setValue(this.plugin.settings.metadataPersonTemplate);
				text.setDisabled(!this.plugin.settings.mapMetadataToFrontmatter);
				text.onChange(async (value) => {
					if (value && !value.includes('{name}') && !value.includes('{value}') && !value.includes('{slug}')) {
						new obsidian.Notice('Warning: People template should include {name}, {value}, or {slug}');
					}
					this.plugin.settings.metadataPersonTemplate = value || '{name}';
					await this.plugin.saveSettings();
				});
			});

		new obsidian.Setting(containerEl)
			.setName('Include attendee tags')
			.setDesc('Add meeting attendees as tags in the frontmatter of each note')
			.addToggle(toggle => {
				toggle.setValue(this.plugin.settings.includeAttendeeTags);
				toggle.onChange(async (value) => {
					this.plugin.settings.includeAttendeeTags = value;
					await this.plugin.saveSettings();
				});
			});

		new obsidian.Setting(containerEl)
			.setName('Exclude my name from tags')
			.setDesc('When adding attendee tags, exclude your own name from the list')
			.addToggle(toggle => {
				toggle.setValue(this.plugin.settings.excludeMyNameFromTags);
				toggle.onChange(async (value) => {
					this.plugin.settings.excludeMyNameFromTags = value;
					await this.plugin.saveSettings();
				});
			});

		new obsidian.Setting(containerEl)
			.setName('My name')
			.setDesc('Your name as it appears in Granola meetings (used to exclude from attendee tags)')
			.addText(text => {
				text.setPlaceholder('Danny McClelland');
				text.setValue(this.plugin.settings.myName);
				text.onChange(async (value) => {
					this.plugin.settings.myName = value;
					await this.plugin.saveSettings();
				});
			});

		new obsidian.Setting(containerEl)
			.setName('Attendee tag template')
			.setDesc('Customize the structure of attendee tags. Use {name} as placeholder for the attendee name. Examples: "person/{name}", "people/{name}", "meeting-attendees/{name}"')
			.addText(text => {
				text.setPlaceholder('person/{name}');
				text.setValue(this.plugin.settings.attendeeTagTemplate);
				text.onChange(async (value) => {
					// Validate the template has {name} placeholder
					if (!value.includes('{name}')) {
						new obsidian.Notice('Warning: Tag template should include {name} placeholder');
					}
					this.plugin.settings.attendeeTagTemplate = value || 'person/{name}';
					await this.plugin.saveSettings();
				});
			});

		new obsidian.Setting(containerEl)
			.setName('Include folder tags')
			.setDesc('Add Granola folder names as tags in the frontmatter of each note (requires Granola folders to be enabled)')
			.addToggle(toggle => {
				toggle.setValue(this.plugin.settings.includeFolderTags);
				toggle.setDisabled(!this.plugin.settings.enableGranolaFolders);
				toggle.onChange(async (value) => {
					this.plugin.settings.includeFolderTags = value;
					await this.plugin.saveSettings();
				});
			});

		new obsidian.Setting(containerEl)
			.setName('Include Granola URL')
			.setDesc('Add a link back to the original Granola note in the frontmatter (e.g., granola_url: "https://notes.granola.ai/d/...")')
			.addToggle(toggle => {
				toggle.setValue(this.plugin.settings.includeGranolaUrl);
				toggle.onChange(async (value) => {
					this.plugin.settings.includeGranolaUrl = value;
					await this.plugin.saveSettings();
				});
			});

		// Create a heading for daily note integration
		containerEl.createEl('h3', {text: 'Daily note integration'});

		new obsidian.Setting(containerEl)
			.setName('Daily note integration')
			.setDesc('Add todays meetings to your daily note')
			.addToggle(toggle => {
				toggle.setValue(this.plugin.settings.enableDailyNoteIntegration);
				toggle.onChange(async (value) => {
					this.plugin.settings.enableDailyNoteIntegration = value;
					await this.plugin.saveSettings();
				});
			});

		new obsidian.Setting(containerEl)
			.setName('Daily note section name')
			.setDesc('The heading name for the Granola meetings section in your daily note')
			.addText(text => {
				text.setPlaceholder('## Granola Meetings');
				text.setValue(this.plugin.settings.dailyNoteSectionName);
				text.onChange(async (value) => {
					this.plugin.settings.dailyNoteSectionName = value;
					await this.plugin.saveSettings();
				});
			});

		// Create a heading for periodic note integration
		containerEl.createEl('h3', {text: 'Periodic note integration'});

		new obsidian.Setting(containerEl)
			.setName('Periodic note integration')
			.setDesc('Add todays meetings to your periodic notes (daily, weekly, or monthly - requires Periodic Notes plugin)')
			.addToggle(toggle => {
				toggle.setValue(this.plugin.settings.enablePeriodicNoteIntegration);
				toggle.onChange(async (value) => {
					this.plugin.settings.enablePeriodicNoteIntegration = value;
					await this.plugin.saveSettings();
				});
			});

		// Add a warning if Periodic Notes plugin is not available
		if (!this.plugin.isPeriodicNotesPluginAvailable()) {
			const warningEl = containerEl.createEl('div', { cls: 'setting-item' });
			warningEl.createEl('div', { cls: 'setting-item-info' });
			const warningNameEl = warningEl.createEl('div', { cls: 'setting-item-name' });
			warningNameEl.setText('⚠️ Periodic Notes plugin not detected');
			const warningDescEl = warningEl.createEl('div', { cls: 'setting-item-description' });
			warningDescEl.setText('The Periodic Notes plugin is not installed or enabled. This integration will not work until the plugin is installed.');
		}

		new obsidian.Setting(containerEl)
			.setName('Periodic note section name')
			.setDesc('The heading name for the Granola meetings section in your periodic notes (works with daily, weekly, or monthly notes)')
			.addText(text => {
				text.setPlaceholder('## Granola Meetings');
				text.setValue(this.plugin.settings.periodicNoteSectionName);
				text.onChange(async (value) => {
					this.plugin.settings.periodicNoteSectionName = value;
					await this.plugin.saveSettings();
				});
			});

		// Create a heading for Granola folders
		containerEl.createEl('h3', {text: 'Granola folders'});

		// Use a button-based approach for the folder toggle
		new obsidian.Setting(containerEl)
			.setName('Enable Granola folders')
			.setDesc('Organize notes into folders based on Granola folder structure. This will create subfolders in your sync directory for each Granola folder.')
			.addButton(button => {
				// Ensure the setting exists and has a default value
				if (this.plugin.settings.enableGranolaFolders === undefined) {
					this.plugin.settings.enableGranolaFolders = false;
				}
				
				const updateButton = () => {
					if (this.plugin.settings.enableGranolaFolders) {
						button.setButtonText('Disable Granola folders');
						button.setCta();
					} else {
						button.setButtonText('Enable Granola folders');
						button.setCta(false);
					}
				};
				
				updateButton();
				
				button.onClick(async () => {
					this.plugin.settings.enableGranolaFolders = !this.plugin.settings.enableGranolaFolders;
					await this.plugin.saveSettings();
					updateButton();
					this.display(); // Refresh the settings display
					new obsidian.Notice(`Granola folders ${this.plugin.settings.enableGranolaFolders ? 'enabled' : 'disabled'}`);
				});
			});

		new obsidian.Setting(containerEl)
			.setName('Folder tag template')
			.setDesc('Customize the structure of folder tags. Use {name} as placeholder for the folder name. Examples: "folder/{name}", "granola/{name}", "meeting-folders/{name}"')
			.addText(text => {
				text.setPlaceholder('folder/{name}');
				text.setValue(this.plugin.settings.folderTagTemplate);
				text.setDisabled(!this.plugin.settings.enableGranolaFolders);
				text.onChange(async (value) => {
					// Validate the template has {name} placeholder
					if (!value.includes('{name}')) {
						new obsidian.Notice('Warning: Tag template should include {name} placeholder');
					}
					this.plugin.settings.folderTagTemplate = value || 'folder/{name}';
					await this.plugin.saveSettings();
				});
			});

		// Folder filtering settings
		containerEl.createEl('h4', {text: 'Folder filtering'});

		new obsidian.Setting(containerEl)
			.setName('Enable folder filter')
			.setDesc('Only sync notes from selected Granola folders. When disabled, all notes are synced.')
			.addToggle(toggle => {
				toggle.setValue(this.plugin.settings.enableFolderFilter);
				toggle.onChange(async (value) => {
					this.plugin.settings.enableFolderFilter = value;
					await this.plugin.saveSettings();
					this.display(); // Refresh to show/hide folder selection
				});
			});

		// Show folder selection only when folder filtering is enabled
		if (this.plugin.settings.enableFolderFilter) {
			// Create a button to fetch available folders
			new obsidian.Setting(containerEl)
				.setName('Refresh folder list')
				.setDesc('Fetch the latest list of folders from Granola')
				.addButton(button => {
					button.setButtonText('Refresh folders');
					button.onClick(async () => {
						try {
							const authContext = await this.plugin.loadCredentials();
							if (!authContext) {
								new obsidian.Notice('Could not load credentials. Please check your auth key path.');
								return;
							}
							const folders = await this.plugin.fetchGranolaFolders(authContext);
							if (folders) {
								this.plugin.availableGranolaFolders = folders;
								new obsidian.Notice(`Found ${folders.length} folders`);
								this.display(); // Refresh to show updated folders
							} else {
								new obsidian.Notice('Could not fetch folders from Granola');
							}
						} catch (error) {
							console.error('Error fetching folders:', error);
							new obsidian.Notice('Error fetching folders. Check console for details.');
						}
					});
				});

			// Show available folders with checkboxes
			const availableFolders = this.plugin.availableGranolaFolders || [];
			if (availableFolders.length > 0) {
				const folderSelectionEl = containerEl.createEl('div', { cls: 'setting-item' });
				const folderInfoEl = folderSelectionEl.createEl('div', { cls: 'setting-item-info' });
				const folderNameEl = folderInfoEl.createEl('div', { cls: 'setting-item-name' });
				folderNameEl.setText('Select folders to sync');
				const folderDescEl = folderInfoEl.createEl('div', { cls: 'setting-item-description' });
				folderDescEl.setText('Check the folders you want to sync. Only notes in selected folders will be synced.');

				const folderListEl = containerEl.createEl('div', { cls: 'granola-folder-list' });
				folderListEl.style.marginLeft = '20px';
				folderListEl.style.marginBottom = '20px';

				for (const folder of availableFolders) {
					const folderItemEl = folderListEl.createEl('div', { cls: 'granola-folder-item' });
					folderItemEl.style.display = 'flex';
					folderItemEl.style.alignItems = 'center';
					folderItemEl.style.marginBottom = '8px';

					const checkbox = folderItemEl.createEl('input', { type: 'checkbox' });
					checkbox.checked = this.plugin.settings.selectedGranolaFolders.includes(folder.id);
					checkbox.style.marginRight = '8px';

					const label = folderItemEl.createEl('label');
					label.setText(folder.title + (folder.document_ids ? ` (${folder.document_ids.length} notes)` : ''));
					label.style.cursor = 'pointer';

					checkbox.addEventListener('change', async () => {
						if (checkbox.checked) {
							if (!this.plugin.settings.selectedGranolaFolders.includes(folder.id)) {
								this.plugin.settings.selectedGranolaFolders.push(folder.id);
							}
						} else {
							this.plugin.settings.selectedGranolaFolders = this.plugin.settings.selectedGranolaFolders.filter(id => id !== folder.id);
						}
						await this.plugin.saveSettings();
					});

					label.addEventListener('click', () => {
						checkbox.click();
					});
				}

				// Add "Select All" / "Deselect All" buttons
				const buttonContainer = containerEl.createEl('div');
				buttonContainer.style.marginBottom = '20px';

				const selectAllBtn = buttonContainer.createEl('button');
				selectAllBtn.setText('Select All');
				selectAllBtn.style.marginRight = '10px';
				selectAllBtn.addEventListener('click', async () => {
					this.plugin.settings.selectedGranolaFolders = availableFolders.map(f => f.id);
					await this.plugin.saveSettings();
					this.display();
				});

				const deselectAllBtn = buttonContainer.createEl('button');
				deselectAllBtn.setText('Deselect All');
				deselectAllBtn.addEventListener('click', async () => {
					this.plugin.settings.selectedGranolaFolders = [];
					await this.plugin.saveSettings();
					this.display();
				});
			} else {
				const noFoldersEl = containerEl.createEl('div', { cls: 'setting-item' });
				const noFoldersInfoEl = noFoldersEl.createEl('div', { cls: 'setting-item-info' });
				const noFoldersDescEl = noFoldersInfoEl.createEl('div', { cls: 'setting-item-description' });
				noFoldersDescEl.setText('No folders loaded. Click "Refresh folders" to fetch available folders from Granola, or run a sync first.');
			}
		}

		// Create a heading for file organization settings
		containerEl.createEl('h3', {text: 'File organization'});

		new obsidian.Setting(containerEl)
			.setName('Sync directory')
			.setDesc('Directory within your vault where Granola notes will be synced')
			.addText(text => {
				text.setPlaceholder('Granola');
				text.setValue(this.plugin.settings.syncDirectory);
				text.onChange(async (value) => {
					this.plugin.settings.syncDirectory = value;
					await this.plugin.saveSettings();
				});
			});

		new obsidian.Setting(containerEl)
			.setName('Manual sync')
			.setDesc('Click to manually sync your Granola notes')
			.addButton(button => {
				button.setButtonText('Sync now');
				button.setCta();
				button.onClick(async () => {
					await this.plugin.syncNotes();
				});
			});

		new obsidian.Setting(containerEl)
			.setName('Reorganize existing notes into folders')
			.setDesc('Move all existing Granola notes into the correct folders based on your current date-based or Granola folder settings. This will not create duplicates.')
			.addButton(button => {
				button.setButtonText('Reorganize notes');
				button.setCta();
				button.onClick(async () => {
					await this.plugin.reorganizeExistingNotes();
				});
			});

		new obsidian.Setting(containerEl)
			.setName('Auto-reorganize notes after sync')
			.setDesc('Automatically move existing Granola notes to their correct folders after each sync. Only works when Granola folders or date-based folders are enabled.')
			.addToggle(toggle => {
				toggle.setValue(this.plugin.settings.enableAutoReorganize);
				toggle.setDisabled(!this.plugin.settings.enableGranolaFolders && !this.plugin.settings.enableDateBasedFolders);
				toggle.onChange(async (value) => {
					this.plugin.settings.enableAutoReorganize = value;
					await this.plugin.saveSettings();
				});
			});

		new obsidian.Setting(containerEl)
			.setName('Find duplicate notes')
			.setDesc('Find and list notes with the same granola-id (helpful after changing search scope settings)')
			.addButton(button => {
				button.setButtonText('Find duplicates');
				button.onClick(async () => {
					await this.plugin.findDuplicateNotes();
				});
			});

		new obsidian.Setting(containerEl)
			.setName('Re-enable auto-sync')
			.setDesc('Re-enable auto-sync after testing new search scope settings (this will restart the auto-sync timer)')
			.addButton(button => {
				button.setButtonText('Re-enable auto-sync');
				button.onClick(async () => {
					await this.plugin.saveSettings(); // This will call setupAutoSync()
					new obsidian.Notice('Auto-sync re-enabled with current settings');
				});
			});

		new obsidian.Setting(containerEl)
			.setName('Reset integration settings')
			.setDesc('Reset Daily Notes and Periodic Notes integration to disabled (useful if toggles seem stuck)')
			.addButton(button => {
				button.setButtonText('Reset integrations');
				button.onClick(async () => {
					this.plugin.settings.enableDailyNoteIntegration = false;
					this.plugin.settings.enablePeriodicNoteIntegration = false;
					await this.plugin.saveSettings();
					this.display(); // Refresh the settings display
					new obsidian.Notice('Integration settings reset to disabled');
				});
			});

		new obsidian.Setting(containerEl)
			.setName('Reset folder settings')
			.setDesc('Reset Granola folder settings to default values (useful if toggles seem stuck)')
			.addButton(button => {
				button.setButtonText('Reset folder settings');
				button.onClick(async () => {
					this.plugin.settings.enableGranolaFolders = false;
					this.plugin.settings.includeFolderTags = false;
					this.plugin.settings.folderTagTemplate = 'folder/{name}';
					await this.plugin.saveSettings();
					this.display(); // Refresh the settings display
					new obsidian.Notice('Folder settings reset to defaults');
				});
			});
	}
}

module.exports = GranolaSyncPlugin;
