// Copyright (c) 2026 Malte Doerper. MIT License. See LICENSE file.
// Server-mode controls for the Data Visualization Grid.
// When the grid is served by the FastAPI server, this module provides
// webhook-based controls that replace the file-system load/save buttons.
// Pipeline definitions are fetched from /api/pipelines/ and drive the
// available actions (Convert Excel, Review → Excel, etc.).

class ServerControlsManager {
    constructor(app) {
        this.app = app;
        this.container = document.querySelector('.server-controls');
        this.debugTag = '[SERVER FILE ACTION]';
        this.storageKeys = {
            fieldPrefix: 'grid-server-field:',
            datasetFilename: 'grid-server-current-dataset'
        };
        this.currentDataset = null;
        this.sourceInfo = null;
        this.pipelineConfig = [];
        this.fieldConfig = [];
        this.fieldState = new Map();
        this.activeTabId = null;
        this.datasetLoadSequence = 0;
        this.latestDatasetLoadRequestId = 0;
        this.initializeEventListeners();
        this.fetchPipelineConfig();
        void this.loadServerCurrentUserName();
        void this.restorePersistedDataset();
    }

    debugLog(message, details) {
        if (typeof details === 'undefined') {
            console.debug(this.debugTag, message);
            return;
        }

        console.debug(this.debugTag, message, details);
    }

    createTimingContext(label, details = {}) {
        return {
            label,
            details,
            startedAt: performance.now(),
        };
    }

    logTiming(timingContext, message, details = {}) {
        if (!timingContext) {
            return;
        }

        this.debugLog(message, {
            ...timingContext.details,
            ...details,
            durationMs: Math.round((performance.now() - timingContext.startedAt) * 10) / 10,
        });
    }

    getCurrentFieldPayload() {
        const payload = {};
        for (const field of this.fieldConfig) {
            const value = this.getFieldValue(field.valueKey);
            if (!value) {
                continue;
            }
            payload[field.valueKey] = value;
        }
        return payload;
    }

    getPipelineRequestPayload(pipeline, options = {}) {
        const payload = {
            pipelineId: pipeline ? pipeline.id : null,
            pipelineLabel: pipeline ? pipeline.label : null,
            activeTabId: this.activeTabId,
            currentDataset: this.currentDataset ? this.currentDataset.filename : null,
            fields: this.getCurrentFieldPayload(),
        };

        if (options.uploadFile) {
            payload.uploadFile = {
                name: options.uploadFile.name,
                size: options.uploadFile.size,
                type: options.uploadFile.type || null,
            };
        }

        if (options.uploadField) {
            payload.uploadField = options.uploadField;
        }

        return payload;
    }

    // ------------------------------------------------------------------
    // Event listeners
    // ------------------------------------------------------------------

    initializeEventListeners() {
        if (this.container) {
            this.container.addEventListener('click', (event) => {
                const tabButton = event.target.closest('[data-server-tab-id]');
                if (tabButton) {
                    const { serverTabId } = tabButton.dataset;
                    if (serverTabId) {
                        this.debugLog('user clicked server tab', {
                            tabId: serverTabId,
                            previousTabId: this.activeTabId,
                        });
                        this.setActiveTab(serverTabId);
                    }
                    return;
                }

                const button = event.target.closest('[data-server-pipeline-id]');
                if (!button) {
                    return;
                }
                const { serverPipelineId: pipelineId } = button.dataset;
                if (pipelineId) {
                    this.debugLog('user clicked server pipeline action', {
                        pipelineId,
                        activeTabId: this.activeTabId,
                        currentDataset: this.currentDataset ? this.currentDataset.filename : null,
                        fields: this.getCurrentFieldPayload(),
                    });
                    void this.startPipeline(pipelineId);
                }
            });

            this.container.addEventListener('keydown', (event) => {
                const input = event.target.closest('[data-server-field-key]');
                if (!input || event.key !== 'Enter') {
                    return;
                }
                event.preventDefault();
                const fieldConfig = this.fieldConfig.find((field) => field.valueKey === input.dataset.serverFieldKey);
                if (fieldConfig && fieldConfig.primaryPipelineId) {
                    this.debugLog('user pressed Enter in server field', {
                        fieldKey: input.dataset.serverFieldKey,
                        value: input.value,
                        pipelineId: fieldConfig.primaryPipelineId,
                    });
                    void this.startPipeline(fieldConfig.primaryPipelineId);
                }
            });

            this.container.addEventListener('input', (event) => {
                const input = event.target.closest('[data-server-field-key]');
                if (!input) {
                    return;
                }
                this.setFieldValue(input.dataset.serverFieldKey, input.value, { persist: false, syncDom: false });
            });

            this.container.addEventListener('change', (event) => {
                const input = event.target.closest('[data-server-field-key]');
                if (!input) {
                    return;
                }
                this.setFieldValue(input.dataset.serverFieldKey, input.value);
            });

            this.container.addEventListener('blur', (event) => {
                const input = event.target.closest('[data-server-field-key]');
                if (!input) {
                    return;
                }
                this.setFieldValue(input.dataset.serverFieldKey, input.value);
            }, true);
        }

        document.getElementById('console-popup-close')
            .addEventListener('click', () => this.closeConsolePopup());
        document.getElementById('console-popup-yes')
            .addEventListener('click', () => this.sendConsoleInput('y'));
        document.getElementById('console-popup-no')
            .addEventListener('click', () => this.sendConsoleInput('n'));
    }

    getStoredValue(key) {
        try {
            return localStorage.getItem(key);
        } catch (err) {
            console.warn('[SERVER] Failed to read local storage:', err);
            return null;
        }
    }

    setStoredValue(key, value) {
        try {
            const normalizedValue = String(value || '').trim();
            if (normalizedValue) {
                localStorage.setItem(key, normalizedValue);
            } else {
                localStorage.removeItem(key);
            }
        } catch (err) {
            console.warn('[SERVER] Failed to persist local storage:', err);
        }
    }

    async loadServerCurrentUserName() {
        try {
            const response = await fetch('/api/settings/current-user', { method: 'GET' });
            if (!response.ok) {
                throw new Error(`HTTP ${response.status}`);
            }

            const payload = await response.json();
            const currentUserName = String(payload.currentUserName || '').trim();
            if (!currentUserName) {
                return;
            }

            this.app.setCurrentUserName(currentUserName, {
                persistLocal: true,
                persistServer: false,
            });
        } catch (error) {
            console.warn('[SERVER] Failed to load current user setting:', error);
        }
    }

    async persistCurrentUserName(name) {
        const currentUserName = String(name || '').trim();
        if (!currentUserName) {
            return false;
        }

        try {
            const response = await fetch('/api/settings/current-user', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({ currentUserName }),
            });

            if (!response.ok) {
                const detail = await response.json().catch(() => ({}));
                throw new Error(detail.detail || `HTTP ${response.status}`);
            }

            return true;
        } catch (error) {
            console.warn('[SERVER] Failed to persist current user setting:', error);
            return false;
        }
    }

    getFieldStorageKey(valueKey) {
        return `${this.storageKeys.fieldPrefix}${valueKey}`;
    }

    getPersistedFieldValue(valueKey) {
        return this.getStoredValue(this.getFieldStorageKey(valueKey)) || '';
    }

    setFieldValue(valueKey, value, options = {}) {
        const {
            persist = true,
            syncDom = true,
        } = options;

        if (!valueKey) {
            return;
        }

        const normalizedValue = String(value || '').trim();
        this.fieldState.set(valueKey, normalizedValue);

        if (persist) {
            this.setStoredValue(this.getFieldStorageKey(valueKey), normalizedValue);
        }

        if (syncDom && this.container) {
            const inputs = this.container.querySelectorAll(`[data-server-field-key="${CSS.escape(valueKey)}"]`);
            for (const input of inputs) {
                if (input.value !== normalizedValue) {
                    input.value = normalizedValue;
                }
            }
        }

        this.updateActionState();
    }

    getFieldValue(valueKey) {
        if (!valueKey) {
            return '';
        }
        if (this.fieldState.has(valueKey)) {
            return this.fieldState.get(valueKey) || '';
        }
        const persistedValue = this.getPersistedFieldValue(valueKey);
        if (persistedValue) {
            this.fieldState.set(valueKey, persistedValue);
        }
        return persistedValue;
    }

    async restorePersistedDataset() {
        const datasetFilename = this.getStoredValue(this.storageKeys.datasetFilename);
        if (!datasetFilename) {
            return;
        }

        this.debugLog('restoring persisted server dataset', {
            datasetFilename,
        });

        await this.loadServerDataset(datasetFilename, {
            suppressErrorNotification: true,
            clearPersistedOnFailure: true,
            reason: 'restore-persisted-dataset',
        });
    }

    // ------------------------------------------------------------------
    // Pipeline config
    // ------------------------------------------------------------------

    async fetchPipelineConfig() {
        const timing = this.createTimingContext('pipeline config fetch');
        try {
            const resp = await fetch('/api/pipelines/');
            if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
            const pipelines = await resp.json();
            this.pipelineConfig = Array.isArray(pipelines)
                ? [...pipelines].sort((left, right) => (left.displayOrder || 0) - (right.displayOrder || 0))
                : [];
            this.logTiming(timing, 'fetched server pipeline config', {
                pipelineCount: this.pipelineConfig.length,
            });
            this.renderPipelineControls();
        } catch (err) {
            console.warn('[SERVER] Failed to fetch pipeline config:', err);
        }
    }

    buildFieldConfig() {
        const fieldsByKey = new Map();
        for (const pipeline of this.pipelineConfig) {
            const tabId = this.normalizeTabId(pipeline.tab || 'General');
            for (const input of pipeline.inputs || []) {
                const valueKey = input.valueKey || input.submissionField || input.name;
                if (!valueKey) {
                    continue;
                }
                const existing = fieldsByKey.get(valueKey);
                if (existing) {
                    if (!existing.tabIds.includes(tabId)) {
                        existing.tabIds.push(tabId);
                    }
                    if (!existing.sourceInfoKey && input.sourceInfoKey) {
                        existing.sourceInfoKey = input.sourceInfoKey;
                    }
                    if (!existing.placeholder && input.placeholder) {
                        existing.placeholder = input.placeholder;
                    }
                    if (!existing.label && input.label) {
                        existing.label = input.label;
                    }
                    if ((existing.default === undefined || existing.default === null || existing.default === '') && input.default) {
                        existing.default = input.default;
                    }
                    continue;
                }
                fieldsByKey.set(valueKey, {
                    ...input,
                    valueKey,
                    primaryPipelineId: pipeline.id,
                    tabIds: [tabId],
                });
            }
        }
        this.fieldConfig = [...fieldsByKey.values()];
    }

    buildTabs() {
        const tabs = [];
        const seen = new Set();
        for (const pipeline of this.pipelineConfig) {
            const id = this.normalizeTabId(pipeline.tab || 'General');
            if (seen.has(id)) {
                continue;
            }
            tabs.push({ id, label: pipeline.tab || 'General' });
            seen.add(id);
        }
        if (!tabs.length) {
            tabs.push({ id: 'general', label: 'General' });
        }
        if (!this.activeTabId || !tabs.some((tab) => tab.id === this.activeTabId)) {
            this.activeTabId = tabs[0].id;
        }
        return tabs;
    }

    buildGroupsForTab(tabId) {
        const groups = [];
        const groupMap = new Map();

        for (const pipeline of this.pipelineConfig) {
            if (this.normalizeTabId(pipeline.tab || 'General') !== tabId) {
                continue;
            }
            const groupLabel = pipeline.group || 'Actions';
            if (!groupMap.has(groupLabel)) {
                const group = { label: groupLabel, pipelines: [] };
                groups.push(group);
                groupMap.set(groupLabel, group);
            }
            groupMap.get(groupLabel).pipelines.push(pipeline);
        }

        return groups;
    }

    createFieldElement(field) {
        const wrapper = document.createElement('div');
        wrapper.className = 'server-control-field';

        const label = document.createElement('label');
        label.className = 'server-control-label';
        label.textContent = field.label || field.name;

        const input = document.createElement('input');
        input.type = 'text';
        input.className = 'server-path-input';
        input.placeholder = field.placeholder || '';
        input.title = field.label || field.placeholder || '';
        input.dataset.serverFieldKey = field.valueKey;
        input.value = this.getFieldValue(field.valueKey) || field.default || '';

        wrapper.appendChild(label);
        wrapper.appendChild(input);
        return wrapper;
    }

    normalizeTabId(tabLabel) {
        return String(tabLabel || 'General')
            .trim()
            .toLowerCase()
            .replace(/[^a-z0-9]+/g, '-')
            .replace(/^-+|-+$/g, '') || 'general';
    }

    setActiveTab(tabId) {
        this.activeTabId = tabId;
        if (!this.container) {
            return;
        }

        for (const button of this.container.querySelectorAll('[data-server-tab-id]')) {
            button.classList.toggle('is-active', button.dataset.serverTabId === tabId);
        }
        for (const panel of this.container.querySelectorAll('[data-server-tab-panel-id]')) {
            panel.hidden = panel.dataset.serverTabPanelId !== tabId;
        }
    }

    getResolvedViewMode(preferredTabId = null) {
        const tabId = String(preferredTabId || this.activeTabId || '').trim().toLowerCase();
        if (tabId === 'excel' || tabId === 'obsidian') {
            return tabId;
        }

        if (this.sourceInfo && this.sourceInfo.source_excel) {
            return 'excel';
        }

        const obsidianConfigPath = this.getFieldValue('obsidian_config_path');
        if (obsidianConfigPath) {
            return 'obsidian';
        }

        const workbookPath = this.getFieldValue('workbook_path');
        if (workbookPath) {
            return 'excel';
        }

        return null;
    }

    buildResolvedViewRequest(preferredTabId = null) {
        const mode = this.getResolvedViewMode(preferredTabId);
        if (!mode) {
            throw new Error('No server view mode is active');
        }

        const request = {
            mode,
            dataset: this.currentDataset ? this.currentDataset.filename : null,
            workbookPath: this.getFieldValue('workbook_path') || null,
            obsidianConfigPath: this.getFieldValue('obsidian_config_path') || null,
        };

        console.debug('[SERVER VIEW]', 'resolved server view request', request);

        return request;
    }

    async fetchResolvedViewConfiguration(preferredTabId = null) {
        const request = this.buildResolvedViewRequest(preferredTabId);
        this.debugLog('requesting resolved server view configuration', request);
        const params = new URLSearchParams();
        params.set('mode', request.mode);
        if (request.dataset) {
            params.set('dataset', request.dataset);
        }
        if (request.workbookPath) {
            params.set('workbook_path', request.workbookPath);
        }
        if (request.obsidianConfigPath) {
            params.set('obsidian_config_path', request.obsidianConfigPath);
        }

        const response = await fetch(`/api/view-config/server?${params.toString()}`, { method: 'GET' });
        if (!response.ok) {
            const detail = await response.json().catch(() => ({}));
            throw new Error(detail.detail || `HTTP ${response.status}`);
        }

        return await response.json();
    }

    async tryAutoLoadResolvedViewConfiguration() {
        if (!this.app.persistenceManager) {
            return false;
        }

        try {
            const payload = await this.fetchResolvedViewConfiguration();
            if (!payload || !payload.content) {
                return false;
            }

            console.debug('[SERVER VIEW]', 'auto-loading server view configuration', {
                filename: payload.filename,
                directory: payload.directory,
            });
            this.app.persistenceManager.processViewFileContent(payload.content, payload.filename);
            return true;
        } catch (error) {
            if (!/View file not found/i.test(error.message)) {
                console.warn('[SERVER] Failed to auto-load resolved view config:', error);
            }
            return false;
        }
    }

    async openResolvedViewConfiguration() {
        if (!this.app.persistenceManager) {
            return;
        }

        try {
            this.debugLog('user triggered server view load', this.buildResolvedViewRequest());
            const payload = await this.fetchResolvedViewConfiguration();
            console.debug('[SERVER VIEW]', 'loading server view configuration via button', {
                filename: payload.filename,
                directory: payload.directory,
            });
            this.app.persistenceManager.processViewFileContent(payload.content, payload.filename);
        } catch (error) {
            this.app.showNotification(`Failed to load view config: ${error.message}`, 'error');
        }
    }

    async saveResolvedViewConfiguration() {
        try {
            const request = this.buildResolvedViewRequest();
            this.debugLog('user triggered server view save', request);
            const content = this.app.persistenceManager
                ? this.app.persistenceManager.prepareViewConfigForSaving()
                : (this.app.getViewConfig ? this.app.getViewConfig() : {});

            const response = await fetch('/api/view-config/server', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    ...request,
                    content,
                }),
            });

            if (!response.ok) {
                const detail = await response.json().catch(() => ({}));
                throw new Error(detail.detail || `HTTP ${response.status}`);
            }

            const payload = await response.json();
            console.debug('[SERVER VIEW]', 'saved server view configuration', payload);
            this.app.showNotification(`Saved view config: ${payload.filename}`, 'success');
        } catch (error) {
            this.app.showNotification(`Failed to save view config: ${error.message}`, 'error');
        }
    }

    renderPipelineControls() {
        if (!this.container) {
            return;
        }

        this.buildFieldConfig();
        this.container.replaceChildren();

        const tabs = this.buildTabs();
        const fragment = document.createDocumentFragment();

        if (tabs.length > 1) {
            const tabsRow = document.createElement('div');
            tabsRow.className = 'server-tabs';
            for (const tab of tabs) {
                const tabButton = document.createElement('button');
                tabButton.type = 'button';
                tabButton.className = 'server-tab-btn';
                tabButton.dataset.serverTabId = tab.id;
                tabButton.textContent = tab.label;
                if (tab.id === this.activeTabId) {
                    tabButton.classList.add('is-active');
                }
                tabsRow.appendChild(tabButton);
            }
            fragment.appendChild(tabsRow);
        }

        const panels = document.createElement('div');
        panels.className = 'server-tab-panels';

        for (const tab of tabs) {
            const panel = document.createElement('section');
            panel.className = 'server-tab-panel';
            panel.dataset.serverTabPanelId = tab.id;
            panel.hidden = tab.id !== this.activeTabId;

            const fieldsRow = document.createElement('div');
            fieldsRow.className = 'server-fields-row';
            for (const field of this.fieldConfig.filter((item) => item.tabIds.includes(tab.id))) {
                fieldsRow.appendChild(this.createFieldElement(field));
            }
            if (fieldsRow.childElementCount > 0) {
                panel.appendChild(fieldsRow);
            }

            const groups = this.buildGroupsForTab(tab.id);
            const actionRow = document.createElement('div');
            actionRow.className = 'server-actions-row';
            for (const group of groups) {
                for (const pipeline of group.pipelines) {
                    const button = document.createElement('button');
                    button.type = 'button';
                    button.className = 'control-btn';
                    button.dataset.serverPipelineId = pipeline.id;
                    button.textContent = pipeline.label;
                    button.title = pipeline.description || pipeline.label;
                    actionRow.appendChild(button);
                }
            }

            if (actionRow.childElementCount > 0) {
                panel.appendChild(actionRow);
            }

            panels.appendChild(panel);
        }

        fragment.appendChild(panels);

        this.container.appendChild(fragment);
        this.syncFieldsFromSourceInfo();
        this.updateActionState();
    }

    // ------------------------------------------------------------------
    // Source info tracking
    // ------------------------------------------------------------------

    async fetchSourceInfo(filename, options = {}) {
        const { applyState = true } = options;
        const timing = this.createTimingContext('source info fetch', {
            filename,
            applyState,
        });
        try {
            const resp = await fetch(`/api/data/source-info/${encodeURIComponent(filename)}`);
            if (resp.ok) {
                const info = await resp.json();
                this.logTiming(timing, 'fetched server source info', {
                    filename,
                    hasSourceInfo: true,
                });
                if (applyState) {
                    this.sourceInfo = info;
                    this.syncFieldsFromSourceInfo();
                    this.updateActionState();
                }
                return info;
            }
        } catch (_ignored) { /* no source info available */ }
        this.logTiming(timing, 'server source info unavailable', {
            filename,
            hasSourceInfo: false,
        });
        if (applyState) {
            this.sourceInfo = null;
            this.syncFieldsFromSourceInfo();
            this.updateActionState();
        }
        return null;
    }

    syncFieldsFromSourceInfo() {
        if (!this.sourceInfo) {
            return;
        }
        for (const field of this.fieldConfig) {
            if (!field.sourceInfoKey) {
                continue;
            }
            const sourceValue = this.sourceInfo[field.sourceInfoKey];
            if (typeof sourceValue === 'string' && sourceValue.trim()) {
                this.setFieldValue(field.valueKey, sourceValue);
            }
        }
    }

    // ------------------------------------------------------------------
    // Load a dataset from the server
    // ------------------------------------------------------------------

    async loadServerDataset(filename, options = {}) {
        const {
            suppressErrorNotification = false,
            clearPersistedOnFailure = false,
            reason = 'direct-load',
        } = options;
        const requestId = ++this.datasetLoadSequence;
        this.latestDatasetLoadRequestId = requestId;

        this.debugLog('starting server dataset load', {
            requestId,
            filename,
            reason,
            currentDataset: this.currentDataset ? this.currentDataset.filename : null,
        });

        try {
            const totalTiming = this.createTimingContext('dataset load total', {
                requestId,
                filename,
                reason,
            });
            const pendingSourceInfo = await this.fetchSourceInfo(filename, { applyState: false });

            const fetchTiming = this.createTimingContext('dataset fetch', {
                requestId,
                filename,
            });
            const resp = await fetch(`/api/data/${encodeURIComponent(filename)}`);
            if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
            const content = await resp.json();
            this.logTiming(fetchTiming, 'fetched server dataset payload', {
                requestId,
                filename,
                rootKeys: content && typeof content === 'object' && !Array.isArray(content) ? Object.keys(content) : [],
            });

            if (requestId !== this.latestDatasetLoadRequestId) {
                this.debugLog('skipping stale server dataset load response', {
                    requestId,
                    latestRequestId: this.latestDatasetLoadRequestId,
                    filename,
                    reason,
                });
                return { ok: false, skipped: true };
            }

            this.currentDataset = { filename };
            this.setStoredValue(this.storageKeys.datasetFilename, filename);

            // Reuse the persistence manager's data processing if available,
            // otherwise feed directly into app.loadDataFromJSON.
            if (this.app.persistenceManager) {
                if (this.app.persistenceManager.fileService) {
                    this.app.persistenceManager.fileService.rememberOpenedFile('data', filename);
                }
                const processTiming = this.createTimingContext('dataset processing', {
                    requestId,
                    filename,
                });
                await this.app.persistenceManager.processDataFileContent(content, filename, {
                    serverSourceInfo: pendingSourceInfo,
                });
                this.logTiming(processTiming, 'processed dataset payload in browser', {
                    requestId,
                    filename,
                });
            } else {
                this.app.loadDataFromJSON(content);
                this.app.showNotification(`Loaded ${filename} from server`, 'success');
            }

            // Refresh source info after load to keep field state in sync.
            if (requestId !== this.latestDatasetLoadRequestId) {
                this.debugLog('skipping stale server dataset post-processing', {
                    requestId,
                    latestRequestId: this.latestDatasetLoadRequestId,
                    filename,
                });
                return { ok: false, skipped: true };
            }

            const postSourceTiming = this.createTimingContext('source info apply', {
                requestId,
                filename,
            });
            await this.fetchSourceInfo(filename, { applyState: true });
            this.logTiming(postSourceTiming, 'applied server source info to UI state', {
                requestId,
                filename,
            });

            if (requestId === this.latestDatasetLoadRequestId) {
                const viewTiming = this.createTimingContext('resolved view auto-load', {
                    requestId,
                    filename,
                });
                await this.tryAutoLoadResolvedViewConfiguration();
                this.logTiming(viewTiming, 'completed resolved view auto-load', {
                    requestId,
                    filename,
                });
            }

            this.updateActionState();
            this.logTiming(totalTiming, 'completed server dataset load', {
                requestId,
                filename,
                sourceInfoKind: this.sourceInfo && this.sourceInfo.source_kind ? this.sourceInfo.source_kind : null,
            });
            this.debugLog('completed server dataset load', {
                requestId,
                filename,
                reason,
                sourceInfo: this.sourceInfo,
            });
            return { ok: true, filename };
        } catch (err) {
            console.error('[SERVER] Failed to load dataset:', err);
            if (clearPersistedOnFailure && requestId === this.latestDatasetLoadRequestId) {
                this.setStoredValue(this.storageKeys.datasetFilename, '');
            }
            if (!suppressErrorNotification && requestId === this.latestDatasetLoadRequestId) {
                this.app.showNotification(`Failed to load ${filename}: ${err.message}`, 'error');
            }
            this.debugLog('server dataset load failed', {
                requestId,
                filename,
                reason,
                error: err.message,
            });
            return { ok: false, error: err };
        }
    }

    // ------------------------------------------------------------------
    // Pipeline execution via server
    // ------------------------------------------------------------------

    async startPipeline(pipelineId) {
        const pipeline = this.getPipeline(pipelineId);
        if (!pipeline) {
            this.app.showNotification(`Pipeline not configured: ${pipelineId}`, 'error');
            return;
        }

        this.debugLog('starting server pipeline', this.getPipelineRequestPayload(pipeline));

        if (pipeline.execution && pipeline.execution.mode === 'interactive') {
            await this.runInteractivePipeline(pipeline);
            return;
        }

        const uploadInput = (pipeline.inputs || []).find((input) => input.source === 'upload');
        if (uploadInput) {
            this.openUploadPicker(pipeline, uploadInput);
            return;
        }

        await this.runRequestPipeline(pipeline);
    }

    openUploadPicker(pipeline, uploadInput) {
        const accept = uploadInput.accept || pipeline.accept || '.xlsx,.xls';
        this.debugLog('opening upload picker for server pipeline', {
            pipelineId: pipeline.id,
            pipelineLabel: pipeline.label,
            accept,
            uploadField: uploadInput.submissionField || pipeline.fileField || 'file',
        });

        const input = document.createElement('input');
        input.type = 'file';
        input.accept = accept;
        input.addEventListener('change', () => {
            if (input.files && input.files.length > 0) {
                this.debugLog('user selected upload file for server pipeline', {
                    pipelineId: pipeline.id,
                    file: {
                        name: input.files[0].name,
                        size: input.files[0].size,
                        type: input.files[0].type || null,
                    },
                });
                void this.runRequestPipeline(pipeline, {
                    uploadFile: input.files[0],
                    uploadField: uploadInput.submissionField || pipeline.fileField || 'file',
                });
            }
        });
        input.click();
    }

    getPipeline(pipelineId) {
        return this.pipelineConfig.find(p => p.id === pipelineId) || null;
    }

    async preparePipelineExecution(pipeline) {
        this.debugLog('preparing server pipeline execution', this.getPipelineRequestPayload(pipeline));
        if (pipeline.requirements && pipeline.requirements.currentDataset && !this.currentDataset) {
            this.app.showNotification('Convert or restore a dataset first', 'info');
            return { ok: false, error: new Error('No current dataset') };
        }

        for (const input of pipeline.inputs || []) {
            const fieldValue = this.getFieldValue(input.valueKey);
            if (input.required && !fieldValue) {
                this.app.showNotification(`Enter ${input.label || input.name} first`, 'info');
                return { ok: false, error: new Error(`Missing required input: ${input.name}`) };
            }
            if (input.source === 'tracked-source' && !fieldValue) {
                const sourceValue = this.sourceInfo && input.sourceInfoKey
                    ? this.sourceInfo[input.sourceInfoKey]
                    : null;
                if (!sourceValue) {
                    this.app.showNotification('No tracked source workbook for this dataset. Convert a workbook first or enter a workbook path.', 'info');
                    return { ok: false, error: new Error('No tracked source workbook') };
                }
            }
        }

        if (pipeline.execution && pipeline.execution.saveCurrentDataset) {
            const initialSave = await this.saveCurrentDataset({ suppressSuccessNotification: true });
            if (!initialSave.ok) {
                return initialSave;
            }
        }

        return { ok: true };
    }

    appendPipelineFormData(form, pipeline) {
        const appended = {};
        if (pipeline.requirements && pipeline.requirements.currentDataset && this.currentDataset) {
            form.append('dataset', this.currentDataset.filename);
            appended.dataset = this.currentDataset.filename;
        }

        for (const input of pipeline.inputs || []) {
            if (input.source === 'upload') {
                continue;
            }
            const value = this.getFieldValue(input.valueKey) || input.default || '';
            if (value) {
                form.append(input.submissionField, value);
                appended[input.submissionField] = value;
            }
        }

        this.debugLog('appended server pipeline form data', {
            pipelineId: pipeline.id,
            payload: appended,
        });
    }

    async runRequestPipeline(pipeline, options = {}) {
        const {
            uploadFile = null,
            uploadField = pipeline.fileField || 'file',
        } = options;

        const preflight = await this.preparePipelineExecution(pipeline);
        if (!preflight.ok) {
            return preflight;
        }

        const targetLabel = uploadFile
            ? `${pipeline.label}: ${uploadFile.name}`
            : `${pipeline.label}${this.currentDataset ? `: ${this.currentDataset.filename}` : ''}`;
        this.debugLog('sending server pipeline request', this.getPipelineRequestPayload(pipeline, {
            uploadFile,
            uploadField,
        }));
        this.app.showNotification(`${targetLabel}…`, 'info');

        try {
            const requestTiming = this.createTimingContext('pipeline request', {
                pipelineId: pipeline.id,
                endpoint: pipeline.execution.endpoint,
            });
            const form = new FormData();
            this.appendPipelineFormData(form, pipeline);
            if (uploadFile) {
                form.append(uploadField, uploadFile);
            }

            const resp = await fetch(pipeline.execution.endpoint, {
                method: 'POST',
                body: form,
            });

            if (!resp.ok) {
                const detail = await resp.json().catch(() => ({}));
                throw new Error(detail.detail || `HTTP ${resp.status}`);
            }

            const result = await resp.json();
            this.logTiming(requestTiming, 'received server pipeline response', {
                pipelineId: pipeline.id,
                filename: result && result.filename ? result.filename : null,
            });
            this.debugLog('received server pipeline response', {
                pipelineId: pipeline.id,
                result,
            });
            return await this.handlePipelineSuccess(pipeline, result);
        } catch (err) {
            console.error('[SERVER] Pipeline failed:', err);
            this.app.showNotification(`${pipeline.label} failed: ${err.message}`, 'error');
            return { ok: false, error: err };
        }
    }

    async handlePipelineSuccess(pipeline, result) {
        const successTiming = this.createTimingContext('pipeline success handling', {
            pipelineId: pipeline.id,
        });
        this.debugLog('handling successful server pipeline result', {
            pipelineId: pipeline.id,
            currentDataset: this.currentDataset ? this.currentDataset.filename : null,
            result,
        });
        if (result.warnings && result.warnings.length > 0) {
            console.warn('[SERVER] Pipeline warnings:', result.warnings);
        }

        if (pipeline.execution && pipeline.execution.promoteAfterSuccess) {
            const promotedSave = await this.saveCurrentDataset({
                promoteChanges: true,
                suppressSuccessNotification: true,
            });
            if (!promotedSave.ok) {
                this.app.showNotification(
                    `${pipeline.label} completed, but promoting the dataset back to JSON failed`,
                    'warning'
                );
                return { ok: false, error: promotedSave.error || new Error('Promoted save failed') };
            }
        }

        const after = result.after || pipeline.after || {};
        const reloadFilename = result.filename || (after.autoLoadDataset && this.currentDataset ? this.currentDataset.filename : null);
        this.debugLog('resolved post-pipeline dataset reload', {
            pipelineId: pipeline.id,
            autoLoadDataset: Boolean(after.autoLoadDataset),
            resultFilename: result.filename || null,
            currentDataset: this.currentDataset ? this.currentDataset.filename : null,
            reloadFilename,
        });
        if (after.autoLoadDataset && reloadFilename) {
            await this.loadServerDataset(reloadFilename, {
                reason: `pipeline:${pipeline.id}`,
            });
        }

        this.logTiming(successTiming, 'completed pipeline success handling', {
            pipelineId: pipeline.id,
            reloadFilename,
            autoLoadDataset: Boolean(after.autoLoadDataset),
        });

        const successMessage = pipeline.execution && pipeline.execution.successMessage
            ? pipeline.execution.successMessage
            : this.buildSuccessMessage(pipeline, result);
        this.app.showNotification(successMessage, 'success');
        return { ok: true, result };
    }

    buildSuccessMessage(pipeline, result) {
        if (result && result.filename && typeof result.items !== 'undefined') {
            return `${pipeline.label}: ${result.filename} (${result.items ?? '?'} items)`;
        }
        return `${pipeline.label} complete`;
    }

    async runInteractivePipeline(pipeline) {
        const preflight = await this.preparePipelineExecution(pipeline);
        if (!preflight.ok) {
            return preflight;
        }

        this.debugLog('opening interactive server pipeline session', this.getPipelineRequestPayload(pipeline));

        this.openConsolePopup(`${pipeline.label}${this.currentDataset ? `: ${this.currentDataset.filename}` : ''}`);

        const wsProtocol = location.protocol === 'https:' ? 'wss:' : 'ws:';
        const wsUrl = `${wsProtocol}//${location.host}${pipeline.execution.endpoint}`;
        this._ws = new WebSocket(wsUrl);

        this._ws.onopen = () => {
            const inputs = {};
            for (const input of pipeline.inputs || []) {
                if (input.source === 'upload') {
                    continue;
                }
                const value = this.getFieldValue(input.valueKey) || input.default || '';
                if (value) {
                    inputs[input.submissionField] = value;
                }
            }

            this._ws.send(JSON.stringify({
                type: 'init',
                dataset: this.currentDataset ? this.currentDataset.filename : '',
                inputs,
            }));
        };

        this._ws.onmessage = async (event) => {
            const msg = JSON.parse(event.data);
            if (msg.type === 'output') {
                this.appendConsoleOutput(msg.text);
            } else if (msg.type === 'prompt') {
                this.showConsolePrompt(msg.text);
            } else if (msg.type === 'complete') {
                this.appendConsoleOutput('\n--- Pipeline complete ---\n');
                if (msg.warnings && msg.warnings.length > 0) {
                    this.appendConsoleOutput('Warnings:\n' + msg.warnings.join('\n') + '\n');
                }
                this.hideConsolePrompt();
                this.setConsoleStatus('Finalizing...');

                const handled = await this.handlePipelineSuccess(pipeline, msg);
                if (handled.ok) {
                    this.setConsoleStatus('Complete.');
                } else {
                    this.appendConsoleOutput('\nWARNING: Pipeline completed, but the follow-up sync failed.\n');
                    this.setConsoleStatus('Pipeline complete, follow-up sync failed.');
                }
            } else if (msg.type === 'error') {
                this.appendConsoleOutput(`\nERROR: ${msg.text}\n`);
                this.setConsoleStatus('Failed.');
                this.hideConsolePrompt();
                this.app.showNotification(`${pipeline.label} failed: ${msg.text}`, 'error');
            }
        };

        this._ws.onerror = () => {
            this.appendConsoleOutput('\nConnection error.\n');
            this.setConsoleStatus('Connection lost.');
            this.hideConsolePrompt();
        };

        this._ws.onclose = () => {
            this._ws = null;
        };

        return { ok: true };
    }

    // ------------------------------------------------------------------
    // Save current dataset back to server
    // ------------------------------------------------------------------

    async saveCurrentDataset(options = {}) {
        if (!this.currentDataset) {
            this.app.showNotification('No dataset loaded to save', 'info');
            return { ok: false, error: new Error('No dataset loaded to save') };
        }

        const {
            promoteChanges = false,
            applyPromotedState = promoteChanges,
            suppressSuccessNotification = false,
        } = options;

        try {
            const data = this.app.persistenceManager
                ? this.app.persistenceManager.prepareDataForSaving({ promoteChanges })
                : { data: this.app.dataset || [] };

            const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
            const form = new FormData();
            form.append('file', blob, this.currentDataset.filename);

            const resp = await fetch(`/api/data/${encodeURIComponent(this.currentDataset.filename)}`, {
                method: 'POST',
                body: form,
            });

            if (!resp.ok) {
                const detail = await resp.json().catch(() => ({}));
                throw new Error(detail.detail || `HTTP ${resp.status}`);
            }

            if (promoteChanges && applyPromotedState && typeof this.app.applyPromotedDataState === 'function') {
                this.app.applyPromotedDataState(data);
            }

            if (!suppressSuccessNotification) {
                this.app.showNotification(
                    promoteChanges
                        ? `Promoted changes and saved ${this.currentDataset.filename} to server`
                        : `Saved ${this.currentDataset.filename} to server`,
                    'success'
                );
            }

            return { ok: true, data };
        } catch (err) {
            console.error('[SERVER] Save failed:', err);
            this.app.showNotification(`Save failed: ${err.message}`, 'error');
            return { ok: false, error: err };
        }
    }

    // ------------------------------------------------------------------
    // Interactive pipeline console
    // ------------------------------------------------------------------

    sendConsoleInput(text) {
        if (this._ws && this._ws.readyState === WebSocket.OPEN) {
            this._ws.send(JSON.stringify({ type: 'input', text }));
            this.appendConsoleOutput(text + '\n');
            this.hideConsolePrompt();
        }
    }

    // ------------------------------------------------------------------
    // Console popup UI
    // ------------------------------------------------------------------

    openConsolePopup(title) {
        const overlay = document.getElementById('console-popup-overlay');
        const titleEl = document.getElementById('console-popup-title');
        const output = document.getElementById('console-popup-output');
        const status = document.getElementById('console-popup-status');
        titleEl.textContent = title || 'Pipeline Console';
        output.textContent = '';
        status.textContent = 'Running…';
        this.hideConsolePrompt();
        overlay.style.display = '';
    }

    closeConsolePopup() {
        document.getElementById('console-popup-overlay').style.display = 'none';
        if (this._ws) {
            this._ws.close();
            this._ws = null;
        }
    }

    appendConsoleOutput(text) {
        const output = document.getElementById('console-popup-output');
        output.textContent += text;
        output.scrollTop = output.scrollHeight;
    }

    showConsolePrompt(promptText) {
        const inputArea = document.getElementById('console-popup-input');
        const promptEl = document.getElementById('console-popup-prompt');
        promptEl.textContent = promptText;
        inputArea.style.display = '';
        // Focus the Yes button for keyboard accessibility
        document.getElementById('console-popup-yes').focus();
    }

    hideConsolePrompt() {
        document.getElementById('console-popup-input').style.display = 'none';
    }

    setConsoleStatus(text) {
        document.getElementById('console-popup-status').textContent = text;
    }

    // ------------------------------------------------------------------
    // UI state
    // ------------------------------------------------------------------

    updateActionState() {
        if (!this.container) {
            return;
        }

        for (const pipeline of this.pipelineConfig) {
            const button = this.container.querySelector(`[data-server-pipeline-id="${CSS.escape(pipeline.id)}"]`);
            if (button) {
                const availability = this.getPipelineAvailability(pipeline);
                button.disabled = !availability.enabled;
                button.title = availability.reason
                    ? `${pipeline.description || pipeline.label}\n${availability.reason}`
                    : (pipeline.description || pipeline.label);
            }
        }
    }

    getPipelineAvailability(pipeline) {
        if (pipeline.requirements && pipeline.requirements.currentDataset && !this.currentDataset) {
            return { enabled: true, reason: 'Load or restore a dataset first.' };
        }

        return { enabled: true, reason: '' };
    }
}


// ===== INTEGRATION WITH EXISTING APP =====

function attachServerControlsIntegration(targetApp) {
    if (!targetApp || targetApp.__serverControlsIntegrated) {
        return;
    }

    targetApp.initializeServerControls = function () {
        if (!this.serverControlsManager) {
            this.serverControlsManager = new ServerControlsManager(this);
        }
        return this.serverControlsManager;
    };

    targetApp.__serverControlsIntegrated = true;
}


// ===== SERVER-MODE DETECTION & INITIALIZATION =====

async function detectServerMode() {
    try {
        const resp = await fetch('/api/health', { method: 'GET' });
        if (resp.ok) {
            const body = await resp.json();
            return body && body.status === 'ok';
        }
    } catch (_ignored) {
        // fetch fails on file:// or non-API servers — not server mode
    }
    return false;
}

document.addEventListener('DOMContentLoaded', async () => {
    if (typeof app === 'undefined' || !app) return;

    // Desktop mode takes precedence — skip server detection inside PySide6 shell
    if (window.desktopBridge || app.isDesktopMode) {
        console.log('[SERVER] Desktop mode detected — skipping server-mode probe');
        return;
    }

    const isServer = await detectServerMode();
    app.isServerMode = isServer;

    const fileControls = document.querySelector('.file-controls');
    const serverControls = document.querySelector('.server-controls');

    if (isServer) {
        if (fileControls) fileControls.style.display = 'none';
        if (serverControls) serverControls.style.display = '';
        attachServerControlsIntegration(app);
        app.initializeServerControls();
        console.log('[SERVER] Server mode active — webhook controls enabled');
    } else {
        if (fileControls) fileControls.style.display = '';
        if (serverControls) serverControls.style.display = 'none';
        console.log('[SERVER] Standalone mode — file controls active');
    }
});
