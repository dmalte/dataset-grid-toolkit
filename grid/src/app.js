// Copyright (c) 2026 Malte Doerper. MIT License. See LICENSE file.

// Core Data Model and Application State
const GridDateUtils = window.GridDateUtils || {
    createDateStamp(date = new Date()) {
        const value = date instanceof Date ? date : new Date(date);
        const pad = (part) => String(part).padStart(2, '0');
        return `${value.getFullYear()}-${pad(value.getMonth() + 1)}-${pad(value.getDate())}`;
    },

    createOffsetStamp(date = new Date()) {
        const value = date instanceof Date ? date : new Date(date);
        const pad = (part) => String(part).padStart(2, '0');
        const offsetMinutes = -value.getTimezoneOffset();
        const sign = offsetMinutes >= 0 ? '+' : '-';
        const absoluteOffsetMinutes = Math.abs(offsetMinutes);
        const offsetHours = Math.floor(absoluteOffsetMinutes / 60);
        const offsetRemainderMinutes = absoluteOffsetMinutes % 60;
        return `${value.getFullYear()}-${pad(value.getMonth() + 1)}-${pad(value.getDate())}` +
            `T${pad(value.getHours())}:${pad(value.getMinutes())}:${pad(value.getSeconds())}` +
            `${sign}${pad(offsetHours)}:${pad(offsetRemainderMinutes)}`;
    },

    normalizeDateStamp(value, fallback = '') {
        if (value instanceof Date && !Number.isNaN(value.getTime())) {
            return this.createDateStamp(value);
        }

        const text = String(value ?? '').trim();
        if (!text) {
            return fallback;
        }

        const directDateMatch = text.match(/^(\d{4}-\d{2}-\d{2})/);
        if (directDateMatch) {
            return directDateMatch[1];
        }

        const parsed = new Date(text);
        if (!Number.isNaN(parsed.getTime())) {
            return this.createDateStamp(parsed);
        }

        return fallback || text;
    },

    createLocalTimestamp(date = new Date()) {
        return this.createOffsetStamp(date);
    },

    normalizeLocalTimestamp(value, fallback = '') {
        if (value instanceof Date && !Number.isNaN(value.getTime())) {
            return this.createLocalTimestamp(value);
        }

        const text = String(value ?? '').trim();
        if (!text) {
            return fallback;
        }

        if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:?\d{2})$/.test(text)) {
            const parsedExisting = new Date(text);
            return Number.isNaN(parsedExisting.getTime()) ? (fallback || text) : this.createLocalTimestamp(parsedExisting);
        }

        const parsed = new Date(text);
        if (!Number.isNaN(parsed.getTime())) {
            return this.createLocalTimestamp(parsed);
        }

        return fallback || text;
    },

    createFilenameTimestamp(date = new Date()) {
        const value = date instanceof Date ? date : new Date(date);
        const pad = (part) => String(part).padStart(2, '0');
        return `${value.getFullYear()}-${pad(value.getMonth() + 1)}-${pad(value.getDate())}` +
            `-${pad(value.getHours())}-${pad(value.getMinutes())}-${pad(value.getSeconds())}`;
    }
};

window.GridDateUtils = GridDateUtils;

class DataVisualizationApp {
    constructor() {
        this.storageKeys = {
            currentUserName: 'grid-current-user-name'
        };

        // Application State
        this.dataset = null;
        this.baselineData = [];
        this.pendingChanges = { version: '1', rows: [] };
        this.viewConfig = null;
        this.metaInfo = null;
        this.filteredData = [];
        this.fieldTypes = new Map();
        this.availableFields = [];
        this.distinctValues = new Map();
        this.schemaFields = {};
        this.relationTypes = [];
        
        // Current UI State
        this.currentAxisSelections = {
            x: null,
            y: null
        };
        this.currentCardSelections = {
            title: null,
            topLeft: null,
            topRight: null,
            bottomRight: null
        };
        this.tableColumnFields = [];
        this.tableSummaryMode = 'none';
        this.currentFilters = new Map();
        this.tagCustomizations = new Map();
        this.selectedCardTag = null;
        this.selectedCardTags = new Set();
        this.selectedControlTags = new Set();
        this.draggedCardTag = null;
        this.currentUserName = this.loadCurrentUserName();
        this.groupMemberships = new Map();
        this.showRelationshipFields = false;
        this.showDerivedFields = false;
        this.showTooltips = false;
        this.growCards = false;
        this.cellRenderMode = 'cards';
        this.aggregationMode = 'sum';
        this.selectedCards = new Set();
        this.selectedTableGraphContexts = new Map();
        
        // Event listeners and initialization
        this.initializeEventListeners();
        
        // V5: Relation UI Manager
        this.relationUIManager = null; // initialized after DOM ready

        // V6: Details Panel Manager
        this.detailsPanelManager = null; // initialized after DOM ready

        // Check for shared state from Compare View before loading built-in data
        if (!this.loadSharedState()) {
            this.loadBuiltInSampleData();
        }
    }

    loadCurrentUserName() {
        try {
            const savedName = localStorage.getItem(this.storageKeys.currentUserName);
            const normalizedName = savedName ? savedName.trim() : '';
            return normalizedName || 'Local User';
        } catch (error) {
            console.warn('Failed to load current user name:', error);
            return 'Local User';
        }
    }

    getCurrentUserName() {
        return this.currentUserName || 'Local User';
    }

    setCurrentUserName(name, options = {}) {
        const {
            persistLocal = true,
            persistServer = this.isServerMode && this.serverControlsManager && typeof this.serverControlsManager.persistCurrentUserName === 'function',
        } = options;
        const normalizedName = String(name || '').trim();
        if (!normalizedName) {
            return false;
        }

        this.currentUserName = normalizedName;

        if (persistLocal) {
            try {
                localStorage.setItem(this.storageKeys.currentUserName, normalizedName);
            } catch (error) {
                console.warn('Failed to persist current user name:', error);
            }
        }

        if (persistServer) {
            void this.serverControlsManager.persistCurrentUserName(normalizedName);
        }

        if (typeof this.renderCurrentUserName === 'function') {
            this.renderCurrentUserName();
        }

        return true;
    }

    getCurrentUserNameStorageKey() {
        return this.storageKeys.currentUserName;
    }
    
    // ===== COMPARE VIEW (V3) =====

    loadSharedState() {
        try {
            const params = new URLSearchParams(window.location.search);
            const sharedKey = params.get('shared');
            if (!sharedKey) return false;

            // Try sessionStorage first (works under same-origin http://)
            let bundle = null;
            try {
                const raw = sessionStorage.getItem(sharedKey);
                if (raw) {
                    bundle = JSON.parse(raw);
                    sessionStorage.removeItem(sharedKey);
                }
            } catch (e) {
                console.warn('[COMPARE] sessionStorage unavailable, trying postMessage fallback');
            }

            if (bundle) {
                console.log('[COMPARE] Loaded shared state from sessionStorage');
                this.loadDataFromJSON(bundle);
                return true;
            }

            // postMessage fallback for file:// protocol
            if (window.opener) {
                this._sharedKey = sharedKey;
                window.addEventListener('message', (event) => {
                    if (event.data && event.data.type === 'shared-state-response' && event.data.key === this._sharedKey) {
                        console.log('[COMPARE] Loaded shared state via postMessage');
                        this.loadDataFromJSON(event.data.bundle);
                        delete this._sharedKey;
                    }
                });
                window.opener.postMessage({ type: 'shared-state-request', key: sharedKey }, '*');
                // Return true to prevent loading built-in sample; data will arrive via message
                return true;
            }

            console.warn('[COMPARE] No shared state found for key:', sharedKey);
            return false;
        } catch (error) {
            console.warn('[COMPARE] Failed to load shared state:', error);
            return false;
        }
    }

    openCompareWindow() {
        const key = 'shared-' + Date.now();
        const bundle = this.serializeFullState();
        const currentUrl = window.location.href.split('?')[0];
        const compareUrl = currentUrl + '?shared=' + encodeURIComponent(key);

        // Try sessionStorage first
        let sessionStorageAvailable = false;
        try {
            sessionStorage.setItem(key, JSON.stringify(bundle));
            sessionStorageAvailable = true;
        } catch (e) {
            console.warn('[COMPARE] sessionStorage unavailable, will use postMessage fallback');
        }

        // Listen for postMessage requests from the child window
        const messageHandler = (event) => {
            if (event.data && event.data.type === 'shared-state-request' && event.data.key === key) {
                event.source.postMessage({
                    type: 'shared-state-response',
                    key: key,
                    bundle: bundle
                }, '*');
                // Clean up after responding
                window.removeEventListener('message', messageHandler);
                if (sessionStorageAvailable) {
                    try { sessionStorage.removeItem(key); } catch (e) { /* ignore */ }
                }
            }
        };
        window.addEventListener('message', messageHandler);

        // Clean up listener after 30 seconds if no request comes
        setTimeout(() => {
            window.removeEventListener('message', messageHandler);
            if (sessionStorageAvailable) {
                try { sessionStorage.removeItem(key); } catch (e) { /* ignore */ }
            }
        }, 30000);

        window.open(compareUrl, '_blank');
        console.log('[COMPARE] Opened compare window:', compareUrl);
    }

    serializeFullState() {
        // Ensure view config is up to date
        this.updateViewConfiguration();

        const state = {
            data: this.cloneDataArray(this.baselineData),
            changes: JSON.parse(JSON.stringify(this.pendingChanges)),
            schema: this.schemaFields ? { fields: JSON.parse(JSON.stringify(this.schemaFields)) } : undefined,
            meta: this.metaInfo ? JSON.parse(JSON.stringify(this.metaInfo)) : undefined,
            view: JSON.parse(JSON.stringify(this.viewConfig))
        };

        // Include relationTypes in schema if present (V4)
        if (this.relationTypes && this.relationTypes.length > 0 && state.schema) {
            state.schema.relationTypes = [...this.relationTypes];
        }

        return state;
    }

    // ===== DATA LOADING AND PARSING =====
    
    loadBuiltInSampleData() {
        // Always use built-in sample data for standalone file:// operation
        // Users can load their own data files through the file picker
        this.loadDataFromJSON(this.getBuiltInSampleData());
    }
    
    getBuiltInSampleData() {
        return {
                    "data": [
                                {
                                            "id": "HOME-001",
                                            "title": "Build seed trays",
                                            "status": "In Progress",
                                            "category": "Gardening",
                                            "owner": "Avery",
                                            "effort": 5,
                                            "tags": [
                                                        "outdoor",
                                                        "weekend"
                                            ],
                                            "dueDate": "2026-04-05",
                                            "comments": {
                                                        "threads": [
                                                                    {
                                                                                "id": "t1",
                                                                                "messages": [
                                                                                            {
                                                                                                        "author": "Avery",
                                                                                                        "timestamp": "2026-03-22T09:30:00Z",
                                                                                                        "text": "Started cutting the boards"
                                                                                            }
                                                                                ]
                                                                    }
                                                        ]
                                            }
                                },
                                {
                                            "id": "HOME-002",
                                            "title": "Plan picnic menu",
                                            "status": "Planned",
                                            "category": "Cooking",
                                            "owner": "Jordan",
                                            "effort": 2,
                                            "tags": [
                                                        "food",
                                                        "friends"
                                            ],
                                            "dueDate": "2026-04-09",
                                            "comments": {
                                                        "threads": []
                                            }
                                },
                                {
                                            "id": "HOME-003",
                                            "title": "Tune guitar pedals",
                                            "status": "Done",
                                            "category": "Music",
                                            "owner": "Avery",
                                            "effort": 3,
                                            "tags": [
                                                        "studio",
                                                        "audio"
                                            ],
                                            "dueDate": "2026-03-20",
                                            "comments": {
                                                        "threads": [
                                                                    {
                                                                                "id": "t2",
                                                                                "messages": [
                                                                                            {
                                                                                                        "author": "Avery",
                                                                                                        "timestamp": "2026-03-20T20:15:00Z",
                                                                                                        "text": "Noise issue resolved"
                                                                                            }
                                                                                ]
                                                                    }
                                                        ]
                                            }
                                },
                                {
                                            "id": "HOME-004",
                                            "title": "Map neighborhood walk",
                                            "status": "In Progress",
                                            "category": "Fitness",
                                            "owner": "Morgan",
                                            "effort": 4,
                                            "tags": [
                                                        "health",
                                                        "outdoors"
                                            ],
                                            "dueDate": "2026-04-02",
                                            "comments": {
                                                        "threads": []
                                            }
                                },
                                {
                                            "id": "HOME-005",
                                            "title": "Sort photo prints",
                                            "status": "Planned",
                                            "category": "Memory Keeping",
                                            "owner": "Taylor",
                                            "effort": 3,
                                            "tags": [
                                                        "archive",
                                                        "family"
                                            ],
                                            "dueDate": "2026-04-12",
                                            "comments": {
                                                        "threads": []
                                            }
                                },
                                {
                                            "id": "HOME-006",
                                            "title": "Refresh balcony lights",
                                            "status": "Done",
                                            "category": "Home Setup",
                                            "owner": "Jordan",
                                            "effort": 2,
                                            "tags": [
                                                        "indoor",
                                                        "weekend"
                                            ],
                                            "dueDate": "2026-03-18",
                                            "comments": {
                                                        "threads": []
                                            }
                                }
                    ],
                    "schema": {
                                "fields": {
                                            "id": {
                                                        "type": "scalar",
                                                        "required": true
                                            },
                                            "title": {
                                                        "type": "scalar",
                                                        "required": true
                                            },
                                            "status": {
                                                        "type": "scalar",
                                                        "required": false
                                            },
                                            "category": {
                                                        "type": "scalar",
                                                        "required": false
                                            },
                                            "owner": {
                                                        "type": "scalar",
                                                        "required": false
                                            },
                                            "effort": {
                                                        "type": "scalar",
                                                        "required": false
                                            },
                                            "tags": {
                                                        "type": "multi-value",
                                                        "required": false
                                            },
                                            "dueDate": {
                                                        "type": "scalar",
                                                        "required": false
                                            },
                                            "comments": {
                                                        "type": "structured",
                                                        "required": false
                                            }
                                }
                    },
                    "view": {
                                "axisSelections": {
                                            "x": "status",
                                            "y": "category",
                                            "title": "title",
                                            "topLeft": "owner",
                                            "topRight": "tags",
                                            "bottomRight": "effort"
                                },
                                "tagCustomizations": {
                                            "outdoor": {
                                                        "label": "Outdoor",
                                                        "color": "#2e8b57"
                                            },
                                            "weekend": {
                                                        "label": "Weekend",
                                                        "color": "#c9771a"
                                            },
                                            "food": {
                                                        "label": "Food",
                                                        "color": "#d1495b"
                                            },
                                            "friends": {
                                                        "label": "Friends",
                                                        "color": "#6c5ce7"
                                            },
                                            "studio": {
                                                        "label": "Studio",
                                                        "color": "#3b82f6"
                                            },
                                            "audio": {
                                                        "label": "Audio",
                                                        "color": "#0f766e"
                                            },
                                            "health": {
                                                        "label": "Health",
                                                        "color": "#15803d"
                                            },
                                            "archive": {
                                                        "label": "Archive",
                                                        "color": "#6b7280"
                                            },
                                            "family": {
                                                        "label": "Family",
                                                        "color": "#b45309"
                                            },
                                            "indoor": {
                                                        "label": "Indoor",
                                                        "color": "#7c3aed"
                                            }
                                }
                    }
        };
    }

    loadDataFromJSON(jsonData) {
        try {
            // Validate and extract data sections
            if (!jsonData.data || !Array.isArray(jsonData.data)) {
                throw new Error('Invalid JSON format: missing or invalid data array');
            }
            
            this.baselineData = this.normalizeDataCommentFields(jsonData.data);
            this.pendingChanges = this.normalizeChangesPayload(jsonData.changes);
            this.dataset = this.buildEffectiveDataset();
            this.metaInfo = jsonData.meta || {};
            this.schemaFields = jsonData.schema && jsonData.schema.fields
                ? jsonData.schema.fields
                : {};
            // V4: Extract relation type registry from schema
            this.relationTypes = jsonData.schema && Array.isArray(jsonData.schema.relationTypes)
                ? [...jsonData.schema.relationTypes]
                : [];
            this.selectedCardTag = null;
            this.selectedCardTags = new Set();
            this.selectedControlTags = new Set();
            this.draggedCardTag = null;
            this.clearTableGraphSelection({ suppressRender: true });
            
            // Process field information
            this.analyzeFields();
            this.pendingChanges = this.normalizeChangesPayload(this.pendingChanges);
            this.dataset = this.buildEffectiveDataset();
            this.viewConfig = this.normalizeViewConfig(jsonData.view);
            this.applyViewConfiguration();
            this.updateFilteredData();
            this.resolveGroupMemberships();

            // Header interactions must exist before the first grid render so
            // drag handles and order diagnostics are available immediately.
            this.initializeAdvancedFeatures();
            
            // Update UI
            this.renderFieldSelectors();
            this.renderSlicers();
            this.renderTags();
            this.renderGroups();
            this.renderCurrentUserName();
            this.renderGrid();
            
            // Initialize interactions after all rendering is complete
            this.initializeInteractions();
            
            // Initialize persistence manager
            this.initializePersistence();
            
            console.log('Data loaded successfully:', {
                baselineItems: this.baselineData.length,
                effectiveItems: this.dataset.length,
                pendingChangeRows: this.pendingChanges.rows.length,
                fields: this.availableFields.length,
                fieldTypes: this.fieldTypes
            });
            
        } catch (error) {
            console.error('Error loading data:', error);
            this.showNotification(error.message, 'error');
        }
    }

    cloneValue(value) {
        if (value === undefined) {
            return undefined;
        }

        if (value === null || typeof value !== 'object') {
            return value;
        }

        if (typeof structuredClone === 'function') {
            return structuredClone(value);
        }

        return JSON.parse(JSON.stringify(value));
    }

    cloneItem(item) {
        if (!item || typeof item !== 'object') {
            return {};
        }

        const clonedItem = this.cloneValue(item);

        if (Array.isArray(clonedItem.relations)) {
            clonedItem.relations = this.normalizeRelationsArray(clonedItem.relations);
        }

        return clonedItem;
    }

    cloneDataArray(data) {
        if (!Array.isArray(data)) {
            return [];
        }

        return data.map((item) => this.cloneItem(item));
    }

    normalizeCommentMessageRecord(message, fallbackId = null) {
        if (!message || typeof message !== 'object') {
            return null;
        }

        const text = String(message.text ?? message.message ?? '').trim();
        if (!text) {
            return null;
        }

        return {
            id: message.id || fallbackId || null,
            author: String(message.author || 'User').trim() || 'User',
            timestamp: GridDateUtils.normalizeLocalTimestamp(message.timestamp, GridDateUtils.createLocalTimestamp()),
            text
        };
    }

    normalizeChangeMetaRecord(meta) {
        const normalizedMeta = meta && typeof meta === 'object'
            ? this.cloneItem(meta)
            : {};
        const now = GridDateUtils.createLocalTimestamp();
        const normalizedCreatedAt = GridDateUtils.normalizeLocalTimestamp(normalizedMeta.createdAt, now);

        return {
            ...normalizedMeta,
            createdAt: normalizedCreatedAt,
            updatedAt: GridDateUtils.normalizeLocalTimestamp(normalizedMeta.updatedAt, normalizedCreatedAt || now)
        };
    }

    normalizeCommentThreadRecord(thread, fallbackId = null) {
        if (!thread || typeof thread !== 'object') {
            return null;
        }

        const threadId = thread.id || fallbackId || `thread-${Date.now()}`;
        const normalizedMessages = [];

        if (Array.isArray(thread.messages)) {
            thread.messages.forEach((message, index) => {
                const normalizedMessage = this.normalizeCommentMessageRecord(message, `${threadId}-${index + 1}`);
                if (normalizedMessage) {
                    normalizedMessages.push(normalizedMessage);
                }
            });
        } else {
            const primaryMessage = this.normalizeCommentMessageRecord(thread, `${threadId}-1`);
            if (primaryMessage) {
                normalizedMessages.push(primaryMessage);
            }

            if (Array.isArray(thread.replies)) {
                thread.replies.forEach((reply, index) => {
                    const normalizedReply = this.normalizeCommentMessageRecord(reply, `${threadId}-reply-${index + 1}`);
                    if (normalizedReply) {
                        normalizedMessages.push(normalizedReply);
                    }
                });
            }
        }

        if (normalizedMessages.length === 0) {
            return null;
        }

        return {
            id: threadId,
            messages: normalizedMessages
        };
    }

    normalizeCommentThreadsValue(value) {
        if (!value || typeof value !== 'object') {
            return null;
        }

        const threads = Array.isArray(value.threads) ? value.threads : null;
        if (!threads) {
            const fallbackThread = this.normalizeCommentThreadRecord(value, value.id || 't1');
            return fallbackThread ? { threads: [fallbackThread] } : null;
        }

        const normalizedThreads = threads
            .map((thread, index) => this.normalizeCommentThreadRecord(thread, `t${index + 1}`))
            .filter((thread) => thread !== null);

        return normalizedThreads.length > 0 ? { threads: normalizedThreads } : null;
    }

    normalizeLegacyCommentsArray(comments) {
        if (!Array.isArray(comments)) {
            return null;
        }

        const normalizedThreads = comments
            .map((comment, index) => this.normalizeCommentThreadRecord(comment, comment && comment.id ? comment.id : `t${index + 1}`))
            .filter((thread) => thread !== null);

        return normalizedThreads.length > 0 ? { threads: normalizedThreads } : null;
    }

    mergeCommentValues(primaryValue, legacyCommentsValue) {
        const normalizedThreads = [];
        const primaryThreads = this.normalizeCommentThreadsValue(primaryValue);
        const legacyThreads = this.normalizeLegacyCommentsArray(legacyCommentsValue);

        if (primaryThreads && Array.isArray(primaryThreads.threads)) {
            normalizedThreads.push(...primaryThreads.threads.map((thread) => this.cloneValue(thread)));
        }

        if (legacyThreads && Array.isArray(legacyThreads.threads)) {
            normalizedThreads.push(...legacyThreads.threads.map((thread) => this.cloneValue(thread)));
        }

        return normalizedThreads.length > 0 ? { threads: normalizedThreads } : null;
    }

    normalizeItemCommentFields(item) {
        if (!item || typeof item !== 'object') {
            return {};
        }

        const normalizedItem = this.cloneItem(item);
        const mergedComment = this.mergeCommentValues(normalizedItem.comment, normalizedItem.comments);

        if (mergedComment) {
            normalizedItem.comment = mergedComment;
        } else {
            delete normalizedItem.comment;
        }

        delete normalizedItem.comments;
        return normalizedItem;
    }

    normalizeDataCommentFields(data) {
        if (!Array.isArray(data)) {
            return [];
        }

        return data.map((item) => this.normalizeItemCommentFields(item));
    }

    getItemCommentThreads(item) {
        return this.mergeCommentValues(item && item.comment, item && item.comments);
    }

    normalizeRelationPriority(priority) {
        if (priority === undefined || priority === null || priority === '') {
            return 1;
        }

        const normalizedPriority = Number(priority);
        return Number.isFinite(normalizedPriority) ? normalizedPriority : 1;
    }

    normalizeRelationObject(relation) {
        if (!relation || typeof relation !== 'object') {
            return null;
        }

        const normalizedRelation = this.cloneValue(relation);
        normalizedRelation.priority = this.normalizeRelationPriority(normalizedRelation.priority);
        if (normalizedRelation.meta && typeof normalizedRelation.meta === 'object') {
            const normalizedCreatedAt = GridDateUtils.normalizeLocalTimestamp(normalizedRelation.meta.createdAt, '');
            const normalizedUpdatedAt = GridDateUtils.normalizeLocalTimestamp(normalizedRelation.meta.updatedAt, normalizedCreatedAt);
            if (normalizedCreatedAt) {
                normalizedRelation.meta.createdAt = normalizedCreatedAt;
            }
            if (normalizedUpdatedAt) {
                normalizedRelation.meta.updatedAt = normalizedUpdatedAt;
            }
        }
        return normalizedRelation;
    }

    normalizeRelationsArray(relations) {
        if (!Array.isArray(relations)) {
            return [];
        }

        return relations
            .map((relation) => this.normalizeRelationObject(relation))
            .filter((relation) => relation !== null);
    }

    normalizeRelationChangeEntry(entry) {
        if (!entry || typeof entry !== 'object') {
            return null;
        }

        const normalizedEntry = this.cloneItem(entry);

        if (normalizedEntry.target && typeof normalizedEntry.target === 'object') {
            if (normalizedEntry.target.ownerItemId !== undefined && normalizedEntry.target.ownerItemId !== null) {
                normalizedEntry.target.ownerItemId = String(normalizedEntry.target.ownerItemId);
            }
            if (normalizedEntry.target.relationId !== undefined && normalizedEntry.target.relationId !== null) {
                normalizedEntry.target.relationId = String(normalizedEntry.target.relationId);
            }
        }

        if (normalizedEntry.proposed && typeof normalizedEntry.proposed === 'object') {
            normalizedEntry.proposed = this.normalizeRelationObject(normalizedEntry.proposed) || {};
        }

        if (normalizedEntry.baseline && typeof normalizedEntry.baseline === 'object' && Object.keys(normalizedEntry.baseline).length > 0) {
            normalizedEntry.baseline = this.normalizeRelationObject(normalizedEntry.baseline) || {};
        }

        if (normalizedEntry.meta && typeof normalizedEntry.meta === 'object') {
            normalizedEntry.meta = this.normalizeChangeMetaRecord(normalizedEntry.meta);
        }

        return normalizedEntry;
    }

    getEditorInputValue(fieldName, value) {
        if (value === undefined || value === null) {
            return '';
        }

        const inputType = this.getFormInputType(fieldName);
        if (inputType === 'date') {
            const normalizedValue = String(value).match(/^(\d{4}-\d{2}-\d{2})/);
            return normalizedValue ? normalizedValue[1] : String(value);
        }

        return String(value);
    }

    normalizeFieldValueForStorage(fieldName, value, referenceValue = undefined) {
        const normalizedFieldName = fieldName === 'comments' ? 'comment' : fieldName;
        const fieldType = this.fieldTypes.get(normalizedFieldName) || this.fieldTypes.get(fieldName);

        if (normalizedFieldName === 'comment') {
            return this.cloneValue(this.getItemCommentThreads({
                comment: fieldName === 'comment' ? value : null,
                comments: fieldName === 'comments' ? value : null
            }));
        }

        if (fieldType === 'multi-value') {
            const normalizedEntries = (Array.isArray(value) ? value : value === null || value === undefined || value === '' ? [] : [value])
                .map((entry) => String(entry).trim())
                .filter((entry) => entry.length > 0);

            if (Array.isArray(referenceValue)) {
                return normalizedEntries;
            }

            if (typeof referenceValue === 'string') {
                if (normalizedEntries.length === 0) {
                    return null;
                }

                if (normalizedEntries.length === 1) {
                    return normalizedEntries[0];
                }
            }

            return normalizedEntries;
        }

        if (value === undefined) {
            return undefined;
        }

        if (value === null) {
            return null;
        }

        if (fieldType === 'structured') {
            return this.cloneValue(value);
        }

        if (typeof value === 'string') {
            const trimmedValue = value.trim();
            if (trimmedValue === '') {
                return null;
            }

            const inputType = this.getFormInputType(normalizedFieldName);
            if (inputType === 'date' && /^\d{4}-\d{2}-\d{2}$/.test(trimmedValue)) {
                return trimmedValue;
            }

            if ((typeof referenceValue === 'number' || (referenceValue === undefined && this.isNumericField(normalizedFieldName))) && !Number.isNaN(Number(trimmedValue))) {
                return Number(trimmedValue);
            }

            return value;
        }

        return this.cloneValue(value);
    }

    normalizeFieldValueForComparison(fieldName, value) {
        const normalizedFieldName = fieldName === 'comments' ? 'comment' : fieldName;
        const fieldType = this.fieldTypes.get(normalizedFieldName) || this.fieldTypes.get(fieldName);

        if (normalizedFieldName === 'comment') {
            return this.cloneValue(this.getItemCommentThreads({
                comment: fieldName === 'comment' ? value : null,
                comments: fieldName === 'comments' ? value : null
            }));
        }

        if (fieldType === 'multi-value') {
            return (Array.isArray(value) ? value : value === null || value === undefined || value === '' ? [] : [value])
                .map((entry) => String(entry).trim())
                .filter((entry) => entry.length > 0);
        }

        if (value === undefined || value === null) {
            return null;
        }

        if (typeof value === 'string') {
            if (this.getFormInputType(normalizedFieldName) === 'date') {
                const normalizedValue = value.match(/^(\d{4}-\d{2}-\d{2})/);
                return normalizedValue ? normalizedValue[1] : value;
            }

            if (this.isNumericField(normalizedFieldName)) {
                const trimmedValue = value.trim();
                if (trimmedValue !== '' && !Number.isNaN(Number(trimmedValue))) {
                    return Number(trimmedValue);
                }
            }
        }

        return this.cloneValue(value);
    }

    areFieldValuesEquivalent(fieldName, firstValue, secondValue) {
        return JSON.stringify(this.normalizeFieldValueForComparison(fieldName, firstValue)) ===
            JSON.stringify(this.normalizeFieldValueForComparison(fieldName, secondValue));
    }

    normalizeChangesPayload(changes) {
        const seenChangeIds = new Set();
        const rows = Array.isArray(changes && changes.rows)
            ? changes.rows
                .filter((row) => row && typeof row === 'object')
                .map((row) => {
                    let changeId = row.changeId ? String(row.changeId) : '';

                    if (!changeId || seenChangeIds.has(changeId)) {
                        do {
                            changeId = this.generateChangeId();
                        } while (seenChangeIds.has(changeId));
                    }

                    seenChangeIds.add(changeId);

                    return this.normalizeChangeRow({
                        changeId,
                        action: row.action === 'create' ? 'create' : 'update',
                        target: this.cloneItem(row.target || {}),
                        baseline: this.cloneItem(row.baseline || {}),
                        proposed: this.cloneItem(row.proposed || {}),
                        meta: this.normalizeChangeMetaRecord(row.meta)
                    });
                })
                .filter((row) => row !== null)
            : [];

        const result = {
            version: String(changes && changes.version ? changes.version : '1'),
            rows
        };

        // Preserve relation changes (V4) — pass through without field normalization
        if (changes && Array.isArray(changes.relations) && changes.relations.length > 0) {
            result.relations = changes.relations
                .map((entry) => this.normalizeRelationChangeEntry(entry))
                .filter((entry) => entry !== null);
        }

        return result;
    }

    normalizeChangeRow(row) {
        if (!row || typeof row !== 'object') {
            return null;
        }

        if (row.action === 'create') {
            const normalizedProposed = {};
            Object.entries(this.cloneItem(row.proposed || {})).forEach(([fieldName, value]) => {
                normalizedProposed[fieldName] = this.normalizeFieldValueForStorage(fieldName, value);
            });

            return {
                ...row,
                baseline: {},
                proposed: normalizedProposed
            };
        }

        const normalizedBaseline = {};
        const normalizedProposed = {};
        const targetItemId = row.target && row.target.itemId !== undefined && row.target.itemId !== null
            ? String(row.target.itemId)
            : null;
        const baselineItem = targetItemId ? this.getBaselineItemById(targetItemId) : null;
        const changedFields = new Set([
            ...Object.keys(row.baseline || {}),
            ...Object.keys(row.proposed || {})
        ]);

        changedFields.forEach((fieldName) => {
            const baselineValue = baselineItem && Object.prototype.hasOwnProperty.call(baselineItem, fieldName)
                ? this.cloneValue(baselineItem[fieldName])
                : this.cloneValue((row.baseline || {})[fieldName]);
            const normalizedNextValue = this.normalizeFieldValueForStorage(
                fieldName,
                (row.proposed || {})[fieldName],
                baselineValue
            );

            if (this.areFieldValuesEquivalent(fieldName, baselineValue, normalizedNextValue)) {
                return;
            }

            normalizedBaseline[fieldName] = baselineValue;
            normalizedProposed[fieldName] = normalizedNextValue;
        });

        if (Object.keys(normalizedProposed).length === 0) {
            return null;
        }

        return {
            ...row,
            baseline: normalizedBaseline,
            proposed: normalizedProposed
        };
    }

    ensureChangesContainer() {
        if (!this.pendingChanges || typeof this.pendingChanges !== 'object') {
            this.pendingChanges = { version: '1', rows: [] };
            return this.pendingChanges;
        }

        if (!Array.isArray(this.pendingChanges.rows)) {
            this.pendingChanges.rows = [];
        }

        if (!this.pendingChanges.version) {
            this.pendingChanges.version = '1';
        }

        // Preserve relations array if present (V4)
        if (this.pendingChanges.relations && !Array.isArray(this.pendingChanges.relations)) {
            this.pendingChanges.relations = [];
        }

        return this.pendingChanges;
    }

    generateChangeId() {
        const existingRows = this.pendingChanges && Array.isArray(this.pendingChanges.rows)
            ? this.pendingChanges.rows
            : [];
        const existingNumericSuffixes = existingRows
            .map((row) => {
                const match = String(row.changeId || '').match(/chg-(\d+)$/);
                return match ? parseInt(match[1], 10) : 0;
            })
            .filter((value) => Number.isFinite(value));

        const nextNumber = existingNumericSuffixes.length > 0
            ? Math.max(...existingNumericSuffixes) + 1
            : 1;

        return `chg-${String(nextNumber).padStart(3, '0')}`;
    }

    getSourceRefIdentity(sourceRef) {
        if (!sourceRef || typeof sourceRef !== 'object') {
            return null;
        }

        const orderedKeys = ['sourceType', 'rawId', 'key'];
        const sourceRefParts = orderedKeys
            .filter((key) => sourceRef[key] !== undefined && sourceRef[key] !== null && String(sourceRef[key]).trim() !== '')
            .map((key) => `${key}=${String(sourceRef[key]).trim()}`);

        if (sourceRefParts.length === 0) {
            return null;
        }

        return `sourceRef:${sourceRefParts.join('|')}`;
    }

    getStableItemIdentity(item, fallbackIndex = null) {
        if (!item || typeof item !== 'object') {
            return null;
        }

        const internalId = item.__gridItemId;
        if (internalId !== undefined && internalId !== null && String(internalId).trim() !== '') {
            return String(internalId);
        }

        if (item.id !== undefined && item.id !== null && String(item.id).trim() !== '') {
            return String(item.id);
        }

        const sourceRefIdentity = this.getSourceRefIdentity(item.sourceRef);
        if (sourceRefIdentity) {
            return sourceRefIdentity;
        }

        const alternateIdentityField = ['TxnID', 'txnId', 'transactionId', 'transactionID', 'rawId', 'key']
            .find((fieldName) => item[fieldName] !== undefined && item[fieldName] !== null && String(item[fieldName]).trim() !== '');

        if (alternateIdentityField) {
            return `${alternateIdentityField}:${String(item[alternateIdentityField]).trim()}`;
        }

        if (Number.isInteger(fallbackIndex) && fallbackIndex >= 0) {
            return `row:${fallbackIndex + 1}`;
        }

        return null;
    }

    getItemIdentity(item) {
        return this.getStableItemIdentity(item);
    }

    assignEffectiveItemIdentity(item, identity) {
        if (!item || typeof item !== 'object') {
            return item;
        }

        Object.defineProperty(item, '__gridItemId', {
            value: identity,
            writable: true,
            configurable: true,
            enumerable: false
        });

        return item;
    }

    buildEffectiveDataset() {
        const effectiveItems = this.baselineData.map((item, index) => {
            const clonedItem = this.cloneItem(item);
            const identity = this.getStableItemIdentity(item, index);
            return this.assignEffectiveItemIdentity(clonedItem, identity);
        });
        const itemsByIdentity = new Map(
            effectiveItems.map((item) => [this.getItemIdentity(item), item])
        );

        // Step 1: Apply item-level changes (changes.rows)
        this.ensureChangesContainer().rows.forEach((row) => {
            const targetItemId = row && row.target && row.target.itemId
                ? String(row.target.itemId)
                : null;

            if (!targetItemId) {
                return;
            }

            if (row.action === 'create') {
                const createdItem = this.cloneItem(row.proposed || {});
                if (!Object.prototype.hasOwnProperty.call(createdItem, 'id') || createdItem.id === null || createdItem.id === '') {
                    createdItem.id = targetItemId;
                }
                this.assignEffectiveItemIdentity(createdItem, targetItemId);
                itemsByIdentity.set(targetItemId, createdItem);
                return;
            }

            const existingItem = itemsByIdentity.get(targetItemId);
            if (!existingItem) {
                return;
            }

            Object.entries(row.proposed || {}).forEach(([fieldName, value]) => {
                existingItem[fieldName] = this.cloneValue(value);
            });
        });

        // Step 2: Apply relation changes (changes.relations) — V4
        this.applyRelationChanges(itemsByIdentity);

        // Step 3: Derive inverse relations — V4
        const allItems = Array.from(itemsByIdentity.values());
        this.deriveInverseRelations(allItems, itemsByIdentity);

        // Step 4: Compute relation-derived fields — V5
        if (this.relationUIManager) {
            this.relationUIManager.computeDerivedRelationFields(allItems);
        }

        // Step 5: Compute non-relation derived fields
        this.computeAdditionalDerivedFields(allItems);

        return allItems;
    }

    computeAdditionalDerivedFields(dataset) {
        if (!Array.isArray(dataset)) {
            return;
        }

        this.computeDateDerivedFields(dataset);
        this.applyGroupDerivedFields(dataset);
    }

    computeDateDerivedFields(dataset) {
        const sourceFields = new Set();

        dataset.forEach((item) => {
            Object.entries(item || {}).forEach(([fieldName, value]) => {
                if (!fieldName || fieldName.startsWith('_')) {
                    return;
                }

                if (Array.isArray(value) || (typeof value === 'object' && value !== null)) {
                    return;
                }

                if (this.parseDerivedDateValue(value, fieldName)) {
                    sourceFields.add(fieldName);
                }
            });
        });

        dataset.forEach((item) => {
            sourceFields.forEach((fieldName) => {
                const parsedDate = this.parseDerivedDateValue(item[fieldName], fieldName);
                const monthField = `_${fieldName}Month`;
                const weekField = `_${fieldName}CalendarWeek`;

                if (!parsedDate) {
                    delete item[monthField];
                    delete item[weekField];
                    return;
                }

                item[monthField] = this.formatDerivedMonthValue(parsedDate);
                item[weekField] = this.formatDerivedCalendarWeekValue(parsedDate);
            });
        });
    }

    parseDerivedDateValue(value, fieldName = '') {
        if (value === undefined || value === null || value === '') {
            return null;
        }

        if (value instanceof Date && !Number.isNaN(value.getTime())) {
            return new Date(Date.UTC(value.getUTCFullYear(), value.getUTCMonth(), value.getUTCDate()));
        }

        const stringValue = String(value).trim();
        if (!stringValue) {
            return null;
        }

        const likelyDateField = /date|due|deadline|start|end|created|updated|modified|timestamp/i.test(String(fieldName));
        const dateOnlyMatch = stringValue.match(/^(\d{4})-(\d{2})-(\d{2})$/);
        if (dateOnlyMatch) {
            const [, year, month, day] = dateOnlyMatch;
            return new Date(Date.UTC(Number(year), Number(month) - 1, Number(day)));
        }

        if (!likelyDateField && !/^\d{4}-\d{2}-\d{2}T/.test(stringValue)) {
            return null;
        }

        const parsed = new Date(stringValue);
        if (Number.isNaN(parsed.getTime())) {
            return null;
        }

        return new Date(Date.UTC(parsed.getUTCFullYear(), parsed.getUTCMonth(), parsed.getUTCDate()));
    }

    formatDerivedMonthValue(date) {
        const year = date.getUTCFullYear();
        const month = String(date.getUTCMonth() + 1).padStart(2, '0');
        return `${year}-${month}`;
    }

    formatDerivedCalendarWeekValue(date) {
        const workingDate = new Date(Date.UTC(
            date.getUTCFullYear(),
            date.getUTCMonth(),
            date.getUTCDate()
        ));
        const dayNumber = workingDate.getUTCDay() || 7;
        workingDate.setUTCDate(workingDate.getUTCDate() + 4 - dayNumber);

        const weekYear = workingDate.getUTCFullYear();
        const yearStart = new Date(Date.UTC(weekYear, 0, 1));
        const weekNumber = Math.ceil((((workingDate - yearStart) / 86400000) + 1) / 7);
        return `${weekYear}-W${String(weekNumber).padStart(2, '0')}`;
    }

    applyGroupDerivedFields(dataset) {
        if (!Array.isArray(dataset)) {
            return;
        }

        const enabledGroups = this.viewConfig && Array.isArray(this.viewConfig.groups)
            ? this.viewConfig.groups.filter((group) => group && group.enabled !== false)
            : [];

        dataset.forEach((item) => {
            if (enabledGroups.length === 0) {
                delete item._groups;
                delete item._groupCount;
                return;
            }

            const memberships = enabledGroups.filter((group) => this.evaluateGroupMembership(item, group));
            if (memberships.length === 0) {
                delete item._groups;
                delete item._groupCount;
                return;
            }

            item._groups = memberships.map((group) => group.name);
            item._groupCount = memberships.length;
        });
    }

    // ===== V4 RELATION SUPPORT =====

    static INVERSE_TYPE_MAP = {
        parent: 'child',
        child: 'parent',
        epic: 'subtask',
        subtask: 'epic',
        blocks: 'blockedBy',
        blockedBy: 'blocks',
        relatesTo: 'relatesTo',
        duplicates: 'duplicates'
    };

    applyRelationChanges(itemsByIdentity) {
        const changes = this.ensureChangesContainer();
        if (!Array.isArray(changes.relations) || changes.relations.length === 0) {
            return;
        }

        changes.relations.forEach((entry) => {
            if (!entry || !entry.target) return;
            const ownerItemId = entry.target.ownerItemId ? String(entry.target.ownerItemId) : null;
            const relationId = entry.target.relationId ? String(entry.target.relationId) : null;
            if (!ownerItemId) return;

            const ownerItem = itemsByIdentity.get(ownerItemId);
            if (!ownerItem) return;

            if (!Array.isArray(ownerItem.relations)) {
                ownerItem.relations = [];
            }

            if (entry.action === 'create' && entry.proposed) {
                const newRelation = this.normalizeRelationObject(entry.proposed);
                if (relationId && !newRelation.relationId) {
                    newRelation.relationId = relationId;
                }
                ownerItem.relations.push(newRelation);
            } else if (entry.action === 'update' && relationId && entry.proposed) {
                const existing = ownerItem.relations.find((r) => r.relationId === relationId);
                if (existing) {
                    Object.entries(entry.proposed).forEach(([key, value]) => {
                        existing[key] = JSON.parse(JSON.stringify(value));
                    });
                }
            } else if (entry.action === 'delete' && relationId) {
                ownerItem.relations = ownerItem.relations.filter((r) => r.relationId !== relationId);
            }

            ownerItem.relations = this.normalizeRelationsArray(ownerItem.relations);
        });
    }

    deriveInverseRelations(allItems, itemsByIdentity) {
        allItems.forEach((item) => {
            if (!Array.isArray(item.relations)) return;

            item.relations.forEach((rel) => {
                if (rel._derived) return;
                const inverseType = DataVisualizationApp.INVERSE_TYPE_MAP[rel.type];
                if (!inverseType) return;

                // Resolve target item
                const targetId = rel.target && rel.target.itemId ? String(rel.target.itemId) : null;
                if (!targetId) return;
                const targetItem = itemsByIdentity.get(targetId);
                if (!targetItem) return;

                if (!Array.isArray(targetItem.relations)) {
                    targetItem.relations = [];
                }

                // For symmetric types, skip if the target already has the relation to this item
                const sourceItemId = this.getItemIdentity(item);
                const alreadyPresent = targetItem.relations.some(
                    (r) => r.type === inverseType && r.target && String(r.target.itemId) === sourceItemId && !r._derived
                );
                if (alreadyPresent) return;

                // Also skip if a derived inverse already exists
                const derivedExists = targetItem.relations.some(
                    (r) => r.type === inverseType && r.target && String(r.target.itemId) === sourceItemId && r._derived
                );
                if (derivedExists) return;

                targetItem.relations.push({
                    relationId: rel.relationId ? `${rel.relationId}-inv` : undefined,
                    type: inverseType,
                    priority: this.normalizeRelationPriority(rel.priority),
                    direction: 'outward',
                    target: {
                        itemId: sourceItemId,
                        sourceRef: item.sourceRef ? JSON.parse(JSON.stringify(item.sourceRef)) : undefined
                    },
                    _derived: true
                });
            });
        });
    }

    rebuildEffectiveDataset(options = {}) {
        const { render = false } = options;
        this.dataset = this.buildEffectiveDataset();
        this.updateFilteredData();

        if (render) {
            this.renderGrid();
        }
    }

    getBaselineItemById(itemId) {
        const normalizedItemId = itemId === undefined || itemId === null ? null : String(itemId);
        if (!normalizedItemId) {
            return null;
        }

        return this.baselineData.find((item, index) => {
            const identity = this.getStableItemIdentity(item, index);
            return identity === normalizedItemId || String(item.id) === normalizedItemId;
        }) || null;
    }

    getEffectiveItemById(itemId) {
        const normalizedItemId = itemId === undefined || itemId === null ? null : String(itemId);
        if (!normalizedItemId || !Array.isArray(this.dataset)) {
            return null;
        }

        return this.dataset.find((item) => {
            const identity = this.getItemIdentity(item);
            return identity === normalizedItemId || String(item.id) === normalizedItemId;
        }) || null;
    }

    getPendingChangeRowForItem(itemId, action = null) {
        const normalizedItemId = itemId === undefined || itemId === null ? null : String(itemId);
        if (!normalizedItemId) {
            return null;
        }

        return this.ensureChangesContainer().rows.find((row) => {
            const matchesItem = row && row.target && String(row.target.itemId) === normalizedItemId;
            const matchesAction = !action || row.action === action;
            return matchesItem && matchesAction;
        }) || null;
    }

    removePendingChangeRow(changeId) {
        const rows = this.ensureChangesContainer().rows;
        const rowIndex = rows.findIndex((row) => row.changeId === changeId);
        if (rowIndex === -1) {
            return false;
        }

        rows.splice(rowIndex, 1);
        return true;
    }

    createChangeMeta(existingMeta = null) {
        const now = GridDateUtils.createLocalTimestamp();
        const baseMeta = existingMeta && typeof existingMeta === 'object'
            ? this.cloneItem(existingMeta)
            : {};

        return {
            author: this.getCurrentUserName(),
            createdAt: GridDateUtils.normalizeLocalTimestamp(baseMeta.createdAt, now),
            updatedAt: now
        };
    }

    getNextGeneratedItemId() {
        const sourceItems = Array.isArray(this.dataset) ? this.dataset : [];
        const maxId = sourceItems.length > 0
            ? Math.max(...sourceItems.map((item) => parseInt(String(item.id || '').replace(/\D/g, ''), 10) || 0))
            : 0;

        return `ITEM-${maxId + 1}`;
    }

    getPendingChangesSnapshot() {
        return this.normalizeChangesPayload(this.ensureChangesContainer());
    }

    getBaselineDataSnapshot() {
        return this.cloneDataArray(this.baselineData);
    }

    getPromotedDataSnapshot() {
        const sourceItems = Array.isArray(this.dataset) ? this.dataset : [];
        return sourceItems.map((item) => this.sanitizeItemForPromotion(item));
    }

    sanitizeItemForPromotion(item) {
        const sanitizedItem = this.cloneItem(item || {});

        Object.keys(sanitizedItem).forEach((fieldName) => {
            const kind = this.getFieldKind(fieldName);
            if (kind === 'derived' || kind === 'relationship') {
                delete sanitizedItem[fieldName];
            }
        });

        if (Array.isArray(sanitizedItem.relations)) {
            const promotedRelations = sanitizedItem.relations
                .filter((relation) => relation && !relation._derived)
                .map((relation) => {
                    const normalizedRelation = this.normalizeRelationObject(relation) || {};
                    delete normalizedRelation._derived;
                    return normalizedRelation;
                });

            if (promotedRelations.length > 0) {
                sanitizedItem.relations = promotedRelations;
            } else {
                delete sanitizedItem.relations;
            }
        }

        return sanitizedItem;
    }

    applyPromotedDataState(savedPayload = {}) {
        const normalizedPayload = savedPayload && typeof savedPayload === 'object'
            ? savedPayload
            : {};
        const schema = normalizedPayload.schema && typeof normalizedPayload.schema === 'object'
            ? normalizedPayload.schema
            : {};
        const changesVersion = normalizedPayload.changes && normalizedPayload.changes.version
            ? String(normalizedPayload.changes.version)
            : '1';

        this.baselineData = this.cloneDataArray(Array.isArray(normalizedPayload.data) ? normalizedPayload.data : []);
        this.pendingChanges = this.normalizeChangesPayload({
            version: changesVersion,
            rows: [],
            relations: []
        });
        this.schemaFields = schema.fields ? this.cloneItem(schema.fields) : {};
        this.relationTypes = Array.isArray(schema.relationTypes)
            ? this.cloneValue(schema.relationTypes)
            : [];
        this.metaInfo = normalizedPayload.meta ? this.cloneItem(normalizedPayload.meta) : (this.metaInfo || {});

        this.dataset = this.buildEffectiveDataset();
        this.analyzeFields();
        this.pendingChanges = this.normalizeChangesPayload(this.pendingChanges);
        this.dataset = this.buildEffectiveDataset();
        this.updateFilteredData();
        this.resolveGroupMemberships();
        this.updateViewConfiguration();
        this.renderFieldSelectors();
        this.renderSlicers();
        this.renderTags();
        this.renderGroups();
        this.renderGrid();

        if (this.detailsPanelManager && typeof this.detailsPanelManager.refreshForSelection === 'function' && this.detailsPanelManager.isOpen()) {
            this.detailsPanelManager.refreshForSelection();
        }
    }
    
    // ===== FIELD ANALYSIS AND TYPE DETECTION =====
    
    analyzeFields() {
        this.fieldTypes.clear();
        this.availableFields = [];
        this.distinctValues.clear();
        
        if (!this.dataset || this.dataset.length === 0) return;
        
        // Collect all possible field names from all items
        const allFieldNames = new Set();
        this.dataset.forEach(item => {
            Object.keys(item).forEach(key => allFieldNames.add(key));
        });
        
        // Analyze each field
        allFieldNames.forEach(fieldName => {
            this.analyzeField(fieldName);
        });
        
        // Sort fields for consistent ordering
        this.availableFields.sort();

        // V5: Register relation-derived fields
        if (this.relationUIManager) {
            this.relationUIManager.registerDerivedFields();
        }
        
        console.log('Field analysis complete:', {
            scalar: this.getFieldsByType('scalar'),
            multiValue: this.getFieldsByType('multi-value'),
            structured: this.getFieldsByType('structured')
        });
    }
    
    analyzeField(fieldName) {
        const values = [];
        const distinctSet = new Set();
        const schemaFieldType = this.schemaFields[fieldName] ? this.schemaFields[fieldName].type : null;
        
        // Sample values from all items
        this.dataset.forEach(item => {
            const value = item[fieldName];
            if (value !== undefined && value !== null && value !== '') {
                values.push(value);
                
                if (Array.isArray(value)) {
                    const isStructuredArray = this.isStructuredArrayValue(fieldName, value);
                    if (!isStructuredArray) {
                        value.forEach(v => distinctSet.add(String(v)));
                    }
                } else if (typeof value === 'object') {
                    // Structured field - don't add to distinct values
                } else {
                    // Scalar field
                    distinctSet.add(String(value));
                }
            }
        });
        
        // Determine field type based on analysis
        let fieldType = 'scalar';
        if (schemaFieldType === 'scalar' || schemaFieldType === 'multi-value' || schemaFieldType === 'structured') {
            fieldType = schemaFieldType;
        } else if (values.length > 0) {
            const firstValue = values[0];
            if (Array.isArray(firstValue)) {
                fieldType = this.isStructuredArrayValue(fieldName, firstValue)
                    ? 'structured'
                    : 'multi-value';
            } else if (typeof firstValue === 'object' && firstValue !== null) {
                fieldType = 'structured';
            }
        } else if (this.isLikelyStructuredFieldName(fieldName)) {
            fieldType = 'structured';
        }
        
        this.fieldTypes.set(fieldName, fieldType);

        const schemaField = this.schemaFields[fieldName];
        if (schemaField && Array.isArray(schemaField.validValues)) {
            this.distinctValues.set(fieldName, Array.from(new Set(
                schemaField.validValues
                    .filter((value) => value !== undefined && value !== null && value !== '')
                    .map((value) => String(value))
            )).sort());
        } else {
            this.distinctValues.set(fieldName, Array.from(distinctSet).sort());
        }
        
        // Add to available fields if selectable
        if (this.isSelectableField(fieldName, fieldType)) {
            this.availableFields.push(fieldName);
        }
    }

    isStructuredArrayValue(fieldName, value) {
        if (!Array.isArray(value)) {
            return false;
        }

        if (this.isLikelyStructuredFieldName(fieldName)) {
            return true;
        }

        return value.some((entry) => Array.isArray(entry) || (typeof entry === 'object' && entry !== null));
    }

    isLikelyStructuredFieldName(fieldName) {
        const normalizedName = String(fieldName).toLowerCase();
        return normalizedName === 'comment' || normalizedName === 'comments' || normalizedName === 'relations';
    }
    
    isSelectableField(fieldName, fieldType) {
        const sf = this.schemaFields[fieldName];
        if (sf && (sf.visible === false || sf.selectable === false)) return false;
        return fieldType !== 'structured';
    }

    getFieldKind(fieldName) {
        const sf = this.schemaFields[fieldName];
        if (sf && sf.kind) return sf.kind;
        // Treat any _-prefixed field as derived
        if (fieldName.startsWith('_')) return 'derived';
        return 'data';
    }
    
    getFieldsByType(type) {
        return Array.from(this.fieldTypes.entries())
            .filter(([, fieldType]) => fieldType === type)
            .map(([fieldName]) => fieldName);
    }
    
    canUseForAxis(fieldName) {
        const fieldType = this.fieldTypes.get(fieldName);
        // Axis selectors are limited to scalar fields to keep placement deterministic.
        return fieldType === 'scalar';
    }

    canUseForValue(fieldName) {
        const fieldType = this.fieldTypes.get(fieldName);
        // Bottom-left value rendering should stay scalar so cards and aggregations remain predictable.
        return fieldType === 'scalar';
    }
    
    canUseForFilter(fieldName) {
        const fieldType = this.fieldTypes.get(fieldName);
        if (fieldType !== 'scalar' && fieldType !== 'multi-value') return false;
        const kind = this.getFieldKind(fieldName);
        if (kind === 'relationship' && !this.showRelationshipFields) return false;
        if (kind === 'derived' && !this.showDerivedFields) return false;
        return true;
    }
    
    // ===== VIEW CONFIGURATION MANAGEMENT =====
    
    getDefaultViewConfig() {
        return {
            axisSelections: {
                x: null,
                y: null
            },
            cardSelections: {
                title: null,
                topLeft: null,
                topRight: null,
                bottomRight: null
            },
            tableColumns: [],
            slicerFilters: {},
            headerOrdering: {
                rows: [],
                columns: []
            },
            tagCustomizations: {},
            cardClick: null,
            groups: [],
            groupSlicerFilter: null,
            cellRenderMode: 'cards',
            tableSummaryMode: 'none'
        };
    }

    normalizeTableSummaryMode(mode) {
        return ['none', 'sum', 'average', 'min', 'max'].includes(mode) ? mode : 'none';
    }

    normalizeViewConfig(viewConfig) {
        const defaultViewConfig = this.getDefaultViewConfig();
        const axisSelections = viewConfig && viewConfig.axisSelections && typeof viewConfig.axisSelections === 'object'
            ? this.cloneItem(viewConfig.axisSelections)
            : {};
        const legacyCardSelections = this.getLegacyCardSelections(axisSelections);
        const normalizedViewConfig = viewConfig && typeof viewConfig === 'object'
            ? {
                ...defaultViewConfig,
                ...this.cloneItem(viewConfig),
                axisSelections: {
                    ...defaultViewConfig.axisSelections,
                    x: axisSelections.x || null,
                    y: axisSelections.y || null
                },
                cardSelections: {
                    ...defaultViewConfig.cardSelections,
                    ...legacyCardSelections,
                    ...(viewConfig.cardSelections && typeof viewConfig.cardSelections === 'object'
                        ? this.cloneItem(viewConfig.cardSelections)
                        : {})
                },
                slicerFilters: this.cloneItem(
                    viewConfig.slicerFilters && typeof viewConfig.slicerFilters === 'object'
                        ? viewConfig.slicerFilters
                        : viewConfig.filters && typeof viewConfig.filters === 'object'
                            ? viewConfig.filters
                            : defaultViewConfig.slicerFilters
                ),
                tagCustomizations: this.cloneItem(
                    viewConfig.tagCustomizations && typeof viewConfig.tagCustomizations === 'object'
                        ? viewConfig.tagCustomizations
                        : viewConfig.tags && typeof viewConfig.tags === 'object'
                            ? viewConfig.tags
                            : defaultViewConfig.tagCustomizations
                ),
                groups: Array.isArray(viewConfig.groups)
                    ? JSON.parse(JSON.stringify(viewConfig.groups))
                    : [],
                tableColumns: this.normalizeFieldList(viewConfig.tableColumns),
                cellRenderMode: viewConfig.cellRenderMode === 'table' ? 'table' : defaultViewConfig.cellRenderMode,
                tableSummaryMode: this.normalizeTableSummaryMode(viewConfig.tableSummaryMode)
            }
            : this.getDefaultViewConfig();

        const hasAxisSelections = Object.values(normalizedViewConfig.axisSelections || {}).some((value) => Boolean(value));
        if (!hasAxisSelections) {
            normalizedViewConfig.axisSelections = this.getDefaultGridAxisSelections();
        }

        const hasCardSelections = Object.values(normalizedViewConfig.cardSelections || {}).some((value) => Boolean(value));
        if (!hasCardSelections) {
            normalizedViewConfig.cardSelections = this.getDefaultCardSelections(normalizedViewConfig.axisSelections);
        }

        if (normalizedViewConfig.tableColumns.length === 0) {
            normalizedViewConfig.tableColumns = this.getDefaultTableColumnFields(normalizedViewConfig.cardSelections);
        }

        return normalizedViewConfig;
    }

    getLegacyCardSelections(axisSelections = {}) {
        return {
            title: axisSelections.title || null,
            topLeft: axisSelections.topLeft || null,
            topRight: axisSelections.topRight || null,
            bottomRight: axisSelections.bottomRight || null
        };
    }

    normalizeFieldList(fieldNames) {
        const normalizedFields = [];
        const seenFields = new Set();

        (Array.isArray(fieldNames) ? fieldNames : []).forEach((fieldName) => {
            if (typeof fieldName !== 'string') {
                return;
            }

            const normalizedFieldName = fieldName.trim();
            if (!normalizedFieldName || seenFields.has(normalizedFieldName)) {
                return;
            }

            seenFields.add(normalizedFieldName);
            normalizedFields.push(normalizedFieldName);
        });

        return normalizedFields;
    }

    getDefaultGridAxisSelections() {
        const scalarFields = this.availableFields.filter((fieldName) => this.fieldTypes.get(fieldName) === 'scalar');
        const unusedFields = new Set(scalarFields);

        const pickField = ({ preferredNames = [], predicate = null } = {}) => {
            let selectedField = preferredNames.find((fieldName) => unusedFields.has(fieldName));

            if (!selectedField && predicate) {
                selectedField = scalarFields.find((fieldName) => unusedFields.has(fieldName) && predicate(fieldName));
            }

            if (!selectedField) {
                selectedField = scalarFields.find((fieldName) => unusedFields.has(fieldName));
            }

            if (!selectedField) {
                selectedField = preferredNames.find((fieldName) => scalarFields.includes(fieldName));
            }

            if (!selectedField && predicate) {
                selectedField = scalarFields.find((fieldName) => predicate(fieldName));
            }

            if (!selectedField) {
                selectedField = scalarFields[0] || null;
            }

            if (selectedField) {
                unusedFields.delete(selectedField);
            }

            return selectedField;
        };

        return {
            x: pickField({ preferredNames: ['status', 'priority', 'assignee', 'initiative', 'milestone', 'dueDate'] }),
            y: pickField({ preferredNames: ['initiative', 'assignee', 'priority', 'status', 'milestone', 'dueDate'] })
        };
    }

    getDefaultCardSelections(axisSelections = this.currentAxisSelections) {
        const scalarFields = this.availableFields.filter((fieldName) => this.fieldTypes.get(fieldName) === 'scalar');
        const unusedFields = new Set(scalarFields);
        const numericScalarFields = new Set(scalarFields.filter((fieldName) => this.isNumericField(fieldName)));

        if (axisSelections && axisSelections.x) {
            unusedFields.delete(axisSelections.x);
        }

        if (axisSelections && axisSelections.y) {
            unusedFields.delete(axisSelections.y);
        }

        const pickField = ({ preferredNames = [], predicate = null } = {}) => {
            let selectedField = preferredNames.find((fieldName) => unusedFields.has(fieldName));

            if (!selectedField && predicate) {
                selectedField = scalarFields.find((fieldName) => unusedFields.has(fieldName) && predicate(fieldName));
            }

            if (!selectedField) {
                selectedField = scalarFields.find((fieldName) => unusedFields.has(fieldName));
            }

            if (!selectedField) {
                selectedField = preferredNames.find((fieldName) => scalarFields.includes(fieldName));
            }

            if (!selectedField && predicate) {
                selectedField = scalarFields.find((fieldName) => predicate(fieldName));
            }

            if (!selectedField) {
                selectedField = scalarFields[0] || null;
            }

            if (selectedField) {
                unusedFields.delete(selectedField);
            }

            return selectedField;
        };

        return {
            title: pickField({ preferredNames: ['title', 'name', 'summary', 'subject', 'id'] }),
            topLeft: pickField({ preferredNames: ['id', 'assignee', 'owner', 'priority', 'status'] }),
            topRight: pickField({ preferredNames: ['priority', 'status', 'initiative', 'assignee', 'dueDate'] }),
            bottomRight: pickField({
                preferredNames: ['effort', 'estimate', 'storyPoints', 'points', 'size', 'score'],
                predicate: (fieldName) => numericScalarFields.has(fieldName)
            })
        };
    }

    getDefaultTableColumnFields(cardSelections = this.currentCardSelections) {
        const candidateFields = this.normalizeFieldList([
            cardSelections && cardSelections.title,
            cardSelections && cardSelections.topLeft,
            cardSelections && cardSelections.topRight,
            cardSelections && cardSelections.bottomRight
        ]);

        if (candidateFields.length > 0) {
            return candidateFields;
        }

        return this.normalizeFieldList(this.availableFields.filter((fieldName) => this.canUseForTableColumn(fieldName)).slice(0, 4));
    }

    canUseForTableColumn(fieldName) {
        const fieldType = this.fieldTypes.get(fieldName);
        return fieldType === 'scalar' || fieldType === 'multi-value';
    }
    
    applyViewConfiguration() {
        this.viewConfig = this.normalizeViewConfig(this.viewConfig);

        if (this.viewConfig.axisSelections) {
            this.currentAxisSelections = { ...this.viewConfig.axisSelections };
        }

        if (this.viewConfig.cardSelections) {
            this.currentCardSelections = { ...this.viewConfig.cardSelections };
        }

        this.tableColumnFields = this.normalizeFieldList(this.viewConfig.tableColumns);
        if (this.tableColumnFields.length === 0) {
            this.tableColumnFields = this.getDefaultTableColumnFields(this.currentCardSelections);
        }

        this.cellRenderMode = this.viewConfig.cellRenderMode === 'table' ? 'table' : 'cards';
        this.tableSummaryMode = this.normalizeTableSummaryMode(this.viewConfig.tableSummaryMode);
        
        if (this.viewConfig.slicerFilters) {
            this.currentFilters.clear();
            Object.entries(this.viewConfig.slicerFilters).forEach(([field, values]) => {
                this.currentFilters.set(field, new Set(values));
            });
        }
        
        if (this.viewConfig.tagCustomizations) {
            this.tagCustomizations.clear();
            Object.entries(this.viewConfig.tagCustomizations).forEach(([tag, config]) => {
                this.tagCustomizations.set(tag, config);
            });
        }
    }
    
    updateViewConfiguration() {
        this.viewConfig.axisSelections = { ...this.currentAxisSelections };
        this.viewConfig.cardSelections = { ...this.currentCardSelections };
        this.viewConfig.tableColumns = this.normalizeFieldList(this.tableColumnFields);
        this.viewConfig.cellRenderMode = this.cellRenderMode === 'table' ? 'table' : 'cards';
        this.viewConfig.tableSummaryMode = this.normalizeTableSummaryMode(this.tableSummaryMode);
        
        // Convert filters to serializable format
        this.viewConfig.slicerFilters = {};
        this.currentFilters.forEach((valueSet, field) => {
            this.viewConfig.slicerFilters[field] = Array.from(valueSet);
        });
        
        // Convert tag customizations to serializable format
        this.viewConfig.tagCustomizations = {};
        this.tagCustomizations.forEach((config, tag) => {
            this.viewConfig.tagCustomizations[tag] = config;
        });

        // Groups are already stored as a plain array in viewConfig
        if (!Array.isArray(this.viewConfig.groups)) {
            this.viewConfig.groups = [];
        }
    }

    getViewConfig() {
        const serializedFilters = {};
        this.currentFilters.forEach((valueSet, field) => {
            serializedFilters[field] = Array.from(valueSet instanceof Set ? valueSet : []);
        });

        const serializedTags = {};
        this.tagCustomizations.forEach((config, tag) => {
            serializedTags[tag] = config;
        });

        return {
            axisSelections: { ...this.currentAxisSelections },
            cardSelections: { ...this.currentCardSelections },
            tableColumns: this.normalizeFieldList(this.tableColumnFields),
            filters: serializedFilters,
            slicerFilters: serializedFilters,
            tagCustomizations: serializedTags,
            tags: serializedTags,
            cardClick: this.getViewCardClickConfig(),
            groups: Array.isArray(this.viewConfig.groups) ? JSON.parse(JSON.stringify(this.viewConfig.groups)) : [],
            groupSlicerFilter: this.viewConfig.groupSlicerFilter || null,
            cellRenderMode: this.cellRenderMode === 'table' ? 'table' : 'cards',
            tableSummaryMode: this.normalizeTableSummaryMode(this.tableSummaryMode)
        };
    }

    getDatasetCardClickConfig() {
        if (!this.metaInfo || !this.metaInfo.cardClick) {
            return null;
        }

        return JSON.parse(JSON.stringify(this.metaInfo.cardClick));
    }

    setDatasetCardClickConfig(config) {
        if (!this.metaInfo) {
            this.metaInfo = {};
        }

        this.metaInfo.cardClick = config
            ? JSON.parse(JSON.stringify(config))
            : null;
    }

    getViewCardClickConfig() {
        if (!this.viewConfig || !this.viewConfig.cardClick) {
            return null;
        }

        return JSON.parse(JSON.stringify(this.viewConfig.cardClick));
    }

    setViewCardClickConfig(config) {
        if (!this.viewConfig) {
            this.viewConfig = this.getDefaultViewConfig();
        }

        this.viewConfig.cardClick = config
            ? JSON.parse(JSON.stringify(config))
            : null;
    }
    
    // ===== DATA FILTERING =====

    applyGroupSlicerFilter(sourceData = []) {
        const groupFilter = this.viewConfig && Array.isArray(this.viewConfig.groupSlicerFilter)
            ? this.viewConfig.groupSlicerFilter
            : null;

        if (!Array.isArray(sourceData)) {
            return [];
        }

        if (!groupFilter) {
            return sourceData;
        }

        if (groupFilter.length === 0) {
            return [];
        }

        const allowedGroupIds = new Set(groupFilter);
        const allowNone = allowedGroupIds.has('_none');
        return sourceData.filter((item) => {
            const itemGroups = this.getItemGroups(item);
            if (itemGroups.length === 0) {
                return allowNone;
            }

            return itemGroups.some((group) => allowedGroupIds.has(group.id));
        });
    }

    computeFilteredDataset() {
        const filterDebug = [];
        const baseData = Array.isArray(this.dataset) ? this.dataset : [];

        const slicerFilteredItems = baseData.filter(item => {
            // Apply all active filters
            for (let [fieldName, allowedValues] of this.currentFilters) {
                const normalizedAllowedValues = allowedValues instanceof Set
                    ? allowedValues
                    : new Set(Array.isArray(allowedValues) ? allowedValues : []);

                if (normalizedAllowedValues !== allowedValues) {
                    this.currentFilters.set(fieldName, normalizedAllowedValues);
                }

                if (!this.fieldTypes.has(fieldName)) {
                    filterDebug.push({
                        itemId: item.id,
                        fieldName,
                        itemValue: undefined,
                        allowedValues: Array.from(normalizedAllowedValues),
                        reason: 'filter field missing from current dataset'
                    });
                    return false;
                }

                const distinctValues = this.distinctValues.get(fieldName) || [];
                if (distinctValues.length === 0) {
                    continue;
                }

                if (normalizedAllowedValues.size === 0) {
                    filterDebug.push({
                        itemId: item.id,
                        fieldName,
                        itemValue: item[fieldName],
                        allowedValues: [],
                        reason: 'explicit empty filter set'
                    });
                    return false;
                }

                const itemValue = item[fieldName];
                const fieldType = this.fieldTypes.get(fieldName);

                if (fieldType === 'multi-value' && Array.isArray(itemValue)) {
                    const hasMatch = itemValue.some(v => normalizedAllowedValues.has(String(v)));
                    if (!hasMatch) {
                        filterDebug.push({
                            itemId: item.id,
                            fieldName,
                            itemValue: itemValue,
                            allowedValues: Array.from(normalizedAllowedValues),
                            reason: 'no multi-value overlap'
                        });
                        return false;
                    }
                } else {
                    if (itemValue !== undefined && itemValue !== null) {
                        if (!normalizedAllowedValues.has(String(itemValue))) {
                            filterDebug.push({
                                itemId: item.id,
                                fieldName,
                                itemValue: itemValue,
                                allowedValues: Array.from(normalizedAllowedValues),
                                reason: 'scalar value excluded'
                            });
                            return false;
                        }
                    } else if (!normalizedAllowedValues.has('')) {
                        filterDebug.push({
                            itemId: item.id,
                            fieldName,
                            itemValue: '',
                            allowedValues: Array.from(normalizedAllowedValues),
                            reason: 'empty value excluded'
                        });
                        return false;
                    }
                }
            }

            return true;
        });

        return {
            items: this.applyGroupSlicerFilter(slicerFilteredItems),
            filterDebug,
            baseCount: baseData.length,
            slicerFilteredCount: slicerFilteredItems.length
        };
    }

    getFilteredData() {
        return this.computeFilteredDataset().items;
    }
    
    updateFilteredData() {
        const { items, filterDebug, baseCount } = this.computeFilteredDataset();
        this.filteredData = items;

        console.log(`Filtering: ${baseCount} total items, ${this.filteredData.length} after filtering`);

        if (this.filteredData.length === 0 && this.currentFilters.size > 0) {
            console.warn('Filtering eliminated all items', {
                filters: Array.from(this.currentFilters.entries()).map(([fieldName, allowedValues]) => ({
                    fieldName,
                    allowedValues: Array.from(allowedValues instanceof Set ? allowedValues : []),
                    fieldExists: this.fieldTypes.has(fieldName),
                    distinctValues: this.distinctValues.get(fieldName) || []
                })),
                sampleExclusions: filterDebug.slice(0, 10)
            });
        }
    }

    snapshotDistinctValues() {
        const snapshot = new Map();

        this.distinctValues.forEach((values, fieldName) => {
            snapshot.set(fieldName, [...values]);
        });

        return snapshot;
    }

    reconcileFiltersWithDistinctValues(previousDistinctValues = new Map()) {
        const reconciledFilters = new Map();

        this.currentFilters.forEach((allowedValues, fieldName) => {
            if (!this.fieldTypes.has(fieldName) || !this.canUseForFilter(fieldName)) {
                console.warn('Dropping stale filter for unavailable field', fieldName);
                return;
            }

            const currentDistinctValues = this.distinctValues.get(fieldName) || [];
            if (currentDistinctValues.length === 0) {
                console.warn('Dropping stale filter for field with no distinct values', fieldName);
                return;
            }

            const normalizedAllowedValues = allowedValues instanceof Set
                ? allowedValues
                : new Set(Array.isArray(allowedValues) ? allowedValues : []);

            if (normalizedAllowedValues.size === 0) {
                reconciledFilters.set(fieldName, normalizedAllowedValues);
                return;
            }

            const currentDistinctSet = new Set(currentDistinctValues);
            const selectedValues = Array.from(normalizedAllowedValues);
            const missingSelectedValues = selectedValues.filter((value) => !currentDistinctSet.has(value));
            const preserveMissingAxisValues = this.isPinnedAxisFilterField(fieldName) && missingSelectedValues.length > 0;

            if (preserveMissingAxisValues) {
                console.debug('[HEADER] preserving missing axis filter values during reconciliation', {
                    fieldName,
                    missingSelectedValues,
                    currentDistinctValues,
                    selectedValues
                });
                reconciledFilters.set(fieldName, new Set(selectedValues));
                return;
            }

            const previousValues = previousDistinctValues.get(fieldName) || [];
            const effectiveAllowedForAll = new Set(normalizedAllowedValues);
            const hadNoneSelected = effectiveAllowedForAll.delete('');
            const wasEffectivelyAll = previousValues.length > 0 &&
                effectiveAllowedForAll.size === previousValues.length &&
                previousValues.every((value) => effectiveAllowedForAll.has(value));

            if (wasEffectivelyAll) {
                const newFilter = new Set(currentDistinctValues);
                if (hadNoneSelected && this.hasItemsWithEmptyValue(fieldName)) {
                    newFilter.add('');
                }
                reconciledFilters.set(fieldName, newFilter);
                return;
            }

            // Auto-select values that are newly appeared (not in previous distinct set)
            const previousDistinctSet = new Set(previousValues);
            const newValues = currentDistinctValues.filter(v => !previousDistinctSet.has(v));
            const kept = Array.from(normalizedAllowedValues).filter(v => {
                if (v === '') return this.hasItemsWithEmptyValue(fieldName);
                return currentDistinctSet.has(v);
            });
            reconciledFilters.set(fieldName, new Set([...kept, ...newValues]));
        });

        this.currentFilters = reconciledFilters;
        this.updateFilteredData();
    }

    isPinnedAxisFilterField(fieldName) {
        return fieldName === this.currentAxisSelections.x || fieldName === this.currentAxisSelections.y;
    }

    isFilterEquivalentToAll(fieldName, allowedValues) {
        const normalizedAllowedValues = allowedValues instanceof Set
            ? allowedValues
            : new Set(Array.isArray(allowedValues) ? allowedValues : []);
        const distinctValues = this.distinctValues.get(fieldName) || [];

        if (distinctValues.length === 0) return false;
        if (!distinctValues.every((value) => normalizedAllowedValues.has(value))) return false;
        if (this.hasItemsWithEmptyValue(fieldName) && !normalizedAllowedValues.has('')) return false;
        return true;
    }

    hasItemsWithEmptyValue(fieldName) {
        return this.dataset.some(item => {
            const value = item[fieldName];
            return value === undefined || value === null || value === '';
        });
    }
    
    setFilter(fieldName, selectedValues) {
        this.currentFilters.set(fieldName, new Set(selectedValues));
        this.updateFilteredData();
        this.renderGrid();
    }
    
    // ===== FIELD VALUE UTILITIES =====
    
    getFieldValue(item, fieldName) {
        if (!item || !fieldName) {
            return '';
        }

        const value = item[fieldName];
        // Missing field values are treated as empty according to design spec
        return value !== undefined && value !== null ? value : '';
    }
    
    getDisplayValue(item, fieldName) {
        const value = this.getFieldValue(item, fieldName);
        const fieldType = this.fieldTypes.get(fieldName);
        
        if (fieldType === 'multi-value' && Array.isArray(value)) {
            return value.join(', ');
        } else if (typeof value === 'object' && value !== null) {
            // For structured fields, try to find a reasonable display value
            return value.toString();
        } else {
            return String(value);
        }
    }

    formatFixedDecimalValue(value, fractionDigits = 2) {
        const parsedValue = this.parseNumericValue(value);
        if (parsedValue === null) {
            return value === undefined || value === null ? '' : String(value);
        }

        return parsedValue.toLocaleString(undefined, {
            minimumFractionDigits: fractionDigits,
            maximumFractionDigits: fractionDigits
        });
    }

    getSummaryValueFromNumericValues(values, mode, fallbackMode = 'none') {
        if (!Array.isArray(values) || values.length === 0) {
            return null;
        }

        const normalizedMode = this.normalizeTableSummaryMode(mode);
        const resolvedMode = normalizedMode === 'none'
            ? this.normalizeTableSummaryMode(fallbackMode)
            : normalizedMode;

        if (resolvedMode === 'sum') {
            return values.reduce((total, value) => total + value, 0);
        }

        if (resolvedMode === 'average') {
            return values.reduce((total, value) => total + value, 0) / values.length;
        }

        if (resolvedMode === 'min') {
            return Math.min(...values);
        }

        if (resolvedMode === 'max') {
            return Math.max(...values);
        }

        return null;
    }
    
    isNumericField(fieldName) {
        const distinctValues = this.distinctValues.get(fieldName) || [];
        return distinctValues.length > 0 &&
            distinctValues.every((value) => this.parseNumericValue(value) !== null);
    }

    parseNumericValue(value) {
        if (typeof value === 'number') {
            return isFinite(value) ? value : null;
        }

        if (typeof value === 'string') {
            const trimmed = value.trim();
            if (!trimmed) {
                return null;
            }

            const normalized = trimmed.replace(/,/g, '');
            if (!/^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(normalized)) {
                return null;
            }

            const parsed = Number(normalized);
            return isFinite(parsed) ? parsed : null;
        }

        return null;
    }

    formatNumericSummaryValue(value) {
        if (!isFinite(value)) {
            return '';
        }

        return this.formatFixedDecimalValue(value, 2);
    }

    parseSeriesDateValue(value) {
        if (value instanceof Date) {
            return Number.isNaN(value.getTime()) ? null : value;
        }

        if (typeof value !== 'string') {
            return null;
        }

        const trimmed = value.trim();
        if (!trimmed) {
            return null;
        }

        const looksLikeDate =
            /^\d{4}-\d{2}-\d{2}(?:[T\s].*)?$/.test(trimmed) ||
            /^\d{4}\/\d{1,2}\/\d{1,2}$/.test(trimmed) ||
            /^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(trimmed);

        if (!looksLikeDate) {
            return null;
        }

        const parsed = new Date(trimmed);
        return Number.isNaN(parsed.getTime()) ? null : parsed;
    }

    getGraphCandidateFields(items, preferredFields = []) {
        const fieldSet = new Set((Array.isArray(preferredFields) ? preferredFields : []).filter(Boolean));

        if (Array.isArray(items)) {
            items.forEach((item) => {
                Object.keys(item || {}).forEach((fieldName) => {
                    if (!fieldName.startsWith('_')) {
                        fieldSet.add(fieldName);
                    }
                });
            });
        }

        return Array.from(fieldSet);
    }

    inferTimeSeriesDateField(items, candidateFields = []) {
        const rankedFields = this.getGraphCandidateFields(items, candidateFields)
            .map((fieldName) => {
                const nonEmptyValues = items
                    .map((item) => this.getFieldValue(item, fieldName))
                    .filter((value) => value !== undefined && value !== null && value !== '');

                if (nonEmptyValues.length < 2) {
                    return null;
                }

                const parsedValues = nonEmptyValues
                    .map((value) => this.parseSeriesDateValue(value))
                    .filter((value) => value !== null);

                if (parsedValues.length < 2) {
                    return null;
                }

                const coverage = parsedValues.length / nonEmptyValues.length;
                if (coverage < 0.8) {
                    return null;
                }

                let score = coverage;
                const lowerField = fieldName.toLowerCase();
                if (/(^date$|date|time|timestamp|created|updated|posted|occurred|start|end|due)/.test(lowerField)) {
                    score += 4;
                }
                if (this.isNumericField(fieldName)) {
                    score -= 3;
                }

                return { fieldName, score };
            })
            .filter(Boolean)
            .sort((left, right) => right.score - left.score || left.fieldName.localeCompare(right.fieldName));

        return rankedFields.length > 0 ? rankedFields[0].fieldName : null;
    }

    inferTimeSeriesValueField(items, candidateFields = [], excludedField = null) {
        const rankedFields = this.getGraphCandidateFields(items, candidateFields)
            .filter((fieldName) => fieldName !== excludedField)
            .map((fieldName) => {
                const nonEmptyValues = items
                    .map((item) => this.getFieldValue(item, fieldName))
                    .filter((value) => value !== undefined && value !== null && value !== '');

                if (nonEmptyValues.length === 0) {
                    return null;
                }

                const numericValues = nonEmptyValues
                    .map((value) => this.parseNumericValue(value))
                    .filter((value) => value !== null);

                if (numericValues.length === 0) {
                    return null;
                }

                const coverage = numericValues.length / nonEmptyValues.length;
                if (coverage < 0.8) {
                    return null;
                }

                let score = coverage;
                const lowerField = fieldName.toLowerCase();
                if (/(amount|value|total|sum|price|cost|balance|score|count|qty|quantity|size)/.test(lowerField)) {
                    score += 4;
                }
                if (/(year|month|week|day|hour|minute|second)/.test(lowerField)) {
                    score -= 2;
                }

                return { fieldName, score };
            })
            .filter(Boolean)
            .sort((left, right) => right.score - left.score || left.fieldName.localeCompare(right.fieldName));

        return rankedFields.length > 0 ? rankedFields[0].fieldName : null;
    }

    buildTimeSeriesGraphContext(items, preferredFields = [], contextLabel = '') {
        if (!Array.isArray(items) || items.length === 0) {
            return null;
        }

        const candidateFields = this.getGraphCandidateFields(items, preferredFields);
        const dateField = this.inferTimeSeriesDateField(items, candidateFields);
        if (!dateField) {
            return null;
        }

        const valueField = this.inferTimeSeriesValueField(items, candidateFields, dateField);
        if (!valueField) {
            return null;
        }

        const graphSummaryMode = this.normalizeTableSummaryMode(this.tableSummaryMode);
        const groupedPoints = new Map();

        items.forEach((item, index) => {
            const parsedDate = this.parseSeriesDateValue(this.getFieldValue(item, dateField));
            const parsedValue = this.parseNumericValue(this.getFieldValue(item, valueField));
            if (!parsedDate || parsedValue === null) {
                return;
            }

            const itemId = this.getItemIdentity(item) || item.id || `row-${index + 1}`;
            const labelField = this.currentCardSelections.title || 'Description';
            const labelValue = this.getFieldValue(item, labelField) || this.getFieldValue(item, 'title') || itemId;
            const normalizedDate = new Date(parsedDate.getFullYear(), parsedDate.getMonth(), parsedDate.getDate());
            const dateKey = normalizedDate.toISOString();

            if (!groupedPoints.has(dateKey)) {
                groupedPoints.set(dateKey, {
                    timestamp: normalizedDate.getTime(),
                    isoDate: dateKey,
                    dateLabel: normalizedDate.toLocaleDateString(),
                    itemIds: [],
                    labels: [],
                    values: []
                });
            }

            const pointGroup = groupedPoints.get(dateKey);
            pointGroup.itemIds.push(itemId);
            pointGroup.labels.push(String(labelValue));
            pointGroup.values.push(parsedValue);
        });

        const points = Array.from(groupedPoints.values())
            .map((pointGroup) => {
                const aggregatedValue = this.getSummaryValueFromNumericValues(
                    pointGroup.values,
                    graphSummaryMode,
                    'sum'
                );
                if (aggregatedValue === null) {
                    return null;
                }

                return {
                    itemId: pointGroup.itemIds[0] || pointGroup.isoDate,
                    itemIds: pointGroup.itemIds,
                    label: pointGroup.values.length > 1
                        ? `${pointGroup.values.length} rows`
                        : (pointGroup.labels[0] || pointGroup.itemIds[0] || pointGroup.isoDate),
                    timestamp: pointGroup.timestamp,
                    isoDate: pointGroup.isoDate,
                    dateLabel: pointGroup.dateLabel,
                    value: aggregatedValue,
                    valueLabel: this.formatFixedDecimalValue(aggregatedValue, 2),
                    sourceCount: pointGroup.values.length
                };
            })
            .filter(Boolean)
            .sort((left, right) => left.timestamp - right.timestamp || left.itemId.localeCompare(right.itemId));

        if (points.length === 0) {
            return null;
        }

        return {
            contextLabel,
            filterLabel: this.getActiveFilterContextLabel(),
            dateField,
            valueField,
            pointCount: points.length,
            points
        };
    }

    getActiveFilterContextLabel() {
        const activeFilters = [];

        this.currentFilters.forEach((allowedValues, fieldName) => {
            if (this.isFilterEquivalentToAll(fieldName, allowedValues)) {
                return;
            }

            const normalizedAllowedValues = allowedValues instanceof Set
                ? Array.from(allowedValues)
                : (Array.isArray(allowedValues) ? allowedValues : []);

            if (normalizedAllowedValues.length === 0) {
                return;
            }

            const displayValues = normalizedAllowedValues
                .map((value) => value === '' ? '(none)' : String(value))
                .sort((left, right) => left.localeCompare(right));

            activeFilters.push(`${this.formatFieldLabel(fieldName)}: ${displayValues.join(', ')}`);
        });

        return activeFilters.join('; ');
    }

    formatTableGraphContextLabel(rowValue, colValue) {
        const labels = [];
        const rowField = this.currentAxisSelections.y;
        const colField = this.currentAxisSelections.x;

        if (rowField) {
            labels.push(`${this.formatFieldLabel(rowField)}: ${rowValue === '' ? '(empty)' : rowValue}`);
        }

        if (colField) {
            labels.push(`${this.formatFieldLabel(colField)}: ${colValue === '' ? '(empty)' : colValue}`);
        }

        const filterLabel = this.getActiveFilterContextLabel();
        if (filterLabel) {
            labels.push(`Filters: ${filterLabel}`);
        }

        return labels.join(' | ');
    }

    stripFilterSuffixFromTableGraphLabel(contextLabel) {
        if (typeof contextLabel !== 'string' || !contextLabel) {
            return '';
        }

        const filterMarker = ' | Filters:';
        const markerIndex = contextLabel.indexOf(filterMarker);
        return markerIndex >= 0
            ? contextLabel.slice(0, markerIndex)
            : contextLabel;
    }

    createTableGraphSelectionKey(contextLabel, itemIds = [], fields = []) {
        const normalizedItemIds = Array.from(new Set(
            (Array.isArray(itemIds) ? itemIds : [itemIds])
                .map((itemId) => itemId === undefined || itemId === null ? null : String(itemId))
                .filter(Boolean)
        ));
        const normalizedFields = this.normalizeFieldList(fields);
        const baseLabel = this.stripFilterSuffixFromTableGraphLabel(contextLabel);

        return `${baseLabel}::${normalizedFields.join('|')}::${normalizedItemIds.join('|')}`;
    }

    isTableGraphSelected(tableGraphContext) {
        return Boolean(tableGraphContext && tableGraphContext.selectionKey) &&
            this.selectedTableGraphContexts.has(tableGraphContext.selectionKey);
    }

    toggleTableGraphSelection(tableGraphContext) {
        if (!tableGraphContext || !tableGraphContext.selectionKey) {
            return false;
        }

        if (this.selectedTableGraphContexts.has(tableGraphContext.selectionKey)) {
            this.selectedTableGraphContexts.delete(tableGraphContext.selectionKey);
            this.applyTableGraphSelectionClasses();
            return false;
        }

        this.selectedTableGraphContexts.set(tableGraphContext.selectionKey, {
            ...tableGraphContext,
            itemIds: Array.isArray(tableGraphContext.itemIds) ? [...tableGraphContext.itemIds] : [],
            fields: Array.isArray(tableGraphContext.fields) ? [...tableGraphContext.fields] : []
        });
        this.applyTableGraphSelectionClasses();
        return true;
    }

    clearTableGraphSelection(options = {}) {
        const { suppressRender = false } = options;
        this.selectedTableGraphContexts.clear();

        if (!suppressRender) {
            this.applyTableGraphSelectionClasses();
        }
    }

    getSelectedTableGraphContexts() {
        return Array.from(this.selectedTableGraphContexts.values());
    }

    getTableGraphContextsForOpen(tableGraphContext) {
        if (!tableGraphContext) {
            return [];
        }

        const selectedContexts = this.getSelectedTableGraphContexts();
        if (selectedContexts.length > 1 && this.isTableGraphSelected(tableGraphContext)) {
            return selectedContexts;
        }

        return [tableGraphContext];
    }

    getCombinedTableGraphFields(tableGraphContexts = []) {
        const combinedFields = [];
        const seenFields = new Set();

        tableGraphContexts.forEach((context) => {
            this.normalizeFieldList(context && context.fields).forEach((fieldName) => {
                if (seenFields.has(fieldName)) {
                    return;
                }

                seenFields.add(fieldName);
                combinedFields.push(fieldName);
            });
        });

        return combinedFields;
    }

    buildCombinedTableGraphContextLabel(tableGraphContexts = []) {
        if (!Array.isArray(tableGraphContexts) || tableGraphContexts.length === 0) {
            return '';
        }

        if (tableGraphContexts.length === 1) {
            return tableGraphContexts[0].contextLabel || '';
        }

        const baseLabels = Array.from(new Set(
            tableGraphContexts
                .map((context) => this.stripFilterSuffixFromTableGraphLabel(context.contextLabel || ''))
                .filter(Boolean)
        ));

        if (baseLabels.length === 0) {
            return `Combined Tables (${tableGraphContexts.length})`;
        }

        const visibleLabels = baseLabels.slice(0, 3);
        const overflowCount = baseLabels.length - visibleLabels.length;
        const suffix = overflowCount > 0 ? ` +${overflowCount} more` : '';
        return `Combined Tables (${tableGraphContexts.length}): ${visibleLabels.join(' || ')}${suffix}`;
    }

    buildTableExportSheetName(tableGraphContext, fallbackIndex = 1) {
        const baseLabel = this.stripFilterSuffixFromTableGraphLabel(tableGraphContext && tableGraphContext.contextLabel || '');
        const candidate = (baseLabel || `Table ${fallbackIndex}`)
            .replace(/[\\/*?:\[\]]/g, ' ')
            .replace(/\s+/g, ' ')
            .trim();

        return candidate.slice(0, 31) || `Table ${fallbackIndex}`;
    }

    buildTableExportRows(items = [], fields = [], extraColumns = {}) {
        return (Array.isArray(items) ? items : []).map((item) => {
            const row = { ...extraColumns };

            this.normalizeFieldList(fields).forEach((fieldName) => {
                const rawValue = this.getFieldValue(item, fieldName);
                if (Array.isArray(rawValue)) {
                    row[fieldName] = rawValue.join(', ');
                } else if (rawValue && typeof rawValue === 'object') {
                    row[fieldName] = JSON.stringify(rawValue);
                } else {
                    row[fieldName] = rawValue;
                }
            });

            return row;
        });
    }

    async exportTableGraphContextsToExcel(tableGraphContexts = []) {
        if (!window.XLSX) {
            this.showNotification('Excel export library is not available', 'error');
            return false;
        }

        const normalizedContexts = Array.isArray(tableGraphContexts)
            ? tableGraphContexts.filter(Boolean)
            : [];
        if (normalizedContexts.length === 0) {
            this.showNotification('No table data available for Excel export', 'warning');
            return false;
        }

        const workbook = window.XLSX.utils.book_new();
        const seenSheetNames = new Set();
        const combinedRows = [];

        normalizedContexts.forEach((context, index) => {
            const itemIds = Array.isArray(context.itemIds) ? context.itemIds : [];
            const items = itemIds
                .map((itemId) => this.getEffectiveItemById(itemId))
                .filter(Boolean);
            if (items.length === 0) {
                return;
            }

            const rows = this.buildTableExportRows(items, context.fields || []);
            if (rows.length === 0) {
                return;
            }

            let sheetName = this.buildTableExportSheetName(context, index + 1);
            let dedupeIndex = 2;
            while (seenSheetNames.has(sheetName)) {
                const suffix = ` ${dedupeIndex}`;
                sheetName = `${sheetName.slice(0, Math.max(0, 31 - suffix.length))}${suffix}`;
                dedupeIndex += 1;
            }
            seenSheetNames.add(sheetName);

            const worksheet = window.XLSX.utils.json_to_sheet(rows);
            window.XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);

            rows.forEach((row) => {
                combinedRows.push({
                    __table: sheetName,
                    ...row
                });
            });
        });

        if (combinedRows.length === 0) {
            this.showNotification('No current table rows could be resolved for Excel export', 'warning');
            return false;
        }

        if (normalizedContexts.length > 1) {
            const combinedSheet = window.XLSX.utils.json_to_sheet(combinedRows);
            window.XLSX.utils.book_append_sheet(workbook, combinedSheet, 'Combined');
        }

        const workbookArray = window.XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob(
            [workbookArray],
            { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }
        );
        const filename = `table-export-${window.GridDateUtils.createFilenameTimestamp()}.xlsx`;
        const savedFilename = await this.persistenceManager.fileService.saveBlob(blob, filename, {
            pickerType: 'data',
            fileType: 'xlsx',
            mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });

        if (!savedFilename) {
            return false;
        }

        this.showNotification(
            normalizedContexts.length > 1
                ? `Exported ${normalizedContexts.length} selected tables to Excel`
                : 'Exported table to Excel',
            'success'
        );
        return true;
    }

    applyTableGraphSelectionClasses() {
        document.querySelectorAll('.grid-cell-table[data-table-graph-key]').forEach((wrapper) => {
            const selectionKey = wrapper.dataset.tableGraphKey;
            wrapper.classList.toggle('table-graph-selected', Boolean(selectionKey) && this.selectedTableGraphContexts.has(selectionKey));
        });
    }

    openTableGraphWindow(items, preferredFields = [], options = {}) {
        const contextLabel = options && typeof options.contextLabel === 'string'
            ? options.contextLabel
            : '';
        const graphContext = this.buildTimeSeriesGraphContext(items, preferredFields, contextLabel);

        if (!graphContext) {
            this.showNotification('No time-series graph could be inferred from this table', 'warning');
            return false;
        }

        const graphWindow = window.open('', '_blank', 'width=1100,height=760');
        if (!graphWindow) {
            this.showNotification('Popup blocked while opening graph window', 'warning');
            return false;
        }

        const payload = JSON.stringify(graphContext).replace(/</g, '\\u003c');
        const title = graphContext.contextLabel
            ? `Time Series Graph - ${graphContext.contextLabel}`
            : 'Time Series Graph';

        graphWindow.document.open();
        graphWindow.document.write(`<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>${title}</title>
    <style>
        :root {
            color-scheme: light;
            font-family: 'Segoe UI', Tahoma, sans-serif;
        }
        body {
            margin: 0;
            padding: 24px;
            background: #f4f6f8;
            color: #1f2937;
        }
        .shell {
            max-width: 1100px;
            margin: 0 auto;
            background: #ffffff;
            border: 1px solid #d0d7de;
            border-radius: 16px;
            box-shadow: 0 12px 32px rgba(15, 23, 42, 0.12);
            overflow: hidden;
        }
        .header {
            padding: 24px 28px 12px;
            border-bottom: 1px solid #e5e7eb;
            background: linear-gradient(180deg, #ffffff 0%, #f8fafc 100%);
        }
        .title {
            margin: 0;
            font-size: 24px;
            font-weight: 700;
        }
        .subtitle {
            margin: 8px 0 0;
            color: #475467;
            font-size: 14px;
        }
        .meta {
            display: flex;
            gap: 16px;
            flex-wrap: wrap;
            padding: 16px 28px 0;
            color: #344054;
            font-size: 13px;
        }
        .toolbar {
            padding: 12px 28px 0;
            display: flex;
            justify-content: flex-end;
            gap: 10px;
        }
        .toolbar button {
            border: 1px solid #cbd5e1;
            background: #ffffff;
            color: #1f2937;
            border-radius: 999px;
            padding: 8px 14px;
            font-size: 13px;
            cursor: pointer;
        }
        .toolbar button.is-active {
            background: #1d4ed8;
            border-color: #1d4ed8;
            color: #ffffff;
        }
        .meta strong {
            color: #111827;
        }
        .chart-wrap {
            padding: 16px 20px 8px;
        }
        svg {
            width: 100%;
            height: auto;
            display: block;
            background: #fff;
        }
        .table-wrap {
            padding: 0 28px 28px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
        }
        th, td {
            text-align: left;
            padding: 10px 12px;
            border-bottom: 1px solid #e5e7eb;
        }
        th {
            font-size: 12px;
            text-transform: uppercase;
            letter-spacing: 0.04em;
            color: #667085;
            background: #f8fafc;
        }
        td:last-child, th:last-child {
            text-align: right;
        }
        .empty {
            padding: 36px 28px;
            color: #667085;
        }
    </style>
</head>
<body>
    <div class="shell">
        <div class="header">
            <h1 class="title">Time Series Graph</h1>
            <p class="subtitle" id="context-label"></p>
        </div>
        <div class="meta">
            <div><strong>Date field:</strong> <span id="date-field"></span></div>
            <div><strong>Value field:</strong> <span id="value-field"></span></div>
            <div><strong>Points:</strong> <span id="point-count"></span></div>
        </div>
        <div class="toolbar">
            <button type="button" id="flip-values-btn">Flip Values</button>
            <button type="button" id="accumulate-values-btn">Show Accumulated</button>
        </div>
        <div class="chart-wrap">
            <svg id="chart" viewBox="0 0 1000 420" aria-label="Time series chart"></svg>
        </div>
        <div class="table-wrap">
            <table>
                <thead>
                    <tr>
                        <th>Date</th>
                        <th>Item</th>
                        <th id="value-column-header"></th>
                    </tr>
                </thead>
                <tbody id="points-body"></tbody>
            </table>
            <div id="empty-state" class="empty" hidden>No plottable points were found.</div>
        </div>
    </div>
    <script>
        const graphData = ${payload};
        const svgNamespace = 'http://www.w3.org/2000/svg';
        let invertValues = false;
        let showAccumulated = false;

        const createSvgNode = (tagName, attributes = {}) => {
            const node = document.createElementNS(svgNamespace, tagName);
            Object.entries(attributes).forEach(([name, value]) => {
                node.setAttribute(name, value);
            });
            return node;
        };

        const formatValue = (value) => Number(value).toLocaleString(undefined, {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        });

        const getRenderedPoints = () => {
            let runningTotal = 0;

            return graphData.points.map((point) => {
                const signedValue = invertValues ? point.value * -1 : point.value;
                const renderedValue = showAccumulated
                    ? (runningTotal += signedValue)
                    : signedValue;

                return {
                    ...point,
                    renderedValue
                };
            });
        };

        const renderChart = () => {
            document.getElementById('context-label').textContent = graphData.contextLabel || 'Current table selection';
            document.getElementById('date-field').textContent = graphData.dateField;
            const displayedValueField = showAccumulated
                ? graphData.valueField + ' (Accumulated)'
                : graphData.valueField;
            document.getElementById('value-field').textContent = displayedValueField;
            document.getElementById('point-count').textContent = String(graphData.pointCount);
            document.getElementById('value-column-header').textContent = displayedValueField;
            const flipButton = document.getElementById('flip-values-btn');
            const accumulateButton = document.getElementById('accumulate-values-btn');
            flipButton.classList.toggle('is-active', invertValues);
            flipButton.textContent = invertValues ? 'Unflip Values' : 'Flip Values';
            accumulateButton.classList.toggle('is-active', showAccumulated);
            accumulateButton.textContent = showAccumulated ? 'Show Individual Values' : 'Show Accumulated';

            const pointsBody = document.getElementById('points-body');
            const emptyState = document.getElementById('empty-state');
            pointsBody.innerHTML = '';

            if (!Array.isArray(graphData.points) || graphData.points.length === 0) {
                emptyState.hidden = false;
                return;
            }

            const renderedPoints = getRenderedPoints();
            emptyState.hidden = true;
            renderedPoints.forEach((point) => {
                const row = document.createElement('tr');
                row.innerHTML = '<td>' + point.dateLabel + '</td><td>' + point.label + '</td><td>' + formatValue(point.renderedValue) + '</td>';
                pointsBody.appendChild(row);
            });

            const svg = document.getElementById('chart');
            svg.innerHTML = '';

            const width = 1000;
            const height = 420;
            const margin = { top: 20, right: 20, bottom: 54, left: 72 };
            const plotWidth = width - margin.left - margin.right;
            const plotHeight = height - margin.top - margin.bottom;
            const points = renderedPoints;
            const timestamps = points.map((point) => point.timestamp);
            const values = points.map((point) => point.renderedValue);
            const minX = Math.min(...timestamps);
            const maxX = Math.max(...timestamps);
            const minYRaw = Math.min(...values);
            const maxYRaw = Math.max(...values);
            const xSpan = Math.max(maxX - minX, 1);
            const yPadding = Math.max((maxYRaw - minYRaw) * 0.1, 1);
            const minY = minYRaw - yPadding;
            const maxY = maxYRaw + yPadding;
            const ySpan = Math.max(maxY - minY, 1);

            const chartGroup = createSvgNode('g');
            svg.appendChild(chartGroup);

            chartGroup.appendChild(createSvgNode('rect', {
                x: margin.left,
                y: margin.top,
                width: plotWidth,
                height: plotHeight,
                fill: '#ffffff'
            }));

            for (let index = 0; index <= 4; index += 1) {
                const value = minY + (ySpan * (index / 4));
                const y = margin.top + plotHeight - ((value - minY) / ySpan) * plotHeight;
                chartGroup.appendChild(createSvgNode('line', {
                    x1: margin.left,
                    y1: y,
                    x2: margin.left + plotWidth,
                    y2: y,
                    stroke: '#e5e7eb',
                    'stroke-width': '1'
                }));

                const label = createSvgNode('text', {
                    x: margin.left - 10,
                    y: y + 4,
                    'text-anchor': 'end',
                    fill: '#667085',
                    'font-size': '12'
                });
                label.textContent = formatValue(value);
                chartGroup.appendChild(label);
            }

            chartGroup.appendChild(createSvgNode('line', {
                x1: margin.left,
                y1: margin.top + plotHeight,
                x2: margin.left + plotWidth,
                y2: margin.top + plotHeight,
                stroke: '#344054',
                'stroke-width': '1.5'
            }));

            chartGroup.appendChild(createSvgNode('line', {
                x1: margin.left,
                y1: margin.top,
                x2: margin.left,
                y2: margin.top + plotHeight,
                stroke: '#344054',
                'stroke-width': '1.5'
            }));

            const projectX = (timestamp, index) => {
                if (points.length === 1) {
                    return margin.left + plotWidth / 2;
                }
                return margin.left + ((timestamp - minX) / xSpan) * plotWidth;
            };

            const projectY = (value) => margin.top + plotHeight - ((value - minY) / ySpan) * plotHeight;

            const polylinePoints = points
                .map((point, index) => projectX(point.timestamp, index) + ',' + projectY(point.renderedValue))
                .join(' ');

            chartGroup.appendChild(createSvgNode('polyline', {
                points: polylinePoints,
                fill: 'none',
                stroke: '#2563eb',
                'stroke-width': '3',
                'stroke-linejoin': 'round',
                'stroke-linecap': 'round'
            }));

            points.forEach((point, index) => {
                const circle = createSvgNode('circle', {
                    cx: projectX(point.timestamp, index),
                    cy: projectY(point.renderedValue),
                    r: '4.5',
                    fill: '#ffffff',
                    stroke: '#1d4ed8',
                    'stroke-width': '2'
                });
                const title = createSvgNode('title');
                title.textContent = point.dateLabel + ' | ' + point.label + ' | ' + formatValue(point.renderedValue);
                circle.appendChild(title);
                chartGroup.appendChild(circle);
            });

            const axisLabelIndexes = Array.from(new Set([
                0,
                Math.floor((points.length - 1) / 2),
                points.length - 1
            ])).filter((index) => index >= 0);

            axisLabelIndexes.forEach((index) => {
                const point = points[index];
                const label = createSvgNode('text', {
                    x: projectX(point.timestamp, index),
                    y: height - 18,
                    'text-anchor': 'middle',
                    fill: '#667085',
                    'font-size': '12'
                });
                label.textContent = point.dateLabel;
                chartGroup.appendChild(label);
            });

            const yAxisTitle = createSvgNode('text', {
                x: 18,
                y: margin.top + (plotHeight / 2),
                transform: 'rotate(-90 18 ' + (margin.top + (plotHeight / 2)) + ')',
                'text-anchor': 'middle',
                fill: '#344054',
                'font-size': '12',
                'font-weight': '600'
            });
            yAxisTitle.textContent = displayedValueField;
            chartGroup.appendChild(yAxisTitle);
        };

        document.getElementById('flip-values-btn').addEventListener('click', () => {
            invertValues = !invertValues;
            renderChart();
        });

        document.getElementById('accumulate-values-btn').addEventListener('click', () => {
            showAccumulated = !showAccumulated;
            renderChart();
        });

        renderChart();
    </script>
</body>
</html>`);
        graphWindow.document.close();
        return true;
    }

    getFormInputType(fieldName) {
        const lowerField = String(fieldName || '').toLowerCase();

        if (lowerField.includes('email')) return 'email';
        if (lowerField.includes('url') || lowerField.includes('link')) return 'url';
        if (lowerField.includes('date')) return 'date';
        if (lowerField.includes('time')) return 'time';
        if (lowerField.includes('phone') || lowerField.includes('tel')) return 'tel';
        if (lowerField.includes('password')) return 'password';

        return 'text';
    }

    validateItemFormData(formData, isCreate = false) {
        const errors = [];

        if (isCreate) {
            if (!formData.id || !String(formData.id).trim()) {
                errors.push({ field: 'id', message: 'ID is required' });
            }

            if (!formData.title || !String(formData.title).trim()) {
                errors.push({ field: 'title', message: 'Title is required' });
            }
        }

        Object.entries(formData).forEach(([fieldName, value]) => {
            const fieldType = this.fieldTypes.get(fieldName);

            if (fieldType === 'structured' && typeof value === 'string' && value.trim()) {
                try {
                    JSON.parse(value);
                } catch (error) {
                    errors.push({
                        field: fieldName,
                        message: `Invalid JSON format: ${error.message}`
                    });
                }
            }
        });

        return {
            isValid: errors.length === 0,
            errors
        };
    }
    
    // ===== GRID DATA ORGANIZATION =====
    
    getGridData() {
        const xField = this.currentAxisSelections.x;
        const yField = this.currentAxisSelections.y;
        
        if (!xField || !yField) {
            return { rows: [], columns: [], cells: new Map() };
        }

        let sourceData = this.filteredData;
        
        // Get distinct values for x and y axes
        const xValues = new Set();
        const yValues = new Set();
        
        sourceData.forEach(item => {
            const xVal = String(this.getFieldValue(item, xField));
            const yVal = String(this.getFieldValue(item, yField));
            xValues.add(xVal);
            yValues.add(yVal);
        });
        
        // Apply header ordering using AdvancedFeaturesManager
        let columns = this.getVisibleAxisHeaders(xField, xValues).sort();
        let rows = this.getVisibleAxisHeaders(yField, yValues).sort();
        const pinnedColumnHeaders = this.getPinnedAxisHeaderValues(xField);
        const pinnedRowHeaders = this.getPinnedAxisHeaderValues(yField);
        
        if (this.getOrderedHeaders) {
            columns = this.getOrderedHeaders(columns, 'column');
            rows = this.getOrderedHeaders(rows, 'row');
        }
        
        // Organize items into grid cells
        const cells = new Map();
        sourceData.forEach(item => {
            const xVal = String(this.getFieldValue(item, xField));
            const yVal = String(this.getFieldValue(item, yField));
            
            // Only include cells that are not collapsed
            if ((!this.isHeaderCollapsed || !this.isHeaderCollapsed(xVal, 'column')) &&
                (!this.isHeaderCollapsed || !this.isHeaderCollapsed(yVal, 'row'))) {
                
                const cellKey = `${yVal}:${xVal}`;
                
                if (!cells.has(cellKey)) {
                    cells.set(cellKey, []);
                }
                cells.get(cellKey).push(item);
            }
        });
        
        return { rows, columns, cells };
    }

    getVisibleAxisHeaders(fieldName, discoveredValues) {
        const normalizedDiscoveredValues = Array.from(discoveredValues || [], (value) => String(value));
        const activeFilter = this.currentFilters.get(fieldName);

        if (!(activeFilter instanceof Set)) {
            // No filter active — include all validValues (from header editor)
            // so that configured columns/rows appear even when empty.
            const schemaField = this.schemaFields[fieldName];
            if (schemaField && Array.isArray(schemaField.validValues) && schemaField.validValues.length > 0) {
                const validSet = new Set(schemaField.validValues.map((v) => String(v)));
                const discoveredSet = new Set(normalizedDiscoveredValues);
                // Start with discovered, then append any missing validValues
                const merged = [...normalizedDiscoveredValues];
                for (const v of schemaField.validValues) {
                    const s = String(v);
                    if (!discoveredSet.has(s)) merged.push(s);
                }
                return merged;
            }
            return normalizedDiscoveredValues;
        }

        const selectedValues = Array.from(activeFilter, (value) => String(value));
        if (selectedValues.length === 0) {
            return [];
        }

        const distinctValues = (this.distinctValues.get(fieldName) || []).map((value) => String(value));
        const selectedValueSet = new Set(selectedValues);
        const orderedSelectedValues = distinctValues.filter((value) => selectedValueSet.has(value));
        const remainingSelectedValues = selectedValues.filter((value) => !orderedSelectedValues.includes(value));

        const discoveredValueSet = new Set(normalizedDiscoveredValues);
        const emptySelectedValues = [...orderedSelectedValues, ...remainingSelectedValues]
            .filter((value) => !discoveredValueSet.has(value));

        if (emptySelectedValues.length > 0) {
            console.debug('[HEADER] preserving empty filtered headers', {
                fieldName,
                preservedHeaders: emptySelectedValues,
                selectedValues,
                discoveredValues: normalizedDiscoveredValues
            });
        }

        return [...orderedSelectedValues, ...remainingSelectedValues];
    }

    getPinnedAxisHeaderValues(fieldName) {
        const activeFilter = this.currentFilters.get(fieldName);
        if (!(activeFilter instanceof Set)) {
            return new Set();
        }

        return new Set(Array.from(activeFilter, (value) => String(value)));
    }

    shouldRenderHeader(headerValue, headerType, pinnedHeaders = new Set()) {
        const normalizedHeaderValue = String(headerValue);

        if (!this.isHeaderCollapsed || !this.isHeaderCollapsed(normalizedHeaderValue, headerType)) {
            return true;
        }

        if (pinnedHeaders.has(normalizedHeaderValue)) {
            console.debug('[HEADER] preserving filtered header despite collapsed state', {
                headerType,
                headerValue: normalizedHeaderValue,
                reason: 'selected in active axis filter'
            });
            return true;
        }

        console.debug('[HEADER] removing header', {
            headerType,
            headerValue: normalizedHeaderValue,
            reason: 'collapsed'
        });
        return false;
    }

    organizeDataForGrid() {
        const { rows, columns, cells } = this.getGridData();
        const structure = {};

        rows.forEach(rowValue => {
            structure[rowValue] = {};

            columns.forEach(colValue => {
                const cellKey = `${rowValue}:${colValue}`;
                structure[rowValue][colValue] = cells.get(cellKey) || [];
            });
        });

        return structure;
    }

    getTagFieldName() {
        if (this.fieldTypes instanceof Map) {
            const matchedField = Array.from(this.fieldTypes.keys()).find((fieldName) =>
                String(fieldName).toLowerCase() === 'tags'
            );
            if (matchedField) {
                return matchedField;
            }
        }

        if (Array.isArray(this.availableFields)) {
            const matchedField = this.availableFields.find((fieldName) =>
                String(fieldName).toLowerCase() === 'tags'
            );
            if (matchedField) {
                return matchedField;
            }
        }

        return null;
    }

    isTagFieldName(fieldName) {
        return String(fieldName || '').toLowerCase() === 'tags';
    }

    normalizeTagValues(value) {
        if (Array.isArray(value)) {
            return value
                .map((entry) => String(entry || '').trim())
                .filter((entry) => entry.length > 0);
        }

        if (typeof value === 'string') {
            return value
                .split(';')
                .map((entry) => entry.trim())
                .filter((entry) => entry.length > 0);
        }

        return [];
    }

    getItemTags(item) {
        if (!item || typeof item !== 'object') {
            return [];
        }

        const tagFieldName = this.getTagFieldName();
        if (!tagFieldName) {
            return [];
        }

        return this.normalizeTagValues(item[tagFieldName]);
    }
    
    // ===== TAG MANAGEMENT =====
    
    getAllTags() {
        const allTags = new Set();
        this.dataset.forEach(item => {
            this.getItemTags(item).forEach(tag => allTags.add(tag));
        });
        this.tagCustomizations.forEach((config, tagName) => {
            allTags.add(tagName);
        });
        return Array.from(allTags).sort();
    }
    
    getTagConfig(tagName) {
        return this.tagCustomizations.get(tagName) || {
            label: tagName,
            color: this.generateTagColor(tagName)
        };
    }

    isFieldValueColorApplicable(fieldName) {
        if (!fieldName) {
            return false;
        }

        if (typeof this.isTagFieldName === 'function' && this.isTagFieldName(fieldName)) {
            return false;
        }

        const fieldType = this.fieldTypes.get(fieldName) || (this.schemaFields[fieldName] && this.schemaFields[fieldName].type) || null;
        return fieldType === 'scalar';
    }

    getConfiguredFieldValueColors(fieldName) {
        if (!fieldName || !this.isFieldValueColorApplicable(fieldName)) {
            return {};
        }

        const schemaField = this.schemaFields[fieldName] || {};
        return schemaField && typeof schemaField.valueColors === 'object' && schemaField.valueColors !== null
            ? { ...schemaField.valueColors }
            : {};
    }

    getFieldValueColor(fieldName, rawValue, fallbackValue = null) {
        if (!fieldName || !this.isFieldValueColorApplicable(fieldName)) {
            return '';
        }

        const valueColors = this.getConfiguredFieldValueColors(fieldName);

        const candidates = [rawValue, fallbackValue]
            .filter((value) => value !== undefined && value !== null && value !== '')
            .map((value) => String(value));

        for (const candidate of candidates) {
            const configuredColor = valueColors[candidate];
            if (typeof configuredColor === 'string' && configuredColor.trim()) {
                return configuredColor.trim();
            }
        }

        return '';
    }

    renderConfiguredFieldValue(element, item, fieldName) {
        if (!element) {
            return false;
        }

        const rawValue = this.getFieldValue(item, fieldName);
        const displayValue = this.getDisplayValue(item, fieldName);
        const shouldRenderAsChip = this.isFieldValueColorApplicable(fieldName);
        const valueColor = this.getFieldValueColor(fieldName, rawValue, displayValue);

        element.innerHTML = '';
        if (displayValue === undefined || displayValue === null || displayValue === '') {
            element.textContent = '';
            return false;
        }

        if (!shouldRenderAsChip || !valueColor) {
            element.textContent = displayValue;
            return false;
        }

        const chip = document.createElement('span');
        chip.className = 'card-tag card-field-value-tag';
        chip.style.backgroundColor = valueColor;
        chip.textContent = displayValue;
        element.appendChild(chip);
        return true;
    }
    
    generateTagColor(tagName) {
        // Simple color generation based on tag name
        let hash = 0;
        for (let i = 0; i < tagName.length; i++) {
            hash = tagName.charCodeAt(i) + ((hash << 5) - hash);
        }
        const hue = Math.abs(hash) % 360;
        return `hsl(${hue}, 60%, 45%)`;
    }
    
    updateTagConfig(tagName, config) {
        this.tagCustomizations.set(tagName, config);
        this.updateViewConfiguration();
        this.renderTags();
        this.renderGrid();
    }

    isCardTagSelected(itemId, tagName) {
        return this.selectedCardTags.has(this.getCardTagSelectionKey(itemId, tagName));
    }

    getCardTagSelectionKey(itemId, tagName) {
        return `${itemId}::${tagName}`;
    }

    getSelectedCardTags() {
        return Array.from(this.selectedCardTags).map((entry) => {
            const separatorIndex = entry.indexOf('::');
            if (separatorIndex === -1) {
                return null;
            }

            return {
                itemId: entry.slice(0, separatorIndex),
                tagName: entry.slice(separatorIndex + 2)
            };
        }).filter(Boolean);
    }

    isControlTagSelected(tagName) {
        return this.selectedControlTags.has(String(tagName));
    }

    getSelectedControlTags() {
        return Array.from(this.selectedControlTags);
    }

    setSelectedControlTag(tagName, options = {}) {
        const { additive = false } = options;
        const normalizedTagName = String(tagName || '').trim();
        if (!normalizedTagName) {
            return;
        }

        if (additive) {
            if (this.selectedControlTags.has(normalizedTagName)) {
                this.selectedControlTags.delete(normalizedTagName);
            } else {
                this.selectedControlTags.add(normalizedTagName);
            }
        } else if (this.selectedControlTags.size === 1 && this.selectedControlTags.has(normalizedTagName)) {
            this.selectedControlTags.clear();
        } else {
            this.selectedControlTags.clear();
            this.selectedControlTags.add(normalizedTagName);
        }

        this.syncSelectedControlTagSelection();
    }

    clearSelectedControlTags() {
        if (this.selectedControlTags.size === 0) {
            return;
        }

        this.selectedControlTags.clear();
        this.syncSelectedControlTagSelection();
    }

    syncSelectedControlTagSelection() {
        const selectedTags = this.selectedControlTags;

        document.querySelectorAll('#tags-container .tag').forEach((tagElement) => {
            tagElement.classList.toggle('is-selected', selectedTags.has(tagElement.dataset.tagName));
        });
    }

    updateSelectedCardTagSnapshot() {
        const [firstSelection] = this.getSelectedCardTags();
        this.selectedCardTag = firstSelection || null;
    }

    deselectCardTag(itemId, tagName) {
        const selectionKey = this.getCardTagSelectionKey(itemId, tagName);
        if (!this.selectedCardTags.delete(selectionKey)) {
            return false;
        }

        this.updateSelectedCardTagSnapshot();
        this.syncSelectedCardTagSelection();
        return true;
    }

    setSelectedCardTag(itemId, tagName, options = {}) {
        const { additive = false } = options;
        const selectionKey = this.getCardTagSelectionKey(itemId, tagName);

        if (additive) {
            if (this.selectedCardTags.has(selectionKey)) {
                this.selectedCardTags.delete(selectionKey);
            } else {
                this.selectedCardTags.add(selectionKey);
            }
        } else if (this.selectedCardTags.size === 1 && this.selectedCardTags.has(selectionKey)) {
            this.selectedCardTags.clear();
        } else {
            this.selectedCardTags.clear();
            this.selectedCardTags.add(selectionKey);
        }

        this.updateSelectedCardTagSnapshot();

        this.syncSelectedCardTagSelection();
    }

    clearSelectedCardTag() {
        if (this.selectedCardTags.size === 0 && !this.selectedCardTag) {
            return;
        }

        this.selectedCardTags.clear();
        this.selectedCardTag = null;
        this.syncSelectedCardTagSelection();
    }

    syncSelectedCardTagSelection() {
        const selectedTags = this.selectedCardTags;

        document.querySelectorAll('.card-tag').forEach((tagElement) => {
            const card = tagElement.closest('.card');
            const isSelected = Boolean(
                card &&
                selectedTags.has(this.getCardTagSelectionKey(card.dataset.itemId, tagElement.dataset.tagName))
            );

            tagElement.classList.toggle('is-selected', isSelected);
        });
    }

    // ===== ITEM GROUPS (V3) =====

    getGroups() {
        return Array.isArray(this.viewConfig.groups) ? this.viewConfig.groups : [];
    }

    getEnabledGroups() {
        return this.getGroups().filter(g => g.enabled !== false);
    }

    addGroup(group) {
        if (!Array.isArray(this.viewConfig.groups)) {
            this.viewConfig.groups = [];
        }
        if (!group.id) {
            group.id = 'grp-' + Date.now();
        }
        this.viewConfig.groups.push(group);
        this.resolveGroupMemberships();
        this.updateViewConfiguration();
    }

    updateGroup(groupId, updates) {
        const groups = this.getGroups();
        const idx = groups.findIndex(g => g.id === groupId);
        if (idx >= 0) {
            Object.assign(groups[idx], updates);
            this.resolveGroupMemberships();
            this.updateViewConfiguration();
        }
    }

    removeGroup(groupId) {
        if (Array.isArray(this.viewConfig.groups)) {
            this.viewConfig.groups = this.viewConfig.groups.filter(g => g.id !== groupId);
            this.resolveGroupMemberships();
            this.updateViewConfiguration();
        }
    }

    toggleGroup(groupId) {
        const groups = this.getGroups();
        const group = groups.find(g => g.id === groupId);
        if (group) {
            group.enabled = !group.enabled;
            this.resolveGroupMemberships();
            this.updateViewConfiguration();
        }
    }

    addItemToGroup(groupId, itemId) {
        const groups = this.getGroups();
        const group = groups.find(g => g.id === groupId);
        if (!group) return;
        if (!Array.isArray(group.manualMembers)) {
            group.manualMembers = [];
        }
        if (group.manualMembers.includes(itemId)) return;
        group.manualMembers.push(itemId);
        this.resolveGroupMemberships();
        this.updateViewConfiguration();
    }

    addItemsToGroup(groupId, itemIds = []) {
        const groups = this.getGroups();
        const group = groups.find(g => g.id === groupId);
        if (!group) return 0;
        if (!Array.isArray(group.manualMembers)) {
            group.manualMembers = [];
        }

        const normalizedItemIds = Array.from(new Set(
            (Array.isArray(itemIds) ? itemIds : [itemIds])
                .map((itemId) => itemId === undefined || itemId === null ? null : String(itemId))
                .filter(Boolean)
        ));

        let addedCount = 0;
        normalizedItemIds.forEach((itemId) => {
            if (group.manualMembers.includes(itemId)) {
                return;
            }

            group.manualMembers.push(itemId);
            addedCount += 1;
        });

        if (addedCount > 0) {
            this.resolveGroupMemberships();
            this.updateViewConfiguration();
        }

        return addedCount;
    }

    removeItemFromGroup(groupId, itemId) {
        const groups = this.getGroups();
        const group = groups.find(g => g.id === groupId);
        if (!group || !Array.isArray(group.manualMembers)) return;
        const idx = group.manualMembers.indexOf(itemId);
        if (idx < 0) return;
        group.manualMembers.splice(idx, 1);
        this.resolveGroupMemberships();
        this.updateViewConfiguration();
    }

    evaluateGroupMembership(item, group) {
        if (!group || group.enabled === false) return false;

        const itemId = this.getItemIdentity(item) || item.id;

        // Manual member check
        if (Array.isArray(group.manualMembers) && group.manualMembers.includes(itemId)) {
            return true;
        }

        // Rule-based check
        if (!group.rule || !group.rule.field) return false;

        const rule = group.rule;
        let itemValue;

        if (rule.field === '_hasChanges') {
            // Virtual field: true if item has a pending change row
            itemValue = Boolean(this.getPendingChangeRowForItem(itemId));
        } else {
            itemValue = this.getFieldValue(item, rule.field);
        }

        const ruleValues = Array.isArray(rule.values) ? rule.values : [];

        switch (rule.operator) {
            case 'equals':
                if (Array.isArray(itemValue)) {
                    return itemValue.some(v => ruleValues.includes(String(v)) || ruleValues.includes(v));
                }
                return ruleValues.includes(String(itemValue)) || ruleValues.includes(itemValue);

            case 'not-equals':
                if (Array.isArray(itemValue)) {
                    return !itemValue.some(v => ruleValues.includes(String(v)) || ruleValues.includes(v));
                }
                return !ruleValues.includes(String(itemValue)) && !ruleValues.includes(itemValue);

            case 'contains':
                if (ruleValues.length === 0) return false;
                // OR logic: match if item value contains ANY of the rule values
                if (Array.isArray(itemValue)) {
                    return ruleValues.some(rv => {
                        const needle = String(rv).toLowerCase().trim();
                        return itemValue.some(v => String(v).toLowerCase().includes(needle));
                    });
                }
                return ruleValues.some(rv => {
                    const needle = String(rv).toLowerCase().trim();
                    return String(itemValue || '').toLowerCase().includes(needle);
                });

            case 'is-empty':
                if (Array.isArray(itemValue)) return itemValue.length === 0;
                return itemValue === undefined || itemValue === null || itemValue === '';

            case 'is-not-empty':
                if (Array.isArray(itemValue)) return itemValue.length > 0;
                return itemValue !== undefined && itemValue !== null && itemValue !== '';

            default:
                return false;
        }
    }

    resolveGroupMemberships() {
        this.groupMemberships = new Map();
        const enabledGroups = this.getEnabledGroups();

        if (enabledGroups.length === 0) {
            this.applyGroupDerivedFields(this.dataset || []);
            if (Array.isArray(this.dataset) && this.dataset.length > 0) {
                this.analyzeFields();
                this.updateFilteredData();
                if (typeof this.renderFieldSelectors === 'function') {
                    this.renderFieldSelectors();
                }
                if (typeof this.renderSlicers === 'function') {
                    this.renderSlicers();
                }
            }
            return;
        }

        const items = this.dataset || [];
        items.forEach(item => {
            const itemId = this.getItemIdentity(item) || item.id;
            const memberOf = enabledGroups.filter(g => this.evaluateGroupMembership(item, g));
            if (memberOf.length > 0) {
                this.groupMemberships.set(itemId, memberOf);
            }
        });

        this.applyGroupDerivedFields(items);

        if (items.length > 0) {
            this.analyzeFields();
            this.updateFilteredData();
            if (typeof this.renderFieldSelectors === 'function') {
                this.renderFieldSelectors();
            }
            if (typeof this.renderSlicers === 'function') {
                this.renderSlicers();
            }
        }

        console.log('[GROUPS] Resolved memberships:', this.groupMemberships.size, 'items in groups');
    }

    getItemGroups(item) {
        const itemId = this.getItemIdentity(item) || item.id;
        return this.groupMemberships ? (this.groupMemberships.get(itemId) || []) : [];
    }

    // ===== INTERACTION MANAGER INITIALIZATION =====
    
    initializeInteractions() {
        // This method will be overridden by interactions.js
        console.log('Interactions initialization placeholder');
    }
    
    initializePersistence() {
        // This method will be overridden by persistence.js
        console.log('Persistence initialization placeholder');
    }
    
    initializeAdvancedFeatures() {
        // This method will be overridden by advanced-features.js
        console.log('Advanced features initialization placeholder');
    }
    
    // ===== PLACEHOLDER METHODS FOR CARD INTERACTIONS =====
    // These will be overridden by interactions.js
    
    showTooltip(event, item) {
        console.log('Tooltip requested for item:', item.id);
        // Will be implemented in interactions.js
    }
    
    hideTooltip() {
        console.log('Hide tooltip requested');
        // Will be implemented in interactions.js
    }
    
    handleCardClick(event, item) {
        console.log('Card click for item:', item.id, event);
        // Will be implemented in interactions.js
    }
    
    handleCardDrop(itemId, newRowValue, newColValue) {
        console.log(`Card drop: ${itemId} to ${newRowValue}:${newColValue}`);
        // Will be implemented in interactions.js
    }
    
    openCommentsDialog(item) {
        console.log('Comments dialog requested for item:', item.id);
        // Will be overridden by advanced-features.js
    }
    
    // ===== CARD MULTI-SELECT =====

    toggleCardSelection(itemId) {
        if (this.selectedCards.has(itemId)) {
            this.selectedCards.delete(itemId);
        } else {
            this.selectedCards.add(itemId);
        }
        this.applyCardSelectionClasses();
    }

    setCardSelection(itemIds = []) {
        this.selectedCards = new Set(
            (Array.isArray(itemIds) ? itemIds : [itemIds])
                .map((itemId) => itemId === undefined || itemId === null ? null : String(itemId))
                .filter(Boolean)
        );
        this.applyCardSelectionClasses();
    }

    clearCardSelection() {
        this.selectedCards.clear();
        this.applyCardSelectionClasses();
    }

    findRenderedItemElement(itemId) {
        const normalizedItemId = itemId === undefined || itemId === null ? null : String(itemId);
        if (!normalizedItemId) {
            return null;
        }

        return Array.from(document.querySelectorAll('.card[data-item-id]')).find(
            (element) => String(element.dataset.itemId || '') === normalizedItemId
        ) || null;
    }

    focusItemInGrid(itemId, options = {}) {
        const {
            select = true,
            behavior = 'auto',
            block = 'center',
            inline = 'center',
            defer = true,
            suppressWarning = false
        } = options;
        const normalizedItemId = itemId === undefined || itemId === null ? null : String(itemId);
        if (!normalizedItemId) {
            return false;
        }

        if (select) {
            this.setCardSelection([normalizedItemId]);
        }

        const revealItem = () => {
            const renderedElement = this.findRenderedItemElement(normalizedItemId);
            if (!renderedElement) {
                if (!suppressWarning) {
                    this.showNotification(`Item ${normalizedItemId} is not currently visible in the grid`, 'warning');
                }
                return false;
            }

            if (typeof renderedElement.scrollIntoView === 'function') {
                renderedElement.scrollIntoView({ behavior, block, inline });
            }
            return true;
        };

        if (defer && typeof window !== 'undefined' && typeof window.requestAnimationFrame === 'function') {
            window.requestAnimationFrame(revealItem);
            return true;
        }

        return revealItem();
    }

    applyCardSelectionClasses() {
        document.querySelectorAll('.card').forEach((card) => {
            const id = card.dataset.itemId;
            if (this.selectedCards.has(id)) {
                card.classList.add('card-selected');
            } else {
                card.classList.remove('card-selected');
            }
        });
    }

    getSelectedItems() {
        return Array.from(this.selectedCards)
            .map((id) => this.getEffectiveItemById(id))
            .filter(Boolean);
    }

    // ===== ADVANCED FEATURES PLACEHOLDER METHODS =====
    
    calculateAggregations(gridData, bottomRightField) {
        // Will be overridden by advanced-features.js
        return { rows: new Map(), columns: new Map() };
    }
    
    getOrderedHeaders(values, headerType) {
        // Will be overridden by advanced-features.js
        return values;
    }
    
    attachHeaderListeners(headerElement, headerValue, headerType) {
        // Will be overridden by advanced-features.js
        console.log('Header listeners attachment placeholder');
    }
    
    isHeaderCollapsed(headerValue, headerType) {
        // Will be overridden by advanced-features.js
        return false;
    }
    
    formatAggregationValue(aggregation) {
        // Will be overridden by advanced-features.js
        return '';
    }
    
    openEditDialog(item) {
        console.log('Edit dialog requested for item:', item.id);
        // Will be implemented in interactions.js
    }

    openInlineValueEditor(card, item, fieldName, element) {
        console.log('Inline value editor requested for item:', item.id, fieldName, card, element);
        // Will be implemented in interactions.js
    }

    openCreateTooltip(event, initialValues) {
        console.log('Create tooltip requested:', event, initialValues);
        // Will be implemented in interactions.js
    }

    addTagToItem(itemId, tagName) {
        console.log(`Add tag requested: ${tagName} -> ${itemId}`);
        // Will be implemented in interactions.js
    }

    removeTagFromItem(itemId, tagName) {
        console.log(`Remove tag requested: ${tagName} -> ${itemId}`);
        // Will be implemented in interactions.js
    }
    
    initializeEventListeners() {
        // Placeholder - will be implemented in controls.js
    }
    
    renderFieldSelectors() {
        // Placeholder - will be implemented in controls.js
        console.log('Field selectors render requested');
    }

    renderCurrentUserName() {
        const currentUserElement = document.getElementById('current-user-name');
        if (currentUserElement) {
            currentUserElement.textContent = this.getCurrentUserName();
        }
    }
    
    renderSlicers() {
        // Placeholder - will be implemented in controls.js
        console.log('Slicers render requested');
    }
    
    renderTags() {
        // Placeholder - will be implemented in controls.js
        console.log('Tags render requested');
    }

    renderGroups() {
        // Placeholder - will be implemented in controls.js
        console.log('Groups render requested');
    }
    
    renderGroupLegend() {
        const legend = document.getElementById('group-legend');
        if (!legend) return;

        const enabledGroups = this.getEnabledGroups();
        if (enabledGroups.length === 0) {
            legend.classList.add('hidden');
            legend.innerHTML = '';
            return;
        }

        legend.classList.remove('hidden');
        legend.innerHTML = '';

        enabledGroups.forEach(g => {
            const item = document.createElement('span');
            item.className = 'group-legend-item';

            const swatch = document.createElement('span');
            swatch.className = 'group-legend-swatch';
            swatch.style.backgroundColor = g.color;

            const name = document.createElement('span');
            name.className = 'group-legend-name';
            name.textContent = g.name;

            item.appendChild(swatch);
            item.appendChild(name);
            legend.appendChild(item);
        });
    }

    renderGrid() {
        this.renderGroupLegend();
        const gridContainer = document.getElementById('grid-container');
        const { rows, columns, cells } = this.getGridData();
        const bottomRightField = this.currentCardSelections.bottomRight;
        const gridData = { structure: this.buildGridStructure(rows, columns, cells) };
        const aggregations = bottomRightField && this.isNumericField(bottomRightField)
            ? this.calculateAggregations(gridData, bottomRightField)
            : null;
        
        if (rows.length === 0 || columns.length === 0) {
            this.renderEmptyGrid(gridContainer);
            return;
        }
        
        gridContainer.innerHTML = '';
        
        // Create the grid wrapper
        const dataGrid = document.createElement('div');
        dataGrid.className = 'data-grid';
        const columnTracks = columns.map((colValue) => (
            this.isHeaderCollapsed && this.isHeaderCollapsed(colValue, 'column')
                ? 'var(--grid-collapsed-column-width)'
                : 'minmax(var(--grid-data-column-width), 1fr)'
        ));
        const rowTracks = rows.map((rowValue) => (
            this.isHeaderCollapsed && this.isHeaderCollapsed(rowValue, 'row')
                ? 'var(--grid-collapsed-row-height)'
                : 'minmax(120px, auto)'
        ));
        const minColumnWidths = columns.map((colValue) => (
            this.isHeaderCollapsed && this.isHeaderCollapsed(colValue, 'column')
                ? 'var(--grid-collapsed-column-width)'
                : 'var(--grid-data-column-width)'
        ));
        
        // Configure CSS Grid layout
        dataGrid.style.gridTemplateColumns = `var(--grid-row-header-width) ${columnTracks.join(' ')}`;
        dataGrid.style.gridTemplateRows = `var(--grid-column-header-height) ${rowTracks.join(' ')}`;
        dataGrid.style.minWidth = `calc(var(--grid-row-header-width) + ${minColumnWidths.join(' + ')})`;
        
        // Add corner cell
        const cornerCell = this.createCornerHeader(aggregations);
        dataGrid.appendChild(cornerCell);
        
        // Add column headers
        columns.forEach((colValue, colIndex) => {
            const colHeader = this.createColumnHeader(colValue, colIndex + 2, aggregations);
            dataGrid.appendChild(colHeader);
        });
        
        // Add row headers and cells
        rows.forEach((rowValue, rowIndex) => {
            const rowHeader = this.createRowHeader(rowValue, rowIndex + 2, aggregations);
            dataGrid.appendChild(rowHeader);
            
            columns.forEach((colValue, colIndex) => {
                const cell = this.createGridCell(rowValue, colValue, cells, rowIndex + 2, colIndex + 2);
                if (cell) {
                    dataGrid.appendChild(cell);
                }
            });
        });
        
        gridContainer.appendChild(dataGrid);
        const isTableMode = this.cellRenderMode === 'table';
        dataGrid.classList.toggle('data-grid-table-mode', isTableMode);
        gridContainer.classList.toggle('grow-cards', this.growCards && !isTableMode);
        gridContainer.classList.toggle('table-mode', isTableMode);
        this.applyCardSelectionClasses();
        this.applyTableGraphSelectionClasses();
        console.log(`Grid rendered: ${rows.length} rows x ${columns.length} columns`);

        // Right-click on empty grid space → show all changes
        dataGrid.addEventListener('contextmenu', (e) => {
            if (e.target.closest('.card') || e.target.closest('.grid-cell-table') || e.target.closest('.grid-header')) return;
            e.preventDefault();
            e.stopPropagation();
            if (this.interactionManager && typeof this.interactionManager.showGridChangesContextMenu === 'function') {
                this.interactionManager.showGridChangesContextMenu(e);
            }
        });

        // V5: Post-render hook for focus mode
        if (this.relationUIManager) {
            this.relationUIManager.onGridRendered();
        }
    }

    buildGridStructure(rows, columns, cells) {
        const structure = {};

        rows.forEach((rowValue) => {
            structure[rowValue] = {};

            columns.forEach((colValue) => {
                const cellKey = `${rowValue}:${colValue}`;
                structure[rowValue][colValue] = cells.get(cellKey) || [];
            });
        });

        return structure;
    }
    
    renderEmptyGrid(container) {
        container.innerHTML = '';
        const placeholder = document.createElement('div');
        placeholder.className = 'grid-placeholder empty-state no-data';
        
        if (!this.currentAxisSelections.x || !this.currentAxisSelections.y) {
            placeholder.innerHTML = '<p>Select X and Y axis fields to view grid</p>';
        } else if (this.filteredData.length === 0) {
            placeholder.innerHTML = '<p>No data matches current filters</p>';
        } else {
            placeholder.innerHTML = '<p>No data to display</p>';
        }
        
        container.appendChild(placeholder);
    }
    
    createCornerHeader(aggregations) {
        const cornerCell = document.createElement('div');
        cornerCell.className = 'grid-header corner-header';
        cornerCell.style.gridRow = '1';
        cornerCell.style.gridColumn = '1';

        if (aggregations && aggregations.total) {
            const prefix = this.aggregationMode === 'count' ? '#' : 'Σ';
            const totalElement = document.createElement('span');
            totalElement.className = 'header-aggregation corner-aggregation';
            totalElement.textContent = `${prefix} ${this.formatAggregationValue(aggregations.total)}`;
            totalElement.style.cursor = 'pointer';
            totalElement.title = `Click to switch to ${this.aggregationMode === 'sum' ? 'count' : 'sum'}`;
            totalElement.addEventListener('click', (e) => {
                e.stopPropagation();
                this.aggregationMode = this.aggregationMode === 'sum' ? 'count' : 'sum';
                this.renderGrid();
            });
            cornerCell.appendChild(totalElement);
        }

        return cornerCell;
    }

    createColumnHeader(colValue, gridColumn, aggregations) {
        const header = document.createElement('div');
        header.className = 'grid-header column-header';
        header.style.gridRow = '1';
        header.style.gridColumn = gridColumn;
        header.dataset.value = colValue || '';
        header.dataset.headerType = 'column';
        header.title = colValue || '(empty)';
        const isCollapsed = this.isHeaderCollapsed && this.isHeaderCollapsed(colValue, 'column');
        
        if (!isCollapsed) {
            const headerContent = document.createElement('div');
            headerContent.className = 'header-label';
            headerContent.textContent = colValue || '(empty)';
            header.appendChild(headerContent);
        }
        
        if (aggregations && !isCollapsed) {
            const columnAgg = aggregations.columns.get(colValue);
            
            if (columnAgg) {
                const prefix = this.aggregationMode === 'count' ? '#' : 'Σ';
                const aggElement = document.createElement('span');
                aggElement.className = 'header-aggregation';
                aggElement.textContent = `${prefix} ${this.formatAggregationValue(columnAgg)}`;
                aggElement.style.cursor = 'pointer';
                aggElement.title = `Click to switch to ${this.aggregationMode === 'sum' ? 'count' : 'sum'}`;
                aggElement.addEventListener('click', (e) => {
                    e.stopPropagation();
                    this.aggregationMode = this.aggregationMode === 'sum' ? 'count' : 'sum';
                    this.renderGrid();
                });
                header.appendChild(aggElement);
            }
        }
        
        // Attach header interactions using AdvancedFeaturesManager
        this.attachHeaderListeners(header, colValue, 'column');
        
        return header;
    }
    
    createRowHeader(rowValue, gridRow, aggregations) {
        const header = document.createElement('div');
        header.className = 'grid-header row-header';
        header.style.gridRow = gridRow;
        header.style.gridColumn = '1';
        header.dataset.value = rowValue || '';
        header.dataset.headerType = 'row';
        header.title = rowValue || '(empty)';
        const isCollapsed = this.isHeaderCollapsed && this.isHeaderCollapsed(rowValue, 'row');
        
        if (!isCollapsed) {
            const headerContent = document.createElement('div');
            headerContent.className = 'header-label';
            headerContent.textContent = rowValue || '(empty)';
            header.appendChild(headerContent);
        }
        
        if (aggregations && !isCollapsed) {
            const rowAgg = aggregations.rows.get(rowValue);
            
            if (rowAgg) {
                const prefix = this.aggregationMode === 'count' ? '#' : 'Σ';
                const aggElement = document.createElement('span');
                aggElement.className = 'header-aggregation';
                aggElement.textContent = `${prefix} ${this.formatAggregationValue(rowAgg)}`;
                aggElement.style.cursor = 'pointer';
                aggElement.title = `Click to switch to ${this.aggregationMode === 'sum' ? 'count' : 'sum'}`;
                aggElement.addEventListener('click', (e) => {
                    e.stopPropagation();
                    this.aggregationMode = this.aggregationMode === 'sum' ? 'count' : 'sum';
                    this.renderGrid();
                });
                header.appendChild(aggElement);
            }
        }
        
        // Attach header interactions using AdvancedFeaturesManager
        this.attachHeaderListeners(header, rowValue, 'row');
        
        return header;
    }
    
    createGridCell(rowValue, colValue, cells, gridRow, gridColumn) {
        const isCollapsedRow = this.isHeaderCollapsed && this.isHeaderCollapsed(rowValue, 'row');
        const isCollapsedColumn = this.isHeaderCollapsed && this.isHeaderCollapsed(colValue, 'column');

        if (isCollapsedRow || isCollapsedColumn) {
            return null;
        }

        const cell = document.createElement('div');
        cell.className = 'grid-cell';
        cell.style.gridRow = gridRow;
        cell.style.gridColumn = gridColumn;
        cell.dataset.rowValue = rowValue === undefined || rowValue === null ? '' : String(rowValue);
        cell.dataset.colValue = colValue === undefined || colValue === null ? '' : String(colValue);
        
        const cellKey = `${rowValue}:${colValue}`;
        const items = cells.get(cellKey) || [];
        
        if (this.cellRenderMode === 'table') {
            const table = this.createCellTable(items, rowValue, colValue);
            if (table) {
                cell.appendChild(table);
            }
        } else {
            // Create cards for items in this cell
            items.forEach(item => {
                const card = this.createCard(item);
                cell.appendChild(card);
            });
        }
        
        // Add drop zone functionality for drag & drop
        this.makeDropZone(cell, rowValue, colValue);
        
        return cell;
    }

    createCellTable(items, rowValue = '', colValue = '') {
        if (!Array.isArray(items) || items.length === 0) {
            return null;
        }

        const fields = this.getTableColumnFields();
        const summary = this.getTableSummaryForItems(items, fields);
        const wrapper = document.createElement('div');
        wrapper.className = 'grid-cell-table';
        const tableGraphContext = {
            itemIds: items.map((item) => this.getItemIdentity(item) || item.id).filter(Boolean),
            fields: [...fields],
            contextLabel: this.formatTableGraphContextLabel(
                rowValue === undefined || rowValue === null ? '' : String(rowValue),
                colValue === undefined || colValue === null ? '' : String(colValue)
            )
        };
        tableGraphContext.selectionKey = this.createTableGraphSelectionKey(
            tableGraphContext.contextLabel,
            tableGraphContext.itemIds,
            tableGraphContext.fields
        );
        wrapper._tableGraphContext = tableGraphContext;
        wrapper.dataset.tableGraphKey = tableGraphContext.selectionKey;

        wrapper.addEventListener('contextmenu', (event) => {
            if (event.target.closest('.grid-table-row')) {
                return;
            }

            if (event.shiftKey && this.interactionManager && typeof this.interactionManager.handleShiftTableContextSelection === 'function') {
                this.interactionManager.handleShiftTableContextSelection(event, wrapper._tableGraphContext);
                return;
            }

            event.preventDefault();
            event.stopPropagation();
            if (this.interactionManager && typeof this.interactionManager.showTableContextMenu === 'function') {
                this.interactionManager.showTableContextMenu(event, wrapper._tableGraphContext);
            }
        });

        const table = document.createElement('table');
        table.className = 'grid-table';

        const thead = document.createElement('thead');
        const headerRow = document.createElement('tr');
        fields.forEach((fieldName) => {
            const headerCell = document.createElement('th');
            headerCell.textContent = this.formatFieldLabel(fieldName);
            headerCell.title = fieldName;
            if (this.isNumericField(fieldName)) {
                headerCell.classList.add('grid-table-cell-numeric');
            }
            headerRow.appendChild(headerCell);
        });
        thead.appendChild(headerRow);
        table.appendChild(thead);

        const tbody = document.createElement('tbody');
        items.forEach((item) => {
            tbody.appendChild(this.createTableRow(item, fields));
        });
        table.appendChild(tbody);

        if (summary) {
            table.appendChild(this.createTableSummaryFooter(fields, summary));
        }

        wrapper.appendChild(table);
        wrapper.classList.toggle('table-graph-selected', this.isTableGraphSelected(tableGraphContext));
        return wrapper;
    }

    getTableSummaryForItems(items, fields) {
        const summaryMode = this.normalizeTableSummaryMode(this.tableSummaryMode);
        if (summaryMode === 'none') {
            return null;
        }

        const numericFields = fields.filter((fieldName) => this.isNumericField(fieldName));
        if (numericFields.length === 0) {
            return null;
        }

        const summaryValues = new Map();
        numericFields.forEach((fieldName) => {
            const numericValues = items
                .map((item) => this.parseNumericValue(this.getFieldValue(item, fieldName)))
                .filter((value) => value !== null);

            if (numericValues.length === 0) {
                return;
            }

            const summaryValue = this.getSummaryValueFromNumericValues(numericValues, summaryMode);

            if (summaryValue !== null) {
                summaryValues.set(fieldName, summaryValue);
            }
        });

        if (summaryValues.size === 0) {
            return null;
        }

        return {
            mode: summaryMode,
            values: summaryValues
        };
    }

    createTableSummaryFooter(fields, summary) {
        const tfoot = document.createElement('tfoot');
        const row = document.createElement('tr');
        row.className = 'grid-table-summary-row';

        const label = summary.mode.charAt(0).toUpperCase() + summary.mode.slice(1);
        const nonNumericLabelIndex = fields.findIndex((fieldName) => !summary.values.has(fieldName));

        fields.forEach((fieldName, index) => {
            const cell = document.createElement('td');
            const summaryValue = summary.values.get(fieldName);

            if (summaryValue !== undefined) {
                cell.textContent = nonNumericLabelIndex === -1 && index === 0
                    ? `${label}: ${this.formatNumericSummaryValue(summaryValue)}`
                    : this.formatNumericSummaryValue(summaryValue);
                cell.classList.add('grid-table-cell-numeric');
                cell.title = `${label} of ${fieldName}`;
            } else if (index === nonNumericLabelIndex) {
                cell.textContent = label;
                cell.classList.add('grid-table-summary-label');
                cell.title = `${label} summary row`;
            } else {
                cell.textContent = '';
            }

            row.appendChild(cell);
        });

        tfoot.appendChild(row);
        return tfoot;
    }

    getTableColumnFields() {
        const configuredFields = this.normalizeFieldList(this.tableColumnFields);
        if (configuredFields.length > 0) {
            return configuredFields;
        }

        const fallbackFields = this.getDefaultTableColumnFields(this.currentCardSelections);
        return fallbackFields.length > 0 ? fallbackFields : ['id'];
    }

    createTableRow(item, fields) {
        const row = document.createElement('tr');
        row.className = 'card grid-table-row';
        row.draggable = true;
        row.dataset.itemId = this.getItemIdentity(item) || item.id;

        const itemGroups = this.getItemGroups(item);
        this.applyTableRowGroupStyling(row, itemGroups);

        fields.forEach((fieldName) => {
            const cell = document.createElement('td');
            const displayValue = this.isNumericField(fieldName)
                ? this.formatFixedDecimalValue(this.getFieldValue(item, fieldName), 2)
                : this.getDisplayValue(item, fieldName);
            const normalizedValue = displayValue === '' ? ' ' : displayValue;
            cell.textContent = normalizedValue;
            cell.title = displayValue || '(empty)';
            cell.dataset.fieldName = fieldName;

            if (this.isNumericField(fieldName)) {
                cell.classList.add('grid-table-cell-numeric');
            }

            row.appendChild(cell);
        });

        this.addCardEventHandlers(row, item);
        return row;
    }

    applyTableRowGroupStyling(row, itemGroups) {
        if (!Array.isArray(itemGroups) || itemGroups.length === 0) {
            row.style.removeProperty('--table-row-group-color');
            row.style.removeProperty('--table-row-background');
            row.removeAttribute('title');
            return;
        }

        row.style.setProperty('--table-row-group-color', itemGroups[0].color);
        row.style.setProperty('--table-row-background', this.hexToTint(itemGroups[0].color, 0.08));
        row.title = itemGroups.length > 1
            ? `Groups: ${itemGroups.map((group) => group.name).join(', ')}`
            : `Group: ${itemGroups[0].name}`;
    }

    formatFieldLabel(fieldName) {
        return fieldName.charAt(0).toUpperCase() +
            fieldName.slice(1).replace(/([A-Z])/g, ' $1');
    }
    
    createCard(item) {
        const template = document.getElementById('card-template');
        const cardElement = template.content.cloneNode(true);
        
        const card = cardElement.querySelector('.card');
        const topLeft = cardElement.querySelector('.card-top-left');
        const topRight = cardElement.querySelector('.card-top-right');
        const title = cardElement.querySelector('.card-title');
        const bottomRight = cardElement.querySelector('.card-bottom-right');
        
        // Set card data
        card.dataset.itemId = this.getItemIdentity(item) || item.id;
        
        // Populate card content based on current axis selections
        const selections = this.currentCardSelections;
        const tagFieldName = this.getTagFieldName();
        const fallbackTopRightField = !selections.topRight && tagFieldName && this.getItemTags(item).length > 0
            ? tagFieldName
            : null;
        const resolvedTopRightField = selections.topRight || fallbackTopRightField;
        
        if (selections.topLeft) {
            this.renderConfiguredFieldValue(topLeft, item, selections.topLeft);
            this.applyFieldStyling(topLeft, item, selections.topLeft);
        }
        
        if (resolvedTopRightField) {
            this.populateTopRight(topRight, item, resolvedTopRightField);
        }
        
        if (selections.title) {
            title.textContent = this.getDisplayValue(item, selections.title);
        } else {
            title.textContent = item.id || 'Untitled';
        }
        
        if (selections.bottomRight) {
            this.renderConfiguredFieldValue(bottomRight, item, selections.bottomRight);
            bottomRight.dataset.fieldName = selections.bottomRight;
            bottomRight.classList.add('is-inline-editable');
        }
        
        // Add card event handlers
        this.addCardEventHandlers(card, item);

        // Apply group highlighting — each group gets its own edge
        // Order: left, right, top, bottom
        const itemGroups = this.getItemGroups(item);
        if (itemGroups.length > 0) {
            card.dataset.groupCount = itemGroups.length;
            const edges = ['borderLeft', 'borderRight', 'borderTop', 'borderBottom'];
            itemGroups.forEach((g, i) => {
                if (i < edges.length) {
                    card.style[edges[i]] = `4px solid ${g.color}`;
                }
            });
            // Tint with first group color
            card.style.backgroundColor = this.hexToTint(itemGroups[0].color, 0.06);
        }

        // V5: Render relation badges
        if (this.relationUIManager) {
            this.relationUIManager.renderBadges(card, item);
        }
        
        return card;
    }
    
    populateTopRight(element, item, fieldName) {
        element.innerHTML = '';
        const isTagField = this.isTagFieldName(fieldName);
        element.classList.toggle('card-tags', isTagField);
        element.classList.toggle('tags', isTagField);
        const itemIdentity = this.getItemIdentity(item) || item.id;
        
        const fieldType = this.fieldTypes.get(fieldName);
        const value = this.getFieldValue(item, fieldName);
        const tagValues = isTagField ? this.normalizeTagValues(value) : [];
        
        if ((fieldType === 'multi-value' && Array.isArray(value)) || (isTagField && tagValues.length > 0)) {
            // Render tags
            const renderedTags = isTagField ? tagValues : value;
            renderedTags.forEach(tagName => {
                const tagConfig = this.getTagConfig(tagName);
                const tagSpan = document.createElement('span');
                tagSpan.className = 'card-tag';
                tagSpan.style.backgroundColor = tagConfig.color;
                tagSpan.textContent = tagConfig.label;

                if (isTagField) {
                    tagSpan.dataset.tagName = tagName;
                    tagSpan.draggable = true;
                    tagSpan.classList.add('is-removable');
                    tagSpan.classList.toggle('is-selected', this.isCardTagSelected(itemIdentity, tagName));

                    tagSpan.addEventListener('dragstart', (e) => {
                        e.stopPropagation();
                        e.dataTransfer.setData('application/x-tag-name', tagName);
                        e.dataTransfer.setData('application/x-card-tag', JSON.stringify({ itemId: itemIdentity, tagName }));
                        e.dataTransfer.setData('text/plain', `tag:${tagName}`);
                        e.dataTransfer.effectAllowed = 'copyMove';
                        this.draggedCardTag = {
                            itemId: itemIdentity,
                            tagName,
                            wasDroppedOnCard: false
                        };
                        document.body.classList.add('tag-drag-mode');
                    });

                    tagSpan.addEventListener('dragend', (e) => {
                        e.stopPropagation();

                        const draggedCardTag = this.draggedCardTag;
                        document.body.classList.remove('tag-drag-mode');

                        if (
                            draggedCardTag &&
                            draggedCardTag.itemId === itemIdentity &&
                            draggedCardTag.tagName === tagName
                        ) {
                            this.draggedCardTag = null;

                            if (!draggedCardTag.wasDroppedOnCard) {
                                this.removeTagFromItem(itemIdentity, tagName);
                            }
                        }
                    });
                }

                element.appendChild(tagSpan);
            });
        } else {
            this.renderConfiguredFieldValue(element, item, fieldName);
        }
    }
    
    applyFieldStyling(element, item, fieldName) {
        const fieldType = this.fieldTypes.get(fieldName);
        
        if (fieldType === 'multi-value') {
            element.style.fontSize = '10px';
            element.style.maxWidth = '100px';
        }
    }

    hexToTint(hex, alpha) {
        const r = parseInt(hex.slice(1, 3), 16);
        const g = parseInt(hex.slice(3, 5), 16);
        const b = parseInt(hex.slice(5, 7), 16);
        return `rgba(${r}, ${g}, ${b}, ${alpha})`;
    }
    
    addCardEventHandlers(card, item) {
        let pendingCardClickTimer = null;

        const isTagDrag = (event) => {
            const types = event.dataTransfer ? Array.from(event.dataTransfer.types || []) : [];
            return (this.controlPanel && this.controlPanel.draggedTagName !== null) ||
                this.draggedCardTag !== null ||
                types.includes('application/x-tag-name') ||
                types.includes('text/plain');
        };

        const isGroupDrag = (event) => {
            const types = event.dataTransfer ? Array.from(event.dataTransfer.types || []) : [];
            return (this.controlPanel && this.controlPanel.draggedGroupId !== null) ||
                types.includes('application/x-group-id');
        };

        const isTagOrGroupDrag = (event) => isTagDrag(event) || isGroupDrag(event);

        const isCardItemDrag = (event) => {
            const types = event.dataTransfer ? Array.from(event.dataTransfer.types || []) : [];
            return types.includes('application/x-item-id');
        };

        const getDraggedTagName = (event) => {
            const explicitTag = event.dataTransfer.getData('application/x-tag-name');
            if (explicitTag) {
                return explicitTag;
            }

            const plainText = event.dataTransfer.getData('text/plain');
            if (plainText.startsWith('tag:')) {
                return plainText.slice(4);
            }

            return this.controlPanel ? this.controlPanel.draggedTagName : null;
        };

        const getDraggedCardTag = (event) => {
            const encodedTag = event.dataTransfer.getData('application/x-card-tag');
            if (encodedTag) {
                try {
                    return JSON.parse(encodedTag);
                } catch (error) {
                    console.warn('Failed to parse dragged card tag payload', error);
                }
            }

            return this.draggedCardTag;
        };

        // Add drag functionality
        card.addEventListener('dragstart', (e) => {
            e.stopPropagation();
            const itemIdentity = this.getItemIdentity(item) || item.id;
            const xField = this.currentAxisSelections.x;
            const yField = this.currentAxisSelections.y;

            console.log('[CARD DRAG] start', {
                itemId: itemIdentity,
                pointer: {
                    x: e.clientX,
                    y: e.clientY
                },
                axisFields: {
                    x: xField,
                    y: yField
                },
                currentCell: {
                    rowValue: yField ? this.getFieldValue(item, yField) : null,
                    colValue: xField ? this.getFieldValue(item, xField) : null
                }
            });

            if (typeof this.hideTooltip === 'function') {
                this.hideTooltip();
            }

            if (this.interactionManager && typeof this.interactionManager.cancelPendingEditorActivation === 'function') {
                this.interactionManager.cancelPendingEditorActivation();
            }

            e.dataTransfer.setData('application/x-item-id', itemIdentity);
            e.dataTransfer.setData('text/plain', String(itemIdentity));
            e.dataTransfer.effectAllowed = 'all';

            const dragImage = this.getOrCreateCardDragImage();
            if (dragImage && typeof e.dataTransfer.setDragImage === 'function') {
                e.dataTransfer.setDragImage(dragImage, 0, 0);
            }

            card.classList.add('dragging');
        });
        
        card.addEventListener('dragend', (e) => {
            console.log('[CARD DRAG] end', {
                itemId: this.getItemIdentity(item) || item.id,
                pointer: {
                    x: e.clientX,
                    y: e.clientY
                }
            });
            this.clearDebugDropZone();
            card.classList.remove('dragging');
        });

        card.addEventListener('dragover', (e) => {
            if (isCardItemDrag(e)) {
                e.preventDefault();
                e.stopPropagation();
                e.dataTransfer.dropEffect = 'link';
                card.classList.add('relation-drop-target-active');
                return;
            }

            if (!isTagOrGroupDrag(e)) {
                return;
            }

            e.preventDefault();
            e.stopPropagation();
            e.dataTransfer.dropEffect = 'copy';
            card.classList.add('tag-drop-target-active');
            if (isGroupDrag(e)) {
                card.classList.add('group-drop-target-active');
            }
        });

        card.addEventListener('dragleave', (e) => {
            if (!card.contains(e.relatedTarget)) {
                card.classList.remove('tag-drop-target-active');
                card.classList.remove('group-drop-target-active');
                card.classList.remove('relation-drop-target-active');
            }
        });

        card.addEventListener('drop', (e) => {
            // Handle card-on-card drop: create parent-child relation
            if (isCardItemDrag(e)) {
                e.preventDefault();
                e.stopPropagation();
                card.classList.remove('relation-drop-target-active');

                const draggedItemId = e.dataTransfer.getData('application/x-item-id');
                const targetItemId = this.getItemIdentity(item) || item.id;

                if (draggedItemId && draggedItemId !== targetItemId && this.relationUIManager) {
                    this.relationUIManager.setParent(draggedItemId, targetItemId);
                }
                return;
            }

            // Handle group drops
            if (isGroupDrag(e)) {
                e.preventDefault();
                e.stopPropagation();
                card.classList.remove('tag-drop-target-active');
                card.classList.remove('group-drop-target-active');

                const groupId = e.dataTransfer.getData('application/x-group-id') ||
                    (this.controlPanel ? this.controlPanel.draggedGroupId : null);
                if (groupId) {
                    const itemId = this.getItemIdentity(item) || item.id;
                    const targetItemIds = this.selectedCards.size > 1 && this.selectedCards.has(String(itemId))
                        ? Array.from(this.selectedCards)
                        : [itemId];
                    const addedCount = this.addItemsToGroup(groupId, targetItemIds);
                    if (addedCount > 0) {
                        this.renderGrid();
                    }
                    if (this.controlPanel) {
                        this.controlPanel.renderGroups();
                    }
                }
                return;
            }

            if (!isTagDrag(e)) {
                return;
            }

            e.preventDefault();
            e.stopPropagation();
            card.classList.remove('tag-drop-target-active');

            const tagName = getDraggedTagName(e);
            const draggedCardTag = getDraggedCardTag(e);
            if (
                draggedCardTag &&
                this.draggedCardTag &&
                draggedCardTag.itemId === this.draggedCardTag.itemId &&
                draggedCardTag.tagName === this.draggedCardTag.tagName
            ) {
                this.draggedCardTag.wasDroppedOnCard = true;
            }

            if (tagName) {
                const itemId = this.getItemIdentity(item) || item.id;
                const targetItemIds = this.selectedCards.size > 1 && this.selectedCards.has(String(itemId))
                    ? Array.from(this.selectedCards)
                    : [itemId];
                this.addTagToItem(targetItemIds, tagName);
            }
        });

        card.addEventListener('mouseenter', (e) => {
            console.log('[CARD HOVER] mouseenter', { itemId: item.id });
            if (this.showTooltips) {
                this.showTooltip(e, item, 'preview');
            }
        });

        card.addEventListener('mousedown', (e) => {
            if (e.button !== 0) {
                return;
            }

            // Shift+click is for multi-select, do not begin editor activation
            if (e.shiftKey) {
                return;
            }

            if (e.target.closest('.card-tag.is-removable') || e.target.closest('.card-bottom-right.is-inline-editable') || e.target.closest('.card-action-btn')) {
                return;
            }

            if (this.interactionManager && typeof this.interactionManager.beginEditorActivation === 'function') {
                this.interactionManager.beginEditorActivation();
            }
        });

        card.addEventListener('mouseleave', () => {
            console.log('[CARD HOVER] mouseleave', {
                itemId: item.id,
                pendingEditorActivation: this.interactionManager && typeof this.interactionManager.isEditorActivationPending === 'function'
                    ? this.interactionManager.isEditorActivationPending()
                    : null
            });
            if (
                this.interactionManager &&
                typeof this.interactionManager.isEditorActivationPending === 'function' &&
                this.interactionManager.isEditorActivationPending()
            ) {
                return;
            }

            this.hideTooltip('preview');
        });
        
        // Add click handler for URL navigation
        card.addEventListener('click', (e) => {
            console.log('[CARD CLICK] click', { itemId: item.id, shiftKey: e.shiftKey });

            // Shift+click: toggle card selection for multi-select
            const tagElement = e.target.closest('.card-tag.is-removable');
            if (tagElement) {
                if (pendingCardClickTimer) {
                    clearTimeout(pendingCardClickTimer);
                    pendingCardClickTimer = null;
                }
                e.stopPropagation();
                const itemId = this.getItemIdentity(item) || item.id;
                this.setSelectedCardTag(itemId, tagElement.dataset.tagName, { additive: e.shiftKey });
                return;
            }

            // Shift+click: toggle card selection for multi-select
            if (e.shiftKey) {
                if (pendingCardClickTimer) {
                    clearTimeout(pendingCardClickTimer);
                    pendingCardClickTimer = null;
                }
                e.stopPropagation();
                const itemId = this.getItemIdentity(item) || item.id;
                this.toggleCardSelection(itemId);
                this.hideTooltip('preview');
                if (this.interactionManager && typeof this.interactionManager.cancelPendingEditorActivation === 'function') {
                    this.interactionManager.cancelPendingEditorActivation();
                }
                // Refresh details panel if open
                if (this.detailsPanelManager && this.detailsPanelManager.isOpen()) {
                    this.detailsPanelManager.refreshForSelection();
                }
                return;
            }

            const inlineValue = e.target.closest('.card-bottom-right.is-inline-editable');
            if (inlineValue) {
                if (pendingCardClickTimer) {
                    clearTimeout(pendingCardClickTimer);
                    pendingCardClickTimer = null;
                }
                e.stopPropagation();
                this.clearSelectedCardTag();
                this.openInlineValueEditor(card, item, inlineValue.dataset.fieldName, inlineValue);
                return;
            }

            // Don't navigate if clicking on action buttons
            if (e.target.closest('.card-action-btn')) {
                if (pendingCardClickTimer) {
                    clearTimeout(pendingCardClickTimer);
                    pendingCardClickTimer = null;
                }
                return;
            }

            if (pendingCardClickTimer) {
                clearTimeout(pendingCardClickTimer);
            }

            if (this.interactionManager && typeof this.interactionManager.beginEditorActivation === 'function') {
                this.interactionManager.beginEditorActivation();
            }

            pendingCardClickTimer = window.setTimeout(() => {
                pendingCardClickTimer = null;
                if (this.interactionManager && typeof this.interactionManager.cancelPendingEditorActivation === 'function') {
                    this.interactionManager.cancelPendingEditorActivation();
                }
                this.clearCardSelection();
                this.clearSelectedCardTag();
                this.handleCardClick(e, item);
            }, 250);
        });

        card.addEventListener('dblclick', (e) => {
            if (e.shiftKey || e.target.closest('.card-action-btn') || e.target.closest('.card-tag.is-removable') || e.target.closest('.card-bottom-right.is-inline-editable')) {
                return;
            }

            if (pendingCardClickTimer) {
                clearTimeout(pendingCardClickTimer);
                pendingCardClickTimer = null;
            }

            if (this.interactionManager && typeof this.interactionManager.cancelPendingEditorActivation === 'function') {
                this.interactionManager.cancelPendingEditorActivation();
            }

            this.clearCardSelection();
            this.clearSelectedCardTag();
            this.handleCardDoubleClick(e, item);
        });

        // Right-click: show card context menu
        card.addEventListener('contextmenu', (e) => {
            if (pendingCardClickTimer) {
                clearTimeout(pendingCardClickTimer);
                pendingCardClickTimer = null;
            }

            if (e.shiftKey && this.interactionManager && typeof this.interactionManager.handleShiftTableContextSelection === 'function') {
                const handled = this.interactionManager.handleShiftTableContextSelection(e, card);
                if (handled) {
                    return;
                }
            }

            e.preventDefault();
            e.stopPropagation();
            if (this.interactionManager && typeof this.interactionManager.showCardContextMenu === 'function') {
                this.interactionManager.showCardContextMenu(e, item, card);
            }
        });
        
        // Add action button handlers
        const commentsBtn = card.querySelector('.comments-btn');
        
        if (commentsBtn) {
            commentsBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                this.openCommentsDialog(item);
            });
        }
    }
    
    makeDropZone(cell, rowValue, colValue) {
        const isItemDrag = (event) => {
            const types = event.dataTransfer ? Array.from(event.dataTransfer.types || []) : [];
            return types.includes('application/x-item-id');
        };

        const getDropTargetValues = () => ({
            rowValue: Object.prototype.hasOwnProperty.call(cell.dataset, 'rowValue')
                ? cell.dataset.rowValue
                : rowValue,
            colValue: Object.prototype.hasOwnProperty.call(cell.dataset, 'colValue')
                ? cell.dataset.colValue
                : colValue
        });

        const logHighlightedDropZone = (event) => {
            const targetValues = getDropTargetValues();
            const zoneKey = `${targetValues.rowValue}:${targetValues.colValue}`;

            if (this.debugHighlightedDropZoneKey === zoneKey) {
                return;
            }

            this.debugHighlightedDropZoneKey = zoneKey;
            console.log('[CARD DRAG] highlighted drop-zone', {
                pointer: {
                    x: event.clientX,
                    y: event.clientY
                },
                dropTarget: {
                    rowValue: targetValues.rowValue,
                    colValue: targetValues.colValue,
                    gridRow: cell.style.gridRow,
                    gridColumn: cell.style.gridColumn
                }
            });
        };

        cell.addEventListener('click', (e) => {
            if (e.target.closest('.card') || e.target.closest('.grid-cell-table')) {
                return;
            }

            const xField = this.currentAxisSelections.x;
            const yField = this.currentAxisSelections.y;
            const initialValues = {};
            const targetValues = getDropTargetValues();

            if (xField) {
                initialValues[xField] = targetValues.colValue === '' ? null : targetValues.colValue;
            }

            if (yField) {
                initialValues[yField] = targetValues.rowValue === '' ? null : targetValues.rowValue;
            }

            if (typeof this.openCreatePanel === 'function') {
                this.openCreatePanel(initialValues);
            } else {
                this.openCreateTooltip(e, initialValues);
            }
        });

        cell.addEventListener('dragover', (e) => {
            if (!isItemDrag(e)) {
                return;
            }
            e.preventDefault();
            e.stopPropagation();
            e.dataTransfer.dropEffect = 'move';
            cell.classList.add('drop-zone');
            logHighlightedDropZone(e);
        });
        
        cell.addEventListener('dragleave', (e) => {
            if (!isItemDrag(e)) {
                return;
            }
            if (!cell.contains(e.relatedTarget)) {
                cell.classList.remove('drop-zone');

                const targetValues = getDropTargetValues();
                const zoneKey = `${targetValues.rowValue}:${targetValues.colValue}`;
                if (this.debugHighlightedDropZoneKey === zoneKey) {
                    this.clearDebugDropZone();
                }
            }
        });
        
        cell.addEventListener('drop', (e) => {
            if (!isItemDrag(e)) {
                return;
            }
            e.preventDefault();
            e.stopPropagation();
            cell.classList.remove('drop-zone');
            const targetValues = getDropTargetValues();
            this.clearDebugDropZone();
            
            const itemId = e.dataTransfer.getData('application/x-item-id') ||
                e.dataTransfer.getData('text/plain');
            if (itemId) {
                console.log('[CARD DRAG] drop', {
                    itemId,
                    pointer: {
                        x: e.clientX,
                        y: e.clientY
                    },
                    dropTarget: {
                        rowValue: targetValues.rowValue,
                        colValue: targetValues.colValue,
                        gridRow: cell.style.gridRow,
                        gridColumn: cell.style.gridColumn
                    }
                });
                this.handleCardDrop(itemId, targetValues.rowValue, targetValues.colValue);
            }
        });
    }

    getOrCreateCardDragImage() {
        if (this.cardDragImage && document.body.contains(this.cardDragImage)) {
            return this.cardDragImage;
        }

        const dragImage = document.createElement('div');
        dragImage.className = 'card-drag-image';
        dragImage.setAttribute('aria-hidden', 'true');
        document.body.appendChild(dragImage);
        this.cardDragImage = dragImage;
        return dragImage;
    }

    clearDebugDropZone() {
        if (!this.debugHighlightedDropZoneKey) {
            return;
        }

        console.log('[CARD DRAG] clear highlighted drop-zone', {
            dropZone: this.debugHighlightedDropZoneKey
        });
        this.debugHighlightedDropZoneKey = null;
    }
    
    calculateColumnAggregation(colValue, cells, allColumns) {
        const bottomRightField = this.currentCardSelections.bottomRight;
        let total = 0;
        let count = 0;
        
        // Sum all items in this column across all rows
        cells.forEach((items, cellKey) => {
            const [rowVal, colVal] = cellKey.split(':');
            if (colVal === colValue) {
                items.forEach(item => {
                    const value = this.getFieldValue(item, bottomRightField);
                    const numValue = parseFloat(value);
                    if (!isNaN(numValue)) {
                        total += numValue;
                        count++;
                    }
                });
            }
        });
        
        return count > 0 ? total.toFixed(1) : '0';
    }
    
    calculateRowAggregation(rowValue, cells, allRows) {
        const bottomRightField = this.currentCardSelections.bottomRight;
        let total = 0;
        let count = 0;
        
        // Sum all items in this row across all columns
        cells.forEach((items, cellKey) => {
            const [rowVal, colVal] = cellKey.split(':');
            if (rowVal === rowValue) {
                items.forEach(item => {
                    const value = this.getFieldValue(item, bottomRightField);
                    const numValue = parseFloat(value);
                    if (!isNaN(numValue)) {
                        total += numValue;
                        count++;
                    }
                });
            }
        });
        
        return count > 0 ? total.toFixed(1) : '0';
    }
    
    toggleColumnVisibility(gridColumn) {
        const dataGrid = document.querySelector('.data-grid');
        if (!dataGrid) return;
        
        const columnElements = dataGrid.querySelectorAll(`[style*="grid-column: ${gridColumn}"]`);
        const isCollapsed = columnElements[0]?.style.display === 'none';
        
        columnElements.forEach(element => {
            element.style.display = isCollapsed ? '' : 'none';
        });
        
        // Update collapse button
        const header = dataGrid.querySelector(`[style*="grid-row: 1"][style*="grid-column: ${gridColumn}"]`);
        const btn = header?.querySelector('.header-collapse-btn');
        if (btn) {
            btn.textContent = isCollapsed ? '−' : '+';
        }
    }
    
    toggleRowVisibility(gridRow) {
        const dataGrid = document.querySelector('.data-grid');
        if (!dataGrid) return;
        
        const rowElements = dataGrid.querySelectorAll(`[style*="grid-row: ${gridRow}"]`);
        const isCollapsed = rowElements[0]?.style.display === 'none';
        
        rowElements.forEach(element => {
            element.style.display = isCollapsed ? '' : 'none';
        });
        
        // Update collapse button
        const header = dataGrid.querySelector(`[style*="grid-row: ${gridRow}"][style*="grid-column: 1"]`);
        const btn = header?.querySelector('.header-collapse-btn');
        if (btn) {
            btn.textContent = isCollapsed ? '−' : '+';
        }
    }
    
    showNotification(message, type = 'info') {
        // Simple console notification for now
        console.log(`[${type.toUpperCase()}] ${message}`);
    }
    
    // ===== DATA MODIFICATION =====
    
    updateItem(itemId, updates, options = {}) {
        const { render = true } = options;
        const effectiveItem = this.getEffectiveItemById(itemId);
        if (!effectiveItem) {
            return false;
        }

        const identity = this.getItemIdentity(effectiveItem);
        if (!identity) {
            return false;
        }

        const normalizedUpdates = this.cloneItem(updates || {});
        const createChange = this.getPendingChangeRowForItem(identity, 'create');

        if (createChange) {
            const nextProposed = this.cloneItem(createChange.proposed || {});
            Object.entries(normalizedUpdates).forEach(([fieldName, nextValue]) => {
                nextProposed[fieldName] = this.normalizeFieldValueForStorage(fieldName, nextValue, nextProposed[fieldName]);
            });

            createChange.proposed = nextProposed;
            createChange.meta = this.createChangeMeta(createChange.meta);
            this.rebuildEffectiveDataset({ render });
            return true;
        }

        const baselineItem = this.getBaselineItemById(identity);
        if (!baselineItem) {
            return false;
        }

        let updateChange = this.getPendingChangeRowForItem(identity, 'update');
        if (!updateChange) {
            updateChange = {
                changeId: this.generateChangeId(),
                action: 'update',
                target: {
                    itemId: identity
                },
                baseline: {},
                proposed: {},
                meta: this.createChangeMeta()
            };

            if (baselineItem.sourceRef !== undefined) {
                updateChange.target.sourceRef = this.cloneValue(baselineItem.sourceRef);
            }

            this.ensureChangesContainer().rows.push(updateChange);
        }

        Object.entries(normalizedUpdates).forEach(([fieldName, nextValue]) => {
            const baselineValue = this.cloneValue(baselineItem[fieldName]);
            const normalizedNextValue = this.normalizeFieldValueForStorage(fieldName, nextValue, baselineValue);

            if (this.areFieldValuesEquivalent(fieldName, baselineValue, normalizedNextValue)) {
                delete updateChange.baseline[fieldName];
                delete updateChange.proposed[fieldName];
                return;
            }

            updateChange.baseline[fieldName] = baselineValue;
            updateChange.proposed[fieldName] = normalizedNextValue;
        });

        updateChange.meta = this.createChangeMeta(updateChange.meta);

        if (Object.keys(updateChange.proposed).length === 0) {
            this.removePendingChangeRow(updateChange.changeId);
        }

        this.rebuildEffectiveDataset({ render });
        return true;
    }
    
    addItem(newItem) {
        // Generate ID if not provided
        if (!newItem.id) {
            newItem.id = this.getNextGeneratedItemId();
        }

        const normalizedNewItem = {};
        Object.entries(this.cloneItem(newItem)).forEach(([fieldName, value]) => {
            normalizedNewItem[fieldName] = this.normalizeFieldValueForStorage(fieldName, value);
        });

        const newItemIdentity = String(normalizedNewItem.id);
        const createChange = {
            changeId: this.generateChangeId(),
            action: 'create',
            target: {
                itemId: newItemIdentity
            },
            baseline: {},
            proposed: this.cloneItem(normalizedNewItem),
            meta: this.createChangeMeta()
        };

        if (normalizedNewItem.sourceRef !== undefined) {
            createChange.target.sourceRef = this.cloneValue(normalizedNewItem.sourceRef);
        }

        this.ensureChangesContainer().rows.push(createChange);
        this.rebuildEffectiveDataset({ render: true });
        return newItem.id;
    }
    
    removeItem(itemId) {
        const effectiveItem = this.getEffectiveItemById(itemId);
        if (!effectiveItem) {
            return false;
        }

        const identity = this.getItemIdentity(effectiveItem);
        const createChange = this.getPendingChangeRowForItem(identity, 'create');

        if (createChange) {
            this.removePendingChangeRow(createChange.changeId);
            this.rebuildEffectiveDataset({ render: true });
            return true;
        }

        return false;
    }
}

// Initialize the application
let app;

document.addEventListener('DOMContentLoaded', () => {
    app = new DataVisualizationApp();
    if (typeof window !== 'undefined') {
        window.app = app;
    }
    console.log('Data Visualization Grid application initialized');
});

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { DataVisualizationApp };
}