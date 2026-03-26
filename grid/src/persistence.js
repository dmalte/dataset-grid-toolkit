// Copyright (c) 2026 Malte Doerper. MIT License. See LICENSE file.

// Data Persistence Layer - Uses Centralized File Service
// Handles loading/saving JSON data and view configurations via FileService

class PersistenceManager {
    constructor(app) {
        this.app = app;
        this.fileService = new FileService({
            getCurrentUserName: () => this.app.getCurrentUserName()
        });
        this.initializeFileOperations();
    }
    
    initializeFileOperations() {
        this.setupFileInputHandlers();
        this.setupDragDropFileHandling();
        this.loadPersistedViewConfiguration();
        this.refreshFileActionState();
        this.fileService.logPageLoadState().catch((error) => {
            console.warn('[FILE HANDLER] Failed to log page load state:', error);
        });
    }
    
    setupFileInputHandlers() {
        // File input handlers are now managed by the control panel
        // This method is kept for backward compatibility
    }
    
    setupDragDropFileHandling() {
        const dropZone = document.body;

        const isFileDrag = (event) => {
            const types = event.dataTransfer ? Array.from(event.dataTransfer.types || []) : [];
            return types.includes('Files');
        };
        
        dropZone.addEventListener('dragover', (e) => {
            if (!isFileDrag(e)) {
                return;
            }
            e.preventDefault();
            e.dataTransfer.dropEffect = 'copy';
            document.body.classList.add('drag-over');
        });
        
        dropZone.addEventListener('dragleave', (e) => {
            if (!isFileDrag(e)) {
                return;
            }
            if (!e.relatedTarget || !dropZone.contains(e.relatedTarget)) {
                document.body.classList.remove('drag-over');
            }
        });
        
        dropZone.addEventListener('drop', (e) => {
            if (!isFileDrag(e)) {
                return;
            }
            e.preventDefault();
            document.body.classList.remove('drag-over');
            this.handleFileDrop(e);
        });
    }
    
    // ===== FILE LOADING OPERATIONS =====
    
    async openDataFile() {
        try {
            const result = await this.fileService.openFile('data');
            if (!result) {
                return; // User cancelled
            }
            
            await this.processDataFileContent(result.content, result.filename);
        } catch (error) {
            console.error('Error opening data file:', error);
            this.app.showNotification(`Failed to open file: ${error.message}`, 'error');
        } finally {
            this.refreshFileActionState();
        }
    }

    async reloadDataFile() {
        const reloadAction = this.fileService.getReloadAction('data');

        if (reloadAction === 'reselect') {
            this.app.showNotification('Select the dataset again to reload it in this browser', 'info');
            await this.openDataFile();
            return;
        }

        try {
            const result = await this.fileService.reopenFile('data');
            if (!result) {
                return;
            }

            await this.processDataFileContent(result.content, result.filename);
            this.app.showNotification(`Reloaded ${result.filename}`, 'success');
        } catch (error) {
            console.error('Error reloading data file:', error);
            this.app.showNotification(`Failed to reload data file: ${error.message}`, 'warning');
        } finally {
            this.refreshFileActionState();
        }
    }
    
    async openViewFile() {
        try {
            const result = await this.fileService.openFile('view');
            if (!result) {
                return; // User cancelled
            }
            
            this.processViewFileContent(result.content, result.filename);
        } catch (error) {
            console.error('Error opening view file:', error);
            this.app.showNotification(`Failed to open view file: ${error.message}`, 'error');
        } finally {
            this.refreshFileActionState();
        }
    }
    
    async handleFileDrop(event) {
        try {
            const results = await this.fileService.handleFileDrop(event);
            if (!results || results.length === 0) {
                return; // No file dropped or not supported
            }

            for (const result of results) {
                const content = result.content;
                const fileName = result.filename;

                if (this.isDataFile(content)) {
                    await this.processDataFileContent(content, fileName);
                } else if (this.isViewFile(content)) {
                    this.processViewFileContent(content, fileName);
                } else {
                    this.app.showNotification(`File format not recognized: ${fileName}`, 'warning');
                }
            }
        } catch (error) {
            console.error('Error handling dropped file:', error);
            this.app.showNotification(`Error processing file: ${error.message}`, 'error');
        } finally {
            this.refreshFileActionState();
        }
    }
    
    isDataFile(jsonData) {
        return jsonData.hasOwnProperty('data') && Array.isArray(jsonData.data);
    }
    
    isViewFile(jsonData) {
        return jsonData.hasOwnProperty('axisSelections') || 
               jsonData.hasOwnProperty('filters') ||
               jsonData.hasOwnProperty('tagCustomizations') ||
               jsonData.hasOwnProperty('cardClick') ||
               jsonData.hasOwnProperty('urlConfig');
    }
    
    async processDataFileContent(content, fileName) {
        try {
            const jsonData = typeof content === 'string' ? JSON.parse(content) : content;
            
            // Handle different data formats
            let payload;
            if (jsonData.data && Array.isArray(jsonData.data)) {
                payload = jsonData;  // Full app format
            } else if (Array.isArray(jsonData)) {
                payload = {
                    data: jsonData,
                    changes: { version: '1', rows: [] },
                    view: this.app.viewConfig || {}
                };
            } else {
                throw new Error('Invalid data format - expected array of items');
            }
            
            const dataset = payload.data;
            if (dataset.length === 0) {
                this.app.showNotification('Dataset is empty', 'warning');
                return;
            }
            
            // Explicitly loaded datasets should start from their own embedded view
            // and optional companion .view.json, not the legacy global localStorage view.
            this.app.skipPendingViewConfigOnce = true;

            // Explicitly loaded datasets should also start from a clean header
            // collapse/order state. Embedded and companion views can still
            // restore gridState after the dataset is loaded.
            if (typeof this.app.resetHeaderState === 'function') {
                this.app.resetHeaderState();
            }

            // Load the data into the application
            this.app.loadDataFromJSON(payload);
            
            // Apply embedded view as initial default (if present)
            if (payload.view) {
                this.applyViewConfiguration(payload.view);
            }

            // Try to auto-load companion view file from cached handle.
            // A companion .view.json takes priority over the embedded view
            // so that item-manager re-exports don't overwrite user-saved views.
            const companionView = await this.tryLoadCompanionView(fileName);
            if (companionView) {
                this.applyViewConfiguration(companionView.content);
                this.app.showNotification(
                    `Loaded ${dataset.length} items from ${fileName} · view from ${companionView.filename}`,
                    'success'
                );
            } else {
                this.app.showNotification(`Loaded ${dataset.length} items from ${fileName}`, 'success');
            }

            this.refreshFileActionState();
            
        } catch (error) {
            console.error('Error processing data file:', error);
            this.app.showNotification(`Invalid file format: ${error.message}`, 'error');
        }
    }
    
    processViewFileContent(content, fileName) {
        try {
            const viewConfig = typeof content === 'string' ? JSON.parse(content) : content;
            this.applyViewConfiguration(viewConfig);
            this.app.showNotification(`Applied view configuration from ${fileName}`, 'success');
            this.refreshFileActionState();
            
        } catch (error) {
            console.error('Error processing view file:', error);
            this.app.showNotification(`Invalid view file: ${error.message}`, 'error');
        }
    }

    /**
     * Try to auto-load a companion .view.json from the cached view file handle.
     * Returns {content, filename} or null if no companion view is available.
     */
    async tryLoadCompanionView(dataFilename) {
        try {
            const expectedFilename = this.fileService.deriveCompanionViewFilename(dataFilename);
            if (!expectedFilename) {
                return null;
            }

            const viewHandle = await this.fileService.getCachedFileHandle('view');
            if (!viewHandle) {
                return null;
            }

            if (viewHandle.name !== expectedFilename) {
                this.fileService.debugLog('companion view skipped: cached view handle does not match dataset', {
                    expectedFilename,
                    cachedViewFilename: viewHandle.name,
                    dataFilename,
                    userScope: this.fileService.getCurrentUserScope()
                });
                return null;
            }

            const file = await viewHandle.getFile();
            const parsedContent = await this.fileService.readJSONFile(file);

            // Verify it looks like a view config
            if (!this.isViewFile(parsedContent)) {
                return null;
            }

            this.fileService.rememberOpenedFile('view', file.name);
            this.fileService.debugLog('auto-loaded companion view', {
                filename: file.name,
                userScope: this.fileService.getCurrentUserScope()
            });

            return { content: parsedContent, filename: file.name };
        } catch (error) {
            // Permission denied, handle expired, or file not found — all fine
            this.fileService.debugLog('companion view not available', {
                reason: error.message
            });
            return null;
        }
    }
    
    // ===== FILE SAVING OPERATIONS =====
    
    async saveDataFile() {
        try {
            const data = this.prepareDataForSaving();
            const itemCount = this.app.dataset ? this.app.dataset.length : 0;
            const filename = this.fileService.getSuggestedFilename('data', this.fileService.generateDataFilename(itemCount));
            
            const savedFilename = await this.fileService.saveFile(data, filename, 'data');
            if (savedFilename) {
                this.app.showNotification(`Saved data file: ${savedFilename}`, 'success');

                // Auto-save view to companion .view.json if a view handle is cached
                await this.autoSaveCompanionView(savedFilename);
            }
        } catch (error) {
            console.error('Error saving data file:', error);
            this.app.showNotification(`Failed to save data: ${error.message}`, 'error');
        } finally {
            this.refreshFileActionState();
        }
    }

    /**
     * Auto-save view config to companion .view.json via quick-save.
     * Only writes if a view file handle is already cached (no picker shown).
     * @param {string} dataFilename - The data filename just saved
     */
    async autoSaveCompanionView(dataFilename) {
        try {
            const expectedFilename = this.fileService.deriveCompanionViewFilename(dataFilename);
            if (!expectedFilename) {
                return;
            }

            const cachedViewHandle = await this.fileService.getCachedFileHandle('view');
            if (!cachedViewHandle || cachedViewHandle.name !== expectedFilename) {
                this.fileService.debugLog('auto-save companion view skipped', {
                    reason: 'cached view handle does not match companion filename',
                    expectedFilename,
                    cachedViewFilename: cachedViewHandle ? cachedViewHandle.name : null,
                    dataFilename,
                    userScope: this.fileService.getCurrentUserScope()
                });
                return;
            }

            const viewConfig = this.prepareViewConfigForSaving();
            const success = await this.fileService.quickSave(viewConfig, 'view');
            if (success) {
                this.fileService.debugLog('auto-saved companion view', {
                    viewFilename: expectedFilename,
                    dataFilename,
                    userScope: this.fileService.getCurrentUserScope()
                });
            }
        } catch (error) {
            // Non-critical — silent failure
            this.fileService.debugLog('auto-save companion view skipped', {
                reason: error.message
            });
        }
    }
    
    async saveViewConfiguration() {
        try {
            const viewConfig = this.prepareViewConfigForSaving();
            const xAxis = this.app.currentAxisSelections.x || 'none';
            const yAxis = this.app.currentAxisSelections.y || 'none';

            // Prefer companion filename derived from the data file (e.g. "foo.view.json")
            const dataFilename = this.fileService.getLastOpenedFilename('data');
            const companionFilename = dataFilename
                ? this.fileService.deriveCompanionViewFilename(dataFilename)
                : null;
            const fallbackFilename = this.fileService.generateViewFilename(xAxis, yAxis);
            const filename = companionFilename || this.fileService.getSuggestedFilename('view', fallbackFilename);
            
            const savedFilename = await this.fileService.saveFile(viewConfig, filename, 'view');
            if (savedFilename) {
                this.app.showNotification(`Saved view config: ${savedFilename}`, 'success');
            }
        } catch (error) {
            console.error('Error saving view configuration:', error);
            this.app.showNotification(`Failed to save view config: ${error.message}`, 'error');
        } finally {
            this.refreshFileActionState();
        }
    }
    
    // ===== DATA PREPARATION =====
    
    prepareDataForSaving() {
        const now = window.GridDateUtils.createLocalTimestamp();
        const existingMeta = this.app.metaInfo || {};
        const datasetCardClick = typeof this.app.getDatasetCardClickConfig === 'function'
            ? this.app.getDatasetCardClickConfig()
            : (existingMeta.cardClick || null);
        const baselineData = typeof this.app.getBaselineDataSnapshot === 'function'
            ? this.app.getBaselineDataSnapshot()
            : (this.app.baselineData || []);
        const effectiveItemCount = this.app.dataset ? this.app.dataset.length : 0;

        // Pre-filter raw pending changes to drop semantic no-ops before normalization
        const rawChanges = this.app.pendingChanges || { version: '1', rows: [] };
        this.app.pendingChanges = this.dropSemanticNoOpChanges(rawChanges);

        const normalizedPendingChanges = typeof this.app.getPendingChangesSnapshot === 'function'
            ? this.app.getPendingChangesSnapshot()
            : (this.app.pendingChanges || { version: '1', rows: [] });

        // Restore original pending changes
        this.app.pendingChanges = rawChanges;

        return {
            data: baselineData,
            changes: normalizedPendingChanges,
            schema: this.generateSchemaFromData(),
            meta: {
                ...existingMeta,
                cardClick: datasetCardClick,
                version: existingMeta.version || '1.0',
                created: existingMeta.created || now,
                lastModified: now,
                itemCount: effectiveItemCount
            }
        };
    }

    dropSemanticNoOpChanges(pendingChanges) {
        if (!pendingChanges || !Array.isArray(pendingChanges.rows)) {
            return pendingChanges;
        }

        const filteredRows = pendingChanges.rows.filter(row => {
            if (!row || row.action !== 'update') return true;

            const baseline = row.baseline || {};
            const proposed = row.proposed || {};
            const allFields = new Set([...Object.keys(baseline), ...Object.keys(proposed)]);

            for (const field of allFields) {
                const bVal = baseline[field];
                const pVal = proposed[field];
                if (String(bVal) !== String(pVal)) return true;
            }
            return false;
        });

        return { ...pendingChanges, rows: filteredRows };
    }
    
    generateSchemaFromData() {
        const schema = {
            fields: {},
            fieldTypes: {}
        };
        
        // Generate schema from field analysis, preserving imported metadata
        this.app.fieldTypes.forEach((type, fieldName) => {
            schema.fieldTypes[fieldName] = type;
            const imported = this.app.schemaFields[fieldName] || {};
            
            schema.fields[fieldName] = {
                type: type,
                required: false,
                distinctValues: this.app.distinctValues.get(fieldName) || []
            };
            // Preserve imported schema properties on round-trip
            if (imported.kind !== undefined) schema.fields[fieldName].kind = imported.kind;
            if (imported.editable !== undefined) schema.fields[fieldName].editable = imported.editable;
            if (imported.selectable !== undefined) schema.fields[fieldName].selectable = imported.selectable;
            if (imported.visible !== undefined) schema.fields[fieldName].visible = imported.visible;
            if (Array.isArray(imported.validValues)) schema.fields[fieldName].validValues = imported.validValues;
        });

        // V4: Include relation type registry if present
        if (Array.isArray(this.app.relationTypes) && this.app.relationTypes.length > 0) {
            schema.relationTypes = [...this.app.relationTypes];
        }
        
        return schema;
    }
    
    prepareViewConfigForSaving() {
        const serializedFilters = {};
        this.app.currentFilters.forEach((values, field) => {
            const normalizedValues = values instanceof Set ? values : new Set(Array.isArray(values) ? values : []);
            serializedFilters[field] = Array.from(normalizedValues);
        });

        const serializedTags = Object.fromEntries(this.app.tagCustomizations);

        return {
            axisSelections: this.app.currentAxisSelections,
            cardSelections: this.app.currentCardSelections,
            tableColumns: Array.isArray(this.app.tableColumnFields)
                ? [...this.app.tableColumnFields]
                : [],
            filters: serializedFilters,
            slicerFilters: serializedFilters,
            tagCustomizations: serializedTags,
            tags: serializedTags,
            cardClick: typeof this.app.getViewCardClickConfig === 'function'
                ? this.app.getViewCardClickConfig()
                : (this.app.viewConfig && this.app.viewConfig.cardClick) || null,
            urlConfig: this.app.interactionManager ? this.app.interactionManager.urlConfig : {},
            groups: Array.isArray(this.app.viewConfig.groups)
                ? JSON.parse(JSON.stringify(this.app.viewConfig.groups))
                : [],
            groupSlicerFilter: this.app.viewConfig.groupSlicerFilter || null,
            gridState: {
                collapsedHeaders: this.getCollapsedHeaders(),
                columnOrder: this.getColumnOrder(),
                rowOrder: this.getRowOrder()
            },
            relations: this.app.relationUIManager ? {
                focusMode: this.app.relationUIManager.focusMode || false,
                focusRootId: this.app.relationUIManager.focusRootId || null,
                focusDepth: this.app.relationUIManager.focusDepth || 1,
                focusTypes: this.app.relationUIManager.focusTypes || null
            } : null,
            showRelationshipFields: this.app.showRelationshipFields || false,
            showDerivedFields: this.app.showDerivedFields || false,
            showTooltips: this.app.showTooltips || false,
            growCards: this.app.growCards || false,
            cellRenderMode: this.app.cellRenderMode || 'cards',
            tableSummaryMode: typeof this.app.normalizeTableSummaryMode === 'function'
                ? this.app.normalizeTableSummaryMode(this.app.tableSummaryMode)
                : 'none',
            created: window.GridDateUtils.createLocalTimestamp()
        };
    }
    
    getCollapsedHeaders() {
        // Extract collapsed header state from grid
        const collapsedHeaders = {
            rows: [],
            columns: []
        };
        
        document.querySelectorAll('.grid-header.collapsed').forEach(header => {
            const value = header.dataset.value;
            const type = header.dataset.headerType;
            
            if (type === 'row') {
                collapsedHeaders.rows.push(value);
            } else if (type === 'column') {
                collapsedHeaders.columns.push(value);
            }
        });
        
        return collapsedHeaders;
    }
    
    getColumnOrder() {
        // Get current column order from grid
        const headers = document.querySelectorAll('.grid-header[data-header-type="column"]');
        return Array.from(headers).map(h => h.dataset.value);
    }
    
    getRowOrder() {
        // Get current row order from grid
        const headers = document.querySelectorAll('.grid-header[data-header-type="row"]');
        return Array.from(headers).map(h => h.dataset.value);
    }
    
    // ===== VIEW CONFIGURATION APPLICATION =====
    
    applyViewConfiguration(viewConfig) {
        try {
            const normalizedViewConfig = typeof this.app.normalizeViewConfig === 'function'
                ? this.app.normalizeViewConfig(viewConfig)
                : viewConfig;

            // Apply axis selections
            if (normalizedViewConfig.axisSelections) {
                Object.assign(this.app.currentAxisSelections, normalizedViewConfig.axisSelections);
            }

            if (normalizedViewConfig.cardSelections) {
                Object.assign(this.app.currentCardSelections, normalizedViewConfig.cardSelections);
            }

            if (Array.isArray(normalizedViewConfig.tableColumns)) {
                this.app.tableColumnFields = [...normalizedViewConfig.tableColumns];
            }
            
            // Apply filters
            if (normalizedViewConfig.filters) {
                this.app.currentFilters.clear();
                Object.entries(normalizedViewConfig.filters).forEach(([field, value]) => {
                    const normalizedValues = value instanceof Set
                        ? value
                        : new Set(Array.isArray(value) ? value : []);
                    this.app.currentFilters.set(field, normalizedValues);
                });
            }
            
            // Apply tag customizations
            const tagConfigs = normalizedViewConfig.tagCustomizations || normalizedViewConfig.tags;
            if (tagConfigs) {
                this.app.tagCustomizations.clear();
                Object.entries(tagConfigs).forEach(([tag, config]) => {
                    this.app.tagCustomizations.set(tag, config);
                });
            }

            if (Object.prototype.hasOwnProperty.call(normalizedViewConfig, 'cardClick') && typeof this.app.setViewCardClickConfig === 'function') {
                this.app.setViewCardClickConfig(normalizedViewConfig.cardClick);
            }
            
            // Apply grid state
            if (normalizedViewConfig.gridState) {
                // Apply collapsed headers
                if (this.app.advancedFeaturesManager && normalizedViewConfig.gridState.collapsedHeaders) {
                    this.app.advancedFeaturesManager.setCollapsedHeaders(normalizedViewConfig.gridState.collapsedHeaders);
                }
                
                // Apply header ordering
                if (this.app.advancedFeaturesManager && normalizedViewConfig.gridState.columnOrder && normalizedViewConfig.gridState.rowOrder) {
                    this.app.advancedFeaturesManager.setHeaderOrdering({
                        columns: normalizedViewConfig.gridState.columnOrder,
                        rows: normalizedViewConfig.gridState.rowOrder
                    });
                }
            }
            
            // Apply URL configuration
            if (normalizedViewConfig.urlConfig && this.app.interactionManager) {
                this.app.interactionManager.setUrlConfig(normalizedViewConfig.urlConfig);
            }

            // Apply groups
            if (Array.isArray(normalizedViewConfig.groups)) {
                this.app.viewConfig.groups = JSON.parse(JSON.stringify(normalizedViewConfig.groups));
            }

            // Apply group slicer filter
            if (Object.prototype.hasOwnProperty.call(normalizedViewConfig, 'groupSlicerFilter')) {
                this.app.viewConfig.groupSlicerFilter = normalizedViewConfig.groupSlicerFilter;
            }

            // Apply relations focus mode state
            if (normalizedViewConfig.relations && this.app.relationUIManager) {
                const rel = normalizedViewConfig.relations;
                if (rel.focusMode && rel.focusRootId) {
                    this.app.relationUIManager.enterFocusMode(
                        rel.focusRootId,
                        rel.focusDepth || 1,
                        rel.focusTypes || null
                    );
                }
            }

            // Apply filter toggle states
            if (Object.prototype.hasOwnProperty.call(normalizedViewConfig, 'showRelationshipFields')) {
                this.app.showRelationshipFields = normalizedViewConfig.showRelationshipFields;
                const cb = document.getElementById('show-relationship-fields-cb');
                if (cb) cb.checked = this.app.showRelationshipFields;
            }
            if (Object.prototype.hasOwnProperty.call(normalizedViewConfig, 'showDerivedFields')) {
                this.app.showDerivedFields = normalizedViewConfig.showDerivedFields;
                const cb = document.getElementById('show-derived-fields-cb');
                if (cb) cb.checked = this.app.showDerivedFields;
            }
            if (Object.prototype.hasOwnProperty.call(normalizedViewConfig, 'showTooltips')) {
                this.app.showTooltips = normalizedViewConfig.showTooltips;
                const cb = document.getElementById('show-tooltips-cb');
                if (cb) cb.checked = this.app.showTooltips;
            }
            if (Object.prototype.hasOwnProperty.call(normalizedViewConfig, 'growCards')) {
                this.app.growCards = normalizedViewConfig.growCards;
                const cb = document.getElementById('grow-cards-cb');
                if (cb) cb.checked = this.app.growCards;
            }
            if (Object.prototype.hasOwnProperty.call(normalizedViewConfig, 'cellRenderMode')) {
                this.app.cellRenderMode = normalizedViewConfig.cellRenderMode === 'table' ? 'table' : 'cards';
                const cb = document.getElementById('table-mode-cb');
                if (cb) cb.checked = this.app.cellRenderMode === 'table';
            }

            if (Object.prototype.hasOwnProperty.call(normalizedViewConfig, 'tableSummaryMode')) {
                this.app.tableSummaryMode = typeof this.app.normalizeTableSummaryMode === 'function'
                    ? this.app.normalizeTableSummaryMode(normalizedViewConfig.tableSummaryMode)
                    : 'none';
                const select = document.getElementById('table-summary-mode-select');
                if (select) select.value = this.app.tableSummaryMode;
            }
            
            // Update the stored view configuration
            this.app.viewConfig = {
                ...this.app.viewConfig,
                ...normalizedViewConfig
            };

            if (typeof this.app.reconcileFiltersWithDistinctValues === 'function') {
                this.app.reconcileFiltersWithDistinctValues(this.app.snapshotDistinctValues());
            } else {
                this.app.updateFilteredData();
            }

            // Resolve group memberships after data is filtered
            if (typeof this.app.resolveGroupMemberships === 'function') {
                this.app.resolveGroupMemberships();
            }
            
            // Re-render everything
            this.app.updateViewConfiguration();
            this.app.renderFieldSelectors();
            this.app.renderSlicers();
            this.app.renderGrid();
            
        } catch (error) {
            console.error('Error applying view configuration:', error);
            throw new Error(`Failed to apply view configuration: ${error.message}`);
        }
    }
    
    // ===== PERSISTENCE OF VIEW CONFIGURATION =====
    
    persistViewConfiguration() {
        try {
            const viewConfig = this.prepareViewConfigForSaving();
            localStorage.setItem('grid-view-config', JSON.stringify(viewConfig));
        } catch (error) {
            console.error('Error persisting view configuration:', error);
        }
    }
    
    loadPersistedViewConfiguration() {
        try {
            const saved = localStorage.getItem('grid-view-config');
            if (saved) {
                const viewConfig = JSON.parse(saved);
                
                // Only apply persisted configuration after the app has loaded data
                if (this.app.dataset && this.app.dataset.length > 0) {
                    this.applyViewConfiguration(viewConfig);
                } else {
                    // Store for later application
                    this.app.persistedViewConfig = viewConfig;
                }
            }
        } catch (error) {
            console.error('Error loading persisted view configuration:', error);
        }
    }
    
    applyPendingViewConfig() {
        if (this.app.persistedViewConfig) {
            try {
                this.applyViewConfiguration(this.app.persistedViewConfig);
                delete this.app.persistedViewConfig;
            } catch (error) {
                console.error('Error applying pending view configuration:', error);
            }
        }
    }
    
    // ===== PUBLIC INTERFACE METHODS =====
    
    // Expose save methods for use by other managers
    async saveData() {
        return this.saveDataFile();
    }
    
    async saveView() {
        return this.saveViewConfiguration();
    }
    
    // Quick save methods (if File System Access supported)
    async quickSaveData() {
        if (this.fileService.isFileSystemAccessSupported()) {
            try {
                const data = this.prepareDataForSaving();
                const success = await this.fileService.quickSave(data, 'data');
                if (success) {
                    this.app.showNotification('Data saved to existing file', 'success');
                } else {
                    // Fallback to regular save
                    await this.saveDataFile();
                }
            } catch (error) {
                console.error('Error in quick save:', error);
                this.app.showNotification(`Quick save failed: ${error.message}`, 'error');
            }
        } else {
            // Fallback to regular save
            await this.saveDataFile();
        }
    }
    
    async quickSaveView() {
        if (this.fileService.isFileSystemAccessSupported()) {
            try {
                const viewConfig = this.prepareViewConfigForSaving();
                const success = await this.fileService.quickSave(viewConfig, 'view');
                if (success) {
                    this.app.showNotification('View configuration saved to existing file', 'success');
                } else {
                    // Fallback to regular save
                    await this.saveViewConfiguration();
                }
            } catch (error) {
                console.error('Error in quick save:', error);
                this.app.showNotification(`Quick save failed: ${error.message}`, 'error');
            }
        } else {
            // Fallback to regular save
            await this.saveViewConfiguration();
        }
    }
    
    // Get file handle status for debugging
    getFileServiceStatus() {
        return this.fileService.getFileHandleStatus();
    }

    refreshFileActionState() {
        if (typeof this.app.updateFileActionState === 'function') {
            this.app.updateFileActionState(this.fileService.getFileHandleStatus());
        }
    }
    
    // Clear file handles (useful for debugging or reset)
    clearFileHandles() {
        this.fileService.clearFileHandles();
        this.refreshFileActionState();
        this.app.showNotification('File handles cleared', 'info');
    }
    
    // ===== AUTO-SAVE FUNCTIONALITY =====
    
    enableAutoSave(intervalMs = 30000) {
        this.autoSaveInterval = setInterval(() => {
            this.persistViewConfiguration();
        }, intervalMs);
    }
    
    disableAutoSave() {
        if (this.autoSaveInterval) {
            clearInterval(this.autoSaveInterval);
            this.autoSaveInterval = null;
        }
    }
}

// ===== INTEGRATION WITH EXISTING APP =====

function attachPersistenceIntegration(targetApp) {
    if (!targetApp || targetApp.__persistenceIntegrated) {
        return;
    }

    targetApp.initializePersistence = function() {
        if (!this.persistenceManager) {
            this.persistenceManager = new PersistenceManager(this);
            this.persistenceManager.enableAutoSave();
        }
        return this.persistenceManager;
    };

    targetApp.saveData = function() {
        return this.initializePersistence().saveDataFile();
    };

    targetApp.saveViewConfig = function() {
        return this.initializePersistence().saveViewConfiguration();
    };

    targetApp.exportViewConfig = function() {
        return this.initializePersistence().saveViewConfiguration();
    };

    targetApp.loadViewConfig = function(viewConfig) {
        const persistenceManager = this.initializePersistence();
        if (viewConfig) {
            return persistenceManager.applyViewConfiguration(viewConfig);
        }
        return persistenceManager.openViewFile();
    };

    const originalUpdateViewConfiguration = targetApp.updateViewConfiguration;
    targetApp.updateViewConfiguration = function() {
        originalUpdateViewConfiguration.call(this);

        if (this.persistenceManager) {
            this.persistenceManager.persistViewConfiguration();
        }
    };

    const originalLoadDataFromJSON = targetApp.loadDataFromJSON;
    targetApp.loadDataFromJSON = function(jsonData) {
        originalLoadDataFromJSON.call(this, jsonData);

        if (this.skipPendingViewConfigOnce) {
            delete this.skipPendingViewConfigOnce;
            return;
        }

        if (this.persistenceManager && this.persistedViewConfig) {
            this.persistenceManager.applyPendingViewConfig();
        }
    };

    targetApp.__persistenceIntegrated = true;
}

document.addEventListener('DOMContentLoaded', () => {
    if (typeof app !== 'undefined' && app) {
        attachPersistenceIntegration(app);
        app.initializePersistence();
    }
});