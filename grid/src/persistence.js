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
        this.fileService.hydratePersistedHandles();
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
            this.fileService.debugLog('user-triggered data load started', {
                action: 'open-data-file',
                currentDataFilename: this.fileService.getLastOpenedFilename('data'),
            });
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
        this.fileService.debugLog('user-triggered data reload started', {
            action: 'reload-data-file',
            reloadAction,
            currentDataFilename: this.fileService.getLastOpenedFilename('data'),
        });

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
            this.fileService.debugLog('user-triggered view load started', {
                action: 'open-view-file',
                currentViewFilename: this.fileService.getLastOpenedFilename('view'),
                currentSourceContext: this.getCurrentSourceContext(),
            });
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
        if (!jsonData || typeof jsonData !== 'object' || Array.isArray(jsonData)) {
            return false;
        }

        return Object.prototype.hasOwnProperty.call(jsonData, 'axisSelections') || 
               Object.prototype.hasOwnProperty.call(jsonData, 'filters') ||
               Object.prototype.hasOwnProperty.call(jsonData, 'tagCustomizations') ||
               Object.prototype.hasOwnProperty.call(jsonData, 'cardClick') ||
               Object.prototype.hasOwnProperty.call(jsonData, 'urlConfig');
    }

    normalizeSourceContext(payload, fileName, options = {}) {
        const { serverSourceInfo = null } = options;
        const existingMeta = payload && payload.meta && typeof payload.meta === 'object' ? payload.meta : {};
        const existingSourceContext = existingMeta.sourceContext && typeof existingMeta.sourceContext === 'object'
            ? JSON.parse(JSON.stringify(existingMeta.sourceContext))
            : null;

        let inferredServerSourceContext = null;
        if (serverSourceInfo && typeof serverSourceInfo === 'object') {
            const sourceKind = String(serverSourceInfo.source_kind || '').trim();
            const workbookPath = String(serverSourceInfo.source_excel || serverSourceInfo.excel || '').trim();
            const obsidianConfigPath = String(serverSourceInfo.obsidian_config_path || serverSourceInfo.config_path || serverSourceInfo.source_path || '').trim();

            if (sourceKind === 'excel-cli-converted' || (!sourceKind && workbookPath)) {
                inferredServerSourceContext = {
                    kind: 'excel-cli-converted',
                    datasetFilename: fileName,
                    workbookPath: workbookPath || null,
                };
            } else if (sourceKind === 'obsidian-cli-converted' || (!sourceKind && obsidianConfigPath)) {
                inferredServerSourceContext = {
                    kind: 'obsidian-cli-converted',
                    datasetFilename: fileName,
                    obsidianConfigPath: obsidianConfigPath || null,
                };
            }
        }

        if (
            existingSourceContext &&
            existingSourceContext.kind &&
            !(existingSourceContext.kind === 'standalone-json' && inferredServerSourceContext)
        ) {
            this.fileService.debugLog('source context loaded from dataset metadata', {
                fileName,
                sourceContext: existingSourceContext,
            });
            return existingSourceContext;
        }

        if (inferredServerSourceContext) {
            this.fileService.debugLog('source context inferred from server source info', {
                fileName,
                sourceContext: inferredServerSourceContext,
                replacedExisting: existingSourceContext,
            });
            return inferredServerSourceContext;
        }

        const sourceContext = {
            kind: 'standalone-json',
            datasetFilename: fileName,
        };
        this.fileService.debugLog('source context defaulted to standalone json', {
            fileName,
            sourceContext,
        });
        return sourceContext;
    }

    getCurrentSourceContext(dataFilename = null) {
        const existingMeta = this.app.metaInfo && typeof this.app.metaInfo === 'object' ? this.app.metaInfo : {};
        const existingSourceContext = existingMeta.sourceContext && typeof existingMeta.sourceContext === 'object'
            ? JSON.parse(JSON.stringify(existingMeta.sourceContext))
            : null;

        if (existingSourceContext && existingSourceContext.kind) {
            return existingSourceContext;
        }

        return {
            kind: 'standalone-json',
            datasetFilename: dataFilename || this.fileService.getLastOpenedFilename('data') || null,
        };
    }

    deriveExpectedViewFilename(sourceContext, dataFilename = null) {
        const context = sourceContext && typeof sourceContext === 'object' ? sourceContext : {};
        const kind = String(context.kind || '').trim();
        const resolvedDataFilename = String(context.datasetFilename || dataFilename || '').trim();

        if (kind === 'excel-cli-converted') {
            const workbookPath = String(context.workbookPath || '').trim();
            if (!workbookPath) {
                return null;
            }
            const fileName = workbookPath.split(/[\\/]/).pop();
            return fileName ? `${fileName}.view.json` : null;
        }

        if (kind === 'obsidian-cli-converted') {
            return 'obsidian.view.json';
        }

        if (!resolvedDataFilename) {
            return null;
        }

        return this.fileService.deriveCompanionViewFilename(resolvedDataFilename);
    }

    /**
     * Derive the absolute path for the Obsidian view sidecar file.
     * Returns null if the source context is not Obsidian or has no config path.
     */
    deriveObsidianViewAbsolutePath(sourceContext) {
        const context = sourceContext && typeof sourceContext === 'object' ? sourceContext : {};
        if (String(context.kind || '') !== 'obsidian-cli-converted') return null;
        const configPath = String(context.obsidianConfigPath || '').trim();
        if (!configPath) return null;
        const dir = configPath.replace(/[\\/][^\\/]+$/, '');
        return dir + '/obsidian.view.json';
    }
    
    async processDataFileContent(content, fileName, options = {}) {
        try {
            const startedAt = performance.now();
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

            payload.meta = payload.meta && typeof payload.meta === 'object' ? payload.meta : {};
            payload.meta.sourceContext = this.normalizeSourceContext(payload, fileName, options);
            this.fileService.debugLog('processing data file content', {
                fileName,
                sourceContext: payload.meta.sourceContext,
                hasEmbeddedView: this.isViewFile(payload.view),
                durationMs: Math.round((performance.now() - startedAt) * 10) / 10,
            });
            
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
            this.fileService.debugLog('loaded dataset payload into application state', {
                fileName,
                itemCount: dataset.length,
                durationMs: Math.round((performance.now() - startedAt) * 10) / 10,
            });
            
            // Apply embedded view as the primary persisted view state.
            const hasEmbeddedView = this.isViewFile(payload.view);
            if (hasEmbeddedView) {
                this.applyViewConfiguration(payload.view);
            }

            // Legacy fallback: only try a companion .view.json when the dataset
            // does not already carry an embedded view.
            const companionView = await this.tryLoadCompanionView(fileName, payload.meta.sourceContext);
            if (companionView) {
                this.applyViewConfiguration(companionView.content);
                this.app.showNotification(
                    `Loaded ${dataset.length} items from ${fileName} · view from ${companionView.filename}`,
                    'success'
                );
            } else if (hasEmbeddedView) {
                this.fileService.debugLog('embedded view remains active because no external sidecar view was available', {
                    fileName,
                    sourceContext: payload.meta.sourceContext,
                });
            } else {
                this.app.showNotification(`Loaded ${dataset.length} items from ${fileName}`, 'success');
            }

            this.refreshFileActionState();
            this.updateGridFilePath(fileName, payload.meta.sourceContext);
            this.fileService.debugLog('completed processing data file content', {
                fileName,
                itemCount: dataset.length,
                durationMs: Math.round((performance.now() - startedAt) * 10) / 10,
            });
            
        } catch (error) {
            console.error('Error processing data file:', error);
            this.app.showNotification(`Invalid file format: ${error.message}`, 'error');
        }
    }
    
    updateGridFilePath(fileName, sourceContext) {
        const el = document.getElementById('grid-file-path');
        if (!el) return;
        // In desktop mode, use the full path from the cached bridge path
        if (this.app.isDesktopMode && window.desktopBridge) {
            const fullPath = window.desktopBridge.getCachedPath('data');
            if (fullPath) {
                el.textContent = fullPath;
                el.title = fullPath;
                return;
            }
        }
        const label = (sourceContext && sourceContext.filename) || fileName || '';
        el.textContent = label;
        el.title = label;
    }

    processViewFileContent(content, fileName) {
        try {
            const viewConfig = typeof content === 'string' ? JSON.parse(content) : content;
            this.fileService.debugLog('processing view file content', {
                fileName,
                keys: Object.keys(viewConfig || {}),
            });
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
    async tryLoadCompanionView(dataFilename, sourceContext = null) {
        try {
            const expectedFilename = this.deriveExpectedViewFilename(sourceContext || this.getCurrentSourceContext(dataFilename), dataFilename);
            if (!expectedFilename) {
                this.fileService.debugLog('view sidecar resolution skipped: no expected filename could be derived', {
                    dataFilename,
                    sourceContext,
                });
                return null;
            }

            this.fileService.debugLog('resolved expected view sidecar filename', {
                dataFilename,
                expectedFilename,
                sourceContext,
            });

            const viewHandle = await this.fileService.getCachedFileHandle('view');
            if (!viewHandle) {
                this.fileService.debugLog('view sidecar not auto-loaded: no cached view handle is available', {
                    expectedFilename,
                    dataFilename,
                    sourceContext,
                });
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
                expectedFilename,
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
    
    async saveDataFile(options = {}) {
        try {
            const { promoteChanges = false } = options;
            const data = this.prepareDataForSaving({ promoteChanges });
            const itemCount = this.app.dataset ? this.app.dataset.length : 0;
            const filename = this.fileService.getSuggestedFilename('data', this.fileService.generateDataFilename(itemCount));
            this.fileService.debugLog('user-triggered data save started', {
                action: promoteChanges ? 'promote-save-data-file' : 'save-data-file',
                filename,
                itemCount,
                changeRows: Array.isArray(data && data.changes && data.changes.rows) ? data.changes.rows.length : 0,
                relationChanges: Array.isArray(data && data.changes && data.changes.relations) ? data.changes.relations.length : 0,
                sourceContext: data && data.meta ? data.meta.sourceContext : null,
            });
            
            const savedFilename = await this.fileService.saveFile(data, filename, 'data');
            if (savedFilename) {
                if (promoteChanges && typeof this.app.applyPromotedDataState === 'function') {
                    this.app.applyPromotedDataState(data);
                }

                this.app.showNotification(
                    promoteChanges
                        ? `Promoted changes and saved data file: ${savedFilename}`
                        : `Saved data file: ${savedFilename}`,
                    'success'
                );
            }
        } catch (error) {
            console.error('Error saving data file:', error);
            this.app.showNotification(`Failed to save data: ${error.message}`, 'error');
        } finally {
            this.refreshFileActionState();
        }
    }

    async saveViewConfiguration() {
        try {
            const sourceContext = this.getCurrentSourceContext();
            console.log('[VIEW-SAVE] saveViewConfiguration triggered', {
                kind: sourceContext?.kind,
                desktopMode: this.fileService.isDesktopMode,
                obsidianConfigPath: sourceContext?.obsidianConfigPath || null,
                datasetFilename: sourceContext?.datasetFilename || null,
            });

            // Desktop + Obsidian: write directly to the derived absolute path
            if (this.fileService.isDesktopMode) {
                const obsViewPath = this.deriveObsidianViewAbsolutePath(sourceContext);
                if (obsViewPath) {
                    const viewConfig = this.prepareViewConfigForSaving();
                    const result = await window.desktopBridge.writeFile(
                        obsViewPath,
                        JSON.stringify(viewConfig, null, 2),
                    );
                    if (result.success) {
                        const savedName = obsViewPath.split(/[\\/]/).pop();
                        this.app.showNotification(`Saved view config: ${savedName}`, 'success');
                    } else {
                        this.app.showNotification(`Failed to save view: ${result.error}`, 'error');
                    }
                    return;
                }
            }

            const expectedFilename = this.deriveExpectedViewFilename(sourceContext);
            const activeDataFilename = this.fileService.getLastOpenedFilename('data')
                || (this.app.serverControlsManager && this.app.serverControlsManager.currentDataset
                    ? this.app.serverControlsManager.currentDataset.filename
                    : null);

            this.fileService.debugLog('save view requested', {
                activeDataFilename,
                sourceContext,
                expectedFilename,
            });

            if (expectedFilename) {
                const viewConfig = this.prepareViewConfigForSaving();
                const savedFilename = await this.fileService.saveFile(viewConfig, expectedFilename, 'view', {
                    preferProvidedFilename: true,
                });
                if (savedFilename) {
                    this.fileService.rememberSavedFilename('view', savedFilename);
                    this.app.showNotification(`Saved view config: ${savedFilename}`, 'success');
                }
                return;
            }

            if (activeDataFilename) {
                const result = await this.saveCurrentDatasetWithEmbeddedView(activeDataFilename);
                if (result) {
                    this.app.showNotification(`Saved view into data file: ${result}`, 'success');
                }
                return;
            }

            const viewConfig = this.prepareViewConfigForSaving();
            const xAxis = this.app.currentAxisSelections.x || 'none';
            const yAxis = this.app.currentAxisSelections.y || 'none';
            const fallbackFilename = this.fileService.generateViewFilename(xAxis, yAxis);
            const filename = this.fileService.getSuggestedFilename('view', fallbackFilename);

            const savedFilename = await this.fileService.saveFile(viewConfig, filename, 'view');
            if (savedFilename) {
                this.app.showNotification(`Saved standalone view config: ${savedFilename}`, 'success');
            }
        } catch (error) {
            console.error('Error saving view configuration:', error);
            this.app.showNotification(`Failed to save view config: ${error.message}`, 'error');
        } finally {
            this.refreshFileActionState();
        }
    }

    async saveCurrentDatasetWithEmbeddedView(dataFilename = null) {
        const targetFilename = dataFilename
            || this.fileService.getLastOpenedFilename('data')
            || (this.app.serverControlsManager && this.app.serverControlsManager.currentDataset
                ? this.app.serverControlsManager.currentDataset.filename
                : null);

        if (!targetFilename) {
            throw new Error('No active data file is available');
        }

        if (this.app.isServerMode && this.app.serverControlsManager && this.app.serverControlsManager.currentDataset) {
            this.fileService.debugLog('saving embedded view through active server dataset', {
                targetFilename,
                currentDataset: this.app.serverControlsManager.currentDataset.filename,
            });
            const result = await this.app.serverControlsManager.saveCurrentDataset({
                suppressSuccessNotification: true,
            });
            if (!result.ok) {
                throw result.error || new Error('Server save failed');
            }
            return this.app.serverControlsManager.currentDataset.filename;
        }

        const dataPayload = this.prepareDataForSaving();
        const cachedDataHandle = await this.fileService.getCachedFileHandle('data');
        this.fileService.debugLog('saving embedded view through active data file', {
            targetFilename,
            cachedDataHandle: cachedDataHandle ? cachedDataHandle.name : null,
        });
        if (cachedDataHandle && cachedDataHandle.name === targetFilename) {
            const success = await this.fileService.quickSave(dataPayload, 'data');
            if (success) {
                this.fileService.rememberSavedFilename('data', targetFilename);
                return targetFilename;
            }
        }

        return await this.fileService.saveFile(dataPayload, targetFilename, 'data', {
            preferProvidedFilename: true,
        });
    }
    
    // ===== DATA PREPARATION =====
    
    prepareDataForSaving(options = {}) {
        const { promoteChanges = false } = options;
        const now = window.GridDateUtils.createLocalTimestamp();
        const existingMeta = this.app.metaInfo || {};
        const datasetCardClick = typeof this.app.getDatasetCardClickConfig === 'function'
            ? this.app.getDatasetCardClickConfig()
            : (existingMeta.cardClick || null);
        const baselineData = promoteChanges && typeof this.app.getPromotedDataSnapshot === 'function'
            ? this.app.getPromotedDataSnapshot()
            : (typeof this.app.getBaselineDataSnapshot === 'function'
                ? this.app.getBaselineDataSnapshot()
                : (this.app.baselineData || []));
        const effectiveItemCount = this.app.dataset ? this.app.dataset.length : 0;

        // Pre-filter raw pending changes to drop semantic no-ops before normalization
        const rawChanges = this.app.pendingChanges || { version: '1', rows: [] };
        this.app.pendingChanges = this.dropSemanticNoOpChanges(rawChanges);

        const normalizedPendingChanges = promoteChanges
            ? {
                version: String(rawChanges && rawChanges.version ? rawChanges.version : '1'),
                rows: [],
                relations: []
            }
            : (typeof this.app.getPendingChangesSnapshot === 'function'
                ? this.app.getPendingChangesSnapshot()
                : (this.app.pendingChanges || { version: '1', rows: [] }));

        // Restore original pending changes
        this.app.pendingChanges = rawChanges;

        return {
            data: baselineData,
            changes: normalizedPendingChanges,
            view: this.prepareViewConfigForSaving(),
            schema: this.generateSchemaFromData(),
            meta: {
                ...existingMeta,
                sourceContext: this.getCurrentSourceContext(this.fileService.getLastOpenedFilename('data')),
                cardClick: datasetCardClick,
                version: existingMeta.version || '1.0',
                created: window.GridDateUtils.normalizeLocalTimestamp(existingMeta.created, now),
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
            const configuredValueColors = typeof this.app.getConfiguredFieldValueColors === 'function'
                ? this.app.getConfiguredFieldValueColors(fieldName)
                : null;
            if (configuredValueColors && Object.keys(configuredValueColors).length > 0) {
                schema.fields[fieldName].valueColors = JSON.parse(JSON.stringify(configuredValueColors));
            }
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
        const fieldValueColors = {};
        Object.keys(this.app.schemaFields || {}).forEach((fieldName) => {
            const configuredValueColors = typeof this.app.getConfiguredFieldValueColors === 'function'
                ? this.app.getConfiguredFieldValueColors(fieldName)
                : null;
            if (configuredValueColors && Object.keys(configuredValueColors).length > 0) {
                fieldValueColors[fieldName] = JSON.parse(JSON.stringify(configuredValueColors));
            }
        });

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
            fieldValueColors,
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

            if (Object.prototype.hasOwnProperty.call(normalizedViewConfig, 'fieldValueColors')) {
                Object.values(this.app.schemaFields || {}).forEach((fieldConfig) => {
                    if (fieldConfig && typeof fieldConfig === 'object' && Object.prototype.hasOwnProperty.call(fieldConfig, 'valueColors')) {
                        delete fieldConfig.valueColors;
                    }
                });

                const serializedFieldValueColors = normalizedViewConfig.fieldValueColors && typeof normalizedViewConfig.fieldValueColors === 'object'
                    ? normalizedViewConfig.fieldValueColors
                    : {};

                Object.entries(serializedFieldValueColors).forEach(([fieldName, valueColors]) => {
                    if (!this.app.schemaFields[fieldName]) {
                        this.app.schemaFields[fieldName] = {};
                    }

                    if (valueColors && typeof valueColors === 'object' && Object.keys(valueColors).length > 0) {
                        this.app.schemaFields[fieldName].valueColors = JSON.parse(JSON.stringify(valueColors));
                    }
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
                if (this.app.controlPanelManager && typeof this.app.controlPanelManager.syncModeConfigurationVisibility === 'function') {
                    this.app.controlPanelManager.syncModeConfigurationVisibility();
                }
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
            console.log('[VIEW-SAVE] persistViewConfiguration', {
                axes: viewConfig.axisSelections,
                filterCount: Object.keys(viewConfig.filters || {}).length,
                cellRenderMode: viewConfig.cellRenderMode,
                desktopMode: this.fileService.isDesktopMode,
                timestamp: viewConfig.created,
            });
            localStorage.setItem('grid-view-config', JSON.stringify(viewConfig));

            // Debounced write to obsidian view sidecar in desktop mode
            if (this.fileService.isDesktopMode) {
                this._scheduleObsidianViewSave(viewConfig);
            }
        } catch (error) {
            console.error('Error persisting view configuration:', error);
        }
    }

    _scheduleObsidianViewSave(viewConfig) {
        if (this._obsViewTimer) clearTimeout(this._obsViewTimer);
        this._obsViewTimer = setTimeout(() => {
            this._writeObsidianView(viewConfig);
        }, 1500);
    }

    async _writeObsidianView(viewConfig) {
        try {
            const sourceContext = this.getCurrentSourceContext();
            const obsViewPath = this.deriveObsidianViewAbsolutePath(sourceContext);
            if (!obsViewPath) {
                console.log('[OBSIDIAN-VIEW] Skipped — no obsidian view path', {
                    kind: sourceContext?.kind,
                    obsidianConfigPath: sourceContext?.obsidianConfigPath || null,
                });
                return;
            }
            const jsonStr = JSON.stringify(viewConfig, null, 2);
            console.log('[OBSIDIAN-VIEW] Writing auto-save', {
                path: obsViewPath,
                sizeBytes: jsonStr.length,
                axes: viewConfig.axisSelections,
                filterCount: Object.keys(viewConfig.filters || {}).length,
                cellRenderMode: viewConfig.cellRenderMode,
                timestamp: viewConfig.created,
            });
            const result = await window.desktopBridge.writeFile(
                obsViewPath,
                jsonStr,
            );
            console.log('[OBSIDIAN-VIEW] Auto-save result', {
                path: obsViewPath,
                success: result.success,
                error: result.success ? null : result.error,
            });
        } catch (err) {
            console.warn('[OBSIDIAN-VIEW] Auto-save failed:', err.message);
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

    async savePromotedData() {
        return this.saveDataFile({ promoteChanges: true });
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

    targetApp.promoteChanges = function() {
        return this.initializePersistence().saveDataFile({ promoteChanges: true });
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