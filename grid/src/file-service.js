// Copyright (c) 2026 Malte Doerper. MIT License. See LICENSE file.

// Centralized File Service Module  
// Implements File System Access API with graceful fallback for browsers that don't support it
// Handles all file loading/saving operations for the data visualization app

class FileService {
    /** Desktop mode — true once the PySide6 QWebChannel bridge is available. */
    get isDesktopMode() {
        return typeof window.desktopBridge !== 'undefined';
    }

    constructor(options = {}) {
        // Check for File System Access API support
        this.supportsFileOpenPicker = 'showOpenFilePicker' in window;
        this.supportsFileSavePicker = 'showSaveFilePicker' in window;
        this.supportsFileSystemAccess = this.supportsFileOpenPicker && this.supportsFileSavePicker;
        this.supportsHandlePersistence = this.supportsFileSystemAccess && typeof indexedDB !== 'undefined';
        this.getCurrentUserName = typeof options.getCurrentUserName === 'function'
            ? options.getCurrentUserName
            : (() => 'Local User');
        this.handleDbName = 'data-visualization-grid-file-handles';
        this.handleStoreName = 'pickerHandles';
        this.handleDbPromise = null;
        
        // File handle storage for repeat operations
        this.fileHandles = {
            data: null,
            view: null
        };

        this.lastOpenedFiles = {
            data: null,
            view: null
        };

        this.lastSavedFilenames = {
            data: null,
            view: null
        };

        this.persistedHandleHydrationPromise = null;
        
        // Supported file types
        this.fileTypes = {
            json: {
                description: 'JSON files',
                accept: {
                    'application/json': ['.json']
                }
            },
            xlsx: {
                description: 'Excel workbooks',
                accept: {
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx']
                }
            }
        };

        this.pickerIds = {
            data: 'data-visualization-grid-data',
            view: 'data-visualization-grid-view'
        };

        this.debugTag = '[FILE HANDLER]';
    }

    debugLog(message, details) {
        if (typeof details === 'undefined') {
            console.debug(this.debugTag, message);
            return;
        }

        console.debug(this.debugTag, message, details);
    }

    describeHandle(handle) {
        if (!handle) {
            return null;
        }

        return {
            kind: handle.kind || 'unknown',
            name: handle.name || null
        };
    }

    async getPersistedHandleSummary(type) {
        if (type !== 'data' && type !== 'view') {
            return null;
        }

        const handle = await this.readPersistedHandle(type);
        return {
            userScope: this.getCurrentUserScope(),
            storage: this.supportsHandlePersistence
                ? {
                    database: this.handleDbName,
                    store: this.handleStoreName,
                    key: this.getPersistedHandleKey(type)
                }
                : null,
            type,
            handle: this.describeHandle(handle)
        };
    }

    async logPageLoadState() {
        const [dataHandle, viewHandle] = await Promise.all([
            this.getPersistedHandleSummary('data'),
            this.getPersistedHandleSummary('view')
        ]);

        this.debugLog('page load: persisted picker location state', {
            userScope: this.getCurrentUserScope(),
            supportsFileSystemAccess: this.supportsFileSystemAccess,
            supportsHandlePersistence: this.supportsHandlePersistence,
            note: this.supportsHandlePersistence
                ? 'Browser storage contains handles, not raw directory paths'
                : 'Persistent picker location is unavailable in this browser',
            data: dataHandle,
            view: viewHandle
        });
    }

    hydratePersistedHandles() {
        if (this.persistedHandleHydrationPromise) {
            return this.persistedHandleHydrationPromise;
        }

        this.persistedHandleHydrationPromise = (async () => {
            const [dataHandle, viewHandle] = await Promise.all([
                this.readPersistedHandle('data'),
                this.readPersistedHandle('view')
            ]);

            if (dataHandle && !this.fileHandles.data) {
                this.fileHandles.data = dataHandle;
            }

            if (viewHandle && !this.fileHandles.view) {
                this.fileHandles.view = viewHandle;
            }

            this.debugLog('hydrated persisted picker handles', {
                userScope: this.getCurrentUserScope(),
                data: this.describeHandle(this.fileHandles.data),
                view: this.describeHandle(this.fileHandles.view)
            });
        })().catch((error) => {
            console.warn('Failed to hydrate persisted file handles:', error);
        });

        return this.persistedHandleHydrationPromise;
    }

    async logPickerIntent(action, type, extra = {}) {
        const persisted = await this.getPersistedHandleSummary(type);
        this.debugLog(`${action}: picker intent`, {
            action,
            type,
            userScope: this.getCurrentUserScope(),
            persisted,
            ...extra
        });
    }
    
    // ===== LOADING OPERATIONS =====
    
    /**
     * Open a file picker for selecting a JSON file
     * @param {string} type - File type hint ('data' or 'view')
     * @returns {Promise<{content: any, filename: string}>}
     */
    async openFile(type = 'json') {
        this.logPickerIntent('open-file', type, {
            supportsOpenPicker: this.supportsFileOpenPicker,
            isDesktopMode: this.isDesktopMode
        }).catch((error) => {
            console.warn('Failed to log open-file picker intent:', error);
        });

        if (this.isDesktopMode) {
            return this.openFileWithDesktopBridge(type);
        } else if (this.supportsFileOpenPicker) {
            return this.openFileWithFileSystemAccess(type);
        } else {
            return this.openFileWithFileInput(type);
        }
    }

    async openFileWithDesktopBridge(type) {
        const path = await window.desktopBridge.pickFile('JSON (*.json)');
        if (!path) return null;

        const result = await window.desktopBridge.readFile(path);
        if (result.error) {
            throw new Error(`Failed to read file: ${result.error}`);
        }

        const parsedContent = JSON.parse(result.content);
        const fileType = this.resolveFileType(type, parsedContent);
        const filename = path.split(/[\\/]/).pop();

        window.desktopBridge.setCachedPath(fileType, path);
        this.rememberOpenedFile(fileType, filename);

        return { content: parsedContent, filename };
    }
    
    async openFileWithFileSystemAccess(type) {
        try {
            const startIn = this.getPickerStartLocation(type);
            this.debugLog('opening file picker', {
                type,
                userScope: this.getCurrentUserScope(),
                startIn: this.describeHandle(startIn) || startIn
            });
            const [fileHandle] = await window.showOpenFilePicker({
                id: this.getPickerId(type),
                types: [this.fileTypes.json],
                startIn,
                excludeAcceptAllOption: true,
                multiple: false
            });
            
            const file = await fileHandle.getFile();
            const parsedContent = await this.readJSONFile(file);
            const fileType = this.resolveFileType(type, parsedContent);
            
            // Store handle for future quick saves
            if (fileType === 'data') {
                await this.cacheFileHandle('data', fileHandle);
            } else if (fileType === 'view') {
                await this.cacheFileHandle('view', fileHandle);
            }

            this.rememberOpenedFile(fileType, file.name);
            
            return {
                content: parsedContent,
                filename: file.name,
                fileHandle: fileHandle
            };
            
        } catch (error) {
            if (error.name !== 'AbortError') {
                throw new Error(`Failed to open file: ${error.message}`);
            }
            return null; // User cancelled
        }
    }

    async reopenFile(type = 'data') {
        const fileHandle = await this.getCachedFileHandle(type);
        if (!fileHandle) {
            throw new Error(this.getReloadUnavailableMessage(type));
        }

        this.debugLog('reload using cached handle', {
            type,
            userScope: this.getCurrentUserScope(),
            handle: this.describeHandle(fileHandle)
        });

        const file = await fileHandle.getFile();
        const parsedContent = await this.readJSONFile(file);
        this.rememberOpenedFile(type, file.name);

        return {
            content: parsedContent,
            filename: file.name,
            fileHandle
        };
    }
    
    async openFileWithFileInput(type = 'json') {
        this.debugLog('opening fallback file input', {
            type,
            userScope: this.getCurrentUserScope()
        });

        return new Promise((resolve, reject) => {
            const input = document.createElement('input');
            input.type = 'file';
            input.accept = '.json,application/json';
            
            input.onchange = async (e) => {
                const file = e.target.files[0];
                if (!file) {
                    resolve(null);
                    return;
                }
                
                try {
                    const parsedContent = await this.readJSONFile(file);
                    const fileType = this.resolveFileType(type, parsedContent);

                    this.rememberOpenedFile(fileType, file.name);
                    
                    resolve({
                        content: parsedContent,
                        filename: file.name,
                        file: file
                    });
                } catch (error) {
                    reject(new Error(`Failed to read file: ${error.message}`));
                }
            };
            
            input.click();
        });
    }
    
    /**
     * Handle drag-and-drop file loading
     * @param {DragEvent} event - Drop event
     * @returns {Promise<Array<{content: any, filename: string}>>}
     */
    async handleFileDrop(event) {
        const files = Array.from(event.dataTransfer.files);
        const results = [];
        
        for (const file of files) {
            if (file.type === 'application/json' || file.name.endsWith('.json')) {
                try {
                    const parsedContent = await this.readJSONFile(file);
                    const fileType = this.resolveFileType('json', parsedContent);

                    this.rememberOpenedFile(fileType, file.name);
                    
                    results.push({
                        content: parsedContent,
                        filename: file.name,
                        fileType,
                        file: file
                    });
                } catch (error) {
                    console.error(`Failed to process dropped file ${file.name}:`, error);
                    // Continue with other files
                }
            }
        }
        
        return results;
    }
    
    async readJSONFile(file) {
        if (!file || typeof file.text !== 'function') {
            throw new Error('Selected file cannot be read in this browser');
        }

        const content = await file.text();
        return JSON.parse(content);
    }
    
    // ===== SAVING OPERATIONS =====
    
    /**
     * Save data to a file using File System Access API or fallback download
     * @param {any} data - Data to save
     * @param {string} filename - Suggested filename
     * @param {string} type - File type hint ('data' or 'view')
     * @returns {Promise<void>}
     */
    async saveFile(data, filename, type = 'json', options = {}) {
        const { preferProvidedFilename = false } = options;
        const suggestedFilename = preferProvidedFilename
            ? filename
            : this.getSuggestedFilename(type, filename);

        this.logPickerIntent('save-file', type, {
            supportsSavePicker: this.supportsFileSavePicker,
            isDesktopMode: this.isDesktopMode,
            suggestedFilename
        }).catch((error) => {
            console.warn('Failed to log save-file picker intent:', error);
        });

        if (this.isDesktopMode) {
            return this.saveFileWithDesktopBridge(data, suggestedFilename, type);
        } else if (this.supportsFileSavePicker) {
            return this.saveFileWithFileSystemAccess(data, suggestedFilename, type);
        } else {
            return this.saveFileWithDownload(data, suggestedFilename, type);
        }
    }

    async saveFileWithDesktopBridge(data, filename, type) {
        // Use cached path for quick re-save; otherwise prompt with native save dialog
        let path = window.desktopBridge.getCachedPath(type);
        if (!path) {
            path = await window.desktopBridge.pickSaveFile(filename, 'JSON (*.json)');
            if (!path) return null;
        }

        const jsonString = JSON.stringify(data, null, 2);
        const result = await window.desktopBridge.writeFile(path, jsonString);
        if (!result.success) {
            throw new Error(`Failed to save file: ${result.error}`);
        }

        const savedFilename = path.split(/[\\/]/).pop();
        window.desktopBridge.setCachedPath(type, path);
        this.rememberSavedFilename(type, savedFilename);
        return savedFilename;
    }
    
    async saveFileWithFileSystemAccess(data, filename, type) {
        try {
            const fileHandle = await this.requestSaveFileHandle(type, filename, 'json');

            // Remember the chosen handle so the next picker starts nearby and
            // explicit quick-save flows can still overwrite directly.
            await this.cacheFileHandle(type, fileHandle);
            
            // Write the file
            const writable = await fileHandle.createWritable();
            const jsonString = JSON.stringify(data, null, 2);
            await writable.write(jsonString);
            await writable.close();

            this.rememberSavedFilename(type, fileHandle.name || filename);
            
            return fileHandle.name || filename;
            
        } catch (error) {
            if (error.name === 'AbortError') {
                return null; // User cancelled
            }
            
            // Fallback to download if File System Access fails
            console.warn('File System Access failed, falling back to download:', error);
            return this.saveFileWithDownload(data, filename, type);
        }
    }
    
    saveFileWithDownload(data, filename, type = 'json') {
        this.debugLog('save fallback download', {
            type,
            userScope: this.getCurrentUserScope(),
            filename
        });

        const jsonString = JSON.stringify(data, null, 2);
        const blob = new Blob([jsonString], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        
        const link = document.createElement('a');
        link.href = url;
        link.download = filename;
        link.style.display = 'none';
        
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);

        this.rememberSavedFilename(type, filename);
        
        // Clean up the URL object
        setTimeout(() => URL.revokeObjectURL(url), 1000);
        
        return filename;
    }

    async requestSaveFileHandle(type, filename, fileType = 'json') {
        const pickerFileType = this.fileTypes[fileType] || this.fileTypes.json;
        const preferredStartIn = this.getPickerStartLocation(type);
        const attempts = [
            {
                label: 'preferred options',
                includeId: true,
                includeStartIn: true,
                startIn: preferredStartIn
            },
            {
                label: 'without remembered location',
                includeId: true,
                includeStartIn: false,
                startIn: null
            },
            {
                label: 'minimal options',
                includeId: false,
                includeStartIn: false,
                startIn: null
            }
        ];

        let lastError = null;

        for (const attempt of attempts) {
            try {
                const pickerOptions = {
                    types: [pickerFileType],
                    suggestedName: filename
                };

                if (attempt.includeId) {
                    pickerOptions.id = this.getPickerId(type);
                }

                if (attempt.includeStartIn && attempt.startIn) {
                    pickerOptions.startIn = attempt.startIn;
                }

                this.debugLog('opening save picker', {
                    type,
                    userScope: this.getCurrentUserScope(),
                    attempt: attempt.label,
                    startIn: this.describeHandle(attempt.startIn) || attempt.startIn || null,
                    suggestedFilename: filename
                });

                return await window.showSaveFilePicker(pickerOptions);
            } catch (error) {
                if (error.name === 'AbortError') {
                    throw error;
                }

                lastError = error;
                this.debugLog('save picker attempt failed', {
                    type,
                    userScope: this.getCurrentUserScope(),
                    attempt: attempt.label,
                    errorName: error.name || 'Error',
                    errorMessage: error.message || String(error)
                });
            }
        }

        throw lastError || new Error('Save picker failed before opening');
    }

            async saveBlob(blob, filename, options = {}) {
                const {
                    pickerType = 'data',
                    fileType = 'xlsx',
                    mimeType = blob && blob.type ? blob.type : 'application/octet-stream'
                } = options;

                if (!(blob instanceof Blob)) {
                    throw new Error('Expected Blob when saving binary file');
                }

                await this.logPickerIntent('save-blob', pickerType, {
                    supportsSavePicker: this.supportsFileSavePicker,
                    suggestedFilename: filename,
                    fileType,
                    mimeType
                }).catch((error) => {
                    console.warn('Failed to log save-blob picker intent:', error);
                });

                if (this.supportsFileSavePicker) {
                    try {
                        const fileHandle = await this.requestSaveFileHandle(pickerType, filename, fileType);

                        await this.cacheFileHandle(pickerType, fileHandle);
                        this.rememberSavedFilename(pickerType, fileHandle.name || filename);

                        const writable = await fileHandle.createWritable();
                        await writable.write(blob);
                        await writable.close();
                        return fileHandle.name || filename;
                    } catch (error) {
                        if (error.name === 'AbortError') {
                            return null;
                        }

                        console.warn('Binary File System Access save failed, falling back to download:', error);
                    }
                }

                const url = URL.createObjectURL(blob);
                const link = document.createElement('a');
                link.href = url;
                link.download = filename;
                link.style.display = 'none';
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                setTimeout(() => URL.revokeObjectURL(url), 1000);
                return filename;
            }
    
    /**
     * Quick save to previously used file handle
     * @param {any} data - Data to save
     * @param {string} type - 'data' or 'view'
     * @returns {Promise<boolean>} - Success status
     */
    async quickSave(data, type) {
        // Desktop mode: quick-save to cached absolute path
        if (this.isDesktopMode) {
            const cachedPath = window.desktopBridge.getCachedPath(type);
            if (!cachedPath) return false;
            const jsonString = JSON.stringify(data, null, 2);
            const result = await window.desktopBridge.writeFile(cachedPath, jsonString);
            return result.success === true;
        }

        if (!this.supportsFileSystemAccess) {
            return false; // Quick save not supported without File System Access
        }
        
        const fileHandle = this.fileHandles[type];
        if (!fileHandle) {
            return false; // No existing file handle
        }
        
        try {
            const writable = await fileHandle.createWritable();
            const jsonString = JSON.stringify(data, null, 2);
            await writable.write(jsonString);
            await writable.close();
            return true;
        } catch (error) {
            console.error('Quick save failed:', error);
            // Clear the invalid handle
            this.fileHandles[type] = null;
            return false;
        }
    }
    
    // ===== FILE TYPE DETECTION =====
    
    /**
     * Determine if a JSON object represents data
     * @param {any} jsonData - Parsed JSON content
     * @returns {boolean}
     */
    isDataFile(jsonData) {
        return jsonData && (
            (jsonData.hasOwnProperty('data') && Array.isArray(jsonData.data)) ||
            (Array.isArray(jsonData) && jsonData.length > 0 && typeof jsonData[0] === 'object')
        );
    }
    
    /**
     * Determine if a JSON object represents a view configuration
     * @param {any} jsonData - Parsed JSON content  
     * @returns {boolean}
     */
    isViewConfigFile(jsonData) {
        return jsonData && (
            jsonData.hasOwnProperty('axisSelections') ||
            jsonData.hasOwnProperty('filters') ||
            jsonData.hasOwnProperty('tagCustomizations') ||
            jsonData.hasOwnProperty('gridState') ||
            jsonData.hasOwnProperty('cardClick') ||
            jsonData.hasOwnProperty('urlConfig')
        );
    }
    
    // ===== FILENAME GENERATION =====
    
    /**
     * Generate a filename for data export
     * @param {number} itemCount - Number of items in dataset
     * @returns {string}
     */
    generateDataFilename(itemCount = 0) {
        const timestamp = window.GridDateUtils
            ? window.GridDateUtils.createFilenameTimestamp()
            : new Date().toISOString().slice(0, 19).replace(/[T:]/g, '-');
        return `data-export-${itemCount}items-${timestamp}.json`;
    }
    
    /**
     * Generate a filename for view configuration export
     * @param {string} xAxis - X-axis field name
     * @param {string} yAxis - Y-axis field name
     * @returns {string}
     */
    generateViewFilename(xAxis = 'none', yAxis = 'none') {
        const timestamp = window.GridDateUtils
            ? window.GridDateUtils.createFilenameTimestamp()
            : new Date().toISOString().slice(0, 19).replace(/[T:]/g, '-');
        return `view-config-${xAxis}x${yAxis}-${timestamp}.json`;
    }

    /**
     * Derive a companion view filename from a data filename.
     * "foo.json" → "foo.view.json"
     * @param {string} dataFilename
     * @returns {string|null}
     */
    deriveCompanionViewFilename(dataFilename) {
        if (!dataFilename || typeof dataFilename !== 'string') {
            return null;
        }
        const match = dataFilename.match(/^(.+)\.json$/i);
        if (match) {
            return match[1] + '.view.json';
        }
        return dataFilename + '.view.json';
    }

    /**
     * Get the filename of the last opened file of the given type.
     * @param {string} type - 'data' or 'view'
     * @returns {string|null}
     */
    getLastOpenedFilename(type) {
        return this.lastOpenedFiles[type] ? this.lastOpenedFiles[type].filename : null;
    }
    
    // ===== UTILITY METHODS =====
    
    /**
     * Check if File System Access API is supported
     * @returns {boolean}
     */
    isFileSystemAccessSupported() {
        return this.supportsFileSystemAccess;
    }

    supportsSavePicker() {
        return this.supportsFileSavePicker;
    }

    resolveFileType(type, parsedContent) {
        if (type === 'data' || type === 'view') {
            return type;
        }

        if (this.isDataFile(parsedContent)) {
            return 'data';
        }

        if (this.isViewConfigFile(parsedContent)) {
            return 'view';
        }

        return 'json';
    }

    rememberOpenedFile(type, filename) {
        if (type !== 'data' && type !== 'view') {
            return;
        }

        this.lastOpenedFiles[type] = {
            filename: filename || null,
            openedAt: window.GridDateUtils.createLocalTimestamp()
        };
    }

    rememberSavedFilename(type, filename) {
        if (type !== 'data' && type !== 'view') {
            return;
        }

        this.lastSavedFilenames[type] = filename || null;
    }

    getSuggestedFilename(type, fallbackFilename) {
        if (type !== 'data' && type !== 'view') {
            return fallbackFilename;
        }

        const openedFilename = this.lastOpenedFiles[type] && this.lastOpenedFiles[type].filename;
        const savedFilename = this.lastSavedFilenames[type];
        return openedFilename || savedFilename || fallbackFilename;
    }

    getReloadAction(type = 'data') {
        if (this.fileHandles[type]) {
            return 'reopen';
        }

        if (this.lastOpenedFiles[type]) {
            return 'reselect';
        }

        return 'unavailable';
    }

    canReload(type = 'data') {
        return this.getReloadAction(type) !== 'unavailable';
    }

    getReloadUnavailableMessage(type = 'data') {
        if (this.getReloadAction(type) === 'reselect') {
            return 'This browser cannot reopen the original file directly. Choose the file again to reload it.';
        }

        return 'No previously opened file is available to reload yet';
    }

    getPickerId(type) {
        return this.pickerIds[type] || 'data-visualization-grid-json';
    }

    getCurrentUserScope() {
        const rawUserName = this.getCurrentUserName();
        const normalizedUserName = String(rawUserName || '').trim();
        return normalizedUserName || 'Local User';
    }

    getPersistedHandleKey(type) {
        return `${this.getCurrentUserScope()}::${type}`;
    }

    openHandleDatabase() {
        if (!this.supportsHandlePersistence) {
            return Promise.resolve(null);
        }

        if (this.handleDbPromise) {
            return this.handleDbPromise;
        }

        this.handleDbPromise = new Promise((resolve) => {
            const request = indexedDB.open(this.handleDbName, 1);

            request.onupgradeneeded = () => {
                const database = request.result;
                if (!database.objectStoreNames.contains(this.handleStoreName)) {
                    database.createObjectStore(this.handleStoreName, { keyPath: 'key' });
                }
            };

            request.onsuccess = () => resolve(request.result);
            request.onerror = () => {
                console.warn('Failed to open file handle database:', request.error);
                resolve(null);
            };
        });

        return this.handleDbPromise;
    }

    async readPersistedHandle(type) {
        if (!this.supportsHandlePersistence || (type !== 'data' && type !== 'view')) {
            return null;
        }

        const database = await this.openHandleDatabase();
        if (!database) {
            return null;
        }

        return new Promise((resolve) => {
            const transaction = database.transaction(this.handleStoreName, 'readonly');
            const store = transaction.objectStore(this.handleStoreName);
            const request = store.get(this.getPersistedHandleKey(type));

            request.onsuccess = () => {
                const result = request.result;
                resolve(result && result.handle ? result.handle : null);
            };

            request.onerror = () => {
                console.warn('Failed to read persisted file handle:', request.error);
                resolve(null);
            };
        });
    }

    async persistHandle(type, handle) {
        if (!this.supportsHandlePersistence || !handle || (type !== 'data' && type !== 'view')) {
            return;
        }

        const database = await this.openHandleDatabase();
        if (!database) {
            return;
        }

        await new Promise((resolve) => {
            try {
                const transaction = database.transaction(this.handleStoreName, 'readwrite');
                const store = transaction.objectStore(this.handleStoreName);
                store.put({
                    key: this.getPersistedHandleKey(type),
                    handle,
                    savedAt: window.GridDateUtils.createLocalTimestamp()
                });
                transaction.oncomplete = () => {
                    this.debugLog('persisted picker location', {
                        type,
                        userScope: this.getCurrentUserScope(),
                        storage: {
                            database: this.handleDbName,
                            store: this.handleStoreName,
                            key: this.getPersistedHandleKey(type)
                        },
                        handle: this.describeHandle(handle),
                        note: 'Browser storage contains a file handle for picker start location, not a raw directory path'
                    });
                    resolve();
                };
                transaction.onerror = () => {
                    console.warn('Failed to persist file handle:', transaction.error);
                    resolve();
                };
            } catch (error) {
                console.warn('Failed to persist file handle:', error);
                resolve();
            }
        });
    }

    async getCachedFileHandle(type) {
        if (this.fileHandles[type]) {
            return this.fileHandles[type];
        }

        const persistedHandle = await this.readPersistedHandle(type);
        if (persistedHandle) {
            this.fileHandles[type] = persistedHandle;
        }

        return this.fileHandles[type] || null;
    }

    async cacheFileHandle(type, handle) {
        if (type !== 'data' && type !== 'view') {
            return;
        }

        this.fileHandles[type] = handle;
        await this.persistHandle(type, handle);
    }

    getPickerStartLocation(type) {
        if (type !== 'data' && type !== 'view') {
            return 'documents';
        }

        const handle = this.fileHandles[type] || (type === 'view' ? this.fileHandles.data : null) || null;
        const startIn = handle || 'documents';
        this.debugLog('resolved picker start location', {
            type,
            userScope: this.getCurrentUserScope(),
            startIn: this.describeHandle(startIn) || startIn
        });
        return startIn;
    }
    
    /**
     * Clear stored file handles
     * @param {string} type - 'data', 'view', or 'all'
     */
    clearFileHandles(type = 'all') {
        if (type === 'all') {
            this.fileHandles.data = null;
            this.fileHandles.view = null;
        } else {
            this.fileHandles[type] = null;
        }
    }
    
    /**
     * Get information about stored file handles
     * @returns {object} Handle status
     */
    getFileHandleStatus() {
        return {
            data: !!this.fileHandles.data,
            view: !!this.fileHandles.view,
            supportsFileSystemAccess: this.supportsFileSystemAccess,
            supportsOpenPicker: this.supportsFileOpenPicker,
            supportsSavePicker: this.supportsFileSavePicker,
            filenames: {
                data: this.getSuggestedFilename('data', null),
                view: this.getSuggestedFilename('view', null)
            },
            reloadDataAction: this.getReloadAction('data'),
            canReloadData: this.canReload('data')
        };
    }
}

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
    module.exports = FileService;
}