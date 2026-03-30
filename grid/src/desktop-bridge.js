// Copyright (c) 2026 Malte Doerper. MIT License. See LICENSE file.

// Desktop Bridge — QWebChannel adapter for the PySide6 desktop shell.
// Detects the QWebChannel transport injected by main.py and exposes
// window.desktopBridge with async wrappers over the Python Bridge object.
//
// Detection:  window.qt && window.qt.webChannelTransport  →  desktop mode
// In browser: window.desktopBridge stays undefined, all code paths unchanged.

(function () {
    'use strict';

    // Only activate when running inside the PySide6 QWebEngineView.
    if (typeof window.qt === 'undefined' || !window.qt.webChannelTransport) {
        return;
    }

    // QWebChannel is injected as a user-script at DocumentCreation by main.py.
    // It defines the global QWebChannel constructor.
    if (typeof QWebChannel === 'undefined') {
        console.error('[DESKTOP] QWebChannel constructor not found — qwebchannel.js may not have been injected');
        return;
    }

    // Track the current save path per type for quick-save support
    const _cachedPaths = { data: null, view: null };

    // Initialise the channel — the callback fires once the bridge is ready.
    new QWebChannel(window.qt.webChannelTransport, function (channel) {
        const bridge = channel.objects.bridge;
        if (!bridge) {
            console.error('[DESKTOP] Bridge object not found on QWebChannel');
            return;
        }

        /**
         * Async wrappers.
         *
         * PySide6 QWebChannel Slots are invoked with a callback as the last
         * argument.  Each wrapper turns that into a Promise.
         */

        function callSlot(method, ...args) {
            return new Promise((resolve) => {
                method.call(bridge, ...args, resolve);
            });
        }

        window.desktopBridge = {
            /**
             * Open a native file-open dialog.
             * @param {string} [filter] - Qt filter string, e.g. "JSON (*.json)"
             * @returns {Promise<string>} Absolute path, or "" if cancelled.
             */
            pickFile(filter) {
                return callSlot(bridge.pick_file, filter || '');
            },

            /**
             * Open a native file-save dialog.
             * @param {string} [suggestedName] - Pre-filled filename.
             * @param {string} [filter] - Qt filter string, e.g. "JSON (*.json)"
             * @returns {Promise<string>} Absolute path, or "" if cancelled.
             */
            pickSaveFile(suggestedName, filter) {
                return callSlot(bridge.pick_save_file, suggestedName || '', filter || '');
            },

            /**
             * Open a native folder dialog.
             * @returns {Promise<string>} Absolute path, or "" if cancelled.
             */
            pickFolder() {
                return callSlot(bridge.pick_folder);
            },

            /**
             * Read a file's text content.
             * @param {string} path - Absolute file path.
             * @returns {Promise<{content: string, error: string}>}
             */
            async readFile(path) {
                const raw = await callSlot(bridge.read_file, path);
                return JSON.parse(raw);
            },

            /**
             * Write text content to a file.
             * @param {string} path - Absolute file path.
             * @param {string} content - Text content to write.
             * @returns {Promise<{success: boolean, error: string}>}
             */
            async writeFile(path, content) {
                const raw = await callSlot(bridge.write_file, path, content);
                return JSON.parse(raw);
            },

            // --- Path cache for quick-save ---

            getCachedPath(type) {
                return _cachedPaths[type] || null;
            },

            setCachedPath(type, path) {
                _cachedPaths[type] = path || null;
            },
        };

        // Expose the raw bridge for desktop-tools.js to call additional slots
        window._desktopBridgeRaw = bridge;

        console.log('[DESKTOP] Desktop bridge initialised');

        // Dispatch a custom event so other modules can detect desktop mode
        // even if they loaded before the channel was ready.
        window.dispatchEvent(new Event('desktop-bridge-ready'));
    });
})();


// ===== DESKTOP CONTROLS =====
// Sets up desktop-mode UI buttons for CLI-backed actions
// (Convert Excel, Save to Excel) once the bridge and app are ready.

(function () {
    'use strict';

    function initDesktopControls() {
        if (typeof app === 'undefined' || !app) return;
        if (!window.desktopBridge) return;

        app.isDesktopMode = true;

        // Hide server controls, show desktop controls (file-info only).
        // .file-controls, .data-view-controls, and section header are hidden
        // by desktop-tools.js initToolsPanel() which runs reliably after delay.
        const serverControls = document.querySelector('.server-controls');
        const desktopControls = document.querySelector('.desktop-controls');

        if (serverControls) serverControls.style.display = 'none';
        if (desktopControls) desktopControls.style.display = '';

        buildDesktopUI(desktopControls);
        console.log('[DESKTOP] Desktop controls active');
    }

    function buildDesktopUI(container) {
        if (!container) return;

        // Desktop-controls now only holds the file info indicator
        container.innerHTML = '<div id="desktop-file-info" class="desktop-file-info" title=""></div>';
    }

    // Expose handlers so desktop-tools.js can wire them into the Json tab
    window._desktopHandlers = {
        load: handleLoadData,
        reload: handleReloadData,
        save: handleSaveData,
    };

    // --- Console popup helpers (reuse existing popup from index.html) ---

    function openConsole(title) {
        const overlay = document.getElementById('console-popup-overlay');
        const titleEl = document.getElementById('console-popup-title');
        const output = document.getElementById('console-popup-output');
        const status = document.getElementById('console-popup-status');
        const inputArea = document.getElementById('console-popup-input');
        if (titleEl) titleEl.textContent = title || 'Desktop Console';
        if (output) output.textContent = '';
        if (status) status.textContent = 'Running…';
        if (inputArea) inputArea.style.display = 'none';
        if (overlay) overlay.style.display = '';
    }

    function appendConsole(text) {
        const output = document.getElementById('console-popup-output');
        if (output) {
            output.textContent += text;
            output.scrollTop = output.scrollHeight;
        }
    }

    function setConsoleStatus(text) {
        const el = document.getElementById('console-popup-status');
        if (el) el.textContent = text;
    }

    // --- Track last-used paths for quick operations ---

    let _lastJsonPath = null;

    function updateFileInfo(fullPath) {
        const el = document.getElementById('desktop-file-info');
        if (!el) return;
        if (!fullPath) {
            el.textContent = '';
            el.title = '';
            document.title = '';
            return;
        }
        const sep = fullPath.includes('/') ? '/' : '\\';
        const parts = fullPath.split(sep);
        const filename = parts.pop();
        const dir = parts.join(sep);
        el.textContent = filename;
        el.title = fullPath;
        document.title = filename;

        const gridPath = document.getElementById('grid-file-path');
        if (gridPath) {
            gridPath.textContent = fullPath;
            gridPath.title = fullPath;
        }
    }

    function updateActionState() {
        const reloadBtn = document.getElementById('desktop-reload-data-btn');
        if (reloadBtn) reloadBtn.disabled = !_lastJsonPath;
    }

    // --- Action handlers ---

    async function handleLoadData() {
        try {
            const path = await window.desktopBridge.pickFile('JSON (*.json)');
            if (!path) return;

            const result = await window.desktopBridge.readFile(path);
            if (result.error) {
                alert('Failed to read file: ' + result.error);
                return;
            }

            const parsed = JSON.parse(result.content);
            _lastJsonPath = path;
            window.desktopBridge.setCachedPath('data', path);
            app.loadDataFromJSON(parsed);
            updateFileInfo(path);
            updateActionState();
        } catch (err) {
            alert('Load failed: ' + err.message);
        }
    }

    async function handleReloadData() {
        if (!_lastJsonPath) return;
        try {
            const result = await window.desktopBridge.readFile(_lastJsonPath);
            if (result.error) {
                alert('Failed to reload: ' + result.error);
                return;
            }
            const parsed = JSON.parse(result.content);
            app.loadDataFromJSON(parsed);
        } catch (err) {
            alert('Reload failed: ' + err.message);
        }
    }

    async function handleSaveData() {
        try {
            if (!app.persistenceManager) return;
            // Use the persistence manager's save flow which respects desktop mode
            await app.persistenceManager.saveDataFile();
        } catch (err) {
            alert('Save failed: ' + err.message);
        }
    }

    // Expose updateFileInfo so desktop-tools.js can update the display
    window._desktopUpdateFileInfo = function (fullPath) {
        _lastJsonPath = fullPath;
        window.desktopBridge.setCachedPath('data', fullPath);
        updateFileInfo(fullPath);
        updateActionState();
    };

    // --- Initialisation ---
    // Wait for both DOMContentLoaded (app exists) and the bridge. The bridge
    // may be ready before or after DOMContentLoaded, so handle both orderings.

    let _domReady = false;

    function tryInit() {
        if (_domReady && window.desktopBridge) {
            initDesktopControls();
        }
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', () => {
            _domReady = true;
            tryInit();
        });
    } else {
        _domReady = true;
    }

    window.addEventListener('desktop-bridge-ready', tryInit);
    // If the bridge was already initialised synchronously, try immediately
    tryInit();
})();
