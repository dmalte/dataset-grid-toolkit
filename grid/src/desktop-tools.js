// Copyright (c) 2026 Malte Doerper. MIT License. See LICENSE file.

// Desktop Tools Panel — renders a YAML-driven tools UI in desktop mode.
// Fetches tool definitions from the bridge (tools.yaml), builds a tabbed
// panel with controls and action buttons, and executes tools via the bridge.

(function () {
    'use strict';

    // Only activate in desktop mode
    if (typeof window.qt === 'undefined' || !window.qt.webChannelTransport) {
        return;
    }

    const STORAGE_PREFIX = 'grid-desktop-tool:';

    // State
    let _toolsConfig = null;   // {settings, tools}
    let _activeTabId = null;
    let _sourceInfo = null;     // tracked-source metadata

    // ------------------------------------------------------------------
    // Console popup helpers (reuse existing popup from index.html)
    // ------------------------------------------------------------------

    function openConsole(title, interactive) {
        const overlay = document.getElementById('console-popup-overlay');
        const titleEl = document.getElementById('console-popup-title');
        const output = document.getElementById('console-popup-output');
        const status = document.getElementById('console-popup-status');
        const inputArea = document.getElementById('console-popup-input');
        if (titleEl) titleEl.textContent = title || 'Desktop Console';
        if (output) output.textContent = '';
        if (status) status.textContent = 'Running…';
        if (inputArea) inputArea.style.display = interactive ? '' : 'none';
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

    function setConsoleInputVisible(visible) {
        const inputArea = document.getElementById('console-popup-input');
        if (inputArea) inputArea.style.display = visible ? '' : 'none';
    }

    // ------------------------------------------------------------------
    // localStorage helpers
    // ------------------------------------------------------------------

    function getStored(toolId, controlId) {
        try {
            return localStorage.getItem(STORAGE_PREFIX + toolId + ':' + controlId) || '';
        } catch (_e) { return ''; }
    }

    function setStored(toolId, controlId, value) {
        try {
            const key = STORAGE_PREFIX + toolId + ':' + controlId;
            if (value) {
                localStorage.setItem(key, value);
            } else {
                localStorage.removeItem(key);
            }
        } catch (_e) { /* ignore */ }
    }

    // ------------------------------------------------------------------
    // Settings helpers (backed by QSettings via bridge)
    // ------------------------------------------------------------------

    async function loadSetting(key) {
        if (!window.desktopBridge) return '';
        return new Promise((resolve) => {
            const bridge = window._desktopBridgeRaw;
            if (bridge && bridge.get_setting) {
                bridge.get_setting(key, resolve);
            } else {
                resolve('');
            }
        });
    }

    async function saveSetting(key, value) {
        if (!window.desktopBridge) return;
        return new Promise((resolve) => {
            const bridge = window._desktopBridgeRaw;
            if (bridge && bridge.set_setting) {
                bridge.set_setting(key, value);
            }
            resolve();
        });
    }

    // ------------------------------------------------------------------
    // Source info (tracked-source metadata from .pipeline-meta.json)
    // ------------------------------------------------------------------

    async function loadSourceInfo() {
        const dataPath = window.desktopBridge.getCachedPath('data');
        if (!dataPath) {
            _sourceInfo = null;
            return;
        }
        // Look for .pipeline-meta.json next to the data file
        const dir = dataPath.replace(/[/\\][^/\\]+$/, '');
        const metaPath = dir + '/.pipeline-meta.json';
        try {
            const result = await window.desktopBridge.readFile(metaPath);
            if (!result.error && result.content) {
                _sourceInfo = JSON.parse(result.content);
                return;
            }
        } catch (_e) { /* no metadata */ }
        _sourceInfo = null;
    }

    function _syncSourceContextToMeta() {
        if (!_sourceInfo || !_sourceInfo.kind) return;
        if (!app.metaInfo || typeof app.metaInfo !== 'object') {
            app.metaInfo = {};
        }
        const kind = _sourceInfo.kind;
        if (kind === 'obsidian-cli-converted') {
            app.metaInfo.sourceContext = {
                kind,
                datasetFilename: app.metaInfo.sourceContext?.datasetFilename || null,
                obsidianConfigPath: _sourceInfo.source_path || _sourceInfo.config_path || null,
            };
        } else if (kind === 'excel-cli-converted') {
            app.metaInfo.sourceContext = {
                kind,
                datasetFilename: app.metaInfo.sourceContext?.datasetFilename || null,
                workbookPath: _sourceInfo.source_excel || _sourceInfo.excel || null,
            };
        }
    }

    async function saveSourceInfo(dataPath, info) {
        if (!dataPath) return;
        const dir = dataPath.replace(/[/\\][^/\\]+$/, '');
        const metaPath = dir + '/.pipeline-meta.json';
        await window.desktopBridge.writeFile(metaPath, JSON.stringify(info, null, 2));
        _sourceInfo = info;
    }

    // ------------------------------------------------------------------
    // Rendering
    // ------------------------------------------------------------------

    function renderToolsPanel(container) {
        if (!container || !_toolsConfig) return;

        container.replaceChildren();
        const tools = _toolsConfig.tools || [];
        const settingsSchema = _toolsConfig.settings || {};

        // Build tabs — start with Data Json and View Json tabs for file operations
        const tabMap = new Map();
        tabMap.set('data-json', { id: 'data-json', label: 'Data Json', tools: [] });
        tabMap.set('view-json', { id: 'view-json', label: 'View Json', tools: [] });

        for (const tool of tools) {
            const tabLabel = tool.ui.tab;
            const tabId = tabLabel.toLowerCase().replace(/\s+/g, '-');
            if (!tabMap.has(tabId)) {
                tabMap.set(tabId, { id: tabId, label: tabLabel, tools: [] });
            }
            tabMap.get(tabId).tools.push(tool);
        }

        const tabs = Array.from(tabMap.values());
        if (!_activeTabId || !tabMap.has(_activeTabId)) {
            _activeTabId = tabs[0].id;
        }

        const fragment = document.createDocumentFragment();

        // Tabs bar
        const tabsRow = document.createElement('div');
        tabsRow.className = 'desktop-tools-tabs';
        for (const tab of tabs) {
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'desktop-tools-tab-btn';
            btn.dataset.toolTabId = tab.id;
            btn.textContent = tab.label;
            if (tab.id === _activeTabId) btn.classList.add('is-active');
            tabsRow.appendChild(btn);
        }
        fragment.appendChild(tabsRow);

        // Tab panels
        const panelsContainer = document.createElement('div');
        panelsContainer.className = 'desktop-tools-panels';

        for (const tab of tabs) {
            const panel = document.createElement('section');
            panel.className = 'desktop-tools-panel';
            panel.dataset.toolTabPanelId = tab.id;
            panel.hidden = tab.id !== _activeTabId;

            // Data Json tab: render the data file operation buttons
            if (tab.id === 'data-json') {
                panel.appendChild(buildDataJsonTabContent());
                panelsContainer.appendChild(panel);
                continue;
            }

            // View Json tab: render the view file operation buttons
            if (tab.id === 'view-json') {
                panel.appendChild(buildViewJsonTabContent());
                panelsContainer.appendChild(panel);
                continue;
            }

            // Per-tab settings (rendered inside the tab, not globally)
            const tabLabel = tab.label;
            let settingsRow = null;
            for (const [key, spec] of Object.entries(settingsSchema)) {
                if (spec.tab === tabLabel) {
                    settingsRow = document.createElement('div');
                    settingsRow.className = 'desktop-tools-settings';
                    settingsRow.appendChild(createSettingField(key, spec));
                    panel.appendChild(settingsRow);
                }
            }

            // Group tools by group label
            const groupMap = new Map();
            for (const tool of tab.tools) {
                const groupLabel = tool.ui.group;
                if (!groupMap.has(groupLabel)) {
                    groupMap.set(groupLabel, []);
                }
                groupMap.get(groupLabel).push(tool);
            }

            // Separate groups into regular (have visible controls) and button-only
            const regularGroups = [];
            const buttonOnlyTools = [];
            for (const [groupLabel, groupTools] of groupMap) {
                const allButtonOnly = groupTools.every(isButtonOnlyTool);
                if (allButtonOnly) {
                    buttonOnlyTools.push(...groupTools);
                } else {
                    regularGroups.push({ label: groupLabel, tools: groupTools });
                }
            }

            let lastRow = null;
            for (const group of regularGroups) {
                const groupEl = document.createElement('div');
                groupEl.className = 'desktop-tools-group';

                const header = document.createElement('div');
                header.className = 'desktop-tools-group-label';
                header.textContent = group.label;
                groupEl.appendChild(header);

                for (const tool of group.tools) {
                    const row = createToolRow(tool);
                    groupEl.appendChild(row);
                    lastRow = row;
                }
                panel.appendChild(groupEl);
            }

            // Append button-only tools: into last regular row, or into the
            // settings row (e.g. Obsidian tab), or as a standalone row.
            if (buttonOnlyTools.length > 0 && !lastRow) {
                // Target: settings input-row if present, otherwise create a row
                const targetRow = settingsRow
                    ? settingsRow.querySelector('.desktop-tools-setting-input-row')
                    : null;
                const container = targetRow || document.createElement('div');
                if (!targetRow) container.className = 'desktop-tools-row';
                for (const tool of buttonOnlyTools) {
                    const actionCtrl = tool.controls.find((c) => c.type === 'action_button');
                    if (actionCtrl) {
                        const btn = document.createElement('button');
                        btn.type = 'button';
                        btn.className = 'control-btn desktop-tools-action-btn';
                        btn.dataset.toolAction = tool.id;
                        btn.textContent = actionCtrl.label || tool.label;
                        btn.title = tool.description || tool.label;
                        container.appendChild(btn);
                    }
                }
                if (!targetRow) panel.appendChild(container);
            } else {
                for (const tool of buttonOnlyTools) {
                    const actionCtrl = tool.controls.find((c) => c.type === 'action_button');
                    if (actionCtrl && lastRow) {
                        const btn = document.createElement('button');
                        btn.type = 'button';
                        btn.className = 'control-btn desktop-tools-action-btn';
                        btn.dataset.toolAction = tool.id;
                        btn.textContent = actionCtrl.label || tool.label;
                        btn.title = tool.description || tool.label;
                        lastRow.appendChild(btn);
                    }
                }
            }

            panelsContainer.appendChild(panel);
        }

        fragment.appendChild(panelsContainer);
        container.appendChild(fragment);

        // Attach event listeners
        attachEventListeners(container);
        updateAllActionState();
    }

    function isButtonOnlyTool(tool) {
        // A tool with only hidden controls + action button (no visible user inputs)
        return tool.controls.every((c) => c.type === 'hidden' || c.type === 'action_button');
    }

    function buildDataJsonTabContent() {
        const wrapper = document.createElement('div');
        wrapper.className = 'desktop-tools-json-tab';

        const handlers = window._desktopHandlers || {};

        // Data file buttons
        const desktopBtns = [
            { id: 'desktop-load-data-btn', label: '📂 Load Data', handler: handlers.load },
            { id: 'desktop-reload-data-btn', label: '🔄 Reload', handler: handlers.reload, disabled: true },
            { id: 'desktop-save-data-btn', label: '💾 Save Data', handler: handlers.save },
        ];
        for (const b of desktopBtns) {
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.id = b.id;
            btn.className = 'control-btn';
            btn.textContent = b.label;
            if (b.disabled) btn.disabled = true;
            if (b.handler) btn.addEventListener('click', b.handler);
            wrapper.appendChild(btn);
        }

        // Move existing Promote / Export from the hidden .file-controls.
        const dataIds = ['promote-data-btn', 'export-filtered-btn'];
        for (const id of dataIds) {
            const el = document.getElementById(id);
            if (el) wrapper.appendChild(el);
        }

        return wrapper;
    }

    function buildViewJsonTabContent() {
        const wrapper = document.createElement('div');
        wrapper.className = 'desktop-tools-json-tab';

        // Move existing Load View / Save View from the hidden .data-view-controls.
        const viewIds = ['load-view-btn', 'save-view-btn'];
        for (const id of viewIds) {
            const el = document.getElementById(id);
            if (el) wrapper.appendChild(el);
        }

        return wrapper;
    }

    function createSettingField(key, spec) {
        const wrapper = document.createElement('div');
        wrapper.className = 'desktop-tools-setting';

        const label = document.createElement('label');
        label.className = 'desktop-tools-setting-label';
        label.textContent = spec.label || key;

        const inputRow = document.createElement('div');
        inputRow.className = 'desktop-tools-setting-input-row';

        const input = document.createElement('input');
        input.type = 'text';
        input.className = 'desktop-tools-input';
        input.placeholder = spec.placeholder || '';
        input.dataset.settingKey = key;
        input.readOnly = spec.type === 'path';

        inputRow.appendChild(input);

        if (spec.type === 'path') {
            const browseBtn = document.createElement('button');
            browseBtn.type = 'button';
            browseBtn.className = 'desktop-tools-browse-btn';
            browseBtn.textContent = '📂';
            browseBtn.title = 'Browse…';
            browseBtn.dataset.settingBrowse = key;
            inputRow.appendChild(browseBtn);
        }

        wrapper.appendChild(label);
        wrapper.appendChild(inputRow);

        // Load persisted value
        void loadSetting(key).then((val) => {
            if (val) input.value = val;
        });

        return wrapper;
    }

    function createToolRow(tool) {
        const row = document.createElement('div');
        row.className = 'desktop-tools-row';
        row.dataset.toolId = tool.id;

        // Render visible controls
        for (const ctrl of tool.controls) {
            if (ctrl.type === 'hidden' || ctrl.type === 'action_button') continue;

            const fieldEl = createControlField(tool.id, ctrl);
            if (fieldEl) row.appendChild(fieldEl);
        }

        // Action button
        const actionCtrl = tool.controls.find((c) => c.type === 'action_button');
        if (actionCtrl) {
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'control-btn desktop-tools-action-btn';
            btn.dataset.toolAction = tool.id;
            btn.textContent = actionCtrl.label || tool.label;
            btn.title = tool.description || tool.label;
            row.appendChild(btn);
        }

        return row;
    }

    function createControlField(toolId, ctrl) {
        const wrapper = document.createElement('div');
        wrapper.className = 'desktop-tools-field';

        const label = document.createElement('label');
        label.className = 'desktop-tools-field-label';
        label.textContent = ctrl.label || ctrl.id;

        wrapper.appendChild(label);

        if (ctrl.type === 'file_picker') {
            const inputRow = document.createElement('div');
            inputRow.className = 'desktop-tools-field-input-row';

            const input = document.createElement('input');
            input.type = 'text';
            input.className = 'desktop-tools-input';
            input.readOnly = true;
            input.placeholder = ctrl.placeholder || '';
            input.dataset.toolControl = ctrl.id;
            input.dataset.toolId = toolId;

            const browseBtn = document.createElement('button');
            browseBtn.type = 'button';
            browseBtn.className = 'desktop-tools-browse-btn';
            browseBtn.textContent = '📂';
            browseBtn.title = 'Browse…';
            browseBtn.dataset.toolBrowse = ctrl.id;
            browseBtn.dataset.toolId = toolId;
            browseBtn.dataset.toolFilters = ctrl.filters || '';

            inputRow.appendChild(input);
            inputRow.appendChild(browseBtn);
            wrapper.appendChild(inputRow);

            // Restore persisted value
            if (ctrl.persist) {
                const stored = getStored(toolId, ctrl.id);
                if (stored) input.value = stored;
            }
        } else if (ctrl.type === 'text_input') {
            const input = document.createElement('input');
            input.type = 'text';
            input.className = 'desktop-tools-input';
            input.placeholder = ctrl.placeholder || '';
            input.value = ctrl.default || '';
            input.dataset.toolControl = ctrl.id;
            input.dataset.toolId = toolId;
            wrapper.appendChild(input);
        } else if (ctrl.type === 'select') {
            const select = document.createElement('select');
            select.className = 'desktop-tools-select';
            select.dataset.toolControl = ctrl.id;
            select.dataset.toolId = toolId;
            for (const choice of (ctrl.choices || [])) {
                const opt = document.createElement('option');
                opt.value = choice;
                opt.textContent = choice;
                if (choice === ctrl.default) opt.selected = true;
                select.appendChild(opt);
            }
            wrapper.appendChild(select);
        }

        return wrapper;
    }

    // ------------------------------------------------------------------
    // Event listeners
    // ------------------------------------------------------------------

    function attachEventListeners(container) {
        container.addEventListener('click', async (event) => {
            // Tab switching
            const tabBtn = event.target.closest('[data-tool-tab-id]');
            if (tabBtn) {
                setActiveTab(tabBtn.dataset.toolTabId, container);
                return;
            }

            // Browse button for file picker controls
            const browseBtn = event.target.closest('[data-tool-browse]');
            if (browseBtn) {
                await handleBrowse(browseBtn);
                return;
            }

            // Browse button for settings
            const settingBrowse = event.target.closest('[data-setting-browse]');
            if (settingBrowse) {
                await handleSettingBrowse(settingBrowse);
                return;
            }

            // Action button
            const actionBtn = event.target.closest('[data-tool-action]');
            if (actionBtn) {
                await executeTool(actionBtn.dataset.toolAction);
                return;
            }
        });

        // Persist text inputs on blur
        container.addEventListener('blur', (event) => {
            const input = event.target.closest('[data-tool-control]');
            if (input && input.dataset.toolId) {
                setStored(input.dataset.toolId, input.dataset.toolControl, input.value);
            }
        }, true);
    }

    function setActiveTab(tabId, container) {
        _activeTabId = tabId;

        // Update tab buttons
        const buttons = container.querySelectorAll('.desktop-tools-tab-btn');
        for (const btn of buttons) {
            btn.classList.toggle('is-active', btn.dataset.toolTabId === tabId);
        }

        // Update panels
        const panels = container.querySelectorAll('.desktop-tools-panel');
        for (const panel of panels) {
            panel.hidden = panel.dataset.toolTabPanelId !== tabId;
        }
    }

    async function handleBrowse(browseBtn) {
        const controlId = browseBtn.dataset.toolBrowse;
        const toolId = browseBtn.dataset.toolId;
        const filters = browseBtn.dataset.toolFilters || '';

        // Build Qt filter string from comma-separated extensions
        let qtFilter = 'All Files (*)';
        if (filters) {
            const exts = filters.split(',').map((f) => f.trim());
            const pattern = exts.map((e) => '*' + e).join(' ');
            qtFilter = 'Files (' + pattern + ');;All Files (*)';
        }

        const path = await window.desktopBridge.pickFile(qtFilter);
        if (!path) return;

        const input = document.querySelector(
            `[data-tool-control="${controlId}"][data-tool-id="${toolId}"]`
        );
        if (input) {
            input.value = path;
            setStored(toolId, controlId, path);
        }

        updateAllActionState();
    }

    async function handleSettingBrowse(browseBtn) {
        const settingKey = browseBtn.dataset.settingBrowse;
        const path = await window.desktopBridge.pickFile('All Files (*)');
        if (!path) return;

        const input = document.querySelector(`[data-setting-key="${settingKey}"]`);
        if (input) input.value = path;

        await saveSetting(settingKey, path);
    }

    // ------------------------------------------------------------------
    // Requirements check & action state
    // ------------------------------------------------------------------

    function updateAllActionState() {
        if (!_toolsConfig) return;
        const dataPath = window.desktopBridge ? window.desktopBridge.getCachedPath('data') : null;

        for (const tool of _toolsConfig.tools) {
            const btn = document.querySelector(`[data-tool-action="${tool.id}"]`);
            if (!btn) continue;

            let enabled = true;
            const requires = tool.requires || [];

            if (requires.includes('dataset') && !dataPath) {
                enabled = false;
            }
            if (requires.includes('source_excel') && (!_sourceInfo || !_sourceInfo.source_path)) {
                enabled = false;
            }

            // Check required visible controls have values
            for (const ctrl of tool.controls) {
                if (ctrl.type === 'hidden' || ctrl.type === 'action_button') continue;
                if (!ctrl.required) continue;
                const input = document.querySelector(
                    `[data-tool-control="${ctrl.id}"][data-tool-id="${tool.id}"]`
                );
                if (input && !input.value.trim()) {
                    enabled = false;
                }
            }

            btn.disabled = !enabled;
        }
    }

    // ------------------------------------------------------------------
    // Tool execution
    // ------------------------------------------------------------------

    async function executeTool(toolId) {
        if (!_toolsConfig || !window.desktopBridge) return;

        const tool = _toolsConfig.tools.find((t) => t.id === toolId);
        if (!tool) return;

        // Save current dataset silently before run if required (no dialog)
        if (tool.result.save_before_run && app && app.persistenceManager) {
            const data = app.persistenceManager.prepareDataForSaving({});
            const saved = await app.persistenceManager.fileService.quickSave(data, 'data');
            if (!saved) {
                app.showNotification('Cannot run tool: save the data file first so the path is known.', 'error');
                return;
            }
        }

        // Collect control values from the DOM
        const controlValues = {};
        for (const ctrl of tool.controls) {
            if (ctrl.type === 'hidden' || ctrl.type === 'action_button') continue;
            const input = document.querySelector(
                `[data-tool-control="${ctrl.id}"][data-tool-id="${toolId}"]`
            );
            if (input) {
                controlValues[ctrl.id] = input.value;
            }
        }

        // Build context for the bridge
        const dataPath = window.desktopBridge.getCachedPath('data') || null;
        const context = {
            control_values: controlValues,
            dataset_path: dataPath,
            source_info: _sourceInfo,
        };

        if (tool.type === 'interactive') {
            await executeInteractiveTool(tool, context);
        } else {
            await executeNonInteractiveTool(tool, context);
        }
    }

    async function executeNonInteractiveTool(tool, context) {
        openConsole(tool.label, false);

        try {
            const raw = await callBridgeSlot('resolve_and_run_tool', tool.id, JSON.stringify(context));
            const result = JSON.parse(raw);

            if (result.stdout) appendConsole(result.stdout);
            if (result.stderr) appendConsole((result.stdout ? '\n' : '') + result.stderr);

            if (result.returncode !== 0) {
                setConsoleStatus('Failed (exit code ' + result.returncode + ')');
                return;
            }

            await handleToolSuccess(tool, result.outputs, result.track_source, result.result);
        } catch (err) {
            setConsoleStatus('Error: ' + err.message);
        }
    }

    // Interactive tool state
    let _interactiveTool = null;
    let _interactiveOutputs = {};
    let _interactiveTrackSource = null;
    let _interactiveResult = {};

    // Register global callbacks for Python bridge push (via page.runJavaScript)
    window._onToolOutput = function (text) {
        onInteractiveOutput(text);
    };
    window._onToolFinished = function (exitCode) {
        onInteractiveFinished(exitCode);
    };

    async function executeInteractiveTool(tool, context) {
        openConsole(tool.label, true);

        const bridge = window._desktopBridgeRaw;
        if (!bridge) {
            setConsoleStatus('Error: Bridge not available');
            return;
        }

        _interactiveTool = tool;

        try {
            const raw = await callBridgeSlot('start_interactive_tool', tool.id, JSON.stringify(context));
            const result = JSON.parse(raw);

            if (!result.ok) {
                setConsoleStatus('Error: ' + result.error);
                _interactiveTool = null;
                return;
            }

            _interactiveOutputs = result.outputs || {};
            _interactiveTrackSource = result.track_source;
            _interactiveResult = result.result || {};
        } catch (err) {
            setConsoleStatus('Error: ' + err.message);
            _interactiveTool = null;
        }
    }

    function onInteractiveOutput(text) {
        appendConsole(text);
        // Show Yes/No buttons when a prompt arrives (line ends with ]: or ?)
        if (/\?\s*$|\]:\s*$/.test(text)) {
            setConsoleInputVisible(true);
        }
    }

    async function onInteractiveFinished(exitCode) {
        setConsoleInputVisible(false);

        if (exitCode !== 0) {
            setConsoleStatus('Failed (exit code ' + exitCode + ')');
        } else if (_interactiveTool) {
            await handleToolSuccess(
                _interactiveTool,
                _interactiveOutputs,
                _interactiveTrackSource,
                _interactiveResult,
            );
        } else {
            setConsoleStatus('✓ Complete');
        }

        _interactiveTool = null;
    }

    function sendInteractiveInput(text) {
        const bridge = window._desktopBridgeRaw;
        if (bridge && bridge.send_tool_input) {
            bridge.send_tool_input(text);
        }
        setConsoleInputVisible(false);
    }

    async function handleToolSuccess(tool, outputs, trackSource, resultSpec) {
        outputs = outputs || {};
        resultSpec = resultSpec || {};

        const dataPath = window.desktopBridge.getCachedPath('data') || null;

        if (resultSpec.auto_load && outputs.json_file) {
            const readResult = await window.desktopBridge.readFile(outputs.json_file);
            if (readResult.error) {
                setConsoleStatus('Failed to read output: ' + readResult.error);
                return;
            }
            const parsed = JSON.parse(readResult.content);
            window.desktopBridge.setCachedPath('data', outputs.json_file);
            app.loadDataFromJSON(parsed);

            if (typeof window._desktopUpdateFileInfo === 'function') {
                window._desktopUpdateFileInfo(outputs.json_file);
            }

            // Auto-load Obsidian companion view if available
            let viewLoaded = false;
            if (trackSource && trackSource.kind === 'obsidian-cli-converted' && trackSource.source_path) {
                const configDir = trackSource.source_path.replace(/[\\/][^\\/]+$/, '');
                const viewPath = configDir + '/obsidian.view.json';
                try {
                    const viewResult = await window.desktopBridge.readFile(viewPath);
                    if (!viewResult.error && viewResult.content) {
                        const viewConfig = JSON.parse(viewResult.content);
                        if (app.persistenceManager && typeof app.persistenceManager.applyViewConfiguration === 'function') {
                            app.persistenceManager.applyViewConfiguration(viewConfig);
                            viewLoaded = true;
                        }
                    }
                } catch (_) { /* view file not found — fine */ }
            }

            const itemCount = parsed.data ? parsed.data.length : 0;
            setConsoleStatus('✓ Loaded ' + itemCount + ' items' + (viewLoaded ? ' · view applied' : ''));
        } else {
            const msg = resultSpec.success_message || '✓ Complete';
            setConsoleStatus(msg);
        }

        if (trackSource && outputs.json_file) {
            await saveSourceInfo(outputs.json_file, trackSource);
        } else if (trackSource && dataPath) {
            await saveSourceInfo(dataPath, trackSource);
        }

        await loadSourceInfo();
        _syncSourceContextToMeta();
        updateAllActionState();
    }

    function callBridgeSlot(method, ...args) {
        return new Promise((resolve) => {
            const bridge = window._desktopBridgeRaw;
            if (bridge && bridge[method]) {
                bridge[method].call(bridge, ...args, resolve);
            } else {
                resolve(JSON.stringify({ stdout: '', stderr: 'Bridge not available', returncode: -1 }));
            }
        });
    }

    // ------------------------------------------------------------------
    // Initialisation
    // ------------------------------------------------------------------

    async function initToolsPanel() {
        if (!window.desktopBridge) return;
        if (typeof app === 'undefined' || !app) return;

        const container = document.querySelector('.desktop-tools');
        if (!container) return;

        try {
            const raw = await callBridgeSlot('get_tools');
            _toolsConfig = JSON.parse(raw);

            if (_toolsConfig.error) {
                console.error('[DESKTOP-TOOLS] Failed to load tools:', _toolsConfig.error);
                return;
            }

            container.style.display = '';
            await loadSourceInfo();
            _syncSourceContextToMeta();
            renderToolsPanel(container);

            // Hide browser-mode controls that are now replaced by the tools panel
            const sectionHeader = document.querySelector('.control-area-data > .section-header');
            const fileControls = document.querySelector('.file-controls');
            const viewControls = document.querySelector('.data-view-controls');
            if (sectionHeader) sectionHeader.style.display = 'none';
            if (fileControls) fileControls.style.display = 'none';
            if (viewControls) viewControls.style.display = 'none';

            // Wire console popup buttons for desktop mode
            const closeBtn = document.getElementById('console-popup-close');
            const yesBtn = document.getElementById('console-popup-yes');
            const noBtn = document.getElementById('console-popup-no');
            if (closeBtn) closeBtn.addEventListener('click', () => {
                document.getElementById('console-popup-overlay').style.display = 'none';
            });
            if (yesBtn) yesBtn.addEventListener('click', () => sendInteractiveInput('y'));
            if (noBtn) noBtn.addEventListener('click', () => sendInteractiveInput('n'));

            console.log('[DESKTOP-TOOLS] Tools panel initialised with', (_toolsConfig.tools || []).length, 'tools');
        } catch (err) {
            console.error('[DESKTOP-TOOLS] Init failed:', err);
        }
    }

    // Wait for both DOMContentLoaded and the bridge
    let _domReady = false;
    let _bridgeReady = false;

    function tryInit() {
        if (_domReady && _bridgeReady) {
            // Small delay to let desktop-bridge init first
            setTimeout(initToolsPanel, 50);
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

    window.addEventListener('desktop-bridge-ready', () => {
        _bridgeReady = true;
        tryInit();
    });

    // If bridge was already initialised
    if (window.desktopBridge) {
        _bridgeReady = true;
        tryInit();
    }
})();
