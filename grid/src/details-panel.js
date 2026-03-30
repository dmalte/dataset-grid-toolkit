// Copyright (c) 2026 Malte Doerper. MIT License. See LICENSE file.

// V6 — Details Panel Manager
// Multi-mode right-side panel: Card Editor, Relations, Header Values, Markdown View
class DetailsPanelManager {
    constructor(app) {
        this.app = app;

        // Panel state
        this.currentMode = null; // 'editor' | 'relations' | 'header-values' | 'markdown'
        this.currentItem = null;
        this.createDraftItem = null;
        this.headerContext = null; // { fieldName, headerType, headerValue } for header-values mode

        this.panel = document.getElementById('details-panel');
        this.initializePanel();
    }

    // ===== INITIALIZATION =====

    initializePanel() {
        if (!this.panel) return;

        // Close button
        const closeBtn = this.panel.querySelector('.details-panel-close');
        if (closeBtn) {
            closeBtn.addEventListener('click', () => this.close());
        }

        // Tab buttons
        this.panel.querySelectorAll('.details-tab').forEach(tab => {
            tab.addEventListener('click', () => {
                const mode = tab.dataset.mode;
                if (mode && this.currentItem) {
                    this.switchMode(mode);
                }
            });
        });

        // Editor save/cancel
        const editorSave = document.getElementById('details-editor-save');
        if (editorSave) {
            editorSave.addEventListener('click', () => this.saveEditor());
        }
        const editorCancel = document.getElementById('details-editor-cancel');
        if (editorCancel) {
            editorCancel.addEventListener('click', () => this.close());
        }

        // Header values save/cancel/add
        const hvSave = document.getElementById('header-values-save');
        if (hvSave) {
            hvSave.addEventListener('click', () => this.saveHeaderValues());
        }
        const hvCancel = document.getElementById('header-values-cancel');
        if (hvCancel) {
            hvCancel.addEventListener('click', () => this.close());
        }
        const hvAddBtn = document.getElementById('header-value-add-btn');
        if (hvAddBtn) {
            hvAddBtn.addEventListener('click', () => this.addHeaderValue());
        }
        const hvInput = document.getElementById('header-value-new-input');
        if (hvInput) {
            hvInput.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    this.addHeaderValue();
                }
            });
        }

        // Escape to close
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape' && this.isOpen()) {
                this.close();
            }
        });

        // Markdown save button
        const mdSaveBtn = document.getElementById('markdown-save-btn');
        if (mdSaveBtn) {
            mdSaveBtn.addEventListener('click', () => this.saveMarkdown());
        }
    }

    // ===== OPEN / CLOSE =====

    isOpen() {
        return this.panel && this.panel.classList.contains('open');
    }

    refreshForSelection() {
        if (!this.isOpen() || !this.currentMode) return;

        const selectedItems = this.app.getSelectedItems();
        if (selectedItems.length > 0) {
            // Use the last selected item as the current item
            this.currentItem = selectedItems[selectedItems.length - 1];
        } else {
            return;
        }

        // Update header
        const titleEl = this.panel.querySelector('.details-panel-title');
        if (titleEl) titleEl.textContent = this.currentItem.title || this.currentItem.id || 'Item';
        const subtitleEl = this.panel.querySelector('.details-panel-subtitle');
        if (subtitleEl) {
            subtitleEl.textContent = selectedItems.length > 1
                ? `${selectedItems.length} items selected`
                : (this.currentItem.id || '');
        }

        // Re-populate current mode
        this.switchMode(this.currentMode);
    }

    openForItem(item, mode = 'editor') {
        if (!this.panel || !item) return;

        this.currentItem = item;
        this.createDraftItem = null;
        this.headerContext = null;

        // Update header
        const titleEl = this.panel.querySelector('.details-panel-title');
        if (titleEl) titleEl.textContent = item.title || item.id || 'Item';
        const subtitleEl = this.panel.querySelector('.details-panel-subtitle');
        if (subtitleEl) subtitleEl.textContent = item.id || '';

        // Show focus button for item modes
        const focusBtn = document.getElementById('details-panel-focus-btn');
        if (focusBtn) focusBtn.style.display = '';

        this.switchMode(mode);
        this.panel.classList.remove('hidden');
        this.panel.classList.add('open');
    }

    openForCreate(initialValues = {}) {
        if (!this.panel) return;

        this.currentItem = null;
        this.createDraftItem = this.buildDraftItem(initialValues);
        this.headerContext = null;

        const titleEl = this.panel.querySelector('.details-panel-title');
        if (titleEl) titleEl.textContent = 'Create New Item';

        const subtitleEl = this.panel.querySelector('.details-panel-subtitle');
        if (subtitleEl) {
            const contextParts = [];
            if (initialValues && initialValues.id) {
                contextParts.push(`Suggested ID: ${initialValues.id}`);
            }
            subtitleEl.textContent = contextParts.join(' · ');
        }

        const focusBtn = document.getElementById('details-panel-focus-btn');
        if (focusBtn) focusBtn.style.display = 'none';

        this.switchMode('editor');
        this.panel.classList.remove('hidden');
        this.panel.classList.add('open');
    }

    openForHeader(fieldName, headerType, headerValue) {
        if (!this.panel) return;

        this.currentItem = null;
        this.createDraftItem = null;
        this.headerContext = { fieldName, headerType, headerValue };

        // Update header
        const titleEl = this.panel.querySelector('.details-panel-title');
        if (titleEl) {
            if (headerType === 'column') {
                titleEl.textContent = `Column: ${headerValue || '(empty)'}`;
            } else if (headerType === 'row') {
                titleEl.textContent = `Row: ${headerValue || '(empty)'}`;
            } else {
                titleEl.textContent = `Field Values: ${fieldName}`;
            }
        }
        const subtitleEl = this.panel.querySelector('.details-panel-subtitle');
        if (subtitleEl) {
            if (headerType === 'filter' && headerValue !== undefined && headerValue !== null) {
                subtitleEl.textContent = `Right-clicked filter value: ${headerValue || '(empty)'}`;
            } else {
                subtitleEl.textContent = `Field: ${fieldName}`;
            }
        }

        // Hide focus button for header mode
        const focusBtn = document.getElementById('details-panel-focus-btn');
        if (focusBtn) focusBtn.style.display = 'none';

        this.switchMode('header-values');
        this.panel.classList.remove('hidden');
        this.panel.classList.add('open');
    }

    close() {
        if (!this.panel) return;
        this.destroyEasyMDE();
        this.panel.classList.add('hidden');
        this.panel.classList.remove('open');
        this.currentItem = null;
        this.createDraftItem = null;
        this.headerContext = null;
        this.currentMode = null;
    }

    // ===== MODE SWITCHING =====

    switchMode(mode) {
        // Destroy EasyMDE when leaving markdown mode
        if (this.currentMode === 'markdown' && mode !== 'markdown') {
            this.destroyEasyMDE();
        }

        this.currentMode = mode;

        // Hide all mode sections
        this.panel.querySelectorAll('.details-mode').forEach(el => {
            el.classList.add('hidden');
        });

        // Update tab active states
        this.panel.querySelectorAll('.details-tab').forEach(tab => {
            tab.classList.toggle('active', tab.dataset.mode === mode);
        });

        // Show/hide tabs based on context
        this.updateTabVisibility();

        // Show the target mode section and populate
        const modeEl = document.getElementById(`details-mode-${mode}`);
        if (modeEl) {
            modeEl.classList.remove('hidden');
        }

        if (mode === 'editor' && (this.currentItem || this.createDraftItem)) {
            this.populateEditor(this.currentItem || this.createDraftItem);
        } else if (mode === 'markdown') {
            this.populateMarkdown(this.currentItem);
        } else if (mode === 'relations') {
            this.populateRelationsMode();
        } else if (mode === 'header-values' && this.headerContext) {
            this.populateHeaderValues();
        }
    }

    updateTabVisibility() {
        const tabs = this.panel.querySelector('.details-panel-tabs');
        if (!tabs) return;

        if (this.headerContext || this.createDraftItem) {
            // Header mode — hide all tabs except header-values
            tabs.style.display = 'none';
        } else {
            tabs.style.display = '';
        }
    }

    populateRelationsMode() {
        const mgr = this.app.relationUIManager;
        if (!mgr) return;

        const modeEl = document.getElementById('details-mode-relations');
        if (!modeEl) return;

        const selectedItems = this.app.getSelectedItems();
        const items = selectedItems.length > 1
            ? selectedItems
            : (this.currentItem ? [this.currentItem] : []);

        if (items.length === 0) return;

        if (items.length === 1) {
            // Single item — use existing single-item population
            mgr.panelItem = items[0];
            mgr.populateRelationshipTree(modeEl, items[0]);
            mgr.populateRelationSections(modeEl, items[0], null);
            mgr.populateRelationTypeDropdown();
            mgr.hideAddRelationForm();
        } else {
            // Multi-card — render a section per card
            const detailsEl = modeEl.querySelector('.relation-panel-details');
            if (detailsEl) detailsEl.innerHTML = '';

            const sectionsEl = modeEl.querySelector('.relation-panel-sections');
            if (sectionsEl) sectionsEl.innerHTML = '';

            // Set panelItem to last selected for add-relation form context
            mgr.panelItem = items[items.length - 1];

            items.forEach(item => {
                const itemId = this.app.getItemIdentity(item) || item.id;
                const itemTitle = item.title || itemId;

                // Card heading
                const heading = document.createElement('div');
                heading.className = 'relation-multi-card-heading';
                heading.textContent = itemTitle;

                // Wrapper with expected child structure for relation methods
                const wrapper = document.createElement('div');
                const treeChild = document.createElement('div');
                treeChild.className = 'relation-panel-details';
                const secChild = document.createElement('div');
                secChild.className = 'relation-panel-sections';
                wrapper.appendChild(treeChild);
                wrapper.appendChild(secChild);

                mgr.populateRelationshipTree(wrapper, item);
                mgr.populateRelationSections(wrapper, item, null);

                if (sectionsEl) {
                    sectionsEl.appendChild(heading);
                    while (treeChild.firstChild) {
                        sectionsEl.appendChild(treeChild.firstChild);
                    }
                    while (secChild.firstChild) {
                        sectionsEl.appendChild(secChild.firstChild);
                    }
                }
            });

            mgr.populateRelationTypeDropdown();
            mgr.hideAddRelationForm();
        }
    }

    // ===== CARD EDITOR MODE =====

    populateEditor(item) {
        const container = document.querySelector('#details-mode-editor .details-editor-fields');
        if (!container) return;
        container.innerHTML = '';

        const fields = this.getEditableFields();

        fields.forEach(fieldName => {
            const fieldType = this.app.fieldTypes.get(fieldName);
            const value = this.app.getFieldValue(item, fieldName);

            const fieldDiv = document.createElement('div');
            fieldDiv.className = 'details-editor-field';

            const label = document.createElement('label');
            label.className = 'details-editor-label';
            label.textContent = this.formatFieldLabel(fieldName);

            const input = this.createFieldInput(fieldName, fieldType, value);
            label.htmlFor = input.id;

            fieldDiv.appendChild(label);
            fieldDiv.appendChild(input);
            container.appendChild(fieldDiv);
        });

        // Supplementary sections below the editable data fields.
        this.populateEditorGroups(item);
        this.populateReadOnlyDerivedFields(item);
    }

    getEditableFields() {
        const allFields = Array.from(this.app.fieldTypes.keys());
        const isEditable = (fn) => {
            const sf = this.app.schemaFields[fn];
            return !(sf && (sf.visible === false || sf.editable === false)) && this.app.getFieldKind(fn) === 'data';
        };
        const prioritizedFields = ['id', 'title'];
        const prioritizedSet = new Set(prioritizedFields);
        const orderedFields = prioritizedFields.filter(fn => allFields.includes(fn) && isEditable(fn));
        const remainingFields = allFields
            .filter(fn => !prioritizedSet.has(fn))
            .filter(fn => this.app.fieldTypes.get(fn) !== 'structured')
            .filter(isEditable)
            .sort();

        return [...orderedFields, ...remainingFields];
    }

    getReadOnlyDerivedFields() {
        return Array.from(this.app.fieldTypes.keys())
            .filter((fieldName) => this.app.fieldTypes.get(fieldName) !== 'structured')
            .filter((fieldName) => {
                const sf = this.app.schemaFields[fieldName];
                if (sf && sf.visible === false) {
                    return false;
                }

                const kind = this.app.getFieldKind(fieldName);
                return kind === 'derived' || kind === 'relationship';
            })
            .sort();
    }

    buildDraftItem(initialValues = {}) {
        const draftItem = {};

        this.getEditableFields().forEach((fieldName) => {
            const fieldType = this.app.fieldTypes.get(fieldName);

            if (Object.prototype.hasOwnProperty.call(initialValues, fieldName)) {
                draftItem[fieldName] = initialValues[fieldName];
                return;
            }

            draftItem[fieldName] = fieldType === 'multi-value' ? [] : null;
        });

        return draftItem;
    }

    createFieldInput(fieldName, fieldType, value) {
        let input;

        if (fieldType === 'multi-value') {
            input = document.createElement('textarea');
            input.rows = fieldName === 'tags' ? 2 : 3;
            input.value = Array.isArray(value) ? value.join('\n') : (value || '');
        } else if (fieldType === 'structured') {
            input = document.createElement('textarea');
            input.rows = 5;
            input.className = 'details-input details-json-input';
            input.value = typeof value === 'object' && value !== null
                ? JSON.stringify(value, null, 2)
                : (value || '');
        } else {
            input = document.createElement('input');
            input.type = this.getInputType(fieldName);
            input.value = typeof this.app.getEditorInputValue === 'function'
                ? this.app.getEditorInputValue(fieldName, value)
                : (value || '');
        }

        input.classList.add('details-input');
        input.name = fieldName;
        input.id = `details-field-${fieldName}`;

        return input;
    }

    getInputType(fieldName) {
        if (typeof this.app.getFormInputType === 'function') {
            return this.app.getFormInputType(fieldName);
        }
        return 'text';
    }

    formatFieldLabel(fieldName) {
        return fieldName
            .replace(/([A-Z])/g, ' $1')
            .replace(/[_-]/g, ' ')
            .replace(/^\w/, c => c.toUpperCase())
            .trim();
    }

    populateEditorGroups(item) {
        const container = document.querySelector('#details-mode-editor .details-editor-groups');
        if (!container) return;
        container.innerHTML = '';

        if (this.createDraftItem) {
            return;
        }

        const allGroups = this.app.getGroups ? this.app.getGroups() : [];
        if (allGroups.length === 0) return;

        const itemId = this.app.getItemIdentity(item) || item.id;
        const itemGroups = this.app.getItemGroups ? this.app.getItemGroups(item) : [];

        const groupLabel = document.createElement('div');
        groupLabel.className = 'details-editor-label';
        groupLabel.textContent = 'Groups';
        container.appendChild(groupLabel);

        const groupValue = document.createElement('div');
        groupValue.className = 'details-editor-group-chips';

        if (itemGroups.length === 0) {
            const empty = document.createElement('span');
            empty.className = 'details-groups-empty';
            empty.textContent = '(none)';
            groupValue.appendChild(empty);
        } else {
            itemGroups.forEach(g => {
                const isManual = Array.isArray(g.manualMembers) && g.manualMembers.includes(itemId);
                const chip = document.createElement('span');
                chip.className = 'details-group-chip';
                chip.style.backgroundColor = g.color;
                chip.textContent = g.name;
                chip.title = isManual ? 'Click to remove from group' : 'Rule-based (cannot remove)';
                if (isManual) {
                    chip.classList.add('is-removable');
                    chip.addEventListener('click', () => {
                        this.app.removeItemFromGroup(g.id, itemId);
                        this.populateEditorGroups(item);
                    });
                } else {
                    chip.classList.add('is-rule-based');
                }
                groupValue.appendChild(chip);
            });
        }

        container.appendChild(groupValue);
    }

    populateReadOnlyDerivedFields(item) {
        const container = document.querySelector('#details-mode-editor .details-editor-groups');
        if (!container || this.createDraftItem) return;

        const derivedFields = this.getReadOnlyDerivedFields();
        if (derivedFields.length === 0) return;

        const section = document.createElement('div');
        section.className = 'details-derived-section';

        const title = document.createElement('div');
        title.className = 'details-editor-label details-derived-section-title';
        title.textContent = 'Derived Fields';
        section.appendChild(title);

        const list = document.createElement('div');
        list.className = 'details-derived-list';

        derivedFields.forEach((fieldName) => {
            const row = document.createElement('div');
            row.className = 'details-derived-field';

            const label = document.createElement('div');
            label.className = 'details-derived-label';
            label.textContent = this.formatFieldLabel(fieldName);

            const value = document.createElement('div');
            value.className = 'details-derived-value';
            value.textContent = this.formatReadOnlyFieldValue(item, fieldName);

            row.appendChild(label);
            row.appendChild(value);
            list.appendChild(row);
        });

        section.appendChild(list);
        container.appendChild(section);
    }

    formatReadOnlyFieldValue(item, fieldName) {
        const rawValue = this.app.getFieldValue(item, fieldName);
        if (rawValue === undefined || rawValue === null || rawValue === '') {
            return '(empty)';
        }

        if (typeof this.app.getDisplayValue === 'function') {
            const displayValue = this.app.getDisplayValue(item, fieldName);
            return displayValue === undefined || displayValue === null || displayValue === ''
                ? '(empty)'
                : String(displayValue);
        }

        if (Array.isArray(rawValue)) {
            return rawValue.join(', ');
        }

        if (typeof rawValue === 'object') {
            return JSON.stringify(rawValue);
        }

        return String(rawValue);
    }

    saveEditor() {
        if (this.createDraftItem) {
            this.saveCreateEditor();
            return;
        }

        if (!this.currentItem) return;

        const result = this.collectEditorFormData();
        if (!result) {
            return;
        }

        const { container, formData } = result;

        const itemId = this.app.getItemIdentity(this.currentItem) || this.currentItem.id;
        const previousDistinctValues = this.app.snapshotDistinctValues();
        const success = this.app.updateItem(itemId, formData);
        if (!success) {
            this.app.showNotification('Failed to update item', 'error');
            return;
        }

        this.app.analyzeFields();
        this.app.reconcileFiltersWithDistinctValues(previousDistinctValues);
        this.app.updateViewConfiguration();
        this.app.renderSlicers();
        this.app.renderTags();
        this.app.renderGrid();

        // Refresh the panel with updated item
        const updatedItem = this.app.getEffectiveItemById(
            this.app.getItemIdentity(this.currentItem) || this.currentItem.id
        );
        if (updatedItem) {
            this.currentItem = updatedItem;
            this.populateEditor(updatedItem);
        }

        this.app.showNotification(`Updated ${formData.id || itemId || 'item'}`, 'success');
    }

    collectEditorFormData() {
        const container = document.querySelector('#details-mode-editor .details-editor-fields');
        if (!container) return null;

        const formData = {};
        const inputs = container.querySelectorAll('input[name], textarea[name], select[name]');

        for (const field of inputs) {
            const fieldName = field.name;
            const fieldType = this.app.fieldTypes.get(fieldName);
            const parsed = this.parseInputValue(fieldName, fieldType, field.value);

            if (parsed.error) {
                this.app.showNotification(parsed.error, 'error');
                field.focus();
                return null;
            }

            formData[fieldName] = parsed.value;
        }

        return { container, formData };
    }

    saveCreateEditor() {
        const result = this.collectEditorFormData();
        if (!result) {
            return;
        }

        const { container, formData } = result;
        const validation = typeof this.app.validateItemFormData === 'function'
            ? this.app.validateItemFormData(formData, true)
            : { isValid: true, errors: [] };

        if (!validation.isValid) {
            const firstError = validation.errors[0];
            if (firstError) {
                this.app.showNotification(firstError.message, 'error');
                const input = container.querySelector(`[name="${firstError.field}"]`);
                if (input) {
                    input.focus();
                }
            }
            return;
        }

        if (formData.id && this.app.dataset.some((item) => item.id === formData.id)) {
            this.app.showNotification('ID already exists', 'error');
            const idInput = container.querySelector('[name="id"]');
            if (idInput) {
                idInput.focus();
            }
            return;
        }

        const previousDistinctValues = this.app.snapshotDistinctValues();
        const newItemId = this.app.addItem(formData);
        if (!newItemId) {
            this.app.showNotification('Failed to create item', 'error');
            return;
        }

        this.app.analyzeFields();
        this.app.reconcileFiltersWithDistinctValues(previousDistinctValues);
        this.app.updateViewConfiguration();
        this.app.renderSlicers();
        this.app.renderTags();
        this.app.renderGrid();

        const createdItem = this.app.getEffectiveItemById(newItemId);
        this.createDraftItem = null;
        if (createdItem) {
            this.openForItem(createdItem, 'editor');
        }

        this.app.showNotification(`Created ${newItemId}`, 'success');
    }

    parseInputValue(fieldName, fieldType, rawValue) {
        if (fieldType === 'multi-value') {
            return {
                value: rawValue
                    .split(/[,\n]+/)
                    .map(entry => entry.trim())
                    .filter(entry => entry.length > 0)
            };
        }
        if (fieldType === 'structured') {
            if (!rawValue.trim()) return { value: null };
            try {
                return { value: JSON.parse(rawValue) };
            } catch (error) {
                return { error: `Invalid JSON in ${fieldName}: ${error.message}` };
            }
        }
        return { value: rawValue || null };
    }

    // ===== HEADER VALUES EDITOR MODE =====

    populateHeaderValues() {
        const container = this.panel.querySelector('.header-values-list');
        if (!container || !this.headerContext) return;
        container.innerHTML = '';

        const { fieldName } = this.headerContext;
        const schemaField = this.app.schemaFields[fieldName] || {};
        const currentValidValues = Array.isArray(schemaField.validValues)
            ? [...schemaField.validValues]
            : [];
        const configuredValueColors = schemaField && typeof schemaField.valueColors === 'object' && schemaField.valueColors !== null
            ? schemaField.valueColors
            : {};

        const dataValues = this.app.distinctValues.get(fieldName) || [];
        const allValues = Array.isArray(schemaField.validValues)
            ? currentValidValues
            : Array.from(new Set([...currentValidValues, ...dataValues])).sort();

        allValues.forEach(value => {
            const isSchemaValue = currentValidValues.includes(value);
            container.appendChild(this.createHeaderValueRow(
                value,
                configuredValueColors[String(value)] || '',
                isSchemaValue
            ));
        });

        // Clear the add input
        const addInput = document.getElementById('header-value-new-input');
        if (addInput) addInput.value = '';
    }

    createHeaderValueRow(value, configuredColor = '', isSchemaValue = false) {
        const row = document.createElement('div');
        row.className = 'header-value-row';
        row.dataset.value = String(value);

        const label = document.createElement('span');
        label.className = 'header-value-label';
        label.textContent = value;

        if (isSchemaValue) {
            label.classList.add('is-schema-value');
        }

        const controls = document.createElement('div');
        controls.className = 'header-value-row-controls';

        const colorButton = document.createElement('button');
        colorButton.type = 'button';
        colorButton.className = 'header-value-color-trigger';
        colorButton.title = 'Choose color for this value';

        const colorInput = document.createElement('input');
        colorInput.type = 'color';
        colorInput.className = 'header-value-color-input';
        colorInput.value = configuredColor || '#808080';
        colorInput.setAttribute('aria-label', `Color for ${value}`);

        const clearBtn = document.createElement('button');
        clearBtn.type = 'button';
        clearBtn.className = 'header-value-clear-btn';
        clearBtn.textContent = 'Clear';
        clearBtn.title = 'Remove configured color';

        if (configuredColor) {
            row.dataset.valueColor = configuredColor;
        } else {
            row.dataset.valueColor = '';
            row.classList.add('is-color-unset');
        }

        this.updateHeaderValueColorButton(colorButton, configuredColor);

        colorButton.addEventListener('click', () => {
            colorInput.click();
        });

        colorInput.addEventListener('input', () => {
            row.dataset.valueColor = colorInput.value;
            row.classList.remove('is-color-unset');
            this.updateHeaderValueColorButton(colorButton, colorInput.value);
        });

        clearBtn.addEventListener('click', () => {
            row.dataset.valueColor = '';
            row.classList.add('is-color-unset');
            colorInput.value = '#808080';
            this.updateHeaderValueColorButton(colorButton, '');
        });

        const removeBtn = document.createElement('button');
        removeBtn.className = 'header-value-remove-btn';
        removeBtn.textContent = '×';
        removeBtn.title = 'Remove from valid values';
        removeBtn.addEventListener('click', () => {
            row.remove();
        });

        controls.appendChild(colorButton);
        controls.appendChild(colorInput);
        controls.appendChild(clearBtn);
        controls.appendChild(removeBtn);
        row.appendChild(label);
        row.appendChild(controls);
        return row;
    }

    updateHeaderValueColorButton(button, configuredColor) {
        if (!button) {
            return;
        }

        const normalizedColor = typeof configuredColor === 'string' ? configuredColor.trim() : '';
        button.classList.toggle('is-color-unset', !normalizedColor);
        button.textContent = normalizedColor ? '' : '×';
        button.style.backgroundColor = normalizedColor || '#d0d5db';
        button.style.color = normalizedColor ? 'transparent' : '#5f6368';
    }

    addHeaderValue() {
        const addInput = document.getElementById('header-value-new-input');
        if (!addInput) return;

        const value = addInput.value.trim();
        if (!value) return;

        const container = this.panel.querySelector('.header-values-list');
        if (!container) return;

        // Check for duplicates
        const existing = container.querySelectorAll('.header-value-label');
        for (const el of existing) {
            if (el.textContent === value) {
                this.app.showNotification('Value already exists', 'warning');
                return;
            }
        }

        container.appendChild(this.createHeaderValueRow(value));

        addInput.value = '';
        addInput.focus();
    }

    saveHeaderValues() {
        if (!this.headerContext) return;

        const container = this.panel.querySelector('.header-values-list');
        if (!container) return;

        const values = [];
        const valueColors = {};
        container.querySelectorAll('.header-value-row').forEach((row) => {
            const label = row.querySelector('.header-value-label');
            const value = label ? label.textContent.trim() : '';
            if (!value) {
                return;
            }

            values.push(value);
            const configuredColor = String(row.dataset.valueColor || '').trim();
            if (configuredColor) {
                valueColors[value] = configuredColor;
            }
        });

        const { fieldName } = this.headerContext;

        // Update schema validValues
        if (!this.app.schemaFields[fieldName]) {
            this.app.schemaFields[fieldName] = {};
        }
        this.app.schemaFields[fieldName].validValues = values;
        if (Object.keys(valueColors).length > 0) {
            this.app.schemaFields[fieldName].valueColors = valueColors;
        } else {
            delete this.app.schemaFields[fieldName].valueColors;
        }

        const previousDistinctValues = typeof this.app.snapshotDistinctValues === 'function'
            ? this.app.snapshotDistinctValues()
            : new Map();

        // Re-analyze fields to pick up the new validValues
        this.app.analyzeFields();
        if (typeof this.app.reconcileFiltersWithDistinctValues === 'function') {
            this.app.reconcileFiltersWithDistinctValues(previousDistinctValues);
        }
        this.app.updateFilteredData();
        this.app.renderSlicers();
        this.app.renderGrid();
        this.populateHeaderValues();

        this.app.showNotification(`Updated valid values for ${fieldName}`, 'success');
    }

    // ===== MARKDOWN VIEW MODE (EasyMDE) =====

    populateMarkdown(item) {
        // Destroy previous EasyMDE instance if it exists
        if (this.easyMDE) {
            this.easyMDE.toTextArea();
            this.easyMDE = null;
        }

        const textarea = document.getElementById('markdown-easymde');
        if (!textarea) return;

        // Collect markdown content: use body_md from selected cards or current item
        const markdownText = this.collectMarkdownContent(item);

        textarea.value = markdownText;

        // Initialize EasyMDE in preview-only mode
        if (typeof EasyMDE !== 'undefined') {
            this.easyMDE = new EasyMDE({
                element: textarea,
                initialValue: markdownText,
                toolbar: ['preview', 'side-by-side', 'fullscreen', '|', 'guide'],
                status: false,
                spellChecker: false,
                minHeight: '200px',
                autoDownloadFontAwesome: true
            });
            // Auto-toggle preview on open
            if (!this.easyMDE.isPreviewActive()) {
                this.easyMDE.togglePreview();
            }
        }
    }

    collectMarkdownContent(item) {
        // Check if multiple cards are selected
        const selectedItems = this.app.getSelectedItems ? this.app.getSelectedItems() : [];

        let items;
        if (selectedItems.length > 1) {
            items = selectedItems;
        } else if (item) {
            items = [item];
        } else {
            return '';
        }

        const sections = items.map(it => {
            const title = it.title || it.id || 'Untitled';
            const bodyMd = it.body_md || '';
            if (bodyMd) {
                return `## ${title}\n\n${bodyMd}`;
            }
            // Fallback: look for any _md fields
            const mdFields = Object.keys(it).filter(k => k.endsWith('_md') && typeof it[k] === 'string' && it[k]);
            if (mdFields.length > 0) {
                return `## ${title}\n\n` + mdFields.map(f => it[f]).join('\n\n');
            }
            return '';
        }).filter(Boolean);

        return sections.join('\n\n---\n\n');
    }

    destroyEasyMDE() {
        if (this.easyMDE) {
            this.easyMDE.toTextArea();
            this.easyMDE = null;
        }
    }

    saveMarkdown() {
        if (!this.easyMDE) return;

        const fullText = this.easyMDE.value();
        const selectedItems = this.app.getSelectedItems ? this.app.getSelectedItems() : [];
        const items = selectedItems.length > 1 ? selectedItems : (this.currentItem ? [this.currentItem] : []);

        if (items.length === 0) return;

        if (items.length === 1) {
            // Single item — write directly to body_md
            const item = items[0];
            const itemId = this.app.getItemIdentity(item) || item.id;
            const titlePrefix = `## ${item.title || item.id || 'Untitled'}\n\n`;
            const bodyContent = fullText.startsWith(titlePrefix)
                ? fullText.slice(titlePrefix.length)
                : fullText;
            this.app.updateItem(itemId, { body_md: bodyContent }, { render: false });
        } else {
            // Multiple items — split by "---" separator and write back to each
            const sections = fullText.split(/\n\n---\n\n/);
            items.forEach((item, i) => {
                if (i >= sections.length) return;
                const section = sections[i];
                const itemId = this.app.getItemIdentity(item) || item.id;
                const titlePrefix = `## ${item.title || item.id || 'Untitled'}\n\n`;
                const bodyContent = section.startsWith(titlePrefix)
                    ? section.slice(titlePrefix.length)
                    : section;
                this.app.updateItem(itemId, { body_md: bodyContent }, { render: false });
            });
        }

        // Rebuild and re-render once
        this.app.rebuildEffectiveDataset({ render: true });
        this.app.showNotification('Markdown saved', 'success');
    }
}

// Initialize after DOM and other managers are ready
document.addEventListener('DOMContentLoaded', () => {
    if (typeof app !== 'undefined' && app) {
        app.detailsPanelManager = new DetailsPanelManager(app);
    }
});
