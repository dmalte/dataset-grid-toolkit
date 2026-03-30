// Copyright (c) 2026 Malte Doerper. MIT License. See LICENSE file.

// Card Interactions Implementation
// Handles drag & drop, tooltips, URL navigation, and card interactions

class InteractionManager {
    constructor(app) {
        this.app = app;
        this.tooltip = null;
        this.previewTooltip = null;
        this.editorOverlay = null;
        this.editorTab = null;
        this.filterSetupDialog = null;
        this.filterSetupContext = null;
        this.lassoSelection = null;
        this.activeTooltipItemId = null;
        this.activeTooltipMode = null;
        this.lastFocusedEditorElement = null;
        this.pendingEditorActivation = false;
        this.urlConfig = {
            prefix: '',
            suffix: '',
            useUrlField: true
        };
        this.initializeInteractions();
    }
    
    initializeInteractions() {
        this.setupTooltip();
        this.initializeFilterSetupDialog();
        this.initializeLassoSelection();
        this.loadUrlConfiguration();
        this.setupTooltipDismissal();
        this.setupCreateItemTrigger();
    }

    initializeLassoSelection() {
        const gridContainer = document.getElementById('grid-container');
        if (!gridContainer || gridContainer.dataset.lassoInitialized === 'true') {
            return;
        }

        gridContainer.dataset.lassoInitialized = 'true';
        gridContainer.addEventListener('mousedown', (event) => {
            if (event.button !== 0 || this.app.cellRenderMode === 'table') {
                return;
            }

            if (!event.target.closest('.grid-cell') || event.target.closest('.card') || event.target.closest('.grid-cell-table')) {
                return;
            }

            const containerRect = gridContainer.getBoundingClientRect();
            this.hideTooltip('preview');
            this.dismissCardContextMenu();

            const selectionBox = document.createElement('div');
            selectionBox.className = 'lasso-selection-box';
            gridContainer.appendChild(selectionBox);

            this.lassoSelection = {
                container: gridContainer,
                box: selectionBox,
                startX: event.clientX,
                startY: event.clientY,
                containerRect,
                moved: false
            };

            gridContainer.classList.add('is-lasso-selecting');
            event.preventDefault();
        });

        document.addEventListener('mousemove', (event) => {
            if (!this.lassoSelection) {
                return;
            }

            const { startX, startY, containerRect, box } = this.lassoSelection;
            const left = Math.min(startX, event.clientX);
            const top = Math.min(startY, event.clientY);
            const width = Math.abs(event.clientX - startX);
            const height = Math.abs(event.clientY - startY);

            this.lassoSelection.moved = this.lassoSelection.moved || width > 4 || height > 4;
            box.style.left = `${left - containerRect.left}px`;
            box.style.top = `${top - containerRect.top}px`;
            box.style.width = `${width}px`;
            box.style.height = `${height}px`;
        });

        document.addEventListener('mouseup', (event) => {
            if (!this.lassoSelection) {
                return;
            }

            const { container, box, startX, startY, moved } = this.lassoSelection;
            const left = Math.min(startX, event.clientX);
            const right = Math.max(startX, event.clientX);
            const top = Math.min(startY, event.clientY);
            const bottom = Math.max(startY, event.clientY);

            if (moved) {
                const selectedItemIds = Array.from(container.querySelectorAll('.card[data-item-id]'))
                    .filter((card) => !card.classList.contains('grid-table-row'))
                    .filter((card) => {
                        const rect = card.getBoundingClientRect();
                        return rect.right >= left && rect.left <= right && rect.bottom >= top && rect.top <= bottom;
                    })
                    .map((card) => card.dataset.itemId)
                    .filter(Boolean);

                this.app.setCardSelection(selectedItemIds);
                if (this.app.detailsPanelManager && this.app.detailsPanelManager.isOpen()) {
                    this.app.detailsPanelManager.refreshForSelection();
                }
            }

            container.classList.remove('is-lasso-selecting');
            box.remove();
            this.lassoSelection = null;
        });
    }

    initializeFilterSetupDialog() {
        const dialog = document.getElementById('filter-setup-dialog');
        if (!dialog) {
            return;
        }

        const closeBtn = dialog.querySelector('.modal-close');
        const cancelBtn = document.getElementById('filter-setup-cancel-btn');
        const form = document.getElementById('filter-setup-form');

        const closeDialog = () => {
            dialog.classList.add('hidden');
            this.filterSetupContext = null;
        };

        if (closeBtn) {
            closeBtn.addEventListener('click', closeDialog);
        }

        if (cancelBtn) {
            cancelBtn.addEventListener('click', closeDialog);
        }

        dialog.addEventListener('click', (event) => {
            if (event.target === dialog) {
                closeDialog();
            }
        });

        if (form) {
            form.addEventListener('submit', (event) => {
                event.preventDefault();
                this.applyFilterSetupSelection();
            });
        }

        this.filterSetupDialog = dialog;
    }
    
    setupTooltip() {
        this.previewTooltip = document.getElementById('tooltip');
        if (!this.previewTooltip) {
            console.warn('Tooltip element not found');
            return;
        }

        if (!this.editorTab) {
            this.editorOverlay = document.createElement('div');
            this.editorOverlay.id = 'edittab-overlay';
            this.editorOverlay.className = 'edittab-overlay hidden';

            this.editorTab = this.previewTooltip.cloneNode(true);
            this.editorTab.id = 'edittab';
            this.editorTab.classList.add('edittab');
            this.editorTab.classList.add('hidden');
            this.editorTab.classList.remove('tooltip-preview-mode', 'tooltip-edit-mode', 'tooltip-create-mode');

            this.editorOverlay.appendChild(this.editorTab);
            document.body.appendChild(this.editorOverlay);
        }

        this.tooltip = this.previewTooltip;
    }

    getSurfaceForMode(mode = 'preview') {
        return mode === 'edit' || mode === 'create'
            ? this.editorTab
            : this.previewTooltip;
    }

    logSurfaceEvent(surfaceType, action, details = {}) {
        console.log(`[SURFACE] ${surfaceType} ${action}`, {
            activeMode: this.activeTooltipMode,
            activeItemId: this.activeTooltipItemId,
            pendingEditorActivation: this.pendingEditorActivation,
            ...details
        });
    }
    
    // ===== CLICK EDITOR FUNCTIONALITY =====

    setupTooltipDismissal() {
        document.addEventListener('mousedown', (event) => {
            if (!this.isEditorTabActive()) {
                return;
            }

            if (event.target.closest('.tooltip')) {
                return;
            }

            event.preventDefault();
            event.stopPropagation();
            this.focusActiveEditorTab();
        }, true);

        document.addEventListener('click', (event) => {
            if (!this.tooltip || this.tooltip.classList.contains('hidden')) {
                return;
            }

            if (event.target.closest('.tooltip') || event.target.closest('.card')) {
                return;
            }

            if (this.activeTooltipMode === 'preview') {
                this.hideTooltip('preview');
                return;
            }

            if (this.isEditorTabActive()) {
                this.focusActiveEditorTab();
            }
        });

        document.addEventListener('focusin', (event) => {
            if (!this.isEditorTabActive()) {
                return;
            }

            if (event.target.closest('.tooltip')) {
                this.lastFocusedEditorElement = event.target;
                return;
            }

            this.focusActiveEditorTab();
        });

        document.addEventListener('keydown', (event) => {
            if (event.key === 'Escape') {
                this.hideTooltip();
            }
        });
    }

    setupCreateItemTrigger() {
        // Create New Item button removed in V3.
        // Item creation is available by clicking empty space inside a grid cell.
    }

    showTooltip(event, item, mode = 'preview') {
        const surface = this.getSurfaceForMode(mode);
        if (!surface) return;

        this.tooltip = surface;

        if ((this.activeTooltipMode === 'edit' || this.activeTooltipMode === 'create') && mode === 'preview') {
            return;
        }

        if (
            mode === 'edit' &&
            this.activeTooltipItemId === item.id &&
            this.activeTooltipMode === 'edit' &&
            !this.tooltip.classList.contains('hidden')
        ) {
            this.hideTooltip();
            return;
        }

        if (mode === 'edit') {
            this.previewTooltip.classList.add('hidden');
            this.editorOverlay.classList.remove('hidden');
            this.populateEditableTooltipContent(item);
            this.tooltip.classList.add('tooltip-edit-mode');
            this.tooltip.classList.remove('tooltip-preview-mode', 'tooltip-create-mode');
        } else {
            this.populatePreviewTooltipContent(item);
            this.tooltip.classList.add('tooltip-preview-mode');
            this.tooltip.classList.remove('tooltip-edit-mode', 'tooltip-create-mode');
        }

        this.activeTooltipItemId = item.id;
        this.activeTooltipMode = mode;
        this.positionTooltip(event, mode);
        this.tooltip.classList.remove('hidden');
        this.pendingEditorActivation = false;

        const firstInput = mode === 'edit'
            ? this.tooltip.querySelector('input, textarea, select')
            : null;
        if (firstInput) {
            this.lastFocusedEditorElement = firstInput;
            requestAnimationFrame(() => firstInput.focus());
        }

        this.logSurfaceEvent(mode === 'preview' ? 'tooltip' : 'edittab', 'show', {
            itemId: item ? item.id : null,
            mode,
            x: event ? event.clientX : null,
            y: event ? event.clientY : null
        });
    }

    showCreateTooltip(event, initialValues = {}) {
        const surface = this.getSurfaceForMode('create');
        if (!surface) return;

        this.tooltip = surface;

        const draftItem = this.buildTooltipDraftItem(initialValues);
        this.populateCreateTooltipContent(draftItem);
        this.tooltip.classList.add('tooltip-create-mode');
        this.tooltip.classList.remove('tooltip-edit-mode', 'tooltip-preview-mode');
        this.previewTooltip.classList.add('hidden');
        this.editorOverlay.classList.remove('hidden');

        this.activeTooltipItemId = null;
        this.activeTooltipMode = 'create';
        this.positionTooltip(event, 'create');
        this.tooltip.classList.remove('hidden');
        this.pendingEditorActivation = false;

        const firstInput = this.tooltip.querySelector('input, textarea, select');
        if (firstInput) {
            this.lastFocusedEditorElement = firstInput;
            requestAnimationFrame(() => firstInput.focus());
        }

        this.logSurfaceEvent('edittab', 'show', {
            itemId: null,
            mode: 'create',
            x: event ? event.clientX : null,
            y: event ? event.clientY : null
        });
    }

    hideTooltip(mode = null) {
        if (!this.previewTooltip || !this.editorTab) return;

        if (mode === 'preview') {
            if (this.activeTooltipMode && this.activeTooltipMode !== 'preview') {
                return;
            }

            this.previewTooltip.classList.add('hidden');
            this.previewTooltip.classList.remove('tooltip-preview-mode');
            this.logSurfaceEvent('tooltip', 'hide', { requestedMode: mode });

            if (this.activeTooltipMode === 'preview') {
                this.activeTooltipItemId = null;
                this.activeTooltipMode = null;
            }
            return;
        }

        this.editorTab.classList.add('hidden');
        this.editorTab.classList.remove('tooltip-edit-mode', 'tooltip-create-mode');
        if (this.editorOverlay) {
            this.editorOverlay.classList.add('hidden');
        }
    this.logSurfaceEvent('edittab', 'hide', { requestedMode: mode || 'editor' });
        this.lastFocusedEditorElement = null;
        this.pendingEditorActivation = false;

        if (this.activeTooltipMode === 'edit' || this.activeTooltipMode === 'create') {
            this.activeTooltipItemId = null;
            this.activeTooltipMode = null;
        }
    }

    isEditorTabActive() {
        return !!this.editorTab &&
            !this.editorTab.classList.contains('hidden') &&
            (this.activeTooltipMode === 'edit' || this.activeTooltipMode === 'create');
    }

    focusActiveEditorTab() {
        if (!this.isEditorTabActive()) {
            return;
        }

        const fallbackTarget = this.editorTab.querySelector('input, textarea, select, button');
        const target = this.lastFocusedEditorElement && this.editorTab.contains(this.lastFocusedEditorElement)
            ? this.lastFocusedEditorElement
            : fallbackTarget;

        if (target && typeof target.focus === 'function') {
            requestAnimationFrame(() => target.focus());
        }
    }

    beginEditorActivation() {
        this.pendingEditorActivation = true;
        this.logSurfaceEvent('edittab', 'pending-open');
    }

    cancelPendingEditorActivation() {
        if (this.activeTooltipMode !== 'edit' && this.activeTooltipMode !== 'create') {
            this.pendingEditorActivation = false;
            this.logSurfaceEvent('edittab', 'pending-cancel');
        }
    }

    isEditorActivationPending() {
        return this.pendingEditorActivation;
    }

    populatePreviewTooltipContent(item) {
        const content = this.tooltip.querySelector('.tooltip-content');
        content.innerHTML = '';

        const header = document.createElement('div');
        header.className = 'tooltip-header';

        const title = document.createElement('div');
        title.className = 'tooltip-title';
        title.textContent = item.title || item.id || 'Item';

        header.appendChild(title);
        content.appendChild(header);

        const fields = this.getItemTooltipFields(item);
        fields.forEach((fieldName) => {
            const fieldType = this.app.fieldTypes.get(fieldName);
            const value = this.app.getFieldValue(item, fieldName);

            if (this.shouldSkipField(fieldName, value)) {
                return;
            }

            const fieldDiv = document.createElement('div');
            fieldDiv.className = 'tooltip-field';

            const label = document.createElement('div');
            label.className = 'tooltip-label';
            label.textContent = this.formatFieldLabel(fieldName);

            const valueDiv = document.createElement('div');
            valueDiv.className = 'tooltip-value';
            this.populateTooltipFieldValue(valueDiv, value, fieldType, fieldName);

            fieldDiv.appendChild(label);
            fieldDiv.appendChild(valueDiv);
            content.appendChild(fieldDiv);
        });

        // Show group memberships
        const itemGroups = this.app.getItemGroups(item);
        if (itemGroups.length > 0) {
            const groupDiv = document.createElement('div');
            groupDiv.className = 'tooltip-field';

            const groupLabel = document.createElement('div');
            groupLabel.className = 'tooltip-label';
            groupLabel.textContent = 'Groups';

            const groupValue = document.createElement('div');
            groupValue.className = 'tooltip-value tooltip-groups';
            itemGroups.forEach(g => {
                const chip = document.createElement('span');
                chip.className = 'tooltip-group-chip';

                const swatch = document.createElement('span');
                swatch.className = 'tooltip-group-swatch';
                swatch.style.backgroundColor = g.color;

                const text = document.createTextNode(g.name);
                chip.appendChild(swatch);
                chip.appendChild(text);
                groupValue.appendChild(chip);
            });

            groupDiv.appendChild(groupLabel);
            groupDiv.appendChild(groupValue);
            content.appendChild(groupDiv);
        }
    }

    populateEditableTooltipContent(item) {
        const content = this.tooltip.querySelector('.tooltip-content');
        content.innerHTML = '';

        const form = document.createElement('form');
        form.className = 'tooltip-form';
        form.addEventListener('submit', (submitEvent) => {
            submitEvent.preventDefault();
            this.saveTooltipEdit(item, form);
        });

        const header = document.createElement('div');
        header.className = 'tooltip-header';

        const title = document.createElement('div');
        title.className = 'tooltip-title';
        title.textContent = item.title || item.id || 'Item';

        header.appendChild(title);
        form.appendChild(header);

        const fields = this.getTooltipEditableFields();

        fields.forEach(fieldName => {
            const fieldType = this.app.fieldTypes.get(fieldName);
            const value = this.app.getFieldValue(item, fieldName);

            const fieldDiv = document.createElement('div');
            fieldDiv.className = 'tooltip-field tooltip-field-editable';

            const label = document.createElement('label');
            label.className = 'tooltip-label';
            label.textContent = this.formatFieldLabel(fieldName);

            const input = this.createTooltipInput(fieldName, fieldType, value);
            label.htmlFor = input.id;

            fieldDiv.appendChild(label);
            fieldDiv.appendChild(input);
            form.appendChild(fieldDiv);
        });

        // Show group memberships as removable tag-like chips
        const allGroups = this.app.getGroups();
        if (allGroups.length > 0) {
            const itemId = this.app.getItemIdentity(item) || item.id;
            const itemGroups = this.app.getItemGroups(item);
            const groupDiv = document.createElement('div');
            groupDiv.className = 'tooltip-field';

            const groupLabel = document.createElement('div');
            groupLabel.className = 'tooltip-label';
            groupLabel.textContent = 'Groups';

            const groupValue = document.createElement('div');
            groupValue.className = 'tooltip-value tooltip-groups tooltip-groups-editable';

            const renderGroupChips = () => {
                groupValue.innerHTML = '';
                const currentGroups = this.app.getItemGroups(item);
                if (currentGroups.length === 0) {
                    const empty = document.createElement('span');
                    empty.className = 'tooltip-groups-empty';
                    empty.textContent = '(none)';
                    groupValue.appendChild(empty);
                    return;
                }
                currentGroups.forEach(g => {
                    const isManual = Array.isArray(g.manualMembers) && g.manualMembers.includes(itemId);
                    const chip = document.createElement('span');
                    chip.className = 'tooltip-group-tag';
                    chip.style.backgroundColor = g.color;
                    chip.textContent = g.name;
                    chip.title = isManual ? 'Click to remove from group' : 'Rule-based (cannot remove)';
                    if (isManual) {
                        chip.classList.add('is-removable');
                        chip.addEventListener('click', (e) => {
                            e.preventDefault();
                            e.stopPropagation();
                            this.app.removeItemFromGroup(g.id, itemId);
                            renderGroupChips();
                        });
                    } else {
                        chip.classList.add('is-rule-based');
                    }
                    groupValue.appendChild(chip);
                });
            };

            renderGroupChips();
            groupDiv.appendChild(groupLabel);
            groupDiv.appendChild(groupValue);
            form.appendChild(groupDiv);
        }

        const actions = document.createElement('div');
        actions.className = 'tooltip-actions';

        const relationsButton = document.createElement('button');
        relationsButton.type = 'button';
        relationsButton.className = 'btn btn-secondary tooltip-action-btn';
        relationsButton.textContent = '🔗 Relations';
        relationsButton.addEventListener('click', () => {
            if (this.app.relationUIManager) {
                this.app.relationUIManager.openPanel(item, null);
            }
        });

        const cancelButton = document.createElement('button');
        cancelButton.type = 'button';
        cancelButton.className = 'btn btn-secondary tooltip-action-btn';
        cancelButton.textContent = 'Cancel';
        cancelButton.addEventListener('click', () => this.hideTooltip());

        const saveButton = document.createElement('button');
        saveButton.type = 'submit';
        saveButton.className = 'btn btn-primary tooltip-action-btn';
        saveButton.textContent = 'Save';

        actions.appendChild(relationsButton);
        actions.appendChild(cancelButton);
        actions.appendChild(saveButton);
        form.appendChild(actions);

        content.appendChild(form);
    }

    populateCreateTooltipContent(draftItem) {
        const content = this.tooltip.querySelector('.tooltip-content');
        content.innerHTML = '';

        const form = document.createElement('form');
        form.className = 'tooltip-form';
        form.addEventListener('submit', (submitEvent) => {
            submitEvent.preventDefault();
            this.saveCreateTooltip(form);
        });

        const header = document.createElement('div');
        header.className = 'tooltip-header';

        const title = document.createElement('div');
        title.className = 'tooltip-title';
        title.textContent = 'Create New Item';

        header.appendChild(title);
        form.appendChild(header);

        this.getTooltipEditableFields().forEach((fieldName) => {
            const fieldType = this.app.fieldTypes.get(fieldName);
            const value = Object.prototype.hasOwnProperty.call(draftItem, fieldName)
                ? draftItem[fieldName]
                : null;

            const fieldDiv = document.createElement('div');
            fieldDiv.className = 'tooltip-field tooltip-field-editable';

            const label = document.createElement('label');
            label.className = 'tooltip-label';
            label.textContent = this.formatFieldLabel(fieldName);

            const input = this.createTooltipInput(fieldName, fieldType, value);
            label.htmlFor = input.id;

            fieldDiv.appendChild(label);
            fieldDiv.appendChild(input);
            form.appendChild(fieldDiv);
        });

        const actions = document.createElement('div');
        actions.className = 'tooltip-actions';

        const cancelButton = document.createElement('button');
        cancelButton.type = 'button';
        cancelButton.className = 'btn btn-secondary tooltip-action-btn';
        cancelButton.textContent = 'Cancel';
        cancelButton.addEventListener('click', () => this.hideTooltip());

        const saveButton = document.createElement('button');
        saveButton.type = 'submit';
        saveButton.className = 'btn btn-primary tooltip-action-btn';
        saveButton.textContent = 'Create';

        actions.appendChild(cancelButton);
        actions.appendChild(saveButton);
        form.appendChild(actions);

        content.appendChild(form);
    }

    getTooltipEditableFields() {
        const allFields = Array.from(this.app.fieldTypes.keys());
        const isEditable = (fn) => {
            const sf = this.app.schemaFields[fn];
            return !(sf && (sf.visible === false || sf.editable === false));
        };
        const prioritizedFields = ['id', 'title'];
        const prioritizedSet = new Set(prioritizedFields);
        const orderedFields = prioritizedFields.filter((fieldName) => allFields.includes(fieldName) && isEditable(fieldName));
        const remainingFields = allFields
            .filter((fieldName) => !prioritizedSet.has(fieldName))
            .filter((fieldName) => this.app.fieldTypes.get(fieldName) !== 'structured')
            .filter(isEditable)
            .sort();

        return [...orderedFields, ...remainingFields];
    }

    getItemTooltipFields(item) {
        const knownFields = Array.from(this.app.fieldTypes.keys());
        const itemFields = item && typeof item === 'object'
            ? Object.keys(item)
            : [];
        const allFields = Array.from(new Set([...knownFields, ...itemFields]));
        const isVisible = (fn) => {
            const sf = this.app.schemaFields[fn];
            return !(sf && sf.visible === false);
        };
        const prioritizedFields = ['id', 'title'];
        const prioritizedSet = new Set(prioritizedFields);
        const orderedFields = prioritizedFields.filter((fieldName) => allFields.includes(fieldName) && isVisible(fieldName));
        const remainingFields = allFields
            .filter((fieldName) => !prioritizedSet.has(fieldName))
            .filter(isVisible)
            .sort();

        return [...orderedFields, ...remainingFields];
    }

    buildTooltipDraftItem(initialValues = {}) {
        const draftItem = {};

        this.getTooltipEditableFields().forEach((fieldName) => {
            const fieldType = this.app.fieldTypes.get(fieldName);

            if (Object.prototype.hasOwnProperty.call(initialValues, fieldName)) {
                draftItem[fieldName] = initialValues[fieldName];
                return;
            }

            if (fieldType === 'multi-value') {
                draftItem[fieldName] = [];
                return;
            }

            draftItem[fieldName] = null;
        });

        return draftItem;
    }

    createTooltipInput(fieldName, fieldType, value) {
        let input;

        if (fieldType === 'multi-value') {
            input = document.createElement('textarea');
            input.rows = fieldName === 'tags' ? 2 : 3;
            input.value = Array.isArray(value) ? value.join('\n') : (value || '');
        } else if (fieldType === 'structured') {
            input = document.createElement('textarea');
            input.rows = 5;
            input.className = 'tooltip-input tooltip-json-input';
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

        input.classList.add('tooltip-input');
        input.name = fieldName;
        input.id = `tooltip-field-${fieldName}`;

        return input;
    }

    saveTooltipEdit(item, form) {
        const formData = this.collectTooltipFormData(form);
        if (!formData) {
            return;
        }

        const validation = this.validateFormData(formData);

        if (!validation.isValid) {
            const firstError = validation.errors[0];
            if (firstError) {
                this.app.showNotification(firstError.message, 'error');
                const input = form.querySelector(`[name="${firstError.field}"]`);
                if (input) {
                    input.focus();
                }
            }
            return;
        }

        const itemId = this.app.getItemIdentity(item) || item.id;
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
        this.hideTooltip();
        this.app.showNotification(`Updated ${formData.id || itemId || 'item'}`, 'success');
    }

    saveCreateTooltip(form) {
        const formData = this.collectTooltipFormData(form);
        if (!formData) {
            return;
        }

        const validation = this.validateFormData(formData, true);

        if (!validation.isValid) {
            const firstError = validation.errors[0];
            if (firstError) {
                this.app.showNotification(firstError.message, 'error');
                const input = form.querySelector(`[name="${firstError.field}"]`);
                if (input) {
                    input.focus();
                }
            }
            return;
        }

        if (formData.id && this.app.dataset.some((item) => item.id === formData.id)) {
            this.app.showNotification('ID already exists', 'error');
            const idInput = form.querySelector('[name="id"]');
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
        this.hideTooltip();
        this.app.showNotification(`Created ${newItemId}`, 'success');
    }

    collectTooltipFormData(form) {
        const formData = {};
        const fields = form.querySelectorAll('input[name], textarea[name], select[name]');

        for (const field of fields) {
            const fieldName = field.name;
            const fieldType = this.app.fieldTypes.get(fieldName);
            const parsedValue = this.parseTooltipInputValue(fieldName, fieldType, field.value);

            if (parsedValue.error) {
                this.app.showNotification(parsedValue.error, 'error');
                field.focus();
                return null;
            }

            formData[fieldName] = parsedValue.value;
        }

        return formData;
    }

    parseTooltipInputValue(fieldName, fieldType, rawValue) {
        if (fieldType === 'multi-value') {
            return {
                value: rawValue
                    .split(/[,\n]+/)
                    .map((entry) => entry.trim())
                    .filter((entry) => entry.length > 0)
            };
        }

        if (fieldType === 'structured') {
            if (!rawValue.trim()) {
                return { value: null };
            }

            try {
                return { value: JSON.parse(rawValue) };
            } catch (error) {
                return { error: `Invalid JSON in ${fieldName}: ${error.message}` };
            }
        }

        return { value: rawValue || null };
    }

    getInputType(fieldName) {
        if (typeof this.app.getFormInputType === 'function') {
            return this.app.getFormInputType(fieldName);
        }

        return 'text';
    }

    validateFormData(formData, isCreate = false) {
        if (typeof this.app.validateItemFormData === 'function') {
            return this.app.validateItemFormData(formData, isCreate);
        }

        return { isValid: true, errors: [] };
    }

    openInlineValueEditor(card, item, fieldName, element) {
        if (!card || !element || !fieldName || element.classList.contains('is-editing')) {
            return;
        }

        this.hideTooltip();

        const fieldType = this.app.fieldTypes.get(fieldName);
        const currentValue = this.app.getFieldValue(item, fieldName);
        const originalDisplayValue = element.textContent;

        card.querySelectorAll('.card-bottom-right.is-editing').forEach((editingElement) => {
            editingElement.classList.remove('is-editing');
            editingElement.textContent = editingElement.dataset.originalDisplayValue || editingElement.textContent;
            delete editingElement.dataset.originalDisplayValue;
        });

        element.dataset.originalDisplayValue = originalDisplayValue;
        element.classList.add('is-editing');
        element.textContent = '';

        const editor = fieldType === 'multi-value' || fieldType === 'structured'
            ? document.createElement('textarea')
            : document.createElement('input');

        editor.className = 'card-inline-editor';
        editor.name = fieldName;

        if (editor.tagName === 'TEXTAREA') {
            editor.rows = fieldType === 'structured' ? 4 : 2;
            editor.value = fieldType === 'multi-value' && Array.isArray(currentValue)
                ? currentValue.join('\n')
                : fieldType === 'structured' && currentValue && typeof currentValue === 'object'
                    ? JSON.stringify(currentValue, null, 2)
                    : (currentValue || '');
        } else {
            editor.type = this.getInputType(fieldName);
            editor.value = typeof this.app.getEditorInputValue === 'function'
                ? this.app.getEditorInputValue(fieldName, currentValue)
                : (currentValue || '');
        }

        const finishEditing = (shouldSave) => {
            if (!element.classList.contains('is-editing')) {
                return;
            }

            if (!shouldSave) {
                element.textContent = originalDisplayValue;
                element.classList.remove('is-editing');
                delete element.dataset.originalDisplayValue;
                return;
            }

            const parsedValue = this.parseInlineEditorValue(editor.value, fieldType, fieldName);
            if (parsedValue.error) {
                this.app.showNotification(parsedValue.error, 'error');
                editor.focus();
                return;
            }

            const itemId = this.app.getItemIdentity(item) || item.id;
            const previousDistinctValues = this.app.snapshotDistinctValues();
            const success = this.app.updateItem(itemId, { [fieldName]: parsedValue.value });
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
            this.app.showNotification(`Updated ${fieldName}`, 'success');
        };

        editor.addEventListener('click', (event) => {
            event.stopPropagation();
        });

        editor.addEventListener('keydown', (event) => {
            if (event.key === 'Escape') {
                event.preventDefault();
                finishEditing(false);
                return;
            }

            if (event.key === 'Enter' && !event.shiftKey && editor.tagName !== 'TEXTAREA') {
                event.preventDefault();
                finishEditing(true);
            }

            if ((event.ctrlKey || event.metaKey) && event.key === 'Enter' && editor.tagName === 'TEXTAREA') {
                event.preventDefault();
                finishEditing(true);
            }
        });

        editor.addEventListener('blur', () => {
            finishEditing(true);
        });

        element.appendChild(editor);
        requestAnimationFrame(() => {
            editor.focus();
            if (typeof editor.select === 'function') {
                editor.select();
            }
        });
    }

    parseInlineEditorValue(rawValue, fieldType, fieldName) {
        if (fieldType === 'multi-value') {
            return {
                value: rawValue
                    .split(/[,\n]+/)
                    .map((entry) => entry.trim())
                    .filter((entry) => entry.length > 0)
            };
        }

        if (fieldType === 'structured') {
            if (!rawValue.trim()) {
                return { value: null };
            }

            try {
                return { value: JSON.parse(rawValue) };
            } catch (error) {
                return { error: `Invalid JSON in ${fieldName}: ${error.message}` };
            }
        }

        return { value: rawValue || null };
    }
    
    shouldSkipField(fieldName, value) {
        // Schema-driven: skip fields marked as not visible
        const sf = this.app.schemaFields[fieldName];
        if (sf && sf.visible === false) return true;

        // Always show these important fields even if empty
        const normalizedFieldName = String(fieldName || '').toLowerCase();
        const alwaysShow = ['id', 'title', 'status', 'priority'];
        if (alwaysShow.includes(fieldName) || normalizedFieldName.includes('date')) return false;
        
        // Skip structured fields entirely in preview tooltip
        const fieldType = this.app.fieldTypes.get(fieldName);
        if (fieldType === 'structured') {
            return true;
        }
        
        // Skip empty values
        return value === '' || value === null || value === undefined ||
               (Array.isArray(value) && value.length === 0);
    }
    
    formatFieldLabel(fieldName) {
        return fieldName.charAt(0).toUpperCase() + 
               fieldName.slice(1).replace(/([A-Z])/g, ' $1');
    }
    
    populateTooltipFieldValue(valueDiv, value, fieldType, fieldName) {
        if (fieldType === 'multi-value' && Array.isArray(value)) {
            // Render tags or list items
            if (this.app.isTagFieldName(fieldName)) {
                this.renderTooltipTags(valueDiv, value);
            } else {
                valueDiv.textContent = value.join(', ');
            }
        } else if (this.app.isTagFieldName(fieldName)) {
            this.renderTooltipTags(valueDiv, this.app.normalizeTagValues(value));
        } else if (fieldType === 'structured' && typeof value === 'object' && value !== null) {
            // Render structured data in a readable format
            this.renderTooltipStructured(valueDiv, value);
        } else {
            // Render scalar value
            valueDiv.textContent = this.app.getDisplayValue({ [fieldName]: value }, fieldName);
        }
    }
    
    renderTooltipTags(container, tags) {
        container.innerHTML = '';
        tags.forEach(tagName => {
            const tagConfig = this.app.getTagConfig(tagName);
            const tagSpan = document.createElement('span');
            tagSpan.className = 'tooltip-tag';
            tagSpan.style.cssText = `
                display: inline-block;
                padding: 2px 6px;
                margin: 2px;
                border-radius: 8px;
                font-size: 11px;
                font-weight: 500;
                color: white;
                background-color: ${tagConfig.color};
            `;
            tagSpan.textContent = tagConfig.label;
            container.appendChild(tagSpan);
        });
    }
    
    renderTooltipStructured(container, obj) {
        container.innerHTML = '';
        const pre = document.createElement('pre');
        pre.style.cssText = `
            font-size: 11px;
            margin: 0;
            white-space: pre-wrap;
            max-height: 100px;
            overflow-y: auto;
            background: #f5f5f5;
            padding: 4px;
            border-radius: 2px;
        `;
        pre.textContent = JSON.stringify(obj, null, 2);
        container.appendChild(pre);
    }
    
    getSurfaceDimensions(mode = 'preview') {
        if (mode === 'edit' || mode === 'create') {
            return { width: 400, height: 400 };
        }

        return { width: 200, height: 400 };
    }

    positionTooltip(event, mode = 'preview') {
        const tooltipRect = this.getSurfaceDimensions(mode);
        const viewportWidth = window.innerWidth;
        const viewportHeight = window.innerHeight;
        
        let left = event.clientX + 10;
        let top = event.clientY + 10;
        
        // Adjust if tooltip would go off screen
        if (left + tooltipRect.width > viewportWidth) {
            left = event.clientX - tooltipRect.width - 10;
        }
        
        if (top + tooltipRect.height > viewportHeight) {
            top = event.clientY - tooltipRect.height - 10;
        }
        
        // Ensure tooltip doesn't go off the left or top edge
        left = Math.max(10, left);
        top = Math.max(10, top);
        
        this.tooltip.style.left = left + 'px';
        this.tooltip.style.top = top + 'px';
    }
    
    // ===== DRAG & DROP FUNCTIONALITY =====
    
    handleCardDrop(itemId, newRowValue, newColValue) {
        const item = typeof this.app.getEffectiveItemById === 'function'
            ? this.app.getEffectiveItemById(itemId)
            : this.app.dataset.find((entry) => entry.id === itemId);
        if (!item) {
            console.error('Item not found for drop:', itemId);
            return;
        }
        
        const xField = this.app.currentAxisSelections.x;
        const yField = this.app.currentAxisSelections.y;
        
        if (!xField || !yField) {
            this.app.showNotification('Cannot move: no axis fields selected', 'error');
            return;
        }

        const preservedHeaderOrdering = this.app.advancedFeaturesManager &&
            typeof this.app.advancedFeaturesManager.getRenderedHeaderOrdering === 'function'
            ? this.app.advancedFeaturesManager.getRenderedHeaderOrdering()
            : null;
        
        // Schema-driven: block drag-drop onto read-only axis fields
        const xSchema = this.app.schemaFields[xField];
        const ySchema = this.app.schemaFields[yField];
        if ((xSchema && xSchema.editable === false) || (ySchema && ySchema.editable === false)) {
            this.app.showNotification('Cannot move: axis field is read-only', 'warning');
            return;
        }

        const blockedDerivedAxes = [xField, yField].filter((fieldName) => {
            const kind = this.app.getFieldKind(fieldName);
            return kind === 'derived' || kind === 'relationship';
        });

        if (blockedDerivedAxes.length > 0) {
            const axisList = blockedDerivedAxes.join(', ');
            this.app.showNotification(`Cannot move: ${axisList} is a derived field`, 'warning');
            return;
        }

        const updates = {};
        let changesMade = false;

        const getNormalizedDropValue = (fieldName, rawValue) => {
            if (rawValue === '') {
                return null;
            }

            if (this.app.getFormInputType(fieldName) === 'date') {
                const normalizedValue = String(rawValue).match(/^(\d{4}-\d{2}-\d{2})/);
                return normalizedValue ? normalizedValue[1] : rawValue;
            }

            return rawValue;
        };
        
        // Update x-axis field if value changed
        const currentXValue = this.app.normalizeFieldValueForComparison(xField, this.app.getFieldValue(item, xField));
        const nextXValue = this.app.normalizeFieldValueForComparison(xField, getNormalizedDropValue(xField, newColValue));
        if (!this.app.areFieldValuesEquivalent(xField, currentXValue, nextXValue)) {
            updates[xField] = getNormalizedDropValue(xField, newColValue);
            changesMade = true;
        }
        
        // Update y-axis field if value changed
        const currentYValue = this.app.normalizeFieldValueForComparison(yField, this.app.getFieldValue(item, yField));
        const nextYValue = this.app.normalizeFieldValueForComparison(yField, getNormalizedDropValue(yField, newRowValue));
        if (!this.app.areFieldValuesEquivalent(yField, currentYValue, nextYValue)) {
            updates[yField] = getNormalizedDropValue(yField, newRowValue);
            changesMade = true;
        }
        
        if (changesMade) {
            const previousDistinctValues = this.app.snapshotDistinctValues();
            const success = this.app.updateItem(itemId, updates, { render: false });
            if (success) {
                this.app.showNotification(`Moved ${item.id || 'item'} successfully`, 'success');
                
                // Update view configuration and re-analyze data
                this.app.analyzeFields();  // Re-analyze in case new values were added
                this.app.reconcileFiltersWithDistinctValues(previousDistinctValues);

                if (
                    preservedHeaderOrdering &&
                    this.app.advancedFeaturesManager &&
                    typeof this.app.advancedFeaturesManager.setHeaderOrdering === 'function'
                ) {
                    this.app.advancedFeaturesManager.setHeaderOrdering(preservedHeaderOrdering);
                    this.app.advancedFeaturesManager.setPendingHeaderOrderReason('row', 'card move preserved header order', {
                        ordering: [...preservedHeaderOrdering.rows],
                        itemId,
                        newRowValue,
                        newColValue
                    });
                    this.app.advancedFeaturesManager.setPendingHeaderOrderReason('column', 'card move preserved header order', {
                        ordering: [...preservedHeaderOrdering.columns],
                        itemId,
                        newRowValue,
                        newColValue
                    });
                }

                this.app.updateViewConfiguration();
                
                // Re-render affected components
                this.app.renderSlicers();  // Update slicers with new values
                this.app.renderGrid();
            } else {
                this.app.showNotification('Failed to update item', 'error');
            }
        } else {
            this.app.showNotification('Item already in target location', 'info');
        }
    }

    addTagToItem(itemId, tagName) {
        const itemIds = Array.from(new Set(
            (Array.isArray(itemId) ? itemId : [itemId])
                .map((value) => value === undefined || value === null ? null : String(value))
                .filter(Boolean)
        ));
        if (itemIds.length === 0) {
            this.app.showNotification('Failed to assign tag: item not found', 'error');
            return false;
        }

        const tagFieldName = this.app.getTagFieldName();
        if (!tagFieldName) {
            this.app.showNotification('No tags field available in dataset', 'error');
            return false;
        }

        let changedCount = 0;

        itemIds.forEach((targetItemId) => {
            const item = typeof this.app.getEffectiveItemById === 'function'
                ? this.app.getEffectiveItemById(targetItemId)
                : this.app.dataset.find((entry) => entry.id === targetItemId);

            if (!item) {
                return;
            }

            const currentTags = this.app.getItemTags(item);
            if (currentTags.includes(tagName)) {
                return;
            }

            const nextTags = [...currentTags, tagName];
            const success = this.app.updateItem(targetItemId, { [tagFieldName]: nextTags }, { render: false });
            if (success) {
                changedCount += 1;
            }
        });

        if (changedCount === 0) {
            this.app.showNotification(itemIds.length === 1 ? `Tag ${tagName} is already assigned` : `Tag ${tagName} was already assigned to all selected items`, 'info');
            return true;
        }

        if (this.app.controlPanel) {
            this.app.controlPanel.draggedTagName = null;
        }

        document.body.classList.remove('tag-drag-mode');
        this.app.updateViewConfiguration();
        this.app.renderSlicers();
        this.app.renderTags();
        this.app.renderGrid();
        this.app.showNotification(`Assigned tag ${tagName} to ${changedCount} item${changedCount === 1 ? '' : 's'}`, 'success');
        return true;
    }

    removeTagFromItem(itemId, tagName, options = {}) {
        const { suppressNotification = false } = options;
        const item = typeof this.app.getEffectiveItemById === 'function'
            ? this.app.getEffectiveItemById(itemId)
            : this.app.dataset.find((entry) => entry.id === itemId);

        if (!item) {
            if (!suppressNotification) {
                this.app.showNotification('Failed to remove tag: item not found', 'error');
            }
            return false;
        }

        const tagFieldName = this.app.getTagFieldName();
        if (!tagFieldName) {
            if (!suppressNotification) {
                this.app.showNotification('No tags field available in dataset', 'error');
            }
            return false;
        }

        const currentTags = this.app.getItemTags(item);
        const updatedTags = currentTags.filter((existingTag) => existingTag !== tagName);

        if (updatedTags.length === currentTags.length) {
            if (!suppressNotification) {
                this.app.showNotification(`Tag ${tagName} is not assigned`, 'info');
            }
            return false;
        }

        if (typeof this.app.deselectCardTag === 'function') {
            this.app.deselectCardTag(itemId, tagName);
        } else if (
            this.app.selectedCardTag &&
            this.app.selectedCardTag.itemId === itemId &&
            this.app.selectedCardTag.tagName === tagName
        ) {
            this.app.clearSelectedCardTag();
        }

        const success = this.app.updateItem(itemId, { [tagFieldName]: updatedTags });

        if (success) {
            this.app.draggedCardTag = null;
            document.body.classList.remove('tag-drag-mode');
            this.app.updateViewConfiguration();
            this.app.renderSlicers();
            this.app.renderTags();
            this.app.renderGrid();

            if (!suppressNotification) {
                this.app.showNotification(`Removed tag ${tagName}`, 'success');
            }

            return true;
        }

        if (!suppressNotification) {
            this.app.showNotification('Failed to remove tag', 'error');
        }
        return false;
    }
    
    // ===== CARD CLICK =====

    handleCardClick(event, item) {
        // Single click always opens the editor/details panel.
        this.hideTooltip();

        if (this.app.detailsPanelManager) {
            const panel = this.app.detailsPanelManager;
            const mode = panel.isOpen() ? (panel.currentMode || 'editor') : 'editor';
            panel.openForItem(item, mode);
        } else {
            this.showTooltip(event, item, 'edit');
        }
    }

    handleCardDoubleClick(event, item) {
        const clickConfig = this.getEffectiveCardClickConfig();
        const url = this.constructItemUrl(item, clickConfig);

        if (url) {
            const target = clickConfig.openInNewTab === false ? '_self' : '_blank';
            window.open(url, target, 'noopener');
            this.hideTooltip();
            return;
        }

        this.handleCardClick(event, item);
    }

    // ===== CARD CONTEXT MENU =====

    getTableGraphContext(surfaceElement = null) {
        if (!surfaceElement || typeof surfaceElement.closest !== 'function') {
            return null;
        }

        const tableWrapper = surfaceElement.closest('.grid-cell-table');
        return tableWrapper && tableWrapper._tableGraphContext
            ? tableWrapper._tableGraphContext
            : null;
    }

    handleShiftTableContextSelection(event, surfaceElementOrContext = null) {
        const tableGraphContext = surfaceElementOrContext && surfaceElementOrContext.selectionKey
            ? surfaceElementOrContext
            : this.getTableGraphContext(surfaceElementOrContext);

        if (!tableGraphContext) {
            return false;
        }

        event.preventDefault();
        event.stopPropagation();
        this.dismissCardContextMenu();

        const isSelected = this.app.toggleTableGraphSelection(tableGraphContext);
        const selectedCount = this.app.getSelectedTableGraphContexts().length;
        this.app.showNotification(
            isSelected
                ? `Selected table for graphing (${selectedCount})`
                : `Removed table from graph selection (${selectedCount})`,
            'info'
        );
        return true;
    }

    getTableGraphMenuLabel(tableGraphContext) {
        const selectedContexts = this.app.getTableGraphContextsForOpen(tableGraphContext);
        return selectedContexts.length > 1
            ? 'Send Selected to Graph'
            : 'Send to Graph';
    }

    openTableGraphFromContext(tableGraphContext) {
        const sourceContexts = this.app.getTableGraphContextsForOpen(tableGraphContext);

        if (!Array.isArray(sourceContexts) || sourceContexts.length === 0) {
            this.app.showNotification('No table data available for graphing', 'warning');
            return;
        }

        const itemIds = Array.from(new Set(
            sourceContexts.flatMap((context) => Array.isArray(context.itemIds) ? context.itemIds : [])
        ));
        if (itemIds.length === 0) {
            this.app.showNotification('No table data available for graphing', 'warning');
            return;
        }

        const items = itemIds
            .map((itemId) => this.app.getEffectiveItemById(itemId))
            .filter(Boolean);

        if (items.length === 0) {
            this.app.showNotification('No current table items could be resolved for graphing', 'warning');
            return;
        }

        this.app.openTableGraphWindow(items, this.app.getCombinedTableGraphFields(sourceContexts), {
            contextLabel: this.app.buildCombinedTableGraphContextLabel(sourceContexts)
        });
    }

    async saveTableContextToExcel(tableGraphContext) {
        const sourceContexts = this.app.getTableGraphContextsForOpen(tableGraphContext);
        if (!Array.isArray(sourceContexts) || sourceContexts.length === 0) {
            this.app.showNotification('No table data available for Excel export', 'warning');
            return;
        }

        try {
            await this.app.exportTableGraphContextsToExcel(sourceContexts);
        } catch (error) {
            console.error('Error exporting table data to Excel:', error);
            this.app.showNotification('Failed to export table data to Excel: ' + error.message, 'error');
        }
    }

    showTableContextMenu(event, tableGraphContext) {
        this.dismissCardContextMenu();

        const menu = document.createElement('div');
        menu.className = 'card-context-menu';
        menu.id = 'card-context-menu';

        const graphOption = document.createElement('div');
        graphOption.className = 'card-context-menu-item';
        graphOption.textContent = this.getTableGraphMenuLabel(tableGraphContext);
        graphOption.addEventListener('click', (clickEvent) => {
            clickEvent.stopPropagation();
            this.dismissCardContextMenu();
            this.openTableGraphFromContext(tableGraphContext);
        });
        menu.appendChild(graphOption);

        const exportOption = document.createElement('div');
        exportOption.className = 'card-context-menu-item';
        exportOption.textContent = this.app.getTableGraphContextsForOpen(tableGraphContext).length > 1
            ? 'Save Selected to Excel'
            : 'Save to Excel';
        exportOption.addEventListener('click', async (clickEvent) => {
            clickEvent.stopPropagation();
            this.dismissCardContextMenu();
            await this.saveTableContextToExcel(tableGraphContext);
        });
        menu.appendChild(exportOption);

        document.body.appendChild(menu);
        const menuRect = menu.getBoundingClientRect();
        let left = event.clientX;
        let top = event.clientY;
        if (left + menuRect.width > window.innerWidth) {
            left = window.innerWidth - menuRect.width - 4;
        }
        if (top + menuRect.height > window.innerHeight) {
            top = window.innerHeight - menuRect.height - 4;
        }
        menu.style.left = `${left}px`;
        menu.style.top = `${top}px`;

        const dismissHandler = (dismissEvent) => {
            if (!menu.contains(dismissEvent.target)) {
                this.dismissCardContextMenu();
                document.removeEventListener('mousedown', dismissHandler, true);
                document.removeEventListener('keydown', escHandler, true);
            }
        };
        const escHandler = (dismissEvent) => {
            if (dismissEvent.key === 'Escape') {
                this.dismissCardContextMenu();
                document.removeEventListener('mousedown', dismissHandler, true);
                document.removeEventListener('keydown', escHandler, true);
            }
        };

        requestAnimationFrame(() => {
            document.addEventListener('mousedown', dismissHandler, true);
            document.addEventListener('keydown', escHandler, true);
        });
    }

    showCardContextMenu(event, item, surfaceElement = null) {
        this.dismissCardContextMenu();

        const menu = document.createElement('div');
        menu.className = 'card-context-menu';
        menu.id = 'card-context-menu';

        // Option a) Open Editor (in details panel)
        const editOption = document.createElement('div');
        editOption.className = 'card-context-menu-item';
        editOption.textContent = 'Open Editor';
        editOption.addEventListener('click', (e) => {
            e.stopPropagation();
            this.dismissCardContextMenu();
            if (this.app.detailsPanelManager) {
                this.app.detailsPanelManager.openForItem(item, 'editor');
            } else {
                this.showTooltip(event, item, 'edit');
            }
        });
        menu.appendChild(editOption);

        const filterOption = document.createElement('div');
        filterOption.className = 'card-context-menu-item';
        filterOption.textContent = 'Setup Filter';
        filterOption.addEventListener('click', (e) => {
            e.stopPropagation();
            this.dismissCardContextMenu();
            this.openFilterSetupDialog(item, surfaceElement);
        });
        menu.appendChild(filterOption);

        const tableGraphContext = this.getTableGraphContext(surfaceElement);
        if (tableGraphContext) {
            const graphOption = document.createElement('div');
            graphOption.className = 'card-context-menu-item';
            graphOption.textContent = this.getTableGraphMenuLabel(tableGraphContext);
            graphOption.addEventListener('click', (e) => {
                e.stopPropagation();
                this.dismissCardContextMenu();
                this.openTableGraphFromContext(tableGraphContext);
            });
            menu.appendChild(graphOption);

            const exportOption = document.createElement('div');
            exportOption.className = 'card-context-menu-item';
            exportOption.textContent = this.app.getTableGraphContextsForOpen(tableGraphContext).length > 1
                ? 'Save Selected to Excel'
                : 'Save to Excel';
            exportOption.addEventListener('click', async (e) => {
                e.stopPropagation();
                this.dismissCardContextMenu();
                await this.saveTableContextToExcel(tableGraphContext);
            });
            menu.appendChild(exportOption);
        }

        // Option b) Open Link (if configured / available)
        const clickConfig = this.getEffectiveCardClickConfig();
        const url = this.constructItemUrl(item, clickConfig);
        if (url) {
            const linkOption = document.createElement('div');
            linkOption.className = 'card-context-menu-item';
            linkOption.textContent = 'Open Link';
            linkOption.addEventListener('click', (e) => {
                e.stopPropagation();
                this.dismissCardContextMenu();
                const target = clickConfig.openInNewTab === false ? '_self' : '_blank';
                window.open(url, target, 'noopener');
            });
            menu.appendChild(linkOption);
        }

        // Option c) Relationship View (details panel)
        if (this.app.detailsPanelManager) {
            const relViewOption = document.createElement('div');
            relViewOption.className = 'card-context-menu-item';
            relViewOption.textContent = 'Open Relationship View';
            relViewOption.addEventListener('click', (e) => {
                e.stopPropagation();
                this.dismissCardContextMenu();
                if (this.app.relationUIManager) {
                    this.app.relationUIManager.openPanel(item, null);
                } else {
                    this.app.detailsPanelManager.openForItem(item, 'relations');
                }
            });
            menu.appendChild(relViewOption);
        }

        // Option d) Open in Graph View
        if (this.app.graphViewManager) {
            const graphViewOption = document.createElement('div');
            graphViewOption.className = 'card-context-menu-item';
            graphViewOption.textContent = 'Open in Graph View';
            graphViewOption.addEventListener('click', (e) => {
                e.stopPropagation();
                this.dismissCardContextMenu();
                this.app.graphViewManager.openGraphView(item);
            });
            menu.appendChild(graphViewOption);
        }

        // Option f) Enter Focus Mode (relations)
        if (this.app.relationUIManager) {
            const focusOption = document.createElement('div');
            focusOption.className = 'card-context-menu-item';
            focusOption.textContent = 'Enter Focus Mode';
            focusOption.addEventListener('click', (e) => {
                e.stopPropagation();
                this.dismissCardContextMenu();
                const itemId = this.app.getItemIdentity(item) || item.id;
                this.app.relationUIManager.enterFocusMode(itemId);
            });
            menu.appendChild(focusOption);
        }

        // Option g) Markdown View
        if (this.app.detailsPanelManager) {
            const mdOption = document.createElement('div');
            mdOption.className = 'card-context-menu-item';
            mdOption.textContent = 'Markdown View';
            mdOption.addEventListener('click', (e) => {
                e.stopPropagation();
                this.dismissCardContextMenu();
                this.app.detailsPanelManager.openForItem(item, 'markdown');
            });
            menu.appendChild(mdOption);
        }

        // Option h) Show Changes
        {
            const itemId = this.app.getItemIdentity(item) || item.id;
            const itemChanges = this._getChangesForItem(itemId);
            if (itemChanges.length > 0) {
                const changesOption = document.createElement('div');
                changesOption.className = 'card-context-menu-item';
                changesOption.textContent = 'Show Changes';
                changesOption.addEventListener('click', (e) => {
                    e.stopPropagation();
                    this.dismissCardContextMenu();
                    this._showChangesPopup(itemChanges, itemId);
                });
                menu.appendChild(changesOption);
            }
        }

        // Position at click coordinates
        document.body.appendChild(menu);
        const menuRect = menu.getBoundingClientRect();
        let left = event.clientX;
        let top = event.clientY;
        if (left + menuRect.width > window.innerWidth) {
            left = window.innerWidth - menuRect.width - 4;
        }
        if (top + menuRect.height > window.innerHeight) {
            top = window.innerHeight - menuRect.height - 4;
        }
        menu.style.left = `${left}px`;
        menu.style.top = `${top}px`;

        // Dismiss on outside click or Escape
        const dismissHandler = (e) => {
            if (!menu.contains(e.target)) {
                this.dismissCardContextMenu();
                document.removeEventListener('mousedown', dismissHandler, true);
                document.removeEventListener('keydown', escHandler, true);
            }
        };
        const escHandler = (e) => {
            if (e.key === 'Escape') {
                this.dismissCardContextMenu();
                document.removeEventListener('mousedown', dismissHandler, true);
                document.removeEventListener('keydown', escHandler, true);
            }
        };
        // Delay listener install so the current right-click doesn't immediately dismiss
        requestAnimationFrame(() => {
            document.addEventListener('mousedown', dismissHandler, true);
            document.addEventListener('keydown', escHandler, true);
        });
    }

    dismissCardContextMenu() {
        const existing = document.getElementById('card-context-menu');
        if (existing) existing.remove();
    }

    // ── Show Changes helpers ──────────────────────────────

    _getChangesForItem(itemId) {
        const rows = (this.app.pendingChanges && Array.isArray(this.app.pendingChanges.rows))
            ? this.app.pendingChanges.rows
            : [];
        return rows.filter((r) => {
            const tid = r.target && (r.target.itemId || r.target.id);
            return String(tid) === String(itemId);
        });
    }

    _getAllChanges() {
        const rows = (this.app.pendingChanges && Array.isArray(this.app.pendingChanges.rows))
            ? this.app.pendingChanges.rows
            : [];
        return rows;
    }

    _flattenChangeRows(changeRows) {
        const flat = [];
        for (const row of changeRows) {
            const itemId = row.target && (row.target.itemId || row.target.id) || '?';
            const action = row.action || 'update';
            const proposed = row.proposed || {};
            const baseline = row.baseline || {};
            const fields = new Set([...Object.keys(proposed), ...Object.keys(baseline)]);
            for (const field of fields) {
                flat.push({
                    id: itemId,
                    action,
                    field,
                    original: baseline[field] !== undefined ? baseline[field] : '',
                    changed: proposed[field] !== undefined ? proposed[field] : '',
                });
            }
        }
        return flat;
    }

    _showChangesPopup(changeRows, title) {
        const flat = this._flattenChangeRows(changeRows);

        // Build modal overlay
        const overlay = document.createElement('div');
        overlay.className = 'changes-popup-overlay';
        overlay.addEventListener('click', (e) => {
            if (e.target === overlay) overlay.remove();
        });

        const modal = document.createElement('div');
        modal.className = 'changes-popup-modal';

        // Header
        const header = document.createElement('div');
        header.className = 'changes-popup-header';
        const h3 = document.createElement('h3');
        h3.textContent = title ? `Changes — ${title}` : 'All Changes';
        header.appendChild(h3);
        const closeBtn = document.createElement('button');
        closeBtn.className = 'modal-close';
        closeBtn.textContent = '×';
        closeBtn.addEventListener('click', () => overlay.remove());
        header.appendChild(closeBtn);
        modal.appendChild(header);

        // Body
        const body = document.createElement('div');
        body.className = 'changes-popup-body';

        if (flat.length === 0) {
            body.textContent = 'No pending changes.';
        } else {
            const table = document.createElement('table');
            table.className = 'changes-popup-table';
            const thead = document.createElement('thead');
            const headRow = document.createElement('tr');
            for (const col of ['ID', 'Field', 'Original Value', 'Changed Value']) {
                const th = document.createElement('th');
                th.textContent = col;
                headRow.appendChild(th);
            }
            thead.appendChild(headRow);
            table.appendChild(thead);

            const tbody = document.createElement('tbody');
            for (const entry of flat) {
                const tr = document.createElement('tr');
                for (const val of [entry.id, entry.field, this._formatCellValue(entry.original), this._formatCellValue(entry.changed)]) {
                    const td = document.createElement('td');
                    td.textContent = val;
                    tr.appendChild(td);
                }
                tbody.appendChild(tr);
            }
            table.appendChild(tbody);
            body.appendChild(table);
        }

        modal.appendChild(body);
        overlay.appendChild(modal);
        document.body.appendChild(overlay);

        // Escape to close
        const escHandler = (e) => {
            if (e.key === 'Escape') {
                overlay.remove();
                document.removeEventListener('keydown', escHandler, true);
            }
        };
        document.addEventListener('keydown', escHandler, true);
    }

    _formatCellValue(val) {
        if (val === null || val === undefined) return '';
        if (typeof val === 'object') return JSON.stringify(val);
        return String(val);
    }

    showGridChangesContextMenu(event) {
        this.dismissCardContextMenu();

        const allChanges = this._getAllChanges();
        if (allChanges.length === 0) return;

        const menu = document.createElement('div');
        menu.className = 'card-context-menu';
        menu.id = 'card-context-menu';

        const option = document.createElement('div');
        option.className = 'card-context-menu-item';
        option.textContent = `Show Changes (${allChanges.length})`;
        option.addEventListener('click', (e) => {
            e.stopPropagation();
            this.dismissCardContextMenu();
            this._showChangesPopup(allChanges, null);
        });
        menu.appendChild(option);

        document.body.appendChild(menu);
        const menuRect = menu.getBoundingClientRect();
        let left = event.clientX;
        let top = event.clientY;
        if (left + menuRect.width > window.innerWidth) left = window.innerWidth - menuRect.width - 4;
        if (top + menuRect.height > window.innerHeight) top = window.innerHeight - menuRect.height - 4;
        menu.style.left = `${left}px`;
        menu.style.top = `${top}px`;

        const dismissHandler = (e) => {
            if (!menu.contains(e.target)) {
                this.dismissCardContextMenu();
                document.removeEventListener('mousedown', dismissHandler, true);
                document.removeEventListener('keydown', escDismiss, true);
            }
        };
        const escDismiss = (e) => {
            if (e.key === 'Escape') {
                this.dismissCardContextMenu();
                document.removeEventListener('mousedown', dismissHandler, true);
                document.removeEventListener('keydown', escDismiss, true);
            }
        };
        requestAnimationFrame(() => {
            document.addEventListener('mousedown', dismissHandler, true);
            document.addEventListener('keydown', escDismiss, true);
        });
    }

    getFilterSetupEntries(item, surfaceElement = null) {
        const fieldNames = surfaceElement && surfaceElement.classList.contains('grid-table-row')
            ? Array.from(surfaceElement.querySelectorAll('td[data-field-name]'), (cell) => cell.dataset.fieldName)
            : this.app.normalizeFieldList(Object.values(this.app.currentCardSelections || {}));

        return this.app.normalizeFieldList(fieldNames).map((fieldName) => {
            const rawValue = this.app.getFieldValue(item, fieldName);
            const displayValue = this.app.getDisplayValue(item, fieldName);
            const filterValues = Array.isArray(rawValue)
                ? Array.from(new Set(rawValue.map((value) => String(value)).filter((value) => value !== '')))
                : [rawValue === undefined || rawValue === null || rawValue === '' ? '' : String(rawValue)];
            const hasMeaningfulValue = !(filterValues.length === 1 && filterValues[0] === '');

            return {
                fieldName,
                displayValue: displayValue || '(empty)',
                filterValues,
                defaultChecked: hasMeaningfulValue
            };
        });
    }

    openFilterSetupDialog(item, surfaceElement = null) {
        if (!this.filterSetupDialog) {
            return;
        }

        const fieldsContainer = document.getElementById('filter-setup-fields');
        if (!fieldsContainer) {
            return;
        }

        const entries = this.getFilterSetupEntries(item, surfaceElement);
        if (entries.length === 0) {
            this.app.showNotification('No displayed fields available for filter setup', 'info');
            return;
        }

        fieldsContainer.innerHTML = '';

        entries.forEach((entry, index) => {
            const row = document.createElement('label');
            row.className = 'filter-setup-row';

            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.name = 'filter-setup-field';
            checkbox.value = String(index);
            checkbox.checked = entry.defaultChecked;

            const fieldName = document.createElement('div');
            fieldName.className = 'filter-setup-field-name';
            fieldName.textContent = entry.fieldName;

            const fieldValue = document.createElement('div');
            fieldValue.className = 'filter-setup-field-value';
            fieldValue.textContent = entry.displayValue;
            fieldValue.classList.toggle('is-empty', entry.displayValue === '(empty)');

            row.appendChild(checkbox);
            row.appendChild(fieldName);
            row.appendChild(fieldValue);
            fieldsContainer.appendChild(row);
        });

        this.filterSetupContext = { entries };
        this.filterSetupDialog.classList.remove('hidden');
    }

    applyFilterSetupSelection() {
        if (!this.filterSetupDialog || !this.filterSetupContext) {
            return;
        }

        const selectedCheckboxes = Array.from(
            this.filterSetupDialog.querySelectorAll('input[name="filter-setup-field"]:checked')
        );

        if (selectedCheckboxes.length === 0) {
            this.app.showNotification('Select at least one field to apply a filter', 'info');
            return;
        }

        selectedCheckboxes.forEach((checkbox) => {
            const entry = this.filterSetupContext.entries[Number(checkbox.value)];
            if (!entry) {
                return;
            }

            this.app.currentFilters.set(entry.fieldName, new Set(entry.filterValues));
        });

        this.app.updateFilteredData();
        this.app.updateViewConfiguration();
        this.app.renderSlicers();
        this.app.renderGrid();
        this.filterSetupDialog.classList.add('hidden');
        this.filterSetupContext = null;
        this.app.showNotification(`Applied ${selectedCheckboxes.length} filter${selectedCheckboxes.length === 1 ? '' : 's'}`, 'success');
    }

    getDefaultCardClickConfig() {
        return {
            mode: 'edit-tooltip',
            openInNewTab: true,
            openUrl: {
                baseUrl: '',
                suffix: '',
                useUrlField: true,
                argumentGenerator: {
                    type: 'template',
                    template: '${id}',
                    field: '',
                    source: ''
                }
            }
        };
    }

    getEffectiveCardClickConfig() {
        const defaultConfig = this.getDefaultCardClickConfig();
        const datasetConfig = this.normalizeCardClickConfig(
            typeof this.app.getDatasetCardClickConfig === 'function'
                ? this.app.getDatasetCardClickConfig()
                : this.app.metaInfo && this.app.metaInfo.cardClick
        );
        const viewConfig = this.normalizeCardClickConfig(
            typeof this.app.getViewCardClickConfig === 'function'
                ? this.app.getViewCardClickConfig()
                : this.app.viewConfig && this.app.viewConfig.cardClick
        );
        const mergedConfig = this.mergeCardClickConfigs(defaultConfig, datasetConfig, viewConfig);

        if (mergedConfig.mode === 'open-url') {
            const legacyOpenUrlConfig = this.getLegacyOpenUrlConfig();
            mergedConfig.openUrl = {
                ...legacyOpenUrlConfig,
                ...mergedConfig.openUrl
            };
            mergedConfig.openUrl.argumentGenerator = {
                ...defaultConfig.openUrl.argumentGenerator,
                ...(legacyOpenUrlConfig.argumentGenerator || {}),
                ...(mergedConfig.openUrl.argumentGenerator || {})
            };
        }

        return mergedConfig;
    }

    mergeCardClickConfigs(...configs) {
        return configs.reduce((merged, config) => {
            if (!config) {
                return merged;
            }

            return {
                ...merged,
                ...config,
                openUrl: {
                    ...(merged.openUrl || {}),
                    ...(config.openUrl || {}),
                    argumentGenerator: {
                        ...((merged.openUrl && merged.openUrl.argumentGenerator) || {}),
                        ...((config.openUrl && config.openUrl.argumentGenerator) || {})
                    }
                }
            };
        }, {});
    }

    normalizeCardClickConfig(config) {
        if (!config || typeof config !== 'object') {
            return null;
        }

        const hasOwn = (key) => Object.prototype.hasOwnProperty.call(config, key);
        const openUrl = config.openUrl || config.url || config.urlConfig || {};
        const rawGenerator = openUrl.argumentGenerator || config.argumentGenerator || {};
        const result = {};
        const hasMode = hasOwn('mode') || hasOwn('behavior') || hasOwn('clickMode');
        const rawMode = String(config.mode || config.behavior || config.clickMode || '').toLowerCase();

        if (hasMode) {
            result.mode = rawMode === 'open-url' || rawMode === 'url' || rawMode === 'new-tab'
                ? 'open-url'
                : 'edit-tooltip';
        }

        if (hasOwn('openInNewTab') || Object.prototype.hasOwnProperty.call(openUrl, 'openInNewTab')) {
            result.openInNewTab = config.openInNewTab !== false && openUrl.openInNewTab !== false;
        }

        const normalizedOpenUrl = {};
        if (hasOwn('baseUrl') || hasOwn('prefix') || Object.prototype.hasOwnProperty.call(openUrl, 'baseUrl') || Object.prototype.hasOwnProperty.call(openUrl, 'prefix')) {
            normalizedOpenUrl.baseUrl = String(openUrl.baseUrl || config.baseUrl || openUrl.prefix || config.prefix || '');
        }

        if (hasOwn('suffix') || Object.prototype.hasOwnProperty.call(openUrl, 'suffix')) {
            normalizedOpenUrl.suffix = String(openUrl.suffix || config.suffix || '');
        }

        if (hasOwn('useUrlField') || Object.prototype.hasOwnProperty.call(openUrl, 'useUrlField')) {
            normalizedOpenUrl.useUrlField = openUrl.useUrlField !== false && config.useUrlField !== false;
        }

        const hasGenerator =
            hasOwn('generatorType') ||
            hasOwn('argumentTemplate') ||
            hasOwn('argumentField') ||
            hasOwn('functionCode') ||
            Object.prototype.hasOwnProperty.call(openUrl, 'generatorType') ||
            Object.prototype.hasOwnProperty.call(openUrl, 'argumentTemplate') ||
            Object.prototype.hasOwnProperty.call(openUrl, 'argumentField') ||
            Object.prototype.hasOwnProperty.call(openUrl, 'functionCode') ||
            Object.prototype.hasOwnProperty.call(rawGenerator, 'type') ||
            Object.prototype.hasOwnProperty.call(rawGenerator, 'kind') ||
            Object.prototype.hasOwnProperty.call(rawGenerator, 'template') ||
            Object.prototype.hasOwnProperty.call(rawGenerator, 'expression') ||
            Object.prototype.hasOwnProperty.call(rawGenerator, 'field') ||
            Object.prototype.hasOwnProperty.call(rawGenerator, 'source') ||
            Object.prototype.hasOwnProperty.call(rawGenerator, 'functionCode') ||
            Object.prototype.hasOwnProperty.call(rawGenerator, 'body');

        if (hasGenerator) {
            const generatorType = this.normalizeArgumentGeneratorType(
                rawGenerator.type ||
                rawGenerator.kind ||
                openUrl.generatorType ||
                config.generatorType ||
                (rawGenerator.template || openUrl.argumentTemplate || config.argumentTemplate || rawGenerator.expression ? 'template' : null) ||
                (rawGenerator.field || openUrl.argumentField || config.argumentField ? 'field' : null) ||
                (rawGenerator.source || rawGenerator.functionCode || openUrl.functionCode || config.functionCode ? 'raw-js' : null)
            );

            normalizedOpenUrl.argumentGenerator = {
                type: generatorType,
                template: String(
                    rawGenerator.template ||
                    rawGenerator.expression ||
                    openUrl.argumentTemplate ||
                    config.argumentTemplate ||
                    ''
                ),
                field: String(rawGenerator.field || openUrl.argumentField || config.argumentField || ''),
                source: String(
                    rawGenerator.source ||
                    rawGenerator.functionCode ||
                    rawGenerator.body ||
                    openUrl.functionCode ||
                    config.functionCode ||
                    ''
                )
            };
        }

        if (Object.keys(normalizedOpenUrl).length > 0) {
            result.openUrl = normalizedOpenUrl;
        }

        return Object.keys(result).length > 0 ? result : null;
    }

    normalizeArgumentGeneratorType(type) {
        const normalizedType = String(type || '').toLowerCase();

        if (normalizedType === 'field') {
            return 'field';
        }

        if (normalizedType === 'raw-js' || normalizedType === 'rawjs' || normalizedType === 'function' || normalizedType === 'js') {
            return 'raw-js';
        }

        return 'template';
    }

    getLegacyOpenUrlConfig() {
        return {
            baseUrl: this.urlConfig.prefix || '',
            suffix: this.urlConfig.suffix || '',
            useUrlField: this.urlConfig.useUrlField !== false,
            argumentGenerator: {
                type: 'field',
                field: ''
            }
        };
    }

    openConfiguredCardUrl(item, clickConfig) {
        const url = this.constructItemUrl(item, clickConfig);
        if (!url) {
            return false;
        }

        const target = clickConfig.openInNewTab === false ? '_self' : '_blank';
        window.open(url, target, 'noopener');
        return true;
    }

    constructItemUrl(item, clickConfig = null) {
        const openUrlConfig = clickConfig && clickConfig.openUrl ? clickConfig.openUrl : null;
        const useUrlField = openUrlConfig
            ? openUrlConfig.useUrlField !== false
            : this.urlConfig.useUrlField !== false;

        // First try to use the url field if present and config allows
        if (useUrlField && item.url) {
            return item.url;
        }

        if (openUrlConfig) {
            const argument = this.generateUrlArgument(item, openUrlConfig.argumentGenerator);
            const composedUrl = this.composeConfiguredUrl(openUrlConfig.baseUrl, argument, openUrlConfig.suffix);

            if (composedUrl) {
                return composedUrl;
            }
        }

        // Otherwise construct from prefix/suffix and selected field
        if (this.urlConfig.prefix || this.urlConfig.suffix) {
            const urlField = this.determineUrlField(item);
            const fieldValue = item[urlField];

            if (fieldValue) {
                return this.urlConfig.prefix + String(fieldValue) + this.urlConfig.suffix;
            }
        }

        return null;
    }

    generateUrlArgument(item, generatorConfig = {}) {
        const normalizedGenerator = {
            type: this.normalizeArgumentGeneratorType(generatorConfig.type),
            template: String(generatorConfig.template || ''),
            field: String(generatorConfig.field || ''),
            source: String(generatorConfig.source || '')
        };

        if (normalizedGenerator.type === 'raw-js') {
            return this.executeRawArgumentGenerator(item, normalizedGenerator.source);
        }

        if (normalizedGenerator.type === 'field') {
            const fieldName = normalizedGenerator.field || this.determineUrlField(item);
            return this.serializeGeneratorValue(item[fieldName]);
        }

        const template = normalizedGenerator.template || '${' + this.determineUrlField(item) + '}';
        return template.replace(/\$\{([^}]+)\}/g, (_, fieldName) => this.serializeGeneratorValue(item[fieldName.trim()]));
    }

    executeRawArgumentGenerator(item, source) {
        if (!source) {
            return null;
        }

        try {
            if (source.includes('=>') || source.trim().startsWith('function')) {
                const compiledFactory = new Function('return (' + source + ');');
                const compiledFunction = compiledFactory();
                return compiledFunction(item, { ...item });
            }

            const compiledBody = new Function('item', 'fields', source);
            return compiledBody(item, { ...item });
        } catch (error) {
            console.warn('Failed to execute raw card-click URL generator:', error);
            return null;
        }
    }

    serializeGeneratorValue(value) {
        if (value === undefined || value === null) {
            return '';
        }

        if (Array.isArray(value)) {
            return value.map((entry) => encodeURIComponent(String(entry))).join(',');
        }

        if (typeof value === 'object') {
            return encodeURIComponent(JSON.stringify(value));
        }

        return encodeURIComponent(String(value));
    }

    composeConfiguredUrl(baseUrl, argument, suffix = '') {
        const normalizedBaseUrl = String(baseUrl || '');
        const normalizedArgument = argument === undefined || argument === null ? '' : String(argument);
        const normalizedSuffix = String(suffix || '');

        if (!normalizedBaseUrl && !normalizedArgument && !normalizedSuffix) {
            return null;
        }

        if (!normalizedBaseUrl) {
            return normalizedArgument + normalizedSuffix;
        }

        if (!normalizedArgument) {
            return normalizedBaseUrl + normalizedSuffix;
        }

        if (/[\/?#=]$/.test(normalizedBaseUrl) || /^[\/?#]/.test(normalizedArgument)) {
            return normalizedBaseUrl + normalizedArgument + normalizedSuffix;
        }

        return normalizedBaseUrl.replace(/\/+$/, '') + '/' + normalizedArgument.replace(/^\/+/, '') + normalizedSuffix;
    }
    
    determineUrlField(item) {
        // Priority order for URL construction field
        const preferredFields = ['id', 'key', 'title'];
        for (const field of preferredFields) {
            if (item[field]) return field;
        }
        
        // Fall back to first non-empty scalar field
        const scalarFields = this.app.getFieldsByType('scalar');
        for (const field of scalarFields) {
            if (item[field]) return field;
        }
        
        return 'id'; // Final fallback
    }
    
    showItemDetails(item) {
        // Create a temporary tooltip-like display for item details
        const detailsDiv = document.createElement('div');
        detailsDiv.className = 'item-details-popup';
        detailsDiv.style.cssText = `
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            width: 300px;
            max-height: 400px;
            overflow-y: auto;
            background: white;
            border: 1px solid #ccc;
            border-radius: 8px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.3);
            padding: 20px;
            z-index: 9999;
        `;
        
        const title = document.createElement('h3');
        title.textContent = item.title || item.id || 'Item Details';
        title.style.marginBottom = '12px';
        detailsDiv.appendChild(title);
        
        // Add close button
        const closeBtn = document.createElement('button');
        closeBtn.textContent = '×';
        closeBtn.style.cssText = `
            position: absolute;
            top: 10px;
            right: 15px;
            background: none;
            border: none;
            font-size: 20px;
            cursor: pointer;
            color: #999;
        `;
        closeBtn.addEventListener('click', () => {
            document.body.removeChild(detailsDiv);
        });
        detailsDiv.appendChild(closeBtn);
        
        // Populate with item details (similar to tooltip)
        this.populateItemDetails(detailsDiv, item);
        
        document.body.appendChild(detailsDiv);
        
        // Auto-close after 5 seconds
        setTimeout(() => {
            if (detailsDiv.parentNode) {
                document.body.removeChild(detailsDiv);
            }
        }, 5000);
    }
    
    populateItemDetails(container, item) {
        const fields = Object.keys(item).sort();
        
        fields.forEach(fieldName => {
            if (fieldName === 'id') return; // Already shown in title
            
            const fieldType = this.app.fieldTypes.get(fieldName);
            const value = this.app.getFieldValue(item, fieldName);
            
            if (this.shouldSkipField(fieldName, value)) return;
            
            const fieldDiv = document.createElement('div');
            fieldDiv.style.marginBottom = '8px';
            
            const label = document.createElement('strong');
            label.textContent = this.formatFieldLabel(fieldName) + ': ';
            label.style.fontSize = '12px';
            
            const valueSpan = document.createElement('span');
            valueSpan.style.fontSize = '13px';
            this.populateTooltipFieldValue(valueSpan, value, fieldType, fieldName);
            
            fieldDiv.appendChild(label);
            fieldDiv.appendChild(valueSpan);
            container.appendChild(fieldDiv);
        });
    }
    
    // ===== URL CONFIGURATION =====
    
    loadUrlConfiguration() {
        // Load URL configuration from localStorage or use defaults
        const savedConfig = localStorage.getItem('grid-url-config');
        if (savedConfig) {
            try {
                this.urlConfig = { ...this.urlConfig, ...JSON.parse(savedConfig) };
            } catch (error) {
                console.warn('Failed to load URL configuration:', error);
            }
        }
    }
    
    updateUrlConfiguration(config) {
        this.urlConfig = { ...this.urlConfig, ...config };
        localStorage.setItem('grid-url-config', JSON.stringify(this.urlConfig));
    }

    setUrlConfig(config) {
        this.updateUrlConfiguration(config);
    }
    
    // ===== MISSING FIELD HANDLING =====
    
    handleMissingField(item, fieldName) {
        // Missing field values are treated as empty according to design spec
        return '';
    }
}

// ===== INTEGRATION WITH MAIN APP =====

function attachInteractionsIntegration(targetApp) {
    if (!targetApp || targetApp.__interactionsIntegrated) {
        return;
    }

    targetApp.initializeInteractions = function() {
        if (!this.interactionManager) {
            this.interactionManager = new InteractionManager(this);
        }
        return this.interactionManager;
    };

    targetApp.showTooltip = function(event, item, mode) {
        if (this.interactionManager) {
            this.interactionManager.showTooltip(event, item, mode);
        }
    };

    targetApp.hideTooltip = function(mode) {
        if (this.interactionManager) {
            this.interactionManager.hideTooltip(mode);
        }
    };

    targetApp.handleCardClick = function(event, item) {
        if (this.interactionManager) {
            this.interactionManager.handleCardClick(event, item);
        }
    };

    targetApp.handleCardDoubleClick = function(event, item) {
        if (this.interactionManager) {
            this.interactionManager.handleCardDoubleClick(event, item);
        }
    };

    targetApp.handleCardDrop = function(itemId, newRowValue, newColValue) {
        if (this.interactionManager) {
            this.interactionManager.handleCardDrop(itemId, newRowValue, newColValue);
        }
    };

    targetApp.addTagToItem = function(itemId, tagName) {
        if (this.interactionManager) {
            return this.interactionManager.addTagToItem(itemId, tagName);
        }
        return false;
    };

    targetApp.removeTagFromItem = function(itemId, tagName, options) {
        if (this.interactionManager) {
            return this.interactionManager.removeTagFromItem(itemId, tagName, options);
        }
        return false;
    };

    targetApp.openInlineValueEditor = function(card, item, fieldName, element) {
        if (this.interactionManager) {
            this.interactionManager.openInlineValueEditor(card, item, fieldName, element);
        }
    };

    targetApp.openCreateTooltip = function(event, initialValues) {
        if (this.interactionManager) {
            this.interactionManager.showCreateTooltip(event, initialValues);
        }
    };

    targetApp.openCreatePanel = function(initialValues) {
        if (this.detailsPanelManager && typeof this.detailsPanelManager.openForCreate === 'function') {
            this.detailsPanelManager.openForCreate(initialValues);
            return;
        }

        if (this.interactionManager) {
            this.interactionManager.showCreateTooltip(null, initialValues);
        }
    };

    targetApp.__interactionsIntegrated = true;
}

document.addEventListener('DOMContentLoaded', () => {
    if (typeof app !== 'undefined' && app) {
        attachInteractionsIntegration(app);
        app.initializeInteractions();
    }
});