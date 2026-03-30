// Copyright (c) 2026 Malte Doerper. MIT License. See LICENSE file.

// Control Panel Implementation
// Handles field selectors, slicers, data management controls, and tag management

class ControlPanelManager {
    constructor(app) {
        this.app = app;
        this.currentEditingTag = null;
        this.tagDialogMode = 'edit';
        this.openSlicerField = null;
        this.slicerSearchTerms = new Map();
        this.draggedTagName = null;
        this.draggedGroupId = null;
        this.initializeEventListeners();
    }
    
    initializeEventListeners() {
        // Field selector change handlers
        document.getElementById('x-axis-select').addEventListener('change', (e) => {
            this.app.currentAxisSelections.x = e.target.value || null;
            this.app.updateViewConfiguration();
            this.app.renderGrid();
            this.updateSlicerAxisShading();
        });
        
        document.getElementById('y-axis-select').addEventListener('change', (e) => {
            this.app.currentAxisSelections.y = e.target.value || null;
            this.app.updateViewConfiguration();
            this.app.renderGrid();
            this.updateSlicerAxisShading();
        });
        
        document.getElementById('title-select').addEventListener('change', (e) => {
            this.app.currentCardSelections.title = e.target.value || null;
            this.app.updateViewConfiguration();
            this.app.renderGrid();
            this.renderTableColumnSelector();
            this.updateSlicerAxisShading();
        });
        
        document.getElementById('top-left-select').addEventListener('change', (e) => {
            this.app.currentCardSelections.topLeft = e.target.value || null;
            this.app.updateViewConfiguration();
            this.app.renderGrid();
            this.renderTableColumnSelector();
            this.updateSlicerAxisShading();
        });
        
        document.getElementById('top-right-select').addEventListener('change', (e) => {
            this.app.currentCardSelections.topRight = e.target.value || null;
            this.app.updateViewConfiguration();
            this.app.renderGrid();
            this.renderTableColumnSelector();
            this.updateSlicerAxisShading();
        });
        
        document.getElementById('bottom-right-select').addEventListener('change', (e) => {
            this.app.currentCardSelections.bottomRight = e.target.value || null;
            this.app.updateViewConfiguration();
            this.app.renderGrid();
            this.renderTableColumnSelector();
            this.updateSlicerAxisShading();
        });

        document.getElementById('card-mode-tab').addEventListener('click', () => {
            this.setLayoutMode('cards');
        });

        document.getElementById('table-mode-tab').addEventListener('click', () => {
            this.setLayoutMode('table');
        });

        document.getElementById('table-columns-add-btn').addEventListener('click', () => {
            this.addSelectedTableColumns();
        });

        document.getElementById('table-columns-remove-btn').addEventListener('click', () => {
            this.removeSelectedTableColumns();
        });

        document.getElementById('table-columns-move-up-btn').addEventListener('click', () => {
            this.moveSelectedTableColumns(-1);
        });

        document.getElementById('table-columns-move-down-btn').addEventListener('click', () => {
            this.moveSelectedTableColumns(1);
        });

        document.getElementById('table-columns-sort-az-btn').addEventListener('click', () => {
            this.sortSelectedTableColumns();
        });

        document.getElementById('table-summary-mode-select').addEventListener('change', (e) => {
            this.app.tableSummaryMode = this.app.normalizeTableSummaryMode(e.target.value);
            this.app.updateViewConfiguration();
            this.app.renderGrid();
        });

        document.getElementById('table-available-fields').addEventListener('change', () => {
            this.updateTableColumnActionState();
        });

        document.getElementById('table-selected-fields').addEventListener('change', () => {
            this.updateTableColumnActionState();
        });
        
        // Data management controls
        document.getElementById('load-data-btn').addEventListener('click', () => {
            this.app.persistenceManager.fileService.debugLog('load button clicked', {
                action: 'load',
                type: 'data',
                userScope: this.app.persistenceManager.fileService.getCurrentUserScope()
            });
            this.app.persistenceManager.openDataFile();
        });

        document.getElementById('reload-data-btn').addEventListener('click', () => {
            this.app.persistenceManager.fileService.debugLog('reload button clicked', {
                action: 'reload',
                type: 'data',
                userScope: this.app.persistenceManager.fileService.getCurrentUserScope(),
                reloadAction: this.app.persistenceManager.fileService.getReloadAction('data')
            });
            this.app.persistenceManager.reloadDataFile();
        });
        
        document.getElementById('save-data-btn').addEventListener('click', () => {
            this.app.persistenceManager.fileService.debugLog('save button clicked', {
                action: 'save',
                type: 'data',
                userScope: this.app.persistenceManager.fileService.getCurrentUserScope(),
                suggestedFilename: this.app.persistenceManager.fileService.getSuggestedFilename(
                    'data',
                    this.app.persistenceManager.fileService.generateDataFilename(this.app.dataset.length)
                )
            });
            this.app.persistenceManager.saveDataFile();
        });

        document.getElementById('promote-data-btn').addEventListener('click', () => {
            this.app.persistenceManager.fileService.debugLog('promote changes button clicked', {
                action: 'promote-save',
                type: 'data',
                userScope: this.app.persistenceManager.fileService.getCurrentUserScope(),
                suggestedFilename: this.app.persistenceManager.fileService.getSuggestedFilename(
                    'data',
                    this.app.persistenceManager.fileService.generateDataFilename(this.app.dataset.length)
                )
            });
            this.app.persistenceManager.saveDataFile({ promoteChanges: true });
        });
        
        document.getElementById('export-filtered-btn').addEventListener('click', () => {
            this.app.persistenceManager.fileService.debugLog('export button clicked', {
                action: 'export',
                type: 'data',
                userScope: this.app.persistenceManager.fileService.getCurrentUserScope(),
                filteredItemCount: this.app.getFilteredData().length
            });
            this.exportFilteredData();
        });
        
        // View configuration controls
        document.getElementById('load-view-btn').addEventListener('click', () => {
            this.app.persistenceManager.fileService.debugLog('load view button clicked', {
                action: 'load',
                type: 'view',
                serverMode: Boolean(this.app.isServerMode),
                currentDataset: this.app.serverControlsManager && this.app.serverControlsManager.currentDataset
                    ? this.app.serverControlsManager.currentDataset.filename
                    : null,
            });
            if (this.app.isServerMode && this.app.serverControlsManager && typeof this.app.serverControlsManager.openResolvedViewConfiguration === 'function') {
                void this.app.serverControlsManager.openResolvedViewConfiguration();
                return;
            }
            this.app.persistenceManager.openViewFile();
        });
        
        document.getElementById('save-view-btn').addEventListener('click', () => {
            this.app.persistenceManager.fileService.debugLog('save view button clicked', {
                action: 'save',
                type: 'view',
                serverMode: Boolean(this.app.isServerMode),
                currentDataset: this.app.serverControlsManager && this.app.serverControlsManager.currentDataset
                    ? this.app.serverControlsManager.currentDataset.filename
                    : null,
            });
            if (this.app.isServerMode && this.app.serverControlsManager && typeof this.app.serverControlsManager.saveResolvedViewConfiguration === 'function') {
                void this.app.serverControlsManager.saveResolvedViewConfiguration();
                return;
            }
            this.app.persistenceManager.saveViewConfiguration();
        });

        document.getElementById('compare-view-btn').addEventListener('click', () => {
            this.app.openCompareWindow();
        });

        document.getElementById('add-tag-btn').addEventListener('click', () => {
            this.openNewTagDialog();
        });

        document.getElementById('add-group-btn').addEventListener('click', () => {
            this.openGroupEditDialog(null);
        });

        this.initializeCurrentUserEditing();

        // Filter toggle checkboxes
        document.getElementById('show-relationship-fields-cb').addEventListener('change', (e) => {
            this.app.showRelationshipFields = e.target.checked;
            this.renderSlicers();
        });
        document.getElementById('show-derived-fields-cb').addEventListener('change', (e) => {
            this.app.showDerivedFields = e.target.checked;
            this.renderSlicers();
        });
        document.getElementById('show-tooltips-cb').addEventListener('change', (e) => {
            this.app.showTooltips = e.target.checked;
        });
        document.getElementById('grow-cards-cb').addEventListener('change', (e) => {
            this.app.growCards = e.target.checked;
            this.app.updateViewConfiguration();
            this.app.renderGrid();
        });
        document.addEventListener('click', (e) => {
            if (!e.target.closest('.slicer')) {
                this.closeAllSlicers();
            }

            if (!e.target.closest('.card-tag.is-removable')) {
                this.app.clearSelectedCardTag();
            }

            if (!e.target.closest('#tags-container .tag')) {
                this.app.clearSelectedControlTags();
            }
        });

        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') {
                this.closeAllSlicers();
                this.app.clearSelectedCardTag();
                this.app.clearSelectedControlTags();
                this.app.clearCardSelection();
                return;
            }

            if (
                (e.key === 'Delete' || e.key === 'Backspace') &&
                this.app.selectedCardTag &&
                !this.isEditableTarget(document.activeElement)
            ) {
                e.preventDefault();
                const selectedTags = typeof this.app.getSelectedCardTags === 'function'
                    ? this.app.getSelectedCardTags()
                    : (this.app.selectedCardTag ? [this.app.selectedCardTag] : []);

                if (selectedTags.length === 0) {
                    return;
                }

                let removedCount = 0;
                selectedTags.forEach(({ itemId, tagName }) => {
                    const removed = this.app.removeTagFromItem(itemId, tagName, { suppressNotification: true });
                    if (removed) {
                        removedCount += 1;
                    }
                });

                if (removedCount > 0) {
                    this.app.showNotification(`Removed ${removedCount} tag${removedCount === 1 ? '' : 's'}`, 'success');
                }
            }
        });
        
        // Tag edit dialog handlers
        this.initializeTagEditDialog();

        // Group edit dialog handlers
        this.initializeGroupEditDialog();

        // Drag and drop for tag assignment
        this.initializeTagDragDrop();
        
    }

    isEditableTarget(element) {
        if (!element) {
            return false;
        }

        const tagName = element.tagName;
        return element.isContentEditable ||
            tagName === 'INPUT' ||
            tagName === 'TEXTAREA' ||
            tagName === 'SELECT';
    }
    
    // ===== FIELD SELECTOR MANAGEMENT =====
    
    renderFieldSelectors() {
        const selectors = [
            { id: 'x-axis-select', usage: 'axis', value: this.app.currentAxisSelections.x },
            { id: 'y-axis-select', usage: 'axis', value: this.app.currentAxisSelections.y },
            { id: 'title-select', usage: 'label', value: this.app.currentCardSelections.title },
            { id: 'top-left-select', usage: 'label', value: this.app.currentCardSelections.topLeft },
            { id: 'top-right-select', usage: 'label', value: this.app.currentCardSelections.topRight },
            { id: 'bottom-right-select', usage: 'value', value: this.app.currentCardSelections.bottomRight }
        ];
        
        selectors.forEach(({ id, usage, value }) => {
            this.populateFieldSelector(id, usage, value);
        });
        
        // Set selected values from current axis selections
        this.updateFieldSelectorValues();
        this.renderTableColumnSelector();
        this.syncModeConfigurationVisibility();
        this.syncTableSummaryControl();

        if (typeof this.app.renderCurrentUserName === 'function') {
            this.app.renderCurrentUserName();
        }
    }

    syncTableSummaryControl() {
        const summarySelect = document.getElementById('table-summary-mode-select');
        if (!summarySelect) {
            return;
        }

        summarySelect.value = this.app.normalizeTableSummaryMode(this.app.tableSummaryMode);
    }

    initializeCurrentUserEditing() {
        const currentUserElement = document.getElementById('current-user-name');
        if (!currentUserElement) {
            return;
        }

        const commitUserName = () => {
            const nextName = currentUserElement.textContent || '';

            if (!nextName.trim()) {
                currentUserElement.textContent = this.app.getCurrentUserName();
                this.app.showNotification('User name cannot be empty', 'warning');
                return;
            }

            const didUpdate = this.app.setCurrentUserName(nextName);
            currentUserElement.textContent = this.app.getCurrentUserName();

            if (didUpdate) {
                this.app.showNotification(`Using ${this.app.getCurrentUserName()} for comments`, 'success');
            }
        };

        currentUserElement.addEventListener('keydown', (event) => {
            if (event.key === 'Enter') {
                event.preventDefault();
                currentUserElement.blur();
                return;
            }

            if (event.key === 'Escape') {
                event.preventDefault();
                currentUserElement.textContent = this.app.getCurrentUserName();
                currentUserElement.blur();
            }
        });

        currentUserElement.addEventListener('blur', commitUserName);
    }
    
    populateFieldSelector(selectorId, usage = 'label', currentSelection = null) {
        const selector = document.getElementById(selectorId);
        selector.innerHTML = '<option value="">Select field...</option>';
        const renderedFields = new Set();
        
        this.app.availableFields.forEach(fieldName => {
            if (usage === 'axis' && !this.app.canUseForAxis(fieldName)) {
                return;
            }

            if (usage === 'value' && !this.app.canUseForValue(fieldName)) {
                return;
            }
            
            const fieldType = this.app.fieldTypes.get(fieldName);
            const option = document.createElement('option');
            option.value = fieldName;
            option.textContent = `${fieldName} (${fieldType})`;
            selector.appendChild(option);
            renderedFields.add(fieldName);
        });

        if (currentSelection && !renderedFields.has(currentSelection)) {
            const fieldType = this.app.fieldTypes.get(currentSelection) || 'unknown';
            const option = document.createElement('option');
            option.value = currentSelection;
            option.textContent = usage === 'axis' || usage === 'value'
                ? `${currentSelection} (${fieldType}, loaded from view)`
                : `${currentSelection} (${fieldType})`;
            selector.appendChild(option);
        }
    }

    updateFieldSelectorValues() {
        const selectorValues = {
            'x-axis-select': this.app.currentAxisSelections.x,
            'y-axis-select': this.app.currentAxisSelections.y,
            'title-select': this.app.currentCardSelections.title,
            'top-left-select': this.app.currentCardSelections.topLeft,
            'top-right-select': this.app.currentCardSelections.topRight,
            'bottom-right-select': this.app.currentCardSelections.bottomRight
        };

        Object.entries(selectorValues).forEach(([selectorId, fieldName]) => {
            const selector = document.getElementById(selectorId);
            if (selector) {
                selector.value = fieldName || '';
            }
        });

        this.updateTagDropTargetSelection();

    }

    renderTableColumnSelector() {
        const availableSelect = document.getElementById('table-available-fields');
        const selectedSelect = document.getElementById('table-selected-fields');

        if (!availableSelect || !selectedSelect) {
            return;
        }

        const availableSelection = new Set(
            Array.from(availableSelect.selectedOptions || [], (option) => option.value)
        );
        const selectedColumnSelection = new Set(
            Array.from(selectedSelect.selectedOptions || [], (option) => option.value)
        );

        const selectedFields = Array.isArray(this.app.tableColumnFields) && this.app.tableColumnFields.length > 0
            ? [...this.app.tableColumnFields]
            : (typeof this.app.getDefaultTableColumnFields === 'function'
                ? this.app.getDefaultTableColumnFields(this.app.currentCardSelections)
                : []);
        const availableFields = this.app.availableFields.filter((fieldName) => {
            if (typeof this.app.canUseForTableColumn === 'function' && !this.app.canUseForTableColumn(fieldName)) {
                return false;
            }

            return !selectedFields.includes(fieldName);
        });

        availableSelect.innerHTML = '';
        selectedSelect.innerHTML = '';

        availableFields.forEach((fieldName) => {
            const option = document.createElement('option');
            option.value = fieldName;
            option.textContent = this.formatFieldOptionLabel(fieldName);
            option.selected = availableSelection.has(fieldName);
            availableSelect.appendChild(option);
        });

        selectedFields.forEach((fieldName) => {
            const option = document.createElement('option');
            option.value = fieldName;
            option.textContent = this.formatFieldOptionLabel(fieldName, !this.app.availableFields.includes(fieldName));
            option.selected = selectedColumnSelection.has(fieldName);
            selectedSelect.appendChild(option);
        });

        this.updateTableColumnActionState();
        this.updateSlicerAxisShading();
    }

    updateTableColumnActionState() {
        const availableSelect = document.getElementById('table-available-fields');
        const selectedSelect = document.getElementById('table-selected-fields');
        const addButton = document.getElementById('table-columns-add-btn');
        const removeButton = document.getElementById('table-columns-remove-btn');
        const moveUpButton = document.getElementById('table-columns-move-up-btn');
        const moveDownButton = document.getElementById('table-columns-move-down-btn');
        const sortButton = document.getElementById('table-columns-sort-az-btn');

        if (!availableSelect || !selectedSelect) {
            return;
        }

        const availableSelectedCount = Array.from(availableSelect.selectedOptions || []).length;
        const selectedIndices = Array.from(selectedSelect.selectedOptions || [], (option) => option.index).sort((a, b) => a - b);
        const hasSelection = selectedIndices.length > 0;
        const selectedCount = selectedSelect.options.length;

        if (addButton) {
            addButton.disabled = availableSelectedCount === 0;
        }

        if (removeButton) {
            removeButton.disabled = !hasSelection;
        }

        if (moveUpButton) {
            moveUpButton.disabled = !hasSelection || selectedIndices[0] === 0;
        }

        if (moveDownButton) {
            moveDownButton.disabled = !hasSelection || selectedIndices[selectedIndices.length - 1] === selectedCount - 1;
        }

        if (sortButton) {
            sortButton.disabled = selectedCount < 2;
        }
    }

    formatFieldOptionLabel(fieldName, isLoadedFromView = false) {
        const fieldType = this.app.fieldTypes.get(fieldName) || 'unknown';
        return isLoadedFromView
            ? `${fieldName} (${fieldType}, loaded from view)`
            : `${fieldName} (${fieldType})`;
    }

    addSelectedTableColumns() {
        const availableSelect = document.getElementById('table-available-fields');
        if (!availableSelect) {
            return;
        }

        const selectedValues = Array.from(availableSelect.selectedOptions || [], (option) => option.value);
        if (selectedValues.length === 0) {
            return;
        }

        const nextColumns = typeof this.app.normalizeFieldList === 'function'
            ? this.app.normalizeFieldList([...(this.app.tableColumnFields || []), ...selectedValues])
            : [...new Set([...(this.app.tableColumnFields || []), ...selectedValues])];

        this.app.tableColumnFields = nextColumns;
        this.app.updateViewConfiguration();
        this.renderTableColumnSelector();
        this.app.renderGrid();
    }

    removeSelectedTableColumns() {
        const selectedSelect = document.getElementById('table-selected-fields');
        if (!selectedSelect) {
            return;
        }

        const valuesToRemove = new Set(Array.from(selectedSelect.selectedOptions || [], (option) => option.value));
        if (valuesToRemove.size === 0) {
            return;
        }

        this.app.tableColumnFields = (this.app.tableColumnFields || []).filter((fieldName) => !valuesToRemove.has(fieldName));
        this.app.updateViewConfiguration();
        this.renderTableColumnSelector();
        this.app.renderGrid();
    }

    moveSelectedTableColumns(direction) {
        const selectedSelect = document.getElementById('table-selected-fields');
        if (!selectedSelect || !Array.isArray(this.app.tableColumnFields)) {
            return;
        }

        const selectedIndices = Array.from(selectedSelect.selectedOptions || [], (option) => option.index).sort((a, b) => a - b);
        if (selectedIndices.length === 0) {
            return;
        }

        const nextColumns = [...this.app.tableColumnFields];
        const traversal = direction < 0 ? selectedIndices : [...selectedIndices].reverse();

        traversal.forEach((index) => {
            const targetIndex = index + direction;
            if (targetIndex < 0 || targetIndex >= nextColumns.length) {
                return;
            }

            const currentValue = nextColumns[index];
            nextColumns[index] = nextColumns[targetIndex];
            nextColumns[targetIndex] = currentValue;
        });

        const nextSelectedValues = selectedIndices
            .map((index) => nextColumns[index + direction])
            .filter(Boolean);

        this.applyTableColumnOrder(nextColumns, nextSelectedValues);
    }

    sortSelectedTableColumns() {
        if (!Array.isArray(this.app.tableColumnFields) || this.app.tableColumnFields.length < 2) {
            return;
        }

        const sortedColumns = [...this.app.tableColumnFields].sort((left, right) =>
            String(left).localeCompare(String(right), undefined, { sensitivity: 'base' })
        );

        this.applyTableColumnOrder(sortedColumns, this.app.tableColumnFields);
    }

    applyTableColumnOrder(nextColumns, selectedValues = []) {
        this.app.tableColumnFields = [...nextColumns];
        this.app.updateViewConfiguration();
        this.renderTableColumnSelector();

        const selectedSelect = document.getElementById('table-selected-fields');
        if (selectedSelect && Array.isArray(selectedValues) && selectedValues.length > 0) {
            const selectedSet = new Set(selectedValues);
            Array.from(selectedSelect.options).forEach((option) => {
                option.selected = selectedSet.has(option.value);
            });
            this.updateTableColumnActionState();
        }

        this.app.renderGrid();
    }

    setLayoutMode(mode) {
        const normalizedMode = mode === 'table' ? 'table' : 'cards';
        if (this.app.cellRenderMode === normalizedMode) {
            this.syncModeConfigurationVisibility();
            return;
        }

        this.app.cellRenderMode = normalizedMode;
        this.app.updateViewConfiguration();
        this.syncModeConfigurationVisibility();
        this.app.renderGrid();
        this.updateSlicerAxisShading();
    }

    syncModeConfigurationVisibility() {
        const isTableMode = this.app.cellRenderMode === 'table';
        this.applyModeConfigurationState('card', !isTableMode);
        this.applyModeConfigurationState('table', isTableMode);

        const cardTab = document.getElementById('card-mode-tab');
        const tableTab = document.getElementById('table-mode-tab');
        if (cardTab) {
            cardTab.classList.toggle('is-active', !isTableMode);
            cardTab.setAttribute('aria-selected', String(!isTableMode));
        }
        if (tableTab) {
            tableTab.classList.toggle('is-active', isTableMode);
            tableTab.setAttribute('aria-selected', String(isTableMode));
        }
    }

    applyModeConfigurationState(panelName, isVisible) {
        const section = document.getElementById(`${panelName}-layout-section`);
        const panel = document.getElementById(`${panelName}-layout-panel`);

        if (!section || !panel) {
            return;
        }

        section.classList.toggle('hidden', !isVisible);
        panel.classList.toggle('hidden', !isVisible);
    }
    
    // ===== SLICER MANAGEMENT =====
    
    renderSlicers() {
        const slicersContainer = document.getElementById('slicers-container');
        slicersContainer.innerHTML = '';
        
        // Create slicers for filterable fields, grouped by kind
        const filterableFields = this.app.availableFields.filter(field => 
            this.app.canUseForFilter(field)
        );

        const dataFields = filterableFields.filter(f => this.app.getFieldKind(f) === 'data');
        const relationshipFields = filterableFields.filter(f => this.app.getFieldKind(f) === 'relationship');
        const derivedFields = filterableFields.filter(f => {
            const kind = this.app.getFieldKind(f);
            return kind === 'derived';
        });

        // Data fields
        dataFields.forEach(fieldName => {
            this.createSlicer(fieldName, slicersContainer);
        });

        // Relationship fields (with separator)
        if (relationshipFields.length > 0) {
            const sep = document.createElement('div');
            sep.className = 'slicer-separator';
            slicersContainer.appendChild(sep);
            relationshipFields.forEach(fieldName => {
                this.createSlicer(fieldName, slicersContainer);
            });
        }

        // Derived fields (with separator)
        if (derivedFields.length > 0) {
            const sep = document.createElement('div');
            sep.className = 'slicer-separator';
            slicersContainer.appendChild(sep);
            derivedFields.forEach(fieldName => {
                this.createSlicer(fieldName, slicersContainer);
            });
        }

        // Add group slicer if groups exist
        const groups = this.app.getGroups();
        if (groups.length > 0) {
            this.createGroupSlicer(slicersContainer);
        }
    }
    
    updateSlicerAxisShading() {
        const slicers = document.querySelectorAll('#slicers-container .slicer[data-field-name]');
        const highlightedFields = this.getHighlightedSlicerFields();
        slicers.forEach(slicer => {
            const field = slicer.dataset.fieldName;
            slicer.classList.toggle('is-axis-selected', highlightedFields.has(field));
        });
    }

    getHighlightedSlicerFields() {
        const highlightedFields = new Set();
        const axisSelections = this.app.currentAxisSelections || {};
        const cardSelections = this.app.currentCardSelections || {};
        const tableColumnFields = Array.isArray(this.app.tableColumnFields)
            ? this.app.tableColumnFields
            : [];
        const isTableMode = this.app.cellRenderMode === 'table';

        Object.values(axisSelections).forEach((fieldName) => {
            if (fieldName) {
                highlightedFields.add(fieldName);
            }
        });

        if (isTableMode) {
            tableColumnFields.forEach((fieldName) => {
                if (fieldName) {
                    highlightedFields.add(fieldName);
                }
            });
        } else {
            Object.values(cardSelections).forEach((fieldName) => {
                if (fieldName) {
                    highlightedFields.add(fieldName);
                }
            });
        }

        return highlightedFields;
    }

    createSlicer(fieldName, container) {
        const template = document.getElementById('slicer-template');
        const slicerElement = template.content.cloneNode(true);
        
        const slicer = slicerElement.querySelector('.slicer');
        const toggle = slicerElement.querySelector('.slicer-toggle');
        const panel = slicerElement.querySelector('.slicer-panel');
        const title = slicerElement.querySelector('.slicer-title');
        const summary = slicerElement.querySelector('.slicer-summary');
        const optionsContainer = slicerElement.querySelector('.slicer-options');
        
        title.textContent = fieldName;
        slicer.dataset.fieldName = fieldName;

        const highlightedFields = this.getHighlightedSlicerFields();
        slicer.classList.toggle('is-axis-selected', highlightedFields.has(fieldName));

        // Get distinct values for this field
        const distinctValues = this.getVisibleSlicerValues(fieldName);
        const hasActiveFilter = this.app.currentFilters.has(fieldName);
        const activeFilters = this.app.currentFilters.get(fieldName) || new Set();
        const isExpanded = this.openSlicerField === fieldName;

        summary.textContent = this.getSlicerSummary(fieldName, distinctValues, activeFilters, hasActiveFilter);
        slicer.classList.toggle('is-open', isExpanded);
        toggle.setAttribute('aria-expanded', String(isExpanded));
        toggle.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            this.toggleSlicer(fieldName);
        });
        slicer.addEventListener('contextmenu', (event) => {
            if (event.target.closest('.slicer-search-input') || event.target.closest('.slicer-control-btn')) {
                return;
            }

            if (this.openFieldValueEditor(fieldName)) {
                event.preventDefault();
                event.stopPropagation();
            }
        });
        
        // Add "select all" / "clear all" controls
        this.addSlicerControls(optionsContainer, fieldName);

        // Add contains-search input like Excel filter lists
        this.addSlicerSearch(panel, fieldName, optionsContainer);

        if (distinctValues.length === 0) {
            const emptyState = document.createElement('div');
            emptyState.className = 'slicer-empty';
            emptyState.textContent = 'No values available';
            optionsContainer.appendChild(emptyState);
            container.appendChild(slicer);
            return;
        }
        
        // Create checkboxes for each distinct value
        distinctValues.forEach((value, index) => {
            const option = document.createElement('div');
            option.className = 'slicer-option';
            option.dataset.searchText = value || '(empty)';
            
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.id = `${fieldName}-${index}`;
            checkbox.value = value;
            checkbox.checked = !hasActiveFilter || activeFilters.has(value);
            
            const label = document.createElement('label');
            label.htmlFor = checkbox.id;
            label.textContent = value || '(empty)';
            
            option.appendChild(checkbox);
            option.appendChild(label);
            optionsContainer.appendChild(option);
            
            // Add event listener
            checkbox.addEventListener('change', () => {
                this.handleSlicerChange(fieldName);
            });

            option.addEventListener('contextmenu', (event) => {
                event.preventDefault();
                event.stopPropagation();
                if (event.shiftKey || !this.openFieldValueEditor(fieldName, value)) {
                    this.applyExclusiveSlicerFilter(fieldName, value);
                }
            });
        });

        // Add (none) option if any items have empty/null values for this field
        if (this.app.hasItemsWithEmptyValue(fieldName)) {
            const option = document.createElement('div');
            option.className = 'slicer-option';
            option.dataset.searchText = '(none)';

            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.id = `${fieldName}-none`;
            checkbox.value = '_none';
            checkbox.checked = !hasActiveFilter || activeFilters.has('');

            const label = document.createElement('label');
            label.htmlFor = checkbox.id;
            label.textContent = '(none)';
            label.style.fontStyle = 'italic';

            option.appendChild(checkbox);
            option.appendChild(label);
            optionsContainer.appendChild(option);

            checkbox.addEventListener('change', () => {
                this.handleSlicerChange(fieldName);
            });

            option.addEventListener('contextmenu', (event) => {
                event.preventDefault();
                event.stopPropagation();
                if (event.shiftKey || !this.openFieldValueEditor(fieldName, '_none')) {
                    this.applyExclusiveSlicerFilter(fieldName, '_none');
                }
            });
        }

        this.applySlicerSearchFilter(fieldName, optionsContainer);
        
        container.appendChild(slicer);
    }

    getVisibleSlicerValues(fieldName) {
        const distinctValues = (this.app.distinctValues.get(fieldName) || []).map((value) => String(value));
        const activeFilters = this.app.currentFilters.get(fieldName);

        // Filter out empty string — empty/null items are handled by the (none) option
        const nonEmptyValues = distinctValues.filter((value) => value !== '');

        if (!(activeFilters instanceof Set)) {
            return nonEmptyValues;
        }

        const selectedValues = Array.from(activeFilters, (value) => String(value));
        // Exclude empty string from missing-value preservation (handled by (none) checkbox)
        const missingSelectedValues = selectedValues.filter((value) => value !== '' && !nonEmptyValues.includes(value));

        if (missingSelectedValues.length > 0 && (fieldName === this.app.currentAxisSelections.x || fieldName === this.app.currentAxisSelections.y)) {
            console.debug('[HEADER] preserving missing axis filter options in slicer', {
                fieldName,
                missingSelectedValues,
                distinctValues: nonEmptyValues,
                selectedValues
            });
        }

        return [...nonEmptyValues, ...missingSelectedValues];
    }

    toggleSlicer(fieldName) {
        this.openSlicerField = this.openSlicerField === fieldName ? null : fieldName;
        this.renderSlicers();
    }

    closeAllSlicers() {
        if (this.openSlicerField === null) {
            return;
        }

        this.openSlicerField = null;
        this.renderSlicers();
    }

    getSlicerSummary(fieldName, distinctValues, activeFilters, hasActiveFilter) {
        if (distinctValues.length === 0) {
            return 'No values';
        }

        if (!hasActiveFilter || this.app.isFilterEquivalentToAll(fieldName, activeFilters)) {
            return 'All';
        }

        if (activeFilters.size === 0) {
            return 'None';
        }

        if (activeFilters.size === 1) {
            const value = Array.from(activeFilters)[0];
            return value === '' ? '(none)' : value;
        }

        return `${activeFilters.size} selected`;
    }
    
    addSlicerControls(container, fieldName) {
        const controlsDiv = document.createElement('div');
        controlsDiv.className = 'slicer-controls';
        
        const selectAllBtn = document.createElement('button');
        selectAllBtn.type = 'button';
        selectAllBtn.textContent = 'All';
        selectAllBtn.className = 'slicer-control-btn';
        
        const clearAllBtn = document.createElement('button');
        clearAllBtn.type = 'button';
        clearAllBtn.textContent = 'None';
        clearAllBtn.className = 'slicer-control-btn';
        
        selectAllBtn.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            this.selectAllSlicerOptions(fieldName, true);
        });
        
        clearAllBtn.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            this.selectAllSlicerOptions(fieldName, false);
        });
        
        controlsDiv.appendChild(selectAllBtn);
        controlsDiv.appendChild(clearAllBtn);
        container.appendChild(controlsDiv);
    }

    addSlicerSearch(panel, fieldName, optionsContainer) {
        if (!panel) {
            return;
        }

        const searchWrapper = document.createElement('div');
        searchWrapper.className = 'slicer-search';

        const searchInput = document.createElement('input');
        searchInput.type = 'search';
        searchInput.className = 'slicer-search-input';
        searchInput.placeholder = 'Search values...';
        searchInput.autocomplete = 'off';
        searchInput.value = this.slicerSearchTerms.get(fieldName) || '';

        searchInput.addEventListener('input', () => {
            this.slicerSearchTerms.set(fieldName, searchInput.value || '');
            this.applySlicerSearchFilter(fieldName, optionsContainer);
        });

        searchWrapper.appendChild(searchInput);
        panel.insertBefore(searchWrapper, optionsContainer);
    }

    applySlicerSearchFilter(fieldName, optionsContainer) {
        if (!optionsContainer) {
            return;
        }

        const query = (this.slicerSearchTerms.get(fieldName) || '').trim().toLowerCase();
        let visibleCount = 0;

        optionsContainer.querySelectorAll('.slicer-option').forEach((option) => {
            const optionText = String(option.dataset.searchText || option.textContent || '').toLowerCase();
            const matches = !query || optionText.includes(query);
            option.classList.toggle('search-hidden', !matches);
            if (matches) {
                visibleCount += 1;
            }
        });

        let emptyState = optionsContainer.parentElement.querySelector('.slicer-search-empty');
        if (!emptyState) {
            emptyState = document.createElement('div');
            emptyState.className = 'slicer-empty slicer-search-empty hidden';
            emptyState.textContent = 'No matching values';
            optionsContainer.insertAdjacentElement('afterend', emptyState);
        }

        emptyState.classList.toggle('hidden', visibleCount !== 0 || !query);
    }
    
    selectAllSlicerOptions(fieldName, selectAll) {
        const slicer = document.querySelector(`[data-field-name="${fieldName}"]`);
        if (!slicer) return;
        
        const checkboxes = slicer.querySelectorAll('input[type="checkbox"]');
        checkboxes.forEach(checkbox => {
            checkbox.checked = selectAll;
        });
        
        this.handleSlicerChange(fieldName);
    }
    
    handleSlicerChange(fieldName) {
        const slicer = document.querySelector(`[data-field-name="${fieldName}"]`);
        if (!slicer) return;
        
        const checkboxes = slicer.querySelectorAll('input[type="checkbox"]');
        const selectedValues = [];
        
        checkboxes.forEach(checkbox => {
            if (checkbox.checked) {
                selectedValues.push(checkbox.value);
            }
        });
        
        // Update the filter (map _none to '' for empty-value matching)
        const filterValues = selectedValues.map(v => v === '_none' ? '' : v);
        this.app.setFilter(fieldName, filterValues);
        
        // Update view configuration
        this.app.updateViewConfiguration();
        this.renderSlicers();
    }

    applyExclusiveSlicerFilter(fieldName, rawValue) {
        const filterValue = rawValue === '_none' ? '' : rawValue;
        this.app.setFilter(fieldName, [filterValue]);
        this.app.updateViewConfiguration();
        this.renderSlicers();
        this.app.showNotification(`Filtered ${fieldName} to ${filterValue === '' ? '(none)' : filterValue}`, 'success');
    }

    openFieldValueEditor(fieldName, contextValue = null) {
        if (!fieldName || !this.app.detailsPanelManager) {
            return false;
        }

        if (typeof this.app.isTagFieldName === 'function' && this.app.isTagFieldName(fieldName)) {
            return false;
        }

        const headerValue = contextValue === '_none' ? '' : contextValue;
        this.app.detailsPanelManager.openForHeader(fieldName, 'filter', headerValue);
        return true;
    }

    createGroupSlicer(container) {
        const groups = this.app.getGroups();
        const activeFilter = this.app.viewConfig.groupSlicerFilter;
        const hasActiveFilter = Array.isArray(activeFilter) && activeFilter.length > 0;
        const isExpanded = this.openSlicerField === '_groups';

        const slicer = document.createElement('div');
        slicer.className = 'slicer' + (isExpanded ? ' is-open' : '');
        slicer.dataset.fieldName = '_groups';

        const toggle = document.createElement('button');
        toggle.type = 'button';
        toggle.className = 'slicer-toggle';
        toggle.setAttribute('aria-expanded', String(isExpanded));

        const title = document.createElement('span');
        title.className = 'slicer-title';
        title.textContent = 'Groups';

        const summary = document.createElement('span');
        summary.className = 'slicer-summary';
        const totalOptions = groups.length + 1; // +1 for (none)
        if (!hasActiveFilter) {
            summary.textContent = 'All';
        } else if (activeFilter.length === 1) {
            if (activeFilter[0] === '_none') {
                summary.textContent = '(none)';
            } else {
                const g = groups.find(gr => gr.id === activeFilter[0]);
                summary.textContent = g ? g.name : '1 selected';
            }
        } else {
            summary.textContent = `${activeFilter.length} of ${totalOptions}`;
        }

        const toggleText = document.createElement('span');
        toggleText.className = 'slicer-toggle-text';
        toggleText.appendChild(title);
        toggleText.appendChild(summary);
        toggle.appendChild(toggleText);

        const toggleIcon = document.createElement('span');
        toggleIcon.className = 'slicer-toggle-icon';
        toggleIcon.setAttribute('aria-hidden', 'true');
        toggleIcon.textContent = '▾';
        toggle.appendChild(toggleIcon);

        slicer.appendChild(toggle);

        const panel = document.createElement('div');
        panel.className = 'slicer-panel';

        const optionsContainer = document.createElement('div');
        optionsContainer.className = 'slicer-options';

        // Add All/None controls
        const controlsDiv = document.createElement('div');
        controlsDiv.className = 'slicer-controls';

        const selectAllBtn = document.createElement('button');
        selectAllBtn.type = 'button';
        selectAllBtn.textContent = 'All';
        selectAllBtn.className = 'slicer-control-btn';

        const clearAllBtn = document.createElement('button');
        clearAllBtn.type = 'button';
        clearAllBtn.textContent = 'None';
        clearAllBtn.className = 'slicer-control-btn';

        selectAllBtn.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            this.app.viewConfig.groupSlicerFilter = null;
            this.app.updateViewConfiguration();
            this.app.updateFilteredData();
            this.renderSlicers();
            this.app.renderGrid();
        });

        clearAllBtn.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            this.app.viewConfig.groupSlicerFilter = [];
            this.app.updateViewConfiguration();
            this.app.updateFilteredData();
            this.renderSlicers();
            this.app.renderGrid();
        });

        controlsDiv.appendChild(selectAllBtn);
        controlsDiv.appendChild(clearAllBtn);
        optionsContainer.appendChild(controlsDiv);

        // Add (none) option for ungrouped items
        const noneOption = document.createElement('div');
        noneOption.className = 'slicer-option';
        noneOption.dataset.searchText = '(none)';

        const noneCheckbox = document.createElement('input');
        noneCheckbox.type = 'checkbox';
        noneCheckbox.id = '_groups-_none';
        noneCheckbox.value = '_none';
        noneCheckbox.checked = !hasActiveFilter || activeFilter.includes('_none');

        const noneLabel = document.createElement('label');
        noneLabel.htmlFor = noneCheckbox.id;

        const noneSwatch = document.createElement('span');
        noneSwatch.style.display = 'inline-block';
        noneSwatch.style.width = '10px';
        noneSwatch.style.height = '10px';
        noneSwatch.style.borderRadius = '2px';
        noneSwatch.style.backgroundColor = '#ccc';
        noneSwatch.style.marginRight = '6px';
        noneSwatch.style.verticalAlign = 'middle';

        noneLabel.appendChild(noneSwatch);
        noneLabel.appendChild(document.createTextNode('(none)'));

        noneOption.appendChild(noneCheckbox);
        noneOption.appendChild(noneLabel);
        optionsContainer.appendChild(noneOption);

        noneCheckbox.addEventListener('change', () => {
            this.handleGroupSlicerChange(slicer);
        });

        noneOption.addEventListener('contextmenu', (event) => {
            event.preventDefault();
            event.stopPropagation();
            this.applyExclusiveGroupSlicerFilter('_none');
        });

        groups.forEach(group => {
            const option = document.createElement('div');
            option.className = 'slicer-option';
            option.dataset.searchText = group.name;

            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.id = `_groups-${group.id}`;
            checkbox.value = group.id;
            checkbox.checked = !hasActiveFilter || activeFilter.includes(group.id);

            const label = document.createElement('label');
            label.htmlFor = checkbox.id;

            const swatch = document.createElement('span');
            swatch.style.display = 'inline-block';
            swatch.style.width = '10px';
            swatch.style.height = '10px';
            swatch.style.borderRadius = '2px';
            swatch.style.backgroundColor = group.color;
            swatch.style.marginRight = '6px';
            swatch.style.verticalAlign = 'middle';

            label.appendChild(swatch);
            label.appendChild(document.createTextNode(group.name));

            option.appendChild(checkbox);
            option.appendChild(label);
            optionsContainer.appendChild(option);

            checkbox.addEventListener('change', () => {
                this.handleGroupSlicerChange(slicer);
            });

            option.addEventListener('contextmenu', (event) => {
                event.preventDefault();
                event.stopPropagation();
                this.applyExclusiveGroupSlicerFilter(group.id);
            });
        });

        panel.appendChild(optionsContainer);
        slicer.appendChild(panel);

    this.addSlicerSearch(panel, '_groups', optionsContainer);

        this.applySlicerSearchFilter('_groups', optionsContainer);

        toggle.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            this.toggleSlicer('_groups');
        });

        container.appendChild(slicer);
    }

    handleGroupSlicerChange(slicerEl) {
        const checkboxes = slicerEl.querySelectorAll('input[type="checkbox"]');
        const allChecked = Array.from(checkboxes).every(cb => cb.checked);

        if (allChecked) {
            this.app.viewConfig.groupSlicerFilter = null;
        } else {
            const selected = [];
            checkboxes.forEach(cb => {
                if (cb.checked) selected.push(cb.value);
            });
            this.app.viewConfig.groupSlicerFilter = selected;
        }

        this.app.updateViewConfiguration();
        this.app.updateFilteredData();
        this.renderSlicers();
        this.app.renderGrid();
    }

    applyExclusiveGroupSlicerFilter(groupId) {
        this.app.viewConfig.groupSlicerFilter = [groupId];
        this.app.updateViewConfiguration();
        this.app.updateFilteredData();
        this.renderSlicers();
        this.app.renderGrid();
        this.app.showNotification(`Filtered Groups to ${groupId === '_none' ? '(none)' : groupId}`, 'success');
    }
    

    

    
    async exportFilteredData() {
        try {
            const filteredItems = this.app.getFilteredData();
            if (!filteredItems || filteredItems.length === 0) {
                this.app.showNotification('No filtered data to export', 'warning');
                return;
            }
            
            const exportData = {
                data: filteredItems,
                view: this.app.viewConfig,
                meta: {
                    ...this.app.metaInfo,
                    lastModified: window.GridDateUtils.createLocalTimestamp(),
                    exportedBy: 'Data Visualization Grid',
                    note: `Filtered export: ${filteredItems.length} of ${this.app.dataset.length} items`
                }
            };
            
            const filename = this.app.persistenceManager.fileService.generateDataFilename(filteredItems.length);
            const savedFilename = await this.app.persistenceManager.fileService.saveFile(
                exportData, 
                filename.replace('data-export', 'filtered-export'), 
                'data'
            );
            
            if (savedFilename) {
                this.app.showNotification(`Exported ${filteredItems.length} filtered items`, 'success');
            }
            
        } catch (error) {
            console.error('Error exporting filtered data:', error);
            this.app.showNotification('Failed to export filtered data: ' + error.message, 'error');
        }
    }
    
    // ===== VIEW CONFIGURATION MANAGEMENT =====
    
    async handleViewFileLoad(file) {
        if (!file) return;
        
        try {
            const text = await file.text();
            const viewConfig = JSON.parse(text);
            
            // Validate view configuration structure
            if (viewConfig.view) {
                this.app.viewConfig = viewConfig.view;
            } else {
                this.app.viewConfig = viewConfig;
            }
            
            this.app.applyViewConfiguration();
            this.app.updateFilteredData();
            
            // Re-render UI components
            this.renderFieldSelectors();
            this.renderSlicers();
            this.app.renderTags();
            this.app.renderGroups();
            this.app.renderGrid();
            
            this.app.showNotification('View configuration loaded successfully', 'success');
        } catch (error) {
            console.error('Error loading view file:', error);
            this.app.showNotification('Failed to load view configuration: ' + error.message, 'error');
        }
    }

    updateFileActionState(status = {}) {
        const reloadButton = document.getElementById('reload-data-btn');
        const saveDataButton = document.getElementById('save-data-btn');
        const promoteDataButton = document.getElementById('promote-data-btn');
        const saveViewButton = document.getElementById('save-view-btn');

        if (reloadButton) {
            const reloadAction = status.reloadDataAction || 'unavailable';
            const canReload = !!status.canReloadData;

            reloadButton.disabled = !canReload;

            if (reloadAction === 'reopen') {
                reloadButton.title = 'Reload the last opened dataset directly from disk';
            } else if (reloadAction === 'reselect') {
                reloadButton.title = 'Choose the dataset again to reload it in this browser';
            } else {
                reloadButton.title = 'Load a dataset first to enable reload';
            }
        }

        if (saveDataButton) {
            saveDataButton.title = status.supportsSavePicker
                ? 'Save to the current file or choose a file with the browser save picker'
                : 'Download the current dataset as a JSON file';
        }

        if (promoteDataButton) {
            promoteDataButton.title = status.supportsSavePicker
                ? 'Save the effective dataset as baseline data and clear pending changes'
                : 'Download the effective dataset as baseline data and clear pending changes';
        }

        if (saveViewButton) {
            saveViewButton.title = status.supportsSavePicker
                ? 'Save to the current view file or choose a file with the browser save picker'
                : 'Download the current view configuration as a JSON file';
        }
    }
    

}

// ===== TAG MANAGEMENT METHODS =====

// Add tag management methods to ControlPanelManager
ControlPanelManager.prototype.renderTags = function() {
    const tagsContainer = document.getElementById('tags-container');
    tagsContainer.innerHTML = '';
    
    const allTags = this.app.getAllTags();
    
    if (allTags.length === 0) {
        const placeholder = document.createElement('div');
        placeholder.className = 'tags-placeholder';
        placeholder.textContent = 'No tags found in dataset';
        placeholder.style.color = '#999';
        placeholder.style.fontStyle = 'italic';
        tagsContainer.appendChild(placeholder);
        return;
    }
    
    allTags.forEach(tagName => {
        this.createTagElement(tagName, tagsContainer);
    });

    this.app.selectedControlTags = new Set(
        this.app.getSelectedControlTags().filter((tagName) => allTags.includes(tagName))
    );
    this.app.syncSelectedControlTagSelection();
};

ControlPanelManager.prototype.createTagElement = function(tagName, container) {
    const template = document.getElementById('tag-template');
    const tagElement = template.content.cloneNode(true);
    
    const tag = tagElement.querySelector('.tag');
    const label = tagElement.querySelector('.tag-label');
    
    const tagConfig = this.app.getTagConfig(tagName);
    let suppressClickAfterDrag = false;
    
    tag.dataset.tagName = tagName;
    tag.style.backgroundColor = tagConfig.color;
    label.textContent = tagConfig.label;
    tag.title = 'Click to edit tag';
    tag.classList.toggle('is-selected', this.app.isControlTagSelected(tagName));
    
    // Add drag functionality
    tag.addEventListener('dragstart', (e) => {
        e.dataTransfer.setData('application/x-tag-name', tagName);
        e.dataTransfer.setData('text/plain', `tag:${tagName}`);
        e.dataTransfer.effectAllowed = 'copy';
        this.draggedTagName = tagName;
        suppressClickAfterDrag = true;
        document.body.classList.add('tag-drag-mode');
        tag.style.opacity = '0.5';
    });
    
    tag.addEventListener('dragend', (e) => {
        this.draggedTagName = null;
        document.body.classList.remove('tag-drag-mode');
        document.querySelectorAll('.tag-drop-target-active').forEach((element) => {
            element.classList.remove('tag-drop-target-active');
        });
        tag.style.opacity = '1';
        setTimeout(() => {
            suppressClickAfterDrag = false;
        }, 0);
    });
    
    tag.addEventListener('click', (e) => {
        if (suppressClickAfterDrag) {
            return;
        }

        if (e.shiftKey) {
            e.preventDefault();
            e.stopPropagation();
            this.app.setSelectedControlTag(tagName, { additive: true });
            return;
        }

        this.openTagEditDialog(tagName, tagConfig);
    });

    tag.addEventListener('contextmenu', (e) => {
        e.preventDefault();
        e.stopPropagation();
        this.applySelectedTagsFilter(tagName);
    });
    
    container.appendChild(tag);
};

ControlPanelManager.prototype.applySelectedTagsFilter = function(fallbackTagName = null) {
    const tagFieldName = this.app.getTagFieldName();
    if (!tagFieldName) {
        this.app.showNotification('No Tags field is available for filtering', 'warning');
        return;
    }

    const selectedTags = this.app.getSelectedControlTags();
    const nextTags = selectedTags.length > 0
        ? selectedTags
        : (fallbackTagName ? [fallbackTagName] : []);

    if (nextTags.length === 0) {
        this.app.showNotification('Select one or more tags first', 'warning');
        return;
    }

    if (selectedTags.length === 0 && fallbackTagName) {
        this.app.setSelectedControlTag(fallbackTagName);
    }

    this.app.setFilter(tagFieldName, nextTags);
    this.app.updateViewConfiguration();
    this.renderSlicers();
    this.app.showNotification(`Filtered ${tagFieldName} to ${nextTags.length} tag${nextTags.length === 1 ? '' : 's'}`, 'success');
};

ControlPanelManager.prototype.initializeTagEditDialog = function() {
    const dialog = document.getElementById('tag-edit-dialog');
    const form = document.getElementById('tag-edit-form');
    const cancelBtn = document.getElementById('tag-edit-cancel-btn');
    const closeBtn = dialog.querySelector('.modal-close');
    
    // Form submit handler
    form.addEventListener('submit', (e) => {
        e.preventDefault();
        this.saveTagEdit();
    });
    
    // Cancel and close handlers
    const closeDialog = () => {
        dialog.classList.add('hidden');
        this.currentEditingTag = null;
        this.tagDialogMode = 'edit';
        form.reset();
    };
    
    cancelBtn.addEventListener('click', closeDialog);
    closeBtn.addEventListener('click', closeDialog);
    
    // Close on background click
    dialog.addEventListener('click', (e) => {
        if (e.target === dialog) {
            closeDialog();
        }
    });
};

ControlPanelManager.prototype.openTagEditDialog = function(tagName, currentConfig) {
    this.currentEditingTag = tagName;
    this.tagDialogMode = 'edit';
    
    const dialog = document.getElementById('tag-edit-dialog');
    const dialogTitle = document.getElementById('tag-edit-dialog-title');
    const labelInput = document.getElementById('tag-label-input');
    const colorInput = document.getElementById('tag-color-input');
    
    dialogTitle.textContent = 'Edit Tag';
    labelInput.value = currentConfig.label;
    colorInput.value = currentConfig.color;
    
    dialog.classList.remove('hidden');
    labelInput.focus();
};

ControlPanelManager.prototype.openNewTagDialog = function() {
    this.currentEditingTag = null;
    this.tagDialogMode = 'create';

    const dialog = document.getElementById('tag-edit-dialog');
    const dialogTitle = document.getElementById('tag-edit-dialog-title');
    const labelInput = document.getElementById('tag-label-input');
    const colorInput = document.getElementById('tag-color-input');

    dialogTitle.textContent = 'Add Tag';
    labelInput.value = '';
    colorInput.value = '#007acc';

    dialog.classList.remove('hidden');
    labelInput.focus();
};

ControlPanelManager.prototype.saveTagEdit = function() {
    const isCreateMode = this.tagDialogMode === 'create';
    const labelInput = document.getElementById('tag-label-input');
    const colorInput = document.getElementById('tag-color-input');
    const rawLabel = labelInput.value.trim();

    if (!rawLabel) {
        this.app.showNotification('Tag label is required', 'warning');
        labelInput.focus();
        return;
    }

    const tagName = isCreateMode
        ? rawLabel
        : this.currentEditingTag;

    if (!tagName) {
        return;
    }

    const existingTags = new Set(this.app.getAllTags().map((tag) => String(tag).toLowerCase()));
    if (isCreateMode && existingTags.has(tagName.toLowerCase())) {
        this.app.showNotification('A tag with that label already exists', 'warning');
        labelInput.focus();
        return;
    }
    
    const newConfig = {
        label: rawLabel,
        color: colorInput.value
    };
    
    this.app.updateTagConfig(tagName, newConfig);
    
    const dialog = document.getElementById('tag-edit-dialog');
    dialog.classList.add('hidden');
    this.currentEditingTag = null;
    this.tagDialogMode = 'edit';
    
    this.app.showNotification(
        isCreateMode ? 'Tag added successfully' : 'Tag updated successfully',
        'success'
    );
};

ControlPanelManager.prototype.initializeTagDragDrop = function() {
    // Native select elements are inconsistent drop targets, so use the selector group wrapper.
    const topLeftSelect = document.getElementById('top-left-select');
    const topRightSelect = document.getElementById('top-right-select');
    
    [topLeftSelect, topRightSelect].forEach(selector => {
        const dropTarget = selector ? selector.closest('.selector-group') : null;
        if (dropTarget) {
            this.makeDropTarget(dropTarget, selector.id);
        }
    });
};

ControlPanelManager.prototype.makeDropTarget = function(element, selectorId) {
    element.classList.add('tag-drop-target');

    const isTagDrag = (event) => {
        const types = event.dataTransfer ? Array.from(event.dataTransfer.types || []) : [];
        return this.draggedTagName !== null ||
            types.includes('application/x-tag-name') ||
            types.includes('text/plain');
    };

    element.addEventListener('dragover', (e) => {
        if (!isTagDrag(e)) {
            return;
        }
        e.preventDefault();
        e.stopPropagation();
        e.dataTransfer.dropEffect = 'copy';
        element.classList.add('tag-drop-target-active');
    });
    
    element.addEventListener('dragleave', (e) => {
        if (!isTagDrag(e)) {
            return;
        }
        if (!element.contains(e.relatedTarget)) {
            element.classList.remove('tag-drop-target-active');
        }
    });
    
    element.addEventListener('drop', (e) => {
        if (!isTagDrag(e)) {
            return;
        }
        e.preventDefault();
        e.stopPropagation();
        element.classList.remove('tag-drop-target-active');
        document.body.classList.remove('tag-drag-mode');
        
        const tagName = e.dataTransfer.getData('application/x-tag-name');
        const fallbackTagName = e.dataTransfer.getData('text/plain').startsWith('tag:')
            ? e.dataTransfer.getData('text/plain').slice(4)
            : '';
        const resolvedTagName = tagName || this.draggedTagName || fallbackTagName;
        if (resolvedTagName) {
            this.assignTagToSelector(selectorId, resolvedTagName);
        }

        this.draggedTagName = null;
    });
};

ControlPanelManager.prototype.assignTagToSelector = function(selectorId, tagName) {
    const tagFieldName = typeof this.app.getTagFieldName === 'function'
        ? this.app.getTagFieldName()
        : (this.app.availableFields.includes('tags') ? 'tags' : null);
    const selectorElement = document.getElementById(selectorId);
    
    if (!tagFieldName || !selectorElement) {
        this.app.showNotification('No tags field available in dataset', 'error');
        return;
    }
    
    // Set the selector to use the tags field
    selectorElement.value = tagFieldName;
    
    // Update the axis selection
    if (selectorElement.id === 'top-left-select') {
        this.app.currentCardSelections.topLeft = tagFieldName;
    } else if (selectorElement.id === 'top-right-select') {
        this.app.currentCardSelections.topRight = tagFieldName;
    }
    
    this.app.updateViewConfiguration();
    this.app.renderFieldSelectors();
    this.renderTableColumnSelector();
    this.app.renderGrid();
    this.updateSlicerAxisShading();
    
    this.app.showNotification(`Assigned tags field to ${selectorElement.id.replace('-select', '')} label`, 'success');
};

ControlPanelManager.prototype.updateTagDropTargetSelection = function() {
    const tagFieldName = typeof this.app.getTagFieldName === 'function'
        ? this.app.getTagFieldName()
        : 'tags';
    ['top-left-select', 'top-right-select'].forEach(selectorId => {
        const selectorElement = document.getElementById(selectorId);
        const dropTarget = selectorElement ? selectorElement.closest('.selector-group') : null;

        if (!dropTarget || !selectorElement) {
            return;
        }

        dropTarget.classList.toggle('tag-drop-target-selected', selectorElement.value === tagFieldName);
    });
};

// ===== GROUP MANAGEMENT METHODS =====

ControlPanelManager.prototype.renderGroups = function() {
    const container = document.getElementById('groups-container');
    if (!container) return;
    container.innerHTML = '';

    const groups = this.app.getGroups();

    if (groups.length === 0) {
        const placeholder = document.createElement('div');
        placeholder.className = 'groups-placeholder';
        placeholder.textContent = 'No groups defined';
        container.appendChild(placeholder);
        return;
    }

    groups.forEach(group => {
        const el = document.createElement('div');
        el.className = 'group-item' + (group.enabled === false ? ' disabled' : '');
        el.dataset.groupId = group.id;

        const swatch = document.createElement('span');
        swatch.className = 'group-item-swatch';
        swatch.style.backgroundColor = group.color || '#3498db';

        const name = document.createElement('span');
        name.className = 'group-item-name';
        name.textContent = group.name || '(unnamed)';

        // Show member count
        const count = document.createElement('span');
        count.className = 'group-item-count';
        let memberCount = 0;
        if (this.app.groupMemberships) {
            this.app.groupMemberships.forEach(memberGroups => {
                if (memberGroups.some(g => g.id === group.id)) memberCount++;
            });
        }
        count.textContent = memberCount > 0 ? `(${memberCount})` : '';

        const toggle = document.createElement('input');
        toggle.type = 'checkbox';
        toggle.className = 'group-item-toggle';
        toggle.checked = group.enabled !== false;
        toggle.title = group.enabled !== false ? 'Disable group' : 'Enable group';

        toggle.addEventListener('click', (e) => {
            e.stopPropagation();
            this.app.toggleGroup(group.id);
            this.renderGroups();
            this.renderSlicers();
            this.app.renderGrid();
        });

        let suppressClickAfterDrag = false;

        el.addEventListener('click', () => {
            if (suppressClickAfterDrag) return;
            this.openGroupEditDialog(group);
        });

        // Drag-and-drop: drag group label onto a card to add it as a manual member
        el.draggable = true;
        el.addEventListener('dragstart', (e) => {
            e.dataTransfer.setData('application/x-group-id', group.id);
            e.dataTransfer.setData('text/plain', `group:${group.id}`);
            e.dataTransfer.effectAllowed = 'copy';
            this.draggedGroupId = group.id;
            suppressClickAfterDrag = true;
            document.body.classList.add('group-drag-mode');
            el.style.opacity = '0.5';
        });

        el.addEventListener('dragend', () => {
            this.draggedGroupId = null;
            document.body.classList.remove('group-drag-mode');
            document.querySelectorAll('.group-drop-target-active').forEach(element => {
                element.classList.remove('group-drop-target-active');
            });
            el.style.opacity = '1';
            setTimeout(() => { suppressClickAfterDrag = false; }, 0);
        });

        el.appendChild(swatch);
        el.appendChild(name);
        el.appendChild(count);
        el.appendChild(toggle);
        container.appendChild(el);
    });
};

ControlPanelManager.prototype.initializeGroupEditDialog = function() {
    const dialog = document.getElementById('group-edit-dialog');
    if (!dialog) return;

    const form = document.getElementById('group-edit-form');
    const cancelBtn = document.getElementById('group-edit-cancel-btn');
    const closeBtn = dialog.querySelector('.modal-close');
    const deleteBtn = document.getElementById('group-edit-delete-btn');
    const operatorSelect = document.getElementById('group-rule-operator');
    const valuesRow = document.getElementById('group-rule-values-row');

    const closeDialog = () => {
        dialog.classList.add('hidden');
        this.currentEditingGroupId = null;
        form.reset();
    };

    form.addEventListener('submit', (e) => {
        e.preventDefault();
        this.saveGroupEdit();
    });

    cancelBtn.addEventListener('click', closeDialog);
    closeBtn.addEventListener('click', closeDialog);
    dialog.addEventListener('click', (e) => {
        if (e.target === dialog) closeDialog();
    });

    deleteBtn.addEventListener('click', () => {
        if (this.currentEditingGroupId) {
            this.app.removeGroup(this.currentEditingGroupId);
            closeDialog();
            this.renderGroups();
            this.renderSlicers();
            this.app.renderGrid();
        }
    });

    // Hide values row for is-empty / is-not-empty operators
    operatorSelect.addEventListener('change', () => {
        const op = operatorSelect.value;
        valuesRow.style.display = (op === 'is-empty' || op === 'is-not-empty') ? 'none' : '';
    });
};

ControlPanelManager.prototype.populateGroupRuleFieldOptions = function(preferredField = null) {
    const fieldSelect = document.getElementById('group-rule-field');
    if (!fieldSelect) return;

    // Preserve current selection
    const current = preferredField || fieldSelect.value;
    fieldSelect.innerHTML = '<option value="">(none)</option>';

    const addedFields = new Set(['']);
    const appendFieldOption = (fieldName, label = fieldName) => {
        if (typeof fieldName !== 'string') {
            return;
        }

        const normalizedFieldName = fieldName.trim();
        if (!normalizedFieldName || addedFields.has(normalizedFieldName)) {
            return;
        }

        const option = document.createElement('option');
        option.value = normalizedFieldName;
        option.textContent = label;
        fieldSelect.appendChild(option);
        addedFields.add(normalizedFieldName);
    };

    // Add _hasChanges virtual field
    appendFieldOption('_hasChanges', '(has changes)');

    // Group rules can target any non-structured field, even if it is hidden from selectors.
    const typedFields = this.app.fieldTypes instanceof Map
        ? Array.from(this.app.fieldTypes.entries())
            .filter(([, fieldType]) => fieldType === 'scalar' || fieldType === 'multi-value')
            .map(([fieldName]) => fieldName)
        : [];
    const fields = Array.from(new Set([...(this.app.availableFields || []), ...typedFields]));
    fields.forEach((fieldName) => {
        appendFieldOption(fieldName);
    });

    // Keep loading/editing robust for older views whose rule field is not in the current dataset metadata.
    appendFieldOption(current);

    // Restore selection if still valid
    if (current && fieldSelect.querySelector(`option[value="${CSS.escape(current)}"]`)) {
        fieldSelect.value = current;
    }
};

ControlPanelManager.prototype.openGroupEditDialog = function(group) {
    this.currentEditingGroupId = group ? group.id : null;

    const dialog = document.getElementById('group-edit-dialog');
    const title = document.getElementById('group-edit-dialog-title');
    const nameInput = document.getElementById('group-name-input');
    const colorInput = document.getElementById('group-color-input');
    const fieldSelect = document.getElementById('group-rule-field');
    const operatorSelect = document.getElementById('group-rule-operator');
    const valuesInput = document.getElementById('group-rule-values');
    const valuesRow = document.getElementById('group-rule-values-row');
    const deleteBtn = document.getElementById('group-edit-delete-btn');

    this.populateGroupRuleFieldOptions(group && group.rule ? group.rule.field : null);

    if (group) {
        title.textContent = 'Edit Group';
        nameInput.value = group.name || '';
        colorInput.value = group.color || '#3498db';
        deleteBtn.classList.remove('hidden');

        if (group.rule && group.rule.field) {
            fieldSelect.value = group.rule.field;
            operatorSelect.value = group.rule.operator || 'equals';
            valuesInput.value = Array.isArray(group.rule.values) ? group.rule.values.join(', ') : '';
        } else {
            fieldSelect.value = '';
            operatorSelect.value = 'equals';
            valuesInput.value = '';
        }
    } else {
        title.textContent = 'Add Group';
        nameInput.value = '';
        colorInput.value = '#3498db';
        fieldSelect.value = '';
        operatorSelect.value = 'equals';
        valuesInput.value = '';
        deleteBtn.classList.add('hidden');
    }

    const op = operatorSelect.value;
    valuesRow.style.display = (op === 'is-empty' || op === 'is-not-empty') ? 'none' : '';

    dialog.classList.remove('hidden');
    nameInput.focus();
};

ControlPanelManager.prototype.saveGroupEdit = function() {
    const nameInput = document.getElementById('group-name-input');
    const colorInput = document.getElementById('group-color-input');
    const fieldSelect = document.getElementById('group-rule-field');
    const operatorSelect = document.getElementById('group-rule-operator');
    const valuesInput = document.getElementById('group-rule-values');

    const name = nameInput.value.trim();
    if (!name) {
        this.app.showNotification('Group name is required', 'warning');
        nameInput.focus();
        return;
    }

    const color = colorInput.value;
    const ruleField = fieldSelect.value;
    let rule = null;

    if (ruleField) {
        const operator = operatorSelect.value;
        const rawValues = valuesInput.value.trim();
        const values = (operator === 'is-empty' || operator === 'is-not-empty')
            ? []
            : rawValues.split(',').map(v => v.trim()).filter(Boolean);
        rule = { field: ruleField, operator, values };
    }

    if (this.currentEditingGroupId) {
        this.app.updateGroup(this.currentEditingGroupId, { name, color, rule });
        this.app.showNotification('Group updated', 'success');
    } else {
        this.app.addGroup({ name, color, rule, enabled: true, manualMembers: [] });
        this.app.showNotification('Group added', 'success');
    }

    const dialog = document.getElementById('group-edit-dialog');
    dialog.classList.add('hidden');
    this.currentEditingGroupId = null;

    this.renderGroups();
    this.renderSlicers();
    this.app.renderGrid();
};

// ===== INTEGRATION WITH MAIN APP =====

function attachControlPanelIntegration(targetApp) {
    if (!targetApp || targetApp.__controlPanelIntegrated) {
        return;
    }

    targetApp.initializeEventListeners = function() {
        if (!this.controlPanel) {
            this.controlPanel = new ControlPanelManager(this);
        }
        return this.controlPanel;
    };

    targetApp.renderFieldSelectors = function() {
        if (this.controlPanel) {
            this.controlPanel.renderFieldSelectors();
        }
    };

    targetApp.renderSlicers = function() {
        if (this.controlPanel) {
            this.controlPanel.renderSlicers();
        }
    };

    targetApp.renderTags = function() {
        if (this.controlPanel) {
            this.controlPanel.renderTags();
        }
    };

    targetApp.renderGroups = function() {
        if (this.controlPanel) {
            this.controlPanel.renderGroups();
        }
    };

    targetApp.updateFileActionState = function(status) {
        if (this.controlPanel) {
            this.controlPanel.updateFileActionState(status);
        }
    };

    targetApp.showNotification = function(message, type = 'info') {
        console.log(`[${type.toUpperCase()}] ${message}`);

        const toast = document.createElement('div');
        toast.className = `notification notification-${type}`;
        toast.textContent = message;
        toast.style.cssText = `
            position: fixed;
            top: 20px;
            right: 20px;
            padding: 12px 20px;
            background: ${type === 'error' ? '#e74c3c' : type === 'success' ? '#27ae60' : '#3498db'};
            color: white;
            border-radius: 4px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.2);
            z-index: 9999;
            font-size: 14px;
            max-width: 300px;
        `;

        document.body.appendChild(toast);

        setTimeout(() => {
            if (toast.parentNode) {
                toast.parentNode.removeChild(toast);
            }
        }, 3000);
    };

    targetApp.__controlPanelIntegrated = true;
}

document.addEventListener('DOMContentLoaded', () => {
    if (typeof app !== 'undefined' && app) {
        attachControlPanelIntegration(app);
        app.initializeEventListeners();
        app.renderFieldSelectors();
        app.renderSlicers();
        app.renderTags();
        app.renderGroups();
        if (app.persistenceManager) {
            app.updateFileActionState(app.persistenceManager.getFileServiceStatus());
        }
    }
});