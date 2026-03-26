// Copyright (c) 2026 Malte Doerper. MIT License. See LICENSE file.

// V5 — Relations UI Manager
// Manages card badges, side panel, derived fields, focus mode, and overlay
class RelationUIManager {
    constructor(app) {
        this.app = app;

        this.MAX_FOCUS_DEPTH = 4;

        // Badge icon mapping
        this.BADGE_ICONS = {
            parent: '⬆',
            child: '⬇',
            blocks: '⛔',
            blockedBy: '🚫',
            relatesTo: '🔗',
            epic: '📦',
            subtask: '📎',
            duplicates: '♊'
        };

        // Badge display labels
        this.BADGE_LABELS = {
            parent: 'parent',
            child: 'children',
            blocks: 'blocks',
            blockedBy: 'blocked by',
            relatesTo: 'related',
            epic: 'epic',
            subtask: 'subtasks',
            duplicates: 'duplicates'
        };

        // Max visible badge types before overflow
        this.MAX_VISIBLE_BADGES = 4;

        // Focus mode state
        this.focusMode = false;
        this.focusRootId = null;
        this.focusDepth = 1;
        this.focusTypes = null; // null = all types
        this.focusSet = new Set();

        // Side panel state
        this.panelItem = null;
        this.panelFilterType = null;

        this.initializePanel();
    }

    // ===== INITIALIZATION =====

    initializePanel() {
        // Relations content now lives inside the multi-mode details panel
        const panel = document.getElementById('details-panel');
        if (!panel) return;

        // Add-relation toggle
        const addBtn = document.getElementById('relation-add-btn');
        if (addBtn) {
            addBtn.addEventListener('click', () => this.toggleAddRelationForm());
        }

        // Add-relation submit
        const submitBtn = document.getElementById('relation-add-submit');
        if (submitBtn) {
            submitBtn.addEventListener('click', () => this.submitAddRelation());
        }

        // Add-relation cancel
        const cancelBtn = document.getElementById('relation-add-cancel');
        if (cancelBtn) {
            cancelBtn.addEventListener('click', () => this.hideAddRelationForm());
        }

        // Target search input
        const searchInput = document.getElementById('relation-target-search');
        if (searchInput) {
            searchInput.addEventListener('input', () => this.updateTargetSearchResults());
        }

        // Focus controls
        this.initializeFocusControls();
    }

    initializeFocusControls() {
        const exitBtn = document.getElementById('focus-exit-btn');
        if (exitBtn) {
            exitBtn.addEventListener('click', () => this.exitFocusMode());
        }

        const depthSelect = document.getElementById('focus-depth-select');
        if (depthSelect) {
            depthSelect.addEventListener('change', () => {
                this.focusDepth = this.normalizeFocusDepth(depthSelect.value);
                if (this.focusMode && this.focusRootId) {
                    this.recomputeFocus();
                }
            });
        }

        const typeCheckboxes = document.querySelectorAll('.focus-type-checkbox');
        typeCheckboxes.forEach((cb) => {
            cb.addEventListener('change', () => {
                this.focusTypes = this.getSelectedFocusTypes();
                if (this.focusMode && this.focusRootId) {
                    this.recomputeFocus();
                }
            });
        });
    }

    // ===== CARD BADGES (§V5.3) =====

    computeRelationBadges(item) {
        const badges = new Map();
        const relations = item && Array.isArray(item.relations) ? item.relations : [];

        relations.forEach((rel) => {
            if (!rel || !rel.type) return;
            if (!badges.has(rel.type)) {
                badges.set(rel.type, { count: 0, items: [] });
            }
            const entry = badges.get(rel.type);
            entry.count++;
            entry.items.push(rel);
        });

        return badges;
    }

    renderBadges(cardElement, item) {
        const badges = this.computeRelationBadges(item);
        if (badges.size === 0) return;

        const container = document.createElement('div');
        container.className = 'card-badges';

        const types = Array.from(badges.keys());
        const visibleTypes = types.slice(0, this.MAX_VISIBLE_BADGES);
        const overflowCount = types.length - visibleTypes.length;

        visibleTypes.forEach((type) => {
            const entry = badges.get(type);
            const icon = this.BADGE_ICONS[type] || '🔗';
            const label = this.BADGE_LABELS[type] || type;

            const badge = document.createElement('span');
            badge.className = 'relation-badge';
            badge.dataset.relationType = type;
            badge.textContent = `${icon}\u00A0${entry.count}`;
            badge.title = `${entry.count} ${label}`;

            // Hover tooltip with item list
            badge.addEventListener('mouseenter', (e) => {
                this.showBadgeTooltip(e, entry, type);
            });
            badge.addEventListener('mouseleave', () => {
                this.hideBadgeTooltip();
            });

            // Click opens side panel filtered to this type
            badge.addEventListener('click', (e) => {
                e.stopPropagation();
                this.openPanel(item, type);
            });

            container.appendChild(badge);
        });

        if (overflowCount > 0) {
            const overflow = document.createElement('span');
            overflow.className = 'relation-badge relation-badge-overflow';
            overflow.textContent = `+${overflowCount}`;
            overflow.title = types.slice(this.MAX_VISIBLE_BADGES).map((t) => this.BADGE_LABELS[t] || t).join(', ');

            overflow.addEventListener('click', (e) => {
                e.stopPropagation();
                this.openPanel(item, null);
            });

            container.appendChild(overflow);
        }

        cardElement.appendChild(container);
    }

    showBadgeTooltip(event, entry, type) {
        this.hideBadgeTooltip();

        const tooltip = document.createElement('div');
        tooltip.className = 'relation-badge-tooltip';
        tooltip.id = 'relation-badge-tooltip';

        const label = this.BADGE_LABELS[type] || type;
        const header = document.createElement('div');
        header.className = 'relation-badge-tooltip-header';
        header.textContent = `${label} (${entry.count})`;
        tooltip.appendChild(header);

        const maxShow = 5;
        const items = entry.items.slice(0, maxShow);

        items.forEach((rel) => {
            const row = document.createElement('div');
            row.className = 'relation-badge-tooltip-item';
            const targetId = rel.target && rel.target.itemId ? String(rel.target.itemId) : null;
            const targetItem = targetId ? this.app.getEffectiveItemById(targetId) : null;

            if (targetItem) {
                row.textContent = targetItem.title || targetItem.id || targetId;
            } else {
                const extKey = rel.target && rel.target.sourceRef && rel.target.sourceRef.key
                    ? rel.target.sourceRef.key
                    : (targetId || 'Unknown');
                row.textContent = extKey;
                row.classList.add('relation-unresolved');
            }
            tooltip.appendChild(row);
        });

        if (entry.count > maxShow) {
            const more = document.createElement('div');
            more.className = 'relation-badge-tooltip-more';
            more.textContent = `… and ${entry.count - maxShow} more`;
            tooltip.appendChild(more);
        }

        // Position near the badge
        document.body.appendChild(tooltip);
        const rect = event.target.getBoundingClientRect();
        tooltip.style.left = `${rect.left}px`;
        tooltip.style.top = `${rect.bottom + 4}px`;

        // Keep within viewport
        const tooltipRect = tooltip.getBoundingClientRect();
        if (tooltipRect.right > window.innerWidth) {
            tooltip.style.left = `${window.innerWidth - tooltipRect.width - 8}px`;
        }
        if (tooltipRect.bottom > window.innerHeight) {
            tooltip.style.top = `${rect.top - tooltipRect.height - 4}px`;
        }
    }

    hideBadgeTooltip() {
        const existing = document.getElementById('relation-badge-tooltip');
        if (existing) existing.remove();
    }

    // ===== SIDE PANEL (§V5.4) =====

    openPanel(item, filterType) {
        if (!item) return;

        this.panelItem = item;
        this.panelFilterType = filterType || null;

        const panel = document.getElementById('details-panel');
        if (!panel) return;

        // Populate relation content
        const modeEl = document.getElementById('details-mode-relations');

        // Populate relationship tree view
        this.populateRelationshipTree(modeEl, item);

        // Populate relation sections (non-hierarchy types)
        this.populateRelationSections(modeEl, item, filterType);

        // Populate the add-relation type dropdown
        this.populateRelationTypeDropdown();

        // Hide the add-relation form by default
        this.hideAddRelationForm();

        // Use the details panel manager to show the panel in relations mode
        if (this.app.detailsPanelManager) {
            this.app.detailsPanelManager.openForItem(item, 'relations');
        }
    }

    closePanel() {
        if (this.app.detailsPanelManager) {
            this.app.detailsPanelManager.close();
        }
        this.panelItem = null;
        this.panelFilterType = null;
    }

    populateRelationshipTree(container, item) {
        const detailsEl = container ? container.querySelector('.relation-panel-details') : null;
        if (!detailsEl) return;
        detailsEl.innerHTML = '';

        const relations = item && Array.isArray(item.relations) ? item.relations : [];
        const childRels = relations.filter(r => r.type === 'child');

        if (!this.getPrimaryParentTarget(item) && childRels.length === 0) {
            const empty = document.createElement('div');
            empty.className = 'relation-tree-empty';
            empty.textContent = 'No parent/child relationships';
            detailsEl.appendChild(empty);
            return;
        }

        const tree = document.createElement('div');
        tree.className = 'relation-tree';

        const treeModel = this.buildRelationshipTreeModel(item);
        if (!treeModel) {
            const empty = document.createElement('div');
            empty.className = 'relation-tree-empty';
            empty.textContent = 'No parent/child relationships';
            detailsEl.appendChild(empty);
            return;
        }

        tree.appendChild(this.renderRelationshipTreeBranch(treeModel, true, true));

        detailsEl.appendChild(tree);
    }

    buildRelationshipTreeModel(item) {
        const itemId = this.app.getItemIdentity(item) || item.id;
        if (!itemId) return null;

        let branch = this.buildRelationshipTreeNode(
            item,
            String(itemId),
            'self',
            true,
            2,
            new Set([String(itemId)])
        );

        const ancestors = this.getAncestorChain(item);
        for (let index = ancestors.length - 1; index >= 0; index--) {
            const ancestor = ancestors[index];
            ancestor.children = [branch];
            branch = ancestor;
        }

        return branch;
    }

    getAncestorChain(item) {
        const chain = [];
        const visited = new Set();
        let currentItem = item;
        let currentId = this.app.getItemIdentity(item) || item.id;
        if (currentId) {
            visited.add(String(currentId));
        }

        while (currentItem) {
            const parentTarget = this.getPrimaryParentTarget(currentItem);
            if (!parentTarget || !parentTarget.itemId) break;
            if (visited.has(parentTarget.itemId)) break;

            visited.add(parentTarget.itemId);
            const parentItem = parentTarget.item ? parentTarget.item : this.app.getEffectiveItemById(parentTarget.itemId);
            chain.unshift(this.createRelationshipTreeModelNode(
                parentItem,
                parentTarget.itemId,
                chain.length === 0 ? 'ancestor' : 'ancestor',
                false,
                []
            ));

            if (!parentItem) break;
            currentItem = parentItem;
            currentId = parentTarget.itemId;
        }

        if (chain.length > 0) {
            chain[0].role = 'root';
        }

        return chain;
    }

    getPrimaryParentTarget(item) {
        if (!item || !Array.isArray(item.relations)) return null;
        const parentRel = item.relations.find((rel) => rel && rel.type === 'parent' && rel.target && rel.target.itemId);
        if (!parentRel || !parentRel.target || !parentRel.target.itemId) return null;

        const parentId = String(parentRel.target.itemId);
        return {
            itemId: parentId,
            item: this.app.getEffectiveItemById(parentId)
        };
    }

    getDirectChildTargets(item) {
        if (!item || !Array.isArray(item.relations)) return [];

        const seen = new Set();
        return item.relations
            .filter((rel) => rel && rel.type === 'child' && rel.target && rel.target.itemId)
            .map((rel) => String(rel.target.itemId))
            .filter((childId) => {
                if (!childId || seen.has(childId)) return false;
                seen.add(childId);
                return true;
            })
            .map((childId) => ({
                itemId: childId,
                item: this.app.getEffectiveItemById(childId)
            }));
    }

    buildRelationshipTreeNode(item, itemId, role, isCurrent, remainingChildLevels, visited) {
        const node = this.createRelationshipTreeModelNode(item, itemId, role, isCurrent, []);
        if (!item || remainingChildLevels <= 0) {
            return node;
        }

        const nextChildren = this.getDirectChildTargets(item)
            .filter((childTarget) => !visited.has(childTarget.itemId))
            .map((childTarget) => {
                const nextVisited = new Set(visited);
                nextVisited.add(childTarget.itemId);
                return this.buildRelationshipTreeNode(
                    childTarget.item,
                    childTarget.itemId,
                    'child',
                    false,
                    remainingChildLevels - 1,
                    nextVisited
                );
            });

        node.children = nextChildren;
        return node;
    }

    createRelationshipTreeModelNode(item, fallbackId, role, isCurrent, children) {
        return {
            label: this.getRelationshipTreeLabel(item, fallbackId),
            role,
            item,
            itemId: fallbackId ? String(fallbackId) : null,
            isCurrent,
            children: Array.isArray(children) ? children : []
        };
    }

    getRelationshipTreeLabel(item, fallbackId) {
        if (item) {
            return item.title || item.id || String(fallbackId || 'Unknown');
        }
        return String(fallbackId || 'Unknown');
    }

    renderRelationshipTreeBranch(nodeData, isRoot, isLast) {
        const branch = document.createElement('div');
        branch.className = 'relation-tree-branch';
        branch.appendChild(this._createTreeNode(nodeData, isRoot, isLast));

        if (Array.isArray(nodeData.children) && nodeData.children.length > 0) {
            const childrenContainer = document.createElement('div');
            childrenContainer.className = 'relation-tree-children';
            nodeData.children.forEach((childNode, index) => {
                childrenContainer.appendChild(
                    this.renderRelationshipTreeBranch(childNode, false, index === nodeData.children.length - 1)
                );
            });
            branch.appendChild(childrenContainer);
        }

        return branch;
    }

    _createTreeNode(nodeData, isRoot, isLast) {
        const { label, role, item: targetItem, isCurrent } = nodeData;
        const node = document.createElement('div');
        node.className = 'relation-tree-node';
        if (isCurrent) node.classList.add('relation-tree-current');

        // Connector line
        const connector = document.createElement('span');
        connector.className = 'relation-tree-connector';
        connector.textContent = isRoot ? '' : (isLast ? '└─' : '├─');
        node.appendChild(connector);

        // Icon
        const icon = document.createElement('span');
        icon.className = 'relation-tree-icon';
        if (role === 'root') icon.textContent = '◎';
        else if (role === 'ancestor') icon.textContent = '○';
        else if (role === 'child') icon.textContent = '↳';
        else icon.textContent = '●';
        node.appendChild(icon);

        // Label
        const labelEl = document.createElement('span');
        labelEl.className = 'relation-tree-label';
        labelEl.textContent = label;

        if (targetItem && !isCurrent) {
            labelEl.classList.add('relation-tree-clickable');
            labelEl.addEventListener('click', (e) => {
                e.stopPropagation();
                this.openPanel(targetItem, null);
            });
        }

        node.appendChild(labelEl);
        return node;
    }

    populateRelationSections(panel, item, filterType) {
        const container = panel.querySelector('.relation-panel-sections');
        if (!container) return;
        container.innerHTML = '';

        const relations = item && Array.isArray(item.relations) ? item.relations : [];
        if (relations.length === 0 && !filterType) {
            const empty = document.createElement('div');
            empty.className = 'relation-panel-empty';
            empty.textContent = 'No relations';
            container.appendChild(empty);
            return;
        }

        // Group by type
        const grouped = new Map();
        relations.forEach((rel) => {
            if (!rel || !rel.type) return;
            if (!grouped.has(rel.type)) grouped.set(rel.type, []);
            grouped.get(rel.type).push(rel);
        });

        // Canonical type ordering
        const typeOrder = ['parent', 'child', 'epic', 'subtask', 'blocks', 'blockedBy', 'relatesTo', 'duplicates'];
        const orderedTypes = typeOrder.filter((t) => grouped.has(t));
        // Add any remaining types not in the canonical list
        grouped.forEach((_, t) => {
            if (!orderedTypes.includes(t)) orderedTypes.push(t);
        });

        orderedTypes.forEach((type) => {
            const rels = grouped.get(type);
            const section = this.createRelationSection(type, rels, item, filterType);
            container.appendChild(section);
        });
    }

    createRelationSection(type, relations, ownerItem, filterType) {
        const section = document.createElement('div');
        section.className = 'relation-section';
        if (filterType && filterType !== type) {
            section.classList.add('collapsed');
        }

        const icon = this.BADGE_ICONS[type] || '🔗';
        const label = this.BADGE_LABELS[type] || type;

        // Section header
        const header = document.createElement('div');
        header.className = 'relation-section-header';
        header.innerHTML = `<span class="relation-section-icon">${icon}</span> <span class="relation-section-label">${label}</span> <span class="relation-section-count">(${relations.length})</span>`;
        header.addEventListener('click', () => {
            section.classList.toggle('collapsed');
        });
        section.appendChild(header);

        // Section body
        const body = document.createElement('div');
        body.className = 'relation-section-body';

        relations.forEach((rel) => {
            const row = this.createRelationItemRow(rel, ownerItem);
            body.appendChild(row);
        });

        section.appendChild(body);
        return section;
    }

    createRelationItemRow(rel, ownerItem) {
        const row = document.createElement('div');
        row.className = 'relation-item-row';

        const targetId = rel.target && rel.target.itemId ? String(rel.target.itemId) : null;
        const targetItem = targetId ? this.app.getEffectiveItemById(targetId) : null;

        // Item label
        const labelEl = document.createElement('span');
        labelEl.className = 'relation-item-label';

        if (targetItem) {
            labelEl.textContent = targetItem.title || targetItem.id || targetId;
            labelEl.title = targetItem.id || targetId;

            // Clickable to open that item in panel
            labelEl.classList.add('relation-item-clickable');
            labelEl.addEventListener('click', (e) => {
                e.stopPropagation();
                this.openPanel(targetItem, null);
            });
        } else {
            // Unresolved reference
            const extKey = rel.target && rel.target.sourceRef && rel.target.sourceRef.key
                ? rel.target.sourceRef.key
                : (targetId || 'Unknown target');
            labelEl.textContent = extKey;
            labelEl.classList.add('relation-item-unresolved');
            labelEl.title = 'Target not in dataset';
        }

        row.appendChild(labelEl);

        // Source ref badge
        if (rel.target && rel.target.sourceRef && rel.target.sourceRef.key) {
            const refBadge = document.createElement('span');
            refBadge.className = 'relation-source-badge';
            refBadge.textContent = rel.target.sourceRef.sourceType || 'ext';
            row.appendChild(refBadge);
        }

        // Remove button (only for non-derived relations)
        if (!rel._derived) {
            const removeBtn = document.createElement('button');
            removeBtn.className = 'relation-remove-btn';
            removeBtn.textContent = '×';
            removeBtn.title = 'Remove relation';
            removeBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                const ownerId = this.app.getItemIdentity(ownerItem);
                this.removeRelation(ownerId, rel.relationId);
            });
            row.appendChild(removeBtn);
        } else {
            const derivedTag = document.createElement('span');
            derivedTag.className = 'relation-derived-tag';
            derivedTag.textContent = 'derived';
            derivedTag.title = 'Inverse relation — derived automatically';
            row.appendChild(derivedTag);
        }

        return row;
    }

    // ===== ADD RELATION (§V5.4) =====

    populateRelationTypeDropdown() {
        const select = document.getElementById('relation-type-select');
        if (!select) return;

        select.innerHTML = '';
        const types = this.app.relationTypes && this.app.relationTypes.length > 0
            ? this.app.relationTypes
            : Object.keys(DataVisualizationApp.INVERSE_TYPE_MAP);

        types.forEach((type) => {
            const opt = document.createElement('option');
            const typeName = (typeof type === 'object' && type !== null && type.name) ? type.name : String(type);
            opt.value = typeName;
            opt.textContent = typeName;
            select.appendChild(opt);
        });
    }

    toggleAddRelationForm() {
        const form = document.getElementById('relation-add-form');
        if (!form) return;
        form.classList.toggle('hidden');
        if (!form.classList.contains('hidden')) {
            const searchInput = document.getElementById('relation-target-search');
            if (searchInput) {
                searchInput.value = '';
                searchInput.focus();
            }
            this.clearTargetSearchResults();
        }
    }

    hideAddRelationForm() {
        const form = document.getElementById('relation-add-form');
        if (form) form.classList.add('hidden');
        this.clearTargetSearchResults();
    }

    updateTargetSearchResults() {
        const searchInput = document.getElementById('relation-target-search');
        const resultsContainer = document.getElementById('relation-target-results');
        if (!searchInput || !resultsContainer) return;

        const query = searchInput.value.trim().toLowerCase();
        resultsContainer.innerHTML = '';

        if (query.length < 1) return;

        const ownerItemId = this.panelItem ? this.app.getItemIdentity(this.panelItem) : null;
        const matches = (this.app.dataset || []).filter((item) => {
            const itemId = this.app.getItemIdentity(item);
            if (itemId === ownerItemId) return false; // exclude self
            const title = (item.title || '').toLowerCase();
            const id = (item.id || '').toLowerCase();
            return title.includes(query) || id.includes(query);
        }).slice(0, 10);

        matches.forEach((item) => {
            const row = document.createElement('div');
            row.className = 'relation-search-result';
            row.textContent = `${item.id || ''} — ${item.title || 'Untitled'}`;
            row.dataset.itemId = this.app.getItemIdentity(item);
            row.addEventListener('click', () => {
                searchInput.value = item.id || item.title || '';
                searchInput.dataset.selectedItemId = this.app.getItemIdentity(item);
                resultsContainer.innerHTML = '';
            });
            resultsContainer.appendChild(row);
        });

        if (matches.length === 0) {
            const noResults = document.createElement('div');
            noResults.className = 'relation-search-no-results';
            noResults.textContent = 'No matching items';
            resultsContainer.appendChild(noResults);
        }
    }

    clearTargetSearchResults() {
        const resultsContainer = document.getElementById('relation-target-results');
        if (resultsContainer) resultsContainer.innerHTML = '';
        const searchInput = document.getElementById('relation-target-search');
        if (searchInput) {
            searchInput.value = '';
            delete searchInput.dataset.selectedItemId;
        }
    }

    submitAddRelation() {
        const typeSelect = document.getElementById('relation-type-select');
        const searchInput = document.getElementById('relation-target-search');
        if (!typeSelect || !searchInput || !this.panelItem) return;

        const type = typeSelect.value;
        if (!type) {
            this.app.showNotification('Please select a relation type', 'warning');
            return;
        }

        const selectedTargetId = searchInput.dataset.selectedItemId;
        const manualEntry = searchInput.value.trim();

        if (!selectedTargetId && !manualEntry) {
            this.app.showNotification('Please select or enter a target', 'warning');
            return;
        }

        const ownerId = this.app.getItemIdentity(this.panelItem);

        // Build target info
        const targetInfo = {};
        if (selectedTargetId) {
            targetInfo.itemId = selectedTargetId;
        } else {
            // Manual entry — treat as unresolved external reference
            targetInfo.itemId = manualEntry;
            targetInfo.sourceRef = { key: manualEntry };
        }

        this.addRelation(ownerId, type, targetInfo);
    }

    addRelation(ownerItemId, type, targetInfo) {
        // Prevent duplicate: same type + same target
        const ownerItem = this.app.getEffectiveItemById(ownerItemId);
        if (ownerItem && Array.isArray(ownerItem.relations)) {
            const targetId = targetInfo.itemId ? String(targetInfo.itemId) : null;
            const duplicate = ownerItem.relations.some(
                (r) => r.type === type && r.target && String(r.target.itemId) === targetId && !r._derived
            );
            if (duplicate) {
                this.app.showNotification('This relation already exists', 'warning');
                return;
            }
        }

        const now = window.GridDateUtils.createLocalTimestamp();
        const relationId = `rel-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
        const changeId = `relchg-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;

        const changeEntry = {
            changeId: changeId,
            action: 'create',
            target: {
                ownerItemId: ownerItemId,
                relationId: relationId
            },
            proposed: {
                relationId: relationId,
                type: type,
                direction: 'outward',
                target: {
                    itemId: targetInfo.itemId || null,
                    sourceRef: targetInfo.sourceRef || undefined
                }
            },
            baseline: {},
            meta: {
                author: this.app.getCurrentUserName(),
                createdAt: now,
                updatedAt: now
            }
        };

        this.app.ensureChangesContainer();
        if (!Array.isArray(this.app.pendingChanges.relations)) {
            this.app.pendingChanges.relations = [];
        }
        this.app.pendingChanges.relations.push(changeEntry);

        // Rebuild effective dataset and re-render
        this.app.rebuildEffectiveDataset({ render: true });

        // Re-open panel to show the new relation
        const updatedItem = this.app.getEffectiveItemById(ownerItemId);
        if (updatedItem) {
            this.openPanel(updatedItem, type);
        }

        this.app.showNotification('Relation added', 'success');
    }

    removeRelation(ownerItemId, relationId) {
        if (!ownerItemId || !relationId) return;

        const now = window.GridDateUtils.createLocalTimestamp();
        const changeId = `relchg-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;

        const changeEntry = {
            changeId: changeId,
            action: 'delete',
            target: {
                ownerItemId: ownerItemId,
                relationId: relationId
            },
            proposed: {},
            baseline: {},
            meta: {
                author: this.app.getCurrentUserName(),
                createdAt: now,
                updatedAt: now
            }
        };

        this.app.ensureChangesContainer();
        if (!Array.isArray(this.app.pendingChanges.relations)) {
            this.app.pendingChanges.relations = [];
        }
        this.app.pendingChanges.relations.push(changeEntry);

        // Rebuild and re-render
        this.app.rebuildEffectiveDataset({ render: true });

        // Re-open panel for the owner item
        const updatedItem = this.app.getEffectiveItemById(ownerItemId);
        if (updatedItem) {
            this.openPanel(updatedItem, this.panelFilterType);
        } else {
            this.closePanel();
        }

        this.app.showNotification('Relation removed', 'success');
    }

    setParent(childItemId, parentItemId) {
        const childItem = this.app.getEffectiveItemById(childItemId);
        if (!childItem) return;

        // Remove any existing parent relation from the child
        if (Array.isArray(childItem.relations)) {
            const existingParent = childItem.relations.find(
                (r) => r.type === 'parent' && !r._derived
            );
            if (existingParent && existingParent.relationId) {
                this.removeRelation(childItemId, existingParent.relationId);
            }
        }

        // Add parent relation: child → parent
        this.addRelation(childItemId, 'parent', { itemId: parentItemId });
    }

    // ===== DERIVED FIELDS (§V5.8) =====

    computeDerivedRelationFields(dataset) {
        if (!Array.isArray(dataset)) return;

        dataset.forEach((item) => {
            const relations = Array.isArray(item.relations) ? item.relations : [];

            item._hasParent = relations.some((r) => r.type === 'parent');
            item._hasChildren = relations.some((r) => r.type === 'child');
            item._childrenCount = relations.filter((r) => r.type === 'child').length;
            item._blocksCount = relations.filter((r) => r.type === 'blocks').length;
            item._isBlocked = relations.some((r) => r.type === 'blockedBy');
            item._relationsCount = relations.length;
        });
    }

    static DERIVED_FIELD_DEFS = {
        _hasParent: { type: 'scalar', subtype: 'boolean', kind: 'relationship' },
        _hasChildren: { type: 'scalar', subtype: 'boolean', kind: 'relationship' },
        _childrenCount: { type: 'scalar', subtype: 'number', kind: 'relationship' },
        _blocksCount: { type: 'scalar', subtype: 'number', kind: 'relationship' },
        _isBlocked: { type: 'scalar', subtype: 'boolean', kind: 'relationship' },
        _relationsCount: { type: 'scalar', subtype: 'number', kind: 'relationship' }
    };

    registerDerivedFields() {
        // Only register if dataset has any items with relations
        const hasRelations = (this.app.dataset || []).some(
            (item) => Array.isArray(item.relations) && item.relations.length > 0
        );
        if (!hasRelations) return;

        Object.entries(RelationUIManager.DERIVED_FIELD_DEFS).forEach(([fieldName, def]) => {
            // Add to fieldTypes
            this.app.fieldTypes.set(fieldName, def.type);

            // Register schema metadata: kind=relationship, visible=false by default
            if (!this.app.schemaFields[fieldName]) {
                this.app.schemaFields[fieldName] = {};
            }
            this.app.schemaFields[fieldName].type = def.type;
            this.app.schemaFields[fieldName].kind = def.kind;
            this.app.schemaFields[fieldName].visible = false;

            // Add to availableFields if not already present
            if (!this.app.availableFields.includes(fieldName)) {
                this.app.availableFields.push(fieldName);
            }

            // Compute distinct values
            const values = new Set();
            (this.app.dataset || []).forEach((item) => {
                const v = item[fieldName];
                if (v !== null && v !== undefined) {
                    values.add(v);
                }
            });
            this.app.distinctValues.set(fieldName, Array.from(values).sort());
        });
    }

    // ===== FOCUS MODE (§V5.6) =====

    computeFocusSet(rootItemId, depth, types) {
        const visited = new Set();
        const queue = [{ id: rootItemId, currentDepth: 0 }];
        visited.add(rootItemId);
        const normalizedDepth = this.normalizeFocusDepth(depth);

        while (queue.length > 0) {
            const { id, currentDepth } = queue.shift();
            if (currentDepth >= normalizedDepth) continue;

            const item = this.app.getEffectiveItemById(id);
            if (!item || !Array.isArray(item.relations)) continue;

            item.relations.forEach((rel) => {
                if (!rel || !rel.target || !rel.target.itemId) return;

                // Filter by type if specified
                if (types && types.length > 0 && !types.includes(rel.type)) return;

                const neighborId = String(rel.target.itemId);
                if (!visited.has(neighborId)) {
                    visited.add(neighborId);
                    queue.push({ id: neighborId, currentDepth: currentDepth + 1 });
                }
            });
        }

        return visited;
    }

    enterFocusMode(itemId, depth, types) {
        this.focusMode = true;
        this.focusRootId = itemId;
        this.focusDepth = this.normalizeFocusDepth(depth || this.focusDepth || 1);
        this.focusTypes = types || this.focusTypes || null;

        this.recomputeFocus();
        this.showFocusControls();
    }

    exitFocusMode() {
        this.focusMode = false;
        this.focusRootId = null;
        this.focusSet.clear();

        const gridContainer = document.getElementById('grid-container');
        if (gridContainer) gridContainer.classList.remove('focus-mode');

        // Remove focus classes from all cards
        document.querySelectorAll('.card').forEach((card) => {
            card.classList.remove('focus-related', 'focus-root');
        });

        this.hideFocusControls();
    }

    recomputeFocus() {
        this.focusSet = this.computeFocusSet(
            this.focusRootId,
            this.focusDepth,
            this.focusTypes
        );
        this.applyFocusClasses();
    }

    applyFocusClasses() {
        const gridContainer = document.getElementById('grid-container');
        if (!gridContainer) return;

        if (this.focusMode) {
            gridContainer.classList.add('focus-mode');
        } else {
            gridContainer.classList.remove('focus-mode');
            return;
        }

        document.querySelectorAll('.card').forEach((card) => {
            const itemId = card.dataset.itemId;
            if (this.focusSet.has(itemId)) {
                card.classList.add('focus-related');
                if (itemId === this.focusRootId) {
                    card.classList.add('focus-root');
                } else {
                    card.classList.remove('focus-root');
                }
            } else {
                card.classList.remove('focus-related', 'focus-root');
            }
        });
    }

    showFocusControls() {
        const controls = document.getElementById('focus-controls');
        if (controls) controls.classList.remove('hidden');

        // Sync depth selector
        const depthSelect = document.getElementById('focus-depth-select');
        if (depthSelect) depthSelect.value = String(this.normalizeFocusDepth(this.focusDepth));
    }

    hideFocusControls() {
        const controls = document.getElementById('focus-controls');
        if (controls) controls.classList.add('hidden');
    }

    getSelectedFocusTypes() {
        const checkboxes = document.querySelectorAll('.focus-type-checkbox:checked');
        if (checkboxes.length === 0) return null; // null = all
        return Array.from(checkboxes).map((cb) => cb.value);
    }

    normalizeFocusDepth(depth) {
        const parsedDepth = Number.parseInt(depth, 10);
        if (!Number.isFinite(parsedDepth)) return 1;
        return Math.min(this.MAX_FOCUS_DEPTH, Math.max(1, parsedDepth));
    }

    // ===== POST-RENDER HOOK =====

    onGridRendered() {
        if (this.focusMode) {
            this.applyFocusClasses();
        }
    }
}

// Attach to app on DOMContentLoaded (relations.js loads after app.js and advanced-features.js)
document.addEventListener('DOMContentLoaded', () => {
    if (typeof app !== 'undefined' && app) {
        app.relationUIManager = new RelationUIManager(app);

        // If data was already loaded before the relation manager was created,
        // fully resync runtime state so derived relation fields
        // (_hasParent, _childrenCount, etc.) are computed and usable everywhere.
        if (app.dataset && app.dataset.length > 0) {
            app.dataset = app.buildEffectiveDataset();
            app.analyzeFields();
            app.updateFilteredData();
            if (typeof app.renderFieldSelectors === 'function') {
                app.renderFieldSelectors();
            }
            if (typeof app.renderSlicers === 'function') {
                app.renderSlicers();
            }
            if (typeof app.renderGrid === 'function') {
                app.renderGrid();
            }
        }

        // Wire the Focus button in the side panel
        const focusBtn = document.getElementById('details-panel-focus-btn');
        if (focusBtn) {
            focusBtn.addEventListener('click', () => {
                if (app.relationUIManager.panelItem) {
                    const itemId = app.getItemIdentity(app.relationUIManager.panelItem);
                    app.relationUIManager.enterFocusMode(itemId);
                }
            });
        }
    }
});
