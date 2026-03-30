// Copyright (c) 2026 Malte Doerper. MIT License. See LICENSE file.

// Advanced Features Manager
// Handles comments system, header interactions, drag reordering, and aggregations

class AdvancedFeaturesManager {
    constructor(app) {
        this.app = app;
        this.currentCommentsItem = null;
        this.collapsedHeaders = new Set();
        this.headerOrdering = { rows: [], columns: [] };
        this.draggedHeader = null;
        this.aggregationCache = new Map();
        this.lastRenderedHeaderSequences = { row: [], column: [] };
        this.pendingHeaderOrderReasons = {
            row: { reason: 'initial render', details: {} },
            column: { reason: 'initial render', details: {} }
        };
        
        this.initializeAdvancedFeatures();
    }

    setPendingHeaderOrderReason(headerType, reason, details = {}) {
        this.pendingHeaderOrderReasons[headerType] = { reason, details };
    }

    arraysEqual(left, right) {
        if (left.length !== right.length) {
            return false;
        }

        return left.every((value, index) => value === right[index]);
    }

    logHeaderSequenceChange(headerType, nextSequence, fallbackReason, fallbackDetails = {}) {
        const previousSequence = this.lastRenderedHeaderSequences[headerType] || [];
        const normalizedNextSequence = [...nextSequence];
        const pendingReason = this.pendingHeaderOrderReasons[headerType];

        if (previousSequence.length === 0) {
            if (pendingReason && pendingReason.reason !== 'initial render') {
                console.log(`[HEADER ORDER DEBUG] ${headerType} sequence initialized`, {
                    reason: pendingReason.reason,
                    next: normalizedNextSequence,
                    ...pendingReason.details
                });
            }

            this.lastRenderedHeaderSequences[headerType] = normalizedNextSequence;
            this.pendingHeaderOrderReasons[headerType] = null;
            return;
        }

        if (this.arraysEqual(previousSequence, normalizedNextSequence)) {
            this.pendingHeaderOrderReasons[headerType] = null;
            return;
        }

        const reason = pendingReason ? pendingReason.reason : fallbackReason;
        const details = pendingReason ? pendingReason.details : fallbackDetails;

        console.log(`[HEADER ORDER DEBUG] ${headerType} sequence changed`, {
            reason,
            previous: previousSequence,
            next: normalizedNextSequence,
            ...details
        });

        this.lastRenderedHeaderSequences[headerType] = normalizedNextSequence;
        this.pendingHeaderOrderReasons[headerType] = null;
    }
    
    initializeAdvancedFeatures() {
        this.loadPersistedHeaderState();
        this.setupCommentsDialog();
        this.setupAggregations();
    }
    
    setupCommentsDialog() {
        const commentsDialog = document.getElementById('comments-dialog');
        const closeBtn = commentsDialog.querySelector('#comments-close-btn');
        const addBtn = commentsDialog.querySelector('#add-comment-btn');
        const modalClose = commentsDialog.querySelector('.modal-close');
        
        if (closeBtn) closeBtn.addEventListener('click', () => this.closeCommentsDialog());
        if (modalClose) modalClose.addEventListener('click', () => this.closeCommentsDialog());
        if (addBtn) addBtn.addEventListener('click', () => this.addComment());
    }
    
    openCommentsDialog(item) {
        this.currentCommentsItem = item;
        
        const dialog = document.getElementById('comments-dialog');
        if (!dialog) {
            this.app.showNotification('Comments dialog not available', 'error');
            return;
        }
        
        this.populateCommentsDialog(item);
        dialog.classList.remove('hidden');
    }
    
    populateCommentsDialog(item) {
        const container = document.querySelector('.comments-container');
        if (!container) return;
        
        container.innerHTML = '';
        
        // Get comments from item
        const comments = this.getItemComments(item);
        
        if (comments.length === 0) {
            const emptyMessage = document.createElement('div');
            emptyMessage.className = 'empty-comments';
            emptyMessage.textContent = 'No comments yet.';
            emptyMessage.style.cssText = `
                text-align: center;
                color: #666;
                padding: 20px;
                font-style: italic;
            `;
            container.appendChild(emptyMessage);
        } else {
            comments.forEach(comment => {
                this.renderCommentThread(container, comment);
            });
        }
    }
    
    getItemComments(item) {
        const commentThreads = typeof this.app.getItemCommentThreads === 'function'
            ? this.app.getItemCommentThreads(item)
            : null;

        if (commentThreads && Array.isArray(commentThreads.threads)) {
            const flat = [];
            commentThreads.threads.forEach((thread) => {
                if (!thread || !Array.isArray(thread.messages)) {
                    return;
                }

                thread.messages.forEach((message, index) => {
                    flat.push({
                        id: message.id || `${thread.id || 'comment'}-${index + 1}`,
                        author: message.author || 'User',
                        message: message.text || message.message || '',
                        timestamp: message.timestamp,
                        replies: []
                    });
                });
            });
            return flat;
        }
        
        return [];
    }
    
    renderCommentThread(container, comment) {
        const threadDiv = document.createElement('div');
        threadDiv.className = 'comment-thread';
        
        // Main comment
        const commentDiv = document.createElement('div');
        commentDiv.className = 'comment-message';
        
        const header = document.createElement('div');
        header.className = 'comment-header';
        header.style.cssText = `
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 6px;
            font-size: 11px;
            color: #666;
        `;
        
        const author = document.createElement('span');
        author.textContent = comment.author || 'User';
        author.style.fontWeight = '600';
        
        const timestamp = document.createElement('span');
        timestamp.textContent = this.formatTimestamp(comment.timestamp);
        
        header.appendChild(author);
        header.appendChild(timestamp);
        
        const messageContent = document.createElement('div');
        messageContent.className = 'comment-content';
        messageContent.textContent = comment.message || '';
        messageContent.style.cssText = `
            font-size: 13px;
            line-height: 1.4;
            color: #333;
        `;
        
        commentDiv.appendChild(header);
        commentDiv.appendChild(messageContent);
        threadDiv.appendChild(commentDiv);
        
        // Replies
        if (comment.replies && comment.replies.length > 0) {
            comment.replies.forEach(reply => {
                this.renderCommentReply(threadDiv, reply);
            });
        }
        
        container.appendChild(threadDiv);
    }
    
    renderCommentReply(threadDiv, reply) {
        const replyDiv = document.createElement('div');
        replyDiv.className = 'comment-reply';
        replyDiv.style.cssText = `
            margin-left: 20px;
            margin-top: 8px;
            padding: 8px;
            background: #f5f5f5;
            border-radius: 4px;
            font-size: 12px;
        `;
        
        const header = document.createElement('div');
        header.style.cssText = `
            display: flex;
            justify-content: space-between;
            margin-bottom: 4px;
            font-size: 10px;
            color: #666;
        `;
        
        const author = document.createElement('span');
        author.textContent = reply.author || 'User';
        author.style.fontWeight = '600';
        
        const timestamp = document.createElement('span');
        timestamp.textContent = this.formatTimestamp(reply.timestamp);
        
        header.appendChild(author);
        header.appendChild(timestamp);
        
        const content = document.createElement('div');
        content.textContent = reply.message || '';
        content.style.cssText = `
            color: #333;
            line-height: 1.3;
        `;
        
        replyDiv.appendChild(header);
        replyDiv.appendChild(content);
        threadDiv.appendChild(replyDiv);
    }
    
    formatTimestamp(timestamp) {
        if (!timestamp) return 'Unknown';

        return window.GridDateUtils.normalizeDateStamp(timestamp, String(timestamp));
    }
    
    addComment() {
        const messageInput = document.getElementById('comment-message-input');
        if (!messageInput || !this.currentCommentsItem) return;
        
        const message = messageInput.value.trim();
        if (!message) {
            this.app.showNotification('Please enter a comment message', 'warning');
            return;
        }
        
        const newComment = {
            id: `comment-${Date.now()}`,
            message: message,
            author: typeof this.app.getCurrentUserName === 'function'
                ? this.app.getCurrentUserName()
                : 'Local User',
            timestamp: window.GridDateUtils.createLocalTimestamp(),
            replies: []
        };
        
        const existingCommentThreads = typeof this.app.getItemCommentThreads === 'function'
            ? this.app.getItemCommentThreads(this.currentCommentsItem)
            : null;
        const nextCommentThreads = existingCommentThreads && Array.isArray(existingCommentThreads.threads)
            ? this.app.cloneValue(existingCommentThreads)
            : { threads: [] };

        nextCommentThreads.threads.push({
            id: newComment.id,
            messages: [{
                id: newComment.id,
                author: newComment.author,
                timestamp: newComment.timestamp,
                text: newComment.message
            }]
        });

        this.currentCommentsItem.comment = nextCommentThreads;
        delete this.currentCommentsItem.comments;
        
        // Clear input
        messageInput.value = '';
        
        // Refresh comments display
        this.populateCommentsDialog(this.currentCommentsItem);
        
        // Update the item in the dataset
        this.app.updateItem(this.app.getItemIdentity(this.currentCommentsItem) || this.currentCommentsItem.id, {
            comment: nextCommentThreads
        });
        
        this.app.showNotification('Comment added', 'success');
    }
    
    closeCommentsDialog() {
        const dialog = document.getElementById('comments-dialog');
        if (dialog) {
            dialog.classList.add('hidden');
        }
        this.currentCommentsItem = null;
    }
    
    // ===== HEADER INTERACTIONS (COLLAPSE/EXPAND) =====
    
    setupHeaderInteractions() {
        // Header interactions will be set up when grid is rendered
        // This is called from the grid rendering process
    }
    
    attachHeaderListeners(headerElement, headerValue, headerType) {
        // Add collapse/expand functionality
        const collapseBtn = document.createElement('button');
        collapseBtn.className = 'header-collapse-btn collapse-btn header-toggle';
        collapseBtn.style.cssText = `
            background: none;
            border: none;
            font-size: 16px;
            font-weight: bold;
            color: #666;
            cursor: pointer;
            padding: 2px 6px;
            margin-left: 6px;
            border-radius: 2px;
        `;
        
        const headerKey = `${headerType}-${headerValue}`;
        const isCollapsed = this.collapsedHeaders.has(headerKey);
        this.applyCollapsedHeaderState(headerElement, collapseBtn, isCollapsed, headerType, headerValue);
        
        collapseBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            this.toggleHeaderCollapse(headerValue, headerType, collapseBtn, headerElement);
        });
        
        headerElement.appendChild(collapseBtn);
        
        // Add drag handle for reordering
        this.addHeaderDragHandle(headerElement, headerValue, headerType);

        // Right-click opens header values editor in details panel
        headerElement.addEventListener('contextmenu', (e) => {
            e.preventDefault();
            e.stopPropagation();
            const fieldName = headerType === 'column'
                ? this.app.currentAxisSelections.x
                : this.app.currentAxisSelections.y;
            if (fieldName && this.app.detailsPanelManager) {
                this.app.detailsPanelManager.openForHeader(fieldName, headerType, headerValue);
            }
        });
    }
    
    toggleHeaderCollapse(headerValue, headerType, button, headerElement) {
        const headerKey = `${headerType}-${headerValue}`;
        let isCollapsed;
        
        if (this.collapsedHeaders.has(headerKey)) {
            this.collapsedHeaders.delete(headerKey);
            isCollapsed = false;
        } else {
            this.collapsedHeaders.add(headerKey);
            isCollapsed = true;
        }

        this.applyCollapsedHeaderState(headerElement, button, isCollapsed, headerType, headerValue);
        
        // Re-render grid to apply collapse state
        this.app.renderGrid();
        
        // Persist header state
        this.persistHeaderState();
    }
    
    isHeaderCollapsed(headerValue, headerType) {
        return this.collapsedHeaders.has(`${headerType}-${headerValue}`);
    }

    applyCollapsedHeaderState(headerElement, button, isCollapsed, headerType, headerValue) {
        headerElement.classList.toggle('collapsed', isCollapsed);
        button.textContent = isCollapsed ? '+' : '−';
        button.title = isCollapsed ? 'Expand header' : 'Collapse header';
        button.setAttribute(
            'aria-label',
            isCollapsed
                ? `Expand ${headerType} ${headerValue || '(empty)'}`
                : `Collapse ${headerType} ${headerValue || '(empty)'}`
        );
    }
    
    // ===== DRAG REORDERING =====
    
    addHeaderDragHandle(headerElement, headerValue, headerType) {
        const dragHandle = document.createElement('span');
        dragHandle.className = 'header-drag-handle';
        dragHandle.textContent = '⋮⋮';
        dragHandle.title = 'Drag to reorder';
        dragHandle.style.cssText = `
            cursor: move;
            color: #999;
            font-size: 12px;
            margin-right: 6px;
            line-height: 1;
            user-select: none;
        `;
        
        headerElement.insertBefore(dragHandle, headerElement.firstChild);
        
        // Make header draggable
        headerElement.draggable = true;
        headerElement.dataset.dragType = 'header';
        headerElement.dataset.headerValue = headerValue;
        headerElement.dataset.headerType = headerType;
        
        headerElement.addEventListener('dragstart', (e) => {
            this.handleHeaderDragStart(e, headerValue, headerType);
        });

        headerElement.addEventListener('dragend', (e) => {
            this.handleHeaderDragEnd(e);
        });
        
        headerElement.addEventListener('dragover', (e) => {
            this.handleHeaderDragOver(e);
        });

        headerElement.addEventListener('dragleave', (e) => {
            this.handleHeaderDragLeave(e);
        });
        
        headerElement.addEventListener('drop', (e) => {
            this.handleHeaderDrop(e, headerValue, headerType);
        });
    }
    
    handleHeaderDragStart(event, headerValue, headerType) {
        const headerElement = event.currentTarget;
        this.draggedHeader = {
            value: headerValue,
            type: headerType,
            element: headerElement
        };

        console.log('[HEADER ORDER DEBUG] drag start', {
            headerType,
            headerValue
        });
        
        event.dataTransfer.effectAllowed = 'move';
        event.dataTransfer.setData('text/plain', `header-${headerType}-${headerValue}`);
        
        // Add visual feedback
        headerElement.classList.add('dragging');
    }
    
    handleHeaderDragOver(event) {
        if (this.draggedHeader) {
            event.preventDefault();
            event.dataTransfer.dropEffect = 'move';
            
            // Add drop indicator
            const headerElement = event.currentTarget;
            if (headerElement !== this.draggedHeader.element) {
                headerElement.classList.add('drag-over-header');
            }
        }
    }

    handleHeaderDragLeave(event) {
        const headerElement = event.currentTarget;
        if (!headerElement.contains(event.relatedTarget)) {
            headerElement.classList.remove('drag-over-header');
        }
    }

    handleHeaderDragEnd(event) {
        this.cleanupHeaderDragState();
    }
    
    handleHeaderDrop(event, targetValue, targetType) {
        event.preventDefault();
        const targetElement = event.currentTarget;
        
        if (!this.draggedHeader || this.draggedHeader.type !== targetType) {
            this.cleanupHeaderDragState();
            return;
        }
        
        const sourceValue = this.draggedHeader.value;

        console.log('[HEADER ORDER DEBUG] drop', {
            sourceValue,
            targetValue,
            targetType
        });
        
        if (sourceValue === targetValue) {
            this.cleanupHeaderDragState();
            return; // Same header
        }
        
        // Reorder headers
        this.reorderHeaders(sourceValue, targetValue, targetType);
        
        // Clean up
        this.cleanupHeaderDragState();
        
        // Re-render grid
        this.app.renderGrid();
        
        // Persist new order
        this.persistHeaderState();
    }
    
    reorderHeaders(sourceValue, targetValue, headerType) {
        const orderArray = headerType === 'row' ? this.headerOrdering.rows : this.headerOrdering.columns;
        const beforeOrder = [...orderArray];

        console.log('[HEADER ORDER DEBUG] reorder request', {
            headerType,
            sourceValue,
            targetValue,
            beforeOrder
        });
        
        const sourceIndex = orderArray.indexOf(sourceValue);
        const targetIndex = orderArray.indexOf(targetValue);
        
        if (sourceIndex !== -1) {
            orderArray.splice(sourceIndex, 1);
        }
        
        const newTargetIndex = orderArray.indexOf(targetValue);
        if (newTargetIndex !== -1) {
            orderArray.splice(newTargetIndex + 1, 0, sourceValue);
        } else {
            orderArray.push(sourceValue);
        }

        this.setPendingHeaderOrderReason(headerType, 'drag reorder', {
            sourceValue,
            targetValue,
            orderBefore: beforeOrder,
            orderAfter: [...orderArray]
        });

        console.log('[HEADER ORDER DEBUG] reorder applied', {
            headerType,
            sourceValue,
            targetValue,
            orderAfter: [...orderArray]
        });
    }

    cleanupHeaderDragState() {
        if (this.draggedHeader && this.draggedHeader.element) {
            this.draggedHeader.element.classList.remove('dragging');
        }

        document.querySelectorAll('.drag-over-header').forEach(el => {
            el.classList.remove('drag-over-header');
        });

        this.draggedHeader = null;
    }
    
    getOrderedHeaders(values, headerType) {
        const orderArray = headerType === 'row' ? this.headerOrdering.rows : this.headerOrdering.columns;
        const missingValues = [];
        
        // Ensure all values are in the order array
        values.forEach(value => {
            if (!orderArray.includes(value)) {
                orderArray.push(value);
                missingValues.push(value);
            }
        });

        if (missingValues.length > 0) {
            this.setPendingHeaderOrderReason(headerType, 'new header values discovered', {
                addedValues: missingValues,
                orderingAfterAppend: [...orderArray]
            });
        }
        
        // Return values sorted by order array
        const orderedValues = [...values].sort((a, b) => {
            const aIndex = orderArray.indexOf(a);
            const bIndex = orderArray.indexOf(b);
            return aIndex - bIndex;
        });

        const previousSequence = this.lastRenderedHeaderSequences[headerType] || [];
        const removedValues = previousSequence.filter(value => !orderedValues.includes(value));
        const addedValues = orderedValues.filter(value => !previousSequence.includes(value));

        this.logHeaderSequenceChange(
            headerType,
            orderedValues,
            'header values changed from current dataset',
            {
                addedValues,
                removedValues
            }
        );

        return orderedValues;
    }

    getRenderedHeaderOrdering() {
        const getSequence = (headerType) => Array.from(
            document.querySelectorAll(`.grid-header[data-header-type="${headerType}"]`)
        ).map((header) => header.dataset.value || '');

        return {
            rows: getSequence('row'),
            columns: getSequence('column')
        };
    }
    
    // ===== AGGREGATIONS =====
    
    setupAggregations() {
        // Aggregations are calculated during grid rendering
        this.aggregationCache.clear();
    }
    
    calculateAggregations(gridData, bottomRightField) {
        if (!bottomRightField) return { rows: new Map(), columns: new Map(), total: null };
        
        const rowAggregations = new Map();
        const columnAggregations = new Map();
        let totalSum = 0;
        let totalCount = 0;
        
        // Cache key for this calculation
        const cacheKey = `${bottomRightField}-${JSON.stringify(gridData.structure)}`;
        
        if (this.aggregationCache.has(cacheKey)) {
            return this.aggregationCache.get(cacheKey);
        }
        
        // Calculate row aggregations
        Object.entries(gridData.structure).forEach(([rowValue, rowData]) => {
            let rowSum = 0;
            let rowCount = 0;
            
            Object.values(rowData).forEach(cellItems => {
                cellItems.forEach(item => {
                    const value = this.app.getFieldValue(item, bottomRightField);
                    const numericValue = this.parseNumericValue(value);
                    if (numericValue !== null) {
                        rowSum += numericValue;
                        rowCount++;
                        totalSum += numericValue;
                        totalCount++;
                    }
                });
            });
            
            if (rowCount > 0) {
                rowAggregations.set(rowValue, {
                    sum: rowSum,
                    count: rowCount,
                    average: rowSum / rowCount
                });
            }
        });
        
        // Calculate column aggregations
        const allColumnValues = new Set();
        Object.values(gridData.structure).forEach(rowData => {
            Object.keys(rowData).forEach(colValue => {
                allColumnValues.add(colValue);
            });
        });
        
        allColumnValues.forEach(colValue => {
            let colSum = 0;
            let colCount = 0;
            
            Object.values(gridData.structure).forEach(rowData => {
                if (rowData[colValue]) {
                    rowData[colValue].forEach(item => {
                        const value = this.app.getFieldValue(item, bottomRightField);
                        const numericValue = this.parseNumericValue(value);
                        if (numericValue !== null) {
                            colSum += numericValue;
                            colCount++;
                        }
                    });
                }
            });
            
            if (colCount > 0) {
                columnAggregations.set(colValue, {
                    sum: colSum,
                    count: colCount,
                    average: colSum / colCount
                });
            }
        });
        
        const result = {
            rows: rowAggregations,
            columns: columnAggregations,
            total: totalCount > 0
                ? {
                    sum: totalSum,
                    count: totalCount,
                    average: totalSum / totalCount
                }
                : null
        };
        
        // Cache the result
        this.aggregationCache.set(cacheKey, result);
        
        return result;
    }
    
    parseNumericValue(value) {
        if (typeof value === 'number') {
            return isFinite(value) ? value : null;
        }
        
        if (typeof value === 'string') {
            const trimmed = value.trim();
            if (trimmed === '') return null;
            
            const parsed = parseFloat(trimmed);
            return isFinite(parsed) ? parsed : null;
        }
        
        return null;
    }
    
    formatAggregationValue(aggregation) {
        if (!aggregation) return '';

        const mode = this.app.aggregationMode || 'sum';
        if (mode === 'count') {
            return aggregation.count.toString();
        }

        // sum mode
        if (Number.isInteger(aggregation.sum)) {
            return aggregation.sum.toString();
        } else {
            return aggregation.sum.toFixed(1);
        }
    }
    
    // ===== PERSISTENCE OF HEADER STATE =====
    
    persistHeaderState() {
        const state = {
            collapsedHeaders: Array.from(this.collapsedHeaders),
            headerOrdering: this.headerOrdering
        };
        
        localStorage.setItem('grid-header-state', JSON.stringify(state));
    }
    
    loadPersistedHeaderState() {
        try {
            const saved = localStorage.getItem('grid-header-state');
            if (saved) {
                const state = JSON.parse(saved);
                
                this.collapsedHeaders = new Set(state.collapsedHeaders || []);
                this.headerOrdering = state.headerOrdering || { rows: [], columns: [] };
            }
        } catch (error) {
            console.error('Failed to load persisted header state:', error);
        }
    }
    
    // ===== PUBLIC INTERFACE =====
    
    getCollapsedHeaders() {
        return Array.from(this.collapsedHeaders);
    }

    normalizeCollapsedHeaderKeys(headerKeys) {
        if (Array.isArray(headerKeys)) {
            return headerKeys;
        }

        if (!headerKeys || typeof headerKeys !== 'object') {
            return [];
        }

        const rowKeys = Array.isArray(headerKeys.rows)
            ? headerKeys.rows.map((value) => `row-${value}`)
            : [];
        const columnKeys = Array.isArray(headerKeys.columns)
            ? headerKeys.columns.map((value) => `column-${value}`)
            : [];

        return [...rowKeys, ...columnKeys];
    }
    
    getHeaderOrdering() {
        return { ...this.headerOrdering };
    }
    
    setCollapsedHeaders(headerKeys) {
        this.collapsedHeaders = new Set(this.normalizeCollapsedHeaderKeys(headerKeys));
    }
    
    setHeaderOrdering(ordering) {
        this.headerOrdering = {
            rows: [...(ordering.rows || [])],
            columns: [...(ordering.columns || [])]
        };
        this.setPendingHeaderOrderReason('row', 'header ordering restored', {
            ordering: [...this.headerOrdering.rows]
        });
        this.setPendingHeaderOrderReason('column', 'header ordering restored', {
            ordering: [...this.headerOrdering.columns]
        });
    }

    resetHeaderState() {
        this.collapsedHeaders = new Set();
        this.headerOrdering = { rows: [], columns: [] };
        this.lastRenderedHeaderSequences = { row: [], column: [] };
        this.pendingHeaderOrderReasons = {
            row: { reason: 'header state reset', details: {} },
            column: { reason: 'header state reset', details: {} }
        };
    }
}

// ===== INTEGRATION WITH MAIN APP =====

function attachAdvancedFeaturesIntegration(targetApp) {
    if (!targetApp || targetApp.__advancedFeaturesIntegrated) {
        return;
    }

    targetApp.initializeAdvancedFeatures = function() {
        if (!this.advancedFeaturesManager) {
            this.advancedFeaturesManager = new AdvancedFeaturesManager(this);
        }
        return this.advancedFeaturesManager;
    };

    targetApp.openCommentsDialog = function(item) {
        if (this.advancedFeaturesManager) {
            this.advancedFeaturesManager.openCommentsDialog(item);
        } else {
            console.warn('Advanced features not initialized');
        }
    };

    targetApp.calculateAggregations = function(gridData, bottomRightField) {
        if (this.advancedFeaturesManager) {
            return this.advancedFeaturesManager.calculateAggregations(gridData, bottomRightField);
        }
        return { rows: new Map(), columns: new Map() };
    };

    targetApp.getOrderedHeaders = function(values, headerType) {
        if (this.advancedFeaturesManager) {
            return this.advancedFeaturesManager.getOrderedHeaders(values, headerType);
        }
        return values;
    };

    targetApp.attachHeaderListeners = function(headerElement, headerValue, headerType) {
        if (this.advancedFeaturesManager) {
            this.advancedFeaturesManager.attachHeaderListeners(headerElement, headerValue, headerType);
        }
    };

    targetApp.reorderHeaders = function(sourceValue, targetValue, headerType) {
        if (this.advancedFeaturesManager) {
            return this.advancedFeaturesManager.reorderHeaders(sourceValue, targetValue, headerType);
        }
    };

    targetApp.resetHeaderState = function() {
        if (this.advancedFeaturesManager) {
            this.advancedFeaturesManager.resetHeaderState();
        }
    };

    targetApp.addComment = function() {
        if (this.advancedFeaturesManager) {
            return this.advancedFeaturesManager.addComment();
        }
    };

    targetApp.isHeaderCollapsed = function(headerValue, headerType) {
        if (this.advancedFeaturesManager) {
            return this.advancedFeaturesManager.isHeaderCollapsed(headerValue, headerType);
        }
        return false;
    };

    targetApp.formatAggregationValue = function(aggregation) {
        if (this.advancedFeaturesManager) {
            return this.advancedFeaturesManager.formatAggregationValue(aggregation);
        }
        return '';
    };

    targetApp.__advancedFeaturesIntegrated = true;
}

document.addEventListener('DOMContentLoaded', () => {
    if (typeof app !== 'undefined' && app) {
        attachAdvancedFeaturesIntegration(app);
        app.initializeAdvancedFeatures();
    }
});