// graph-view.js — Hierarchy graph view (directory-style indented tree) in a popup window
// Single composed focus graph: builds one structural tree from the focus item's root,
// with secondary parents rendered as integrated cards in left-side overlay lanes.

class GraphViewManager {
    constructor(app) {
        this.app = app;
    }

    // ===== STRUCTURAL PARENT SELECTION =====

    getIncomingParentRelations(item) {
        if (!item || !Array.isArray(item.relations)) return [];
        return item.relations
            .filter(r => r && r.type === 'parent' && r.target && r.target.itemId)
            .map(r => ({
                parentId: String(r.target.itemId),
                priority: Number.isFinite(Number(r.priority)) ? Number(r.priority) : Number.MAX_SAFE_INTEGER,
                relationId: r.relationId || '',
                type: r.type
            }))
            .sort((a, b) => {
                if (a.priority !== b.priority) return a.priority - b.priority;
                return a.parentId.localeCompare(b.parentId);
            });
    }

    getStructuralParent(item) {
        const parents = this.getIncomingParentRelations(item);
        return parents.length > 0 ? parents[0] : null;
    }

    getSecondaryParents(item) {
        const parents = this.getIncomingParentRelations(item);
        if (parents.length <= 1) return [];
        return parents.slice(1);
    }

    // ===== ROOT & REACHABILITY =====

    findRoot(item) {
        const visited = new Set();
        let current = item;
        while (current) {
            const currentId = this.app.getItemIdentity(current) || current.id;
            if (visited.has(String(currentId))) break;
            visited.add(String(currentId));
            const structural = this.getStructuralParent(current);
            if (!structural) break;
            const parentItem = this.app.getEffectiveItemById(structural.parentId);
            if (!parentItem) break;
            current = parentItem;
        }
        return current;
    }

    // ===== CHILD ENUMERATION =====

    getStructuralChildren(item) {
        if (!item || !Array.isArray(item.relations)) return [];
        const itemId = String(this.app.getItemIdentity(item) || item.id);
        const seen = new Set();
        return item.relations
            .filter(r => r && r.type === 'child' && r.target && r.target.itemId)
            .map(r => {
                const childId = String(r.target.itemId);
                const childItem = this.app.getEffectiveItemById(childId);
                return { childId, childItem, relation: r };
            })
            .filter(entry => {
                if (!entry.childItem || seen.has(entry.childId)) return false;
                const structural = this.getStructuralParent(entry.childItem);
                if (!structural || structural.parentId !== itemId) return false;
                seen.add(entry.childId);
                return true;
            })
            .sort((a, b) => {
                const orderA = Number.isFinite(Number(a.relation.order)) ? Number(a.relation.order) : Number.MAX_SAFE_INTEGER;
                const orderB = Number.isFinite(Number(b.relation.order)) ? Number(b.relation.order) : Number.MAX_SAFE_INTEGER;
                if (orderA !== orderB) return orderA - orderB;
                return a.childId.localeCompare(b.childId);
            });
    }

    getNonStructuralChildren(item) {
        if (!item || !Array.isArray(item.relations)) return [];
        const itemId = String(this.app.getItemIdentity(item) || item.id);
        const seen = new Set();
        return item.relations
            .filter(r => r && r.type === 'child' && r.target && r.target.itemId)
            .map(r => {
                const childId = String(r.target.itemId);
                const childItem = this.app.getEffectiveItemById(childId);
                return { childId, childItem, relation: r };
            })
            .filter(entry => {
                if (!entry.childItem || seen.has(entry.childId)) return false;
                const structural = this.getStructuralParent(entry.childItem);
                if (structural && structural.parentId === itemId) return false;
                seen.add(entry.childId);
                return true;
            })
            .sort((a, b) => {
                const orderA = Number.isFinite(Number(a.relation.order)) ? Number(a.relation.order) : Number.MAX_SAFE_INTEGER;
                const orderB = Number.isFinite(Number(b.relation.order)) ? Number(b.relation.order) : Number.MAX_SAFE_INTEGER;
                if (orderA !== orderB) return orderA - orderB;
                return a.childId.localeCompare(b.childId);
            });
    }

    // ===== TREE BUILDING =====

    buildTree(rootItem, globalVisited) {
        const rows = [];
        const visited = new Set();
        const walk = (item, depth, parentNode) => {
            const itemId = String(this.app.getItemIdentity(item) || item.id);
            if (visited.has(itemId)) return;
            visited.add(itemId);
            const isMirror = globalVisited ? globalVisited.has(itemId) : false;
            const mirrorOfMainRow = isMirror ? globalVisited.get(itemId) : null;
            if (globalVisited && !globalVisited.has(itemId)) globalVisited.set(itemId, null);
            const secondaryParents = isMirror ? [] : this.getSecondaryParents(item);
            const node = { item, itemId, depth, children: [], secondaryParents, row: rows.length, isMirror, mirrorOfMainRow };
            rows.push(node);
            if (parentNode) parentNode.children.push(node);
            if (!isMirror) {
                const children = this.getStructuralChildren(item);
                children.forEach(child => {
                    if (child.childItem) walk(child.childItem, depth + 1, node);
                });
                if (globalVisited) {
                    const nsChildren = this.getNonStructuralChildren(item);
                    nsChildren.forEach(child => {
                        if (child.childItem) {
                            const cid = String(this.app.getItemIdentity(child.childItem) || child.childItem.id);
                            if (globalVisited.has(cid)) walk(child.childItem, depth + 1, node);
                        }
                    });
                }
            } else if (globalVisited) {
                const children = this.getStructuralChildren(item);
                children.forEach(child => {
                    if (child.childItem) {
                        const cid = String(this.app.getItemIdentity(child.childItem) || child.childItem.id);
                        if (globalVisited.has(cid)) walk(child.childItem, depth + 1, node);
                    }
                });
            }
        };
        walk(rootItem, 0, null);

        const markLastChild = (nodes) => {
            for (const node of nodes) {
                if (node.children.length > 0) {
                    node.children.forEach((child, i) => {
                        child.isLast = i === node.children.length - 1;
                    });
                    markLastChild(node.children);
                }
            }
        };
        if (rows.length > 0) {
            rows[0].isLast = true;
            markLastChild([rows[0]]);
        }
        return rows;
    }

    // ===== SECONDARY PARENT TREE BUILDING =====

    buildSecondaryTrees(mainTreeRows, globalVisited) {
        const mainRootId = mainTreeRows.length > 0 ? mainTreeRows[0].itemId : null;
        const mainNodeIds = new Set(mainTreeRows.map(n => n.itemId));
        const connections = [];

        // Discover via secondary parents of main tree nodes (upward links)
        for (const node of mainTreeRows) {
            for (const sp of node.secondaryParents) {
                connections.push({ secParentId: sp.parentId, childNode: node });
            }
        }

        // Discover via non-structural children of main tree nodes (downward links)
        for (const node of mainTreeRows) {
            const nsChildren = this.getNonStructuralChildren(node.item);
            for (const nsc of nsChildren) {
                if (!nsc.childItem) continue;
                if (mainNodeIds.has(nsc.childId)) continue;
                connections.push({ secParentId: nsc.childId, childNode: node, isDownwardLink: true });
            }
        }

        if (connections.length === 0) return [];

        const rootGroups = new Map();
        for (const conn of connections) {
            const targetItem = this.app.getEffectiveItemById(conn.secParentId);
            if (!targetItem) continue;
            const root = this.findRoot(targetItem);
            const rootId = String(this.app.getItemIdentity(root) || root.id);
            if (rootId === mainRootId) continue;
            if (!rootGroups.has(rootId)) {
                rootGroups.set(rootId, { root, rootId, connections: [] });
            }
            rootGroups.get(rootId).connections.push(conn);
        }

        const secTrees = [];
        for (const [rootId, group] of rootGroups) {
            const treeRows = this.buildTree(group.root, globalVisited);
            if (treeRows.length === 0) continue;
            const treeConnections = [];
            for (const conn of group.connections) {
                let secNode;
                if (conn.isDownwardLink) {
                    secNode = treeRows.find(n => n.itemId === conn.secParentId && !n.isMirror);
                } else {
                    secNode = treeRows.find(n => n.itemId === conn.childNode.itemId && n.isMirror);
                }
                if (secNode) {
                    treeConnections.push({
                        mirrorRow: secNode.row,
                        childRow: conn.childNode.row,
                        spItemId: conn.secParentId,
                        childItemId: conn.childNode.itemId
                    });
                }
            }
            if (treeConnections.length === 0) continue;
            secTrees.push({ rootId, treeRows, connections: treeConnections });
        }
        return secTrees;
    }

    computeGridLayout(mainTreeRows, secTrees) {
        let mainOffset = 0;
        for (const secTree of secTrees) {
            for (const conn of secTree.connections) {
                const required = conn.mirrorRow - conn.childRow;
                if (required > mainOffset) mainOffset = required;
            }
        }
        for (const node of mainTreeRows) {
            node.gridRow = mainOffset + node.row + 1;
        }
        for (const secTree of secTrees) {
            const topConn = secTree.connections.reduce((best, c) =>
                c.mirrorRow < best.mirrorRow ? c : best, secTree.connections[0]);
            const secOffset = topConn.childRow + mainOffset - topConn.mirrorRow;
            secTree.offset = secOffset;
            for (const node of secTree.treeRows) {
                node.gridRow = secOffset + node.row + 1;
            }
            for (const conn of secTree.connections) {
                conn.gridRow = mainOffset + conn.childRow + 1;
            }
        }
        let totalGridRows = 0;
        for (const node of mainTreeRows) {
            if (node.gridRow > totalGridRows) totalGridRows = node.gridRow;
        }
        for (const secTree of secTrees) {
            for (const node of secTree.treeRows) {
                if (node.gridRow > totalGridRows) totalGridRows = node.gridRow;
            }
        }
        return totalGridRows;
    }

    // ===== POPUP RENDERING =====

    openGraphView(item) {
        const focusItemId = String(this.app.getItemIdentity(item) || item.id);
        const rootItem = this.findRoot(item);
        const treeRows = this.buildTree(rootItem);

        if (treeRows.length === 0) {
            this.app.showNotification('No hierarchy found for this item', 'warning');
            return false;
        }

        const globalVisited = new Map();
        treeRows.forEach(n => globalVisited.set(n.itemId, n.row));
        const secTrees = this.buildSecondaryTrees(treeRows, globalVisited);
        const totalGridRows = this.computeGridLayout(treeRows, secTrees);

        const secCount = secTrees.length;
        const gapCol = secCount > 0 ? secCount + 1 : 0;
        const mainCol = secCount > 0 ? secCount + 2 : 1;
        const totalColumns = mainCol;

        const serializeTreeRows = (rows) => rows.map(n => ({
            itemId: n.itemId,
            title: n.item.title || n.itemId,
            id: n.item.id || n.itemId,
            status: n.item.status || '',
            type: n.item.type || '',
            assignee: n.item.assignee || '',
            depth: n.depth,
            row: n.row,
            gridRow: n.gridRow,
            isLast: !!n.isLast,
            isMirror: !!n.isMirror,
            mirrorOfMainRow: n.mirrorOfMainRow != null ? n.mirrorOfMainRow : null,
            secondaryParentIds: (n.secondaryParents || []).map(sp => sp.parentId)
        }));

        const payload = {
            mainTree: {
                rows: serializeTreeRows(treeRows),
                gridColumn: mainCol
            },
            secondaryTrees: secTrees.map((st, i) => ({
                rootId: st.rootId,
                rows: serializeTreeRows(st.treeRows),
                gridColumn: i + 1,
                connections: st.connections.map(c => ({
                    gridRow: c.gridRow,
                    spItemId: c.spItemId,
                    childItemId: c.childItemId
                }))
            })),
            focusItemId,
            rootTitle: rootItem.title || rootItem.id || 'Hierarchy',
            totalGridRows,
            totalColumns,
            gapCol
        };

        const graphWindow = window.open('', '_blank', 'width=1200,height=800');
        if (!graphWindow) {
            this.app.showNotification('Popup blocked while opening graph view', 'warning');
            return false;
        }

        const payloadJson = JSON.stringify(payload).replace(/</g, '\\u003c');
        const title = 'Graph View — ' + (payload.rootTitle || 'Hierarchy');
        graphWindow.document.open();
        graphWindow.document.write(this.buildPopupHtml(title, payloadJson));
        graphWindow.document.close();
        return true;
    }

    buildPopupHtml(title, payloadJson) {
        return '<!DOCTYPE html>\n'
            + '<html lang="en">\n'
            + '<head>\n'
            + '    <meta charset="UTF-8">\n'
            + '    <title>' + this.escapeHtml(title) + '</title>\n'
            + '    <style>\n'
            + this.getPopupCss()
            + '    </style>\n'
            + '</head>\n'
            + '<body>\n'
            + '    <div class="shell">\n'
            + '        <div class="header">\n'
            + '            <h1 class="title">' + this.escapeHtml(title) + '</h1>\n'
            + '        </div>\n'
            + '        <div class="graph-content" id="graph-content"></div>\n'
            + '    </div>\n'
            + '    <script>\n'
            + '        const graphPayload = ' + payloadJson + ';\n'
            + this.getPopupScript()
            + '    </script>\n'
            + '</body>\n'
            + '</html>';
    }

    escapeHtml(text) {
        return String(text)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    getPopupCss() {
        return `
        :root {
            color-scheme: light;
            font-family: 'Segoe UI', Tahoma, sans-serif;
        }
        * { box-sizing: border-box; }
        body {
            margin: 0;
            padding: 24px;
            background: #f4f6f8;
            color: #1f2937;
        }
        .shell {
            margin: 0 auto;
            background: #ffffff;
            border: 1px solid #d0d7de;
            border-radius: 16px;
            box-shadow: 0 12px 32px rgba(15, 23, 42, 0.12);
            overflow: hidden;
            min-width: 600px;
        }
        .header {
            padding: 20px 28px 14px;
            border-bottom: 1px solid #e5e7eb;
            background: linear-gradient(180deg, #ffffff 0%, #f8fafc 100%);
        }
        .title {
            margin: 0;
            font-size: 20px;
            font-weight: 700;
            color: #111827;
        }
        .graph-content {
            padding: 24px;
            overflow: auto;
            max-height: calc(100vh - 120px);
        }

        /* ===== CSS Grid layout ===== */
        .graph-grid {
            display: grid;
            column-gap: 48px;
        }

        /* Horizontal connector in gap column */
        .tree-connector {
            position: relative;
            pointer-events: none;
        }
        .tree-connector::before {
            content: '';
            position: absolute;
            top: 28px;
            left: -24px;
            right: -24px;
            border-top: 2px dashed #6366f1;
        }

        /* Secondary parent card accent */
        .node-card.sec-parent-card {
            border-left: 3px solid #6366f1;
        }

        /* Primary tree row cells */
        .tree-row {
            display: flex;
            align-items: flex-start;
            padding-bottom: 20px;
            position: relative;
            min-height: 56px;
        }
        .indent-guides {
            display: flex;
            flex-shrink: 0;
            align-self: stretch;
        }
        .indent-segment {
            width: 50px;
            position: relative;
            flex-shrink: 0;
            align-self: stretch;
        }
        /* Continuing vertical line (ancestor still has siblings below) */
        .indent-segment.has-line::before {
            content: '';
            position: absolute;
            left: 10px;
            top: 0;
            bottom: 0;
            width: 1px;
            background: #d0d7de;
        }
        /* Branch connector (non-last child): vertical line from top to arm */
        .indent-segment.is-branch::after {
            content: '';
            position: absolute;
            left: 10px;
            top: 0;
            height: 28px;
            width: 1px;
            background: #d0d7de;
        }
        .indent-segment.is-branch .branch-arm {
            position: absolute;
            left: 10px;
            top: 28px;
            width: 30px;
            height: 0;
            border-top: 1px solid #d0d7de;
        }
        /* Elbow connector (last child): vertical line from top to arm, no continuation */
        .indent-segment.is-elbow::after {
            content: '';
            position: absolute;
            left: 10px;
            top: 0;
            height: 28px;
            width: 1px;
            background: #d0d7de;
        }
        .indent-segment.is-elbow .branch-arm {
            position: absolute;
            left: 10px;
            top: 28px;
            width: 30px;
            height: 0;
            border-top: 1px solid #d0d7de;
        }

        /* Node cards */
        .node-card {
            display: inline-flex;
            flex-direction: column;
            gap: 2px;
            width: 220px;
            padding: 8px 12px;
            background: #ffffff;
            border: 1px solid #d0d7de;
            border-radius: 8px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.06);
            flex-shrink: 0;
            cursor: pointer;
            transition: box-shadow 0.15s, border-color 0.15s;
        }
        .node-card:hover {
            border-color: #93c5fd;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        .node-card.is-focus {
            box-shadow: 0 0 0 2px #1a73e8, 0 2px 8px rgba(0,0,0,0.12);
            border-color: #1a73e8;
        }
        .node-card .card-id {
            font-size: 10px;
            font-weight: 600;
            color: #6b7280;
            text-transform: uppercase;
            letter-spacing: 0.03em;
        }
        .node-card .card-title {
            font-size: 13px;
            font-weight: 600;
            color: #111827;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        .node-card .card-meta {
            display: flex;
            gap: 8px;
            font-size: 10px;
            color: #6b7280;
        }
        .node-card .card-meta .chip {
            display: inline-block;
            padding: 1px 6px;
            border-radius: 3px;
            background: #f3f4f6;
            border: 1px solid #e5e7eb;
            font-size: 10px;
            white-space: nowrap;
        }
        .node-card .card-meta .chip.status-done { background: #d1fae5; border-color: #a7f3d0; color: #065f46; }
        .node-card .card-meta .chip.status-progress { background: #dbeafe; border-color: #bfdbfe; color: #1e40af; }
        .node-card .card-meta .chip.status-review { background: #fef3c7; border-color: #fde68a; color: #92400e; }
        .node-card .card-meta .chip.status-blocked { background: #fee2e2; border-color: #fecaca; color: #991b1b; }
        .node-card .card-meta .chip.status-backlog { background: #f3f4f6; border-color: #d1d5db; color: #374151; }

        /* Mirror nodes */
        .node-card.is-mirror {
            opacity: 0.45;
            border-style: dashed;
            border-color: #9ca3af;
            background: #f9fafb;
            cursor: default;
        }
        .node-card.is-mirror:hover {
            border-color: #9ca3af;
            box-shadow: 0 1px 3px rgba(0,0,0,0.06);
        }
        .mirror-label {
            font-size: 9px;
            color: #9ca3af;
            font-style: italic;
            margin-top: 1px;
        }
        `;
    }

    getPopupScript() {
        return `
        var content = document.getElementById('graph-content');
        var mainTree = graphPayload.mainTree;
        var secondaryTrees = graphPayload.secondaryTrees;
        var focusItemId = graphPayload.focusItemId;
        var totalGridRows = graphPayload.totalGridRows;
        var totalColumns = graphPayload.totalColumns;
        var gapCol = graphPayload.gapCol;

        function buildContinuingLines(rows) {
            var result = [];
            for (var i = 0; i < rows.length; i++) { result.push([]); }
            for (var i = 0; i < rows.length; i++) {
                var node = rows[i];
                var lines = [];
                for (var d = 1; d <= node.depth; d++) {
                    var continues = false;
                    for (var j = i + 1; j < rows.length; j++) {
                        if (rows[j].depth < d) break;
                        if (rows[j].depth === d) { continues = true; break; }
                    }
                    lines.push(continues);
                }
                result[i] = lines;
            }
            return result;
        }

        function renderTreeColumn(grid, rows, gridColumn, focusId, ancestorContinues, secParentIds) {
            for (var idx = 0; idx < rows.length; idx++) {
                var node = rows[idx];
                var treeCell = document.createElement('div');
                treeCell.className = 'tree-row';
                treeCell.style.gridRow = String(node.gridRow);
                treeCell.style.gridColumn = String(gridColumn);

                if (node.depth > 0) {
                    var guidesDiv = document.createElement('div');
                    guidesDiv.className = 'indent-guides';
                    for (var d = 1; d <= node.depth; d++) {
                        var seg = document.createElement('div');
                        seg.className = 'indent-segment';
                        if (d < node.depth) {
                            if (ancestorContinues[idx][d - 1]) seg.classList.add('has-line');
                        } else {
                            if (node.isLast) {
                                seg.classList.add('is-elbow');
                            } else {
                                seg.classList.add('is-branch');
                                seg.classList.add('has-line');
                            }
                            var arm = document.createElement('div');
                            arm.className = 'branch-arm';
                            seg.appendChild(arm);
                        }
                        guidesDiv.appendChild(seg);
                    }
                    treeCell.appendChild(guidesDiv);
                }

                var card = document.createElement('div');
                card.className = 'node-card';
                if (node.itemId === focusId) card.classList.add('is-focus');
                if (node.isMirror) card.classList.add('is-mirror');
                if (secParentIds && secParentIds.indexOf(node.itemId) !== -1) card.classList.add('sec-parent-card');
                card.dataset.itemId = node.itemId;
                if (!node.isMirror) {
                    card.addEventListener('click', function() {
                        var id = this.dataset.itemId;
                        if (window.opener && window.opener.app && window.opener.app.focusItemInGrid) {
                            window.opener.app.focusItemInGrid(id, { select: true, suppressWarning: true });
                            window.opener.focus();
                        }
                    });
                }
                var statusClass = getStatusClass(node.status);
                card.innerHTML = '<div class="card-id">' + escapeHtml(node.id) + '</div>'
                    + '<div class="card-title" title="' + escapeHtml(node.title) + '">' + escapeHtml(node.title) + '</div>'
                    + '<div class="card-meta">'
                    + (node.type ? '<span class="chip">' + escapeHtml(node.type) + '</span>' : '')
                    + (node.status ? '<span class="chip ' + statusClass + '">' + escapeHtml(node.status) + '</span>' : '')
                    + (node.assignee ? '<span>' + escapeHtml(node.assignee) + '</span>' : '')
                    + '</div>'
                    + (node.isMirror ? '<div class="mirror-label">\u2192 see main tree</div>' : '');
                treeCell.appendChild(card);
                grid.appendChild(treeCell);
            }
        }

        // Build grid
        var secCount = secondaryTrees.length;
        var grid = document.createElement('div');
        grid.className = 'graph-grid';
        if (secCount > 0) {
            grid.style.gridTemplateColumns = 'repeat(' + secCount + ', 280px) 48px minmax(400px, 1fr)';
        } else {
            grid.style.gridTemplateColumns = 'minmax(400px, 1fr)';
        }

        // Render secondary trees in left columns
        for (var si = 0; si < secCount; si++) {
            var secTree = secondaryTrees[si];
            var secAncestors = buildContinuingLines(secTree.rows);
            var spIds = [];
            for (var ci = 0; ci < secTree.connections.length; ci++) {
                spIds.push(secTree.connections[ci].spItemId);
            }
            renderTreeColumn(grid, secTree.rows, secTree.gridColumn, focusItemId, secAncestors, spIds);
        }

        // Render connector cells in gap column
        if (gapCol > 0) {
            for (var si = 0; si < secCount; si++) {
                var secTree = secondaryTrees[si];
                for (var ci = 0; ci < secTree.connections.length; ci++) {
                    var conn = secTree.connections[ci];
                    var connCell = document.createElement('div');
                    connCell.className = 'tree-connector';
                    connCell.style.gridRow = String(conn.gridRow);
                    connCell.style.gridColumn = String(gapCol);
                    grid.appendChild(connCell);
                }
            }
        }

        // Render main tree
        var mainAncestors = buildContinuingLines(mainTree.rows);
        renderTreeColumn(grid, mainTree.rows, mainTree.gridColumn, focusItemId, mainAncestors, null);

        content.appendChild(grid);

        function getStatusClass(status) {
            if (!status) return '';
            var s = status.toLowerCase();
            if (s === 'done') return 'status-done';
            if (s.indexOf('progress') !== -1) return 'status-progress';
            if (s.indexOf('review') !== -1) return 'status-review';
            if (s === 'blocked') return 'status-blocked';
            if (s === 'backlog') return 'status-backlog';
            return '';
        }

        function escapeHtml(text) {
            var d = document.createElement('div');
            d.textContent = text || '';
            return d.innerHTML;
        }
        `;
    }
}

// Initialize when DOM is ready, after app is available
document.addEventListener('DOMContentLoaded', () => {
    const waitForApp = setInterval(() => {
        if (window.app) {
            clearInterval(waitForApp);
            window.app.graphViewManager = new GraphViewManager(window.app);
        }
    }, 50);
});
