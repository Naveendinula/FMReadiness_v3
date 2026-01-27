// FM Readiness - Web UI Application

// Global state
window.auditData = {
    summary: null,
    rows: []
};

// DOM Elements
const overallReadinessEl = document.getElementById('overallReadiness');
const fullyReadyCountEl = document.getElementById('fullyReadyCount');
const fullyReadySublabelEl = document.getElementById('fullyReadySublabel');
const missingCountEl = document.getElementById('missingCount');
const missingSublabelEl = document.getElementById('missingSublabel');
const totalCountEl = document.getElementById('totalCount');
const groupScoresEl = document.getElementById('groupScores');
const searchInput = document.getElementById('searchInput');
const showMissingOnly = document.getElementById('showMissingOnly');
const tableBody = document.getElementById('tableBody');
const resultsTable = document.getElementById('resultsTable');
const emptyState = document.getElementById('emptyState');

// Cache of elementId -> views[]
const viewsByElementId = new Map();

// Initialize WebView2 message listener
if (window.chrome && window.chrome.webview) {
    window.chrome.webview.addEventListener('message', (event) => {
        const msg = event.data;
        if (msg && msg.type === 'auditResults') {
            receiveAudit(msg);
            return;
        }

        if (msg && msg.type === '2dViewOptions') {
            receive2dViewOptions(msg);
        }
    });
}

// Also expose global function for ExecuteScriptAsync fallback
window.receiveAudit = receiveAudit;

/**
 * Receive audit results from C#
 */
function receiveAudit(data) {
    window.auditData = {
        summary: data.summary,
        rows: data.rows || []
    };
    render();
}

/**
 * Main render function
 */
function render() {
    const { summary, rows } = window.auditData;

    // Update summary cards
    if (summary) {
        const pct = summary.overallReadinessPercent || 0;
        overallReadinessEl.textContent = `${Math.round(pct)}%`;
        overallReadinessEl.className = 'card-value ' + getReadinessClass(pct);

        const fullyReady = summary.fullyReady || 0;
        const totalAudited = summary.totalAudited || 0;
        const missing = totalAudited - fullyReady;

        fullyReadyCountEl.textContent = fullyReady;
        fullyReadySublabelEl.textContent = `${fullyReady} assets complete`;

        missingCountEl.textContent = missing;
        missingSublabelEl.textContent = `${missing} assets incomplete`;

        totalCountEl.textContent = totalAudited;

        // Render group score badges
        renderGroupScores(summary.groupScores || {});
    } else {
        overallReadinessEl.textContent = '--%';
        overallReadinessEl.className = 'card-value';
        fullyReadyCountEl.textContent = '--';
        fullyReadySublabelEl.textContent = '-- assets complete';
        missingCountEl.textContent = '--';
        missingSublabelEl.textContent = '-- assets incomplete';
        totalCountEl.textContent = '--';
        groupScoresEl.innerHTML = '';
    }

    // Filter rows
    const filteredRows = filterRows(rows);

    // Update table
    if (filteredRows.length > 0) {
        resultsTable.classList.add('visible');
        emptyState.classList.add('hidden');
        renderTable(filteredRows);
    } else if (rows.length > 0) {
        // Has data but all filtered out
        resultsTable.classList.add('visible');
        emptyState.classList.add('hidden');
        renderTable([]);
    } else {
        resultsTable.classList.remove('visible');
        emptyState.classList.remove('hidden');
    }
}

/**
 * Filter rows based on search and checkbox
 */
function filterRows(rows) {
    const searchTerm = searchInput.value.toLowerCase().trim();
    const missingOnly = showMissingOnly.checked;

    return rows.filter(row => {
        // Filter by missing only
        if (missingOnly && row.missingCount === 0) {
            return false;
        }

        // Filter by search term
        if (searchTerm) {
            const searchFields = [
                row.category,
                row.family,
                row.type,
                row.missingParams,
                String(row.elementId)
            ].map(f => (f || '').toLowerCase());

            const matches = searchFields.some(field => field.includes(searchTerm));
            if (!matches) return false;
        }

        return true;
    });
}

/**
 * Render group score badges in the summary area
 */
function renderGroupScores(groupScores) {
    if (!groupScoresEl) return;

    const groups = Object.entries(groupScores || {});
    if (groups.length === 0) {
        groupScoresEl.innerHTML = '';
        return;
    }

    groupScoresEl.innerHTML = groups.map(([name, pct]) => {
        const scoreClass = getScoreClass(pct);
        return `<div class="group-badge ${scoreClass}">
            <span class="group-name">${escapeHtml(name)}</span>
            <span class="group-pct">${pct}%</span>
        </div>`;
    }).join('');
}

/**
 * Get score class for badges
 */
function getScoreClass(percent) {
    if (percent >= 100) return 'score-100';
    if (percent >= 75) return 'score-high';
    if (percent >= 50) return 'score-medium';
    return 'score-low';
}

/**
 * Render inline group badges for a row
 */
function renderRowGroupBadges(groupScores) {
    if (!groupScores || Object.keys(groupScores).length === 0) {
        return '-';
    }

    return Object.entries(groupScores).map(([name, pct]) => {
        const scoreClass = getScoreClass(pct);
        // Abbreviate group names for inline display
        const abbrev = name.substring(0, 3);
        return `<span class="group-badge-sm ${scoreClass}" title="${escapeHtml(name)}: ${pct}%">${abbrev} ${pct}%</span>`;
    }).join(' ');
}

/**
 * Render table rows
 */
function renderTable(rows) {
    tableBody.innerHTML = '';

    if (rows.length === 0) {
        const emptyRow = document.createElement('tr');
        emptyRow.innerHTML = '<td colspan="10" style="text-align: center; color: #a3a3a3; padding: 48px;">No matching results</td>';
        tableBody.appendChild(emptyRow);
        return;
    }

    rows.forEach(row => {
        const tr = document.createElement('tr');
        const readinessClass = getReadinessClass(row.readinessPercent);
        const groupBadgesHtml = renderRowGroupBadges(row.groupScores);
        const statusBadge = getStatusBadge(row.readinessPercent, row.missingCount);

        tr.innerHTML = `
            <td>${row.elementId}</td>
            <td>${statusBadge}</td>
            <td>${escapeHtml(row.category)}</td>
            <td>${escapeHtml(row.family)}</td>
            <td>${escapeHtml(row.type)}</td>
            <td class="${readinessClass}">${row.readinessPercent}%</td>
            <td class="group-badges-inline">${groupBadgesHtml}</td>
            <td class="missing-params" title="${escapeHtml(row.missingParams)}">${escapeHtml(row.missingParams) || '-'}</td>
            <td>
                <select class="view-select" data-element-id="${row.elementId}" disabled>
                    <option>Select view</option>
                </select>
                <button class="open-view-btn" data-element-id="${row.elementId}" disabled>Open</button>
            </td>
            <td><button class="select-btn" data-id="${row.elementId}">Select</button></td>
        `;

        tableBody.appendChild(tr);
    });
}

/**
 * Get status badge HTML
 */
function getStatusBadge(readinessPercent, missingCount) {
    if (readinessPercent >= 100) {
        return '<span class="status-badge ready">Ready</span>';
    } else if (missingCount > 0 && readinessPercent > 0) {
        return '<span class="status-badge partial">Partial</span>';
    } else {
        return '<span class="status-badge missing">Missing</span>';
    }
}

/**
 * Get CSS class for readiness percentage
 */
function getReadinessClass(percent) {
    if (percent >= 100) return 'readiness-100';
    if (percent >= 75) return 'readiness-high';
    if (percent >= 50) return 'readiness-medium';
    return 'readiness-low';
}

/**
 * Escape HTML to prevent XSS
 */
function escapeHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

/**
 * Send message to C# (Select + Zoom)
 */
function selectElement(elementId) {
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage({
            type: 'selectZoom',
            elementId: elementId
        });
    }
}

function request2dViews(elementId) {
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage({
            action: 'get2dViews',
            elementId: elementId
        });
    }
}

function open2dView(elementId, viewId) {
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage({
            action: 'open2dView',
            elementId: elementId,
            viewId: viewId
        });
    }
}

function receive2dViewOptions(msg) {
    const elementId = Number(msg.elementId);
    const views = Array.isArray(msg.views) ? msg.views : [];
    if (!Number.isFinite(elementId)) return;

    viewsByElementId.set(elementId, views);

    const select = tableBody.querySelector(`select.view-select[data-element-id="${elementId}"]`);
    if (!select) return;

    const openBtn = tableBody.querySelector(`button.open-view-btn[data-element-id="${elementId}"]`);

    select.innerHTML = '';
    const placeholder = document.createElement('option');
    placeholder.value = '';
    placeholder.textContent = views.length ? 'Select 2D view' : 'No 2D views found';
    select.appendChild(placeholder);

    views.forEach(v => {
        const opt = document.createElement('option');
        opt.value = String(v.viewId);
        opt.textContent = v.name;
        select.appendChild(opt);
    });

    select.disabled = false;
    if (openBtn) {
        openBtn.disabled = true;
    }
}

// Event Listeners

// Search input
searchInput.addEventListener('input', () => {
    render();
});

// Show missing only checkbox
showMissingOnly.addEventListener('change', () => {
    render();
});

// Table row click (event delegation)
tableBody.addEventListener('click', (e) => {
    const btn = e.target.closest('.select-btn');
    if (btn) {
        const elementId = parseInt(btn.dataset.id, 10);
        if (!isNaN(elementId)) {
            selectElement(elementId);
        }
    }

    // Clicking a row (but not the dropdown/select button) triggers 2D view request
    const row = e.target.closest('tr');
    if (row && !e.target.closest('.select-btn') && !e.target.closest('.view-select')) {
        const firstCell = row.querySelector('td');
        const elementId = firstCell ? parseInt(firstCell.textContent, 10) : NaN;
        if (!isNaN(elementId) && !viewsByElementId.has(elementId)) {
            request2dViews(elementId);
        }
    }
});

// Dropdown selection enables the Open button (event delegation)
tableBody.addEventListener('change', (e) => {
    const select = e.target.closest('select.view-select');
    if (!select) return;

    const elementId = parseInt(select.dataset.elementId, 10);
    if (isNaN(elementId)) return;

    const openBtn = tableBody.querySelector(`button.open-view-btn[data-element-id="${elementId}"]`);
    if (!openBtn) return;

    const viewId = parseInt(select.value, 10);
    openBtn.disabled = isNaN(viewId);
});

// Open button click sends open2dView (event delegation)
tableBody.addEventListener('click', (e) => {
    const openBtn = e.target.closest('button.open-view-btn');
    if (!openBtn) return;

    const elementId = parseInt(openBtn.dataset.elementId, 10);
    if (isNaN(elementId)) return;

    const select = tableBody.querySelector(`select.view-select[data-element-id="${elementId}"]`);
    if (!select) return;

    const viewId = parseInt(select.value, 10);
    if (!isNaN(viewId)) {
        open2dView(elementId, viewId);
    }
});

// ========================================
// Tab Navigation
// ========================================
document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        // Update active tab button
        document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        
        // Show corresponding tab content
        const tabId = btn.dataset.tab;
        document.querySelectorAll('.tab-content').forEach(content => {
            content.classList.remove('active');
        });
        document.getElementById(`tab-${tabId}`).classList.add('active');
    });
});

// ========================================
// Parameter Editor State
// ========================================
let editorState = {
    selectedElements: [],
    currentTypeId: null,
    currentTypeInfo: null
};

// ========================================
// Parameter Editor - Message Handling
// ========================================
function handleEditorMessage(msg) {
    switch (msg.type) {
        case 'selectedElementsData':
            receiveSelectedElementsData(msg);
            break;
        case 'categoryStats':
            receiveCategoryStats(msg);
            break;
        case 'operationResult':
            receiveOperationResult(msg);
            break;
    }
}

// Extend the existing message listener
if (window.chrome && window.chrome.webview) {
    window.chrome.webview.addEventListener('message', (event) => {
        const msg = event.data;
        // Handle editor messages
        if (msg && ['selectedElementsData', 'categoryStats', 'operationResult'].includes(msg.type)) {
            handleEditorMessage(msg);
        }
    });
}

// ========================================
// Parameter Editor - C# Communication
// ========================================
function requestSelectedElements() {
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage({
            action: 'getSelectedElements'
        });
    }
}

function requestCategoryStats(category) {
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage({
            action: 'getCategoryStats',
            category: category
        });
    }
}

function setInstanceParams(elementIds, params) {
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage({
            action: 'setInstanceParams',
            elementIds: elementIds,
            params: params
        });
    }
}

function setCategoryParams(category, params, onlyBlanks) {
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage({
            action: 'setCategoryParams',
            category: category,
            params: params,
            onlyBlanks: onlyBlanks
        });
    }
}

function setTypeParams(typeId, params) {
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage({
            action: 'setTypeParams',
            typeId: typeId,
            params: params
        });
    }
}

function copyComputedToParam(elementIds, sourceField, targetParam) {
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage({
            action: 'copyComputedToParam',
            elementIds: elementIds,
            sourceField: sourceField,
            targetParam: targetParam
        });
    }
}

// ========================================
// Parameter Editor - UI Updates
// ========================================
function receiveSelectedElementsData(msg) {
    const elements = msg.elements || [];
    editorState.selectedElements = elements;
    
    const infoEl = document.getElementById('selectedElementInfo');
    const formEl = document.getElementById('instanceParamsForm');
    const typeInfoEl = document.getElementById('typeInfo');
    const typeFormEl = document.getElementById('typeParamsForm');
    
    if (elements.length === 0) {
        infoEl.innerHTML = '<p class="placeholder-text">No elements selected. Select elements in Revit and click Refresh.</p>';
        formEl.classList.add('hidden');
        typeInfoEl.innerHTML = '<p class="placeholder-text">Select an element to edit its type parameters</p>';
        typeFormEl.classList.add('hidden');
        clearComputedValues();
        return;
    }
    
    // Show element badges
    let badgesHtml = elements.map(el => `
        <span class="element-badge">
            <span class="element-id">${el.elementId}</span>
            <span>${el.category} - ${el.family}</span>
        </span>
    `).join('');
    
    if (elements.length > 5) {
        badgesHtml = elements.slice(0, 5).map(el => `
            <span class="element-badge">
                <span class="element-id">${el.elementId}</span>
                <span>${el.category}</span>
            </span>
        `).join('') + `<span class="element-badge">+${elements.length - 5} more</span>`;
    }
    
    infoEl.innerHTML = badgesHtml;
    formEl.classList.remove('hidden');
    
    // Pre-fill instance params if single element
    if (elements.length === 1) {
        const el = elements[0];
        const ip = el.instanceParams || {};
        document.getElementById('edit_FM_Barcode').value = ip.FM_Barcode || '';
        document.getElementById('edit_FM_UniqueAssetId').value = ip.FM_UniqueAssetId || '';
        document.getElementById('edit_FM_InstallationDate').value = ip.FM_InstallationDate || '';
        document.getElementById('edit_FM_WarrantyStart').value = ip.FM_WarrantyStart || '';
        document.getElementById('edit_FM_WarrantyEnd').value = ip.FM_WarrantyEnd || '';
        document.getElementById('edit_FM_Criticality').value = ip.FM_Criticality || '';
        document.getElementById('edit_FM_Trade').value = ip.FM_Trade || '';
        document.getElementById('edit_FM_PMTemplateId').value = ip.FM_PMTemplateId || '';
        document.getElementById('edit_FM_PMFrequencyDays').value = ip.FM_PMFrequencyDays || '';
        document.getElementById('edit_FM_Building').value = ip.FM_Building || '';
        document.getElementById('edit_FM_LocationSpace').value = ip.FM_LocationSpace || '';
        
        // Update type info
        if (el.typeId && el.typeName) {
            editorState.currentTypeId = el.typeId;
            editorState.currentTypeInfo = el;
            typeInfoEl.innerHTML = `
                <div class="type-badge">
                    <span class="type-name">${el.typeName}</span>
                    <span>(${el.typeInstanceCount || '?'} instances)</span>
                </div>
            `;
            typeFormEl.classList.remove('hidden');
            
            const tp = el.typeParams || {};
            document.getElementById('type_Manufacturer').value = tp.Manufacturer || '';
            document.getElementById('type_Model').value = tp.Model || '';
            document.getElementById('type_TypeMark').value = tp.TypeMark || '';
            document.getElementById('typeImpact').textContent = `Will update ${el.typeInstanceCount || '?'} instances`;
        }
        
        // Update computed values
        const computed = el.computed || {};
        document.getElementById('computed_UniqueId').textContent = computed.UniqueId || '--';
        document.getElementById('computed_Level').textContent = computed.Level || '--';
        document.getElementById('computed_RoomSpace').textContent = computed.RoomSpace || '--';
    } else {
        // Multiple selection - clear fields
        clearInstanceForm();
        typeInfoEl.innerHTML = '<p class="placeholder-text">Select a single element to edit type parameters</p>';
        typeFormEl.classList.add('hidden');
        clearComputedValues();
    }
}

function receiveCategoryStats(msg) {
    const countEl = document.getElementById('categoryCount');
    if (msg.count !== undefined) {
        countEl.textContent = `${msg.count} elements`;
    } else {
        countEl.textContent = '';
    }
}

function receiveOperationResult(msg) {
    showToast(msg.message, msg.success ? 'success' : 'error');
    
    // Refresh selection after parameter updates
    if (msg.success && msg.refreshAudit) {
        requestSelectedElements();
        // Trigger audit refresh to update the Audit Results tab
        requestAuditRefresh();
    }
}

/**
 * Request an audit refresh from C# to update the Audit Results tab
 * after parameter changes in the editor.
 */
function requestAuditRefresh() {
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage({
            action: 'refreshAudit'
        });
    }
}

function clearInstanceForm() {
    document.getElementById('edit_FM_Barcode').value = '';
    document.getElementById('edit_FM_UniqueAssetId').value = '';
    document.getElementById('edit_FM_InstallationDate').value = '';
    document.getElementById('edit_FM_WarrantyStart').value = '';
    document.getElementById('edit_FM_WarrantyEnd').value = '';
    document.getElementById('edit_FM_Criticality').value = '';
    document.getElementById('edit_FM_Trade').value = '';
    document.getElementById('edit_FM_PMTemplateId').value = '';
    document.getElementById('edit_FM_PMFrequencyDays').value = '';
    document.getElementById('edit_FM_Building').value = '';
    document.getElementById('edit_FM_LocationSpace').value = '';
}

function clearComputedValues() {
    document.getElementById('computed_UniqueId').textContent = '--';
    document.getElementById('computed_Level').textContent = '--';
    document.getElementById('computed_RoomSpace').textContent = '--';
}

function showToast(message, type = 'info') {
    const existing = document.querySelector('.toast');
    if (existing) existing.remove();
    
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.textContent = message;
    document.body.appendChild(toast);
    
    setTimeout(() => toast.remove(), 3000);
}

// ========================================
// Parameter Editor - Event Handlers
// ========================================

// Refresh selection button
document.getElementById('refreshSelection')?.addEventListener('click', () => {
    requestSelectedElements();
});

// Apply to Selected button
document.getElementById('applyToSelected')?.addEventListener('click', () => {
    if (editorState.selectedElements.length === 0) {
        showToast('No elements selected', 'error');
        return;
    }
    
    const params = {};
    const fields = ['FM_Barcode', 'FM_UniqueAssetId', 'FM_InstallationDate', 'FM_WarrantyStart', 
                    'FM_WarrantyEnd', 'FM_Criticality', 'FM_Trade', 'FM_PMTemplateId', 
                    'FM_PMFrequencyDays', 'FM_Building', 'FM_LocationSpace'];
    
    fields.forEach(field => {
        const el = document.getElementById(`edit_${field}`);
        if (el && el.value.trim() !== '') {
            params[field] = el.value.trim();
        }
    });
    
    if (Object.keys(params).length === 0) {
        showToast('No values to apply', 'error');
        return;
    }
    
    const elementIds = editorState.selectedElements.map(e => e.elementId);
    setInstanceParams(elementIds, params);
});

// Category selection change
document.getElementById('bulkCategory')?.addEventListener('change', (e) => {
    const category = e.target.value;
    if (category) {
        requestCategoryStats(category);
    } else {
        document.getElementById('categoryCount').textContent = '';
    }
});

// Apply to Category button
document.getElementById('applyToCategory')?.addEventListener('click', () => {
    const category = document.getElementById('bulkCategory').value;
    if (!category) {
        showToast('Select a category first', 'error');
        return;
    }
    
    const params = {};
    const fields = ['FM_Building', 'FM_Trade', 'FM_PMFrequencyDays', 'FM_Criticality'];
    
    fields.forEach(field => {
        const el = document.getElementById(`bulk_${field}`);
        if (el && el.value.trim() !== '') {
            params[field] = el.value.trim();
        }
    });
    
    if (Object.keys(params).length === 0) {
        showToast('No values to apply', 'error');
        return;
    }
    
    const onlyBlanks = document.getElementById('onlyFillBlanks').checked;
    setCategoryParams(category, params, onlyBlanks);
});

// Update Type button
document.getElementById('updateType')?.addEventListener('click', () => {
    if (!editorState.currentTypeId) {
        showToast('No type selected', 'error');
        return;
    }
    
    const params = {};
    const manufacturer = document.getElementById('type_Manufacturer').value.trim();
    const model = document.getElementById('type_Model').value.trim();
    const typeMark = document.getElementById('type_TypeMark').value.trim();
    
    if (manufacturer) params.Manufacturer = manufacturer;
    if (model) params.Model = model;
    if (typeMark) params.TypeMark = typeMark;
    
    if (Object.keys(params).length === 0) {
        showToast('No values to update', 'error');
        return;
    }
    
    setTypeParams(editorState.currentTypeId, params);
});

// Copy UniqueId to clipboard
document.getElementById('copyUniqueId')?.addEventListener('click', () => {
    const uniqueId = document.getElementById('computed_UniqueId').textContent;
    if (uniqueId && uniqueId !== '--') {
        navigator.clipboard.writeText(uniqueId).then(() => {
            showToast('Copied to clipboard!', 'success');
        });
    }
});

// Copy UniqueId to FM_UniqueAssetId
document.getElementById('copyUniqueIdToParam')?.addEventListener('click', () => {
    if (editorState.selectedElements.length === 0) {
        showToast('No elements selected', 'error');
        return;
    }
    
    const elementIds = editorState.selectedElements.map(e => e.elementId);
    copyComputedToParam(elementIds, 'UniqueId', 'FM_UniqueAssetId');
});

// Populate FM_LocationSpace from Room
document.getElementById('populateLocationSpace')?.addEventListener('click', () => {
    if (editorState.selectedElements.length === 0) {
        showToast('No elements selected', 'error');
        return;
    }
    
    const elementIds = editorState.selectedElements.map(e => e.elementId);
    copyComputedToParam(elementIds, 'RoomSpace', 'FM_LocationSpace');
});

// Initial render
render();
