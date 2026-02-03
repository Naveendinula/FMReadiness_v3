// FM Readiness - Web UI Application

// Global state
window.auditData = {
    summary: null,
    rows: []
};

// Preset state
window.presetState = {
    currentPreset: null,
    availablePresets: [],
    presetFields: null
};

const COBIE_PRESET_FILE = 'cobie-core.json';
const LEGACY_PRESET_FILE = 'fm-legacy.json';

// COBie validation state
window.cobieState = {
    validationResults: null,
    uniquenessViolations: [],
    parametersEnsured: false
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
const autoSyncSelectionEl = document.getElementById('autoSyncSelection');
const lockSelectionEl = document.getElementById('lockSelection');
const selectionScopeEl = document.getElementById('selectionScope');
const scopeCategoryEl = document.getElementById('scopeCategory');
const applyScopeEl = document.getElementById('applyScope');

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
            return;
        }

        // Preset messages
        if (msg && msg.type === 'availablePresets') {
            receiveAvailablePresets(msg);
            return;
        }

        if (msg && msg.type === 'presetLoaded') {
            receivePresetLoaded(msg);
            return;
        }

        if (msg && msg.type === 'presetFields') {
            receivePresetFields(msg);
            return;
        }

        // COBie validation messages
        if (msg && msg.type === 'cobieReadinessResult') {
            receiveCobieReadinessResult(msg);
            return;
        }

        if (msg && msg.type === 'cobieParametersEnsured') {
            receiveCobieParametersEnsured(msg);
            return;
        }

        if (msg && msg.type === 'legacyParametersEnsured') {
            receiveLegacyParametersEnsured(msg);
            return;
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

        if (tabId === 'editor') {
            requestSelectedElements();
            setSelectionSyncState();
            updateScopeCategoryVisibility();
        }
    });
});

// ========================================
// Parameter Editor State
// ========================================
let editorState = {
    selectedElements: [],
    currentTypeId: null,
    currentTypeInfo: null,
    editorMode: 'cobie', // 'cobie' or 'legacy'
    cobieFields: {
        instance: [],
        type: []
    }
};

// ========================================
// Editor Mode Toggle
// ========================================
document.querySelectorAll('.mode-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        const mode = btn.dataset.mode;
        editorState.editorMode = mode;
        
        // Update button states
        document.querySelectorAll('.mode-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        
        // Toggle editor modes visibility
        document.querySelectorAll('.editor-mode').forEach(m => m.classList.remove('active'));
        const modeEl = document.getElementById(mode === 'cobie' ? 'cobieEditorMode' : 'legacyEditorMode');
        if (modeEl) modeEl.classList.add('active');

        switchPresetForMode(mode);
        // Re-render fields if COBie mode
        if (mode === 'cobie' && window.presetState.presetFields) {
            renderCobieEditorFields();
        }
    });
});

function switchPresetForMode(mode) {
    const desired = mode === 'legacy' ? LEGACY_PRESET_FILE : COBIE_PRESET_FILE;
    const dropdown = document.getElementById('presetDropdown');
    if (dropdown && dropdown.value !== desired) {
        dropdown.value = desired;
        loadPreset(desired);
    }
}

// ========================================
// COBie Editor Field Rendering
// ========================================
function renderCobieEditorFields() {
    const presetFields = window.presetState.presetFields;
    if (!presetFields) return;
    
    const instanceGrid = document.getElementById('cobieInstanceFieldsGrid');
    const typeGrid = document.getElementById('cobieTypeFieldsGrid');
    const bulkGrid = document.getElementById('cobieBulkFieldsGrid');
    
    if (!instanceGrid || !typeGrid) return;
    
    // Clear existing fields
    instanceGrid.innerHTML = '';
    typeGrid.innerHTML = '';
    if (bulkGrid) bulkGrid.innerHTML = '';
    
    editorState.cobieFields.instance = [];
    editorState.cobieFields.type = [];
    
    // Render component (instance) fields
    const componentGroups = Array.isArray(presetFields.componentGroups)
        ? presetFields.componentGroups
        : Object.entries(presetFields.componentGroups || {}).map(([name, fields]) => ({ name, fields }));
    let instanceCount = 0;
    
    componentGroups.forEach(group => {
        const fields = group.fields || [];
        fields.forEach(field => {
            const fieldHtml = createCobieFieldInput(field, 'instance');
            instanceGrid.insertAdjacentHTML('beforeend', fieldHtml);
            editorState.cobieFields.instance.push(field);
            instanceCount++;
        });
    });
    
    // Update instance field count
    const instanceCountEl = document.getElementById('cobieInstanceFieldCount');
    if (instanceCountEl) {
        instanceCountEl.textContent = `(${instanceCount} fields)`;
    }
    
    // Render type fields
    const typeGroups = Array.isArray(presetFields.typeGroups)
        ? presetFields.typeGroups
        : Object.entries(presetFields.typeGroups || {}).map(([name, fields]) => ({ name, fields }));
    let typeCount = 0;
    
    typeGroups.forEach(group => {
        const fields = group.fields || [];
        fields.forEach(field => {
            const fieldHtml = createCobieFieldInput(field, 'type');
            typeGrid.insertAdjacentHTML('beforeend', fieldHtml);
            editorState.cobieFields.type.push(field);
            typeCount++;
        });
    });
    
    // Update type field count
    const typeCountEl = document.getElementById('cobieTypeFieldCount');
    if (typeCountEl) {
        typeCountEl.textContent = `(${typeCount} fields)`;
    }
    
    // Render bulk fill fields (common fields only)
    if (bulkGrid) {
        const commonFields = editorState.cobieFields.instance.filter(f => 
            ['InstallationDate', 'WarrantyStartDate', 'Space'].some(k => f.cobieKey?.includes(k))
        );
        commonFields.forEach(field => {
            const fieldHtml = createCobieFieldInput(field, 'bulk');
            bulkGrid.insertAdjacentHTML('beforeend', fieldHtml);
        });
    }

    applyCobieEditorEnabledState();
}

function createCobieFieldInput(field, prefix) {
    const isRequired = field.required === true;
    const cobieKey = field.cobieKey || field.key || '';
    const fieldId = `${prefix}_${cobieKey.replace(/\./g, '_')}`;
    const label = field.label || cobieKey;
    const dataType = field.dataType || 'string';
    const aliases = field.aliasParams || [];
    
    let inputHtml = '';
    let placeholder = field.revitParam || cobieKey;
    
    if (dataType === 'date') {
        placeholder = 'YYYY-MM-DD';
        inputHtml = `<input type="text" id="${fieldId}" placeholder="${placeholder}" data-cobie-key="${cobieKey}" data-revit-param="${field.revitParam || ''}" class="cobie-field-input">`;
    } else if (dataType === 'number') {
        inputHtml = `<input type="number" id="${fieldId}" placeholder="${placeholder}" data-cobie-key="${cobieKey}" data-revit-param="${field.revitParam || ''}" class="cobie-field-input">`;
    } else if (field.options && field.options.length > 0) {
        const optionsHtml = field.options.map(o => `<option value="${o}">${o}</option>`).join('');
        inputHtml = `<select id="${fieldId}" data-cobie-key="${cobieKey}" data-revit-param="${field.revitParam || ''}" class="cobie-field-input"><option value="">Select...</option>${optionsHtml}</select>`;
    } else {
        inputHtml = `<input type="text" id="${fieldId}" placeholder="${placeholder}" data-cobie-key="${cobieKey}" data-revit-param="${field.revitParam || ''}" class="cobie-field-input">`;
    }
    
    const aliasHint = aliases.length > 0 ? `<span class="alias-hint">Also: ${aliases.slice(0, 2).join(', ')}${aliases.length > 2 ? '...' : ''}</span>` : '';
    
    return `
        <div class="form-group ${isRequired ? 'required' : ''}">
            <label for="${fieldId}">${label}</label>
            ${inputHtml}
            ${aliasHint}
        </div>
    `;
}

// ========================================
// COBie Editor - Apply Values
// ========================================
function ensureCobieParameters() {
    if (window.chrome && window.chrome.webview) {
        const includeAliases = document.getElementById('writeAliases')?.checked ?? true;
        const presetFile = document.getElementById('presetDropdown')?.value;
        const removeFmAliases = document.getElementById('removeFmAliases')?.checked ?? false;
        const effectiveIncludeAliases = includeAliases && !removeFmAliases;
        window.chrome.webview.postMessage({
            action: 'ensureCobieParameters',
            includeAliases: effectiveIncludeAliases,
            presetFile: presetFile,
            mode: 'cobie',
            replaceCobie: false,
            removeFmAliases: removeFmAliases
        });
    }
}

document.getElementById('ensureCobieParams')?.addEventListener('click', () => {
    ensureCobieParameters();
});

function ensureLegacyParameters() {
    if (window.chrome && window.chrome.webview) {
        const includeAliases = document.getElementById('writeAliases')?.checked ?? true;
        window.chrome.webview.postMessage({
            action: 'ensureCobieParameters',
            includeAliases: includeAliases,
            presetFile: LEGACY_PRESET_FILE,
            mode: 'legacy',
            replaceCobie: true
        });
    }
}

document.getElementById('ensureLegacyParams')?.addEventListener('click', () => {
    ensureLegacyParameters();
});

function receiveCobieParametersEnsured(msg) {
    if (msg && msg.success) {
        window.cobieState.parametersEnsured = true;
        applyCobieEditorEnabledState();
    }
}

function receiveLegacyParametersEnsured(msg) {
    if (msg && msg.success) {
        window.cobieState.parametersEnsured = false;
        applyCobieEditorEnabledState();
    }
}

function requireCobieEnsured() {
    if (window.cobieState.parametersEnsured) return true;
    showToast('Click "Ensure COBie Parameters" before editing COBie fields.', 'error');
    return false;
}

function applyCobieEditorEnabledState() {
    const enabled = window.cobieState.parametersEnsured;
    document.querySelectorAll('.cobie-field-input').forEach(input => {
        input.disabled = !enabled;
    });

    const ids = [
        'applyCobieToSelected',
        'applyCobieType',
        'applyCobieToCategory',
        'copyToSerialNumber',
        'copyToSpace'
    ];
    ids.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.disabled = !enabled;
    });

    const notice = document.getElementById('cobieEnsureNotice');
    if (notice) {
        notice.classList.toggle('hidden', enabled);
    }
}

function setCobieFieldValues(elementIds, fieldValues, writeAliases) {
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage({
            action: 'setCobieFieldValues',
            elementIds: elementIds,
            fieldValues: fieldValues,
            writeAliases: writeAliases
        });
    }
}

document.getElementById('applyCobieToSelected')?.addEventListener('click', () => {
    if (!requireCobieEnsured()) return;
    if (editorState.selectedElements.length === 0) {
        showToast('No elements selected', 'error');
        return;
    }
    
    const fieldValues = collectCobieFieldValues('instance');
    if (Object.keys(fieldValues).length === 0) {
        showToast('No values to apply', 'warning');
        return;
    }
    
    const writeAliases = document.getElementById('writeAliases')?.checked ?? true;
    const elementIds = editorState.selectedElements.map(e => e.elementId);
    
    setCobieFieldValues(elementIds, fieldValues, writeAliases);
});

document.getElementById('applyCobieType')?.addEventListener('click', () => {
    if (!requireCobieEnsured()) return;
    if (!editorState.currentTypeId) {
        showToast('No type selected', 'error');
        return;
    }
    
    const fieldValues = collectCobieFieldValues('type');
    if (Object.keys(fieldValues).length === 0) {
        showToast('No values to apply', 'warning');
        return;
    }
    
    const writeAliases = document.getElementById('writeAliases')?.checked ?? true;
    
    // Use type-specific action
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage({
            action: 'setCobieTypeFieldValues',
            typeId: editorState.currentTypeId,
            fieldValues: fieldValues,
            writeAliases: writeAliases
        });
    }
});

document.getElementById('applyCobieToCategory')?.addEventListener('click', () => {
    if (!requireCobieEnsured()) return;
    const category = document.getElementById('cobieBulkCategory')?.value;
    if (!category) {
        showToast('Select a category first', 'error');
        return;
    }
    
    const fieldValues = collectCobieFieldValues('bulk');
    if (Object.keys(fieldValues).length === 0) {
        showToast('No values to apply', 'warning');
        return;
    }
    
    const onlyBlanks = document.getElementById('cobieOnlyFillBlanks')?.checked ?? true;
    const writeAliases = document.getElementById('writeAliases')?.checked ?? true;
    
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage({
            action: 'setCobieCategoryFieldValues',
            category: category,
            fieldValues: fieldValues,
            onlyBlanks: onlyBlanks,
            writeAliases: writeAliases
        });
    }
});

function collectCobieFieldValues(prefix) {
    const fieldValues = {};
    const inputs = document.querySelectorAll(`#cobie${capitalize(prefix)}FieldsGrid .cobie-field-input, #cobie${capitalize(prefix)}Form .cobie-field-input`);
    
    // Fallback for bulk fields which might be in a different container
    const bulkInputs = prefix === 'bulk' ? document.querySelectorAll('#cobieBulkFieldsGrid .cobie-field-input') : [];
    const allInputs = inputs.length > 0 ? inputs : bulkInputs;
    
    // In multi-select mode, only collect fields the user actually modified
    const isMultiSelect = editorState.selectedElements.length > 1;
    
    allInputs.forEach(input => {
        const value = input.value?.trim();
        const cobieKey = input.dataset.cobieKey;
        const revitParam = input.dataset.revitParam;
        
        if (!cobieKey && !revitParam) return;
        
        // In multi-select: only include if user modified the field, OR if it has a value (single select)
        if (isMultiSelect) {
            // Only send if user explicitly modified this field
            if (input.getAttribute('data-user-modified') === 'true' && value) {
                fieldValues[cobieKey || revitParam] = {
                    value: value,
                    cobieKey: cobieKey,
                    revitParam: revitParam
                };
            }
        } else {
            // Single select: send all non-empty values
            if (value) {
                fieldValues[cobieKey || revitParam] = {
                    value: value,
                    cobieKey: cobieKey,
                    revitParam: revitParam
                };
            }
        }
    });
    
    return fieldValues;
}

function capitalize(str) {
    return str.charAt(0).toUpperCase() + str.slice(1);
}

// ========================================
// COBie Editor - Populate from Selection
// ========================================
function populateCobieFieldsFromElement(element) {
    if (!element) return;
    
    const instanceParams = element.instanceParams || {};
    const typeParams = element.typeParams || {};
    const cobieInstanceParams = element.cobieInstanceParams || element.cobieParams || {};
    const cobieTypeParams = element.cobieTypeParams || {};
    
    // Populate instance fields
    editorState.cobieFields.instance.forEach(field => {
        const fieldId = `instance_${(field.cobieKey || '').replace(/\./g, '_')}`;
        const input = document.getElementById(fieldId);
        if (!input) return;
        
        // Try COBie param first, then revitParam, then aliases
        let value = cobieInstanceParams[field.cobieKey] || 
                    instanceParams[field.revitParam] ||
                    (field.aliasParams || []).map(a => instanceParams[a]).find(v => v);
        
        input.value = value || '';
        
        // Update visual state
        if (field.required && !value) {
            input.classList.add('missing-required');
            input.classList.remove('has-value');
        } else if (value) {
            input.classList.add('has-value');
            input.classList.remove('missing-required');
        } else {
            input.classList.remove('has-value', 'missing-required');
        }
    });
    
    // Populate type fields
    editorState.cobieFields.type.forEach(field => {
        const fieldId = `type_${(field.cobieKey || '').replace(/\./g, '_')}`;
        const input = document.getElementById(fieldId);
        if (!input) return;
        
        let value = cobieTypeParams[field.cobieKey] || typeParams[field.revitParam] || typeParams[field.cobieKey] ||
                    (field.aliasParams || []).map(a => typeParams[a]).find(v => v);
        
        input.value = value || '';
    });
    
    // Update computed values for COBie editor
    const computed = element.computed || {};
    const cobieUniqueId = document.getElementById('cobie_computed_UniqueId');
    const cobieLevel = document.getElementById('cobie_computed_Level');
    const cobieRoomSpace = document.getElementById('cobie_computed_RoomSpace');
    
    if (cobieUniqueId) cobieUniqueId.textContent = computed.UniqueId || '--';
    if (cobieLevel) cobieLevel.textContent = computed.Level || '--';
    if (cobieRoomSpace) cobieRoomSpace.textContent = computed.RoomSpace || '--';
}

// COBie category stats
document.getElementById('cobieBulkCategory')?.addEventListener('change', (e) => {
    const category = e.target.value;
    if (category) {
        requestCategoryStats(category);
    } else {
        document.getElementById('cobieCategoryCount').textContent = '';
    }
});

// Copy computed values for COBie
document.getElementById('copyToSerialNumber')?.addEventListener('click', () => {
    if (!requireCobieEnsured()) return;
    if (editorState.selectedElements.length === 0) {
        showToast('No elements selected', 'error');
        return;
    }
    
    const elementIds = editorState.selectedElements.map(e => e.elementId);
    const writeAliases = document.getElementById('writeAliases')?.checked ?? true;
    
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage({
            action: 'copyComputedToCobieParam',
            elementIds: elementIds,
            sourceField: 'UniqueId',
            targetCobieKey: 'Component.SerialNumber',
            writeAliases: writeAliases
        });
    }
});

document.getElementById('copyToSpace')?.addEventListener('click', () => {
    if (!requireCobieEnsured()) return;
    if (editorState.selectedElements.length === 0) {
        showToast('No elements selected', 'error');
        return;
    }
    
    const elementIds = editorState.selectedElements.map(e => e.elementId);
    const writeAliases = document.getElementById('writeAliases')?.checked ?? true;
    
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage({
            action: 'copyComputedToCobieParam',
            elementIds: elementIds,
            sourceField: 'RoomSpace',
            targetCobieKey: 'Component.Space',
            writeAliases: writeAliases
        });
    }
});

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

function setSelectionSyncState() {
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage({
            action: 'setSelectionSyncState',
            autoSync: autoSyncSelectionEl?.checked ?? true,
            lockSelection: lockSelectionEl?.checked ?? false
        });
    }
}

function applySelectionScope() {
    if (!window.chrome || !window.chrome.webview) return;

    const scope = selectionScopeEl?.value || 'selection';
    const payload = { action: 'applySelectionScope', scope: scope };

    if (scope === 'category') {
        const category = scopeCategoryEl?.value;
        if (!category) {
            showToast('Select a category', 'error');
            return;
        }
        payload.category = category;
    }

    if (scope === 'selectedType') {
        const typeIds = [...new Set(editorState.selectedElements.map(e => e.typeId).filter(Boolean))];
        if (typeIds.length !== 1) {
            showToast('Select elements of a single type', 'error');
            return;
        }
        payload.typeId = typeIds[0];
    }

    window.chrome.webview.postMessage(payload);
}

function updateScopeCategoryVisibility() {
    if (!scopeCategoryEl || !selectionScopeEl) return;
    if (selectionScopeEl.value === 'category') {
        scopeCategoryEl.classList.remove('hidden');
    } else {
        scopeCategoryEl.classList.add('hidden');
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
    
    // COBie editor elements
    const cobieInstanceForm = document.getElementById('cobieInstanceForm');
    const cobieTypeInfo = document.getElementById('cobieTypeInfo');
    const cobieTypeForm = document.getElementById('cobieTypeForm');
    
    if (elements.length === 0) {
        infoEl.innerHTML = '<p class="placeholder-text">No elements selected. Select elements in Revit or use Scope.</p>';
        formEl?.classList.add('hidden');
        typeInfoEl.innerHTML = '<p class="placeholder-text">Select an element to edit its type parameters</p>';
        typeFormEl?.classList.add('hidden');
        
        // COBie editor
        cobieInstanceForm?.classList.add('hidden');
        if (cobieTypeInfo) cobieTypeInfo.innerHTML = '<p class="placeholder-text">Select an element to edit type parameters</p>';
        cobieTypeForm?.classList.add('hidden');
        
        // Hide multi-select legend
        document.getElementById('multiSelectLegend')?.classList.add('hidden');
        
        clearComputedValues();
        return;
    }
    
    // Show element count badge for multi-select
    const isMultiSelect = elements.length > 1;
    let badgesHtml = '';
    
    if (isMultiSelect) {
        // Show count badge and category summary
        const categories = [...new Set(elements.map(e => e.category))];
        const categoryText = categories.length === 1 
            ? categories[0] 
            : `${categories.length} categories`;
        badgesHtml = `
            <span class="element-count-badge">
                <span class="icon">☰</span>
                <span>${elements.length} elements selected (${categoryText})</span>
            </span>
        `;
        // Show legend
        document.getElementById('multiSelectLegend')?.classList.remove('hidden');
    } else {
        // Single element - show badge
        const el = elements[0];
        badgesHtml = `
            <span class="element-badge">
                <span class="element-id">${el.elementId}</span>
                <span>${el.category} - ${el.family}</span>
            </span>
        `;
        // Hide legend
        document.getElementById('multiSelectLegend')?.classList.add('hidden');
    }
    
    infoEl.innerHTML = badgesHtml;
    formEl?.classList.remove('hidden');
    cobieInstanceForm?.classList.remove('hidden');
    
    // Handle single or multiple selection
    if (elements.length === 1) {
        // Single selection - pre-fill all fields directly
        const el = elements[0];
        populateLegacyFieldsFromElement(el);
        populateCobieFieldsFromElement(el);
        
        // Update type info (legacy)
        if (el.typeId && el.typeName) {
            editorState.currentTypeId = el.typeId;
            editorState.currentTypeInfo = el;
            typeInfoEl.innerHTML = `
                <div class="type-badge">
                    <span class="type-name">${el.typeName}</span>
                    <span>(${el.typeInstanceCount || '?'} instances)</span>
                </div>
            `;
            typeFormEl?.classList.remove('hidden');
            
            const tp = el.typeParams || {};
            document.getElementById('type_Manufacturer').value = tp.Manufacturer || '';
            document.getElementById('type_Model').value = tp.Model || '';
            document.getElementById('type_TypeMark').value = tp.TypeMark || '';
            document.getElementById('typeImpact').textContent = `Will update ${el.typeInstanceCount || '?'} instances`;
            
            // COBie type info
            if (cobieTypeInfo) {
                cobieTypeInfo.innerHTML = `
                    <div class="type-badge">
                        <span class="type-name">${el.typeName}</span>
                        <span>(${el.typeInstanceCount || '?'} instances)</span>
                    </div>
                `;
            }
            cobieTypeForm?.classList.remove('hidden');
            const cobieTypeImpact = document.getElementById('cobieTypeImpact');
            if (cobieTypeImpact) cobieTypeImpact.textContent = `Will update ${el.typeInstanceCount || '?'} instances`;
        }
        
        // Update computed values
        const computed = el.computed || {};
        document.getElementById('computed_UniqueId').textContent = computed.UniqueId || '--';
        document.getElementById('computed_Level').textContent = computed.Level || '--';
        document.getElementById('computed_RoomSpace').textContent = computed.RoomSpace || '--';
        
        // Populate COBie editor fields
        populateCobieFieldsFromElement(el);
    } else {
        // Multiple selection - compute and display common values
        populateMultiSelectFields(elements);
        
        // Check if all elements share the same type
        const typeIds = [...new Set(elements.map(e => e.typeId).filter(Boolean))];
        if (typeIds.length === 1) {
            // Same type - allow type editing
            const el = elements[0];
            editorState.currentTypeId = el.typeId;
            editorState.currentTypeInfo = el;
            const totalInstances = el.typeInstanceCount || '?';
            typeInfoEl.innerHTML = `
                <div class="type-badge">
                    <span class="type-name">${el.typeName}</span>
                    <span>(${totalInstances} instances total, ${elements.length} selected)</span>
                </div>
            `;
            typeFormEl?.classList.remove('hidden');
            
            // Populate type fields with common values
            populateTypeFieldsFromMultiSelect(elements);
            
            if (cobieTypeInfo) {
                cobieTypeInfo.innerHTML = `
                    <div class="type-badge">
                        <span class="type-name">${el.typeName}</span>
                        <span>(${totalInstances} instances total)</span>
                    </div>
                `;
            }
            cobieTypeForm?.classList.remove('hidden');
        } else {
            // Different types - disable type editing
            typeInfoEl.innerHTML = `<p class="placeholder-text multi-type-warning">⚠ ${typeIds.length} different types selected - type editing disabled</p>`;
            typeFormEl?.classList.add('hidden');
            
            if (cobieTypeInfo) cobieTypeInfo.innerHTML = `<p class="placeholder-text multi-type-warning">⚠ ${typeIds.length} different types selected</p>`;
            cobieTypeForm?.classList.add('hidden');
        }
        
        // Clear computed values for multi-select (they vary per element)
        clearComputedValues();
    }
}

// ========================================
// Multi-Select Value Comparison
// ========================================

/**
 * Compute common values across multiple selected elements.
 * Returns an object with { fieldName: { value, varies, blankCount } }
 */
function computeCommonValues(elements, fieldGetter) {
    const comparison = {};
    const sampleFields = fieldGetter(elements[0]) || {};
    
    Object.keys(sampleFields).forEach(fieldName => {
        const values = elements.map(el => {
            const fields = fieldGetter(el) || {};
            return fields[fieldName] || '';
        });
        
        const nonBlankValues = values.filter(v => v && v.trim());
        const uniqueValues = [...new Set(nonBlankValues)];
        
        comparison[fieldName] = {
            value: uniqueValues.length === 1 ? uniqueValues[0] : null,
            varies: uniqueValues.length > 1,
            blankCount: values.length - nonBlankValues.length,
            uniqueValues: uniqueValues
        };
    });
    
    return comparison;
}

/**
 * Populate legacy FM_ fields from multi-select with common value logic
 */
function populateMultiSelectFields(elements) {
    // Legacy FM_ instance params
    const legacyComparison = computeCommonValues(elements, el => el.instanceParams);
    const legacyFields = [
        'FM_Barcode', 'FM_UniqueAssetId', 'FM_InstallationDate',
        'FM_WarrantyStart', 'FM_WarrantyEnd', 'FM_Criticality',
        'FM_Trade', 'FM_PMTemplateId', 'FM_PMFrequencyDays',
        'FM_Building', 'FM_LocationSpace'
    ];
    
    legacyFields.forEach(fieldName => {
        const input = document.getElementById(`edit_${fieldName}`);
        if (!input) return;
        
        const info = legacyComparison[fieldName] || { value: null, varies: false, blankCount: elements.length };
        applyMultiSelectValueToInput(input, info, elements.length);
    });
    
    // COBie instance params
    const cobieComparison = computeCommonValues(elements, el => el.cobieInstanceParams);
    populateCobieFieldsFromMultiSelect(cobieComparison, elements.length);
}

/**
 * Apply multi-select value info to an input field
 */
function applyMultiSelectValueToInput(input, info, totalCount) {
    // Reset state
    input.classList.remove('common-value', 'varies-value', 'all-blank');
    input.removeAttribute('data-original-value');
    input.removeAttribute('data-user-modified');
    
    if (info.value) {
        // All elements have the same value
        input.value = info.value;
        input.placeholder = '';
        input.classList.add('common-value');
        input.setAttribute('data-original-value', info.value);
    } else if (info.varies) {
        // Values differ across elements
        input.value = '';
        input.placeholder = `(varies: ${info.uniqueValues.length} values)`;
        input.classList.add('varies-value');
        input.setAttribute('data-original-value', '__VARIES__');
    } else {
        // All blank
        input.value = '';
        input.placeholder = `(all ${totalCount} empty)`;
        input.classList.add('all-blank');
        input.setAttribute('data-original-value', '');
    }
    
    // Track user modifications
    input.addEventListener('input', markFieldAsModified, { once: true });
}

function markFieldAsModified(e) {
    e.target.setAttribute('data-user-modified', 'true');
    e.target.classList.add('user-modified');
}

/**
 * Populate COBie fields from multi-select comparison
 */
function populateCobieFieldsFromMultiSelect(cobieComparison, totalCount) {
    editorState.cobieFields.instance.forEach(field => {
        const fieldId = `instance_${(field.cobieKey || '').replace(/\./g, '_')}`;
        const input = document.getElementById(fieldId);
        if (!input) return;
        
        const info = cobieComparison[field.cobieKey] || { value: null, varies: false, blankCount: totalCount };
        applyMultiSelectValueToInput(input, info, totalCount);
        
        // Update required field styling
        if (field.required && !info.value && !info.varies) {
            input.classList.add('missing-required');
        } else {
            input.classList.remove('missing-required');
        }
    });
}

/**
 * Populate type fields from multi-select (when all elements share same type)
 */
function populateTypeFieldsFromMultiSelect(elements) {
    const typeComparison = computeCommonValues(elements, el => el.typeParams);
    
    ['Manufacturer', 'Model', 'TypeMark'].forEach(fieldName => {
        const input = document.getElementById(`type_${fieldName}`);
        if (!input) return;
        
        const info = typeComparison[fieldName] || { value: null, varies: false };
        // Type params should be same across all instances of same type
        input.value = info.value || '';
    });
    
    const impactEl = document.getElementById('typeImpact');
    if (impactEl && elements[0]) {
        impactEl.textContent = `Will update ${elements[0].typeInstanceCount || '?'} instances`;
    }
    
    // COBie type fields
    const cobieTypeComparison = computeCommonValues(elements, el => el.cobieTypeParams);
    editorState.cobieFields.type.forEach(field => {
        const fieldId = `type_${(field.cobieKey || '').replace(/\./g, '_')}`;
        const input = document.getElementById(fieldId);
        if (!input) return;
        
        const info = cobieTypeComparison[field.cobieKey] || { value: null, varies: false };
        input.value = info.value || '';
    });
    
    const cobieImpactEl = document.getElementById('cobieTypeImpact');
    if (cobieImpactEl && elements[0]) {
        cobieImpactEl.textContent = `Will update ${elements[0].typeInstanceCount || '?'} instances`;
    }
}

/**
 * Helper to populate legacy fields from a single element
 */
function populateLegacyFieldsFromElement(el) {
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
    
    // Update computed values
    const computed = el.computed || {};
    document.getElementById('computed_UniqueId').textContent = computed.UniqueId || '--';
    document.getElementById('computed_Level').textContent = computed.Level || '--';
    document.getElementById('computed_RoomSpace').textContent = computed.RoomSpace || '--';
}

function clearCobieInstanceForm() {
    document.querySelectorAll('.cobie-field-input').forEach(input => {
        input.value = '';
        input.classList.remove('has-value', 'missing-required');
    });
    
    // Clear COBie computed values
    const cobieUniqueId = document.getElementById('cobie_computed_UniqueId');
    const cobieLevel = document.getElementById('cobie_computed_Level');
    const cobieRoomSpace = document.getElementById('cobie_computed_RoomSpace');
    if (cobieUniqueId) cobieUniqueId.textContent = '--';
    if (cobieLevel) cobieLevel.textContent = '--';
    if (cobieRoomSpace) cobieRoomSpace.textContent = '--';
}

function receiveCategoryStats(msg) {
    // Legacy category count
    const countEl = document.getElementById('categoryCount');
    if (msg.count !== undefined) {
        countEl.textContent = `${msg.count} elements`;
    } else {
        countEl.textContent = '';
    }
    
    // COBie category count
    const cobieCountEl = document.getElementById('cobieCategoryCount');
    if (cobieCountEl && msg.count !== undefined) {
        cobieCountEl.textContent = `${msg.count} elements`;
    } else if (cobieCountEl) {
        cobieCountEl.textContent = '';
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

autoSyncSelectionEl?.addEventListener('change', () => {
    setSelectionSyncState();
});

lockSelectionEl?.addEventListener('change', () => {
    setSelectionSyncState();
});

selectionScopeEl?.addEventListener('change', () => {
    updateScopeCategoryVisibility();
});

applyScopeEl?.addEventListener('click', () => {
    applySelectionScope();
});

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

// ========================================
// Preset Management
// ========================================
function requestAvailablePresets() {
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage({
            action: 'getAvailablePresets'
        });
    }
}

function loadPreset(fileName) {
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage({
            action: 'loadPreset',
            fileName: fileName
        });
    }
}

function requestPresetFields() {
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage({
            action: 'getPresetFields'
        });
    }
}

function receiveAvailablePresets(msg) {
    const presets = msg.presets || [];
    const currentPreset = msg.currentPreset || '';
    
    window.presetState.availablePresets = presets;
    
    const dropdown = document.getElementById('presetDropdown');
    if (!dropdown) return;
    
    dropdown.innerHTML = '';
    presets.forEach(p => {
        const opt = document.createElement('option');
        opt.value = p.fileName;
        opt.textContent = p.name;
        opt.title = p.description;
        if (p.fileName === currentPreset) {
            opt.selected = true;
        }
        dropdown.appendChild(opt);
    });
}

function receivePresetLoaded(msg) {
    if (msg.success) {
        window.presetState.currentPreset = msg.preset;
        window.cobieState.parametersEnsured = false;
        applyCobieEditorEnabledState();
        showToast(`Loaded preset: ${msg.preset.name}`, 'success');
        requestPresetFields();
    } else {
        showToast('Failed to load preset', 'error');
    }
}

function receivePresetFields(msg) {
    window.presetState.presetFields = {
        presetName: msg.presetName,
        componentGroups: msg.componentGroups,
        typeGroups: msg.typeGroups
    };
    
    // Render both COBie field reference and editor fields
    renderCobieFieldReference();
    renderCobieEditorFields();
}

// Preset dropdown change
document.getElementById('presetDropdown')?.addEventListener('change', (e) => {
    const fileName = e.target.value;
    if (fileName) {
        loadPreset(fileName);
    }
});

// Refresh presets button
document.getElementById('refreshPresets')?.addEventListener('click', () => {
    requestAvailablePresets();
});

// ========================================
// COBie Validation
// ========================================
function validateCobieReadiness() {
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage({
            action: 'validateCobieReadiness'
        });
    }
}

function receiveCobieReadinessResult(msg) {
    window.cobieState.validationResults = msg.elements || [];
    window.cobieState.uniquenessViolations = msg.uniquenessViolations || [];
    
    renderCobieSummary(msg.summary);
    renderCobieValidationResults(msg.elements);
    renderUniquenessViolations(msg.uniquenessViolations);
    renderMissingRequiredFields(msg.elements);
}

function renderCobieSummary(summary) {
    if (!summary) return;
    
    const overallScoreEl = document.getElementById('cobieOverallScore');
    const readyCountEl = document.getElementById('cobieReadyCount');
    const requiredScoreEl = document.getElementById('cobieRequiredScore');
    
    if (overallScoreEl) {
        overallScoreEl.textContent = `${Math.round(summary.averageScore)}%`;
        overallScoreEl.className = 'cobie-card-value ' + getReadinessClass(summary.averageScore);
    }
    
    if (readyCountEl) {
        readyCountEl.textContent = `${summary.cobieReadyCount}/${summary.totalElements}`;
    }
    
    if (requiredScoreEl) {
        const avgRequired = window.cobieState.validationResults.length > 0
            ? window.cobieState.validationResults.reduce((acc, e) => acc + e.requiredScore, 0) / window.cobieState.validationResults.length
            : 0;
        requiredScoreEl.textContent = `${Math.round(avgRequired)}%`;
        requiredScoreEl.className = 'cobie-card-value ' + getReadinessClass(avgRequired);
    }
}

function renderCobieValidationResults(elements) {
    const container = document.getElementById('cobieValidationResults');
    if (!container) return;
    
    if (!elements || elements.length === 0) {
        container.innerHTML = '<p class="placeholder-text">No validation results. Select elements and click "Validate Selected".</p>';
        return;
    }
    
    const html = elements.map(el => {
        const statusClass = el.cobieReady ? 'status-ready' : 'status-incomplete';
        const statusText = el.cobieReady ? '✓ COBie Ready' : '⚠ Incomplete';
        const errorCount = el.validationErrors ? el.validationErrors.length : 0;
        
        return `
            <div class="validation-item ${statusClass}">
                <div class="validation-header">
                    <span class="element-id">${el.elementId}</span>
                    <span class="element-category">${escapeHtml(el.category)}</span>
                    <span class="validation-status">${statusText}</span>
                    <span class="validation-score">${el.overallScore}%</span>
                </div>
                <div class="validation-details">
                    <span>Required: ${el.requiredFields.populated}/${el.requiredFields.total}</span>
                    <span>Optional: ${el.optionalFields.populated}/${el.optionalFields.total}</span>
                    ${errorCount > 0 ? `<span class="error-count">${errorCount} errors</span>` : ''}
                </div>
                ${errorCount > 0 ? `
                    <div class="validation-errors">
                        ${el.validationErrors.map(err => `<span class="validation-error">${escapeHtml(err.message)}</span>`).join('')}
                    </div>
                ` : ''}
            </div>
        `;
    }).join('');
    
    container.innerHTML = html;
}

function renderUniquenessViolations(violations) {
    const container = document.getElementById('uniquenessViolations');
    if (!container) return;
    
    if (!violations || violations.length === 0) {
        container.innerHTML = '<p class="placeholder-text">No uniqueness violations detected.</p>';
        return;
    }
    
    const html = violations.map(v => `
        <div class="violation-item">
            <span class="violation-field">${escapeHtml(v.field)}</span>
            <span class="violation-elements">Duplicate elements: ${v.elementIds.join(', ')}</span>
        </div>
    `).join('');
    
    container.innerHTML = html;
}

function renderMissingRequiredFields(elements) {
    const container = document.getElementById('missingRequiredFields');
    if (!container) return;
    
    if (!elements || elements.length === 0) {
        container.innerHTML = '<p class="placeholder-text">Run validation to see missing required fields.</p>';
        return;
    }
    
    const missingByField = {};
    elements.forEach(el => {
        if (el.missingRequired) {
            el.missingRequired.forEach(field => {
                if (!missingByField[field]) {
                    missingByField[field] = [];
                }
                missingByField[field].push(el.elementId);
            });
        }
    });
    
    const fields = Object.keys(missingByField);
    if (fields.length === 0) {
        container.innerHTML = '<p class="placeholder-text">All required fields are populated! ✓</p>';
        return;
    }
    
    const html = fields.map(field => `
        <div class="missing-field-item">
            <span class="missing-field-name">${escapeHtml(field)}</span>
            <span class="missing-field-count">${missingByField[field].length} elements</span>
        </div>
    `).join('');
    
    container.innerHTML = html;
}

function renderCobieFieldReference() {
    const fields = window.presetState.presetFields;
    if (!fields) return;
    
    const componentContainer = document.querySelector('#componentFieldsRef .field-ref-list');
    const typeContainer = document.querySelector('#typeFieldsRef .field-ref-list');
    
    if (componentContainer && fields.componentGroups) {
        componentContainer.innerHTML = fields.componentGroups.map(group => `
            <div class="field-group">
                <h5>${escapeHtml(group.name)}</h5>
                ${group.fields.map(f => `
                    <div class="field-ref-item ${f.required ? 'required' : ''}">
                        <span class="field-label">${escapeHtml(f.label)}</span>
                        <span class="field-param">${escapeHtml(f.revitParam || f.computedSource || '-')}</span>
                        ${f.required ? '<span class="required-badge">Required</span>' : ''}
                    </div>
                `).join('')}
            </div>
        `).join('');
    }
    
    if (typeContainer && fields.typeGroups) {
        const typeFields = fields.typeGroups.flatMap(g => g.fields);
        typeContainer.innerHTML = typeFields.map(f => `
            <div class="field-ref-item ${f.required ? 'required' : ''}">
                <span class="field-label">${escapeHtml(f.label)}</span>
                <span class="field-param">${escapeHtml(f.revitParam || '-')}</span>
                ${f.required ? '<span class="required-badge">Required</span>' : ''}
            </div>
        `).join('');
    }
}

// Validate COBie button
document.getElementById('validateCobieBtn')?.addEventListener('click', () => {
    validateCobieReadiness();
});

// Selected Elements help panel
document.getElementById('selectionHelpHeader')?.addEventListener('click', () => {
    const content = document.getElementById('selectionHelpContent');
    const icon = document.querySelector('#selectionHelpHeader .collapse-icon');
    if (content && icon) {
        content.classList.toggle('collapsed');
        icon.textContent = content.classList.contains('collapsed') ? '▼' : '▲';
    }
});

// Collapsible field reference
document.getElementById('cobieFieldRefHeader')?.addEventListener('click', () => {
    const content = document.getElementById('cobieFieldReference');
    const icon = document.querySelector('#cobieFieldRefHeader .collapse-icon');
    if (content && icon) {
        content.classList.toggle('collapsed');
        icon.textContent = content.classList.contains('collapsed') ? '▼' : '▲';
    }
});

// Initialize presets on load
requestAvailablePresets();
requestPresetFields();

// Initial render
render();
