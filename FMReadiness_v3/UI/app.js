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
const auditProfileLabelEl = document.getElementById('auditProfileLabel');
const componentGroupSectionEl = document.getElementById('componentGroupSection');
const typeGroupSectionEl = document.getElementById('typeGroupSection');
const componentGroupScoresEl = document.getElementById('componentGroupScores');
const typeGroupScoresEl = document.getElementById('typeGroupScores');
const componentGroupMetaEl = document.getElementById('componentGroupMeta');
const typeGroupMetaEl = document.getElementById('typeGroupMeta');
const groupScoreHelpEl = document.getElementById('groupScoreHelp');
const categoryScoresEl = document.getElementById('categoryScores');
const searchInput = document.getElementById('searchInput');
const showMissingOnly = document.getElementById('showMissingOnly');
const auditScoreModeEl = document.getElementById('auditScoreMode');
const auditViewModeEl = document.getElementById('auditViewMode');
const categoryFilterEl = document.getElementById('categoryFilter');
const groupsHeaderEl = document.getElementById('groupsHeader');
const missingHeaderEl = document.getElementById('missingHeader');
const tableBody = document.getElementById('tableBody');
const resultsTable = document.getElementById('resultsTable');
const emptyState = document.getElementById('emptyState');
const autoSyncSelectionEl = document.getElementById('autoSyncSelection');
const lockSelectionEl = document.getElementById('lockSelection');
const selectionScopeEl = document.getElementById('selectionScope');
const scopeCategoryEl = document.getElementById('scopeCategory');
const applyScopeEl = document.getElementById('applyScope');
const fixBannerEl = document.getElementById('fixBanner');
const cobieSelectedPreviewEl = document.getElementById('cobieSelectedPreview');
const cobieTypePreviewEl = document.getElementById('cobieTypePreview');
const cobieCategoryPreviewEl = document.getElementById('cobieCategoryPreview');
const legacySelectedPreviewEl = document.getElementById('legacySelectedPreview');
const legacyCategoryPreviewEl = document.getElementById('legacyCategoryPreview');
const legacyTypePreviewEl = document.getElementById('legacyTypePreview');
const selectionTrayEl = document.getElementById('selectionTray');
const selectionTrayToggleEl = document.getElementById('selectionTrayToggle');
const selectionTrayCountEl = document.getElementById('selectionTrayCount');
const selectionTrayContentEl = document.getElementById('selectionTrayContent');
const selectionTrayListEl = document.getElementById('selectionTrayList');
const selectionTraySelectAllEl = document.getElementById('selectionTraySelectAll');
const removeCheckedSelectionEl = document.getElementById('removeCheckedSelection');
const keepCheckedSelectionEl = document.getElementById('keepCheckedSelection');
const clearSelectionListEl = document.getElementById('clearSelectionList');
let skipNextEditorSelectionRefresh = false;
let selectionTrayExpanded = false;
const selectionTrayCheckedIds = new Set();

// Cache of elementId -> views[]
const viewsByElementId = new Map();
const legacyDateFields = new Set(['FM_InstallationDate', 'FM_WarrantyStart', 'FM_WarrantyEnd']);
const legacyFixKeyMap = {
    edit_FM_Barcode: ['FM_Barcode', 'Component.TagNumber', 'Component.BarCode', 'COBie.Component.TagNumber', 'COBie.Component.BarCode'],
    edit_FM_UniqueAssetId: ['FM_UniqueAssetId', 'Component.AssetIdentifier', 'COBie.Component.AssetIdentifier'],
    edit_FM_InstallationDate: ['FM_InstallationDate', 'Component.InstallationDate', 'COBie.Component.InstallationDate'],
    edit_FM_WarrantyStart: ['FM_WarrantyStart', 'Component.WarrantyStartDate', 'COBie.Component.WarrantyStartDate'],
    edit_FM_WarrantyEnd: ['FM_WarrantyEnd', 'Component.WarrantyEndDate', 'COBie.Component.WarrantyEndDate'],
    edit_FM_Criticality: ['FM_Criticality', 'FMOps.Criticality', 'COBie.FMOps.Criticality'],
    edit_FM_Trade: ['FM_Trade', 'FMOps.Trade', 'COBie.FMOps.Trade'],
    edit_FM_PMTemplateId: ['FM_PMTemplateId', 'FMOps.PMTemplateId', 'COBie.FMOps.PMTemplateId'],
    edit_FM_PMFrequencyDays: ['FM_PMFrequencyDays', 'FMOps.PMFrequencyDays', 'COBie.FMOps.PMFrequencyDays'],
    edit_FM_Building: ['FM_Building', 'FMOps.Building', 'COBie.FMOps.Building'],
    edit_FM_LocationSpace: ['FM_LocationSpace', 'Component.Space', 'COBie.Component.Space'],
    type_Manufacturer: ['Manufacturer', 'Type.Manufacturer'],
    type_Model: ['Model', 'Type.ModelNumber'],
    type_TypeMark: ['TypeMark', 'Type.TypeMark']
};

const groupExplanationMap = {
    identity: 'Identity fields identify each asset or type for tracking and handover.',
    metadata: 'Metadata fields capture authoring and timestamp context.',
    installwarranty: 'Install and warranty fields support warranty and commissioning workflows.',
    location: 'Location fields tie assets to rooms/spaces/levels.',
    fmops: 'FM operations fields support maintenance planning and operations.',
    classification: 'Classification fields map assets to agreed classification systems.',
    makemodel: 'Make/model fields capture manufacturer and model references.',
    warranty: 'Type warranty fields define coverage terms and units.',
    lifecycle: 'Lifecycle fields support replacement planning and service life.',
    dimensions: 'Dimension fields capture physical size information.',
    reference: 'Reference fields link to external documents or model references.'
};

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
    const { summary, rows: rawRows } = window.auditData;
    const baseSummary = summary || {
        auditProfile: '',
        scoreModeLabel: '',
        scoreMode: '',
        groupScores: {},
        groupStats: {},
        groupDefinitions: {}
    };
    const viewMode = auditViewModeEl?.value || 'combined';
    const rows = projectAuditRowsByView(rawRows || [], viewMode);
    updateAuditTableHeaders(viewMode);

    updateCategoryFilterOptions(rows);
    const scopeCategory = categoryFilterEl?.value || 'all';
    const scopedRows = scopeCategory === 'all'
        ? rows
        : rows.filter(row => row.category === scopeCategory);
    const summaryToRender = rows.length > 0 ? buildScopedSummary(scopedRows, baseSummary) : null;

    // Update summary cards
    if (summaryToRender) {
        const pct = summaryToRender.overallReadinessPercent || 0;
        overallReadinessEl.textContent = `${Math.round(pct)}%`;
        overallReadinessEl.className = 'card-value ' + getReadinessClass(pct);

        const fullyReady = summaryToRender.fullyReady || 0;
        const totalAudited = summaryToRender.totalAudited || 0;
        const missing = totalAudited - fullyReady;

        fullyReadyCountEl.textContent = fullyReady;
        fullyReadySublabelEl.textContent = `${fullyReady} assets complete`;

        missingCountEl.textContent = missing;
        missingSublabelEl.textContent = `${missing} assets incomplete`;

        totalCountEl.textContent = totalAudited;
        if (auditProfileLabelEl) {
            const modeLabel = summary.scoreModeLabel ? ` • ${summary.scoreModeLabel}` : '';
            auditProfileLabelEl.textContent = summaryToRender.auditProfile
                ? `Audit profile: ${summaryToRender.auditProfile}${modeLabel}${scopeCategory !== 'all' ? ` • Category: ${scopeCategory}` : ''}`
                : 'Audit profile: --';
        }

        const viewLabel = viewMode === 'combined'
            ? 'Combined'
            : viewMode === 'component'
                ? 'Component only'
                : 'Type only';
        if (auditProfileLabelEl && summaryToRender.auditProfile) {
            const modeLabel = summary.scoreModeLabel ? ` • ${summary.scoreModeLabel}` : '';
            auditProfileLabelEl.textContent = `Audit profile: ${summaryToRender.auditProfile}${modeLabel} • View: ${viewLabel}${scopeCategory !== 'all' ? ` • Category: ${scopeCategory}` : ''}`;
        }

        if (auditProfileLabelEl && summaryToRender.auditProfile) {
            const viewLabel = viewMode === 'combined'
                ? 'Combined'
                : viewMode === 'component'
                    ? 'Component only'
                    : 'Type only';
            const modeLabel = summaryToRender.scoreModeLabel ? ` | ${summaryToRender.scoreModeLabel}` : '';
            const categoryLabel = scopeCategory !== 'all' ? ` | Category: ${scopeCategory}` : '';
            auditProfileLabelEl.textContent = `Audit profile: ${summaryToRender.auditProfile}${modeLabel} | View: ${viewLabel}${categoryLabel}`;
        }

        // Render group score badges
        renderGroupScores(
            summaryToRender.groupScores || {},
            summaryToRender.groupStats || {},
            summaryToRender.groupDefinitions || {},
            summaryToRender.scoreModeLabel || ''
        );
        renderCategoryScores(rows);

        if (auditScoreModeEl && summaryToRender.scoreMode) {
            auditScoreModeEl.value = summaryToRender.scoreMode;
        }
    } else {
        overallReadinessEl.textContent = '--%';
        overallReadinessEl.className = 'card-value';
        fullyReadyCountEl.textContent = '--';
        fullyReadySublabelEl.textContent = '-- assets complete';
        missingCountEl.textContent = '--';
        missingSublabelEl.textContent = '-- assets incomplete';
        totalCountEl.textContent = '--';
        if (auditProfileLabelEl) {
            auditProfileLabelEl.textContent = 'Audit profile: --';
        }
        renderGroupScoreSections({}, {}, {}, viewMode, '');
        if (categoryScoresEl) {
            categoryScoresEl.innerHTML = '';
        }
    }

    // Filter rows
    const filteredRows = filterRows(scopedRows);

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
                row.missingParamsDisplay || row.missingParams,
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
function renderGroupScores(groupScores, groupStats, groupDefinitions, scoreModeLabel) {
    const viewMode = auditViewModeEl?.value || 'combined';
    renderGroupScoreSections(groupScores, groupStats, groupDefinitions, viewMode, scoreModeLabel);
}

function isTypeGroupName(name) {
    return String(name || '').trim().toLowerCase().startsWith('type:');
}

function splitGroupMapByScope(groupMap) {
    const component = {};
    const type = {};
    Object.entries(groupMap || {}).forEach(([name, value]) => {
        if (isTypeGroupName(name)) {
            type[name] = value;
        } else {
            component[name] = value;
        }
    });
    return { component, type };
}

function splitGroupScoresByScope(groupScores) {
    return splitGroupMapByScope(groupScores);
}

function normalizeGroupStatsMap(groupStats) {
    const normalized = {};
    Object.entries(groupStats || {}).forEach(([groupName, stats]) => {
        const totalChecks = Number(stats?.totalChecks ?? stats?.total ?? 0);
        const failedChecks = Number(stats?.failedChecks ?? stats?.failed ?? 0);
        if (!Number.isFinite(totalChecks) || totalChecks <= 0) return;

        const safeTotal = Math.max(0, Math.round(totalChecks));
        const safeFailed = Math.max(0, Math.min(safeTotal, Math.round(Number.isFinite(failedChecks) ? failedChecks : 0)));
        normalized[groupName] = {
            totalChecks: safeTotal,
            failedChecks: safeFailed,
            passedChecks: Math.max(0, safeTotal - safeFailed)
        };
    });
    return normalized;
}

function sumGroupStats(groupStats) {
    return Object.values(groupStats || {}).reduce((acc, stats) => {
        const totalChecks = Number(stats?.totalChecks) || 0;
        const failedChecks = Number(stats?.failedChecks) || 0;
        acc.totalChecks += totalChecks;
        acc.failedChecks += failedChecks;
        return acc;
    }, { totalChecks: 0, failedChecks: 0 });
}

function getGroupDefinition(groupDefinitions, groupName) {
    if (!groupDefinitions || !groupName) return null;
    return groupDefinitions[groupName] || groupDefinitions[String(groupName).trim()] || null;
}

function getGroupScopeLabel(groupName, definition) {
    const scope = String(definition?.scope || '').trim().toLowerCase();
    if (scope === 'type') return 'Type';
    if (scope === 'component' || scope === 'instance') return 'Component';
    return isTypeGroupName(groupName) ? 'Type' : 'Component';
}

function getGroupExplanation(groupName) {
    const key = String(groupName || '').replace(/^Type:\s*/i, '').trim().toLowerCase();
    return groupExplanationMap[key] || 'Audit group from the active preset.';
}

function getGroupFieldSummary(definition) {
    const fields = Array.isArray(definition?.fields)
        ? definition.fields.filter(Boolean)
        : [];
    if (fields.length === 0) return 'Defined by active preset.';
    if (fields.length <= 5) return fields.join(', ');
    return `${fields.slice(0, 5).join(', ')} +${fields.length - 5} more`;
}

function buildGroupTooltip(groupName, percent, stats, definition, scoreModeLabel) {
    const totalChecks = Number(stats?.totalChecks) || 0;
    const failedChecks = Number(stats?.failedChecks) || 0;
    const passedChecks = Math.max(0, totalChecks - failedChecks);
    const scopeLabel = getGroupScopeLabel(groupName, definition);
    const scoringLabel = scoreModeLabel
        ? `Scoring mode: ${scoreModeLabel}.`
        : 'Scoring mode: current audit scope.';

    return [
        `${groupName}`,
        `Scope: ${scopeLabel}`,
        `Score: ${Math.round(Number(percent) || 0)}%`,
        totalChecks > 0 ? `Checks passed: ${passedChecks}/${totalChecks}` : 'Checks passed: --',
        `Fields: ${getGroupFieldSummary(definition)}`,
        scoringLabel,
        getGroupExplanation(groupName)
    ].join('\n');
}

function getMissingEntryScope(entry) {
    const scope = String(entry?.scope || '').trim().toLowerCase();
    if (scope === 'type') return 'type';
    if (scope === 'component' || scope === 'instance') return 'component';

    const group = String(entry?.group || '').trim();
    if (isTypeGroupName(group)) return 'type';
    if (group) return 'component';

    const key = String(entry?.key || '').trim().toLowerCase();
    if (key.startsWith('type.') || key.includes('.type.') || key.startsWith('cobie.type.')) return 'type';

    return 'component';
}

function summarizeMissingEntries(entries, fallbackText) {
    const labels = (entries || [])
        .map(entry => (entry.label || entry.key || '').trim())
        .filter(Boolean);
    const uniqueLabels = [...new Set(labels)];
    if (uniqueLabels.length === 0) return fallbackText || '-';

    const head = uniqueLabels.slice(0, 3).join(', ');
    const remaining = uniqueLabels.length - 3;
    return remaining > 0 ? `${head} +${remaining} more` : head;
}

function renderGroupBadgeList(containerEl, groupScores, groupStats, groupDefinitions, scoreModeLabel) {
    if (!containerEl) return;
    const groups = Object.entries(groupScores || {});
    if (groups.length === 0) {
        containerEl.innerHTML = '';
        return;
    }

    containerEl.innerHTML = groups
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([name, pct]) => {
        const scoreClass = getScoreClass(pct);
        const stats = groupStats?.[name] || null;
        const totalChecks = Number(stats?.totalChecks) || 0;
        const failedChecks = Number(stats?.failedChecks) || 0;
        const passedChecks = Math.max(0, totalChecks - failedChecks);
        const countLabel = totalChecks > 0 ? `${passedChecks}/${totalChecks}` : '--';
        const definition = getGroupDefinition(groupDefinitions, name);
        const tooltip = buildGroupTooltip(name, pct, stats, definition, scoreModeLabel);
        return `<div class="group-badge ${scoreClass}">
            <span class="group-name">${escapeHtml(name)}</span>
            <span class="group-pct">${pct}%</span>
            <span class="group-count" title="${escapeHtml(tooltip)}">${countLabel}</span>
            <span class="group-info" title="${escapeHtml(tooltip)}" aria-label="Group details">i</span>
        </div>`;
    }).join('');
}

function renderGroupScoreSections(groupScores, groupStats, groupDefinitions, viewMode, scoreModeLabel) {
    const { component, type } = splitGroupScoresByScope(groupScores);
    const { component: componentStats, type: typeStats } = splitGroupMapByScope(normalizeGroupStatsMap(groupStats));
    const { component: componentDefs, type: typeDefs } = splitGroupMapByScope(groupDefinitions || {});
    const hasComponent = Object.keys(component).length > 0;
    const hasType = Object.keys(type).length > 0;

    if (componentGroupSectionEl) {
        componentGroupSectionEl.classList.toggle('hidden', viewMode === 'type' || !hasComponent);
    }
    if (typeGroupSectionEl) {
        typeGroupSectionEl.classList.toggle('hidden', viewMode === 'component' || !hasType);
    }

    const componentTotals = sumGroupStats(componentStats);
    const typeTotals = sumGroupStats(typeStats);
    if (componentGroupMetaEl) {
        componentGroupMetaEl.textContent = hasComponent
            ? `${Object.keys(component).length} groups${componentTotals.totalChecks > 0 ? ` | ${Math.max(0, componentTotals.totalChecks - componentTotals.failedChecks)}/${componentTotals.totalChecks}` : ''}`
            : '';
    }
    if (typeGroupMetaEl) {
        typeGroupMetaEl.textContent = hasType
            ? `${Object.keys(type).length} groups${typeTotals.totalChecks > 0 ? ` | ${Math.max(0, typeTotals.totalChecks - typeTotals.failedChecks)}/${typeTotals.totalChecks}` : ''}`
            : '';
    }

    if (groupScoreHelpEl) {
        const modeLabel = scoreModeLabel ? ` Current scoring mode: ${scoreModeLabel}.` : '';
        groupScoreHelpEl.textContent = `Hover badges for scope, included fields, and scoring details.${modeLabel}`;
    }

    renderGroupBadgeList(componentGroupScoresEl, component, componentStats, componentDefs, scoreModeLabel);
    renderGroupBadgeList(typeGroupScoresEl, type, typeStats, typeDefs, scoreModeLabel);
}

function projectAuditRowsByView(rows, viewMode) {
    return (rows || [])
        .map(row => {
            const { component, type } = splitGroupScoresByScope(row.groupScores || {});
            const { component: componentStats, type: typeStats } = splitGroupMapByScope(normalizeGroupStatsMap(row.groupStats || {}));
            const scopedScores = viewMode === 'type'
                ? type
                : viewMode === 'component'
                    ? component
                    : (row.groupScores || {});
            const scopedStats = viewMode === 'type'
                ? typeStats
                : viewMode === 'component'
                    ? componentStats
                    : normalizeGroupStatsMap(row.groupStats || {});

            const allEntries = normalizeMissingFieldEntries(row.missingFields, row.missingParams);
            const scopedEntries = viewMode === 'combined'
                ? allEntries
                : allEntries.filter(entry => getMissingEntryScope(entry) === viewMode);

            const scopedTotals = sumGroupStats(scopedStats);
            const scopedScoreValues = Object.values(scopedScores)
                .map(value => Number(value))
                .filter(Number.isFinite);
            const projectedReadinessFromStats = scopedTotals.totalChecks > 0
                ? Math.round(((scopedTotals.totalChecks - scopedTotals.failedChecks) / scopedTotals.totalChecks) * 100)
                : null;
            const projectedReadiness = projectedReadinessFromStats != null
                ? projectedReadinessFromStats
                : (viewMode === 'combined'
                    ? (Number(row.readinessPercent) || 0)
                    : (scopedScoreValues.length > 0
                        ? Math.round(scopedScoreValues.reduce((sum, value) => sum + value, 0) / scopedScoreValues.length)
                        : 0));

            const hasScopedData = viewMode === 'combined'
                ? true
                : Object.keys(scopedScores).length > 0 || Object.keys(scopedStats).length > 0 || scopedEntries.length > 0;

            return {
                ...row,
                groupScores: scopedScores,
                groupStats: scopedStats,
                missingFields: scopedEntries,
                missingCount: scopedEntries.length,
                missingParamsDisplay: summarizeMissingEntries(scopedEntries, row.missingParams),
                readinessPercent: projectedReadiness,
                _hasScopedData: hasScopedData
            };
        })
        .filter(row => row._hasScopedData);
}

function computeGroupStats(rows) {
    const aggregate = {};
    rows.forEach(row => {
        const stats = normalizeGroupStatsMap(row.groupStats || {});
        Object.entries(stats).forEach(([name, value]) => {
            if (!aggregate[name]) {
                aggregate[name] = { totalChecks: 0, failedChecks: 0, passedChecks: 0 };
            }
            aggregate[name].totalChecks += value.totalChecks;
            aggregate[name].failedChecks += value.failedChecks;
            aggregate[name].passedChecks += value.passedChecks;
        });
    });
    return aggregate;
}

function computeGroupScores(rows, groupStats) {
    const aggregateStats = groupStats || {};
    if (Object.keys(aggregateStats).length > 0) {
        const scores = {};
        Object.entries(aggregateStats).forEach(([name, stats]) => {
            const totalChecks = Number(stats?.totalChecks) || 0;
            const failedChecks = Number(stats?.failedChecks) || 0;
            if (totalChecks <= 0) return;
            scores[name] = Math.round(((totalChecks - failedChecks) / totalChecks) * 100);
        });
        return scores;
    }

    const totals = {};
    const counts = {};
    rows.forEach(row => {
        const scores = row.groupScores || {};
        Object.entries(scores).forEach(([name, score]) => {
            const numericScore = Number(score);
            if (!Number.isFinite(numericScore)) return;
            totals[name] = (totals[name] || 0) + numericScore;
            counts[name] = (counts[name] || 0) + 1;
        });
    });

    const averages = {};
    Object.keys(totals).forEach(name => {
        averages[name] = Math.round(totals[name] / counts[name]);
    });
    return averages;
}

function renderCategoryScores(rows) {
    if (!categoryScoresEl) return;
    if (!rows || rows.length === 0) {
        categoryScoresEl.innerHTML = '';
        return;
    }

    const stats = {};
    rows.forEach(row => {
        const category = row.category || 'Uncategorized';
        if (!stats[category]) {
            stats[category] = { total: 0, ready: 0, scoreSum: 0 };
        }
        stats[category].total += 1;
        if (row.missingCount === 0) stats[category].ready += 1;
        stats[category].scoreSum += Number(row.readinessPercent) || 0;
    });

    const sorted = Object.keys(stats).sort((a, b) => a.localeCompare(b));
    categoryScoresEl.innerHTML = sorted.map(category => {
        const data = stats[category];
        const pct = data.total > 0 ? Math.round(data.scoreSum / data.total) : 0;
        const scoreClass = getScoreClass(pct);
        const label = `${category} (${data.total})`;
        return `<div class="group-badge ${scoreClass}">
            <span class="group-name">${escapeHtml(label)}</span>
            <span class="group-pct">${pct}%</span>
        </div>`;
    }).join('');
}

function updateCategoryFilterOptions(rows) {
    if (!categoryFilterEl) return;
    const categories = Array.from(new Set((rows || []).map(r => r.category).filter(Boolean))).sort();
    const nextValues = ['all', ...categories];
    const currentValues = Array.from(categoryFilterEl.options).map(o => o.value);
    const currentSelection = categoryFilterEl.value || 'all';

    if (currentValues.join('|') !== nextValues.join('|')) {
        categoryFilterEl.innerHTML = '';
        const allOpt = document.createElement('option');
        allOpt.value = 'all';
        allOpt.textContent = 'All categories';
        categoryFilterEl.appendChild(allOpt);

        categories.forEach(category => {
            const opt = document.createElement('option');
            opt.value = category;
            opt.textContent = category;
            categoryFilterEl.appendChild(opt);
        });
    }

    categoryFilterEl.value = nextValues.includes(currentSelection) ? currentSelection : 'all';
}

function buildScopedSummary(rows, baseSummary) {
    const total = rows.length;
    const fullyReady = rows.filter(r => r.missingCount === 0).length;
    const averageScore = total > 0
        ? rows.reduce((acc, r) => acc + (Number(r.readinessPercent) || 0), 0) / total
        : 0;
    const groupStats = computeGroupStats(rows);
    const groupScores = computeGroupScores(rows, groupStats);

    return {
        ...baseSummary,
        overallReadinessPercent: averageScore,
        fullyReady: fullyReady,
        totalAudited: total,
        groupScores: groupScores,
        groupStats: groupStats
    };
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
function renderRowGroupBadges(groupScores, groupStats) {
    if (!groupScores || Object.keys(groupScores).length === 0) {
        return '-';
    }

    const sorted = Object.entries(groupScores)
        .sort((a, b) => Number(a[1]) - Number(b[1]));
    const maxBadges = 4;
    const visible = sorted.slice(0, maxBadges);
    const remaining = sorted.length - visible.length;

    const badgeHtml = visible.map(([name, pct]) => {
        const scoreClass = getScoreClass(pct);
        const abbrev = name.startsWith('Type:')
            ? `T:${name.replace(/^Type:\s*/i, '').substring(0, 3)}`
            : name.substring(0, 3);
        const stats = groupStats?.[name] || null;
        const totalChecks = Number(stats?.totalChecks) || 0;
        const failedChecks = Number(stats?.failedChecks) || 0;
        const countHint = totalChecks > 0 ? ` | ${Math.max(0, totalChecks - failedChecks)}/${totalChecks}` : '';
        return `<span class="group-badge-sm ${scoreClass}" title="${escapeHtml(name)}: ${pct}%${countHint}">${abbrev} ${pct}%</span>`;
    }).join(' ');

    return remaining > 0
        ? `${badgeHtml} <span class="group-badge-sm more-badge" title="${remaining} additional groups">+${remaining}</span>`
        : badgeHtml;
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
        const groupBadgesHtml = renderRowGroupBadges(row.groupScores, row.groupStats);
        const statusBadge = getStatusBadge(row.readinessPercent, row.missingCount);
        const missingFields = normalizeMissingFieldEntries(row.missingFields, row.missingParams);
        const canFix = row.missingCount > 0;

        tr.innerHTML = `
            <td>${row.elementId}</td>
            <td>${statusBadge}</td>
            <td>${escapeHtml(row.category)}</td>
            <td>${escapeHtml(row.family)}</td>
            <td>${escapeHtml(row.type)}</td>
            <td class="${readinessClass}">${row.readinessPercent}%</td>
            <td class="group-badges-inline">${groupBadgesHtml}</td>
            <td class="missing-params" title="${escapeHtml(row.missingParamsDisplay || row.missingParams)}">${escapeHtml(row.missingParamsDisplay || row.missingParams) || '-'}</td>
            <td>
                <select class="view-select" data-element-id="${row.elementId}" disabled>
                    <option>Select view</option>
                </select>
                <button class="open-view-btn" data-element-id="${row.elementId}" disabled>Open</button>
            </td>
            <td>
                <button class="fix-btn" data-id="${row.elementId}" data-missing-fields="${encodeURIComponent(JSON.stringify(missingFields))}" ${canFix ? '' : 'disabled'}>Fix</button>
                <button class="locate-btn" data-id="${row.elementId}" title="Select and zoom in Revit">Locate</button>
            </td>
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

function normalizeToken(value) {
    if (value === undefined || value === null) return '';
    return String(value).trim().toLowerCase().replace(/\s+/g, ' ');
}

function parseMissingFields(missingParams) {
    if (!missingParams) return [];
    return String(missingParams)
        .split(',')
        .map(x => x.replace(/\[[^\]]+\]/g, '').replace(/\(dup\)/gi, '').trim())
        .filter(Boolean);
}

function normalizeMissingFieldEntries(missingFields, missingParams) {
    const entries = [];

    if (Array.isArray(missingFields) && missingFields.length > 0) {
        missingFields.forEach(item => {
            if (!item) return;

            if (typeof item === 'string') {
                const label = item.trim();
                if (label) entries.push({ key: '', label, group: '', scope: '', required: false, reason: '' });
                return;
            }

            const key = String(item.key || item.fieldKey || '').trim();
            const label = String(item.label || item.fieldLabel || '').trim();
            const group = String(item.group || '').trim();
            const scope = String(item.scope || '').trim();
            const reason = String(item.reason || '').trim();
            const required = item.required === true;

            if (!key && !label) return;
            entries.push({ key, label, group, scope, required, reason });
        });
    }

    if (entries.length === 0) {
        parseMissingFields(missingParams).forEach(label => {
            entries.push({ key: '', label, group: '', scope: '', required: false, reason: '' });
        });
    }

    const deduped = new Map();
    entries.forEach(entry => {
        const dedupeKey = `${normalizeToken(entry.key)}|${normalizeToken(entry.label)}|${normalizeToken(entry.reason)}`;
        if (!deduped.has(dedupeKey)) {
            deduped.set(dedupeKey, entry);
        }
    });

    return Array.from(deduped.values());
}

function createFixTargetMaps(entries) {
    const tokenToTargetIds = new Map();
    const targetLabelById = new Map();

    entries.forEach((entry, index) => {
        const key = normalizeToken(entry.key);
        const reason = normalizeToken(entry.reason);
        const targetId = `${key || `target-${index}`}|${reason}`;
        const displayLabel = entry.label || entry.key || `Missing field ${index + 1}`;
        targetLabelById.set(targetId, displayLabel);

        [entry.key, entry.label].forEach(tokenValue => {
            const token = normalizeToken(tokenValue);
            if (!token) return;
            if (!tokenToTargetIds.has(token)) {
                tokenToTargetIds.set(token, new Set());
            }
            tokenToTargetIds.get(token).add(targetId);
        });
    });

    return { tokenToTargetIds, targetLabelById };
}

function getInputFixTokens(input, labelEl) {
    const tokens = new Set();
    const addToken = (value) => {
        const token = normalizeToken(value);
        if (token) tokens.add(token);
    };

    addToken(labelEl?.textContent || '');
    addToken(input?.dataset?.label || '');
    addToken(input?.dataset?.cobieKey || '');
    addToken(input?.dataset?.revitParam || '');

    const inputId = input?.id || '';
    addToken(inputId);

    if (inputId.includes('_')) {
        const suffix = inputId.substring(inputId.indexOf('_') + 1);
        addToken(suffix);
        addToken(suffix.replace(/_/g, '.'));
    }

    const mappedKeys = legacyFixKeyMap[inputId];
    if (Array.isArray(mappedKeys)) {
        mappedKeys.forEach(addToken);
    }

    return tokens;
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

function updateAuditTableHeaders(viewMode) {
    if (groupsHeaderEl) {
        groupsHeaderEl.textContent = viewMode === 'combined'
            ? 'Groups'
            : viewMode === 'component'
                ? 'Component Groups'
                : 'Type Groups';
    }

    if (missingHeaderEl) {
        missingHeaderEl.textContent = viewMode === 'combined'
            ? 'Missing'
            : viewMode === 'component'
                ? 'Missing (Component)'
                : 'Missing (Type)';
    }
}

function switchToEditorTab(options = {}) {
    const btn = document.querySelector('.tab-btn[data-tab="editor"]');
    if (!btn) return;
    if (options.skipSelectionRefresh) {
        skipNextEditorSelectionRefresh = true;
    }
    btn.click();
}

function queueFixSelectionSync(elementId, attempt = 0) {
    const maxAttempts = 4;
    const delayMs = 150 + (attempt * 150);

    setTimeout(() => {
        const isSynced = editorState.selectedElements.some(e => e.elementId === elementId);
        if (isSynced) return;

        requestSelectedElements();
        if (attempt + 1 < maxAttempts) {
            queueFixSelectionSync(elementId, attempt + 1);
        }
    }, delayMs);
}

function startFixFromAudit(elementId, missingFields) {
    editorState.pendingFix = {
        elementId: elementId,
        missingFields: normalizeMissingFieldEntries(missingFields, null)
    };

    switchToEditorTab({ skipSelectionRefresh: true });
    requestFixElementSelection(elementId);
    queueFixSelectionSync(elementId);
}

function getCurrentSelectionIds() {
    return editorState.selectedElements
        .map(e => Number(e.elementId))
        .filter(Number.isFinite);
}

function normalizeSelectionTrayState() {
    const currentIds = new Set(getCurrentSelectionIds());
    Array.from(selectionTrayCheckedIds).forEach(id => {
        if (!currentIds.has(id)) {
            selectionTrayCheckedIds.delete(id);
        }
    });
}

function updateSelectionTrayActionState() {
    const checkedCount = selectionTrayCheckedIds.size;
    if (removeCheckedSelectionEl) removeCheckedSelectionEl.disabled = checkedCount === 0;
    if (keepCheckedSelectionEl) keepCheckedSelectionEl.disabled = checkedCount === 0;

    const currentIds = getCurrentSelectionIds();
    if (selectionTraySelectAllEl) {
        selectionTraySelectAllEl.checked = currentIds.length > 0 && checkedCount === currentIds.length;
        selectionTraySelectAllEl.indeterminate = checkedCount > 0 && checkedCount < currentIds.length;
    }
}

function renderSelectionTray() {
    if (!selectionTrayEl || !selectionTrayListEl || !selectionTrayCountEl) return;

    const elements = editorState.selectedElements || [];
    if (elements.length <= 1) {
        selectionTrayEl.classList.add('hidden');
        selectionTrayExpanded = false;
        selectionTrayCheckedIds.clear();
        selectionTrayContentEl?.classList.add('hidden');
        if (selectionTraySelectAllEl) {
            selectionTraySelectAllEl.checked = false;
            selectionTraySelectAllEl.indeterminate = false;
        }
        updateSelectionTrayActionState();
        return;
    }

    normalizeSelectionTrayState();
    selectionTrayEl.classList.remove('hidden');
    selectionTrayCountEl.textContent = `${elements.length}`;
    if (selectionTrayContentEl) {
        selectionTrayContentEl.classList.toggle('hidden', !selectionTrayExpanded);
    }

    selectionTrayListEl.innerHTML = elements.map(el => {
        const id = Number(el.elementId);
        const checked = selectionTrayCheckedIds.has(id) ? 'checked' : '';
        const category = escapeHtml(el.category || '-');
        const typeName = escapeHtml(el.typeName || el.type || '-');
        return `
            <div class="selection-tray-row">
                <input type="checkbox" class="selection-tray-check" data-id="${id}" ${checked}>
                <span class="selection-tray-id">${id}</span>
                <span>${category}</span>
                <span>${typeName}</span>
                <button class="selection-tray-remove" data-id="${id}">Remove</button>
            </div>
        `;
    }).join('');

    updateSelectionTrayActionState();
}

function applySelectionIdsFromTray(nextIds, successMessage) {
    setSelectionElements(nextIds);
    if (successMessage) showToast(successMessage, 'success');
}

function clearFixHighlights() {
    document.querySelectorAll('.fix-target').forEach(el => el.classList.remove('fix-target'));
}

function applyFixTargets() {
    const pending = editorState.pendingFix;
    if (!pending || !editorState.selectedElements.some(e => e.elementId === pending.elementId)) {
        if (fixBannerEl) fixBannerEl.classList.add('hidden');
        clearFixHighlights();
        return;
    }

    const entries = normalizeMissingFieldEntries(pending.missingFields, null);
    pending.missingFields = entries;
    clearFixHighlights();

    if (entries.length === 0) {
        if (fixBannerEl) {
            fixBannerEl.textContent = `Fix mode: Element ${pending.elementId}. You can still edit any field, not just highlighted ones.`;
            fixBannerEl.classList.remove('hidden');
        }
        return;
    }

    const { tokenToTargetIds, targetLabelById } = createFixTargetMaps(entries);
    const matchedTargetIds = new Set();
    let highlighted = 0;
    const activeMode = document.querySelector('#tab-editor .editor-mode.active');
    const targetRoot = activeMode || document.getElementById('tab-editor');
    targetRoot?.querySelectorAll('.form-group').forEach(group => {
        const labelEl = group.querySelector('label');
        const input = group.querySelector('input, select');
        if (!input) return;

        const tokens = getInputFixTokens(input, labelEl);
        let isMatch = false;
        tokens.forEach(token => {
            const targetIds = tokenToTargetIds.get(token);
            if (!targetIds) return;
            isMatch = true;
            targetIds.forEach(id => matchedTargetIds.add(id));
        });

        if (isMatch) {
            input.classList.add('fix-target');
            highlighted++;
        }
    });

    const unmatchedLabels = Array.from(targetLabelById.entries())
        .filter(([id]) => !matchedTargetIds.has(id))
        .map(([, label]) => label);

    if (false && fixBannerEl) {
        fixBannerEl.textContent = highlighted > 0
            ? `Fix mode: Element ${pending.elementId} • Highlighted ${highlighted} audit-missing field(s). You can still edit any field.`
            : `Fix mode: Element ${pending.elementId} • Audit-missing fields: ${(pending.missingFields || []).join(', ')}. You can still edit any field.`;
        fixBannerEl.classList.remove('hidden');
    }

    if (fixBannerEl) {
        const unmatchedSuffix = unmatchedLabels.length > 0
            ? ` Not mapped in current editor: ${unmatchedLabels.slice(0, 3).join(', ')}${unmatchedLabels.length > 3 ? '…' : ''}.`
            : '';

        fixBannerEl.textContent = highlighted > 0
            ? `Fix mode: Element ${pending.elementId}. Highlighted ${highlighted} field(s) matching ${matchedTargetIds.size}/${entries.length} audit-missing item(s). You can still edit any field.${unmatchedSuffix}`
            : `Fix mode: Element ${pending.elementId}. Could not map audit-missing fields in the current editor.${unmatchedSuffix} You can still edit any field.`;
        fixBannerEl.classList.remove('hidden');
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

categoryFilterEl?.addEventListener('change', () => {
    render();
});

auditViewModeEl?.addEventListener('change', () => {
    render();
});

if (auditScoreModeEl) {
    auditScoreModeEl.addEventListener('change', () => {
        const mode = auditScoreModeEl.value;
        if (window.chrome && window.chrome.webview) {
            window.chrome.webview.postMessage({
                action: 'setAuditScoreMode',
                mode: mode
            });
        }
    });
}

// Table row click (event delegation)
tableBody.addEventListener('click', (e) => {
    const btn = e.target.closest('.locate-btn');
    if (btn) {
        const elementId = parseInt(btn.dataset.id, 10);
        if (!isNaN(elementId)) {
            selectElement(elementId);
        }
        return;
    }

    const fixBtn = e.target.closest('.fix-btn');
    if (fixBtn) {
        const elementId = parseInt(fixBtn.dataset.id, 10);
        if (!isNaN(elementId)) {
            let missingFields = [];
            try {
                const raw = decodeURIComponent(fixBtn.dataset.missingFields || '');
                missingFields = raw ? JSON.parse(raw) : [];
            } catch {
                missingFields = decodeURIComponent(fixBtn.dataset.missingFields || '')
                    .split('|')
                    .map(f => f.trim())
                    .filter(Boolean);
            }
            startFixFromAudit(elementId, missingFields);
        }
        return;
    }

    // Clicking a row (but not the dropdown/select button) triggers 2D view request
    const row = e.target.closest('tr');
    if (row && !e.target.closest('.locate-btn') && !e.target.closest('.fix-btn') && !e.target.closest('.view-select') && !e.target.closest('.open-view-btn')) {
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
            if (skipNextEditorSelectionRefresh) {
                skipNextEditorSelectionRefresh = false;
            } else {
                requestSelectedElements();
            }
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
    pendingFix: null,
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

        applyFixTargets();
        updateApplyPreviews();
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
    updateApplyPreviews();
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
        inputHtml = `<input type="text" id="${fieldId}" placeholder="${placeholder}" data-cobie-key="${cobieKey}" data-revit-param="${field.revitParam || ''}" data-label="${escapeHtml(label)}" data-type="${dataType}" data-required="${isRequired}" class="cobie-field-input">`;
    } else if (dataType === 'number') {
        inputHtml = `<input type="number" id="${fieldId}" placeholder="${placeholder}" data-cobie-key="${cobieKey}" data-revit-param="${field.revitParam || ''}" data-label="${escapeHtml(label)}" data-type="${dataType}" data-required="${isRequired}" class="cobie-field-input">`;
    } else if (field.options && field.options.length > 0) {
        const optionsHtml = field.options.map(o => `<option value="${o}">${o}</option>`).join('');
        inputHtml = `<select id="${fieldId}" data-cobie-key="${cobieKey}" data-revit-param="${field.revitParam || ''}" data-label="${escapeHtml(label)}" data-type="${dataType}" data-required="${isRequired}" class="cobie-field-input"><option value="">Select...</option>${optionsHtml}</select>`;
    } else {
        inputHtml = `<input type="text" id="${fieldId}" placeholder="${placeholder}" data-cobie-key="${cobieKey}" data-revit-param="${field.revitParam || ''}" data-label="${escapeHtml(label)}" data-type="${dataType}" data-required="${isRequired}" class="cobie-field-input">`;
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

    updateApplyPreviews();
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
    
    const { fieldValues, includedInputs } = collectCobieFieldValues('instance', true);
    if (Object.keys(fieldValues).length === 0) {
        showToast('No values to apply', 'warning');
        return;
    }

    const invalidCobieFields = getInvalidCobieFields(includedInputs);
    if (invalidCobieFields.length > 0) {
        showToast(`Fix invalid field values: ${formatInvalidFieldList(invalidCobieFields)}`, 'error');
        return;
    }
    
    const writeAliases = document.getElementById('writeAliases')?.checked ?? true;
    const elementIds = editorState.selectedElements.map(e => e.elementId);

    const message = `Apply ${Object.keys(fieldValues).length} field(s) to ${elementIds.length} selected element(s)?`;
    if (!confirmSafeApply(message)) return;
    
    setCobieFieldValues(elementIds, fieldValues, writeAliases);
});

document.getElementById('applyCobieType')?.addEventListener('click', () => {
    if (!requireCobieEnsured()) return;
    if (!editorState.currentTypeId) {
        showToast('No type selected', 'error');
        return;
    }
    
    const { fieldValues, includedInputs } = collectCobieFieldValues('type', true);
    if (Object.keys(fieldValues).length === 0) {
        showToast('No values to apply', 'warning');
        return;
    }

    const invalidCobieFields = getInvalidCobieFields(includedInputs);
    if (invalidCobieFields.length > 0) {
        showToast(`Fix invalid field values: ${formatInvalidFieldList(invalidCobieFields)}`, 'error');
        return;
    }
    
    const writeAliases = document.getElementById('writeAliases')?.checked ?? true;
    const affected = editorState.currentTypeInfo?.typeInstanceCount || '?';
    const message = `Apply ${Object.keys(fieldValues).length} type field(s)? This affects ${affected} instance(s).`;
    if (!confirmSafeApply(message)) return;
    
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
    
    const { fieldValues, includedInputs } = collectCobieFieldValues('bulk', true);
    if (Object.keys(fieldValues).length === 0) {
        showToast('No values to apply', 'warning');
        return;
    }

    const invalidCobieFields = getInvalidCobieFields(includedInputs);
    if (invalidCobieFields.length > 0) {
        showToast(`Fix invalid field values: ${formatInvalidFieldList(invalidCobieFields)}`, 'error');
        return;
    }
    
    const onlyBlanks = document.getElementById('cobieOnlyFillBlanks')?.checked ?? true;
    const writeAliases = document.getElementById('writeAliases')?.checked ?? true;
    const targetCount = readElementCount('cobieCategoryCount');
    const scopeText = targetCount !== null ? `${targetCount} element(s)` : 'selected category elements';
    const modeText = onlyBlanks ? 'Fill blank values only' : 'Overwrite existing values';
    const message = `Apply ${Object.keys(fieldValues).length} field(s) to ${scopeText}? ${modeText}.`;
    if (!confirmSafeApply(message)) return;
    
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

function collectCobieFieldValues(prefix, includeMetadata = false) {
    const fieldValues = {};
    const includedInputs = [];
    const inputs = document.querySelectorAll(`#cobie${capitalize(prefix)}FieldsGrid .cobie-field-input, #cobie${capitalize(prefix)}Form .cobie-field-input`);
    
    // Fallback for bulk fields which might be in a different container
    const bulkInputs = prefix === 'bulk' ? document.querySelectorAll('#cobieBulkFieldsGrid .cobie-field-input') : [];
    const allInputs = inputs.length > 0 ? inputs : bulkInputs;
    
    allInputs.forEach(input => {
        const value = input.value?.trim();
        const cobieKey = input.dataset.cobieKey;
        const revitParam = input.dataset.revitParam;
        const isModified = input.getAttribute('data-user-modified') === 'true';
        const dataType = (input.dataset?.type || '').toLowerCase();
        let normalizedValue = value;
        if (dataType === 'date' && value) {
            const normalized = normalizeDateValue(value);
            if (normalized) {
                normalizedValue = normalized;
            }
        }
        
        if (!cobieKey && !revitParam) return;
        
        if (isModified && normalizedValue) {
            fieldValues[cobieKey || revitParam] = {
                value: normalizedValue,
                cobieKey: cobieKey,
                revitParam: revitParam
            };
            includedInputs.push(input);
        }
    });
    
    if (includeMetadata) {
        return { fieldValues, includedInputs };
    }
    return fieldValues;
}

function capitalize(str) {
    return str.charAt(0).toUpperCase() + str.slice(1);
}

function confirmSafeApply(message) {
    if (typeof window.confirm !== 'function') return true;
    return window.confirm(message);
}

function readElementCount(elementId) {
    const text = document.getElementById(elementId)?.textContent || '';
    const match = text.match(/(\d+)/);
    return match ? parseInt(match[1], 10) : null;
}

function formatDateValue(year, month, day) {
    if (!Number.isInteger(year) || year < 1) return null;
    if (!Number.isInteger(month) || month < 1 || month > 12) return null;
    if (!Number.isInteger(day) || day < 1 || day > 31) return null;
    const date = new Date(Date.UTC(year, month - 1, day));
    if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) {
        return null;
    }
    const yyyy = String(year).padStart(4, '0');
    const mm = String(month).padStart(2, '0');
    const dd = String(day).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
}

function normalizeDateValue(value) {
    const trimmed = (value || '').trim();
    if (!trimmed) return null;

    const isoWithTime = trimmed.match(/^(\d{4}-\d{2}-\d{2})(?:[T ].*)$/);
    if (isoWithTime) {
        return normalizeDateValue(isoWithTime[1]);
    }

    let match = trimmed.match(/^(\d{4})(\d{2})(\d{2})$/);
    if (match) {
        return formatDateValue(parseInt(match[1], 10), parseInt(match[2], 10), parseInt(match[3], 10));
    }

    match = trimmed.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})$/);
    if (match) {
        return formatDateValue(parseInt(match[1], 10), parseInt(match[2], 10), parseInt(match[3], 10));
    }

    match = trimmed.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
    if (match) {
        const first = parseInt(match[1], 10);
        const second = parseInt(match[2], 10);
        const year = parseInt(match[3], 10);
        let month = first;
        let day = second;
        if (first > 12 && second <= 12) {
            day = first;
            month = second;
        } else if (second > 12 && first <= 12) {
            day = second;
            month = first;
        }
        return formatDateValue(year, month, day);
    }

    return null;
}

function isValidDateValue(value) {
    return !!normalizeDateValue(value);
}

function normalizeDateInputValue(input) {
    if (!input) return null;
    const value = input.value?.trim() || '';
    if (!value) {
        input.classList.remove('invalid-value');
        return null;
    }
    const normalized = normalizeDateValue(value);
    if (normalized) {
        if (input.getAttribute('data-user-modified') === 'true' && normalized !== value) {
            input.value = normalized;
        }
        input.classList.remove('invalid-value');
        return normalized;
    }
    input.classList.add('invalid-value');
    return null;
}

function validateCobieInput(input) {
    const value = input?.value?.trim() || '';
    const dataType = (input?.dataset?.type || 'string').toLowerCase();

    if (!value) {
        input?.classList.remove('invalid-value');
        return true;
    }

    let valid = true;
    if (dataType === 'date') {
        valid = isValidDateValue(value);
    } else if (dataType === 'number') {
        valid = Number.isFinite(Number(value));
    }

    input?.classList.toggle('invalid-value', !valid);
    return valid;
}

function validateCobieInputs(inputs) {
    if (!Array.isArray(inputs) || inputs.length === 0) return true;
    return inputs.every(validateCobieInput);
}

function getInputLabel(inputElement) {
    if (!inputElement) return '';
    const datasetLabel = inputElement.dataset?.label;
    if (datasetLabel) return datasetLabel.trim();
    const formGroup = inputElement.closest('.form-group');
    const labelEl = formGroup?.querySelector('label');
    const labelText = labelEl?.textContent?.trim();
    if (labelText) return labelText;
    const cobieKey = inputElement.dataset?.cobieKey;
    if (cobieKey) return cobieKey.trim();
    const revitParam = inputElement.dataset?.revitParam;
    if (revitParam) return revitParam.trim();
    return inputElement.id || '';
}

function formatInvalidFieldList(labels, maxItems = 4) {
    const uniqueLabels = Array.from(new Set(labels.filter(Boolean)));
    if (uniqueLabels.length === 0) return 'unknown fields';
    if (uniqueLabels.length <= maxItems) return uniqueLabels.join(', ');
    return `${uniqueLabels.slice(0, maxItems).join(', ')} +${uniqueLabels.length - maxItems} more`;
}

function getInvalidCobieFields(inputs) {
    if (!Array.isArray(inputs) || inputs.length === 0) return [];
    const invalidLabels = [];
    inputs.forEach(inputElement => {
        if (!validateCobieInput(inputElement)) {
            const label = getInputLabel(inputElement);
            if (label) invalidLabels.push(label);
        }
    });
    return Array.from(new Set(invalidLabels));
}

function getInvalidLegacyFields(inputEntries) {
    if (!Array.isArray(inputEntries) || inputEntries.length === 0) return [];
    const invalidLabels = [];
    inputEntries.forEach(entry => {
        const inputElement = entry.input;
        const fieldName = entry.field;
        if (!validateLegacyInput(fieldName, inputElement)) {
            const label = getInputLabel(inputElement) || fieldName;
            if (label) invalidLabels.push(label);
        }
    });
    return Array.from(new Set(invalidLabels));
}

function validateLegacyInput(fieldName, input) {
    if (!input) return true;
    const value = input.value?.trim() || '';

    if (!value) {
        input.classList.remove('invalid-value');
        return true;
    }

    let valid = true;
    if (legacyDateFields.has(fieldName)) {
        valid = isValidDateValue(value);
    } else if (fieldName === 'FM_PMFrequencyDays') {
        valid = Number.isInteger(Number(value)) && Number(value) > 0;
    }

    input.classList.toggle('invalid-value', !valid);
    return valid;
}

function collectLegacySelectedParams() {
    const params = {};
    const includedInputs = [];
    const fields = ['FM_Barcode', 'FM_UniqueAssetId', 'FM_InstallationDate', 'FM_WarrantyStart',
        'FM_WarrantyEnd', 'FM_Criticality', 'FM_Trade', 'FM_PMTemplateId',
        'FM_PMFrequencyDays', 'FM_Building', 'FM_LocationSpace'];

    fields.forEach(field => {
        const input = document.getElementById(`edit_${field}`);
        if (!input) return;
        const value = input.value?.trim() || '';
        if (!value) return;
        if (input.getAttribute('data-user-modified') !== 'true') return;

        let normalizedValue = value;
        if (legacyDateFields.has(field)) {
            const normalized = normalizeDateValue(value);
            if (normalized) {
                normalizedValue = normalized;
            }
        }
        params[field] = normalizedValue;
        includedInputs.push({ field, input });
    });

    return { params, includedInputs };
}

function collectLegacyCategoryParams() {
    const params = {};
    const includedInputs = [];
    const fields = ['FM_Building', 'FM_Trade', 'FM_PMFrequencyDays', 'FM_Criticality'];

    fields.forEach(field => {
        const input = document.getElementById(`bulk_${field}`);
        if (!input) return;
        const value = input.value?.trim() || '';
        if (!value) return;
        if (input.getAttribute('data-user-modified') !== 'true') return;
        let normalizedValue = value;
        if (legacyDateFields.has(field)) {
            const normalized = normalizeDateValue(value);
            if (normalized) {
                normalizedValue = normalized;
            }
        }
        params[field] = normalizedValue;
        includedInputs.push({ field, input });
    });

    return { params, includedInputs };
}

function collectLegacyTypeParams() {
    const params = {};
    const manufacturer = document.getElementById('type_Manufacturer');
    const model = document.getElementById('type_Model');
    const typeMark = document.getElementById('type_TypeMark');

    if (manufacturer?.value?.trim() && manufacturer.getAttribute('data-user-modified') === 'true') params.Manufacturer = manufacturer.value.trim();
    if (model?.value?.trim() && model.getAttribute('data-user-modified') === 'true') params.Model = model.value.trim();
    if (typeMark?.value?.trim() && typeMark.getAttribute('data-user-modified') === 'true') params.TypeMark = typeMark.value.trim();

    return params;
}

function updateApplyPreviews() {
    const selectedCount = editorState.selectedElements.length;
    const categoryCount = readElementCount('categoryCount');
    const cobieCategoryCount = readElementCount('cobieCategoryCount');
    const legacySelected = collectLegacySelectedParams();
    const legacyCategory = collectLegacyCategoryParams();
    const legacyType = collectLegacyTypeParams();
    const cobieSelected = collectCobieFieldValues('instance', true);
    const cobieType = collectCobieFieldValues('type', true);
    const cobieCategory = collectCobieFieldValues('bulk', true);

    if (legacySelectedPreviewEl) {
        legacySelectedPreviewEl.textContent = selectedCount > 0 && Object.keys(legacySelected.params).length > 0
            ? `Will update ${Object.keys(legacySelected.params).length} field(s) on ${selectedCount} element(s).`
            : '';
    }

    if (legacyCategoryPreviewEl) {
        const categoryValue = document.getElementById('bulkCategory')?.value || '';
        if (categoryValue && Object.keys(legacyCategory.params).length > 0) {
            const mode = document.getElementById('onlyFillBlanks')?.checked ? 'fill blanks only' : 'overwrite';
            legacyCategoryPreviewEl.textContent = categoryCount !== null
                ? `Will update ${Object.keys(legacyCategory.params).length} field(s) on ${categoryCount} element(s) (${mode}).`
                : `Will update ${Object.keys(legacyCategory.params).length} field(s) (${mode}).`;
        } else {
            legacyCategoryPreviewEl.textContent = '';
        }
    }

    if (legacyTypePreviewEl) {
        const affected = editorState.currentTypeInfo?.typeInstanceCount || null;
        legacyTypePreviewEl.textContent = editorState.currentTypeId && Object.keys(legacyType).length > 0
            ? `Will update ${Object.keys(legacyType).length} field(s)${affected ? ` affecting ${affected} instance(s)` : ''}.`
            : '';
    }

    if (cobieSelectedPreviewEl) {
        cobieSelectedPreviewEl.textContent = selectedCount > 0 && Object.keys(cobieSelected.fieldValues).length > 0
            ? `Will update ${Object.keys(cobieSelected.fieldValues).length} field(s) on ${selectedCount} element(s).`
            : '';
    }

    if (cobieTypePreviewEl) {
        const affected = editorState.currentTypeInfo?.typeInstanceCount || null;
        cobieTypePreviewEl.textContent = editorState.currentTypeId && Object.keys(cobieType.fieldValues).length > 0
            ? `Will update ${Object.keys(cobieType.fieldValues).length} field(s)${affected ? ` affecting ${affected} instance(s)` : ''}.`
            : '';
    }

    if (cobieCategoryPreviewEl) {
        const categoryValue = document.getElementById('cobieBulkCategory')?.value || '';
        if (categoryValue && Object.keys(cobieCategory.fieldValues).length > 0) {
            const mode = document.getElementById('cobieOnlyFillBlanks')?.checked ? 'fill blanks only' : 'overwrite';
            cobieCategoryPreviewEl.textContent = cobieCategoryCount !== null
                ? `Will update ${Object.keys(cobieCategory.fieldValues).length} field(s) on ${cobieCategoryCount} element(s) (${mode}).`
                : `Will update ${Object.keys(cobieCategory.fieldValues).length} field(s) (${mode}).`;
        } else {
            cobieCategoryPreviewEl.textContent = '';
        }
    }
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
        resetInputFromMultiState(input);
        
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

        validateCobieInput(input);
    });
    
    // Populate type fields
    editorState.cobieFields.type.forEach(field => {
        const fieldId = `type_${(field.cobieKey || '').replace(/\./g, '_')}`;
        const input = document.getElementById(fieldId);
        if (!input) return;
        resetInputFromMultiState(input);
        
        let value = cobieTypeParams[field.cobieKey] || typeParams[field.revitParam] || typeParams[field.cobieKey] ||
                    (field.aliasParams || []).map(a => typeParams[a]).find(v => v);
        
        input.value = value || '';
        validateCobieInput(input);
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
    updateApplyPreviews();
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

function requestFixElementSelection(elementId) {
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage({
            action: 'focusFixElement',
            elementId: elementId
        });
    }
}

function setSelectionElements(elementIds) {
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage({
            action: 'setSelectionElements',
            elementIds: elementIds
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
    editorState.currentTypeId = null;
    editorState.currentTypeInfo = null;
    renderSelectionTray();
    
    const infoEl = document.getElementById('selectedElementInfo');
    const formEl = document.getElementById('instanceParamsForm');
    const typeInfoEl = document.getElementById('typeInfo');
    const typeFormEl = document.getElementById('typeParamsForm');
    
    // COBie editor elements
    const cobieInstanceForm = document.getElementById('cobieInstanceForm');
    const cobieTypeInfo = document.getElementById('cobieTypeInfo');
    const cobieTypeForm = document.getElementById('cobieTypeForm');
    
    if (elements.length === 0) {
        editorState.currentTypeId = null;
        editorState.currentTypeInfo = null;
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
        applyFixTargets();
        updateApplyPreviews();
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
            ['type_Manufacturer', 'type_Model', 'type_TypeMark'].forEach(id => {
                const input = document.getElementById(id);
                if (!input) return;
                input.removeAttribute('data-user-modified');
                input.classList.remove('invalid-value', 'fix-target');
            });
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
            editorState.currentTypeId = null;
            editorState.currentTypeInfo = null;
            typeInfoEl.innerHTML = `<p class="placeholder-text multi-type-warning">⚠ ${typeIds.length} different types selected - type editing disabled</p>`;
            typeFormEl?.classList.add('hidden');
            
            if (cobieTypeInfo) cobieTypeInfo.innerHTML = `<p class="placeholder-text multi-type-warning">⚠ ${typeIds.length} different types selected</p>`;
            cobieTypeForm?.classList.add('hidden');
        }
        
        // Clear computed values for multi-select (they vary per element)
        clearComputedValues();
    }

    applyFixTargets();
    updateApplyPreviews();
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

function rememberBasePlaceholder(input) {
    if (!input?.dataset) return '';
    if (input.dataset.basePlaceholder === undefined) {
        input.dataset.basePlaceholder = input.getAttribute('placeholder') || '';
    }
    return input.dataset.basePlaceholder || '';
}

function resetInputFromMultiState(input) {
    if (!input) return;
    const basePlaceholder = rememberBasePlaceholder(input);
    input.placeholder = basePlaceholder;
    input.classList.remove('common-value', 'varies-value', 'all-blank', 'user-modified', 'fix-target', 'invalid-value');
    input.removeAttribute('data-original-value');
    input.removeAttribute('data-user-modified');
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
    rememberBasePlaceholder(input);

    // Reset state
    input.classList.remove('common-value', 'varies-value', 'all-blank', 'invalid-value');
    input.removeAttribute('data-original-value');
    input.removeAttribute('data-user-modified');
    input.classList.remove('user-modified', 'fix-target');
    
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
    const target = e?.target || e;
    if (!target || typeof target.setAttribute !== 'function') return;
    target.setAttribute('data-user-modified', 'true');
    target.classList.add('user-modified');
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
        resetInputFromMultiState(input);
        
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
        resetInputFromMultiState(input);
        
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
    const mappings = {
        FM_Barcode: ip.FM_Barcode,
        FM_UniqueAssetId: ip.FM_UniqueAssetId,
        FM_InstallationDate: ip.FM_InstallationDate,
        FM_WarrantyStart: ip.FM_WarrantyStart,
        FM_WarrantyEnd: ip.FM_WarrantyEnd,
        FM_Criticality: ip.FM_Criticality,
        FM_Trade: ip.FM_Trade,
        FM_PMTemplateId: ip.FM_PMTemplateId,
        FM_PMFrequencyDays: ip.FM_PMFrequencyDays,
        FM_Building: ip.FM_Building,
        FM_LocationSpace: ip.FM_LocationSpace
    };

    Object.entries(mappings).forEach(([field, value]) => {
        const input = document.getElementById(`edit_${field}`);
        if (!input) return;
        resetInputFromMultiState(input);
        input.value = value || '';
        validateLegacyInput(field, input);
    });
    
    // Update computed values
    const computed = el.computed || {};
    document.getElementById('computed_UniqueId').textContent = computed.UniqueId || '--';
    document.getElementById('computed_Level').textContent = computed.Level || '--';
    document.getElementById('computed_RoomSpace').textContent = computed.RoomSpace || '--';
}

function clearCobieInstanceForm() {
    document.querySelectorAll('.cobie-field-input').forEach(input => {
        resetInputFromMultiState(input);
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

    updateApplyPreviews();
}

function receiveOperationResult(msg) {
    showToast(msg.message, msg.success ? 'success' : 'error');
    
    // Refresh selection after parameter updates
    if (msg.success && msg.refreshAudit) {
        requestSelectedElements();
        // Trigger audit refresh to update the Audit Results tab
        requestAuditRefresh();
    }

    updateApplyPreviews();
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
    [
        'edit_FM_Barcode',
        'edit_FM_UniqueAssetId',
        'edit_FM_InstallationDate',
        'edit_FM_WarrantyStart',
        'edit_FM_WarrantyEnd',
        'edit_FM_Criticality',
        'edit_FM_Trade',
        'edit_FM_PMTemplateId',
        'edit_FM_PMFrequencyDays',
        'edit_FM_Building',
        'edit_FM_LocationSpace'
    ].forEach(id => {
        const el = document.getElementById(id);
        if (!el) return;
        resetInputFromMultiState(el);
        el.value = '';
    });
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

document.getElementById('tab-editor')?.addEventListener('input', (e) => {
    const target = e.target;
    if (!(target instanceof HTMLInputElement || target instanceof HTMLSelectElement)) return;

    if (target.type !== 'checkbox' && (target.classList.contains('cobie-field-input') || target.id?.startsWith('edit_') || target.id?.startsWith('bulk_') || target.id?.startsWith('type_'))) {
        markFieldAsModified(target);
    }

    if (target.classList.contains('cobie-field-input')) {
        validateCobieInput(target);
    } else if (target.id?.startsWith('edit_')) {
        validateLegacyInput(target.id.replace('edit_', ''), target);
    } else if (target.id?.startsWith('bulk_')) {
        validateLegacyInput(target.id.replace('bulk_', ''), target);
    }

    updateApplyPreviews();
});

document.getElementById('tab-editor')?.addEventListener('change', (e) => {
    const target = e.target;
    if ((target instanceof HTMLInputElement || target instanceof HTMLSelectElement)
        && target.type !== 'checkbox'
        && (target.classList.contains('cobie-field-input') || target.id?.startsWith('edit_') || target.id?.startsWith('bulk_') || target.id?.startsWith('type_'))) {
        markFieldAsModified(target);
    }
    updateApplyPreviews();
});

document.getElementById('tab-editor')?.addEventListener('focusout', (e) => {
    const target = e.target;
    if (!(target instanceof HTMLInputElement || target instanceof HTMLSelectElement)) return;
    if (target.type === 'checkbox') return;

    if (target.classList.contains('cobie-field-input')) {
        const dataType = (target.dataset?.type || '').toLowerCase();
        if (dataType === 'date') {
            normalizeDateInputValue(target);
        }
        validateCobieInput(target);
    } else if (target.id?.startsWith('edit_') || target.id?.startsWith('bulk_') || target.id?.startsWith('type_')) {
        const fieldName = target.id.replace(/^(edit_|bulk_|type_)/, '');
        if (legacyDateFields.has(fieldName)) {
            normalizeDateInputValue(target);
        }
        validateLegacyInput(fieldName, target);
    }

    updateApplyPreviews();
});

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

selectionTrayToggleEl?.addEventListener('click', () => {
    selectionTrayExpanded = !selectionTrayExpanded;
    const icon = selectionTrayToggleEl.querySelector('.collapse-icon');
    if (icon) icon.textContent = selectionTrayExpanded ? '▲' : '▼';
    renderSelectionTray();
});

selectionTrayListEl?.addEventListener('click', (e) => {
    const removeBtn = e.target.closest('.selection-tray-remove');
    if (!removeBtn) return;

    const elementId = Number(removeBtn.dataset.id);
    if (!Number.isFinite(elementId)) return;

    const nextIds = getCurrentSelectionIds().filter(id => id !== elementId);
    selectionTrayCheckedIds.delete(elementId);
    applySelectionIdsFromTray(nextIds, `Removed element ${elementId} from selection`);
});

selectionTrayListEl?.addEventListener('change', (e) => {
    const check = e.target.closest('.selection-tray-check');
    if (!check) return;

    const elementId = Number(check.dataset.id);
    if (!Number.isFinite(elementId)) return;

    if (check.checked) {
        selectionTrayCheckedIds.add(elementId);
    } else {
        selectionTrayCheckedIds.delete(elementId);
    }
    updateSelectionTrayActionState();
});

selectionTraySelectAllEl?.addEventListener('change', () => {
    const currentIds = getCurrentSelectionIds();
    selectionTrayCheckedIds.clear();
    if (selectionTraySelectAllEl.checked) {
        currentIds.forEach(id => selectionTrayCheckedIds.add(id));
    }
    renderSelectionTray();
});

removeCheckedSelectionEl?.addEventListener('click', () => {
    const checkedIds = new Set(selectionTrayCheckedIds);
    if (checkedIds.size === 0) return;

    const nextIds = getCurrentSelectionIds().filter(id => !checkedIds.has(id));
    selectionTrayCheckedIds.clear();
    applySelectionIdsFromTray(nextIds, `Removed ${checkedIds.size} selected element(s)`);
});

keepCheckedSelectionEl?.addEventListener('click', () => {
    const checkedIds = new Set(selectionTrayCheckedIds);
    if (checkedIds.size === 0) return;

    const nextIds = getCurrentSelectionIds().filter(id => checkedIds.has(id));
    selectionTrayCheckedIds.clear();
    applySelectionIdsFromTray(nextIds, `Kept ${nextIds.length} selected element(s)`);
});

clearSelectionListEl?.addEventListener('click', () => {
    if (editorState.selectedElements.length === 0) return;
    selectionTrayCheckedIds.clear();
    applySelectionIdsFromTray([], 'Selection cleared');
});

// Apply to Selected button
document.getElementById('applyToSelected')?.addEventListener('click', () => {
    if (editorState.selectedElements.length === 0) {
        showToast('No elements selected', 'error');
        return;
    }
    
    const { params, includedInputs } = collectLegacySelectedParams();
    
    if (Object.keys(params).length === 0) {
        showToast('No values to apply', 'error');
        return;
    }

    const invalidLegacyFields = getInvalidLegacyFields(includedInputs);
    if (invalidLegacyFields.length > 0) {
        showToast(`Fix invalid field values: ${formatInvalidFieldList(invalidLegacyFields)}`, 'error');
        return;
    }
    
    const elementIds = editorState.selectedElements.map(e => e.elementId);
    const message = `Apply ${Object.keys(params).length} field(s) to ${elementIds.length} selected element(s)?`;
    if (!confirmSafeApply(message)) return;
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
    updateApplyPreviews();
});

// Apply to Category button
document.getElementById('applyToCategory')?.addEventListener('click', () => {
    const category = document.getElementById('bulkCategory').value;
    if (!category) {
        showToast('Select a category first', 'error');
        return;
    }
    
    const { params, includedInputs } = collectLegacyCategoryParams();
    
    if (Object.keys(params).length === 0) {
        showToast('No values to apply', 'error');
        return;
    }

    const invalidLegacyFields = getInvalidLegacyFields(includedInputs);
    if (invalidLegacyFields.length > 0) {
        showToast(`Fix invalid field values: ${formatInvalidFieldList(invalidLegacyFields)}`, 'error');
        return;
    }
    
    const onlyBlanks = document.getElementById('onlyFillBlanks').checked;
    const targetCount = readElementCount('categoryCount');
    const modeText = onlyBlanks ? 'Fill blank values only' : 'Overwrite existing values';
    const scopeText = targetCount !== null ? `${targetCount} element(s)` : 'selected category elements';
    const message = `Apply ${Object.keys(params).length} field(s) to ${scopeText}? ${modeText}.`;
    if (!confirmSafeApply(message)) return;
    setCategoryParams(category, params, onlyBlanks);
});

// Update Type button
document.getElementById('updateType')?.addEventListener('click', () => {
    if (!editorState.currentTypeId) {
        showToast('No type selected', 'error');
        return;
    }
    
    const params = collectLegacyTypeParams();
    
    if (Object.keys(params).length === 0) {
        showToast('No values to update', 'error');
        return;
    }

    const affected = editorState.currentTypeInfo?.typeInstanceCount || '?';
    const message = `Apply ${Object.keys(params).length} type field(s)? This affects ${affected} instance(s).`;
    if (!confirmSafeApply(message)) return;
    
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
        requestAuditRefresh();
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
updateApplyPreviews();
