// FM Readiness - Audit Results UI

window.auditData = {
    summary: null,
    rows: []
};

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

// Fallback hook if ExecuteScriptAsync is used.
window.receiveAudit = receiveAudit;

function receiveAudit(data) {
    window.auditData = {
        summary: data.summary,
        rows: data.rows || []
    };
    render();
}

function render() {
    const { summary, rows } = window.auditData;

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

    const filteredRows = filterRows(rows);

    if (filteredRows.length > 0) {
        resultsTable.classList.add('visible');
        emptyState.classList.add('hidden');
        renderTable(filteredRows);
    } else if (rows.length > 0) {
        resultsTable.classList.add('visible');
        emptyState.classList.add('hidden');
        renderTable([]);
    } else {
        resultsTable.classList.remove('visible');
        emptyState.classList.remove('hidden');
    }
}

function filterRows(rows) {
    const searchTerm = searchInput.value.toLowerCase().trim();
    const missingOnly = showMissingOnly.checked;

    return rows.filter(row => {
        if (missingOnly && row.missingCount === 0) {
            return false;
        }

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

function getScoreClass(percent) {
    if (percent >= 100) return 'score-100';
    if (percent >= 75) return 'score-high';
    if (percent >= 50) return 'score-medium';
    return 'score-low';
}

function renderRowGroupBadges(groupScores) {
    if (!groupScores || Object.keys(groupScores).length === 0) {
        return '-';
    }

    return Object.entries(groupScores).map(([name, pct]) => {
        const scoreClass = getScoreClass(pct);
        const abbrev = name.substring(0, 3);
        return `<span class="group-badge-sm ${scoreClass}" title="${escapeHtml(name)}: ${pct}%">${abbrev} ${pct}%</span>`;
    }).join(' ');
}

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

function getStatusBadge(readinessPercent, missingCount) {
    if (readinessPercent >= 100) {
        return '<span class="status-badge ready">Ready</span>';
    } else if (missingCount > 0 && readinessPercent > 0) {
        return '<span class="status-badge partial">Partial</span>';
    } else {
        return '<span class="status-badge missing">Missing</span>';
    }
}

function getReadinessClass(percent) {
    if (percent >= 100) return 'readiness-100';
    if (percent >= 75) return 'readiness-high';
    if (percent >= 50) return 'readiness-medium';
    return 'readiness-low';
}

function escapeHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

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

// Event listeners
searchInput.addEventListener('input', () => {
    render();
});

showMissingOnly.addEventListener('change', () => {
    render();
});

tableBody.addEventListener('click', (e) => {
    const btn = e.target.closest('.select-btn');
    if (btn) {
        const elementId = parseInt(btn.dataset.id, 10);
        if (!isNaN(elementId)) {
            selectElement(elementId);
        }
    }

    const row = e.target.closest('tr');
    if (row && !e.target.closest('.select-btn') && !e.target.closest('.view-select')) {
        const firstCell = row.querySelector('td');
        const elementId = firstCell ? parseInt(firstCell.textContent, 10) : NaN;
        if (!isNaN(elementId) && !viewsByElementId.has(elementId)) {
            request2dViews(elementId);
        }
    }
});

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

// Initial render
render();

