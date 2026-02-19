const pages = {
    async exports(params) {
        const actions = document.getElementById('topbar-actions');
        const content = document.getElementById('page-content');
        setPage('Exports');

        const searchVal = params.q || '';
        actions.innerHTML = `
            <div class="search-box">
                ${icons.search}
                <input type="text" id="export-search" placeholder="Search drawing, part, user..." value="${esc(searchVal)}">
            </div>`;

        content.innerHTML = `<div class="loading">Loading exports</div>`;

        const searchInput = document.getElementById('export-search');
        let debounce;
        searchInput.addEventListener('input', () => {
            clearTimeout(debounce);
            debounce = setTimeout(() => router.go('exports', { q: searchInput.value }), 400);
        });

        try {
            const searchQ = searchVal ? `&search=${encodeURIComponent(searchVal)}` : '';
            const data = await api.get(`/api/exports?take=500${searchQ}`);

            if (data.items.length === 0) {
                content.innerHTML = `<div class="empty">No exports found.</div>`;
                return;
            }

            // Deduplicate: keep only the latest export per drawing number
            const seen = new Set();
            const unique = data.items.filter(e => {
                const dn = e.drawingNumber || '';
                if (seen.has(dn)) return false;
                seen.add(dn);
                return true;
            });

            // Group by equipment number (first token of drawing number)
            const groups = new Map();
            unique.forEach(e => {
                const dn = e.drawingNumber || '';
                const spaceIdx = dn.indexOf(' ');
                const equip = spaceIdx > 0 ? dn.substring(0, spaceIdx) : (dn || 'Other');
                if (!groups.has(equip)) groups.set(equip, []);
                groups.get(equip).push(e);
            });

            // Sort equipment groups by number descending (most recent equipment first)
            const sortedGroups = [...groups.entries()].sort((a, b) => {
                const numA = parseInt(a[0]) || 0;
                const numB = parseInt(b[0]) || 0;
                return numB - numA;
            });

            const uniqueEquip = sortedGroups.length;
            const uniqueDrawings = unique.length;
            setPage('Exports', `${uniqueDrawings} drawings / ${uniqueEquip} equipment`);

            const groupsHtml = sortedGroups.map(([equip, items], gi) => {
                const totalBom = items.reduce((s, e) => s + e.bomItemCount, 0);

                const rows = items.map((e, i) => {
                    const dn = e.drawingNumber || '';
                    const spaceIdx = dn.indexOf(' ');
                    const drawingPart = spaceIdx > 0 ? dn.substring(spaceIdx + 1) : dn;

                    return `
                        <tr class="clickable" onclick="router.go('export-detail', {id: ${e.id}})" style="animation: fadeSlideIn 0.2s ease ${0.02 * i}s forwards; opacity: 0">
                            <td style="font-family:var(--font-mono);color:var(--text-dim);font-size:13px">${e.id}</td>
                            <td><strong>${esc(drawingPart) || '<span style="color:var(--text-dim)">\u2014</span>'}</strong></td>
                            <td style="color:var(--text-secondary);font-size:13px">${esc(e.title) || ''}</td>
                            <td><span class="badge badge-count">${e.bomItemCount}</span></td>
                            <td style="color:var(--text-secondary)">${esc(e.exportedBy)}</td>
                            <td style="font-family:var(--font-mono);font-size:13px;color:var(--text-secondary);white-space:nowrap">${fmtDate(e.exportedAt)}</td>
                        </tr>`;
                }).join('');

                return `
                    <div class="equip-group animate-in" id="equip-${esc(equip)}" style="animation-delay:${0.04 * gi}s">
                        <div class="equip-header" onclick="toggleEquipGroup('equip-${esc(equip)}')">
                            <span class="chevron-toggle open" id="equip-${esc(equip)}-icon">${icons.chevron}</span>
                            <span class="equip-header-number">${esc(equip)}</span>
                            <div class="equip-header-meta">
                                <span class="equip-header-stat"><strong>${items.length}</strong> exports</span>
                                <span class="equip-header-stat"><strong>${totalBom}</strong> items</span>
                            </div>
                        </div>
                        <div class="equip-body">
                            <table>
                                <thead><tr>
                                    <th style="width:50px">#</th>
                                    <th>Drawing</th>
                                    <th>Title</th>
                                    <th style="width:80px">Items</th>
                                    <th>Exported By</th>
                                    <th style="width:180px">Date</th>
                                </tr></thead>
                                <tbody>${rows}</tbody>
                            </table>
                        </div>
                    </div>`;
            }).join('');

            content.innerHTML = `
                <div class="stats-grid">
                    <div class="stat-card animate-in"><div class="stat-label">Drawings</div><div class="stat-value">${uniqueDrawings}</div></div>
                    <div class="stat-card animate-in"><div class="stat-label">Equipment</div><div class="stat-value">${uniqueEquip}</div></div>
                </div>
                ${groupsHtml}`;
        } catch (err) {
            content.innerHTML = `<div class="empty">Error: ${esc(err.message)}</div>`;
        }
    },

    async exportDetail(id) {
        const actions = document.getElementById('topbar-actions');
        const content = document.getElementById('page-content');
        setPage('Loading...');
        actions.innerHTML = '';
        content.innerHTML = `<div class="loading">Loading export</div>`;

        try {
            const exp = await api.get(`/api/exports/${id}`);
            setPage(exp.drawingNumber || `Export #${exp.id}`, 'export detail');

            const dxfCount = (exp.bomItems || []).filter(b => b.cutTemplate?.contentHash).length;

            const bomRows = (exp.bomItems || []).map((b, i) => {
                const hasDetails = b.cutTemplate || b.formProgram;
                const toggleId = `bom-${b.id}`;
                return `
                    <tr class="${hasDetails ? 'clickable' : ''}" ${hasDetails ? `onclick="toggleBomRow('${toggleId}')"` : ''} style="animation: fadeSlideIn 0.25s ease ${0.03 * i}s forwards; opacity: 0">
                        <td style="width:32px">${hasDetails ? `<span class="chevron-toggle" id="${toggleId}-icon">${icons.chevron}</span>` : ''}</td>
                        <td style="font-family:var(--font-mono);font-weight:600;color:var(--cyan)">${esc(b.itemNo)}</td>
                        <td><strong>${esc(b.partName)}</strong></td>
                        <td style="color:var(--text-secondary)">${esc(b.description)}</td>
                        <td><span style="font-family:var(--font-mono);font-size:13px">${esc(b.material)}</span></td>
                        <td style="font-family:var(--font-mono);text-align:center">${b.qty ?? ''}</td>
                        <td style="font-family:var(--font-mono);text-align:center">${b.totalQty ?? ''}</td>
                        <td>
                            ${b.cutTemplate ? `<span class="badge badge-cyan">${icons.laser} DXF</span>` : ''}
                            ${b.formProgram ? `<span class="badge badge-amber">${icons.bend} Form</span>` : ''}
                        </td>
                    </tr>
                    ${hasDetails ? `<tr class="bom-expand-row" id="${toggleId}" style="display:none"><td colspan="8">${renderBomDetails(b)}</td></tr>` : ''}`;
            }).join('');

            content.innerHTML = `
                <a class="back-link" onclick="router.go('exports')">${icons.back} Back to exports</a>

                <div class="card animate-in" style="margin-bottom:20px">
                    <div class="card-header">Export Information</div>
                    <div class="card-body">
                        <div class="detail-grid">
                            <div class="detail-field"><label>Drawing Number</label><div class="value">${esc(exp.drawingNumber) || '\u2014'}</div></div>
                            ${exp.title ? `<div class="detail-field"><label>Title</label><div class="value">${esc(exp.title)}</div></div>` : ''}
                            <div class="detail-field"><label>Exported By</label><div class="value">${esc(exp.exportedBy)}</div></div>
                            <div class="detail-field"><label>Date</label><div class="value mono">${fmtDate(exp.exportedAt)}</div></div>
                            <div class="detail-field"><label>Source File</label><div class="value mono">${esc(exp.sourceFilePath)}</div></div>
                        </div>
                    </div>
                </div>

                <div class="card animate-in">
                    <div class="card-header">
                        BOM Items
                        <span class="badge badge-count">${exp.bomItems?.length || 0} items</span>
                        <span style="margin-left:auto;display:flex;gap:6px">
                            ${exp.pdfContentHash ? `<a class="btn btn-amber btn-sm" href="/api/filebrowser/download?hash=${encodeURIComponent(exp.pdfContentHash)}&ext=pdf&name=${encodeURIComponent((exp.drawingNumber || 'drawing') + '.pdf')}">${icons.download} PDF</a>` : ''}
                            ${dxfCount > 0 ? `<a class="btn btn-cyan btn-sm" href="/api/exports/${exp.id}/download-dxfs">${icons.download} All DXFs</a>` : ''}
                        </span>
                    </div>
                    ${exp.bomItems?.length ? `
                    <table>
                        <thead><tr>
                            <th style="width:32px"></th>
                            <th style="width:60px">Item</th>
                            <th>Part Name</th>
                            <th>Description</th>
                            <th>Material</th>
                            <th style="width:50px;text-align:center">Qty</th>
                            <th style="width:55px;text-align:center">Total</th>
                            <th style="width:120px">Data</th>
                        </tr></thead>
                        <tbody>${bomRows}</tbody>
                    </table>` : '<div class="empty">No BOM items for this export.</div>'}
                </div>`;
        } catch (err) {
            content.innerHTML = `<div class="empty">Error: ${esc(err.message)}</div>`;
        }
    },

    async drawings() {
        const actions = document.getElementById('topbar-actions');
        const content = document.getElementById('page-content');
        setPage('Drawings');
        actions.innerHTML = '';
        content.innerHTML = `<div class="loading">Loading drawings</div>`;

        try {
            const numbers = await api.get('/api/exports/drawing-numbers');
            if (numbers.length === 0) {
                content.innerHTML = `<div class="empty">No drawings found.</div>`;
                return;
            }

            numbers.sort();
            setPage('Drawings', `${numbers.length} drawings`);

            const cards = numbers.map((d, i) => `
                <div class="drawing-card" onclick="router.go('drawing-detail', {id: '${encodeURIComponent(d)}'})" style="animation: fadeSlideIn 0.3s ease ${0.025 * Math.min(i, 20)}s forwards; opacity: 0">
                    <div class="drawing-card-title">${esc(d)}</div>
                    <div class="drawing-card-sub">Drawing</div>
                </div>`).join('');

            content.innerHTML = `
                <div class="stats-grid">
                    <div class="stat-card animate-in"><div class="stat-label">Total Drawings</div><div class="stat-value">${numbers.length}</div></div>
                </div>
                <div class="drawings-grid">${cards}</div>`;
        } catch (err) {
            content.innerHTML = `<div class="empty">Error: ${esc(err.message)}</div>`;
        }
    },

    async drawingDetail(drawingEncoded) {
        const drawingNumber = decodeURIComponent(drawingEncoded);
        const actions = document.getElementById('topbar-actions');
        const content = document.getElementById('page-content');
        setPage(drawingNumber, 'drawing');
        actions.innerHTML = '';
        content.innerHTML = `<div class="loading">Loading drawing</div>`;

        try {
            const exports = await api.get(`/api/exports/by-drawing?drawingNumber=${encodeURIComponent(drawingNumber)}`);

            if (exports.length === 0) {
                content.innerHTML = `
                    <a class="back-link" onclick="router.go('drawings')">${icons.back} Back to drawings</a>
                    <div class="empty">No exports found for this drawing.</div>`;
                return;
            }

            const allBom = [];
            exports.forEach(exp => {
                (exp.bomItems || []).forEach(b => {
                    allBom.push({ ...b, exportId: exp.id, exportedAt: exp.exportedAt });
                });
            });

            const bomRows = allBom.map((b, i) => {
                const hasDetails = b.cutTemplate || b.formProgram;
                const toggleId = `dbom-${b.id}`;
                return `
                    <tr class="${hasDetails ? 'clickable' : ''}" ${hasDetails ? `onclick="toggleBomRow('${toggleId}')"` : ''} style="animation: fadeSlideIn 0.25s ease ${0.03 * i}s forwards; opacity: 0">
                        <td style="width:32px">${hasDetails ? `<span class="chevron-toggle" id="${toggleId}-icon">${icons.chevron}</span>` : ''}</td>
                        <td style="font-family:var(--font-mono);font-weight:600;color:var(--cyan)">${esc(b.itemNo)}</td>
                        <td><strong>${esc(b.partName)}</strong></td>
                        <td style="color:var(--text-secondary)">${esc(b.description)}</td>
                        <td><span style="font-family:var(--font-mono);font-size:13px">${esc(b.material)}</span></td>
                        <td style="font-family:var(--font-mono);text-align:center">${b.qty ?? ''}</td>
                        <td style="font-family:var(--font-mono);text-align:center">${b.totalQty ?? ''}</td>
                        <td>
                            ${b.cutTemplate ? `<span class="badge badge-cyan">${icons.laser} DXF</span>` : ''}
                            ${b.formProgram ? `<span class="badge badge-amber">${icons.bend} Form</span>` : ''}
                        </td>
                    </tr>
                    ${hasDetails ? `<tr class="bom-expand-row" id="${toggleId}" style="display:none"><td colspan="8">${renderBomDetails(b)}</td></tr>` : ''}`;
            }).join('');

            content.innerHTML = `
                <a class="back-link" onclick="router.go('drawings')">${icons.back} Back to drawings</a>

                <div class="stats-grid">
                    <div class="stat-card animate-in"><div class="stat-label">Exports</div><div class="stat-value">${exports.length}</div></div>
                    <div class="stat-card animate-in"><div class="stat-label">BOM Items</div><div class="stat-value">${allBom.length}</div></div>
                    <div class="stat-card animate-in"><div class="stat-label">Latest Export</div><div class="stat-value stat-sm">${fmtDate(exports[0].exportedAt)}</div></div>
                </div>

                <div class="card animate-in">
                    <div class="card-header">
                        All BOM Items
                        <span class="badge badge-count">${allBom.length} items</span>
                    </div>
                    ${allBom.length ? `
                    <table>
                        <thead><tr>
                            <th style="width:32px"></th>
                            <th style="width:60px">Item</th>
                            <th>Part Name</th>
                            <th>Description</th>
                            <th>Material</th>
                            <th style="width:50px;text-align:center">Qty</th>
                            <th style="width:55px;text-align:center">Total</th>
                            <th style="width:120px">Data</th>
                        </tr></thead>
                        <tbody>${bomRows}</tbody>
                    </table>` : '<div class="empty">No BOM items.</div>'}
                </div>`;
        } catch (err) {
            content.innerHTML = `<div class="empty">Error: ${esc(err.message)}</div>`;
        }
    },

    async files(params) {
        const actions = document.getElementById('topbar-actions');
        const content = document.getElementById('page-content');
        setPage('Files');

        const searchVal = params.q || '';
        actions.innerHTML = `
            <div style="display:flex;gap:8px;align-items:center">
                <div class="search-box">
                    ${icons.search}
                    <input type="text" id="file-search" placeholder="Search drawing number, filename..." value="${esc(searchVal)}">
                </div>
                <select id="file-type-filter" style="background:var(--surface);border:1px solid var(--border);border-radius:4px;padding:6px 10px;color:var(--text);font-family:var(--font-body);font-size:14px;height:36px">
                    <option value="">All types</option>
                    <option value="dxf">DXF only</option>
                    <option value="pdf">PDF only</option>
                </select>
            </div>`;

        content.innerHTML = `<div class="loading">Loading files</div>`;

        const searchInput = document.getElementById('file-search');
        const typeFilter = document.getElementById('file-type-filter');
        let debounce;
        const refresh = () => {
            clearTimeout(debounce);
            debounce = setTimeout(() => router.go('files', { q: searchInput.value + (typeFilter.value ? '&type=' + typeFilter.value : '') }), 400);
        };
        searchInput.addEventListener('input', refresh);
        typeFilter.addEventListener('change', refresh);

        // Parse search and type from combined param
        let searchQ = searchVal;
        let typeQ = '';
        if (searchVal.includes('&type=')) {
            const parts = searchVal.split('&type=');
            searchQ = parts[0];
            typeQ = parts[1] || '';
            searchInput.value = searchQ;
            typeFilter.value = typeQ;
        }

        try {
            let url = '/api/filebrowser/files?';
            if (searchQ) url += `search=${encodeURIComponent(searchQ)}&`;
            if (typeQ) url += `type=${encodeURIComponent(typeQ)}&`;
            const data = await api.get(url);

            setPage('Files', `${data.total} files`);

            if (data.files.length === 0) {
                content.innerHTML = `<div class="empty">No files found.</div>`;
                return;
            }

            const rows = data.files.map((f, i) => {
                const ext = f.fileType || f.fileName.split('.').pop().toLowerCase();
                const hashShort = f.contentHash ? f.contentHash.substring(0, 12) : '';
                return `
                    <tr style="animation: fadeSlideIn 0.25s ease ${0.02 * i}s forwards; opacity: 0">
                        <td><div class="file-name-cell">${ext === 'pdf' ? icons.filePdf : icons.fileDxf}<a href="/api/filebrowser/download?hash=${encodeURIComponent(f.contentHash)}&ext=${ext}&name=${encodeURIComponent(f.fileName)}">${esc(f.fileName)}</a></div></td>
                        <td><span class="badge ${ext === 'dxf' ? 'badge-cyan' : 'badge-amber'}">${ext.toUpperCase()}</span></td>
                        <td style="color:var(--text-secondary)">${esc(f.drawingNumber)}</td>
                        <td style="font-family:var(--font-mono);font-size:13px;color:var(--text-secondary)">${f.thickness != null ? f.thickness.toFixed(4) + '"' : '\u2014'}</td>
                        <td style="font-family:var(--font-mono);font-size:13px;color:var(--text-secondary)">${fmtDate(f.createdAt)}</td>
                        <td style="font-family:var(--font-mono);font-size:12px;color:var(--text-dim)">${esc(hashShort)}</td>
                        <td style="white-space:nowrap">
                            <a class="btn btn-cyan btn-sm" href="/api/filebrowser/download?hash=${encodeURIComponent(f.contentHash)}&ext=${ext}&name=${encodeURIComponent(f.fileName)}">${icons.download}</a>
                        </td>
                    </tr>`;
            }).join('');

            content.innerHTML = `
                <div class="card animate-in">
                    <table>
                        <thead><tr>
                            <th>Name</th>
                            <th style="width:60px">Type</th>
                            <th>Drawing</th>
                            <th style="width:90px">Thickness</th>
                            <th style="width:170px">Date</th>
                            <th style="width:100px">Hash</th>
                            <th style="width:90px">Actions</th>
                        </tr></thead>
                        <tbody>${rows}</tbody>
                    </table>
                </div>`;
        } catch (err) {
            content.innerHTML = `<div class="empty">Error: ${esc(err.message)}</div>`;
        }
    }
};
