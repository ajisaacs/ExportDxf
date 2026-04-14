/* ─── BOM Detail Expansion ─── */
function renderBomDetails(b) {
    let html = '<div class="bom-expand-content">';

    if (b.cutTemplate) {
        const ct = b.cutTemplate;
        const displayName = ct.dxfFilePath?.split(/[/\\]/).pop() || '';
        html += `
            <div class="bom-section-title">${icons.laser} Cut Template</div>
            <div class="info-grid">
                <div class="info-item"><span class="lbl">File</span><span class="val">${esc(displayName)}</span></div>
                <div class="info-item"><span class="lbl">Thickness</span><span class="val">${fmtThickness(ct.thickness)}</span></div>
                <div class="info-item"><span class="lbl">K-Factor</span><span class="val">${ct.kFactor != null ? ct.kFactor : '\u2014'}</span></div>
                <div class="info-item"><span class="lbl">Bend Radius</span><span class="val">${ct.defaultBendRadius != null ? ct.defaultBendRadius.toFixed(4) + '"' : '\u2014'}</span></div>
            </div>`;

        if (ct.contentHash) {
            html += `<div style="margin-top:10px">
                <a class="btn btn-cyan btn-sm" href="/api/files/blob/${encodeURIComponent(ct.contentHash)}?ext=dxf&download=true&name=${encodeURIComponent(displayName)}" onclick="event.stopPropagation()">${icons.download} Download DXF</a>
                <span style="font-family:var(--font-mono);font-size:13px;color:var(--text-dim);margin-left:8px">${esc(displayName)}</span>
            </div>`;
        }
    }

    if (b.formProgram) {
        const fp = b.formProgram;
        html += `
            <div class="bom-section-title">${icons.bend} Form Program</div>
            <div class="info-grid">
                <div class="info-item"><span class="lbl">Program</span><span class="val">${esc(fp.programName)}</span></div>
                <div class="info-item"><span class="lbl">Thickness</span><span class="val">${fmtThickness(fp.thickness)}</span></div>
                <div class="info-item"><span class="lbl">Material</span><span class="val">${esc(fp.materialType)}</span></div>
                <div class="info-item"><span class="lbl">K-Factor</span><span class="val">${fp.kFactor != null ? fp.kFactor : '\u2014'}</span></div>
                <div class="info-item"><span class="lbl">Bends</span><span class="val">${fp.bendCount}</span></div>
                <div class="info-item"><span class="lbl">Upper Tools</span><span class="val">${esc(fp.upperToolNames) || '\u2014'}</span></div>
                <div class="info-item"><span class="lbl">Lower Tools</span><span class="val">${esc(fp.lowerToolNames) || '\u2014'}</span></div>
            </div>
            ${fp.setupNotes ? `<div style="margin-top:8px;padding:8px 12px;background:var(--amber-dim);border-radius:4px;font-size:13px;font-family:var(--font-mono);color:var(--amber)"><span class="lbl">Setup Notes</span>${esc(fp.setupNotes)}</div>` : ''}`;
    }

    html += '</div>';
    return html;
}

function toggleEquipGroup(id) {
    const group = document.getElementById(id);
    const icon = document.getElementById(id + '-icon');
    if (!group) return;
    group.classList.toggle('collapsed');
    if (icon) icon.classList.toggle('open', !group.classList.contains('collapsed'));
}

function toggleBomRow(id) {
    const row = document.getElementById(id);
    const icon = document.getElementById(id + '-icon');
    if (!row) return;
    const visible = row.style.display !== 'none';
    row.style.display = visible ? 'none' : '';
    if (icon) icon.classList.toggle('open', !visible);
}
