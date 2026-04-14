function fmtSize(b) {
    if (!b) return '0 B';
    const k = 1024, s = ['B','KB','MB','GB'];
    const i = Math.floor(Math.log(b) / Math.log(k));
    return parseFloat((b / Math.pow(k, i)).toFixed(1)) + ' ' + s[i];
}

function fmtDate(d) {
    if (!d) return '';
    const dt = new Date(d);
    return dt.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }) +
        ' ' + dt.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
}

function fmtThickness(t) {
    if (t == null) return '\u2014';
    return `<span style="font-family:var(--font-mono)">${t.toFixed(4)}"</span>`;
}

function esc(s) {
    return s ? s.replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;') : '';
}

function setPage(title, tag = '') {
    document.getElementById('page-title').textContent = title;
    document.getElementById('page-tag').textContent = tag;
    document.getElementById('page-tag').style.display = tag ? '' : 'none';
}

const api = {
    async get(url) {
        const r = await fetch(url);
        if (!r.ok) throw new Error(`${r.status} ${r.statusText}`);
        return r.json();
    },
    async del(url) {
        const r = await fetch(url, { method: 'DELETE' });
        if (!r.ok) throw new Error(`${r.status} ${r.statusText}`);
    }
};

async function deleteExport(id) {
    if (!confirm('Delete this export record? This cannot be undone.')) return;
    try {
        await api.del(`/api/exports/${id}`);
        router.dispatch();
    } catch (err) {
        alert('Failed to delete: ' + err.message);
    }
}
