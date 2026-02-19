const icons = {
    search: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>`,
    folder: `<svg viewBox="0 0 24 24" fill="var(--amber-dim)" stroke="var(--amber)" stroke-width="1.5"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/></svg>`,
    fileDxf: `<svg viewBox="0 0 24 24" fill="none" stroke="var(--cyan)" stroke-width="1.5"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>`,
    filePdf: `<svg viewBox="0 0 24 24" fill="none" stroke="var(--red)" stroke-width="1.5"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>`,
    fileGeneric: `<svg viewBox="0 0 24 24" fill="none" stroke="var(--text-dim)" stroke-width="1.5"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>`,
    download: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>`,
    back: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="19" y1="12" x2="5" y2="12"/><polyline points="12 19 5 12 12 5"/></svg>`,
    chevron: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="9 18 15 12 9 6"/></svg>`,
    laser: `<svg viewBox="0 0 16 16" fill="none" stroke="var(--cyan)" stroke-width="1.2"><circle cx="8" cy="8" r="2"/><path d="M8 2v3M8 11v3M2 8h3M11 8h3" opacity="0.5"/></svg>`,
    bend: `<svg viewBox="0 0 16 16" fill="none" stroke="var(--amber)" stroke-width="1.2"><path d="M3 13V7a4 4 0 0 1 4-4h6"/><polyline points="10 6 13 3 10 0" transform="translate(0,2)"/></svg>`,
};

function fileIcon(name) {
    const ext = name.split('.').pop().toLowerCase();
    if (ext === 'dxf') return icons.fileDxf;
    if (ext === 'pdf') return icons.filePdf;
    return icons.fileGeneric;
}
