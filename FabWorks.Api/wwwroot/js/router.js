const router = {
    go(page, params = {}) {
        const hash = page + (params.id ? '/' + params.id : '') + (params.q ? '?q=' + encodeURIComponent(params.q) : '');
        location.hash = hash;
    },
    parse() {
        const h = location.hash.slice(1) || 'exports';
        const [path, qs] = h.split('?');
        const parts = path.split('/');
        const params = {};
        if (qs) qs.split('&').forEach(p => { const [k,v] = p.split('='); params[k] = decodeURIComponent(v); });
        return { page: parts[0], id: parts[1], params };
    },
    init() {
        window.addEventListener('hashchange', () => this.dispatch());
        this.dispatch();
    },
    dispatch() {
        const { page, id, params } = this.parse();
        document.querySelectorAll('.nav-item').forEach(el => {
            el.classList.toggle('active',
                el.dataset.page === page ||
                (page === 'export-detail' && el.dataset.page === 'exports') ||
                (page === 'drawing-detail' && el.dataset.page === 'drawings'));
        });
        switch(page) {
            case 'exports': pages.exports(params); break;
            case 'export-detail': pages.exportDetail(id); break;
            case 'drawings': pages.drawings(params); break;
            case 'drawing-detail': pages.drawingDetail(id, params); break;
            case 'files': pages.files(params); break;
            default: pages.exports(params);
        }
    }
};
