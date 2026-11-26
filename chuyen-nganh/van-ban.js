/* =====================================================================
 * TAB "VĂN BẢN" — Thư viện văn bản pháp quy KTCN 2026 & Đính kèm tài liệu
 * - Văn bản hệ thống: khai báo trong SYSTEM_DOCS (file nằm cùng thư mục web).
 * - Văn bản đính kèm của người dùng: lưu trong IndexedDB của trình duyệt
 *   (không mất khi tải lại trang, chỉ hiển thị trên máy/trình duyệt đã đính kèm).
 * ===================================================================== */
(function () {
    'use strict';

    // ---------- 1. DANH MỤC VĂN BẢN HỆ THỐNG (4 VĂN BẢN WORD GỌN NHẸ) ----------
    const SYSTEM_DOCS = [
        { id: 'tt33',  so: '33/2026/TT-BCT',  loai: 'Thông tư', bo: 'BCT',   title: 'Danh mục hàng hóa kiểm tra chuyên ngành thuộc Bộ Công Thương', file: '33_2026_TT-BCT_710258.doc', size: 371896 },
        { id: 'tt36',  so: '36/2026/TT-BKHCN', loai: 'Thông tư', bo: 'BKHCN', title: 'Danh mục hàng hóa kiểm tra chuyên ngành thuộc Bộ Khoa học và Công nghệ', file: '36_2026_TT-BKHCN_710259.doc', size: 891225 },
        { id: 'tt41',  so: '41/2026/TT-BXD',  loai: 'Thông tư', bo: 'BXD',   title: 'Danh mục hàng hóa kiểm tra chuyên ngành thuộc Bộ Xây dựng', file: '41_2026_TT-BXD_697288.doc', size: 228123 },
        { id: 'tt16',  so: '16/2026/TT-BNV',  loai: 'Thông tư', bo: 'BNV',   title: 'Danh mục máy, thiết bị, vật tư có yêu cầu nghiêm ngặt về ATVSLĐ (Bộ Nội vụ)', file: '16_2026_TT-BNV_716907.doc', size: 70502 }
    ];

    const BO = {
        ALL:   { label: 'Tất cả' },
        CP:    { label: 'Chính phủ',  pill: 'text-rose-700 bg-rose-50 border-rose-200 dark:text-rose-300 dark:bg-rose-950/40 dark:border-rose-900/60' },
        BCT:   { label: 'Bộ Công Thương', pill: 'text-emerald-700 bg-emerald-50 border-emerald-200 dark:text-emerald-300 dark:bg-emerald-950/40 dark:border-emerald-900/60' },
        BKHCN: { label: 'Bộ KH&CN',   pill: 'text-blue-700 bg-blue-50 border-blue-200 dark:text-blue-300 dark:bg-blue-950/40 dark:border-blue-900/60' },
        BXD:   { label: 'Bộ Xây dựng', pill: 'text-orange-700 bg-orange-50 border-orange-200 dark:text-orange-300 dark:bg-orange-950/40 dark:border-orange-900/60' },
        BNV:   { label: 'Bộ Nội vụ',  pill: 'text-purple-700 bg-purple-50 border-purple-200 dark:text-purple-300 dark:bg-purple-950/40 dark:border-purple-900/60' },
        KHAC:  { label: 'Khác',       pill: 'text-slate-700 bg-slate-50 border-slate-200 dark:text-slate-300 dark:bg-slate-800/60 dark:border-slate-700' }
    };

    // ---------- 2. TIỆN ÍCH ----------
    const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
    const noTone = s => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase();
    const fmtSize = b => !b ? '' : b >= 1048576 ? (b / 1048576).toFixed(1) + ' MB' : Math.max(1, Math.round(b / 1024)) + ' KB';
    const extOf = name => (String(name).split('.').pop() || '').toLowerCase();
    const fmtDate = t => { try { return new Date(t).toLocaleDateString('vi-VN'); } catch (e) { return ''; } };

    function fileIcon(ext) {
        if (ext === 'pdf') return { i: 'fa-file-pdf', c: 'text-rose-600 bg-rose-50 dark:bg-rose-950/40' };
        if (ext === 'doc' || ext === 'docx') return { i: 'fa-file-word', c: 'text-blue-600 bg-blue-50 dark:bg-blue-950/40' };
        if (ext === 'xls' || ext === 'xlsx' || ext === 'csv') return { i: 'fa-file-excel', c: 'text-emerald-600 bg-emerald-50 dark:bg-emerald-950/40' };
        if (['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp'].includes(ext)) return { i: 'fa-file-image', c: 'text-amber-600 bg-amber-50 dark:bg-amber-950/40' };
        if (ext === 'zip' || ext === 'rar' || ext === '7z') return { i: 'fa-file-zipper', c: 'text-slate-600 bg-slate-100 dark:bg-slate-800' };
        return { i: 'fa-file-lines', c: 'text-slate-600 bg-slate-100 dark:bg-slate-800' };
    }
    const canPreview = ext => ['pdf', 'png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'txt'].includes(ext);

    function guessBo(name) {
        const n = noTone(name).toUpperCase();
        if (/BKHCN/.test(n)) return 'BKHCN';
        if (/BCT/.test(n)) return 'BCT';
        if (/BXD/.test(n)) return 'BXD';
        if (/BNV/.test(n)) return 'BNV';
        if (/ND-CP|ND CP|NGHI DINH|\bND\b/.test(n)) return 'CP';
        return 'KHAC';
    }
    function guessSo(name) {
        const m = String(name).match(/(\d{1,4})\s*[\/_.\- ]\s*(20\d{2})(?:\s*[\/_\- ]\s*((?:TT|NĐ|ND|QĐ|QD|VBHN)[\-_ ]?[A-ZĐ]*))?/i);
        if (!m) return '';
        return m[1] + '/' + m[2] + (m[3] ? '/' + m[3].replace(/[_ ]/g, '-').toUpperCase() : '');
    }

    // ---------- 3. INDEXEDDB (lưu file đính kèm) ----------
    const DB_NAME = 'ktcn_docs_db', STORE = 'files';
    let dbPromise = null;
    function openDB() {
        if (!dbPromise) dbPromise = new Promise((res, rej) => {
            const r = indexedDB.open(DB_NAME, 1);
            r.onupgradeneeded = () => r.result.createObjectStore(STORE, { keyPath: 'id' });
            r.onsuccess = () => res(r.result);
            r.onerror = () => rej(r.error);
        });
        return dbPromise;
    }
    async function dbRun(mode, fn) {
        const db = await openDB();
        return new Promise((res, rej) => {
            const t = db.transaction(STORE, mode);
            const req = fn(t.objectStore(STORE));
            let out;
            if (req) req.onsuccess = () => { out = req.result; };
            t.oncomplete = () => res(out);
            t.onerror = () => rej(t.error);
            t.onabort = () => rej(t.error);
        });
    }
    const dbAll = () => dbRun('readonly', s => s.getAll());
    const dbPut = rec => dbRun('readwrite', s => s.put(rec));
    const dbDel = id => dbRun('readwrite', s => s.delete(id));
    const dbGet = id => dbRun('readonly', s => s.get(id));

    // ---------- 4. TRẠNG THÁI ----------
    let docsMode = false;
    let userDocs = [];
    let query = '';
    let boFilter = 'ALL';
    const availability = {}; // id -> true/false (văn bản hệ thống có trên server hay không)
    const sysUrl = d => './' + encodeURIComponent(d.file);

    // ---------- 5. CSS & KHUNG GIAO DIỆN ----------
    const style = document.createElement('style');
    style.textContent = `
        body.docs-mode #filter-bar, body.docs-mode #layout-dual-view, body.docs-mode #layout-cards-view,
        body.docs-mode #layout-table-view, body.docs-mode #empty-state { display: none !important; }
        #layout-docs-view .doc-card { transition: transform .15s ease, box-shadow .15s ease; }
        #layout-docs-view .doc-card:hover { transform: translateY(-2px); box-shadow: 0 10px 24px -12px rgba(76, 29, 149, .35); }
        #docs-dropzone.drag-over { border-color: #7c3aed; background: rgba(124, 58, 237, .08); }
        .docs-scroll::-webkit-scrollbar { width: 5px; }
    `;
    document.head.appendChild(style);

    function ensureContainer() {
        let el = document.getElementById('layout-docs-view');
        if (!el) {
            el = document.createElement('div');
            el.id = 'layout-docs-view';
            el.className = 'hidden';
            const main = document.querySelector('main');
            (main || document.body).appendChild(el);
        }
        return el;
    }

    function renderShell() {
        const el = ensureContainer();
        el.innerHTML = `
        <div class="flex flex-col gap-2.5">
            <!-- Thanh công cụ floating 1 hàng -->
            <div class="flex items-center gap-2 py-0.5 whitespace-nowrap">
                <div class="relative flex items-center w-52 sm:w-60 lg:w-64 shrink-0">
                    <i class="fa-solid fa-magnifying-glass absolute left-2.5 text-slate-400 text-xs pointer-events-none"></i>
                    <input id="docs-search" type="text" placeholder="Tìm số hiệu, tên văn bản..." class="w-full pl-7 pr-3 py-1.5 bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-700/80 rounded-lg text-slate-900 dark:text-white placeholder-slate-400 text-xs font-medium shadow-xs focus:outline-none focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500/20 transition">
                </div>
                <div id="docs-bo-pills" class="flex items-center gap-1 flex-1 min-w-0 overflow-x-auto no-scrollbar"></div>
                <div class="flex items-center gap-1.5 shrink-0">
                    <button id="docs-btn-attach" class="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold text-white shadow-xs transition hover:brightness-110" style="background: linear-gradient(135deg, #2b1565 0%, #5f3fc0 100%);">
                        <i class="fa-solid fa-paperclip"></i><span>Đính kèm văn bản</span>
                    </button>
                    <input id="docs-file-input" type="file" multiple class="hidden" accept=".pdf,.doc,.docx,.xls,.xlsx,.csv,.txt,.png,.jpg,.jpeg,.webp,.zip,.rar,.7z">
                </div>
            </div>

            <div class="bg-white/90 dark:bg-slate-900/85 backdrop-blur-xl rounded-2xl border border-white/80 dark:border-slate-800 shadow-xl shadow-slate-200/50 dark:shadow-none p-4 sm:p-5 docs-scroll overflow-y-auto" style="height: calc(100vh - 150px); min-height: 520px;">
                <!-- Vùng kéo thả -->
                <label id="docs-dropzone" for="docs-file-input" class="flex items-center justify-center gap-3 px-4 py-3 mb-4 rounded-xl border-2 border-dashed border-purple-300/70 dark:border-purple-800/70 text-slate-500 dark:text-slate-400 cursor-pointer hover:border-purple-500 hover:bg-purple-50/40 dark:hover:bg-purple-950/20 transition">
                    <span class="w-9 h-9 rounded-full bg-purple-100 dark:bg-purple-950/60 text-purple-600 dark:text-purple-300 flex items-center justify-center"><i class="fa-solid fa-cloud-arrow-up"></i></span>
                    <span class="text-xs leading-relaxed">
                        <b class="text-slate-700 dark:text-slate-200">Kéo thả file vào đây</b> hoặc bấm để chọn (PDF, Word, Excel, ảnh...)<br>
                        <span class="text-[11px] text-slate-400">File đính kèm được lưu an toàn trong trình duyệt của máy này, không tải lên máy chủ.</span>
                    </span>
                </label>

                <!-- Văn bản đính kèm của tôi -->
                <div class="flex items-center justify-between mb-2">
                    <h3 class="text-[11px] font-black uppercase tracking-wider text-slate-500 dark:text-slate-400"><i class="fa-solid fa-paperclip text-purple-500 mr-1.5"></i>Văn bản đính kèm của tôi (<span id="docs-user-count">0</span>)</h3>
                    <span id="docs-storage" class="text-[10px] text-slate-400"></span>
                </div>
                <div id="docs-user-grid" class="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-3 gap-2.5 mb-5"></div>

                <!-- Văn bản pháp quy hệ thống -->
                <div class="flex items-center justify-between mb-2">
                    <h3 class="text-[11px] font-black uppercase tracking-wider text-slate-500 dark:text-slate-400"><i class="fa-solid fa-scale-balanced text-indigo-500 mr-1.5"></i>Văn bản pháp quy KTCN 2026 (<span id="docs-sys-count">0</span>)</h3>
                </div>
                <div id="docs-sys-grid" class="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-3 gap-2.5"></div>
            </div>
        </div>

        <!-- Modal xem trước -->
        <div id="docs-preview" class="fixed inset-0 z-[60] hidden items-center justify-center bg-slate-950/70 backdrop-blur-sm p-3 sm:p-6">
            <div class="w-full max-w-6xl h-full bg-white dark:bg-slate-900 rounded-2xl shadow-2xl flex flex-col overflow-hidden border border-white/20">
                <div class="flex items-center justify-between gap-3 px-4 py-2.5 text-white shrink-0" style="background: linear-gradient(135deg, #2b1565 0%, #5f3fc0 100%);">
                    <div class="min-w-0"><div id="docs-preview-title" class="text-sm font-bold truncate"></div><div id="docs-preview-sub" class="text-[11px] text-purple-200 truncate"></div></div>
                    <div class="flex items-center gap-1.5 shrink-0">
                        <a id="docs-preview-open" target="_blank" rel="noopener" class="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg text-xs font-semibold bg-white/15 hover:bg-white/25 border border-white/25"><i class="fa-solid fa-up-right-from-square"></i><span class="hidden sm:inline">Tab mới</span></a>
                        <a id="docs-preview-dl" class="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg text-xs font-semibold bg-white/15 hover:bg-white/25 border border-white/25"><i class="fa-solid fa-download"></i><span class="hidden sm:inline">Tải về</span></a>
                        <button id="docs-preview-close" class="w-8 h-8 rounded-full bg-white/15 hover:bg-white/25 border border-white/25 flex items-center justify-center" title="Đóng (Esc)"><i class="fa-solid fa-xmark"></i></button>
                    </div>
                </div>
                <div id="docs-preview-body" class="flex-1 bg-slate-100 dark:bg-slate-950 overflow-auto"></div>
            </div>
        </div>

        <!-- Modal sửa thông tin -->
        <div id="docs-edit" class="fixed inset-0 z-[60] hidden items-center justify-center bg-slate-950/60 backdrop-blur-sm p-4">
            <form id="docs-edit-form" class="w-full max-w-md bg-white dark:bg-slate-900 rounded-2xl shadow-2xl border border-slate-200 dark:border-slate-700 p-5 space-y-3">
                <h4 class="text-sm font-black text-slate-800 dark:text-white"><i class="fa-solid fa-pen-to-square text-purple-500 mr-1.5"></i>Thông tin văn bản đính kèm</h4>
                <input type="hidden" name="docid">
                <label class="block text-[11px] font-bold text-slate-500 uppercase">Số hiệu<input name="so" class="mt-1 w-full px-2.5 py-1.5 rounded-lg border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-950 text-xs text-slate-800 dark:text-white font-medium normal-case" placeholder="VD: 36/2026/TT-BKHCN"></label>
                <label class="block text-[11px] font-bold text-slate-500 uppercase">Tên / trích yếu<input name="title" required class="mt-1 w-full px-2.5 py-1.5 rounded-lg border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-950 text-xs text-slate-800 dark:text-white font-medium normal-case"></label>
                <label class="block text-[11px] font-bold text-slate-500 uppercase">Cơ quan ban hành<select name="bo" class="mt-1 w-full px-2.5 py-1.5 rounded-lg border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-950 text-xs text-slate-800 dark:text-white font-medium normal-case">
                    ${Object.keys(BO).filter(k => k !== 'ALL').map(k => `<option value="${k}">${esc(BO[k].label)}</option>`).join('')}
                </select></label>
                <label class="block text-[11px] font-bold text-slate-500 uppercase">Ghi chú<textarea name="note" rows="2" class="mt-1 w-full px-2.5 py-1.5 rounded-lg border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-950 text-xs text-slate-800 dark:text-white font-medium normal-case" placeholder="VD: Hiệu lực từ 01/01/2027..."></textarea></label>
                <div class="flex justify-end gap-2 pt-1">
                    <button type="button" id="docs-edit-cancel" class="px-3 py-1.5 rounded-lg text-xs font-semibold text-slate-600 dark:text-slate-300 bg-slate-100 dark:bg-slate-800 hover:bg-slate-200">Hủy</button>
                    <button type="submit" class="px-3 py-1.5 rounded-lg text-xs font-bold text-white" style="background: linear-gradient(135deg, #2b1565 0%, #5f3fc0 100%);">Lưu</button>
                </div>
            </form>
        </div>`;

        // Sự kiện
        document.getElementById('docs-search').addEventListener('input', e => { query = e.target.value; renderLists(); });
        document.getElementById('docs-btn-attach').addEventListener('click', () => document.getElementById('docs-file-input').click());
        document.getElementById('docs-file-input').addEventListener('change', e => { addFiles(e.target.files); e.target.value = ''; });
        const dz = document.getElementById('docs-dropzone');
        ['dragenter', 'dragover'].forEach(ev => el.addEventListener(ev, e => { if (hasFiles(e)) { e.preventDefault(); dz.classList.add('drag-over'); } }));
        ['dragleave', 'drop'].forEach(ev => el.addEventListener(ev, e => { if (ev === 'dragleave' && el.contains(e.relatedTarget)) return; dz.classList.remove('drag-over'); }));
        el.addEventListener('drop', e => { if (hasFiles(e)) { e.preventDefault(); addFiles(e.dataTransfer.files); } });
        document.getElementById('docs-preview-close').addEventListener('click', closePreview);
        document.getElementById('docs-preview').addEventListener('click', e => { if (e.target.id === 'docs-preview') closePreview(); });
        document.getElementById('docs-edit-cancel').addEventListener('click', closeEdit);
        document.getElementById('docs-edit-form').addEventListener('submit', saveEdit);
        el.addEventListener('click', onGridClick);
        document.addEventListener('keydown', e => { if (e.key === 'Escape') { closePreview(); closeEdit(); } });
    }
    const hasFiles = e => e.dataTransfer && Array.from(e.dataTransfer.types || []).includes('Files');

    // ---------- 6. RENDER ----------
    function matches(d) {
        if (boFilter !== 'ALL' && d.bo !== boFilter) return false;
        if (!query.trim()) return true;
        const q = noTone(query.trim());
        return noTone([d.so, d.title, d.file || d.name, d.note, d.loai, (BO[d.bo] || {}).label].join(' ')).includes(q);
    }

    function renderPills() {
        const box = document.getElementById('docs-bo-pills');
        if (!box) return;
        const all = SYSTEM_DOCS.concat(userDocs);
        box.innerHTML = Object.keys(BO).map(k => {
            const n = k === 'ALL' ? all.length : all.filter(d => d.bo === k).length;
            if (k !== 'ALL' && !n) return '';
            const active = boFilter === k;
            const cls = active
                ? 'bg-slate-900 text-white dark:bg-white dark:text-slate-900 border-transparent font-bold'
                : (BO[k].pill || 'text-slate-700 bg-white border-slate-200 dark:text-slate-200 dark:bg-slate-900 dark:border-slate-700') + ' font-semibold hover:brightness-95';
            return `<button data-bo="${k}" class="docs-bo shrink-0 px-2.5 py-1.5 rounded-lg border text-xs shadow-xs transition ${cls}">${esc(BO[k].label)} (${n})</button>`;
        }).join('');
    }

    function cardHtml(d, isUser) {
        const name = isUser ? d.name : d.file;
        const ext = extOf(name);
        const ic = fileIcon(ext);
        const bo = BO[d.bo] || BO.KHAC;
        const missing = !isUser && availability[d.id] === false;
        const meta = [fmtSize(d.size), ext.toUpperCase(), isUser && d.added ? 'Thêm ' + fmtDate(d.added) : (d.loai || '')].filter(Boolean).join(' • ');
        const btn = 'inline-flex items-center gap-1 px-2 py-1 rounded-md text-[11px] font-semibold border transition';
        const actions = missing
            ? `<span class="text-[10px] font-semibold text-amber-600 dark:text-amber-400"><i class="fa-solid fa-triangle-exclamation mr-1"></i>Chưa đăng tải lên web</span>`
            : `${canPreview(ext) ? `<button data-act="view" data-id="${esc(d.id)}" data-user="${isUser ? 1 : 0}" class="${btn} text-indigo-700 dark:text-indigo-300 bg-indigo-50 dark:bg-indigo-950/40 border-indigo-200 dark:border-indigo-900/60 hover:bg-indigo-100"><i class="fa-solid fa-eye"></i>Xem</button>` : ''}
               <button data-act="dl" data-id="${esc(d.id)}" data-user="${isUser ? 1 : 0}" class="${btn} text-slate-700 dark:text-slate-200 bg-white dark:bg-slate-800 border-slate-200 dark:border-slate-700 hover:bg-slate-50"><i class="fa-solid fa-download"></i>Tải</button>
               ${isUser ? `<button data-act="edit" data-id="${esc(d.id)}" class="${btn} text-slate-600 dark:text-slate-300 bg-white dark:bg-slate-800 border-slate-200 dark:border-slate-700 hover:bg-slate-50" title="Sửa thông tin"><i class="fa-solid fa-pen"></i></button>
               <button data-act="del" data-id="${esc(d.id)}" class="${btn} text-rose-600 dark:text-rose-400 bg-white dark:bg-slate-800 border-rose-200 dark:border-rose-900/60 hover:bg-rose-50" title="Xóa"><i class="fa-solid fa-trash-can"></i></button>` : ''}`;
        return `
        <div class="doc-card group relative flex gap-3 p-3 rounded-xl bg-white/95 dark:bg-slate-900/90 border border-slate-200/90 dark:border-slate-800 shadow-xs ${missing ? 'opacity-70' : ''}">
            <div class="w-10 h-10 rounded-lg ${ic.c} flex items-center justify-center shrink-0"><i class="fa-solid ${ic.i} text-lg"></i></div>
            <div class="min-w-0 flex-1">
                <div class="flex items-center gap-1.5 flex-wrap">
                    ${d.so ? `<span class="px-1.5 py-0.5 rounded font-mono text-[10px] font-bold text-indigo-700 dark:text-indigo-300 bg-indigo-50 dark:bg-indigo-950/50 border border-indigo-100 dark:border-indigo-900/60">${esc(d.so)}</span>` : ''}
                    <span class="px-1.5 py-0.5 rounded border text-[10px] font-semibold ${bo.pill}">${esc(bo.label)}</span>
                    ${isUser ? '<span class="px-1.5 py-0.5 rounded text-[10px] font-bold text-purple-700 dark:text-purple-300 bg-purple-50 dark:bg-purple-950/50"><i class="fa-solid fa-paperclip mr-0.5"></i>Đính kèm</span>' : ''}
                </div>
                <div class="mt-1 text-xs font-bold text-slate-800 dark:text-slate-100 leading-snug line-clamp-2" title="${esc(name)}">${esc(d.title)}</div>
                ${d.note ? `<div class="mt-0.5 text-[11px] text-slate-500 dark:text-slate-400 line-clamp-2">${esc(d.note)}</div>` : ''}
                <div class="mt-1 text-[10px] text-slate-400 truncate">${esc(meta)}</div>
                <div class="mt-2 flex items-center gap-1.5 flex-wrap">${actions}</div>
            </div>
        </div>`;
    }

    function emptyHtml(text) {
        return `<div class="col-span-full text-center text-xs text-slate-400 py-5 rounded-xl border border-dashed border-slate-200 dark:border-slate-800">${text}</div>`;
    }

    function renderLists() {
        if (!document.getElementById('docs-user-grid')) return;
        renderPills();
        const u = userDocs.filter(matches).sort((a, b) => b.added - a.added);
        const s = SYSTEM_DOCS.filter(matches);
        document.getElementById('docs-user-count').textContent = userDocs.length;
        document.getElementById('docs-sys-count').textContent = SYSTEM_DOCS.length;
        document.getElementById('docs-user-grid').innerHTML = u.length ? u.map(d => cardHtml(d, true)).join('')
            : emptyHtml(userDocs.length ? 'Không có văn bản đính kèm phù hợp bộ lọc.' : 'Chưa có văn bản đính kèm. Bấm <b>Đính kèm văn bản</b> hoặc kéo thả file vào khung phía trên.');
        document.getElementById('docs-sys-grid').innerHTML = s.length ? s.map(d => cardHtml(d, false)).join('') : emptyHtml('Không có văn bản phù hợp.');
        updateBadge();
    }

    function updateBadge() {
        const b = document.getElementById('docs-tab-count');
        if (b) b.textContent = SYSTEM_DOCS.length + userDocs.length;
    }

    async function updateStorage() {
        const el = document.getElementById('docs-storage');
        if (!el || !navigator.storage || !navigator.storage.estimate) return;
        try {
            const { usage, quota } = await navigator.storage.estimate();
            el.textContent = `Dung lượng đã dùng: ${fmtSize(usage) || '0 KB'} / ${fmtSize(quota)}`;
        } catch (e) { /* bỏ qua */ }
    }

    // ---------- 7. THAO TÁC ----------
    async function loadUserDocs() {
        try { userDocs = (await dbAll()) || []; } catch (e) { userDocs = []; console.warn('IndexedDB lỗi:', e); }
        renderLists(); updateStorage();
    }

    async function addFiles(fileList) {
        const files = Array.from(fileList || []);
        if (!files.length) return;
        const big = files.filter(f => f.size > 200 * 1048576);
        if (big.length && !confirm(`Có ${big.length} file lớn hơn 200 MB. Vẫn tiếp tục đính kèm?`)) return;
        if (navigator.storage && navigator.storage.persist) { try { await navigator.storage.persist(); } catch (e) { } }
        let ok = 0;
        for (const f of files) {
            const base = f.name.replace(/\.[^.]+$/, '');
            const rec = {
                id: 'u' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7),
                name: f.name, title: base.replace(/[_]+/g, ' ').trim(), so: guessSo(base), bo: guessBo(base),
                note: '', type: f.type, size: f.size, added: Date.now(), blob: f
            };
            try { await dbPut(rec); ok++; } catch (e) { alert('Không lưu được file "' + f.name + '": ' + (e && e.message || e)); }
        }
        await loadUserDocs();
        if (ok) toast(`Đã đính kèm ${ok} văn bản.`);
    }

    function findDoc(id, isUser) { return isUser ? userDocs.find(d => d.id === id) : SYSTEM_DOCS.find(d => d.id === id); }

    let currentBlobUrl = null;
    function urlFor(d, isUser) {
        if (!isUser) return sysUrl(d);
        if (currentBlobUrl) URL.revokeObjectURL(currentBlobUrl);
        const blob = d.blob instanceof Blob && !d.blob.type && extOf(d.name) === 'pdf' ? new Blob([d.blob], { type: 'application/pdf' }) : d.blob;
        currentBlobUrl = URL.createObjectURL(blob);
        return currentBlobUrl;
    }

    function download(d, isUser) {
        const a = document.createElement('a');
        a.href = urlFor(d, isUser);
        a.download = isUser ? d.name : d.file;
        document.body.appendChild(a); a.click(); a.remove();
    }

    function openPreview(d, isUser) {
        const name = isUser ? d.name : d.file;
        const ext = extOf(name);
        const url = urlFor(d, isUser);
        document.getElementById('docs-preview-title').textContent = (d.so ? d.so + ' — ' : '') + d.title;
        document.getElementById('docs-preview-sub').textContent = name + (d.size ? ' • ' + fmtSize(d.size) : '');
        const open = document.getElementById('docs-preview-open'); open.href = url;
        const dl = document.getElementById('docs-preview-dl'); dl.href = url; dl.download = name;
        const body = document.getElementById('docs-preview-body');
        body.innerHTML = ['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp'].includes(ext)
            ? `<div class="min-h-full flex items-center justify-center p-4"><img src="${url}" alt="" class="max-w-full shadow-lg rounded"></div>`
            : `<iframe src="${url}" class="w-full h-full border-0 bg-white" title="Xem trước văn bản"></iframe>`;
        const m = document.getElementById('docs-preview'); m.classList.remove('hidden'); m.classList.add('flex');
    }
    function closePreview() {
        const m = document.getElementById('docs-preview');
        if (!m || m.classList.contains('hidden')) return;
        m.classList.add('hidden'); m.classList.remove('flex');
        document.getElementById('docs-preview-body').innerHTML = '';
    }

    function openEdit(d) {
        const f = document.getElementById('docs-edit-form').elements;
        f.docid.value = d.id; f.so.value = d.so || ''; f.title.value = d.title || ''; f.bo.value = d.bo || 'KHAC'; f.note.value = d.note || '';
        const m = document.getElementById('docs-edit'); m.classList.remove('hidden'); m.classList.add('flex');
        setTimeout(() => f.title.focus(), 30);
    }
    function closeEdit() { const m = document.getElementById('docs-edit'); if (m) { m.classList.add('hidden'); m.classList.remove('flex'); } }
    async function saveEdit(e) {
        e.preventDefault();
        const f = e.target.elements;
        const rec = await dbGet(f.docid.value);
        if (!rec) return closeEdit();
        Object.assign(rec, { so: f.so.value.trim(), title: f.title.value.trim() || rec.name, bo: f.bo.value, note: f.note.value.trim() });
        await dbPut(rec);
        closeEdit(); await loadUserDocs(); toast('Đã lưu thông tin văn bản.');
    }

    async function onGridClick(e) {
        const pill = e.target.closest('.docs-bo');
        if (pill) { boFilter = pill.dataset.bo; renderLists(); return; }
        const b = e.target.closest('[data-act]');
        if (!b) return;
        const isUser = b.dataset.user === '1' || ['edit', 'del'].includes(b.dataset.act);
        const d = findDoc(b.dataset.id, isUser);
        if (!d) return;
        if (b.dataset.act === 'view') openPreview(d, isUser);
        else if (b.dataset.act === 'dl') download(d, isUser);
        else if (b.dataset.act === 'edit') openEdit(d);
        else if (b.dataset.act === 'del') {
            if (!confirm(`Xóa văn bản đính kèm "${d.title}"?`)) return;
            await dbDel(d.id); await loadUserDocs(); toast('Đã xóa văn bản.');
        }
    }

    function toast(msg) {
        const t = document.createElement('div');
        t.className = 'fixed left-1/2 -translate-x-1/2 bottom-6 z-[70] px-4 py-2 rounded-full text-xs font-semibold text-white shadow-lg';
        t.style.background = 'linear-gradient(135deg, #2b1565 0%, #5f3fc0 100%)';
        t.textContent = msg;
        document.body.appendChild(t);
        setTimeout(() => { t.style.transition = 'opacity .3s'; t.style.opacity = '0'; setTimeout(() => t.remove(), 300); }, 1800);
    }

    // Kiểm tra văn bản hệ thống có trên máy chủ không (GitHub Pages có thể chưa đăng tải file nặng)
    async function checkAvailability() {
        if (location.protocol === 'file:') return;
        await Promise.all(SYSTEM_DOCS.map(async d => {
            try { const r = await fetch(sysUrl(d), { method: 'HEAD', cache: 'no-store' }); availability[d.id] = r.ok; }
            catch (e) { availability[d.id] = false; }
        }));
        renderLists();
    }

    // ---------- 8. CHUYỂN TAB ----------
    const BTN_IDLE = 'inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold bg-white/15 hover:bg-white/25 border border-white/25 text-white shadow-xs transition';
    const BTN_ACTIVE = 'inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold bg-white text-indigo-700 border border-white shadow-xs transition';
    const LAYOUT_IDLE = 'w-8 h-8 rounded-lg text-white/80 hover:text-white hover:bg-white/10 flex items-center justify-center transition';
    let savedLayoutClasses = null;

    function setDocsMode(on) {
        docsMode = !!on;
        document.body.classList.toggle('docs-mode', docsMode);
        ensureContainer().classList.toggle('hidden', !docsMode);
        const btn = document.getElementById('btn-docs-tab');
        if (btn) btn.className = docsMode ? BTN_ACTIVE : BTN_IDLE;
        const ids = ['btn-layout-dual', 'btn-layout-cards', 'btn-layout-table'];
        if (docsMode) {
            savedLayoutClasses = ids.map(id => { const b = document.getElementById(id); return b ? b.className : ''; });
            ids.forEach(id => { const b = document.getElementById(id); if (b) b.className = LAYOUT_IDLE; });
            if (!document.getElementById('docs-user-grid')) { renderShell(); loadUserDocs(); checkAvailability(); }
            else renderLists();
            try { history.replaceState(null, '', '#van-ban'); } catch (e) { }
        } else {
            if (savedLayoutClasses) ids.forEach((id, i) => { const b = document.getElementById(id); if (b && savedLayoutClasses[i]) b.className = savedLayoutClasses[i]; });
            if (location.hash === '#van-ban') { try { history.replaceState(null, '', location.pathname + location.search); } catch (e) { } }
        }
        window.scrollTo({ top: 0 });
    }

    window.toggleDocsView = () => setDocsMode(!docsMode);
    window.showDocsView = () => setDocsMode(true);

    // Khi bấm chuyển layout (2 cột / thẻ / bảng) thì tự thoát tab Văn bản
    if (typeof window.setLayout === 'function') {
        const origSetLayout = window.setLayout;
        window.setLayout = function (layout) {
            if (docsMode) { savedLayoutClasses = null; setDocsMode(false); }
            return origSetLayout.apply(this, arguments);
        };
    }

    // Khởi tạo
    ensureContainer();
    const initBtn = document.getElementById('btn-docs-tab');
    if (initBtn) initBtn.className = BTN_IDLE;
    updateBadge();
    openDB().then(() => dbAll()).then(list => { userDocs = list || []; updateBadge(); }).catch(() => { });
    // Mở sẵn tab Văn bản khi URL có #van-ban (chờ trang khởi tạo xong layout mặc định)
    const openFromHash = () => { if (location.hash === '#van-ban' && !docsMode) setTimeout(() => setDocsMode(true), 150); };
    if (document.readyState === 'complete') openFromHash(); else window.addEventListener('load', openFromHash);
    window.addEventListener('hashchange', () => { if (location.hash === '#van-ban') setDocsMode(true); });
})();
