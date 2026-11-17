/* UNIVERSAL GRI-6 ENGINE (JS) - Sibling Tree + Polarity + Quantitative
 * Cong 1:1 voi universal_engine.py. Du lieu: window.TARIFF_ENGINE (tariff_engine.js).
 */
(function (root) {
    'use strict';
    const EPS = 1e-6;
    function norm(s) {
        if (!s) return '';
        return String(s).replace(/đ/g, 'd').replace(/Đ/g, 'D').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
    }
    const NUM = '(\\d{1,3}(?:\\.\\d{3})+|\\d+(?:[.,]\\d+)?)';
    const UNIT = '(kva|va|kv|kw|mw|ka|ma|hp|vac|vdc|volt|v|ampe|amp|a|w)';
    const UNIT_DIM = {
        v: ['V', 1], vac: ['V', 1], vdc: ['V', 1], volt: ['V', 1], kv: ['V', 1000],
        a: ['A', 1], amp: ['A', 1], ampe: ['A', 1], ma: ['A', 0.001], ka: ['A', 1000],
        w: ['W', 1], kw: ['W', 1000], mw: ['W', 1e6], hp: ['W', 746],
        va: ['VA', 1], kva: ['VA', 1000]
    };
    function parseNum(tok) {
        tok = tok.trim();
        if (/^\d{1,3}(?:\.\d{3})+$/.test(tok)) return parseFloat(tok.replace(/\./g, ''));
        return parseFloat(tok.replace(',', '.'));
    }
    const CMP = [
        ['khong\\s+qua', '<='], ['khong\\s+vuot\\s+qua', '<='], ['khong\\s+it\\s+hon', '>='],
        ['vuot\\s+qua', '>'], ['duoi', '<'], ['tren', '>']
    ];
    const ALIASES = {
        'vien thong': ['vien thong', 'telecom', 'telecommunication', 'mang lan', 'lan', 'cat5', 'cat5e', 'cat6', 'cat6a', 'cat7', 'cat 5', 'cat 6', 'ethernet', 'utp', 'ftp', 'sftp', 'dien thoai', 'tin hieu mang', 'cap mang'],
        'hoat dong bang dien': ['hoat dong bang dien', 'chay bang dien', 'chay dien', 'dung dien', 'dong co dien', 'motor dien', 'electrically operated', 'electric motor']
    };
    function escRe(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

    const MATERIALS = {
        kim_loai: 'kim\\s+loai|thep|sat|inox|nhom|hop\\s+kim|steel|metal|iron|brass|copper|alumin(?:i)?um|dong\\s+thau|bang\\s+dong',
        det: 'vat\\s+lieu\\s+det|soi\\s+det|vai\\s+det|soi\\s+vai|vai|textile|fabric|polyester',
        plastic: 'plastic|nhua|pvc|polyethylene|polypropylene|abs|polyme|polymer',
        cao_su: 'cao\\s+su|rubber',
        giay: 'giay|paper|bia\\s+cung',
        go: 'go|wood',
        thuy_tinh: 'thuy\\s+tinh|glass',
        gom: 'gom|ceramic|porcelain'
    };
    const MAT_RE = {};
    Object.keys(MATERIALS).forEach(k => { MAT_RE[k] = new RegExp('(?<![a-z0-9])(?:' + MATERIALS[k] + ')(?![a-z0-9])'); });
    function materialsOf(t) { return Object.keys(MAT_RE).filter(k => MAT_RE[k].test(t)).sort(); }

    function extractQuery(text) {
        const qn = norm(text);
        const vals = {};
        const add = (d, v) => { (vals[d] = vals[d] || []).push(v); };
        const consumed = [];
        for (const [pat, op] of CMP) {
            const re = new RegExp('\\b' + pat + '\\s+(?:khoang\\s+)?' + NUM + '\\s*' + UNIT + '\\b', 'g');
            let m;
            while ((m = re.exec(qn)) !== null) {
                const [dim, f] = UNIT_DIM[m[2]];
                let v = parseNum(m[1]) * f;
                if (op === '<') v -= EPS; else if (op === '>') v += EPS;
                add(dim, v);
                consumed.push([m.index, m.index + m[0].length]);
            }
        }
        const re2 = new RegExp('(?<![a-z0-9.,])' + NUM + '\\s*' + UNIT + '\\b', 'g');
        let m2;
        while ((m2 = re2.exec(qn)) !== null) {
            if (consumed.some(([s, e]) => s <= m2.index && m2.index < e)) continue;
            const [dim, f] = UNIT_DIM[m2[2]];
            add(dim, parseNum(m2[1]) * f);
        }
        const q = {};
        for (const d in vals) q[d] = Math.max.apply(null, vals[d]);
        return { qn: qn, q: q, mt: materialsOf(qn) };
    }

    function condOk(conds, q) {
        let checked = false;
        for (const [dim, op, val] of conds) {
            if (!(dim in q)) continue;
            checked = true;
            const v = q[dim];
            if (op === '<' && !(v < val)) return [false, true];
            if (op === '<=' && !(v <= val + 1e-9)) return [false, true];
            if (op === '>' && !(v > val)) return [false, true];
            if (op === '>=' && !(v >= val - 1e-9)) return [false, true];
        }
        return [true, checked];
    }

    function queryAssertions(qn, phrase) {
        const cands = ALIASES[phrase] || [phrase];
        let has = false, nothas = false;
        for (const p of cands) {
            const pe = escRe(p);
            if (new RegExp('\\b(?:khong|chua)\\s+(?:co\\s+|duoc\\s+|dung\\s+|lap\\s+|gan\\s+|su\\s+dung\\s+|hoat\\s+dong\\s+bang\\s+)?' + pe).test(qn) ||
                new RegExp('\\bnon[- ]' + pe).test(qn) || new RegExp('\\bwithout\\s+' + pe).test(qn)) {
                nothas = true;
            } else if (new RegExp('\\b' + pe + '(?![a-z0-9])').test(qn)) {
                has = true;
            }
        }
        if (phrase === 'hoat dong bang dien' && /\bkhong\s+(?:hoat\s+dong\s+bang\s+)?dien\b|\bbang\s+tay\b|\bthu\s+cong\b/.test(qn)) {
            nothas = true; has = false;
        }
        return [has, nothas];
    }

    function groupFailed(members, qn, q, qmt) {
        const failed = {};
        let matchedNumeric = false;
        members.forEach((m, i) => {
            const c = m.c || [];
            if (c.length) {
                const [ok, checked] = condOk(c, q);
                if (!ok) failed[i] = 'Khong thoa dieu kien so cua nhanh: ' + c.map(x => x[0] + x[1] + x[2]).join(', ');
                else if (checked && m.p) matchedNumeric = true;
            }
        });
        members.forEach((m, i) => {
            if (m.n) {
                const [has, nothas] = queryAssertions(qn, m.n);
                if (has && !nothas && !(i in failed)) failed[i] = 'Hang hoa co "' + m.n + '" nen loai nhanh "khong ' + m.n + '"';
            }
            (m.pp || []).forEach(p => {
                const [has, nothas] = queryAssertions(qn, p);
                if (nothas && !has && !(i in failed)) failed[i] = 'Hang hoa khong co "' + p + '" nen loai nhanh co "' + p + '"';
            });
        });
        let matchedMat = false, hitPure = false;
        const qm = new Set(qmt || []);
        if (qm.size) {
            const hit = members.map(m => (m.mt || []).some(x => qm.has(x)));
            hitPure = members.some((m, i) => hit[i] && m.mp);
            const contrast = members.some((m, i) => (m.mt || []).length && !hit[i]);
            if (hit.some(Boolean) && contrast) {
                matchedMat = true;
                members.forEach((m, i) => {
                    if ((m.mt || []).length && !hit[i] && !(i in failed)) failed[i] = 'Vat lieu cua hang khong khop nhanh nay (' + m.mt.join(', ') + '), loai nhanh';
                });
            }
        }
        if (matchedNumeric || (matchedMat && hitPure)) {
            members.forEach((m, i) => {
                if (m.o && !(i in failed)) failed[i] = 'Da co nhanh cung cap khop dieu kien (so/vat lieu), loai nhanh "Loai khac"';
            });
        }
        return failed;
    }

    // Tra ve [disqualified, reason, bonus]
    function evaluate(code, qinfo) {
        const data = root.TARIFF_ENGINE;
        if (!data) return [false, '', 0];
        const c = String(code).replace(/\./g, '');
        const qn = qinfo.qn, q = qinfo.q;
        const hc = data.heads[c.substring(0, 4)];
        if (hc) {
            const [ok] = condOk(hc, q);
            if (!ok) return [true, 'Ten nhom ' + c.substring(0, 4) + ' co dieu kien so khong thoa: ' + hc.map(x => x[0] + x[1] + x[2]).join(', '), 0];
        }
        let bonus = 0;
        const chain = data.chains[c] || [];
        for (const [gi, mi] of chain) {
            const members = data.groups[gi];
            const failed = groupFailed(members, qn, q, qinfo.mt);
            if (mi in failed) return [true, failed[mi], 0];
            const m = members[mi];
            if (m.c && m.c.length) { const [, checked] = condOk(m.c, q); if (checked) bonus += 300; }
            const ps = (m.n ? [m.n] : []).concat(m.pp || []);
            for (const p of ps) {
                const [has, nothas] = queryAssertions(qn, p);
                if ((m.n && nothas) || (m.pp && has)) bonus += 300;
            }
        }
        return [false, '', bonus];
    }

    let _lineIdx = null;
    function lineIndex() {
        if (_lineIdx) return _lineIdx;
        _lineIdx = [];
        const db = root.HS_DATABASE || [];
        for (const it of db) {
            const c = (it.c || '').replace(/\./g, '');
            if (c.length !== 8 || !it.lin) continue;
            const segs = it.lin.split(' > ').slice(1);
            _lineIdx.push([c, norm(segs.join(' > '))]);
        }
        return _lineIdx;
    }

    function fmt(c) { return c.length === 8 ? c.substring(0, 4) + '.' + c.substring(4, 6) + '.' + c.substring(6) : c; }

    function genericResolve(headNounNorm, fullText) {
        const clean = [];
        for (const t of (String(headNounNorm || '').match(/[a-z0-9]+/g) || [])) {
            if (/\d/.test(t)) break;
            clean.push(t);
        }
        const qi = extractQuery(fullText);
        const data = lineIndex();
        for (let n = clean.length; n >= 2; n--) {
            const phrase = clean.slice(0, n).join(' ');
            const re = new RegExp('(?<![a-z])' + escRe(phrase) + '(?![a-z])');
            const cands = data.filter(x => re.test(x[1]));
            if (!cands.length) continue;
            const surv = [], dropped = [];
            for (const [c] of cands) {
                const [dq, why] = evaluate(c, qi);
                (dq ? dropped : surv).push([c, why]);
            }
            if (!surv.length) return null;
            const subs = new Set(surv.map(x => x[0].substring(0, 6)));
            const removed = Array.from(new Set(dropped.map(x => fmt(x[0]) + ': ' + x[1]))).sort().join('; ').substring(0, 600);
            if (surv.length === 1) return { code: fmt(surv[0][0]), level: 8, status: 'EXACT_8_DIGIT', phrase: phrase, survivors: [surv[0][0]], dropped: removed };
            if (subs.size === 1) {
                const sub = Array.from(subs)[0];
                return { code: sub.substring(0, 4) + '.' + sub.substring(4, 6), level: 6, status: 'STOP_6_DIGIT', phrase: phrase, survivors: surv.map(x => x[0]), dropped: removed };
            }
            return null;
        }
        return null;
    }


    const TOK_STOP = new Set('loai khac cap dien cua tu hoac va dung cho khong chua co bang la nhung duoc de'.split(' '));
    function rerouteScope(code, fullText, prefix) {
        const c0 = String(code).replace(/\./g, '');
        const qi = extractQuery(fullText);
        const qtoks = new Set((qi.qn.match(/[a-z0-9]+/g) || []).filter(t => !TOK_STOP.has(t)));
        const surv = [];
        for (const it of (root.HS_DATABASE || [])) {
            const c = (it.c || '').replace(/\./g, '');
            if (c.length !== 8 || c.indexOf(prefix) !== 0) continue;
            const r = evaluate(c, qi);
            if (!r[0]) surv.push([c, it, r[2] || 0]);
        }
        if (!surv.length) return null;
        const why0 = c0.length === 8 ? evaluate(c0, qi)[1] : '';
        if (surv.length === 1) return { code: fmt(surv[0][0]), level: 8, status: 'EXACT_8_DIGIT', phrase: c0.substring(0, 4), survivors: [surv[0][0]], dropped: why0 };
        const scored = surv.map(([c, it, bn]) => {
            const txt = norm((it.vn || '').replace(/^[\s\-]+/, ''));
            const toks = new Set((txt.match(/[a-z0-9]+/g) || []).filter(t => !TOK_STOP.has(t)));
            let sc = 0;
            toks.forEach(t => { if (qtoks.has(t)) sc++; });
            const mm = materialsOf(txt).filter(x => qi.mt.indexOf(x) >= 0).length;
            return [sc + 2 * mm + 3 * Math.floor(bn / 300), c];
        }).sort((a, b) => b[0] - a[0]);
        if (scored[0][0] > 0 && (scored.length === 1 || scored[0][0] > scored[1][0])) {
            return { code: fmt(scored[0][1]), level: 8, status: 'EXACT_8_DIGIT', phrase: c0.substring(0, 4), survivors: scored.map(x => x[1]), dropped: why0 };
        }
        const subs = new Set(scored.map(x => x[1].substring(0, 6)));
        if (subs.size === 1) {
            const sub = Array.from(subs)[0];
            return { code: sub.substring(0, 4) + '.' + sub.substring(4, 6), level: 6, status: 'STOP_6_DIGIT', phrase: c0.substring(0, 4), survivors: scored.map(x => x[1]), dropped: why0 };
        }
        return null;
    }

    function reroute(code, fullText) {
        const c0 = String(code).replace(/\./g, '');
        const scopes = c0.length >= 6 ? [c0.substring(0, 6), c0.substring(0, 4)] : [c0.substring(0, 4)];
        for (const pre of scopes) {
            const r = rerouteScope(code, fullText, pre);
            if (r) return r;
        }
        return null;
    }

    root.UniversalEngine = { norm, extractQuery, evaluate, genericResolve, reroute, condOk, queryAssertions };
})(typeof window !== 'undefined' ? window : globalThis);
