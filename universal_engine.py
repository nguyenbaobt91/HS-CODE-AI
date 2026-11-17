# -*- coding: utf-8 -*-
"""
UNIVERSAL GRI-6 ENGINE (Sibling Tree + Polarity + Quantitative)
----------------------------------------------------------------
Doc dieu kien truc tiep tu cay phan cap (full_lineage) cua Bieu thue 2026,
KHONG viet luat rieng cho tung ma HS.

  1. Sibling Tree   : cac nhanh cung cap gach noi (cung cha) duoc so sanh voi nhau.
  2. Polarity       : cap "co X" / "khong X" (chua X) tu dong phat hien trong bieu thue.
  3. Quantitative   : dieu kien so (duoi / khong qua / tren ... V, A, W, VA) duoc doi
                      ve don vi chuan roi so voi thong so nguoi dung nhap.

Du lieu tien tinh (tariff_engine.json / tariff_engine.js) duoc sinh boi build_engine_data()
va dung chung cho Python (server) va JavaScript (index.html).
"""
import json
import os
import re
import unicodedata

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TARIFF_FILE = os.path.join(BASE_DIR, 'tariff_8digit_details.json')
ENGINE_JSON = os.path.join(BASE_DIR, 'tariff_engine.json')
ENGINE_JS = os.path.join(BASE_DIR, 'tariff_engine.js')

EPS = 1e-6

# ---------------------------------------------------------------------------
# Chuan hoa van ban
# ---------------------------------------------------------------------------
def norm(s):
    if not s:
        return ''
    s = s.replace('đ', 'd').replace('Đ', 'D')
    s = unicodedata.normalize('NFD', s)
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    return s.lower()


NUM = r'(\d{1,3}(?:\.\d{3})+|\d+(?:[.,]\d+)?)'
UNIT = r'(kva|va|kv|kw|mw|ka|ma|hp|vac|vdc|volt|v|ampe|amp|a|w)'
UNIT_DIM = {
    'v': ('V', 1.0), 'vac': ('V', 1.0), 'vdc': ('V', 1.0), 'volt': ('V', 1.0), 'kv': ('V', 1000.0),
    'a': ('A', 1.0), 'amp': ('A', 1.0), 'ampe': ('A', 1.0), 'ma': ('A', 0.001), 'ka': ('A', 1000.0),
    'w': ('W', 1.0), 'kw': ('W', 1000.0), 'mw': ('W', 1e6), 'hp': ('W', 746.0),
    'va': ('VA', 1.0), 'kva': ('VA', 1000.0),
}


def parse_num(tok):
    tok = tok.strip()
    if re.fullmatch(r'\d{1,3}(?:\.\d{3})+', tok):
        return float(tok.replace('.', ''))
    return float(tok.replace(',', '.'))


# ---------------------------------------------------------------------------
# Dieu kien so trong dong bieu thue
# ---------------------------------------------------------------------------
CMP = [
    (r'khong\s+qua', '<='),
    (r'khong\s+vuot\s+qua', '<='),
    (r'khong\s+it\s+hon', '>='),
    (r'vuot\s+qua', '>'),
    (r'duoi', '<'),
    (r'tren', '>'),
]


def parse_conditions(text_norm):
    """Tra ve list [dim, op, value_base_unit]. Tat ca dieu kien cung dim la AND."""
    conds = []
    t = text_norm
    # "tu X [unit] den Y unit"
    for m in re.finditer(r'tu\s+' + NUM + r'\s*' + UNIT + r'?\s+(?:den|toi)\s+' + NUM + r'\s*' + UNIT + r'\b', t):
        u2 = m.group(4)
        u1 = m.group(2) or u2
        d1, f1 = UNIT_DIM[u1]
        d2, f2 = UNIT_DIM[u2]
        if d1 == d2:
            conds.append([d1, '>=', parse_num(m.group(1)) * f1])
            conds.append([d2, '<=', parse_num(m.group(3)) * f2])
    # "tren X unit den Y unit"
    for m in re.finditer(r'tren\s+' + NUM + r'\s*' + UNIT + r'?\s+(?:den|toi)\s+' + NUM + r'\s*' + UNIT + r'\b', t):
        u2 = m.group(4)
        u1 = m.group(2) or u2
        d1, f1 = UNIT_DIM[u1]
        d2, f2 = UNIT_DIM[u2]
        if d1 == d2:
            conds.append([d1, '>', parse_num(m.group(1)) * f1])
            conds.append([d2, '<=', parse_num(m.group(3)) * f2])
    # "tu X unit tro len" / "tu X unit tro xuong"
    for m in re.finditer(r'tu\s+' + NUM + r'\s*' + UNIT + r'\s+tro\s+(len|xuong)\b', t):
        dim, f = UNIT_DIM[m.group(2)]
        conds.append([dim, '>=' if m.group(3) == 'len' else '<=', parse_num(m.group(1)) * f])
    # dieu kien don le
    for pat, op in CMP:
        for m in re.finditer(pat + r'\s+(?:khoang\s+)?' + NUM + r'\s*' + UNIT + r'\b', t):
            dim, f = UNIT_DIM[m.group(2)]
            c = [dim, op, parse_num(m.group(1)) * f]
            if c not in conds:
                conds.append(c)
    return conds


def cond_ok(conds, qvals):
    """conds: list [dim, op, val]; qvals: dict dim -> value. Dim vang mat -> bo qua (khong loai)."""
    checked = False
    for dim, op, val in conds:
        if dim not in qvals:
            continue
        checked = True
        q = qvals[dim]
        if op == '<' and not q < val:
            return False, True
        if op == '<=' and not q <= val + 1e-9:
            return False, True
        if op == '>' and not q > val:
            return False, True
        if op == '>=' and not q >= val - 1e-9:
            return False, True
    return True, checked


# ---------------------------------------------------------------------------
# Phu dinh (Polarity)
# ---------------------------------------------------------------------------
NEG_RE = re.compile(r'\b(?:khong|chua)\s+(?:co\s+|duoc\s+|dung\s+cho\s+|lap\s+|gan\s+|su\s+dung\s+)?([a-z0-9][a-z0-9 \-/]{2,60}?)(?=\s*(?:[,;:()]|$|\bva\b|\bhoac\b|\bnhung\b|\bdung\b|\bdoi\b))')

ALIASES = {
    'vien thong': ['vien thong', 'telecom', 'telecommunication', 'mang lan', 'lan', 'cat5', 'cat5e', 'cat6', 'cat6a', 'cat7',
                   'cat 5', 'cat 6', 'ethernet', 'utp', 'ftp', 'sftp', 'dien thoai', 'tin hieu mang', 'cap mang'],
    'hoat dong bang dien': ['hoat dong bang dien', 'chay bang dien', 'chay dien', 'dung dien', 'dong co dien',
                            'motor dien', 'electrically operated', 'electric motor'],
}


def neg_phrase(text_norm):
    m = NEG_RE.search(text_norm)
    if not m:
        return None
    p = m.group(1).strip()
    if len(p) < 3 or re.search(r'\d', p) or re.match(r'(?:qua|vuot|it\s+hon|hon)\b', p):
        return None
    return p


COND_STRIP = re.compile(r'(?:khong\s+vuot\s+qua|khong\s+qua|khong\s+it\s+hon|vuot\s+qua|duoi|tren|tu)\s+(?:khoang\s+)?'
                        + NUM + r'\s*' + UNIT + r'?(?:\s+(?:den|toi)\s+' + NUM + r'\s*' + UNIT + r')?\b')
STOP = set('loai khac co dong dien danh dinh muc dung cho ap cong suat tren duoi khong qua den va hoac nhung voi cua la trong tro len xuong'.split())


MAT_FILLER = set('loai khac chi duoc gia co bang cach dien cua tu hoac va dung cho khong chua'.split())


def is_pure_material(text_norm):
    rest = text_norm
    for rx in _MAT_RE.values():
        rest = rx.sub(' ', rest)
    toks = [t for t in re.findall(r'[a-z]+', rest) if t not in MAT_FILLER]
    return len(toks) == 0


def is_pure_condition(text_norm, max_tokens=3):
    """Chi nhanh mang dieu kien so thuan tuy (khong kem san pham cu the) moi so sanh dinh luong."""
    if ';' in text_norm:
        return False
    rest = COND_STRIP.sub(' ', text_norm)
    toks = [t for t in re.findall(r'[a-z]+', rest) if t not in STOP]
    return len(toks) <= max_tokens


def query_assertions(qn, phrase):
    """Tra ve (has, not_has) cua phrase (kem alias) trong cau truy van da chuan hoa."""
    cands = ALIASES.get(phrase, [phrase])
    has = False
    nothas = False
    for p in cands:
        pe = re.escape(p)
        if re.search(r'\b(?:khong|chua)\s+(?:co\s+|duoc\s+|dung\s+|lap\s+|gan\s+|su\s+dung\s+|hoat\s+dong\s+bang\s+)?' + pe, qn) \
                or re.search(r'\bnon[- ]' + pe, qn) or re.search(r'\bwithout\s+' + pe, qn):
            nothas = True
        elif re.search(r'\b' + pe + r'(?![a-z0-9])', qn):
            has = True
    if phrase == 'hoat dong bang dien' and re.search(r'\bkhong\s+(?:hoat\s+dong\s+bang\s+)?dien\b|\bbang\s+tay\b|\bthu\s+cong\b', qn):
        nothas = True
        has = False
    return has, nothas


# ---------------------------------------------------------------------------
# Vat lieu (so sanh vat lieu giua cac nhanh anh em)
# ---------------------------------------------------------------------------
MATERIALS = {
    'kim_loai': r'kim\s+loai|thep|sat|inox|nhom|hop\s+kim|steel|metal|iron|brass|copper|alumin(?:i)?um|dong\s+thau|bang\s+dong',
    'det': r'vat\s+lieu\s+det|soi\s+det|vai\s+det|soi\s+vai|vai|textile|fabric|polyester',
    'plastic': r'plastic|nhua|pvc|polyethylene|polypropylene|abs|polyme|polymer',
    'cao_su': r'cao\s+su|rubber',
    'giay': r'giay|paper|bia\s+cung',
    'go': r'go|wood',
    'thuy_tinh': r'thuy\s+tinh|glass',
    'gom': r'gom|ceramic|porcelain',
}
_MAT_RE = {k: re.compile(r'(?<![a-z0-9])(?:' + v + r')(?![a-z0-9])') for k, v in MATERIALS.items()}


def materials_of(text_norm):
    return sorted(k for k, rx in _MAT_RE.items() if rx.search(text_norm))


# ---------------------------------------------------------------------------
# Truy van: trich xuat dai luong
# ---------------------------------------------------------------------------
def extract_query(text):
    qn = norm(text)
    vals = {}

    def add(dim, v):
        vals.setdefault(dim, []).append(v)

    consumed = []
    # "duoi 16A" / "tren 16A" / "khong qua 16A" mo ta nguong -> suy ra gia tri
    for pat, op in CMP:
        for m in re.finditer(r'\b' + pat + r'\s+(?:khoang\s+)?' + NUM + r'\s*' + UNIT + r'\b', qn):
            dim, f = UNIT_DIM[m.group(2)]
            v = parse_num(m.group(1)) * f
            if op == '<':
                v -= EPS
            elif op == '>':
                v += EPS
            add(dim, v)
            consumed.append((m.start(), m.end()))
    for m in re.finditer(r'(?<![a-z0-9.,])' + NUM + r'\s*' + UNIT + r'\b', qn):
        if any(s <= m.start() < e for s, e in consumed):
            continue
        # tranh ket qua sai nhu "m12 x 50 w" khong co; 'a' don le can co so ngay truoc
        dim, f = UNIT_DIM[m.group(2)]
        add(dim, parse_num(m.group(1)) * f)
    q = {d: max(v) for d, v in vals.items()}
    return {'qn': qn, 'q': q, 'mt': materials_of(qn)}


# ---------------------------------------------------------------------------
# Xay du lieu tien tinh tu cay bieu thue
# ---------------------------------------------------------------------------
def _seg_text(seg):
    return seg.rstrip(':').strip()


def build_engine_data(tariff_path=TARIFF_FILE):
    with open(tariff_path, 'r', encoding='utf-8') as f:
        tariff = json.load(f)

    groups = {}          # key tuple -> {member_text: idx}
    group_members = []   # idx -> list of member dict
    key_index = {}       # tuple(prefix) -> group idx
    chains = {}
    heads = {}

    def member_of(gkey, text):
        if gkey not in key_index:
            key_index[gkey] = len(group_members)
            group_members.append({'order': [], 'by': {}})
        g = group_members[key_index[gkey]]
        if text not in g['by']:
            tn = norm(text)
            g['by'][text] = len(g['order'])
            g['order'].append({'t': tn})
        return key_index[gkey], g['by'][text]

    for code, it in tariff.items():
        lineage = it.get('full_lineage') or ''
        segs = [s.strip() for s in lineage.split(' > ')]
        if len(segs) < 2:
            continue
        head_txt = it.get('head_title') or segs[0]
        hn = norm(head_txt)
        hc = parse_conditions(hn)
        if hc:
            heads[it.get('head') or code[:4]] = hc
        chain = []
        for k in range(1, len(segs)):
            gkey = tuple(segs[:k])
            gi, mi = member_of(gkey, segs[k])
            chain.append([gi, mi])
        chains[code] = chain

    # Phan tich thanh vien
    out_groups = []
    remap = {}
    for gi, g in enumerate(group_members):
        members = g['order']
        negs = []
        for m in members:
            tn = m['t']
            m['c'] = parse_conditions(tn) if is_pure_condition(tn) else []
            m['p'] = 1 if (m['c'] and is_pure_condition(tn, 1)) else 0
            m['o'] = 1 if re.fullmatch(r'(?:loai\s+)?khac\s*:?', tn.strip().rstrip(':').strip()) else 0
            m['n'] = neg_phrase(tn)
            m['pp'] = []
            m['mt'] = materials_of(tn)
            m['mp'] = 1 if (m['mt'] and is_pure_material(tn)) else 0
        for m in members:
            if m['n']:
                for o in members:
                    if o is m:
                        continue
                    if not o['n'] and m['n'] in o['t']:
                        if m['n'] not in o['pp']:
                            o['pp'].append(m['n'])
        active = any(m['c'] or m['n'] or m['pp'] or m['mt'] for m in members)
        if active and len(members) >= 1:
            remap[gi] = len(out_groups)
            out_groups.append([{k: v for k, v in m.items() if (v or k == 't') and k != 't'} for m in members])
    # strip text
    out_chains = {}
    for code, chain in chains.items():
        keep = [[remap[gi], mi] for gi, mi in chain if gi in remap]
        if keep:
            out_chains[code] = keep
    return {'groups': out_groups, 'chains': out_chains, 'heads': heads}


def write_engine_data():
    data = build_engine_data()
    with open(ENGINE_JSON, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, separators=(',', ':'))
    with open(ENGINE_JS, 'w', encoding='utf-8') as f:
        f.write('window.TARIFF_ENGINE=' + json.dumps(data, ensure_ascii=False, separators=(',', ':')) + ';')
    return data


_DATA = None


def load_data():
    global _DATA
    if _DATA is None:
        if os.path.exists(ENGINE_JSON):
            with open(ENGINE_JSON, 'r', encoding='utf-8') as f:
                _DATA = json.load(f)
        else:
            _DATA = write_engine_data()
    return _DATA


# ---------------------------------------------------------------------------
# Danh gia: mot ma 8 so co bi loai boi cay gach noi khong?
# ---------------------------------------------------------------------------
def _group_failed(members, qn, q, qinfo_mt=None):
    """Tra ve {member_idx: ly do} cho cac nhanh bi loai trong 1 nhom anh em."""
    failed = {}
    matched_numeric = False
    for i, m in enumerate(members):
        c = m.get('c') or []
        if c:
            ok, checked = cond_ok(c, q)
            if not ok:
                failed[i] = 'Khong thoa dieu kien so cua nhanh: ' + ', '.join(f'{d}{o}{v:g}' for d, o, v in c)
            elif checked and m.get('p'):
                matched_numeric = True
    for i, m in enumerate(members):
        n = m.get('n')
        if n:
            has, nothas = query_assertions(qn, n)
            if has and not nothas:
                failed.setdefault(i, f'Hang hoa co "{n}" nen loai nhanh "khong {n}"')
        for p in m.get('pp') or []:
            has, nothas = query_assertions(qn, p)
            if nothas and not has:
                failed.setdefault(i, f'Hang hoa khong co "{p}" nen loai nhanh co "{p}"')
    qm = set(qinfo_mt or [])
    matched_mat = False
    hit_pure = False
    if qm:
        hit = [bool(set(m.get('mt') or []) & qm) for m in members]
        hit_pure = any(h and m.get('mp') for h, m in zip(hit, members))
        contrast = any((m.get('mt') and not hit[i]) for i, m in enumerate(members))
        if any(hit) and contrast:
            matched_mat = True
            for i, m in enumerate(members):
                if m.get('mt') and not hit[i] and i not in failed:
                    failed[i] = 'Vat lieu cua hang khong khop nhanh nay (' + ', '.join(m['mt']) + '), loai nhanh'
    if matched_numeric or (matched_mat and hit_pure):
        for i, m in enumerate(members):
            if m.get('o') and i not in failed:
                failed[i] = 'Da co nhanh cung cap khop dieu kien (so/vat lieu), loai nhanh "Loai khac"'
    return failed


def evaluate(code, qinfo):
    """Tra ve (disqualified: bool, reason: str, bonus: int)."""
    data = load_data()
    c = code.replace('.', '')
    qn, q = qinfo['qn'], qinfo['q']
    hc = data['heads'].get(c[:4])
    if hc:
        ok, _ = cond_ok(hc, q)
        if not ok:
            return True, 'Ten nhom ' + c[:4] + ' co dieu kien so khong thoa: ' + ', '.join(f'{d}{o}{v:g}' for d, o, v in hc), 0
    bonus = 0
    for gi, mi in data['chains'].get(c, []):
        members = data['groups'][gi]
        failed = _group_failed(members, qn, q, qinfo.get('mt'))
        if mi in failed:
            return True, failed[mi], 0
        m = members[mi]
        if m.get('c'):
            _, checked = cond_ok(m['c'], q)
            if checked:
                bonus += 300
        if m.get('n') or m.get('pp'):
            for p in ([m['n']] if m.get('n') else []) + (m.get('pp') or []):
                has, nothas = query_assertions(qn, p)
                if (m.get('n') and nothas) or (m.get('pp') and has):
                    bonus += 300
    return False, '', bonus


# ---------------------------------------------------------------------------
# Generic resolver: khop cum danh tu chinh voi cay bieu thue roi loc bang engine
# ---------------------------------------------------------------------------
_TARIFF = None


def _tariff():
    global _TARIFF
    if _TARIFF is None:
        with open(TARIFF_FILE, 'r', encoding='utf-8') as f:
            raw = json.load(f)
        _TARIFF = []
        for code, it in raw.items():
            segs = [x.strip() for x in (it.get('full_lineage') or '').split(' > ')][1:]
            _TARIFF.append((code, norm(' > '.join(segs)), it))
    return _TARIFF


def generic_resolve(head_noun_norm, full_text):
    """Tra ve dict {code, level, status, reasoning, survivors} hoac None neu van nhap nhang.
    Chi chap nhan khi sau khi engine loai nhanh sai, cac ma con lai CUNG 1 phan nhom 6 so
    (hoac duy nhat 1 ma 8 so). Khong doan mo."""
    toks = [t for t in re.findall(r'[a-z]+', head_noun_norm or '') ]
    # cat tai token dau tien co so
    clean = []
    for t in re.findall(r'[a-z0-9]+', head_noun_norm or ''):
        if re.search(r'\d', t):
            break
        clean.append(t)
    qi = extract_query(full_text)
    data = _tariff()
    for n in range(len(clean), 1, -1):
        phrase = ' '.join(clean[:n])
        pat = re.compile(r'(?<![a-z])' + re.escape(phrase) + r'(?![a-z])')
        cands = [(c, it) for c, ln, it in data if pat.search(ln)]
        if not cands:
            continue
        surv, dropped = [], []
        for c, it in cands:
            dq, why, _ = evaluate(c, qi)
            (dropped if dq else surv).append((c, it, why))
        if not surv:
            return None
        subs = {c[:6] for c, _, _ in surv}
        removed = '; '.join(sorted({f"{c[:4]}.{c[4:6]}.{c[6:]}: {w}" for c, _, w in dropped}))[:600]
        if len(surv) == 1:
            c, it, _ = surv[0]
            return {'code': f"{c[:4]}.{c[4:6]}.{c[6:]}", 'level': 8, 'status': 'EXACT_8_DIGIT', 'phrase': phrase,
                    'survivors': [c], 'dropped': removed}
        if len(subs) == 1:
            sub = next(iter(subs))
            return {'code': f"{sub[:4]}.{sub[4:6]}", 'level': 6, 'status': 'STOP_6_DIGIT', 'phrase': phrase,
                    'survivors': [c for c, _, _ in surv], 'dropped': removed}
        return None
    return None


TOK_STOP = set('loai khac cap dien cua tu hoac va dung cho khong chua co bang la nhung duoc de'.split())


def _reroute_scope(code, full_text, prefix):
    """Ma `code` (do walker/AI de xuat) bi engine loai -> tim ma thay the trong cung nhom 4 so.
    Chap nhan khi chi con 1 ma, hoac 1 ma co do khop dac thu (vat lieu / tu khoa) vuot troi; neu khong thi dung o 6 so."""
    c0 = code.replace('.', '')
    qi = extract_query(full_text)
    qtoks = set(re.findall(r'[a-z0-9]+', qi['qn'])) - TOK_STOP
    surv, dropped = [], []
    for c, ln, it in _tariff():
        if c.startswith(prefix):
            dq, why, bn = evaluate(c, qi)
            (dropped if dq else surv).append((c, it, why, bn))
    if not surv:
        return None
    why0 = evaluate(c0, qi)[1] if len(c0) == 8 else ''
    def fmt(c): return f"{c[:4]}.{c[4:6]}.{c[6:]}"
    if len(surv) == 1:
        return {'code': fmt(surv[0][0]), 'level': 8, 'status': 'EXACT_8_DIGIT', 'phrase': c0[:4], 'survivors': [surv[0][0]], 'dropped': why0}
    scored = []
    for c, it, _, bn in surv:
        txt = norm(re.sub(r'^[\s\-]+', '', it.get('vn', '')))
        toks = set(re.findall(r'[a-z0-9]+', txt)) - TOK_STOP
        sc = len(toks & qtoks) + 2 * len(set(materials_of(txt)) & set(qi['mt'])) + 3 * (bn // 300)
        scored.append((sc, c))
    scored.sort(reverse=True)
    if scored[0][0] > 0 and (len(scored) == 1 or scored[0][0] > scored[1][0]):
        c = scored[0][1]
        return {'code': fmt(c), 'level': 8, 'status': 'EXACT_8_DIGIT', 'phrase': c0[:4], 'survivors': [x[1] for x in scored], 'dropped': why0}
    subs = {c[:6] for _, c in scored}
    if len(subs) == 1:
        sub = next(iter(subs))
        return {'code': f"{sub[:4]}.{sub[4:6]}", 'level': 6, 'status': 'STOP_6_DIGIT', 'phrase': c0[:4], 'survivors': [x[1] for x in scored], 'dropped': why0}
    return None


def reroute(code, full_text):
    """Thu trong cung phan nhom 6 so cua ma bi loai truoc, neu khong ra moi mo rong ra nhom 4 so."""
    c0 = code.replace('.', '')
    scopes = [c0[:6], c0[:4]] if len(c0) >= 6 else [c0[:4]]
    for pre in scopes:
        r = _reroute_scope(code, full_text, pre)
        if r:
            return r
    return None


if __name__ == '__main__':
    d = write_engine_data()
    print('groups:', len(d['groups']), 'chains:', len(d['chains']), 'heads:', len(d['heads']))
    print('size json:', os.path.getsize(ENGINE_JSON))

