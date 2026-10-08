import json, os, glob, re, sys, unicodedata
import docx

def clean_code(c):
    if not c or c == 'N/A':
        return ''
    digits = re.sub(r'[^0-9]', '', str(c))
    if len(digits) >= 8:
        return digits[:8]
    elif len(digits) == 6:
        return digits + '00'
    return digits

def extract_keywords(text):
    if not text:
        return []
    nfkd = unicodedata.normalize('NFKD', text)
    clean = ''.join([c for c in nfkd if not unicodedata.combining(c)]).replace('đ', 'd').replace('Đ', 'D').lower()
    words = re.findall(r'[a-z0-9]+', clean)
    stop = {'va', 'cua', 'cho', 'dung', 'loai', 'co', 'bang', 'the', 'trong', 'duoc', 'voi', 'theo', 'nhu', 'de', 'dang', 'kich', 'thuoc', 'hang', 'muc', 'san', 'pham', 'danh'}
    kw = [w for w in words if w not in stop and len(w) > 1]
    return list(dict.fromkeys(kw))[:15]

rulings = []

# 1. Load from tavily_classification_updates.json
tavily_path = r'E:\AI\MOL\HS CODE\DL\tavily_classification_updates.json'
if os.path.exists(tavily_path):
    with open(tavily_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
        for item in data:
            code = clean_code(item.get('proposed_hs_code'))
            name = item.get('product_name') or item.get('document_title') or ''
            if code and name:
                rulings.append({
                    'doc': item.get('id') or item.get('document_title') or 'Phán quyết Hải quan',
                    'name': name,
                    'hs': code,
                    'basis': item.get('legal_basis') or 'Căn cứ Biểu thuế XNK & Chú giải WCO',
                    'summary': item.get('key_summary') or f'Phân loại mặt hàng {name} vào mã {code}.',
                    'keywords': extract_keywords(name)
                })

# 2. Load from classification_index.json
index_path = r'E:\AI\MOL\HS CODE\DL\classification_index.json'
if os.path.exists(index_path):
    with open(index_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
        for item in data:
            code = clean_code(item.get('hs_code'))
            name = item.get('description') or ''
            doc_name = item.get('filename') or item.get('decision_number') or ''
            if code and name and code != 'N/A':
                rulings.append({
                    'doc': doc_name.replace('.pdf', ''),
                    'name': name,
                    'hs': code,
                    'basis': 'Thông báo kết quả phân loại kiểm định của Tổng cục Hải quan',
                    'summary': f"Phân loại mặt hàng '{name}' vào mã HS {code}.",
                    'keywords': extract_keywords(name)
                })

# 3. Load from official 50+ .docx case files in E:\AI\MOL\HS CODE\
docx_files = glob.glob(r'E:\AI\MOL\HS CODE\*.docx')
for f in docx_files:
    fname = os.path.basename(f)
    if fname in ('HS CODE.docx', 'temp_weekly.docx') or fname.startswith('~$') or fname.endswith('.tmp'):
        continue
    m = re.match(r'^(\d{6,8})\.docx$', fname)
    if not m:
        continue
    file_hs = clean_code(m.group(1))
    try:
        doc = docx.Document(f)
        name = ''
        desc = ''
        use = ''
        basis = ''
        proposed_hs = file_hs
        for p in doc.paragraphs:
            txt = p.text.strip()
            if 'DANH MỤC SẢN PHẨM:' in txt:
                name = txt.replace('DANH MỤC SẢN PHẨM:', '').strip()
            elif 'Tên hàng' in txt and not name:
                name = txt.replace('Tên hàng / Model:', '').replace('Tên hàng:', '').strip()
            if 'Cấu tạo & Thông số kỹ thuật:' in txt:
                desc = txt.replace('Cấu tạo & Thông số kỹ thuật:', '').strip()
            if 'Công dụng & Chức năng:' in txt:
                use = txt.replace('Công dụng & Chức năng:', '').strip()
            if 'MÃ HS ĐỀ XUẤT ÁP DỤNG THỐNG NHẤT:' in txt:
                m_hs = re.search(r'(\d{4}\.\d{2}\.\d{2})', txt)
                if m_hs:
                    proposed_hs = m_hs.group(1).replace('.', '')
            if '· GRI 1' in txt or 'GRI 1 (' in txt:
                basis = txt

        if name:
            combined_text = f"{name} {desc} {use}"
            rulings.append({
                'doc': f'Hồ sơ Thẩm định Antigravity ({proposed_hs})',
                'name': name,
                'hs': proposed_hs,
                'basis': basis or 'Áp dụng 6 Quy tắc GRI và Chú giải WCO 2022',
                'summary': f"{name}. {desc}. {use}".strip(),
                'keywords': extract_keywords(combined_text)
            })
    except Exception as e:
        print(f"Error reading {fname}: {e}")

# 4. Add Authoritative Statutory Precedents from SKILL.md (15+ Gold Benchmarks)
statutory_benchmarks = [
    {
        'doc': 'Chú giải 1(g) Phần XVI & GRI 6 (TCHQ)',
        'name': 'Tấm che bảo vệ an toàn bằng nhựa dùng cho máy dập kim loại',
        'hs': '39269099',
        'basis': 'Áp dụng GRI 1 & Chú giải 1(g) Phần XVI, các sản phẩm bằng chất dẻo thuộc Chương 39 (Nhóm 39.26) bị loại trừ khỏi Phần XVI máy móc (loại trừ nhóm phụ tùng 84.66). Áp dụng GRI 6 xếp vào mã 3926.90.99.',
        'summary': 'Tấm che bảo vệ bằng nhựa cho máy móc bắt buộc phân loại vào 3926.90.99, không xếp vào nhóm phụ tùng máy 8466.',
        'keywords': ['tam', 'che', 'bao', 've', 'nhua', 'an', 'toan', 'may', 'dap']
    },
    {
        'doc': 'Chú giải 2 Phần XV & GRI 6 (TCHQ)',
        'name': 'Miếng thép gắn cố định cảm biến kích thước 45x75x5mm, bản mã thép',
        'hs': '73269099',
        'basis': 'Áp dụng GRI 1 & Chú giải 2 Phần XV, các sản phẩm gia công cơ khí bằng kim loại cơ bản chưa định danh riêng phân loại tại Nhóm 73.26.',
        'summary': 'Sản phẩm gia công bằng sắt thép chưa định danh xếp vào 7326.90.99.',
        'keywords': ['mieng', 'thep', 'gia', 'cong', 'co', 'dinh', 'cam', 'bien', 'ban', 'ma']
    },
    {
        'doc': 'Phân nhóm 8536.50 Biểu thuế XNK 2026 (TCHQ)',
        'name': 'Công tắc điện chuyển mạch đóng ngắt dưới 16A, 220V',
        'hs': '85365096',
        'basis': 'Áp dụng GRI 1 & Chú giải 2(a) Phần XVI vào Nhóm 85.36. Áp dụng GRI 6 so sánh cùng cấp gạch nối: điện áp <= 1000V và dòng điện định mức dưới 16A bắt buộc vào mã 8536.50.96.',
        'summary': 'Công tắc đóng ngắt mạch điện hạ thế dưới 16A bắt buộc phân loại 8536.50.96 (loại trừ 8536.50.99 dành cho từ 16A trở lên).',
        'keywords': ['cong', 'tac', 'dien', 'dong', 'ngat', 'chuyen', 'mach', '10a', '220v']
    },
    {
        'doc': 'Phân nhóm 8536.50 Biểu thuế XNK 2026 (TCHQ)',
        'name': 'Công tắc điện chuyển mạch đóng ngắt từ 16A trở lên, 220V, 20A, 30A',
        'hs': '85365099',
        'basis': 'Áp dụng GRI 1 & Chú giải 2(a) Phần XVI vào Nhóm 85.36. Áp dụng GRI 6 so sánh cùng cấp gạch nối: điện áp <= 1000V và dòng điện định mức từ 16A trở lên bắt buộc vào mã 8536.50.99.',
        'summary': 'Công tắc đóng ngắt mạch điện hạ thế từ 16A trở lên bắt buộc phân loại 8536.50.99.',
        'keywords': ['cong', 'tac', 'dien', 'dong', 'ngat', '16a', '20a', '30a', '220v']
    },
    {
        'doc': 'Chú giải 2(a) Phần XVI & Chú giải chi tiết WCO 84.82',
        'name': 'Thanh trượt bằng thép có ổ bi dẫn hướng tịnh tiến (con trượt bi, linear guide)',
        'hs': '84821000',
        'basis': 'Áp dụng GRI 1 & Chú giải 2(a) Phần XVI, ổ bi trong mọi trường hợp ưu tiên tuyệt đối phân loại vào Nhóm 84.82, loại trừ khỏi nhóm bộ phận máy 84.66. Cơ chế trượt có ổ bi thuộc Nhóm 84.82 (GRI 6 mã 8482.10.00).',
        'summary': 'Thanh trượt / con trượt có ổ bi dẫn hướng bắt buộc phân loại vào 8482.10.00.',
        'keywords': ['thanh', 'truot', 'o', 'bi', 'con', 'truot', 'dan', 'huong', 'linear', 'guide']
    },
    {
        'doc': 'Chú giải Nhóm 84.83 Biểu thuế XNK 2026',
        'name': 'Thân trượt dẫn hướng bằng gang không có ổ bi dùng cho máy tiện',
        'hs': '84833090',
        'basis': 'Áp dụng GRI 1 & Chú giải Nhóm 84.83 (Thân gối đỡ và bạc lót trục không lắp ổ bi). Áp dụng GRI 6 loại trừ nhánh ô tô 8483.30.30, xếp vào loại khác dùng cho máy công nghiệp mã 8483.30.90.',
        'summary': 'Thân trượt / gối đỡ không có bi thuộc phân nhóm 8483.30, mã 8483.30.90.',
        'keywords': ['than', 'truot', 'khong', 'co', 'bi', 'goi', 'do', 'bac', 'lot']
    },
    {
        'doc': 'Chú giải 2 Phần XV & Phân nhóm 7318.15 Biểu thuế 2026',
        'name': 'Bu lông lục giác bằng thép M12x50 có ren',
        'hs': '73181510',
        'basis': 'Áp dụng GRI 1 & Chú giải 2 Phần XV, các chi tiết ghép nối kim loại cơ bản phân loại tại Nhóm 73.18. Áp dụng GRI 6 so sánh đường kính thân d = 12mm <= 16mm vào mã 7318.15.10.',
        'summary': 'Bu lông lục giác đường kính từ 16mm trở xuống bắt buộc vào mã 7318.15.10 (loại trừ mã 7318.15.90 trên 16mm).',
        'keywords': ['bu', 'long', 'luc', 'giac', 'thep', 'm12', 'm12x50', 'ren']
    },
    {
        'doc': 'Chú giải Nhóm 85.23 Biểu thuế 2026',
        'name': 'Ổ lưu trữ thể rắn USB flash drive dung lượng 32GB, thẻ nhớ',
        'hs': '85235111',
        'basis': 'Áp dụng GRI 1 & Chú giải Nhóm 85.23: Thiết bị lưu trữ dữ liệu bán dẫn bền vững thể rắn (Solid-state non-volatile storage). Áp dụng GRI 6 phân nhóm: Phân nhóm 8523.51 -> Chưa ghi dữ liệu -> Mã 8523.51.11.',
        'summary': 'USB flash drive chưa ghi dữ liệu phân loại vào mã 8523.51.11.',
        'keywords': ['usb', 'flash', 'drive', 'o', 'luu', 'tru', 'the', 'ran', 'nho', '32gb']
    },
    {
        'doc': 'Chú giải Nhóm 85.44 & Phân nhóm 8544.20 Biểu thuế 2026',
        'name': 'Dây cáp đồng trục có gắn đầu nối, vỏ bọc cách điện nhựa',
        'hs': '85442011',
        'basis': 'Áp dụng GRI 1 Nhóm 85.44. Áp dụng GRI 6: Phân nhóm 8544.20 (Cáp đồng trục) -> Cáp cách điện ĐÃ GẮN VỚI ĐẦU NỐI, dùng cho điện áp không quá 66 kV -> Cách điện bằng cao su hoặc plastic -> Bắt buộc phân loại mã 8544.20.11 (loại trừ mã 8544.20.21 là loại CHƯA gắn đầu nối).',
        'summary': 'Cáp đồng trục có gắn đầu nối điện áp dưới 66kV cách điện nhựa/cao su bắt buộc phân loại 8544.20.11.',
        'keywords': ['day', 'cap', 'dong', 'truc', 'dau', 'noi', 'coaxial', 'cable']
    },
    {
        'doc': 'Chú giải Nhóm 85.44 & Phân nhóm 8544.42 Biểu thuế 2026',
        'name': 'Dây cáp điện cách điện bằng nhựa PVC có gắn đầu nối điện áp dưới 1000V',
        'hs': '85444294',
        'basis': 'Áp dụng GRI 1 Nhóm 85.44. Áp dụng GRI 6: Phân nhóm 8544.42 (Dây dẫn điện khác, dùng cho điện áp không quá 1000V, đã gắn với các đầu nối) -> Mã 8544.42.94.',
        'summary': 'Dây cáp điện hạ thế có gắn đầu nối phân loại vào 8544.42.94.',
        'keywords': ['day', 'cap', 'dien', 'pvc', 'dau', 'noi', 'ha', 'the', '1000v']
    },
    {
        'doc': 'Chú giải Nhóm 84.87 & Chú giải loại trừ Nhóm 84.81 (HQKV8 Trang 214194)',
        'name': 'Vú tra mỡ bò, vú bơm mỡ bò bằng thép không tự động (grease nipple)',
        'hs': '84879000',
        'basis': 'Căn cứ Chú giải chi tiết WCO Nhóm 84.87 và Chú giải loại trừ Nhóm 84.81, vú tra mỡ không tự động là bộ phận cơ khí máy móc bắt buộc phân loại vào Nhóm 84.87, mã 8487.90.00. Loại trừ khỏi 73.18, 73.07, 84.81.',
        'summary': 'Vú tra mỡ bò bằng thép phân loại vào mã 8487.90.00 (Thuế MFN: 0%, VAT: 8%).',
        'keywords': ['vu', 'tra', 'mo', 'bom', 'bo', 'grease', 'nipple', 'zerk']
    },
    {
        'doc': 'Công văn TCHQ & Chú giải Nhóm 39.26 / 87.08 (Dual Ruling)',
        'name': 'Nẹp nhựa cửa xe ô tô dùng để gắn trang trí thân xe',
        'hs': '39263000',
        'basis': 'Áp dụng GRI 1 & GRI 6 phân loại vào Nhóm 39.26 (Phụ kiện lắp vào đồ nội thất, trên thân xe - coachwork - mã 3926.30.00). Đối chiếu quan điểm phân loại bộ phận thân xe ô tô Nhóm 87.08 (mã 8708.29.93).',
        'summary': 'Nẹp trang trí / phụ kiện thân vỏ ô tô bằng nhựa ưu tiên mã 3926.30.00, đối chiếu mã phụ tùng thân xe 8708.29.93.',
        'keywords': ['nep', 'nhua', 'cua', 'xe', 'o', 'to', 'trang', 'tri', 'than', 'coachwork']
    },
    {
        'doc': 'Chú giải 2(a) Phần XVI & Nhóm 85.04',
        'name': 'Bộ chuyển đổi nguồn điện tĩnh AC sang DC (adapter, bộ nguồn switching)',
        'hs': '85044090',
        'basis': 'Áp dụng GRI 1 & Chú giải 2(a) Phần XVI, máy biến đổi điện tĩnh thuộc Nhóm 85.04. Áp dụng GRI 6 phân nhóm 8504.40 mã 8504.40.90.',
        'summary': 'Bộ nguồn adapter chuyển đổi AC sang DC phân loại vào 8504.40.90.',
        'keywords': ['bo', 'nguon', 'chuyen', 'doi', 'dien', 'tinh', 'adapter', 'switching', 'power', 'supply']
    },
    {
        'doc': 'Chú giải Nhóm 90.32 Biểu thuế 2026',
        'name': 'Bộ điều nhiệt tự động (thermostat, rơ le nhiệt điều chỉnh nhiệt độ)',
        'hs': '90321000',
        'basis': 'Áp dụng GRI 1 & Chú giải 1(m) Phần XVI, thiết bị thuộc Chương 90 bị loại trừ khỏi Phần XVI. Căn cứ Nhóm 90.32, bộ điều nhiệt tự động (Thermostats) thuộc mã 9032.10.00.',
        'summary': 'Bộ điều nhiệt Thermostat tự động phân loại vào mã 9032.10.00.',
        'keywords': ['bo', 'dieu', 'nhiet', 'tu', 'dong', 'thermostat', 'ro', 'le', 'nhiet']
    }
]

for b in statutory_benchmarks:
    rulings.append(b)

# Deduplicate by (hs, name lower)
seen = set()
unique_rulings = []
for r in rulings:
    k = (r['hs'], r['name'].strip().lower()[:60])
    if k not in seen:
        seen.add(k)
        unique_rulings.append(r)

out_file = r'E:\AI\MOL\HS CODE\web\rulings_data.js'
with open(out_file, 'w', encoding='utf-8') as f:
    f.write('// AUTHORITATIVE CUSTOMS RULINGS & PRECEDENTS DATABASE\n')
    f.write('// Extracted from E:\\AI\\MOL\\HS CODE (classification_index, tavily_updates, 54 docx cases, statutory benchmarks)\n')
    f.write('window.CUSTOMS_RULINGS_DATABASE = ' + json.dumps(unique_rulings, indent=2, ensure_ascii=False) + ';\n')

print(f"Generated {out_file} with {len(unique_rulings)} authoritative customs rulings!")
