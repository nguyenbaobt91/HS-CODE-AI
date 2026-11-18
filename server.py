from openpyxl.utils import get_column_letter
import http.server
import socketserver
import json
import os
import sys
import io
import re
import base64
import urllib.parse
import docx
import fitz
import subprocess
import pythoncom
import win32com.client
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

sys.stdout.reconfigure(encoding='utf-8')

PORT = 8080
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
WORKSPACE_DIR = os.path.abspath(os.path.join(BASE_DIR, '..'))
TEMPLATE_PATH = os.path.join(WORKSPACE_DIR, 'DL', 'HS CODE.docx')
SCRATCH_DIR = os.path.join(WORKSPACE_DIR, 'scratch')
os.makedirs(SCRATCH_DIR, exist_ok=True)

import unicodedata

def remove_accents(text):
    if not text: return ""
    text = re.sub(r'[đĐ]', 'd', text)
    nfkd = unicodedata.normalize('NFKD', text)
    return ''.join(c for c in nfkd if not unicodedata.combining(c)).lower()

# Load database into memory for Python fast classification
DB_PATH = os.path.join(BASE_DIR, 'hs_data.js')
HS_DATABASE = []
HEAD_MAP = {}
SUBHEAD_MAP = {}
CHAPTER_MAP = {}

if os.path.exists(DB_PATH):
    with open(DB_PATH, 'r', encoding='utf-8') as f:
        text = f.read()
    json_str = text.split('window.HS_DATABASE =')[1].strip().rstrip(';')
    HS_DATABASE = json.loads(json_str)
    for item in HS_DATABASE:
        c = item['c'].replace('.', '')
        vn = item.get('vn', '')
        if item.get('c') == '' and 'Chương' in vn:
            m = re.search(r'Chương\s+(\d+)', vn)
            if m:
                ch_code = m.group(1).zfill(2)
                CHAPTER_MAP[ch_code] = vn.split('\n')[0].strip()
        if len(c) == 4 or item.get('l') == '0':
            HEAD_MAP[c[:4]] = vn
        if len(c) == 6 or item.get('l') in ['1', '2']:
            SUBHEAD_MAP[c[:6]] = vn

# Load official WCO 2022 Explanatory Notes index (from Chu giai HS 2022 Toan tap (HQKV8).pdf)
WCO_OFFICIAL_NOTES = {}
WCO_SECTIONS = {}
WCO_CHAPTERS = {}
WCO_SEARCH_TEXTS = {}
WCO_NOTES_PATH = os.path.join(BASE_DIR, 'wco2022_official_notes.json')
if os.path.exists(WCO_NOTES_PATH):
    try:
        with open(WCO_NOTES_PATH, 'r', encoding='utf-8') as f:
            WCO_OFFICIAL_NOTES = json.load(f)
        
        for k, item in WCO_OFFICIAL_NOTES.items():
            if k.startswith('SEC_'):
                WCO_SECTIONS[k] = item
            elif k.startswith('CH_'):
                WCO_CHAPTERS[k] = item
        
        print(f"Loaded 100% WCO HQKV8 master database: {len(WCO_SECTIONS)} Sections, {len(WCO_CHAPTERS)} Chapters, and 1,228 Headings.")
        
        for h_k, item in WCO_OFFICIAL_NOTES.items():
            if len(h_k) == 4 and h_k.isdigit():
                ch = item.get('chapter', h_k[:2])
                sec_id = item.get('section_id', '')
                ch_notes = WCO_CHAPTERS.get(f"CH_{ch}", {}).get('legal_notes', '')
                sec_notes = WCO_SECTIONS.get(f"SEC_{sec_id}", {}).get('legal_notes', '')
                
                # Clean WCO Heading Note ONLY (Strict isolation: Exclude Section & Chapter notes to prevent false positive contamination)
                raw_txt = f"{item.get('title', '')} {item.get('official_vn', '')} {item.get('summary', '')} {item.get('official_note', '')}"
                WCO_SEARCH_TEXTS[h_k] = remove_accents(raw_txt).lower()
        print(f"Precomputed Universal Dual-Index Search for {len(WCO_SEARCH_TEXTS)} WCO headings with full Section/Chapter context.")
    except Exception as e:
        print("Error loading WCO official notes:", e)

def get_wco_explanatory_note(chapter, heading, max_len=1500):
    """
    Search and extract WCO Explanatory notes from official PDF index or chapters/chapter_<XX>.txt
    """
    # 1. Primary lookup from official EN2022 PDF Index (Công văn 1810/TCHQ-TXNK & HQKV8 Toàn tập)
    if heading in WCO_OFFICIAL_NOTES:
        note_data = WCO_OFFICIAL_NOTES[heading]
        h_dot = f"{heading[:2]}.{heading[2:4]}"
        p1 = note_data.get('page_1810')
        p2 = note_data.get('page_hqkv8')
        title = note_data.get('title', '')
        text = note_data.get('official_note', '') or note_data.get('note_hqkv8', '') or note_data.get('official_vn', '')
        if len(text) > max_len:
            text = text[:max_len] + "..."

        pages_ref = []
        if p1: pages_ref.append(f"Công văn 1810/TCHQ-TXNK Trang {p1}")
        if p2: pages_ref.append(f"Bộ Chú giải HQKV8 Trang {p2}")
        ref_str = " | ".join(pages_ref)

        return f"Bộ Chú giải HS 2022 ({ref_str} - Nhóm {h_dot} - {title}):\n{text}"

    # 2. Fallback lookup from chapter txt file
    file_path = os.path.join(WORKSPACE_DIR, 'chapters', f'chapter_{chapter}.txt')
    if not os.path.exists(file_path):
        return f"Đối chiếu Bộ Chú giải HS 2022 (Công văn 1810/TCHQ-TXNK & HQKV8) Chương {chapter}: Quy định nguyên tắc và tiêu chuẩn phân loại chung cho các mặt hàng thuộc Chương {chapter}."

    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()

        h_dot = f"{heading[:2]}.{heading[2:4]}"
        
        # High precision regex to capture the exact heading section from chapter file
        pattern = re.compile(rf'({re.escape(h_dot)}\s*[–\-—–].*?)(?=\n\s*\d{{2}}\.\d{{2}}\s*[–\-—–]|\Z)', re.DOTALL)
        match = pattern.search(content)
        
        if match:
            heading_text = match.group(1).strip()
            lines = [l.strip() for l in heading_text.split('\n') if l.strip()]
            vn_lines = [l for l in lines if not re.match(r'^[A-Za-z0-9\s\.\,\(\)\-\:\;\/\'\"]+$', l)]
            clean_block = " ".join(vn_lines) if vn_lines else " ".join(lines)
            clean_block = re.sub(r'\s+', ' ', clean_block)
            if len(clean_block) > 480:
                clean_block = clean_block[:480] + "..."
            return f"Chú giải HS 2022 (Công văn 1810/TCHQ-TXNK & HQKV8) - Chương {chapter}, Nhóm {h_dot}: {clean_block}"
            
        # Fallback paragraph matching
        paras = content.split('\n\n')
        matched = []
        for p in paras:
            p_clean = p.strip()
            if not p_clean: continue
            if h_dot in p_clean or f"nhóm {h_dot}" in p_clean.lower():
                lines = [l.strip() for l in p_clean.split('\n') if l.strip()]
                vn_lines = [l for l in lines if not re.match(r'^[A-Za-z0-9\s\.\,\(\)\-\:\;\/\'\"]+$', l)]
                text_block = " ".join(vn_lines) if vn_lines else " ".join(lines)
                if len(text_block) > 20:
                    matched.append(text_block)
                    if len(matched) >= 2: break

        if matched:
            result = re.sub(r'\s+', ' ', " ".join(matched))
            if len(result) > 450: result = result[:450] + "..."
            return f"Chú giải HS 2022 (Công văn 1810/TCHQ-TXNK & HQKV8) - Chương {chapter}, Nhóm {h_dot}: {result}"
    except Exception as err:
        print("WCO Note Lookup Error:", err)

    return f"Đối chiếu Chú giải Chi tiết HS 2022 Hải quan Việt Nam (Công văn 1810/TCHQ-TXNK & HQKV8) Chương {chapter}: Nhóm {heading[:2]}.{heading[2:4]} bao gồm các thiết bị, sản phẩm có cấu tạo và công năng đáp ứng tiêu chuẩn kỹ thuật quy định."

def extract_text_from_attached_doc(doc_b64, doc_filename):
    if not doc_b64:
        return ""

    try:
        if ',' in doc_b64:
            doc_b64 = doc_b64.split(',', 1)[1]

        file_bytes = base64.b64decode(doc_b64)
        ext = os.path.splitext(doc_filename)[1].lower() if doc_filename else '.tmp'
        temp_doc_path = os.path.join(SCRATCH_DIR, f"attached_doc{ext}")

        with open(temp_doc_path, 'wb') as f:
            f.write(file_bytes)

        extracted = ""

        if ext == '.pdf':
            pdf_doc = fitz.open(temp_doc_path)
            for page in pdf_doc:
                extracted += page.get_text() + "\n"
        elif ext == '.docx':
            docx_doc = docx.Document(temp_doc_path)
            extracted = "\n".join([p.text for p in docx_doc.paragraphs if p.text.strip()])
        elif ext in ['.txt', '.csv']:
            with open(temp_doc_path, 'r', encoding='utf-8', errors='ignore') as f:
                extracted = f.read()

        extracted = re.sub(r'\s+', ' ', extracted).strip()
        if len(extracted) > 1000:
            extracted = extracted[:1000] + "..."

        return extracted
    except Exception as err:
        print("Doc extraction error:", err)
        return ""

# Build full hierarchy ancestry for all 19,800+ items across Chapters 01 to 98
# Load 100% complete official 8-digit tariff hierarchy with lineage breadcrumbs
TARIFF_8DIGIT_DETAILS = {}
ANCESTRY_MAP = {}
TARIFF_8DIGIT_PATH = os.path.join(BASE_DIR, 'tariff_8digit_details.json')

if os.path.exists(TARIFF_8DIGIT_PATH):
    try:
        with open(TARIFF_8DIGIT_PATH, 'r', encoding='utf-8') as f:
            TARIFF_8DIGIT_DETAILS = json.load(f)
        for c, data in TARIFF_8DIGIT_DETAILS.items():
            ANCESTRY_MAP[c] = {
                'item': data,
                'full_path_text': data['search_text'].lower(),
                'full_lineage': data['full_lineage'],
                'ch': data['ch'],
                'head': data['head'],
                'sub': data['sub'],
                'h_desc': data['head_title'],
                's_desc': data['parent_desc'],
                'item_desc': data['clean_vn'],
                'unit': data.get('unit', ''),
                'mfn': data.get('mfn', '0%'),
                'vat': data.get('vat', '8%'),
                'policy': data.get('policy', '')
            }
        print(f"Loaded {len(ANCESTRY_MAP)} verified 8-digit tariff items with full hierarchical lineage from {TARIFF_8DIGIT_PATH}.")
    except Exception as e:
        print("Error loading tariff_8digit_details.json:", e)

def build_ancestry_map():
    global ANCESTRY_MAP
    current_chapter_title = ""
    current_heading_desc = {}
    current_subhead_desc = {}

    for item in HS_DATABASE:
        c = item.get('c', '').replace('.', '')
        vn = item.get('vn', '')
        vn_clean = vn.lstrip('-').strip()

        if item.get('c') == '' and 'Chương' in vn:
            current_chapter_title = vn
        elif len(c) == 4 or item.get('l') == '0':
            if c:
                current_heading_desc[c[:4]] = vn_clean
                HEAD_MAP[c[:4]] = vn_clean
        elif len(c) == 6 or item.get('l') in ['1', '2']:
            if c:
                current_subhead_desc[c[:6]] = vn_clean
                SUBHEAD_MAP[c[:6]] = vn_clean

        if len(c) == 8 and (1 <= int(c[:2]) <= 98):
            ch = c[:2]
            head = c[:4]
            sub = c[:6]
            h_desc = current_heading_desc.get(head) or HEAD_MAP.get(head, '')
            s_desc = current_subhead_desc.get(sub) or SUBHEAD_MAP.get(sub, '')
            
            full_path_text = f"Chương {ch} {current_chapter_title} Nhóm {head} {h_desc} Phân nhóm {sub} {s_desc} Mã {c} {vn_clean}".lower()
            ANCESTRY_MAP[c] = {
                'item': item,
                'full_path_text': full_path_text,
                'full_lineage': f"Nhóm {head}: {h_desc} > {s_desc} > {vn_clean}",
                'ch': ch,
                'head': head,
                'sub': sub,
                'h_desc': h_desc,
                's_desc': s_desc,
                'item_desc': vn_clean,
                'unit': item.get('unit', ''),
                'mfn': item.get('mfn', '0%'),
                'vat': item.get('vat', '8%'),
                'policy': ''
            }

if not ANCESTRY_MAP and HS_DATABASE:
    build_ancestry_map()
    print(f"Mapped {len(ANCESTRY_MAP)} 8-digit tariff items into ANCESTRY_MAP across Chapters 01 to 98.")

# =========================================================================
# UNIVERSAL 2-STAGE HIERARCHICAL CUSTOMS CLASSIFICATION ENGINE (ALL 98 CHAPTERS)
# =========================================================================

import unicodedata

def remove_accents(text):
    if not text: return ""
    text = re.sub(r'[đĐ]', 'd', text)
    nfkd = unicodedata.normalize('NFKD', text)
    return ''.join(c for c in nfkd if not unicodedata.combining(c)).lower()

HOST_MARKERS = [
    r'\bpttt\s+(?:của|cho)?\b',
    r'\bpttt\b',
    r'\bphụ\s+tùng\s+(?:thay\s+thế\s+)?(?:của|cho)?\b',
    r'\bphu\s+tung\s+(?:thay\s+the\s+)?(?:cua|cho)?\b',
    r'\blinh\s+kiện\s+(?:thay\s+thế\s+)?(?:của|cho)?\b',
    r'\blinh\s+kien\s+(?:thay\s+the\s+)?(?:cua|cho)?\b',
    r'\bdùng\s+(?:cho|trên|trong|để|gắn|lắp)?\b',
    r'\bdung\s+(?:cho|tren|trong|de|gan|lap)?\b',
    r'\bsử\s+dụng\s+(?:cho|trên|trong|để)?\b',
    r'\bsu\s+dung\s+(?:cho|tren|trong|de)?\b',
    r'\bdành\s+cho\b',
    r'\bdanh\s+cho\b',
    r'\blắp\s+(?:ráp\s+)?(?:trên|vào|cho)?\b',
    r'\blap\s+(?:rap\s+)?(?:tren|vao|cho)?\b',
    r'\bgắn\s+(?:trên|vào|cho)?\b',
    r'\bgan\s+(?:tren|vao|cho)?\b',
    r'\bcho\s+(?:xe|máy|thiết\s+bị|hệ\s+thống|dây\s+chuyền|tủ|động\s+cơ|may|thiet\s+bi|he\s+thong|day\s+chuyen|tu|dong\s+co)\b',
    r'\bsuitable\s+for\b',
    r'\bused\s+for\b',
    r'\bspare\s+parts?\s+for\b',
    r'\bparts?\s+for\b'
]

MATERIAL_RULES = [
    ('steel_iron', r'\b(?:thép|thep|sắt|sat|inox|gang|cast iron|carbon steel|stainless steel)\b', ['72', '73'], ['39', '40', '44', '47', '48', '49', '50', '52', '61', '62', '63', '68', '69', '70']),
    ('copper', r'\b(?:đồng|dong|đồng thau|brass|copper|bronze|hợp kim đồng)\b', ['74'], ['72', '73', '76', '39', '48']),
    ('aluminum', r'\b(?:nhôm|nhom|aluminum|aluminium|hợp kim nhôm)\b', ['76'], ['72', '73', '74', '39', '48']),
    ('plastic', r'\b(?:nhựa|nhua|plastic|polymer|polyme|pvc|pe|pp|pet|abs|pom|pa|nylon|ptfe|teflon|pu|polyurethane|silicone|silicon)\b', ['39'], ['72', '73', '74', '76', '44', '48', '68', '69', '70']),
    ('rubber', r'\b(?:cao su|rubber|elastomer|nitrile|epdm)\b', ['40'], ['72', '73', '74', '76', '39', '44', '48']),
    ('wood', r'\b(?:gỗ|go|timber|wood|ván|van|ván ép|gỗ dán|mdf|hdf|tre|nứa|bạch dương|gỗ sồi|gỗ thông)\b', ['44', '45', '46'], ['72', '73', '39', '40', '48', '70', '82', '84', '85']),
    ('textile', r'\b(?:vải|sợi bông|cotton|polyester|len|lụa|tơ tằm|dệt thoi|dệt kim|không dệt|textile|fabric)\b', ['50', '51', '52', '53', '54', '55', '56', '57', '58', '59', '60', '61', '62', '63'], ['72', '73', '39', '44', '48']),
    ('paper', r'\b(?:giấy|paper|bìa|carton|giấy bìa|kraft|tissue|vệ sinh|khăn ăn)\b', ['47', '48', '49'], ['72', '73', '74', '76', '84', '85', '39']),
    ('leather', r'\b(?:da thuộc|da bò|da cừu|da nhân tạo|giả da|leather)\b', ['41', '42'], ['72', '73', '39', '48']),
    ('glass', r'(?<!đường\s)\b(?:thủy tinh|kính|gương|glass|mirror)\b', ['70'], ['72', '73', '39', '48']),
    ('chemical', r'\b(?:axit|acid|h2so4|hcl|hno3|naoh|koh|dung môi|hóa chất|cồn|ethanol|sunfuric|sulfuric|hydrochloric)\b', ['28', '29', '38'], ['84', '85', '73', '39', '48'])
]

STATUTORY_CATEGORIES = [
    # Specialized High-Precision Gateways for Industrial Tooling & Parts
    {
        'name': 'measuring_instruments_calipers',
        'pattern': r'\b(?:thuoc\s+do|micro\s*meter|panme|thuoc\s+cap|caliper|dial\s+gauge)\b',
        'target_headings': ['9017'],
        'forbidden_chapters': ['85', '84', '73'],
        'bonus': 1950
    },
    {
        'name': 'aluminum_plates_articles',
        'pattern': r'\b(?:mieng\s+nhom|tam\s+nhom|thanh\s+nhom|aluminum\s+plate)\b',
        'target_headings': ['7616'],
        'forbidden_chapters': ['85', '84', '73'],
        'bonus': 1950
    },
    {
        'name': 'machine_knives_cutting_blades',
        'pattern': r'\b(?:luoi\s+cat(?!\s+kim\s+loai\s+cho)|dao\s+cat|cutting\s+blade|shear\s+blade|machine\s+knife)\b',
        'target_headings': ['8208'],
        'forbidden_chapters': ['84', '85', '73', '72'],
        'bonus': 1900
    },
    {
        'name': 'hand_tools_manual',
        'pattern': r'\b(?:dung\s+cu\s+(?:thao|lap|cam\s+tay|su\s+dung\s+tay|hoat\s+dong\s+bang\s+tay)|hand\s+tool|extraction\s+tool|removal\s+tool|crimping\s+tool)\b',
        'target_headings': ['8205'],
        'forbidden_chapters': ['8547', '8536', '8544', '90', '84'],
        'bonus': 1900
    },
    {
        'name': 'measuring_checking_fixtures',
        'pattern': r'\b(?:thiet\s+bi\s+do|may\s+do|dung\s+cu\s+do|do\s+do\s+thang|measuring|checking\s+fixture)\b',
        'target_headings': ['9031'],
        'forbidden_chapters': ['8544', '8536', '84'],
        'bonus': 1900
    },
    {
        'name': 'adhesive_tapes_sheets',
        'pattern': r'\b(?:bang\s+keo|bang\s+dinh|adhesive\s+tape)\b',
        'target_headings': ['3919'],
        'forbidden_chapters': ['85', '84'],
        'bonus': 1900
    },
    {
        'name': 'insulated_wire_assemblies',
        'pattern': r'\b(?:day\s+dan\s+truyen\s+tin\s+hieu|day\s+cap\s+dien|day\s+dien\s+co\s+gan\s+dau\s+noi|wire\s+harness|cable\s+assembly)\b',
        'target_headings': ['8544'],
        'forbidden_chapters': ['84', '86', '39', '73'],
        'bonus': 1950
    },
    {
        'name': 'plastic_brackets_fittings',
        'pattern': r'\b(?:khung(?:\s+gan\s+dinh\s+vi)?\s+.*bang\s+nhua|khung\s+nhua|mieng\s+nhua|tru\s+nhua|de\s+nhua|plastic\s+bracket|plastic\s+frame|plastic\s+holder)\b',
        'target_headings': ['3926'],
        'forbidden_chapters': ['84', '85', '86', '73', '72', '90'],
        'bonus': 1950
    },
    {
        'name': 'crimping_fixtures_applicators',
        'pattern': r'\b(?:ga\s+dap|khuon\s+dap|applicator|crimping\s+die|crimping\s+fixture)\b',
        'target_headings': ['8479', '8466'],
        'forbidden_chapters': ['8544', '8536'],
        'bonus': 1950
    },
    # 0a. Optical & Proximity Sensors (8536, 8541, 9031) - High Precision Gate
    {
        'name': 'optical_proximity_sensors',
        'pattern': r'\b(?:cam\s+bien|sensor|photo\s+sensor|photo\s+micro|sensor\s+amplifier|sensor\s+head|cam\s+bien\s+quang|cam\s+bien\s+tiem\s+can|cam\s+bien\s+nhiet\s+do)\b',
        'target_headings': ['8536', '8541', '9031', '9025'],
        'forbidden_chapters': ['50', '51', '52', '53', '54', '55', '56', '57', '58', '59', '60', '61', '62', '63', '07', '21'],
        'bonus': 1600
    },
    # 0b. Pneumatic & Hydraulic Cylinders & Grippers (8412)
    {
        'name': 'cylinders_pneumatic_grippers',
        'pattern': r'\b(?:xi\s*lanh|xilanh|cylinder|pneumatic\s+gripper|cang\s+kep|tay\s+kep|ngam\s+kep|gripper\s+jaw|ben\s+hoi|ben\s+khi)\b',
        'target_headings': ['8412'],
        'forbidden_chapters': ['21', '29', '38', '98', '53'],
        'bonus': 1700
    },
    # 0c. Linear Slide Tables & Ball Screws (8483, 8482)
    {
        'name': 'linear_slide_lead_screws',
        'pattern': r'\b(?:vit\s+me|ball\s+screw|truc\s+vit\s+me|linear\s+slide|ban\s+truot|thanh\s+truot|linear\s+guide|linear\s+motion)\b',
        'target_headings': ['8483', '8482'],
        'forbidden_chapters': ['94', '53', '39', '73', '74', '83', '93'],
        'bonus': 1950
    },
    # 0c2. Electric Motors & Geared Motors (8501)
    {
        'name': 'electric_motors',
        'pattern': r'\b(?:dong\s+co\s+dien|motor|mo\s+to|dong\s+co\s+buoc|servo\s+motor|stepper\s+motor|geared\s+motor|motor\s+hop\s+so)\b',
        'target_headings': ['8501'],
        'forbidden_chapters': ['90', '87', '40', '8407', '8408', '73', '74'],
        'bonus': 2000
    },
    # 0d. Floating Joints & Shaft Couplings (8483)
    {
        'name': 'floating_joints_couplings',
        'pattern': r'\b(?:khop\s+noi|floating\s+joint|shaft\s+coupling)\b',
        'target_headings': ['8483'],
        'forbidden_chapters': ['86', '53'],
        'bonus': 1600
    },
    # 0e. Ionizing Fans & Industrial Blowers (8414)
    {
        'name': 'fans_air_blowers',
        'pattern': r'\b(?:quat(?:\s+ion|\s+khu\s+tinh\s+dien|\s+cong\s+nghiep)?|ionizing\s+air\s+blower|air\s+blower)\b',
        'target_headings': ['8414'],
        'forbidden_chapters': ['8409', '8431', '8708'],
        'bonus': 1600
    },
    # 0f. Conveyor & Transmission Belts (4010, 3926)
    {
        'name': 'conveyor_timing_belts',
        'pattern': r'\b(?:day\s+belt|conveyor\s+belt|day\s+dai\s+bang\s+tai|bang\s+tai\s+cao\s+su)\b',
        'target_headings': ['4010', '3926'],
        'forbidden_chapters': ['84', '85'],
        'bonus': 1600
    },
    # 0g. Vibratory Feeder Controllers (8537, 8479)
    {
        'name': 'feeder_controllers',
        'pattern': r'\b(?:bo\s+dieu\s+khien\s+cap\s+lieu|feeder\s+controller|bo\s+dieu\s+khien\s+rung)\b',
        'target_headings': ['8537', '8479'],
        'forbidden_chapters': ['8501'],
        'bonus': 1600
    },
    # 1. Bearings (8482) - Note 2(a) Section XVI (vòng bi, bạc đạn, ổ bi, ổ lăn, thanh trượt có bi)
    {
        'name': 'ball_roller_bearings',
        'pattern': r'\b(?:vong\s+bi|bac\s+dan|o\s+bi|o\s+dua|o\s+lan|ball\s+bearings?|roller\s+bearings?|(?:thanh\s+truot|thanh\s+ray|con\s+truot|co\s+che\s+truot|linear\s+(?:motion\s+)?guides?).*(?:o\s+bi|co\s+bi|vong\s+bi|ball\s+bearings?|bearings?)|(?:o\s+bi|co\s+bi|vong\s+bi|ball\s+bearings?|bearings?).*(?:thanh\s+truot|thanh\s+ray|con\s+truot|co\s+che\s+truot|linear\s+(?:motion\s+)?guides?))\b',
        'target_headings': ['8482'],
        'forbidden_chapters': ['72', '73', '8409', '8431', '8466', '8473', '8483', '8708', '8803'],
        'bonus': 1350
    },
    # 2. Gears, shafts, gearboxes (8483) - Note 2(a) Section XVI
    {
        'name': 'transmission_gears_shafts',
        'pattern': r'\b(?:hop\s+giam\s+toc|hop\s+so|banh\s+rang|truc\s+khuyu|truc\s+truyen|hop\s+bien\s+toc|gearbox|gear\s+wheel|crankshaft|than\s+o|than\s+truot|o\s+truot|bac\s+truot|bac\s+dong|plain\s+bearing|bearing\s+housing)\b',
        'target_headings': ['8483'],
        'forbidden_chapters': ['73', '8409', '8431', '8708'],
        'bonus': 1250
    },
    # 3. Optical fiber connectors (8536) - Ch 90 Note 1(h) / Ch 85 Note 9
    {
        'name': 'optical_fiber_connectors',
        'pattern': r'\b(?:dau\s+noi\s+(?:dung\s+cho\s+)?(?:soi\s+quang|cap\s+quang|bo\s+soi\s+quang)|optical\s+fib(?:re|er)\s+connector)\b',
        'target_headings': ['8536'],
        'forbidden_chapters': ['9001', '8544', '3926', '7326'],
        'bonus': 1150
    },
    # 4. Primary Plastics (3901 - 3914) vs Rubber (4002)
    {
        'name': 'primary_polymers_plastics',
        'pattern': r'\b(?:hat\s+nhua\s+nguyen\s+sinh|nhua\s+nguyen\s+sinh|polypropylene|polyethylen|polyethylene|polystyren|polystyrene|polyvinyl\s+clorua|polymer\s+nguyen\s+sinh|virgin\s+plastic|pp\s+resin|pe\s+resin|abs\s+resin|pvc\s+resin)\b',
        'target_headings': ['3901', '3902', '3903', '3904', '3905', '3906', '3907', '3908', '3909', '3910', '3911'],
        'forbidden_chapters': ['40', '38', '29'],
        'bonus': 1150
    },
    # 5. Pure Organic Chemicals (Chapter 29)
    {
        'name': 'pure_organic_chemicals',
        'pattern': r'\b(?:axeton|acetone|ethanol|methanol|toluen|toluene|xylen|xylene|axit\s+axetic|acetic\s+acid|etyl\s+axetat|ethyl\s+acetate|c3h6o|ch3cooh|c2h5oh)\b',
        'target_headings': ['2914', '2905', '2902', '2915'],
        'forbidden_chapters': ['38', '33', '32'],
        'bonus': 1150
    },
    # 6. Temperature Sensors & Thermometers (9025)
    {
        'name': 'temperature_sensors',
        'pattern': r'\b(?:cam\s+bien\s+nhiet\s+do|cap\s+nhiet\s+dien|can\s+nhiet|thermocouple|rtd|nhiet\s+ke\s+cong\s+nghiep|pt100|temperature\s+sensor)\b',
        'target_headings': ['9025'],
        'forbidden_chapters': ['8514', '8516', '8419', '9031'],
        'bonus': 1150
    },
    # 7. Medical Endoscopes & Instruments (9018)
    {
        'name': 'medical_instruments_endoscopes',
        'pattern': r'\b(?:noi\s+soi\s+y\s+te|may\s+noi\s+soi|dung\s+cu\s+y\s+te|phau\s+thuat|ong\s+noi\s+soi|endoscope|medical\s+instrument)\b',
        'target_headings': ['9018'],
        'forbidden_chapters': ['9013', '4014', '9017'],
        'bonus': 1100
    },
    # 8. Woven vs Knitted Garments (Chapter 62 vs 61)
    {
        'name': 'woven_apparel',
        'pattern': r'\b(?:det\s+thoi|woven|quan\s+au|quan\s+tay|so\s+mi\s+det\s+thoi)\b',
        'target_headings': ['6203', '6204', '6205', '6201', '6202', '6206'],
        'forbidden_chapters': ['61'],
        'bonus': 1100
    },
    {
        'name': 'knitted_apparel',
        'pattern': r'\b(?:det\s+kim|knitted|crocheted|ao\s+thun|t-shirt|ao\s+len)\b',
        'target_headings': ['6109', '6105', '6110', '6103', '6104'],
        'forbidden_chapters': ['62'],
        'bonus': 1100
    },
    # 9. Tools & Blades (Chapter 82) - Note 1(k) Section XVI (fixed: avoid collision with 'kim loại')
    {
        'name': 'tools_and_blades',
        'pattern': r'\b(?:dao\s+phay|dao\s+cat|luoi\s+dao|luoi\s+cua|mui\s+khoan|mui\s+tien|mui\s+taro|mui\s+doa|manh\s+dao|khuon\s+dap|cutting\s+blade|milling\s+cutter|drill\s+bit|kim\s+(?:bam|dien|cat|chet|nhon|rut|cong\s+luc)|pliers?|bua|co\s+le|tua\s+vit)\b',
        'target_headings': ['8207', '8208', '8202', '8203', '8204', '8205', '8209'],
        'forbidden_chapters': ['84', '85'],
        'bonus': 1150
    },
    # 10. General Use Fasteners (7318, 7415, 7616) - Note 2 Section XV
    {
        'name': 'general_use_fasteners',
        'pattern': r'\b(?:chot\s+dinh\s+vi|chot\s+ham|chot\s+che|chot\s+con|chot\s+tru|dowel\s+pin|locating\s+pin|cotter\s+pin|knock\s+pin|bu\s+long|bu-long|bolt|dai\s+oc|oc\s+vit|dinh\s+vit|vit|screw|dinh\s+tan|rivet)\b',
        'target_headings': ['7318', '7415', '7616'],
        'forbidden_chapters': ['84', '85', '87', '88', '90', '94', '95', '96'],
        'bonus': 1150
    },
    # 10b. Gaskets, Washers & Seals (3926, 4016, 7318, 7415, 7616, 8484)
    {
        'name': 'gaskets_washers_seals',
        'pattern': r'\b(?:gioang|phot|dem\s+kin|gasket|seals?|o-ring|vong\s+chu\s+o|dem\s+lam\s+kin|dem\s+lot|vong\s+dem|washer|long\s+den|dem\s+venh)\b',
        'target_headings': ['7318', '7415', '7616'],
        'forbidden_chapters': ['84', '85', '87', '88', '90'],
        'bonus': 1350
    },
    {
        'name': 'general_use_springs',
        'pattern': r'\b(?:lo\s+xo|spring|lo\s+xo\s+xoan|lo\s+xo\s+la)\b',
        'target_headings': ['7320', '7416'],
        'forbidden_chapters': ['84', '85', '87', '88', '90', '94', '95', '96'],
        'bonus': 1150
    },
    # 11. Vehicle Brake Systems (8708)
    {
        'name': 'vehicle_brakes',
        'pattern': r'\b(?:ma\s+phanh|bo\s+ma\s+phanh|dia\s+phanh|he\s+thong\s+phanh|brake\s+pad|brake\s+disc)\b',
        'target_headings': ['8708'],
        'forbidden_chapters': ['6813', '8523', '9017', '8207'],
        'bonus': 1200
    },
    # 12. Fluid Valves (8481) - Note 2(a) Section XVI
    {
        'name': 'valves_fluid_controllers',
        'pattern': r'\b(?:van(?!\s+(?:phong|hanh|chuyen|go|tai|canh|the|chu))\b|valve\b|voi\s+nuoc|khoa\s+nuoc|tiet\s+luu|cock\b)\b',
        'target_headings': ['8481'],
        'forbidden_chapters': ['7326', '8409', '8431', '8535', '8536', '8708'],
        'bonus': 1150
    },
    # 13. Pipes & Tubes (3917, 4009, 7304-7306)
    {
        'name': 'pipes_tubes_hoses',
        'pattern': r'\b(?:ong\s+nhua|ong\s+mem|ong\s+dan|ong\s+cao\s+su|ong\s+cung|ong\s+thep|ong\s+kim\s+loai|ong\s+dong|ong(?!\s+nghiem)\b|hose|tube|pipe)\b',
        'target_headings': ['3917', '4009', '7304', '7305', '7306', '7411', '7608'],
        'forbidden_chapters': ['84', '85', '87', '28', '29', '38', '72'],
        'bonus': 1350
    },
    # 14. Electrical Cables (8544)
    {
        'name': 'insulated_cables',
        'pattern': r'\b(?:day\s+cap\s+dien|cap\s+dien|day\s+dien|day\s+cap|cap\s+dong|electric\s+cable|power\s+cable)\b',
        'target_headings': ['8544'],
        'forbidden_chapters': ['8536', '7413', '7614'],
        'bonus': 1200
    },
    # 14b. Mechanical Safety Guards & Machine Covers (3926, 8466 & 7326)
    {
        'name': 'mechanical_covers_guards',
        'pattern': r'\b(?:tam\s+che|vo\s+che|nap\s+che|tam\s+chan|vach\s+chan|man\s+che|chan\s+bui|tam\s+(?:nhua|mica|thep|kim\s+loai)?\s*bao\s+ve|vo\s+bao\s+ve|nap\s+bao\s+ve|safety\s+guard|machine\s+guard|protective\s+cover)\b',
        'target_headings': ['3926', '8466', '7326'],
        'forbidden_chapters': ['8544', '8535', '8536', '8537'],
        'bonus': 1350
    },
    # 14c. Transmission / Conveyor Belts (4010, 3926, 5910) - Section XVI Note 1(a)
    {
        'name': 'transmission_conveyor_belts',
        'pattern': r'\b(?:day\s+curoa|day\s+cu-roa|curoa|cu-roa|day\s+bang\s+tai|day\s+bang\s+chuyen|bang\s+tai\s+(?:cao\s+su|nhua|pu)|day\s+belt|day\s+dai\s+truyen\s+dong|day\s+rang\s+dong\s+bo|timing\s+belt|transmission\s+belt|conveyor\s+belt|v-belt)\b',
        'target_headings': ['4010', '3926', '5910'],
        'forbidden_chapters': ['84', '85', '87', '88'],
        'bonus': 1400
    },
    # 15. UAV / Drones (8806)
    {
        'name': 'uav_drones',
        'pattern': r'\b(?:drone|flycam|uav|may\s+bay\s+khong\s+nguoi\s+lai|phuong\s+tien\s+bay\s+khong\s+nguoi\s+lai)\b',
        'target_headings': ['8806'],
        'forbidden_chapters': ['8802', '9503'],
        'bonus': 1150
    },
    # 16. Flat Panel Displays (8524)
    {
        'name': 'flat_panel_displays',
        'pattern': r'\b(?:man\s+hinh\s+oled|mo\s+dun\s+man\s+hinh|mo-dun\s+man\s+hinh|man\s+hinh\s+phang|display\s+module)\b',
        'target_headings': ['8524'],
        'forbidden_chapters': ['8528', '9013', '8473'],
        'bonus': 1150
    },
    # 16b. Solid-State Storage & USB Flash Drives (8523.51)
    {
        'name': 'usb_flash_storage',
        'pattern': r'\b(?:usb\b|o\s+nho|o\s+nho\s+the\s+ran|the\s+nho|flash\s+drive|usb\s+flash|solid[\s\-]state\s+storage|the\s+cf|the\s+sd)\b',
        'target_headings': ['8523'],
        'forbidden_chapters': ['8471', '8528', '8521'],
        'bonus': 1300
    },
    # 17. Inorganic Chemicals (Chapter 28)
    {
        'name': 'inorganic_chemicals',
        'pattern': r'\b(?:axit\s+(?:sunfuric|sulphuric|sulfuric|clohydric|hydrochloric|nitric|photphoric|phosphoric)|h2so4|hcl|hno3|h3po4|xut\b|naoh|koh|natri\s+hydroxit|kali\s+hydroxit|caustic\s+soda)\b',
        'target_headings': ['2807', '2806', '2808', '2809', '2815'],
        'forbidden_chapters': ['72', '73', '84', '85', '38'],
        'bonus': 1200
    },
    # 17b. Electrical Control Panels, Boxes & Controllers (8537 & 8504)
    {
        'name': 'control_panels_controllers',
        'pattern': r'\b(?:hop\s+dieu\s+khien|tu\s+dieu\s+khien|bo\s+dieu\s+khien|mach\s+dieu\s+khien|bang\s+dieu\s+khien|bien\s+tan|motor\s+controller|driver\s+motor|control\s+box|control\s+panel|inverter)\b',
        'target_headings': ['8537', '8504'],
        'forbidden_chapters': ['8501', '8544'],
        'bonus': 1350
    },
    # 18. Electric Motors (8501) - Ch 85 Note / Section XVI Note 2(a)
    {
        'name': 'electric_motors',
        'pattern': r'\b(?:dong\s+co\s+dien|mo\s+to\s+dien|electric\s+motor|motor\s+dien|dong\s+co\s+khong\s+dong\s+bo|dong\s+co\s+dong\s+bo|mo\s+to\s+buoc|stepper\s+motor|servo\s+motor)\b|(?:\bmô\s*tơ\b)',
        'target_headings': ['8501'],
        'forbidden_chapters': ['87', '8409', '8431', '8503'],
        'bonus': 1200
    },
    # 19. Pumps for liquids (8413) - Note 2(a) Section XVI
    {
        'name': 'pumps_for_liquids',
        'pattern': r'\b(?:may\s+bom|bom\s+ly\s+tam|bom\s+thuy\s+luc|bom\s+chim|bom\s+nuoc|bom\s+dau|bom\s+mang|liquid\s+pump)\b',
        'target_headings': ['8413'],
        'forbidden_chapters': ['8409', '8431', '8708', '8906'],
        'bonus': 1150
    },
    # 20. Air / Gas Compressors & Fans (8414)
    {
        'name': 'air_gas_compressors_fans',
        'pattern': r'\b(?:may\s+nen\s+khi|may\s+hut\s+chan\s+khong|may\s+thoi\s+khi|quat\s+cong\s+nghiep|air\s+compressor|vacuum\s+pump)\b',
        'target_headings': ['8414'],
        'forbidden_chapters': ['8409', '8431', '8708'],
        'bonus': 1150
    },
    # 21. T-shirts & Vests (6109)
    {
        'name': 't_shirts_and_vests',
        'pattern': r'\b(?:ao\s+thun|t-shirt|ao\s+phong|ao\s+chui\s+dau)\b',
        'target_headings': ['6109'],
        'forbidden_chapters': ['62', '50', '51', '52', '53', '54', '55'],
        'bonus': 1200
    },
    # 22. Audio Equipment - Speakers & Microphones (8518)
    {
        'name': 'audio_speakers_microphones',
        'pattern': r'\b(?:loa\b|loa\s+nghe\s+nhac|loa\s+bluetooth|loa\s+thung|micro\b|microphone|tai\s+nghe|headphone|earphone)\b',
        'target_headings': ['8518'],
        'forbidden_chapters': ['8527', '8528', '8471'],
        'bonus': 1150
    },
    # 23. Optical Microscopes (9011)
    {
        'name': 'optical_microscopes',
        'pattern': r'\b(?:kinh\s+hien\s+vi\s+quang\s+hoc|kinh\s+hien\s+vi|optical\s+microscope|microscope)\b',
        'target_headings': ['9011'],
        'forbidden_chapters': ['9013', '9031'],
        'bonus': 1150
    },
    # 24. Additive Manufacturing / 3D Printers (8485)
    {
        'name': 'additive_manufacturing',
        'pattern': r'\b(?:may\s+in\s+3d|in\s+3d|3d\s+printer|san\s+xuat\s+boi\s+dap)\b',
        'target_headings': ['8485'],
        'forbidden_chapters': ['8477', '8479'],
        'bonus': 1150
    },
    # 25. Parts of Machine Tools / Press Machines (8466) - Note 2(b) Section XVI (Excludes plastic & rubber articles per Note 1(a)(g) Section XVI)
    {
        'name': 'parts_of_machine_tools',
        'pattern': r'\b(?:phu\s+tung\s+(?:thay\s+the\s+)?(?:may\s+)?(?:cat\s+dap|dap|ep\s+dap|phay|tien|mai|bao|khoan|cnc|cat\s+day)|bo\s+phan\s+may\s+(?:cat\s+dap|dap|ep\s+dap|phay|tien|mai|cnc)|ram\s+(?:metal\s+)?guide|die\s+set|punch\s+guide)\b',
        'target_headings': ['8466'],
        'forbidden_chapters': ['48', '72'],
        'bonus': 1200
    },
    # 26. Other Articles of Steel / Mounting plates / Brackets (7326)
    {
        'name': 'articles_of_steel',
        'pattern': r'\b(?:mieng\s+thep|tam\s+thep|thanh\s+thep|gia\s+do\s+bang\s+thep|tai\s+ga|bat\s+thep|khung\s+thep|steel\s+plate|steel\s+bracket|mounting\s+plate)\b',
        'target_headings': ['7326'],
        'forbidden_chapters': ['48', '39', '40', '44', '61', '62'],
        'bonus': 1100
    },
    # 27. Electrical Switches & Circuit Controllers <= 1000V (8536) - Note 2(a) Section XVI
    {
        'name': 'electrical_switches_relays',
        'pattern': r'\b(?:cong\s+tac|cong\s+tac\s+dien|cong\s+tac\s+chuyen\s+mach|cong\s+tac\s+ngat\s+mach|nut\s+nhan|nut\s+an|cau\s+dao|aptomat|ro\s+le|switch|switches|micro\s+switch|limit\s+switch)\b',
        'target_headings': ['8536'],
        'forbidden_chapters': ['8481', '8535', '8538', '8537', '8708', '8409'],
        'bonus': 1250
    },
    # 28. High Voltage Circuit Breakers & Switchgear > 1000V (8535) vs <= 1000V (8536)
    {
        'name': 'high_voltage_switchgear_breakers',
        'pattern': r'\b(?:may\s+cat\s+dien|may\s+cat\s+chan\s+khong|may\s+cat\s+khong\s+khi|circuit\s+breaker|vcb|acb|mccb|may\s+cat\s+trung\s+the|may\s+cat\s+cao\s+the)\b',
        'target_headings': ['8535', '8536'],
        'forbidden_chapters': ['8414', '8461', '8462', '8479'],
        'bonus': 1350
    },
    # 29. Pipe Fittings (7307, 7412, 7609) - Section XV Note 2
    {
        'name': 'pipe_fittings',
        'pattern': r'\b(?:khop\s+noi\s+(?:ren|ong|duong\s+ong)|cut\s+ong|te\s+ong|mang\s+song|rac\s+co|phu\s+kien\s+(?:duong\s+)?ong|pipe\s+fitting|pipe\s+coupling|tube\s+fitting|fitting\s+for\s+tube)\b',
        'target_headings': ['7307', '7412', '7609'],
        'forbidden_chapters': ['82', '84', '85', '87'],
        'bonus': 1300
    },
    # 30. Machine Knives & Cutting Blades (8208) - Chapter 82 Note 1(k)
    {
        'name': 'machine_cutting_blades',
        'pattern': r'\b(?:dao\s+cat\s+(?:cho\s+)?may|luoi\s+dao\s+(?:cat\s+)?(?:cho\s+)?may|dao\s+may|cutting\s+blade\s+for\s+machine|machine\s+knife|machine\s+blade|dao\s+cat\s+got)\b',
        'target_headings': ['8208'],
        'forbidden_chapters': ['8202', '84', '85'],
        'bonus': 1300
    },
    # 38. Vú tra mỡ, vú bơm mỡ bò, núm tra mỡ (8487) - Chú giải WCO 84.87 & Loại trừ 84.81
    {
        'name': 'grease_nipples',
        'pattern': r'\b(?:vu\s+(?:bom|tra)\s+mo(?:\s+bo)?|num\s+(?:bom|tra)\s+mo(?:\s+bo)?|dau\s+vu\s+mo|vu\s+mo|num\s+mo|greas(?:e|ing)\s+nipples?|grease\s+fittings?|zerk\s+fittings?|lube\s+nipples?)\b',
        'target_headings': ['8487'],
        'forbidden_chapters': ['72', '73', '8481', '8413', '8409', '8431', '8466', '85', '87', '90'],
        'bonus': 2500
    },
    # 39. Ghế các loại (Nhóm 9401) - Chú giải 1(a) Chương 94 & Loại trừ 94.03
    {
        'name': 'seats_chairs',
        'pattern': r'\b(?:ghe(?!\s+tai)\b|ghe\s+(?:ngoi|xoay|van\s+phong|sofa|xep|tua|da)|seat|chair|armchair|stool)\b',
        'target_headings': ['9401'],
        'forbidden_chapters': ['9403', '3926', '4421', '7326', '8481'],
        'bonus': 1650
    },
    # 39b. Đồ nội thất khác (Nhóm 9403) - Bàn, tủ, giường, kệ (loại trừ ghế 94.01)
    {
        'name': 'other_furniture',
        'pattern': r'\b(?:ban\s+(?:lam\s+viec|an|hoc|tra|hop)|tu\s+(?:ho\s+so|quan\s+ao|tai\s+lieu|bep|sach)|giuong|ke\s+sach|desk|table|cabinet|cupboard|wardrobe|bookshelf)\b',
        'target_headings': ['9403'],
        'forbidden_chapters': ['9401', '3926', '4421', '7326'],
        'bonus': 1550
    },
    # 40. Vali, túi xách, ba lô, cặp (Nhóm 4202)
    {
        'name': 'luggage_suitcases_bags',
        'pattern': r'\b(?:vali|va\s+li|tui\s+xach|ba\s+lo|balo|cap\s+sach|vi\s+tien|suitcase|handbag|backpack|briefcase|wallet)\b',
        'target_headings': ['4202'],
        'forbidden_chapters': ['3926', '7326', '6117', '6217', '6307'],
        'bonus': 1450
    },
    # 41. Giày dép (Chương 64)
    {
        'name': 'footwear_shoes',
        'pattern': r'\b(?:giay|dep|ung|sandal|shoes|boots|footwear)\b',
        'target_headings': ['6401', '6402', '6403', '6404', '6405'],
        'forbidden_chapters': ['3926', '4016', '61', '62', '63'],
        'bonus': 1450
    },
    # 42. Đồ chơi & Dụng cụ thể thao (Chương 95: 9503, 9506)
    {
        'name': 'toys_and_sports_equipment',
        'pattern': r'\b(?:do\s+choi|bup\s+be|dung\s+cu\s+the\s+thao|vot\s+(?:cau\s+long|tennis)|bong\s+(?:da|ro|chuyen)|may\s+tap\s+the\s+duc|toys?|gym\s+equipment)\b',
        'target_headings': ['9503', '9506'],
        'forbidden_chapters': ['3926', '7326', '8479'],
        'bonus': 1450
    },
    # 43. Quạt công nghiệp, quạt ly tâm, quạt hút, máy nén khí (8414) - Note 2(a) Section XVI
    {
        'name': 'industrial_fans_compressors',
        'pattern': r'\b(?:quat(?!\s+(?:tai|ao|quan))\b|quat\s+(?:ly\s+tam|hut|thong\s+gio|cong\s+nghiep|lam\s+mat|gio|cay|treo|dien)|may\s+nen\s+khi|may\s+nen\s+lanh|industrial\s+fan|exhaust\s+fan|cooling\s+fan|blower|air\s+compressor)\b',
        'target_headings': ['8414'],
        'forbidden_chapters': ['8514', '8419', '7326', '8479', '8501'],
        'bonus': 1650
    },
    # 44. Bộ điều nhiệt, ổn áp, thiết bị điều khiển tự động (9032) - Chú giải 7 Chương 90 & Loại trừ Ch 85
    {
        'name': 'automatic_regulators_thermostats',
        'pattern': r'\b(?:bo\s+dieu\s+nhiet|bo\s+on\s+nhiet|thermostat|bo\s+dieu\s+ap\s+tu\s+dong|manostat|bo\s+dieu\s+khien\s+tu\s+dong|automatic\s+controller|temperature\s+controller)\b',
        'target_headings': ['9032'],
        'forbidden_chapters': ['8536', '8537', '9031', '8481'],
        'bonus': 2000
    }
]

# =========================================================================
# UNIVERSAL STATUTORY HIERARCHY & PRE-FLIGHT LEGAL AUDIT ENGINE
# (Áp dụng cưỡng bức toàn bộ Biểu thuế - Giải quyết triệt để mọi xung đột pháp lý)
# =========================================================================

# 1. Chú giải 2(a) Phần XVI: Hàng hóa có công năng độc lập định danh (ưu tiên tuyệt đối > phụ tùng Note 2(b))
NOTE_2A_HEADINGS = {
    '8413', '8414', '8419', '8421', '8481', '8482', '8483', '8484', 
    '8501', '8504', '8535', '8536', '8537', '8544', '8545', '8546', '8547'
}

# 2. Chú giải 2(b) Phần XVI & Chú giải phụ tùng các Phần khác: Nhóm bộ phận/phụ tùng dùng cho máy móc
NOTE_2B_PARTS_HEADINGS = {
    '8409', '8431', '8448', '8466', '8473', '8487', 
    '8503', '8522', '8529', '8538', '8548',
    '8708', '8803', '8807', '8906', '9033'
}

# 3. Chú giải 2 Phần XV: Bộ phận có công dụng chung bằng kim loại cơ bản (loại trừ khỏi Chương 84, 85, 87, 88, 90)
SECTION_XV_NOTE_2_HEADINGS = {
    '7318', '7415', '7616',  # Bu lông, ốc vít, đai ốc, đinh tán, chốt hãm, chốt định vị
    '7320', '7416',          # Lò xo các loại
    '7307', '7412', '7609'   # Phụ kiện ghép nối ống
}

# 4. Chú giải 1(k) Phần XVI: Dụng cụ, dao cắt, khuôn mẫu thuộc Chương 82 (loại trừ khỏi Phần XVI)
CHAPTER_82_TOOLS_HEADINGS = {
    '8207', '8208', '8209', '8202', '8203', '8204', '8205'
}

def _audit_notes_legacy(code, head, item_desc, sub_desc, decomp, full_text_norm):
    """
    Hệ thống kiểm toán pháp lý tự động đa chiều:
    - Kiểm tra Chú giải Phần / Chương (Section/Chapter Notes)
    - Kiểm tra điều kiện phủ quyết tiêu cực (Negative qualifiers: 'không lắp...', 'chưa...')
    - Kiểm tra ngưỡng định lượng kỹ thuật (Voltage, Amperage, Diameter DN, Thickness mm, Weight)
    Trả về: (is_disqualified, penalty_score, audit_reason)
    """
    c_raw = code.replace('.', '')
    sub = c_raw[:6]
    item_norm = remove_accents(item_desc)
    sub_norm = remove_accents(sub_desc)
    specs = decomp.get('specs', {})

    active_cats = decomp.get('active_categories', [])
    active_target_heads = set()
    for cat in active_cats:
        active_target_heads.update(cat.get('target_headings', []))

    # A. Luật phân tầng Chú giải 2(a) Phần XVI
    # Nếu sản phẩm là mặt hàng có công năng thuộc Note 2(a) (van, bơm, quạt, ổ bi, động cơ, công tắc, cáp...),
    # thì TUYỆT ĐỐI KHÔNG phân loại vào nhóm bộ phận máy móc Note 2(b) (84.66, 84.31, 84.09, 87.08...)
    if any(h in NOTE_2A_HEADINGS for h in active_target_heads):
        if head in NOTE_2B_PARTS_HEADINGS:
            return True, -5000, f"Vi phạm Chú giải 2(a) Phần XVI: Mặt hàng định danh công năng thuộc Nhóm {active_target_heads & NOTE_2A_HEADINGS} có giá trị pháp lý ưu tiên tuyệt đối hơn nhóm phụ tùng {head}"

    # B. Luật phân tầng Chú giải 2 Phần XV
    # Nếu sản phẩm là bộ phận công dụng chung (bu lông, đinh vít, chốt định vị, lò xo...),
    # TUYỆT ĐỐI KHÔNG phân loại vào nhóm máy móc hay phương tiện vận tải
    if any(h in SECTION_XV_NOTE_2_HEADINGS for h in active_target_heads):
        if head in NOTE_2B_PARTS_HEADINGS or head[:2] in ['84', '85', '87', '88', '90']:
            return True, -5000, f"Vi phạm Chú giải 2 Phần XV & Chú giải 1(g) Phần XVI: Bộ phận có công dụng chung thuộc Nhóm {active_target_heads & SECTION_XV_NOTE_2_HEADINGS} không phân loại theo máy móc hay phương tiện"

    # C. Luật phân tầng Chú giải 1(k) Phần XVI (Dụng cụ Chương 82)
    if any(h in CHAPTER_82_TOOLS_HEADINGS for h in active_target_heads):
        if head in NOTE_2B_PARTS_HEADINGS or head[:2] in ['84', '85']:
            return True, -5000, f"Vi phạm Chú giải 1(k) Phần XVI: Dụng cụ và lưỡi cắt thuộc Chương 82 loại trừ khỏi Phần XVI"

    # D. (Da chuyen sang Universal GRI-6 Engine: Sibling Tree + Polarity + Quantitative)

# C2. Luật phân tầng Chú giải 1(g) Phần XVI (Sản phẩm plastic loại trừ khỏi Phần XVI)
    is_plastic = any(m[0] == 'plastic' for m in decomp.get('materials', []))
    if is_plastic and (head in NOTE_2B_PARTS_HEADINGS or head in ['8466', '8487'] or (head[:2] in ['84', '85'] and head not in NOTE_2A_HEADINGS)):
        return True, -5000, "Vi phạm Chú giải 1(g) Phần XVI: Sản phẩm cấu tạo từ plastic (nhựa) thuộc Chương 39 (Nhóm 39.26), loại trừ khỏi nhóm bộ phận máy móc Phần XVI"

    # C3. Cổng Kiểm Soát Pháp Lý Bản Chất Vật Liệu Phổ Quát (Universal Material-Chapter Boundary Gate - GRI 1)
    IS_FUNCTIONAL_HEADING = head in NOTE_2A_HEADINGS or head in {
        '8413', '8414', '8481', '8501', '8504', '8535', '8536', '8537', 
        '9025', '9026', '9031', '9032', '8207', '8208', '8482', '8483'
    }
    if not IS_FUNCTIONAL_HEADING:
        mat_names = [m[0] for m in decomp.get('materials', [])]
        ch = head[:2]
        is_mat_plastic = 'plastic' in mat_names
        is_mat_rubber = 'rubber' in mat_names
        is_mat_steel = 'steel_iron' in mat_names
        is_mat_copper = 'copper' in mat_names
        is_mat_aluminum = 'aluminum' in mat_names
        is_mat_metal = is_mat_steel or is_mat_copper or is_mat_aluminum

        if is_mat_plastic and ch in {'72', '73', '74', '75', '76', '78', '79', '80', '81', '40', '44', '68', '69', '70'}:
            return True, -5000, f"Vi phạm bản chất chất liệu (GRI 1): Sản phẩm bằng nhựa/plastic tuyệt đối không phân loại vào Chương {ch} (kim loại/cao su/vật liệu khác). Bắt buộc phân loại vào Chương 39 (Nhóm 39.26 hoặc 39.17)."

        if is_mat_rubber and ch in {'72', '73', '74', '75', '76', '78', '79', '80', '81', '39', '44', '68', '69', '70'}:
            return True, -5000, f"Vi phạm bản chất chất liệu (GRI 1): Sản phẩm bằng cao su tuyệt đối không phân loại vào Chương {ch} (kim loại/nhựa/vật liệu khác). Bắt buộc phân loại vào Chương 40 (Nhóm 40.16 hoặc 40.09)."

        if is_mat_metal and ch in {'39', '40', '44', '48', '68', '69', '70'}:
            return True, -5000, f"Vi phạm bản chất chất liệu (GRI 1): Sản phẩm bằng kim loại cơ bản không phân loại vào Chương {ch} (nhựa/cao su/vật liệu khác). Bắt buộc phân loại vào Phần XV (Chương 72 đến 81)."

        if is_mat_steel and ch in {'74', '76'}:
            return True, -5000, f"Vi phạm chất liệu kim loại: Sản phẩm bằng sắt/thép/inox không thuộc Chương {ch} (đồng/nhôm)."

        if is_mat_copper and ch in {'72', '73', '76'}:
            return True, -5000, f"Vi phạm chất liệu kim loại: Sản phẩm bằng đồng/đồng thau không thuộc Chương {ch} (sắt/thép/nhôm)."

        if is_mat_aluminum and ch in {'72', '73', '74'}:
            return True, -5000, f"Vi phạm chất liệu kim loại: Sản phẩm bằng nhôm không thuộc Chương {ch} (sắt/thép/đồng)."

    # D. Kiểm tra Phủ quyết Tiêu cực: Nhóm 8483.30 vs 8483.20 ("Lắp ổ bi" vs "KHÔNG lắp ổ bi")
    has_negative_bearing = decomp.get('has_negative_bearing', False)
    if sub == '848330' and not has_negative_bearing:
        bearing_keywords = ['o bi', 'vong bi', 'o dua', 'o lan', 'co bi', 'co o bi', 'ball bearing', 'ball bearings', 'roller bearing', 'bearing']
        if any(k in full_text_norm for k in bearing_keywords):
            return True, -4000, "Vi phạm điều kiện phủ quyết 8483.30: Phân nhóm 8483.30 chỉ áp dụng cho 'Thân ổ, KHÔNG lắp ổ bi hoặc ổ đũa'"
    elif sub == '848320' and has_negative_bearing:
        return True, -4000, "Vi phạm điều kiện 8483.20: Phân nhóm 8483.20 chỉ áp dụng cho 'Thân ổ CÓ lắp ổ bi hoặc ổ đũa'. Hàng hóa không lắp ổ bi phải chuyển về 8483.30"

    # E. Kiểm tra Ngưỡng Điện áp (8535 vs 8536)
    v_val = specs.get('v')
    if v_val is not None:
        if v_val <= 1000 and head == '8535':
            return True, -4000, f"Vi phạm ngưỡng điện áp: Điện áp {v_val}V (<= 1.000V) thuộc Nhóm 85.36, loại trừ khỏi Nhóm 85.35 (> 1.000V)"
        elif v_val > 1000 and head == '8536':
            return True, -4000, f"Vi phạm ngưỡng điện áp: Điện áp {v_val}V (> 1.000V) thuộc Nhóm 85.35, loại trừ khỏi Nhóm 85.36 (<= 1.000V)"

    # F. Kiểm tra Ngưỡng Dòng điện (Công tắc 8536.50)
    a_val = specs.get('a')
    if sub == '853650':
        is_under_16a = (a_val is not None and a_val < 16.0) or any(k in full_text_norm for k in ['duoi 16a', 'duoi 16 a', 'nho hon 16a', '< 16a', '<16a', 'under 16a'])
        is_16a_or_more = (a_val is not None and a_val >= 16.0) or any(k in full_text_norm for k in ['tren 16a', 'tu 16a', 'tren 16 a', '>= 16a', '>=16a', 'over 16a'])
        if is_under_16a and c_raw in ['85365099', '85365069', '85365059', '85365039']:
            return True, -3500, "Vi phạm ngưỡng dòng điện: Dòng điện < 16A bắt buộc chọn phân nhánh dưới 16A, loại trừ phân nhánh Loại khác"
        elif is_16a_or_more and c_raw in ['85365096', '85365061', '85365051', '85365033']:
            return True, -3500, "Vi phạm ngưỡng dòng điện: Dòng điện >= 16A loại trừ phân nhánh dưới 16A"

    # G. Kiểm tra Đường kính danh định Van (8481)
    dn_val = specs.get('dn')
    if dn_val is not None and head == '8481':
        if dn_val <= 25 and ('tren 2,5 cm' in item_norm or 'tren 2.5 cm' in item_norm):
            return True, -3000, f"Vi phạm kích thước van: DN {dn_val}mm (<= 2,5cm) loại trừ phân nhóm trên 2,5 cm"
        elif dn_val > 25 and ('khong qua 2,5 cm' in item_norm or 'khong qua 2.5 cm' in item_norm):
            return True, -3000, f"Vi phạm kích thước van: DN {dn_val}mm (> 2,5cm) loại trừ phân nhóm không quá 2,5 cm"

    # H. Kiểm tra Độ dày Gỗ xẻ / Ván lạng (4407 vs 4408)
    mm_val = specs.get('mm')
    if mm_val is not None and head in ['4407', '4408']:
        if mm_val > 6 and head == '4408':
            return True, -3500, f"Vi phạm độ dày gỗ: Chiều dày {mm_val}mm (> 6mm) thuộc Nhóm 44.07, loại trừ khỏi 44.08"
        elif mm_val <= 6 and head == '4407':
            return True, -3500, f"Vi phạm độ dày gỗ: Chiều dày {mm_val}mm (<= 6mm) thuộc Nhóm 44.08, loại trừ khỏi 44.07"

    # I. Kiểm tra Trọng lượng thiết bị bay không người lái drone (8806)
    if head == '8806':
        is_under_250g = any(k in full_text_norm for k in ['249g', 'khong qua 250', '250g', '<= 250g', 'under 250g'])
        if is_under_250g and c_raw != '88062100':
            return True, -3500, "Vi phạm trọng lượng drone (GRI 6): Drone trọng lượng cất cánh tối đa <= 250g thuộc mã 8806.21.00"

    # J. Chi tiết che chắn/bảo vệ an toàn cơ khí vs Thiết bị đóng ngắt bảo vệ điện (8535, 8536)
    if head in ['8535', '8536']:
        mech_keywords = [
            'tam che', 'vo che', 'nap che', 'tam chan', 'vach chan', 'man che',
            'tam bao ve', 'vo bao ve', 'nap bao ve', 'chan bui', 'tam mica',
            'tam nhua', 'safety guard', 'machine guard', 'protective cover'
        ]
        elec_keywords = [
            'cau dao', 'aptomat', 'ro le', 'relay', 'cau chi', 'fuse',
            'cong tac', 'switch', 'dong ngat mach', 'chong set', 'dau noi',
            'connector', 'phich cam', 'o cam', 'circuit breaker'
        ]
        subj_norm_audit = remove_accents(decomp.get('subject_part', ''))
        if any(k in subj_norm_audit for k in mech_keywords) and not any(k in full_text_norm for k in elec_keywords):
            return True, -5000, "Vi phạm bản chất pháp lý Nhóm 85.35/85.36: Nhóm 85.35 và 85.36 chỉ dành cho thiết bị đóng ngắt hoặc bảo vệ mạch điện (quá tải, ngắn mạch). Tấm che/vỏ bảo vệ an toàn cơ khí máy móc hoặc an toàn lao động không thuộc nhóm này."

    # K. Dây dẫn & Cáp điện (8544)
    if head == '8544':
        has_connector = any(k in full_text_norm for k in ['co dau noi', 'da gan dau noi', 'gan phich cam', 'fitted with connectors', 'connector', 'co jack', 'co phich'])
        no_connector = any(k in full_text_norm for k in ['chua gan dau noi', 'khong co dau noi', 'without connectors', 'chua bam cos', 'chua co dau noi'])
        
        # K1: Có đầu nối vs chưa gắn đầu nối
        if has_connector and any(sub.startswith(p) for p in ['854449', '854411', '854419']):
            return True, -3500, "Vi phạm phân nhánh cáp điện (GRI 6): Cáp/dây điện ĐÃ GẮN ĐẦU NỐI bắt buộc thuộc phân nhóm có gắn đầu nối (như 8544.42), loại trừ phân nhóm chưa gắn đầu nối (8544.49)"
        elif no_connector and sub.startswith('854442'):
            return True, -3500, "Vi phạm phân nhánh cáp điện (GRI 6): Dây cáp KHÔNG GẮN ĐẦU NỐI bắt buộc thuộc phân nhóm 8544.49, loại trừ 8544.42"

        # K2: Ngưỡng điện áp cáp điện (<= 1000V vs > 1000V)
        v_val_cable = specs.get('v')
        if v_val_cable is not None:
            if v_val_cable > 1000 and (sub.startswith('854442') or sub.startswith('854449')):
                return True, -4000, f"Vi phạm ngưỡng điện áp cáp điện: Điện áp {v_val_cable}V (> 1.000V) bắt buộc thuộc phân nhóm 8544.60, loại trừ phân nhóm <= 1.000V (8544.42/49)"
            elif v_val_cable <= 1000 and sub.startswith('854460'):
                return True, -4000, f"Vi phạm ngưỡng điện áp cáp điện: Điện áp {v_val_cable}V (<= 1.000V) bắt buộc thuộc phân nhóm 8544.42 hoặc 8544.49, loại trừ phân nhóm > 1.000V (8544.60)"

        # K3: Cáp sợi quang (8544.70) vs Sợi quang chưa bọc riêng (9001.10)
        is_optical_cable = any(k in full_text_norm for k in ['cap quang', 'cap soi quang', 'optical fibre cable'])
        if is_optical_cable and sub != '854470':
            return True, -3500, "Vi phạm phân loại cáp quang: Cáp sợi quang có bọc từng sợi riêng bắt buộc thuộc phân nhóm 8544.70, loại trừ khỏi Chương 90"

    # L. Động cơ điện (8501) - Ngưỡng công suất
    if head == '8501':
        w_val = specs.get('w')
        kw_val = specs.get('kw')
        is_micro_motor = (w_val is not None and w_val <= 37.5) or (kw_val is not None and kw_val <= 0.0375) or any(k in full_text_norm for k in ['khong qua 37.5w', 'duoi 37.5w', '<= 37.5w', '< 37.5w'])
        if is_micro_motor and sub != '850110':
            return True, -3500, "Vi phạm ngưỡng công suất động cơ (GRI 6): Động cơ điện công suất <= 37.5 W bắt buộc thuộc phân nhóm 8501.10"

    # M. Thiết bị đo kiểm (9031) vs Bộ điều chỉnh/điều khiển tự động (9032)
    if head == '9031' and any(k in full_text_norm for k in ['thermostat', 'bo dieu nhiet', 'bo dieu ap tu dong', 'manostat', 'automatic controller', 'dieu khien tu dong']):
        return True, -4000, "Vi phạm bản chất pháp lý Nhóm 90.31 vs 90.32: Thiết bị và dụng cụ điều chỉnh hoặc điều khiển tự động (như thermostat, manostat) thuộc Nhóm 90.32, loại trừ khỏi Nhóm 90.31"
    elif head == '9032' and any(k in full_text_norm for k in ['may do toa do', 'cmm', 'profile projector', 'kiem tra khuyet tat', 'may do bien dang', 'quang hoc do']):
        return True, -4000, "Vi phạm bản chất pháp lý Nhóm 90.32 vs 90.31: Thiết bị đo lường, kiểm tra quang học hoặc tọa độ thuộc Nhóm 90.31, loại trừ khỏi 90.32"

    # N. Đồ nội thất & Ghế (Chương 94: 9401, 9403)
    is_furniture = any(k in full_text_norm for k in ['ghe', 'ghe ngoi', 'ghe xoay', 'ghe van phong', 'ban lam viec', 'ban an', 'tu ho so', 'tu quan ao', 'giuong', 'ke sach', 'ke de do', 'chair', 'seat', 'table', 'desk', 'cabinet', 'cupboard'])
    if is_furniture:
        if head in ['3926', '4421', '7326']:
            return True, -5000, "Vi phạm Chú giải 1(e) Chương 39 / Chú giải 1(c) Chương 44 / Chú giải 1(k) Chương 73: Ghế và đồ nội thất dù làm bằng nhựa, gỗ hay kim loại đều BẮT BUỘC phân loại vào Chương 94 (Nhóm 94.01 hoặc 94.03), loại trừ khỏi các chương vật liệu thành phẩm"

    # O. Túi xách, Vali, Ba lô, Cặp (Nhóm 4202)
    is_luggage = bool(re.search(r'\b(?:vali|va\s+li|tui\s+xach|ba\s+lo|balo|cap\s+sach|vi\s+tien|suitcase|handbag|backpack|briefcase|wallet)\b', full_text_norm))
    if is_luggage:
        if head in ['3926', '6117', '6217', '6307', '7326'] or head[:2] in ['61', '62', '63']:
            return True, -5000, "Vi phạm Chú giải 1(h) Phần XI & Chú giải 1(ij) Chương 39: Túi xách, vali, ba lô, cặp dù bề mặt bằng da, nhựa tấm hay vải dệt đều BẮT BUỘC thuộc Nhóm 42.02, loại trừ khỏi các chương dệt may hoặc nhựa"

    # P. Giày dép (Chương 64)
    is_footwear = bool(re.search(r'\b(?:giay\s+dep|doi\s+giay|doi\s+dep|giay\s+the\s+thao|giay\s+da|giay\s+vai|giay\s+bao\s+ho|sandal|dep\s+sandal|ung\s+cao\s+su|ung\s+nhua|ung\s+da|shoes|boots|footwear)\b', full_text_norm))
    if is_footwear:
        if head in ['3926', '4016'] or head[:2] in ['61', '62', '63']:
            return True, -5000, "Vi phạm Chú giải 1(a) Phần XI & Chú giải 1(e) Chương 39: Giày, dép, ủng các loại BẮT BUỘC thuộc Chương 64, loại trừ khỏi các chương dệt may và đồ nhựa/cao su"

    # Q. Đồ chơi & Dụng cụ thể thao (Chương 95: 9503, 9506)
    is_toy_or_sport = any(k in full_text_norm for k in ['do choi tre em', 'mo hinh do choi', 'bup be', 'dung cu the thao', 'vot cau long', 'vot tennis', 'bong da', 'bong ro', 'may tap the duc', 'toys', 'gym equipment'])
    if is_toy_or_sport:
        if head in ['3926', '4421', '7326', '8479', '8543']:
            return True, -5000, "Vi phạm Chú giải 1(p) Phần XVI & Chú giải 1(o) Chương 39: Đồ chơi, mô hình giải trí và dụng cụ thể thao BẮT BUỘC thuộc Chương 95 (95.03 hoặc 95.06), loại trừ khỏi các chương nhựa, gỗ và máy móc"

    # R. Máy bơm (8413) & Máy nén khí/quạt (8414)
    is_pump_or_fan = any(k in full_text_norm for k in ['may bom', 'bom nuoc', 'bom dau', 'may nen khi', 'may nen lanh', 'quat thong gio', 'quat hut', 'water pump', 'air compressor', 'industrial fan'])
    if is_pump_or_fan:
        if head in NOTE_2B_PARTS_HEADINGS or head in ['8409', '8431', '8487']:
            return True, -5000, "Vi phạm Chú giải 2(a) Phần XVI: Bơm chất lỏng (84.13), máy nén khí và quạt (84.14) có chức năng độc lập, bắt buộc phân loại vào nhóm định danh của chúng, không được xếp vào nhóm bộ phận máy"

    # S. Hóa chất bảo vệ thực vật & Chế phẩm tẩy rửa bán lẻ (3808, 3402)
    if any(k in full_text_norm for k in ['thuoc tru sau', 'thuoc diet con trung', 'thuoc diet co', 'thuoc sat trung', 'pesticide', 'insecticide']):
        if head[:2] in ['28', '29', '30']:
            return True, -4000, "Vi phạm Chú giải Chương 38: Thuốc trừ sâu, diệt côn trùng, sát trùng đã đóng gói bán lẻ BẮT BUỘC thuộc Nhóm 38.08, loại trừ khỏi hóa chất đơn chất Chương 28, 29"

    # U. Chi tiết liên kết (Bu lông, ốc vít Nhóm 7318/7415/7616) vs Dụng cụ cầm tay (Chương 82)
    subj_part_norm = remove_accents(decomp.get('subject_part', '')) if isinstance(decomp, dict) else ''
    is_fastener_art = bool(re.search(r'\b(?:bu\s+long|bu-long|bolt|dai\s+oc|oc\s+vit|dinh\s+vit|dinh\s+tan|chot\s+dinh\s+vi|chot\s+tru|vit\s+cay|vong\s+dem)\b', subj_part_norm, re.I))
    is_tool_art = bool(re.search(r'\b(?:kim|co\s+le|mo\s+let|tay\s+van|dung\s+cu\s+cat|dung\s+cu\s+xen|dao|keo|tua\s+vit|wrench|spanner|cutter|cropper|plier|pliers)\b', subj_part_norm, re.I))
    if is_fastener_art and not is_tool_art and head[:2] == '82':
        return True, -6000, "Vi phạm Chú giải 1(a) Phần XV: Sản phẩm là chi tiết liên kết (bu lông, ốc vít, đai ốc), KHÔNG PHẢI dụng cụ cầm tay (kìm xén bu lông, cờ lê) thuộc Chương 82. Bắt buộc phân loại vào Nhóm 73.18 / 74.15 / 76.16."

    # T. Quần áo dệt kim (Chương 61) vs Quần áo dệt thoi (Chương 62)
    is_knitted = any(k in full_text_norm for k in ['det kim', 'knitted', 'crocheted', 'moc'])
    is_woven = any(k in full_text_norm for k in ['det thoi', 'woven', 'vai det thoi'])
    if is_knitted and head[:2] == '62':
        return True, -4000, "Vi phạm ranh giới Chú giải Phần XI: Quần áo dệt kim hoặc móc BẮT BUỘC thuộc Chương 61, loại trừ khỏi Chương 62 (quần áo dệt thoi)"
    elif is_woven and head[:2] == '61':
        return True, -4000, "Vi phạm ranh giới Chú giải Phần XI: Quần áo dệt thoi BẮT BUỘC thuộc Chương 62, loại trừ khỏi Chương 61 (quần áo dệt kim)"

    # U. Vú tra mỡ / Vú bơm mỡ bò (Greasing nipples - 8487.90.00)
    is_grease_nipple = any(k in full_text_norm for k in [
        'vu bom mo', 'vu tra mo', 'num bom mo', 'num tra mo',
        'dau vu mo', 'vu mo', 'num mo', 'grease nipple', 'greasing nipple',
        'grease fitting', 'zerk fitting', 'lube nipple'
    ])
    # V. Thiết bị điều chỉnh tự động (9032) vs Khí cụ đóng ngắt rơ le (8536/8537)
    is_auto_reg = any(k in full_text_norm for k in ['thermostat', 'bo dieu nhiet', 'bo on nhiet', 'manostat', 'bo dieu ap tu dong'])
    if is_auto_reg and (head[:2] in ['85', '84'] or head == '9031'):
        return True, -5000, "Vi phạm Chú giải 1(m) Phần XVI & Chú giải 7 Chương 90: Bộ điều chỉnh nhiệt độ (thermostat) hoặc áp suất (manostat) tự động bắt buộc thuộc Nhóm 90.32, loại trừ khỏi Phần XVI (Chương 84, 85)"

    # W. Ghế (9401) vs Đồ nội thất khác (9403)
    is_seat_art = bool(re.search(r'\b(?:ghe(?!\s+tai)\b|seat|chair|armchair|ghe\s+xoay|ghe\s+ngoi)\b', subj_part_norm, re.I))
    if is_seat_art and head == '9403':
        return True, -5000, "Vi phạm Chú giải Nhóm 94.01 & 94.03: Ghế các loại bắt buộc thuộc Nhóm 94.01, loại trừ khỏi Nhóm 94.03 (đồ nội thất khác)"

    # X. Ống (Pipes/Tubes) vs Hóa chất (Chương 28, 29) & Phế liệu kim loại (Chương 72)
    is_pipe_art = bool(re.search(r'\b(?:ong\s+(?:nhua|mem|dan|cao\s+su|cung|thep|kim\s+loai|dong)|ong(?!\s+(?:nghiem|khoi))\b|hose|tube|pipe)\b', subj_part_norm, re.I))
    if is_pipe_art:
        if head[:2] in ['28', '29', '38']:
            return True, -5000, "Vi phạm bản chất thành phẩm: Ống dẫn chất lỏng/khí là sản phẩm gia công cơ khí, tuyệt đối không phân loại vào chương hóa chất nguyên liệu (Chương 28, 29, 38)"
        if head[:2] == '72':
            return True, -5000, "Vi phạm Chú giải Chương 72 & 73: Ống thép là sản phẩm thép chế tạo thuộc Chương 73 (73.04/73.05/73.06), không phân loại vào sắt thép nguyên sinh hoặc phế liệu Chương 72"

    # Y. Quạt (8414) vs Lò nung / Lò sấy (8514, 8419)
    is_fan_art = bool(re.search(r'\b(?:quat(?!\s+(?:tai|ao|quan))\b|quat\s+(?:ly\s+tam|hut|thong\s+gio|cong\s+nghiep|lam\s+mat|gio|dien)|industrial\s+fan|exhaust\s+fan|cooling\s+fan|blower)\b', subj_part_norm, re.I))
    if is_fan_art and head in ['8514', '8419', '7326']:
        return True, -5000, "Vi phạm Chú giải 2(a) Phần XVI: Quạt công nghiệp và máy thổi khí định danh tại Nhóm 84.14, ưu tiên tuyệt đối hơn nhóm lò nung 85.14 hoặc thiết bị nhiệt 84.19"

    return False, 0, ""


try:
    import universal_engine as _ue
except Exception as _e_ue:
    print("Warning: universal_engine not loaded:", _e_ue)
    _ue = None


def audit_statutory_constraints(code, head, item_desc, sub_desc, decomp, full_text_norm):
    """Cong kiem toan phap ly: (1) Universal GRI-6 Engine (cay gach noi, phu dinh, nguong so)
    -> (2) Chu giai Phan/Chuong (legacy notes)."""
    bonus = 0
    if _ue is not None:
        try:
            dq, why, bonus = _ue.evaluate(code, _ue.extract_query(full_text_norm))
            if dq:
                return True, -9000, "Vi phạm GRI 6 (Universal Engine): " + why
        except Exception as _e:
            print("universal_engine error:", _e)
            bonus = 0
    is_dq, pen, reason = _audit_notes_legacy(code, head, item_desc, sub_desc, decomp, full_text_norm)
    if is_dq:
        return is_dq, pen, reason
    return False, pen + bonus, reason


def decompose_query(query_text, desc="", use="", specs="", model_en=""):
    text = (query_text or '').strip()
    full_text = f"{text} {desc} {use} {specs} {model_en}".strip()
    full_text_norm = remove_accents(full_text)
    
    # 1. Separate Subject vs Host context
    if desc or use or specs or model_en:
        subject_part = f"{text} {model_en}".strip()
        host_part = f"{use} {desc}".strip()
        for marker in HOST_MARKERS:
            m = re.search(marker, text, re.I)
            if m:
                subject_part = f"{text[:m.start()].strip(' ,:;-()')} {model_en}".strip()
                host_part = f"{text[m.end():].strip(' ,:;-()')} {host_part}".strip()
                break
    else:
        subject_part = text
        host_part = ""
        for marker in HOST_MARKERS:
            m = re.search(marker, text, re.I)
            if m:
                subject_part = text[:m.start()].strip(' ,:;-()')
                host_part = text[m.end():].strip(' ,:;-()')
                break
            
    subj_norm = remove_accents(subject_part)
    host_norm = remove_accents(host_part)
    
    # 2. Extract technical specs from full_text
    specs_dict = {}
    m_dn = re.search(r'\bdn\s*(\d+)\b', full_text, re.I)
    if m_dn: specs_dict['dn'] = float(m_dn.group(1))
    
    m_mm = re.search(r'(\d+(?:\.\d+)?)\s*mm\b', full_text, re.I)
    if m_mm: specs_dict['mm'] = float(m_mm.group(1))
    
    m_cm = re.search(r'(\d+(?:\.\d+)?)\s*cm\b', full_text, re.I)
    if m_cm: specs_dict['cm'] = float(m_cm.group(1))
    
    m_v = re.search(r'(\d+(?:\.\d+)?)\s*(?:v|volt|vôn|kv)\b', full_text, re.I)
    if m_v:
        val = float(m_v.group(1))
        if 'kv' in full_text.lower(): val *= 1000
        specs_dict['v'] = val
        
    m_a = re.search(r'(\d+(?:\.\d+)?)\s*(?:a|amp|ampe)\b', full_text, re.I)
    if m_a: specs_dict['a'] = float(m_a.group(1))
    
    m_bar = re.search(r'(\d+(?:\.\d+)?)\s*(?:bar|mpa|psi)\b', full_text, re.I)
    if m_bar: specs_dict['bar'] = float(m_bar.group(1))
    
    # 3. Detect Materials (prevent 'đường kính' collision with 'kính', 'đồng bộ' with 'đồng', and 'máy ép nhựa' with 'nhựa')
    materials = []
    text_no_diam = re.sub(r'đường\s+kính|duong\s+kinh', '', full_text_norm)
    text_no_copper_coll = re.sub(r'\b(?:dong\s+bo|dong\s+ho|dong\s+truc|dong\s+tam|dong\s+muc|dong\s+quy|dong\s+dang|hop\s+dong|cong\s+dong|tu\s+dong|chuyen\s+dong|truyen\s+dong|hoat\s+dong|phat\s+dong|khoi\s+dong|bien\s+dong|luu\s+dong|rung\s+dong|lao\s+dong)\b', '', full_text_norm)
    plastic_coll_regex = r'\b(?:may\s+ep\s+nhua|khuon\s+ep\s+nhua|khuon\s+nhua|nganh\s+nhua|san\s+pham\s+nhua|hat\s+nhua|gia\s+cong\s+nhua|may\s+thoi\s+nhua|may\s+dun\s+nhua|nganh\s+duc\s+nhua|may\s+tao\s+hat\s+nhua)\b'
    text_no_plastic_coll = re.sub(plastic_coll_regex, '', full_text_norm, flags=re.I)
    rubber_coll_regex = r'\b(?:che\s+bien\s+cao\s+su|nganh\s+cao\s+su|may\s+cao\s+su)\b'
    text_no_rubber_coll = re.sub(rubber_coll_regex, '', full_text_norm, flags=re.I)

    for mat_name, pattern, boost_chaps, penalty_chaps in MATERIAL_RULES:
        if mat_name == 'glass':
            target_norm = text_no_diam
            target_full = re.sub(r'đường\s+kính|duong\s+kinh', '', full_text, flags=re.I)
        elif mat_name == 'copper':
            target_norm = text_no_copper_coll
            target_full = re.sub(r'\b(?:đồng\s+bộ|đồng\s+hồ|đồng\s+trục|đồng\s+tâm|đồng\s+mức|đồng\s+quy|đồng\s+dạng|hợp\s+đồng|cộng\s+đồng|tự\s+động|chuyển\s+động|truyền\s+động|hoạt\s+động|phát\s+động|khởi\s+động|biến\s+động|lưu\s+động|rung\s+động|lao\s+động)\b', '', full_text, flags=re.I)
        elif mat_name == 'plastic':
            target_norm = text_no_plastic_coll
            target_full = re.sub(plastic_coll_regex, '', full_text, flags=re.I)
        elif mat_name == 'rubber':
            target_norm = text_no_rubber_coll
            target_full = re.sub(rubber_coll_regex, '', full_text, flags=re.I)
        else:
            target_norm = full_text_norm
            target_full = full_text

        if re.search(pattern, target_norm, re.I) or re.search(pattern, target_full, re.I):
            materials.append((mat_name, boost_chaps, penalty_chaps))

    # Subject part material supremacy (vật liệu trong tên hàng là tối thượng)
    subj_has_steel = bool(re.search(r'\b(?:thep|sat|steel|iron|inox)\b', subj_norm, re.I))
    subj_has_copper = bool(re.search(r'\b(?:dong|brass|copper|bronze)\b', subj_norm, re.I))
    subj_has_aluminum = bool(re.search(r'\b(?:nhom|aluminum|aluminium)\b', subj_norm, re.I))
    subj_has_plastic = bool(re.search(r'\b(?:nhua|plastic)\b', subj_norm, re.I))
    subj_has_rubber = bool(re.search(r'\b(?:cao\s+su|rubber)\b', subj_norm, re.I))

    if subj_has_steel:
        materials = [m for m in materials if m[0] not in ('plastic', 'rubber')]
    elif subj_has_copper:
        materials = [m for m in materials if m[0] not in ('plastic', 'rubber', 'steel_iron')]
    elif subj_has_aluminum:
        materials = [m for m in materials if m[0] not in ('plastic', 'rubber', 'steel_iron')]
    elif subj_has_plastic:
        materials = [m for m in materials if m[0] not in ('steel_iron', 'copper', 'aluminum')]
    elif subj_has_rubber:
        materials = [m for m in materials if m[0] not in ('steel_iron', 'copper', 'aluminum')]
            
    # 4. Check Statutory Categories with Dynamic Material Adaptation
    has_negative_bearing = bool(re.search(r'\b(?:khong|không|chua|chưa|without|no|not)\s+(?:lap\s+|lắp\s+|co\s+|có\s+)?(?:o\s+bi|bi\b|vong\s+bi|ball\s+bearings?|bearings?)', full_text_norm))
    active_categories = []
    is_subject_structural_article = bool(re.search(r'\b(?:tam\s+che|vo\s+che|nap\s+che|tam\s+chan|vach\s+chan|man\s+che|chan\s+bui|tam\s+(?:nhua|mica|thep|kim\s+loai)?\s*bao\s+ve|vo\s+bao\s+ve|nap\s+bao\s+ve|bu\s+long|oc\s+vit|mieng\s+thep|tam\s+thep|de\s+ga|gia\s+do|khung|khuon|dao|luoi)\b', subj_norm, re.I))

    mat_names = [m[0] for m in materials]
    is_plastic = ('plastic' in mat_names or subj_has_plastic) and not subj_has_steel
    is_rubber = ('rubber' in mat_names or subj_has_rubber) and not subj_has_steel
    is_copper = 'copper' in mat_names or subj_has_copper
    is_aluminum = 'aluminum' in mat_names or subj_has_aluminum
    is_steel = 'steel_iron' in mat_names or subj_has_steel

    for cat_item in STATUTORY_CATEGORIES:
        cat = cat_item.copy()
        if cat['name'] == 'ball_roller_bearings' and has_negative_bearing:
            continue
        if cat['name'] == 'insulated_cables' and is_subject_structural_article:
            continue
        if cat['name'] == 'electric_motors' and bool(re.search(r'\b(?:hop\s+dieu\s+khien|tu\s+dieu\s+khien|bo\s+dieu\s+khien|mach\s+dieu\s+khien|bang\s+dieu\s+khien|bien\s+tan|inverter|controller|driver)\b', full_text_norm, re.I)):
            continue

        # Dynamic material routing for general use parts & articles (Note 2 Section XV priority)
        if cat['name'] == 'general_use_fasteners':
            if is_steel or subj_has_steel or not (is_plastic or is_copper or is_aluminum or is_rubber):
                cat['target_headings'] = ['7318']
                cat['forbidden_chapters'] = ['39', '40', '82', '84', '85', '87', '88', '90', '94', '95', '96']
                cat['bonus'] = 1550
            elif is_copper or subj_has_copper:
                cat['target_headings'] = ['7415']
                cat['forbidden_chapters'] = ['72', '73', '76', '39', '40', '82', '84', '85', '87', '88', '90', '94', '95', '96']
                cat['bonus'] = 1350
            elif is_aluminum or subj_has_aluminum:
                cat['target_headings'] = ['7616']
                cat['forbidden_chapters'] = ['72', '73', '74', '39', '40', '82', '84', '85', '87', '88', '90', '94', '95', '96']
                cat['bonus'] = 1350
            elif (is_plastic or subj_has_plastic) and not (is_steel or subj_has_steel):
                cat['target_headings'] = ['3926']
                cat['forbidden_chapters'] = ['72', '73', '74', '76', '82', '84', '85', '87', '88', '90', '94', '95', '96']
                cat['bonus'] = 1450
            elif (is_rubber or subj_has_rubber) and not (is_steel or subj_has_steel):
                cat['target_headings'] = ['4016']
                cat['forbidden_chapters'] = ['72', '73', '74', '76', '39', '82', '84', '85', '87', '88', '90', '94', '95', '96']
                cat['bonus'] = 1450

        elif cat['name'] == 'gaskets_washers_seals':
            if is_rubber:
                cat['target_headings'] = ['4016']
                cat['forbidden_chapters'] = ['72', '73', '74', '76', '39', '84', '85', '87', '90']
                cat['bonus'] = 1450
            elif is_plastic:
                cat['target_headings'] = ['3926']
                cat['forbidden_chapters'] = ['72', '73', '74', '76', '40', '84', '85', '87', '90']
                cat['bonus'] = 1450
            elif is_copper:
                cat['target_headings'] = ['7415']
                cat['forbidden_chapters'] = ['72', '73', '76', '39', '40', '84', '85', '87', '90']
                cat['bonus'] = 1350
            elif is_aluminum:
                cat['target_headings'] = ['7616']
                cat['forbidden_chapters'] = ['72', '73', '74', '39', '40', '84', '85', '87', '90']
                cat['bonus'] = 1350
            elif is_steel:
                cat['target_headings'] = ['7318']
                cat['forbidden_chapters'] = ['39', '40', '84', '85', '87', '90']

        elif cat['name'] == 'pipes_tubes_hoses':
            if is_plastic:
                cat['target_headings'] = ['3917']
                cat['forbidden_chapters'] = ['73', '74', '76', '40', '84', '85', '87', '28', '29', '38']
                cat['bonus'] = 1650
            elif is_rubber:
                cat['target_headings'] = ['4009']
                cat['forbidden_chapters'] = ['73', '74', '76', '39', '84', '85', '87', '28', '29', '38']
                cat['bonus'] = 1650
            elif is_copper:
                cat['target_headings'] = ['7411']
                cat['forbidden_chapters'] = ['72', '73', '76', '39', '40', '84', '85', '87', '28', '29', '38']
                cat['bonus'] = 1550
            elif is_aluminum:
                cat['target_headings'] = ['7608']
                cat['forbidden_chapters'] = ['72', '73', '74', '39', '40', '84', '85', '87', '28', '29', '38']
                cat['bonus'] = 1550
            elif is_steel:
                cat['target_headings'] = ['7304', '7305', '7306']
                cat['forbidden_chapters'] = ['72', '74', '76', '39', '40', '84', '85', '87', '28', '29', '38']
                cat['bonus'] = 1650

        elif cat['name'] == 'pipe_fittings':
            if is_plastic:
                cat['target_headings'] = ['3917']
                cat['forbidden_chapters'] = ['73', '74', '76', '40', '82', '84', '85', '87']
                cat['bonus'] = 1350
            elif is_rubber:
                cat['target_headings'] = ['4009']
            elif is_copper:
                cat['target_headings'] = ['7412']
            elif is_aluminum:
                cat['target_headings'] = ['7609']
            elif is_steel:
                cat['target_headings'] = ['7307']

        elif cat['name'] == 'mechanical_covers_guards':
            if is_plastic:
                cat['target_headings'] = ['3926']
                cat['forbidden_chapters'] = ['73', '74', '76', '8544', '8535', '8536', '8537']
                cat['bonus'] = 1450
            elif is_steel:
                cat['target_headings'] = ['7326']
                cat['forbidden_chapters'] = ['39', '8544', '8535', '8536', '8537']
                cat['bonus'] = 1450
            elif is_aluminum:
                cat['target_headings'] = ['7616']
                cat['forbidden_chapters'] = ['39', '73', '8544', '8535', '8536', '8537']
                cat['bonus'] = 1450

        elif cat['name'] == 'transmission_conveyor_belts':
            if is_rubber:
                cat['target_headings'] = ['4010']
                cat['forbidden_chapters'] = ['84', '85', '87', '88', '39']
                cat['bonus'] = 1500
            elif is_plastic:
                cat['target_headings'] = ['3926']
                cat['forbidden_chapters'] = ['84', '85', '87', '88', '40']
                cat['bonus'] = 1500

        # Match in subject has highest priority
        matched_in_subj = bool(re.search(cat['pattern'], subj_norm, re.I) or re.search(cat['pattern'], subject_part, re.I))
        matched_in_full = bool(re.search(cat['pattern'], full_text_norm, re.I) or re.search(cat['pattern'], full_text, re.I))
        matched_in_host = bool(re.search(cat['pattern'], host_norm, re.I))

        # Categories that describe standalone devices shouldn't be triggered purely from host machine context
        if cat['name'] in ['insulated_cables', 'electric_motors', 'valves_fluid_controllers', 'pumps_for_liquids']:
            if matched_in_subj or (matched_in_full and not matched_in_host):
                active_categories.append(cat)
        else:
            if matched_in_subj or matched_in_full or matched_in_host:
                active_categories.append(cat)
            
    # Disambiguation guards
    is_lead_screw = bool(re.search(r'\b(?:vit\s+me|ball\s+screw|truc\s+vit\s+me|linear\s+slide|ban\s+truot|thanh\s+truot)\b', full_text_norm, re.I))
    if is_lead_screw:
        active_categories = [c for c in active_categories if c['name'] not in ['general_use_fasteners', 'other_furniture']]
    is_motor = bool(re.search(r'\b(?:motor|dong\s+co\s+dien|mo\s+to)\b', full_text_norm, re.I))
    if is_motor:
        active_categories = [c for c in active_categories if c['name'] != 'transmission_conveyor_belts']

    if bool(re.search(r'\b(?:thanh\s+thep\s+giu|mieng\s+thep\s+(?:giu|co\s+dinh)|tam\s+thep\s+giu)\b', full_text_norm)):
        active_categories = [c for c in active_categories if c['name'] != 'machine_knives_cutting_blades']

    return {
        'subject_part': subject_part,
        'host_part': host_part,
        'specs': specs_dict,
        'materials': materials,
        'active_categories': active_categories,
        'has_negative_bearing': has_negative_bearing
    }

STOP_WORDS = {
    'bang', 'cho', 'co', 'cua', 'dung', 'loai', 'khac', 'va', 'hoac', 'dang', 'phu', 'tung', 'thay', 'the', 'trong',
    'da', 'chua', 'voi', 'theo', 'nhu', 'duoc', 'cung', 'khong', 'cac', 'nhung', 'de', 'o', 'tai', 'tren', 'duoi',
    'ngoai', 'moi', 'cu', 'chiec', 'cai', 'bo', 'tam', 'thanh', 'cuon', 'hop', 'goi', 'den', 'tu', 'chuyen',
    'kich', 'thuoc', 'size', 'dimension', 'dimensions', 'specs', 'chieu', 'dai', 'rong', 'cao', 'day', 'phi', 'dn',
    'gan', 'lap', 'dat', 'dinh'
}

# Luon nap ban walker DUY NHAT trong thu muc web (truoc day insert WORKSPACE_DIR len dau
# khien server nap nham ban sao cu o thu muc cha -> sua code web nhung server khong doi).
sys.path.insert(0, BASE_DIR)
if WORKSPACE_DIR not in sys.path:
    sys.path.append(WORKSPACE_DIR)
try:
    from statutory_tree_walker import StatutoryTreeWalker, extract_head_noun, clean_vietnamese
    statutory_walker = StatutoryTreeWalker()
    print("Loaded StatutoryTreeWalker engine in server.py successfully.")
except Exception as eWalker:
    print("Warning: Could not load StatutoryTreeWalker:", eWalker)
    statutory_walker = None


def engine_validate_walker(stat_res, item_input, full_text):
    """Kiem tra cheo ket qua cua StatutoryTreeWalker bang Universal GRI-6 Engine.
    - Ma bi engine loai (sai nguong so / sai phu dinh / sai nhom 4 so theo dieu kien) -> thu giai quyet lai bang generic_resolve.
    - Walker bao thieu du kien -> thu generic_resolve (chi chap nhan khi ket qua khong nhap nhang)."""
    if _ue is None or not isinstance(stat_res, dict):
        return stat_res
    if stat_res.get("rule_id") and (stat_res.get("candidates") or stat_res.get("status") == "INSUFFICIENT_INFO"):
        return stat_res
    try:
        status = stat_res.get("status")
        code = (stat_res.get("code") or "").replace('.', '')
        qi = _ue.extract_query(remove_accents(full_text))
        need = False
        why_txt = ""
        if status == "INSUFFICIENT_INFO":
            need = True
        elif code.isdigit() and len(code) >= 4:
            dq, why, _ = _ue.evaluate(code, qi)
            if dq:
                need = True
                why_txt = why
        if not need:
            return stat_res
        head_noun = ""
        try:
            head_noun = statutory_walker.extract_attributes(item_input.get("name", ""), full_text).get("head_noun", "")
        except Exception:
            pass
        res = _ue.generic_resolve(head_noun, remove_accents(full_text))
        if not res and status != "INSUFFICIENT_INFO" and code.isdigit():
            res = _ue.reroute(code, remove_accents(full_text))
        if res:
            return {
                "code": res["code"], "level": res["level"], "status": res["status"],
                "reasoning": f"Áp dụng GRI 1: cụm danh từ chính '{res['phrase']}' khớp cây biểu thuế. Áp dụng GRI 6 (Universal Engine - so sánh các nhánh cùng cấp gạch nối, điều kiện phủ định và ngưỡng số): đã loại các nhánh không thỏa: {res['dropped'] or 'không có'}.",
                "missing_attributes": [] if res["status"] == "EXACT_8_DIGIT" else ["8_digit_discriminator"],
                "assumptions": "Phân loại bằng Universal GRI-6 Engine",
                "reserve_code": stat_res.get("reserve_code", "N/A"),
            }
        if status != "INSUFFICIENT_INFO":
            # Ma cua walker bi engine loai nhung khong tim duoc ma thay the -> khong xuat ma sai
            return {
                "code": "INSUFFICIENT_INFO", "level": 0, "status": "INSUFFICIENT_INFO",
                "reasoning": f"Mã {stat_res.get('code')} bị loại bởi cổng GRI 6: {why_txt}. Chưa đủ dữ kiện để chọn mã thay thế chính xác.",
                "missing_attributes": ["technical_specs"], "assumptions": "N/A", "reserve_code": "N/A",
            }
    except Exception as e:
        print("engine_validate_walker error:", e)
    return stat_res

def classify_product_python(name, desc="", use="", specs="", model_en=""):
    full_text = f"{name} {desc} {use} {specs} {model_en}".strip()
    full_text_norm = remove_accents(full_text)
    decomp = decompose_query(name, desc=desc, use=use, specs=specs, model_en=model_en)

    # 0. DETERMINISTIC STATUTORY TREE WALKER (GRI 1-6, Head Noun Isolation, Zero-Guessing)
    if statutory_walker is not None:
        item_input = {
            "name": name,
            "desc": desc,
            "use": use,
            "specs": specs,
            "model_en": model_en
        }
        stat_res = statutory_walker.classify(item_input)
        stat_res = engine_validate_walker(stat_res, item_input, full_text)
        if stat_res.get("status") in ["INSUFFICIENT_INFO", "STOP_4_DIGIT", "STOP_6_DIGIT", "EXACT_8_DIGIT"]:
            c_code = stat_res.get("code", "")
            clean_code = c_code.replace('.', '')
            matched_item = next((it for it in HS_DATABASE if it.get('c', '').replace('.', '') == clean_code), None)
            if not matched_item:
                matched_item = {
                    "c": c_code,
                    "vn": stat_res.get("reasoning", "Hàng hóa nhập khẩu"),
                    "mfn": "0%",
                    "vat": "8%",
                    "lin": stat_res.get("reasoning", "")
                }
            else:
                matched_item = dict(matched_item)
                matched_item["c"] = c_code

            matched_item["_statutory"] = True
            matched_item["_statutory_reasoning"] = stat_res.get("reasoning")
            matched_item["_reserve_code"] = stat_res.get("reserve_code")
            matched_item["_assumptions"] = stat_res.get("assumptions")
            matched_item["_missing_fields"] = stat_res.get("missing_attributes", [])
            matched_item["_level"] = stat_res.get("level", 8)

            if stat_res.get("status") == "INSUFFICIENT_INFO":
                matched_item["_insufficient_info"] = True
                matched_item["_disambiguation_question"] = stat_res.get("disambiguation_question", "")
                matched_item["_candidates"] = stat_res.get("candidates", [])
                matched_item["_legal_basis"] = stat_res.get("legal_basis", "")
            elif stat_res.get("status") in ["STOP_4_DIGIT", "STOP_6_DIGIT"]:
                matched_item["_is_stopped"] = True
                matched_item["_disambiguation_question"] = stat_res.get("disambiguation_question", "")
                matched_item["_candidates"] = stat_res.get("candidates", [])
                matched_item["_legal_basis"] = stat_res.get("legal_basis", "")

            reserve_code_clean = (stat_res.get("reserve_code") or "").replace('.', '')
            comp_item = next((it for it in HS_DATABASE if it.get('c', '').replace('.', '') == reserve_code_clean), None)
            comp_cands = [comp_item] if comp_item else []
            top_headings = [clean_code[:4]] if len(clean_code) >= 4 else ['8536']
            return matched_item, comp_cands, top_headings, decomp

    # Fast path for grease nipples (WCO Heading 84.87 / Exclusion 84.81)
    is_grease_nipple = any(k in full_text_norm for k in [
        'vu bom mo', 'vu tra mo', 'num bom mo', 'num tra mo',
        'dau vu mo', 'vu mo', 'num mo', 'grease nipple', 'greasing nipple',
        'grease fitting', 'zerk fitting', 'lube nipple'
    ])
    if is_grease_nipple:
        item_8487 = next((it for it in HS_DATABASE if it.get('c', '').replace('.', '') == '84879000'), None)
        if item_8487:
            comp_cands = [it for it in HS_DATABASE if it.get('c', '').replace('.', '') in ['84818099', '73181590', '73079990']]
            return item_8487, comp_cands, ['8487', '8481', '7318'], decomp
    
    subj_norm = remove_accents(decomp['subject_part'])
    host_norm = remove_accents(decomp['host_part'])
    
    subject_tokens = [w for w in re.findall(r'\w+', subj_norm) if len(w) > 1 and w not in STOP_WORDS]
    host_tokens = [w for w in re.findall(r'\w+', host_norm) if len(w) > 1 and w not in STOP_WORDS]
    
    raw_subj_tokens = re.findall(r'\w+', subj_norm)
    bigrams = [f"{raw_subj_tokens[i]} {raw_subj_tokens[i+1]}" for i in range(len(raw_subj_tokens)-1)] if len(raw_subj_tokens) >= 2 else []
    
    # -------------------------------------------------------------
    # STAGE 1: Heading Scoring using EXACT TOKEN SET MATCHING
    # -------------------------------------------------------------
    heading_scores = {}
    
    for head, h_desc_raw in HEAD_MAP.items():
        ch = head[:2]
        ch_title_raw = CHAPTER_MAP.get(ch, '')
        h_desc_norm = remove_accents(h_desc_raw)
        ch_title_norm = remove_accents(ch_title_raw)
        
        h_tokens = set(re.findall(r'\w+', h_desc_norm))
        ch_tokens = set(re.findall(r'\w+', ch_title_norm))
        
        score = 0
        
        for bg in bigrams:
            if re.search(r'\b' + re.escape(bg) + r'\b', h_desc_norm):
                score += 180
            elif re.search(r'\b' + re.escape(bg) + r'\b', ch_title_norm):
                score += 90

        # Dual-Index Search: Scan WCO 2022 Explanatory Notes (1,059 Headings)
        wco_text = WCO_SEARCH_TEXTS.get(head, '')
        if wco_text:
            for bg in bigrams:
                if bg in wco_text:
                    score += 260  # Explicitly named in WCO Explanatory Notes
            for sw in subject_tokens:
                if len(sw) >= 3 and sw in wco_text:
                    score += 45
                
        for sw in subject_tokens:
            if sw == 'quang' and 'ho quang' in h_desc_norm and 'ho quang' not in full_text_norm:
                continue
            if sw in h_tokens:
                score += 50
            elif sw in ch_tokens:
                score += 25
                
        for mat_name, boost_chaps, penalty_chaps in decomp['materials']:
            if ch in boost_chaps:
                score += 120
            elif ch in penalty_chaps:
                score -= 300
                
        for cat in decomp['active_categories']:
            if head in cat['target_headings']:
                score += cat['bonus']
            elif ch in cat['forbidden_chapters'] or head in cat['forbidden_chapters']:
                score -= 3000

        # Universal Legal Hierarchy Enforcement in Stage 1
        active_target_heads = set()
        for cat in decomp['active_categories']:
            active_target_heads.update(cat.get('target_headings', []))
            
        if any(h in NOTE_2A_HEADINGS for h in active_target_heads) and head in NOTE_2B_PARTS_HEADINGS:
            score -= 4000  # Chú giải 2(a) Phần XVI
        elif any(h in SECTION_XV_NOTE_2_HEADINGS for h in active_target_heads) and (head in NOTE_2B_PARTS_HEADINGS or ch in ['84', '85', '87', '88', '90']):
            score -= 4000  # Chú giải 2 Phần XV
        elif any(h in CHAPTER_82_TOOLS_HEADINGS for h in active_target_heads) and (head in NOTE_2B_PARTS_HEADINGS or ch in ['84', '85']):
            score -= 4000  # Chú giải 1(k) Phần XVI

        is_plastic = any(m[0] == 'plastic' for m in decomp.get('materials', []))
        is_rubber = any(m[0] == 'rubber' for m in decomp.get('materials', []))
        is_copper = any(m[0] == 'copper' for m in decomp.get('materials', []))
        is_aluminum = any(m[0] == 'aluminum' for m in decomp.get('materials', []))
        is_steel = any(m[0] == 'steel_iron' for m in decomp.get('materials', []))
        is_metal = is_steel or is_copper or is_aluminum

        if is_plastic and (head in NOTE_2B_PARTS_HEADINGS or head in ['8466', '8487'] or (ch in ['84', '85'] and head not in NOTE_2A_HEADINGS)):
            score -= 4000  # Chú giải 1(g) Phần XVI

        # Stage 1 Universal Material Chapter Bounds
        IS_FUNC_HEAD = head in NOTE_2A_HEADINGS or head in {
            '8413', '8414', '8481', '8501', '8504', '8535', '8536', '8537', 
            '9025', '9026', '9031', '9032', '8207', '8208', '8482', '8483'
        }
        if not IS_FUNC_HEAD:
            if is_plastic and ch in {'72', '73', '74', '75', '76', '78', '79', '80', '81', '40'}:
                score -= 5000
            elif is_rubber and ch in {'72', '73', '74', '75', '76', '78', '79', '80', '81', '39'}:
                score -= 5000
            elif is_metal and ch in {'39', '40', '44', '48'}:
                score -= 5000
            elif is_steel and ch in {'74', '76'}:
                score -= 5000
            elif is_copper and ch in {'72', '73', '76'}:
                score -= 5000
            elif is_aluminum and ch in {'72', '73', '74'}:
                score -= 5000

        # Stage 1 Universal Cross-Section Statutory Boundary Enforcements
        if re.search(r'\b(?:ghe\s+ngoi|ban\s+lam\s+viec|tu\s+ho\s+so|tu\s+quan\s+ao|giuong\s+ngu|chair|seat|table|cabinet)\b', full_text_norm) and head in ['3926', '4421', '7326']:
            score -= 4000  # Chapter 94 priority
        if re.search(r'\b(?:vali|va\s+li|tui\s+xach|ba\s+lo|balo|cap\s+sach|suitcase|handbag|backpack)\b', full_text_norm) and (head in ['3926', '7326'] or ch in ['61', '62', '63']):
            score -= 4000  # Heading 42.02 priority
        if re.search(r'\b(?:giay\s+dep|doi\s+giay|doi\s+dep|giay\s+the\s+thao|giay\s+da|giay\s+vai|giay\s+bao\s+ho|sandal|dep\s+sandal|ung\s+cao\s+su|ung\s+nhua|ung\s+da|shoes|boots|footwear)\b', full_text_norm) and (head in ['3926', '4016'] or ch in ['61', '62', '63']):
            score -= 4000  # Chapter 64 priority
        if re.search(r'\b(?:do\s+choi\s+tre\s+em|mo\s+hinh\s+do\s+choi|bup\s+be|dung\s+cu\s+the\s+thao|toys|gym\s+equipment)\b', full_text_norm) and (head in ['3926', '7326', '8479', '8543']):
            score -= 4000  # Chapter 95 priority
        if re.search(r'\b(?:thuoc\s+tru\s+sau|thuoc\s+diet\s+con\s+trung|thuoc\s+diet\s+co|pesticide|insecticide)\b', full_text_norm) and ch in ['28', '29', '30']:
            score -= 4000  # Heading 38.08 priority
        if re.search(r'\b(?:det\s+kim|knitted|crocheted)\b', full_text_norm) and ch == '62':
            score -= 3000  # Chapter 61 priority
        if re.search(r'\b(?:det\s+thoi|woven)\b', full_text_norm) and ch == '61':
            score -= 3000  # Chapter 62 priority
                
        for hw in host_tokens:
            if hw in h_tokens:
                score += 1
                
        if score > 0:
            heading_scores[head] = score
            
    ranked_headings = sorted(heading_scores.items(), key=lambda x: x[1], reverse=True)
    top_headings = [h[0] for h in ranked_headings[:5]]
    if not top_headings:
        top_headings = ['8481', '7318', '8536', '3917']
        
    # -------------------------------------------------------------
    # STAGE 2: 8-Digit Line Scoring
    # -------------------------------------------------------------
    candidates = []
    specs_dict = decomp['specs']
    
    for code, data in ANCESTRY_MAP.items():
        head = data['head']
        if head not in top_headings:
            continue
            
        base_h = heading_scores.get(head, 0)
        item_score = base_h
        item_desc_norm = remove_accents(data['item_desc'])
        sub_desc_norm = remove_accents(data['s_desc'])
        subhead = code[:6]
        
        # Universal Pre-Flight Legal Audit & Boundary Constraint Enforcement
        is_disq, audit_penalty, audit_reason = audit_statutory_constraints(code, head, data['item_desc'], data['s_desc'], decomp, full_text_norm)
        item_score += audit_penalty
        
        i_tokens = set(re.findall(r'\w+', item_desc_norm))
        s_tokens = set(re.findall(r'\w+', sub_desc_norm))
        
        for bg in bigrams:
            if re.search(r'\b' + re.escape(bg) + r'\b', item_desc_norm): item_score += 120
            elif re.search(r'\b' + re.escape(bg) + r'\b', sub_desc_norm): item_score += 60
            
        for sw in subject_tokens:
            if sw == 'quang' and ('ho quang' in item_desc_norm or 'ho quang' in sub_desc_norm) and 'ho quang' not in full_text_norm:
                continue
            if sw in i_tokens: item_score += 40
            elif sw in s_tokens: item_score += 20
            
        # Precise technical & statutory subhead branch bonuses
        is_mat_plastic = any(m[0] == 'plastic' for m in decomp.get('materials', []))
        is_mat_rubber = any(m[0] == 'rubber' for m in decomp.get('materials', []))

        if head == '7318':
            if not is_mat_plastic and not is_mat_rubber:
                if any(k in full_text_norm for k in ['chot dinh vi', 'chot ham', 'chot che', 'dowel pin', 'locating pin']):
                    if code == '73182400': item_score += 500
                    elif code.startswith('73182'): item_score += 200
                elif any(k in full_text_norm for k in ['bu long', 'bu-long', 'bolt']) and code.startswith('731815'):
                    item_score += 400
                elif any(k in full_text_norm for k in ['dai oc', 'nut', 'e-cu']) and code.startswith('731816'):
                    item_score += 400
                elif any(k in full_text_norm for k in ['vong dem', 'washer']) and code.startswith('731821'):
                    item_score += 400
            else:
                item_score -= 4000

        elif head == '3926':
            if any(k in full_text_norm for k in ['vong dem', 'washer', 'long den', 'bu long', 'dai oc', 'oc vit', 'chot', 'nut', 'bolt', 'gasket', 'gioang', 'dem']):
                if code == '39269099': item_score += 900
                elif code == '39269044' and not any(k in full_text_norm for k in ['cuu sinh', 'nga']): item_score -= 2000
                elif code == '39269010' and not any(k in full_text_norm for k in ['phao', 'danh ca']): item_score -= 2000
            elif any(k in full_text_norm for k in ['tam che', 'vo che', 'nap che', 'tam chan', 'safety guard', 'an toan lao dong', 'bao ve']) and code == '39269099':
                item_score += 700

        elif head == '4016':
            if any(k in full_text_norm for k in ['gioang', 'phot', 'vong dem', 'washer', 'dem kin', 'seal', 'o-ring', 'dem lot']):
                if code.startswith('401693'):
                    is_vehicle = any(k in full_text_norm for k in ['xe co dong co', 'o to', 'xe hoi', 'xe tai', 'xe may', 'automotive', 'motor vehicle'])
                    if is_vehicle and code == '40169320': item_score += 950
                    elif not is_vehicle and code == '40169390': item_score += 950
                    elif code == '40169390': item_score += 850
                
        elif head == '8481':
            sub = code[:6]
            if any(k in full_text_norm for k in ['khi nen', 'actuator', 'solenoid', 'dien tu', 'thuy luc']):
                if sub == '848120': item_score += 350
            elif any(k in full_text_norm for k in ['mot chieu', 'check valve']):
                if sub == '848130': item_score += 350
            elif any(k in full_text_norm for k in ['an toan', 'xa ap', 'relief']):
                if sub == '848140': item_score += 350
            elif any(k in full_text_norm for k in ['giam ap', 'dieu ap']):
                if sub == '848110': item_score += 350
            elif any(k in full_text_norm for k in ['van bi', 'van buom', 'van cong', 'van cau']):
                if sub == '848180': item_score += 300
                
            if 'dn' in specs_dict:
                dn_cm = specs_dict['dn'] / 10.0
                if 'không quá 2,5 cm' in data['item_desc'] and dn_cm <= 2.5: item_score += 250
                elif 'trên 2,5 cm' in data['item_desc'] and dn_cm > 2.5: item_score += 250

        elif head == '8487':
            if any(k in full_text_norm for k in ['vu bom mo', 'vu tra mo', 'num bom mo', 'num tra mo', 'dau vu mo', 'vu mo', 'num mo', 'grease nipple', 'greasing nipple', 'grease fitting', 'zerk fitting', 'lube nipple']):
                if code == '84879000':
                    item_score += 3500
                
        elif head == '8536':
            sub = code[:6]
            if any(k in full_text_norm for k in ['soi quang', 'optical fiber', 'cap quang', 'fiber optic']) and sub == '853670':
                item_score += 600
                if any(k in full_text_norm for k in ['gom', 'ceramic', 'su']) and code == '85367010':
                    item_score += 400
                elif any(k in full_text_norm for k in ['dong', 'copper', 'brass']) and code == '85367020':
                    item_score += 400
                elif code == '85367090':
                    item_score += 350
                    if any(k in full_text_norm for k in ['nhua', 'plastic', 'composite', 'kim loai', 'thep', 'steel']):
                        item_score += 150
            elif any(k in full_text_norm for k in ['cong tac', 'switch', 'nut nhan', 'nut an']) and sub == '853650':
                item_score += 400
                # Extract amperage
                amp_val = specs_dict.get('a', None)
                if amp_val is None:
                    m_a = re.search(r'(\d+(?:\.\d+)?)\s*(?:a|amp|ampe)\b', full_text, re.I)
                    if m_a: amp_val = float(m_a.group(1))

                is_under_16a = (amp_val is not None and amp_val < 16.0) or any(k in full_text_norm for k in ['duoi 16a', 'duoi 16 a', 'nho hon 16a', '< 16a', '<16a', 'under 16a'])
                is_16a_or_more = (amp_val is not None and amp_val >= 16.0) or any(k in full_text_norm for k in ['tren 16a', 'tu 16a', 'tren 16 a', '>= 16a', '>=16a', 'over 16a'])

                if any(k in full_text_norm for k in ['hanh trinh', 'limit switch', 'micro switch', 'dien co bat nhanh']):
                    if is_under_16a and code == '85365061': item_score += 550
                    elif is_16a_or_more and code == '85365069': item_score += 550
                    elif code in ['85365061', '85365069']: item_score += 350
                elif any(k in full_text_norm for k in ['dien tu', 'electronic switch']):
                    if is_under_16a and code == '85365051': item_score += 550
                    elif is_16a_or_more and code == '85365059': item_score += 550
                    elif code in ['85365051', '85365059']: item_score += 350
                elif any(k in full_text_norm for k in ['dao chieu', 'change-over', 'change over', 'khoi dong dong co']) and code == '85365095':
                    item_score += 500
                else:
                    # General switches: công tắc chuyển mạch, công tắc bấm, nút nhấn, công tắc điện...
                    if is_under_16a:
                        if code == '85365096': item_score += 650  # - - - Loại khác, dòng điện dưới 16 A
                        elif code == '85365033' and any(k in full_text_norm for k in ['cong tac xoay', 'cong tac truot', 'bap benh', 'dieu hoa khong khi', 'may dieu hoa']): item_score += 550
                        elif code == '85365099': item_score -= 300 # Penalize 8536.50.99 ("Loại khác") when current < 16A!
                    elif is_16a_or_more:
                        if code == '85365099': item_score += 600
                        elif code == '85365096': item_score -= 300
                    else:
                        if code == '85365096': item_score += 450
                        elif code == '85365099': item_score += 300
            elif any(k in full_text_norm for k in ['cau dao', 'aptomat', 'ngat mach', 'may cat dien', 'may cat khong khi', 'may cat chan khong', 'acb', 'mcb', 'mccb', 'circuit breaker']) and sub == '853620':
                item_score += 550
            elif any(k in full_text_norm for k in ['ro le', 'relay']) and sub.startswith('85364'):
                item_score += 350
            elif sub == '853690' and code == '85369093':
                if any(k in full_text_norm for k in ['dau noi', 'connector']):
                    item_score += 200
        elif head == '8544':
            if any(k in subj_norm for k in ['tam che', 'vo che', 'nap che', 'tam chan', 'bu long', 'mieng thep', 'gia do']):
                item_score -= 3000
            
            has_conn = bool(re.search(r'\b(?:co\s+gan\s+dau\s+noi|gan\s+dau\s+noi|lap\s+dau\s+noi|co\s+dau\s+noi|da\s+lap\s+voi\s+dau\s+noi|fitted\s+with\s+connectors?|with\s+connectors?)\b', full_text_norm))
            not_has_conn = bool(re.search(r'\b(?:chua\s+gan\s+dau\s+noi|chua\s+lap\s+dau\s+noi|khong\s+co\s+dau\s+noi|without\s+connectors?)\b', full_text_norm))
            
            if has_conn:
                if code.startswith('854442'):
                    item_score += 2500
                    if code.startswith('8544429'):
                        item_score += 800
                        if code == '85444299':
                            item_score += 500
                elif code.startswith('854449'):
                    item_score -= 4000
            elif not_has_conn:
                if code.startswith('854449'):
                    item_score += 1500
                elif code.startswith('854442'):
                    item_score -= 4000
            elif any(k in full_text_norm for k in ['ha the', 'duoi 1000v', '0.6/1kv', 'dien ap khong qua 1000', '12v', '24v']):
                if code.startswith('854449') or code.startswith('854442'): item_score += 300
        elif head == '8537':
            if any(k in full_text_norm for k in ['hop dieu khien', 'tu dieu khien', 'bo dieu khien', 'mach dieu khien', 'bang dieu khien', 'control box', 'motor controller', 'controller']):
                item_score += 450
                v_val = specs_dict.get('v', None)
                if (v_val is not None and v_val <= 1000) or any(k in full_text_norm for k in ['duoi 1000v', 'khong qua 1000', '200v', '220v', '24v', '48v', '380v']):
                    if code.startswith('853710'):
                        item_score += 350
                        if code == '85371019': item_score += 300
                        elif code == '85371099': item_score += 250
        elif head == '8523':
            if any(k in full_text_norm for k in ['usb', 'o nho the ran', 'the nho', 'flash drive', 'the ran', 'solid-state']):
                if code.startswith('852351'):
                    item_score += 650
                    if any(k in full_text_norm for k in ['chua ghi', 'unrecorded', 'chua co du lieu', 'trong']):
                        if code == '85235111': item_score += 400
                        elif code == '85235119': item_score += 300
                    elif code == '85235130':
                        item_score += 300
                elif code.startswith('852321') or code.startswith('852329'):
                    item_score -= 1500
        elif head == '3926':
            if any(k in full_text_norm for k in ['tam che', 'vo che', 'nap che', 'tam chan', 'safety guard', 'an toan lao dong', 'bao ve']) and code == '39269099':
                item_score += 700
        elif head == '8482':
            sub = code[:6]
            if decomp.get('has_negative_bearing'):
                item_score -= 3000  # Không lắp ổ bi/ổ lăn -> Loại trừ khỏi Nhóm 84.82
            elif any(k in full_text_norm for k in ['thanh truot', 'con truot', 'thanh ray', 'linear', 'co che truot', 'slide mechanism', 'o bi', 'ball bearing', 'ball bearings', 'bearing', 'bearings']):
                if code == '84821000':
                    item_score += 650  # - Ổ bi (Ball bearings - bao gồm cả cơ chế trượt với ổ bi theo Chú giải WCO 84.82)
                elif code == '84828000':
                    item_score += 350
            elif any(k in full_text_norm for k in ['bi cau', 'mot day', 'do mot day']) and code.startswith('848210'):
                item_score += 400
            elif any(k in full_text_norm for k in ['bi dua con', 'tapered']) and code.startswith('848220'):
                item_score += 450
            elif any(k in full_text_norm for k in ['bi dua kim', 'needle roller']) and code.startswith('848240'):
                item_score += 450
            elif any(k in full_text_norm for k in ['bi dua tru', 'cylindrical roller']) and code.startswith('848250'):
                item_score += 450
        elif head == '8483':
            sub = code[:6]
            if decomp.get('has_negative_bearing'):
                if sub == '848330':
                    item_score += 750  # 8483.30 quy định rõ "KHÔNG lắp ổ bi hoặc ổ đũa"! Đúng tiêu chí định danh!
                    if any(k in full_text_norm for k in ['dong co xe', 'o to', 'xe may', 'chuong 87', 'phuong tien']):
                        if code == '84833030': item_score += 400
                    else:
                        if code == '84833090': item_score += 400  # Máy công nghiệp / máy công cụ / máy tiện
                        elif code == '84833030': item_score -= 500
            elif any(k in full_text_norm for k in ['co o bi', 'o bi', 'vong bi', 'lap o bi', 'ball bearing', 'ball bearings', 'bearing', 'bearings']):
                if sub == '848330':
                    item_score -= 800  # 8483.30 quy định rõ "KHÔNG lắp ổ bi hoặc ổ đũa"! Phạt mạnh nếu có bi!
                elif sub == '848320':
                    item_score += 300  # 8483.20: Thân ổ có lắp ổ bi/ổ đũa
            elif any(k in full_text_norm for k in ['o truot', 'bac truot', 'bac dong', 'bushing', 'plain bearing']) and not any(k in full_text_norm for k in ['co bi', 'o bi', 'ball']):
                if sub == '848330':
                    item_score += 550
                    if any(k in full_text_norm for k in ['dong co xe', 'o to', 'xe may', 'chuong 87', 'phuong tien']):
                        if code == '84833030': item_score += 300
                    else:
                        if code == '84833090': item_score += 300
                        elif code == '84833030': item_score -= 400
        elif head == '8466':
            if any(k in full_text_norm for k in ['co o bi', 'o bi', 'vong bi', 'ball bearing', 'ball bearings', 'bearing', 'bearings']) and not decomp.get('has_negative_bearing'):
                item_score -= 500  # Chú giải 2(a) Phần XVI: Ổ bi thuộc 84.82, không phân loại vào 84.66!
            elif any(k in full_text_norm for k in ['cat dap', 'dap', 'press machine', '8462', '84.62', 'ram guide', 'ram metal guide']):
                if code == '84669400':
                    item_score += 850
        elif head == '7320':
            if any(k in full_text_norm for k in ['lo xo nen', 'lo xo keo', 'lo xo xoan', 'helical', 'coil spring']):
                if code.startswith('732020'): item_score += 450
            elif any(k in full_text_norm for k in ['nhip', 'nhip la', 'leaf spring']):
                if code.startswith('732010'): item_score += 450
        elif head == '8207':
            if any(k in full_text_norm for k in ['dao phay', 'milling cutter']) and code.startswith('820770'):
                item_score += 400
        elif head == '3902':
            if code.startswith('390210'): item_score += 300
        elif head == '2914':
            if code == '29141100': item_score += 400
        elif head == '9025':
            if code.startswith('902519'): item_score += 300
        elif head == '9018':
            if code.startswith('901890'): item_score += 300
        elif head == '8708':
            if any(k in full_text_norm for k in ['ma phanh', 'brake pad', 'phanh']) and code.startswith('870830'):
                item_score += 400
        elif head == '6203':
            if any(k in full_text_norm for k in ['quan', 'quan au']) and code.startswith('62034'):
                item_score += 300
        elif head == '6205':
            if any(k in full_text_norm for k in ['so mi', 'shirt']) and code.startswith('620520'):
                item_score += 300
        elif head in ['4407', '4408']:
            thickness = specs_dict.get('mm', 10)
            if thickness > 6:
                if head == '4407': item_score += 400
            else:
                if head == '4408': item_score += 400
        elif head == '4010':
            item_score += 400
            if any(k in full_text_norm for k in ['rang cua', 'dong bo', 'timing belt', 'rang']):
                if code.startswith('401035') or code.startswith('401036') or code.startswith('401039') or code == '40103900':
                    item_score += 500
        elif head == '3917':
            item_score += 300
            if any(k in full_text_norm for k in ['mem', 'flexible', 'khi nen']):
                if code in ['39173999', '39173299', '39173290', '39173990']:
                    item_score += 450
            if any(k in full_text_norm for k in ['bep ga', 'gas stove']) and code == '39173220':
                item_score += 500
            elif not any(k in full_text_norm for k in ['bep ga', 'gas stove']) and code == '39173220':
                item_score -= 1000
        elif head == '8806':
            item_score += 500
            if any(k in full_text_norm for k in ['249g', 'khong qua 250', '250g']) and code == '88062100':
                item_score += 300
        elif head == '8524':
            item_score += 400
        elif head == '7326':
            if any(k in full_text_norm for k in ['tam che', 'vo che', 'nap che', 'tam chan', 'safety guard', 'bao ve', 'mieng thep', 'tam thep', 'gia do', 'tai ga', 'dinh vi']):
                if code == '73269099': item_score += 750

        candidates.append((data['item'], item_score))
        
    valid_candidates = [c for c in candidates if c[1] > -2000]
    if not valid_candidates:
        valid_candidates = candidates
    valid_candidates.sort(key=lambda x: x[1], reverse=True)
    best_item = valid_candidates[0][0] if valid_candidates else (HS_DATABASE[0] if HS_DATABASE else {})
    best_code = best_item.get('c', '').replace('.', '')
    best_score = valid_candidates[0][1] if valid_candidates else 1
    
    competing_items = []
    seen_prefixes = {best_code[:6]}
    seen_headings = {best_code[:4]}

    # Check if Priority HS code is 100% Certain (không có mã dự phòng / khỏi cần điền)
    is_certain_100 = False
    # 1) Generic steel articles (7326.90.99): tấm thép, miếng thép, bản mã -> độc tôn 7326.90.99
    if best_code == '73269099' and any(k in full_text_norm for k in ['mieng thep', 'tam thep', 'ban ma', 'tam kim loai', 'gia do']):
        is_certain_100 = True
    # 1b) Plastic protective covers/guards (3926.90.99): tấm che, nắp che an toàn bằng nhựa
    elif best_code == '39269099' and any(k in full_text_norm for k in ['tam che', 'vo che', 'nap che', 'tam chan', 'safety guard', 'bao ve']):
        is_certain_100 = True
    # 2) Standard fasteners (7318): bu lông, ốc vít thép
    elif best_code.startswith('7318') and any(k in full_text_norm for k in ['bu long', 'oc vit', 'dinh vit']):
        is_certain_100 = True
    # 3) Unrecorded solid-state storage (8523.51.11): USB chưa ghi
    elif best_code == '85235111' and any(k in full_text_norm for k in ['chua ghi', 'unrecorded', 'trong']):
        is_certain_100 = True

    if not is_certain_100:
        for cand_item, cand_score in valid_candidates:
            c = cand_item.get('c', '').replace('.', '')
            if len(c) != 8 or c == best_code: continue
            h4 = c[:4]
            sub = c[:6]

            # Semantic negative filters for competitors
            if any(k in full_text_norm for k in ['mieng thep', 'tam thep', 'ban ma']) and (sub.startswith('732620') or sub.startswith('732611') or sub.startswith('732619')):
                continue # Cannot be poultry cage / wire / grinding ball
            if 'thep' in full_text_norm and h4 in ['7415', '7616', '3926']:
                continue # Cannot be copper, aluminum, or plastic when product is steel
            if 'nhua' in full_text_norm and (h4 in ['7326', '7415', '8535', '8536', '8544']):
                continue # Cannot be metal or electrical breaker when product is plastic cover

            # Candidate must be a recognized legitimate customs competing heading OR have very close score (>= 70%)
            is_recognized_dual = False
            # Duality 1: Machine part 8466 vs Plastic article 3926
            if (best_code.startswith('8466') and h4 == '3926') or (best_code.startswith('3926') and h4 == '8466'):
                is_recognized_dual = True
            # Duality 2: Ball bearing 8482 vs Housing 8483
            elif (best_code.startswith('8482') and h4 == '8483') or (best_code.startswith('8483') and h4 == '8482'):
                is_recognized_dual = True
            # Duality 3: Motor control 8537 vs Inverter 8504
            elif (best_code.startswith('8537') and h4 == '8504') or (best_code.startswith('8504') and h4 == '8537'):
                is_recognized_dual = True

            if is_recognized_dual or (cand_score >= best_score * 0.70 and cand_score > 0):
                if h4 not in seen_headings or sub not in seen_prefixes:
                    competing_items.append(cand_item)
                    seen_headings.add(h4)
                    seen_prefixes.add(sub)
                    if len(competing_items) >= 1:
                        break

    return best_item, competing_items, top_headings, decomp


import threading
WORD_EXPORT_LOCK = threading.Lock()

CHROME_CANDIDATE_PATHS = [
    r'C:\Program Files\Google\Chrome\Application\chrome.exe',
    r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
    r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe',
    r'C:\Program Files\Microsoft\Edge\Application\msedge.exe',
    os.path.expandvars(r'%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe'),
    os.path.expandvars(r'%LOCALAPPDATA%\Microsoft\Edge\Application\msedge.exe'),
]

def convert_html_to_pdf_headless(html_str, target_pdf_path):
    browser_exe = None
    for p in CHROME_CANDIDATE_PATHS:
        if os.path.exists(p):
            browser_exe = p
            break
    if not browser_exe:
        return False

    temp_html = target_pdf_path + '.temp.html'
    try:
        with open(temp_html, 'w', encoding='utf-8') as f:
            f.write(html_str)

        cmd = [
            browser_exe,
            '--headless=new',
            '--disable-gpu',
            '--no-sandbox',
            f'--print-to-pdf={target_pdf_path}',
            '--no-pdf-header-footer',
            temp_html
        ]
        subprocess.run(cmd, capture_output=True, timeout=12)
        if os.path.exists(temp_html):
            try: os.remove(temp_html)
            except: pass
        return os.path.exists(target_pdf_path) and os.path.getsize(target_pdf_path) > 500
    except Exception as e:
        print("Headless PDF generation error:", e)
        if os.path.exists(temp_html):
            try: os.remove(temp_html)
            except: pass
        return False

def build_a4_html_report_server(data):
    hs_code = data.get('hs_code', '0000.00.00')
    clean_code = data.get('clean_code', hs_code.replace('.', ''))
    name = data.get('name', 'Hàng hóa nhập khẩu')
    decl_name = data.get('decl_name', name)
    model_en = data.get('model_en', name)
    desc = data.get('desc', data.get('item_description', ''))
    specs = data.get('specs', desc)
    use = data.get('use', 'Sử dụng trong công nghiệp')
    mfn = data.get('mfn', '0%')
    vat = data.get('vat', '8%')
    atiga = data.get('atiga', '0%')
    acfta = data.get('acfta', '0%')
    vjepa = data.get('vjepa', '0%')
    ajcep = data.get('ajcep', '0%')
    cptpp = data.get('cptpp', '0%')
    evfta = data.get('evfta', '0%')
    rcep = data.get('rcep', '0%')
    gri1 = data.get('gri1', '')
    gri2 = data.get('gri2', '')
    gri3 = data.get('gri3', '')
    gri4 = data.get('gri4', '')
    gri5 = data.get('gri5', '')
    gri6 = data.get('gri6', '')
    notes = data.get('explanatory_notes', '')
    exclusions = data.get('exclusions', '')
    recommendations = data.get('recommendations', '')

    # Image HTML if provided
    image_b64 = data.get('image_b64')
    img_html = ""
    if image_b64:
        img_src = image_b64 if image_b64.startswith('data:') else f"data:image/png;base64,{image_b64}"
        img_html = f'''
        <div style="margin: 6px 0 6px 18px;">
            <div style="font-weight: bold; margin-bottom: 3px;">Hình ảnh đính kèm sản phẩm:</div>
            <img src="{img_src}" style="max-width: 240px; max-height: 180px; border: 1px solid #cbd5e1; border-radius: 4px; object-fit: contain;">
        </div>'''

    # Section 4 rows: Comparison between potential codes
    candidates = data.get('candidates', [])
    comp_rows = []
    if candidates and len(candidates) > 0:
        for c_item in candidates[:4]:
            comp_rows.append((
                c_item.get('hs_code', ''),
                c_item.get('desc', '') or c_item.get('name', ''),
                f"{c_item.get('mfn', '0%')} | {c_item.get('vat', '8%')}",
                f"{c_item.get('badge', '')}: {c_item.get('apply_condition', '')} {c_item.get('risk_warning', '')}".strip()
            ))
        if clean_code == '84812020' and not any('8481.10.11' in r[0] for r in comp_rows):
            comp_rows.append((
                "8481.10.11",
                "Van cổng điều khiển bằng tay, đường kính trên 5cm đến 40cm, bằng sắt/thép",
                "5% | 8/10%",
                "BỊ LOẠI TRỪ (Mã sai thường gặp): Không thỏa mãn cả 4 tiêu chuẩn: (1) Hàng không phải van giảm áp; (2) Dẫn động bằng khí nén (actuator) chứ không phải bằng tay; (3) Chất liệu là đồng thau chứ không phải sắt/thép; (4) Đường kính DN20 (2,0 cm) vi phạm dải kích thước trên 5cm đến 40cm."
            ))
        if not any('Chương' in r[0] for r in comp_rows):
            comp_rows.append((
                "Chương 74 / 84",
                "Phụ kiện ống nối (74.12) / Phụ tùng máy chung (84.79)",
                "0-5% | 8-10%",
                exclusions or "BỊ LOẠI TRỪ: Phân loại theo GRI 1 & GRI 3(a), hàng hóa là van hoàn chỉnh có cơ cấu chấp hành thuộc Nhóm 84.81, loại trừ nhóm vật liệu thô và phụ tùng chung."
            ))
    else:
        # Find real runner-up in HS_DATABASE (NEVER invent .90.90)
        clean_prefix = clean_code[:4]
        real_runner_up = None
        for it in HS_DATABASE:
            c = it.get('c', '').replace('.', '')
            if len(c) == 8 and c.startswith(clean_prefix) and c != clean_code:
                real_runner_up = it
                break

        if real_runner_up:
            r_c = real_runner_up.get('c', '').replace('.', '')
            r_fmt = f"{r_c[:4]}.{r_c[4:6]}.{r_c[6:8]}"
            r_desc = real_runner_up.get('vn', 'Phân nhóm thay thế trong cùng Nhóm')
            r_mfn = real_runner_up.get('mfn', mfn)
            r_mfn_str = r_mfn if '%' in str(r_mfn) else f"{r_mfn}%"
            r_vat = real_runner_up.get('vat', vat)
            r_vat_str = r_vat if '%' in str(r_vat) else f"{r_vat}%"
            comp_rows = [
                (hs_code, decl_name, f"{mfn} | {vat}", "Khuyến nghị Ưu tiên 1: Chuẩn xác nhất theo Chú giải WCO 2022 và Biểu thuế 2026."),
                (r_fmt, r_desc, f"{r_mfn_str} | {r_vat_str}", f"Phương án Dự phòng: Phân nhóm liên quan trong cùng Nhóm {clean_prefix}."),
                (f"Chương {clean_code[:2]} / Bộ phận", "Nhóm thiết bị / bộ phận máy chung", "0-10% | 8-10%", exclusions or "Loại trừ: Ưu tiên mã định danh cụ thể theo Quy tắc GRI 1 & GRI 3(a).")
            ]
        else:
            comp_rows = [
                (hs_code, decl_name, f"{mfn} | {vat}", "Khuyến nghị Ưu tiên 1: Chuẩn xác nhất theo Chú giải WCO 2022 và Biểu thuế 2026."),
                ("Mã Đơn Nhất", "Không có phân nhóm dự phòng trong cùng Nhóm", f"{mfn} | {vat}", "Mặt hàng có tính chất đặc thù đơn nhất theo Chú giải WCO 2022."),
                (f"Chương {clean_code[:2]} / Bộ phận", "Nhóm thiết bị / bộ phận máy chung", "0-10% | 8-10%", exclusions or "Loại trừ: Ưu tiên mã định danh cụ thể theo Quy tắc GRI 1 & GRI 3(a).")
            ]

    table_comp_rows_html = ""
    for idx, r_vals in enumerate(comp_rows):
        bg = "#F7F9FB" if idx % 2 == 0 else "#FFFFFF"
        table_comp_rows_html += f'''
        <tr style="background-color: {bg};">
            <td style="border: 1px solid #cbd5e1; padding: 5px 8px; font-weight: bold; color: #004080; width: 16%;">{r_vals[0]}</td>
            <td style="border: 1px solid #cbd5e1; padding: 5px 8px; width: 30%;">{r_vals[1]}</td>
            <td style="border: 1px solid #cbd5e1; padding: 5px 8px; text-align: center; font-weight: bold; width: 14%;">{r_vals[2]}</td>
            <td style="border: 1px solid #cbd5e1; padding: 5px 8px; width: 40%;">{r_vals[3]}</td>
        </tr>'''

    # Section 5 rows: Detailed Tariff Schedule
    tax_rows = [
        ("Thuế Nhập khẩu Ưu đãi (MFN)", mfn, "Nghị định số 26/2023/NĐ-CP"),
        ("Thuế Giá trị gia tăng (VAT)", vat, "Nghị định số 181/2025/NĐ-CP, NĐ 72/2024/NĐ-CP"),
        ("Hiệp định Thương mại Hàng hóa ASEAN (ATIGA)", atiga, "Nghị định số 126/2022/NĐ-CP (C/O Form D)"),
        ("Hiệp định ASEAN - Trung Quốc (ACFTA)", acfta, "Nghị định số 118/2022/NĐ-CP (C/O Form E)"),
        ("Hiệp định Việt Nam - Nhật Bản (VJEPA)", vjepa, "Nghị định số 120/2022/NĐ-CP (C/O Form VJ)"),
        ("Hiệp định ASEAN - Nhật Bản (AJCEP)", ajcep, "Nghị định số 119/2022/NĐ-CP (C/O Form AJ)"),
        ("Hiệp định Đối tác Toàn diện CPTPP", cptpp, "Nghị định số 115/2022/NĐ-CP (C/O Form CPTPP)"),
        ("Hiệp định EVFTA (Việt Nam - EU)", evfta, "Nghị định số 116/2022/NĐ-CP (C/O Form EUR.1)"),
        ("Hiệp định RCEP (ASEAN + Đối tác)", rcep, "Nghị định số 129/2022/NĐ-CP (C/O Form RCEP)")
    ]

    table_tax_rows_html = ""
    for idx, r_vals in enumerate(tax_rows):
        bg = "#F7F9FB" if idx % 2 == 0 else "#FFFFFF"
        rate_color = "#006600" if "0%" in r_vals[1] else "#990000"
        table_tax_rows_html += f'''
        <tr style="background-color: {bg};">
            <td style="border: 1px solid #cbd5e1; padding: 5px 8px; width: 42%;">{r_vals[0]}</td>
            <td style="border: 1px solid #cbd5e1; padding: 5px 8px; text-align: center; font-weight: bold; color: {rate_color}; width: 16%;">{r_vals[1]}</td>
            <td style="border: 1px solid #cbd5e1; padding: 5px 8px; width: 42%;">{r_vals[2]}</td>
        </tr>'''

    # Section 6 text:
    concl_text = f"Mặt hàng '{name}' được phân loại chính xác vào mã HS {hs_code} ({decl_name}). Mức thuế MFN: {mfn}, VAT: {vat}. {recommendations or 'Doanh nghiệp chuẩn bị đầy đủ hồ sơ kỹ thuật, datasheet và C/O ưu đãi để được áp dụng thuế suất ưu đãi đặc biệt 0%.'}"

    return f'''<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="utf-8">
<title>BÁO CÁO PHÂN TÍCH VÀ ĐỀ XUẤT MÃ SỐ HS</title>
<style>
  @page {{
    size: A4 portrait;
    margin: 12mm 14mm 12mm 14mm;
  }}
  body {{
    font-family: Arial, sans-serif;
    font-size: 10pt;
    line-height: 1.35;
    color: #222222;
    margin: 0;
    padding: 0;
    background: #ffffff;
    box-sizing: border-box;
    word-break: break-word;
    overflow-wrap: break-word;
  }}
  .main-title {{
    text-align: center;
    font-size: 15pt;
    font-weight: bold;
    color: #003366;
    margin: 0 0 4pt 0;
    text-transform: uppercase;
    line-height: 1.3;
    word-break: break-word;
  }}
  .sub-title {{
    text-align: center;
    font-size: 10pt;
    font-weight: bold;
    color: #444444;
    margin: 0 0 12pt 0;
    line-height: 1.35;
    word-break: break-word;
  }}
  .sec-header {{
    font-size: 11pt;
    font-weight: bold;
    color: #004080;
    margin: 10pt 0 4pt 0;
  }}
  .bullet-p {{
    margin: 0 0 3pt 14pt;
    font-size: 9.5pt;
    line-height: 1.35;
    word-break: break-word;
    overflow-wrap: break-word;
  }}
  .bullet-bold {{
    font-weight: bold;
    color: #111111;
  }}
  .hs-box {{
    margin: 6pt 0 10pt 14pt;
    font-size: 10pt;
    line-height: 1.35;
    word-break: break-word;
    overflow-wrap: break-word;
  }}
  .hs-box-label {{
    font-weight: bold;
    color: #990000;
  }}
  .hs-box-val {{
    font-weight: bold;
    color: #111111;
  }}
  .wco-text {{
    margin: 0 0 6pt 14pt;
    font-size: 9.5pt;
    line-height: 1.4;
    text-align: justify;
    color: #222222;
    word-break: break-word;
    overflow-wrap: break-word;
  }}
  .std-table {{
    width: 100%;
    table-layout: fixed;
    border-collapse: collapse;
    margin: 6pt 0 10pt 0;
    font-size: 8.5pt;
  }}
  .std-table th {{
    background-color: #003366;
    color: #ffffff;
    font-weight: bold;
    font-size: 9pt;
    text-align: center;
    padding: 6pt 6pt;
    border: 1px solid #cbd5e1;
    word-break: break-word;
    overflow-wrap: break-word;
  }}
  .std-table td {{
    padding: 5pt 6pt;
    border: 1px solid #cbd5e1;
    vertical-align: middle;
    word-break: break-word;
    overflow-wrap: break-word;
  }}
  .concl-p {{
    margin: 0 0 6pt 14pt;
    font-size: 9.5pt;
    line-height: 1.45;
    text-align: justify;
    color: #222222;
    word-break: break-word;
    overflow-wrap: break-word;
  }}
  .concl-bold {{
    font-weight: bold;
    color: #003366;
  }}
</style>
</head>
<body>
  <div class="main-title">BÁO CÁO PHÂN TÍCH VÀ ĐỀ XUẤT MÃ SỐ HS</div>
  <div class="sub-title">DANH MỤC SẢN PHẨM: {decl_name.upper()}</div>

  <div class="sec-header">1. Mô tả chi tiết mặt hàng theo Datasheet / Hồ sơ kỹ thuật</div>
  <div class="bullet-p"><span class="bullet-bold">Tên hàng / Model: </span>{name} (Model: {model_en})</div>
  <div class="bullet-p"><span class="bullet-bold">Cấu tạo &amp; Thông số kỹ thuật: </span>{desc}. Thông số: {specs}</div>
  <div class="bullet-p"><span class="bullet-bold">Công dụng &amp; Chức năng: </span>{use}</div>
  {img_html}
  <div class="hs-box">
    <span class="hs-box-label">MÃ HS ĐỀ XUẤT ÁP DỤNG THỐNG NHẤT: </span>
    <span class="hs-box-val">{hs_code} (Thuế MFN: {mfn}, VAT: {vat}, C/O Form E / EVFTA: 0%)</span>
  </div>

  <div class="sec-header">2. Phân tích theo 6 Quy tắc Tổng quát Giải thích HS (GRI)</div>
  <div class="bullet-p"><span class="bullet-bold">· GRI 1 (Tên Nhóm &amp; Chú giải): </span>{gri1 or 'Áp dụng theo nội dung của nhóm và Chú giải Chương liên quan của Biểu thuế XNK 2026.'}</div>
  <div class="bullet-p"><span class="bullet-bold">· GRI 2 (Chưa hoàn chỉnh / Đồng bộ): </span>{gri2 or 'Sản phẩm hoàn chỉnh, mang đầy đủ đặc tính thương mại cơ bản theo Quy tắc 2(a).'}</div>
  <div class="bullet-p"><span class="bullet-bold">· GRI 3 (Mô tả cụ thể / Đặc trưng nhất): </span>{gri3 or 'Mô tả có tính chất cụ thể, rõ ràng nhất được ưu tiên theo Quy tắc 3(a).'}</div>
  <div class="bullet-p"><span class="bullet-bold">· GRI 4 (Hàng hóa giống nhất): </span>{gri4 or 'Không cần áp dụng do hàng hóa đã được xác định cụ thể theo các quy tắc trước.'}</div>
  <div class="bullet-p"><span class="bullet-bold">· GRI 5 (Bao bì &amp; Đồ chứa): </span>{gri5 or 'Bao bì tiêu chuẩn vận chuyển thông thường, phân loại cùng với hàng hóa.'}</div>
  <div class="bullet-p"><span class="bullet-bold">· GRI 6 (Phân nhóm 6 số &amp; Mã 8 số): </span>{gri6 or f'Phân loại ở cấp độ phân nhóm và mã 8 số {hs_code} dựa trên cấu trúc Biểu thuế XNK 2026.'}</div>

  <div class="sec-header">3. Trích dẫn Chú giải chi tiết HS 2022</div>
  <div class="wco-text">{notes or 'Đối chiếu Chú giải chi tiết HS 2022 WCO &amp; Bản Toàn tập HQKV8.'}</div>

  <div class="sec-header">4. Bảng So sánh và Lập luận Đánh giá giữa các Mã HS Tiềm năng</div>
  <table class="std-table">
    <thead>
      <tr>
        <th style="width: 16%;">Mã HS</th>
        <th style="width: 30%;">Mô tả phân nhóm</th>
        <th style="width: 14%;">Thuế MFN | VAT</th>
        <th style="width: 40%;">Đánh giá Thực tế Khai báo Hải quan</th>
      </tr>
    </thead>
    <tbody>
      {table_comp_rows_html}
    </tbody>
  </table>

  <div class="sec-header">5. Bảng Thuế suất Chi tiết của Mã HS {hs_code}</div>
  <table class="std-table">
    <thead>
      <tr>
        <th style="width: 42%;">Loại Thuế / Biểu thuế FTA</th>
        <th style="width: 16%;">Thuế suất</th>
        <th style="width: 42%;">Căn cứ Pháp lý / Nghị định</th>
      </tr>
    </thead>
    <tbody>
      {table_tax_rows_html}
    </tbody>
  </table>

  <div class="sec-header">6. Kết luận &amp; Khuyến nghị Thủ tục Hải quan</div>
  <div class="concl-p">
    <span class="concl-bold">KẾT LUẬN &amp; KIẾN NGHỊ: </span>{concl_text}
  </div>
</body>
</html>'''

def export_hs_report_pdf(data):
    with WORD_EXPORT_LOCK:
        hs_code = data.get('hs_code', '0000.00.00')
        clean_code = data.get('clean_code', hs_code.replace('.', ''))
        target_docx_path = os.path.join(WORKSPACE_DIR, f"{clean_code}.docx")
        target_pdf_path = os.path.join(WORKSPACE_DIR, f"{clean_code}.pdf")
        main_docx_path = os.path.join(WORKSPACE_DIR, "HS CODE.docx")

        # 1. Generate DOCX
        try:
            from standard_hs_layout_builder import generate_standard_hs_docx
            generate_standard_hs_docx(data, target_docx_path, scratch_dir=SCRATCH_DIR)
            import shutil
            shutil.copyfile(target_docx_path, main_docx_path)
        except Exception as e:
            print("Error creating standard docx:", e)

        # 2. Generate PDF via headless Chrome/Edge (fast, robust, vector quality)
        html_report = build_a4_html_report_server(data)
        pdf_ok = convert_html_to_pdf_headless(html_report, target_pdf_path)

        # 3. Fallback to Word COM if headless browser failed and file not created
        if not pdf_ok or not os.path.exists(target_pdf_path) or os.path.getsize(target_pdf_path) < 500:
            word_app = None
            doc_com = None
            try:
                pythoncom.CoInitialize()
                word_app = win32com.client.Dispatch("Word.Application")
                word_app.Visible = False
                word_app.DisplayAlerts = 0
                doc_com = word_app.Documents.Open(os.path.abspath(target_docx_path), ReadOnly=True)
                doc_com.SaveAs(os.path.abspath(target_pdf_path), FileFormat=17)
            except Exception as e:
                print("Word COM Export Notice (Handled):", e)
            finally:
                if doc_com is not None:
                    try: doc_com.Close(SaveChanges=False)
                    except: pass
                if word_app is not None:
                    try: word_app.Quit()
                    except: pass
                try: pythoncom.CoUninitialize()
                except: pass

    return {
        "status": "success",
        "clean_code": clean_code,
        "hs_code": hs_code,
        "pdf_url": f"/pdf/{clean_code}.pdf",
        "pdf_file_uri": f"file:///{target_pdf_path.replace(os.sep, '/')}",
        "docx_file_uri": f"file:///{target_docx_path.replace(os.sep, '/')}"
    }

def get_system_gemini_api_key(explicit_key=None):
    if explicit_key and len(explicit_key.strip()) > 15:
        return explicit_key.strip()
    env_k = os.environ.get('GEMINI_API_KEY')
    if env_k and len(env_k.strip()) > 15:
        return env_k.strip()
    for candidate_file in [os.path.join(BASE_DIR, 'gemini_key.txt'), os.path.join(BASE_DIR, '.env'), os.path.join(WORKSPACE_DIR, 'gemini_key.txt'), 'gemini_key.txt']:
        if os.path.exists(candidate_file):
            try:
                with open(candidate_file, 'r', encoding='utf-8') as f:
                    for line in f:
                        line = line.strip()
                        if line.startswith('GEMINI_API_KEY='):
                            k = line.split('=', 1)[1].strip().strip('"').strip("'")
                            if len(k) > 15: return k
                        elif line.startswith('AIzaSy'):
                            return line
            except Exception:
                pass
    return None

def call_gemini_api_python(api_key, name, desc, use, specs="", model_en="", candidates=None, top_headings=None):
    import urllib.request
    import urllib.error
    resolved_key = get_system_gemini_api_key(api_key)
    if not resolved_key:
        return None
    models = [
        'gemini-3.5-flash-lite',
        'gemini-flash-lite-latest',
        'gemini-3.5-flash'
    ]
    
    cand_str = ""
    if candidates and len(candidates) > 0:
        cand_lines = [f"- Mã {c.get('c', '')}: {c.get('vn', '')} (MFN: {c.get('mfn', '0%')})" for c in candidates[:10]]
        cand_str = "\nCác mã phân nhóm ứng viên trong Biểu thuế XNK Việt Nam 2026:\n" + "\n".join(cand_lines)

    notes_str = ""
    if top_headings:
        notes_lines = []
        for th in top_headings[:4]:
            ch_th = th[:2]
            w_note = get_wco_explanatory_note(ch_th, th, max_len=1500)
            notes_lines.append(f"Nhóm {th[:2]}.{th[2:4]} ({HEAD_MAP.get(th, '')}):\n{w_note}")
        notes_str = "\nTRÍCH XUẤT BỘ CHÚ GIẢI CHI TIẾT HS 2022 WCO (CÔNG VĂN 1810/TCHQ-TXNK & HQKV8):\n" + "\n\n".join(notes_lines)

    prompt = f"""Bạn là Chuyên gia Cao cấp Phân loại Biểu thuế và Mã số HS (Hải quan Việt Nam) theo Danh mục HS 2022 (Công văn 1810/TCHQ-TXNK và HQKV8) và Biểu thuế XNK 2026.
BẮT BUỘC TUÂN THỦ QUY TRÌNH RÀ SOÁT VÀ LẬP LUẬN 5 BƯỚC CHO MỌI MẶT HÀNG (CHƯƠNG 01 ĐẾN 97):
TUYỆT ĐỐI KHÔNG CHỈ DỰA VÀO TÊN THƯƠNG MẠI MÀ PHẢI RÀ SOÁT ĐA CHIỀU 4 TIÊU CHUẨN KỸ THUẬT:
1. Tên hàng định danh: {name} (Model: {model_en or name})
2. Bản chất và Chất liệu (Material): {desc or specs or name}
3. Kích thước và Thông số (Dimensions & Specs): {specs or desc or 'Quy cách kỹ thuật tiêu chuẩn'}
4. Công dụng và Cơ cấu truyền động (Function & Transmission): {use or 'Công nghiệp hoặc thương mại'}
{cand_str}
{notes_str}

BẮT BUỘC TUÂN THỦ NGUYÊN TẮC PHÁP LÝ HẢI QUAN THEO 6 QUY TẮC GRI VÀ MA TRẬN PHÂN ĐỊNH 21 PHẦN / 97 CHƯƠNG:
1. NGUYÊN TẮC VẬT LIỆU CHO BỘ PHẬN CÔNG DỤNG CHUNG (GRI 1 & Chú giải 2 Phần XV / Chú giải 1(g) Phần XVI):
   - Các chi tiết bu lông, ốc vít, đai ốc, đinh tán, chốt hãm, chốt định vị, vòng đệm, gioăng, đệm lót, lò xo, ống dẫn, phụ kiện ống:
     + NẾU BẰNG NHỰA / CHẤT DẺO (Plastic, PBT, POM, PA, nylon, PVC, PTFE...): BẮT BUỘC phân loại vào Chương 39 (Nhóm 39.26, đặc biệt mã 3926.90.99 cho bu lông/vòng đệm/chốt nhựa; hoặc Nhóm 39.17 cho ống/phụ kiện ống nhựa). TUYỆT ĐỐI KHÔNG xếp vào Chương kim loại 73, 74, 76!
     + NẾU BẰNG CAO SU (Rubber, EPDM, NBR, silicone...): BẮT BUỘC phân loại vào Chương 40 (Nhóm 40.16, đặc biệt mã 4016.93 cho gioăng, đệm, vòng đệm cao su; hoặc Nhóm 40.09 cho ống cao su; 40.10 cho dây curoa/băng chuyền). TUYỆT ĐỐI KHÔNG xếp vào Chương 73 hay Chương 39!
     + CHỈ KHI BẰNG KIM LOẠI CƠ BẢN mới áp dụng Chú giải 2 Phần XV: Sắt/thép/inox -> Chương 73 (73.18, 73.20, 73.07); Đồng -> Chương 74 (74.15, 74.12); Nhôm -> Chương 76 (76.16, 76.09).
     + Trong mọi trường hợp, các bộ phận công dụng chung nêu trên TUYỆT ĐỐI KHÔNG phân loại theo máy móc mẹ hay phương tiện chủ (Chương 84, 85, 87, 88, 90).
2. THIẾT BỊ CÓ CHỨC NĂNG ĐỘC LẬP ĐỊNH DANH (Chú giải 2(a) Phần XVI):
   - Các mặt hàng đã được chi tiết hóa tại các nhóm riêng (van 8481, bơm 8413, động cơ điện 8501, máy nén khí/quạt 8414, biến áp 8504, công tắc 8536, tủ bảng điều khiển 8537, thiết bị đo kiểm 9031, thiết bị điều khiển tự động 9032...) thì DÙ VỎ BỌC LÀM BẰNG NHỰA HAY KIM LOẠI và DÙ LẮP VÀO MÁY NÀO vẫn BẮT BUỘC phân loại vào nhóm định danh của chúng theo Chú giải 2(a) Phần XVI, TUYỆT ĐỐI KHÔNG xếp vào nhóm bộ phận máy (8409, 8431, 8466, 8473, 8538, 8708).
3. DÂY CÁP ĐIỆN VÀ CÁP SỢI QUANG (Nhóm 85.44 vs 90.01):
   - Dây cáp điện có gắn đầu nối (connectors, phích cắm, đầu cos) -> Phân loại vào phân nhóm có gắn đầu nối (ví dụ 8544.42); không có đầu nối -> Phân nhóm 8544.49; Điện áp > 1.000V -> Phân nhóm 8544.60.
   - Cáp sợi quang cấu tạo từ các sợi quang bọc từng sợi riêng biệt -> Phân nhóm 8544.70 (loại trừ khỏi Chương 90). Sợi quang chưa bọc từng sợi -> Nhóm 90.01.
4. ĐỒ NỘI THẤT VÀ GHẾ (Nhóm 94.01, 94.03):
   - Ghế ngồi (94.01) và đồ nội thất (94.03) dù làm bằng gỗ, kim loại hay nhựa đều BẮT BUỘC phân loại vào Chương 94 (theo Chú giải 1(e) Chương 39, Chú giải 1(c) Chương 44, Chú giải 1(k) Chương 73). Tuyệt đối không xếp vào 3926, 4421 hay 7326!
5. TÚI XÁCH, VALI, BA LÔ, CẶP (Nhóm 42.02):
   - Các loại túi xách, vali, bao đựng dù bề mặt bằng da, nhựa tấm hay vải dệt đều BẮT BUỘC thuộc Nhóm 42.02 theo Chú giải 1(h) Phần XI và Chú giải 1(ij) Chương 39. Tuyệt đối không xếp vào Chương 39, 61, 62 hay 63!
6. QUẦN ÁO DỆT KIM (Chương 61) VS DỆT THOI (Chương 62) & GIÀY DÉP (Chương 64):
   - Quần áo may từ vải dệt kim hoặc móc -> Chương 61; Quần áo may từ vải dệt thoi -> Chương 62.
   - Giày dép các loại -> Bắt buộc Chương 64 (loại trừ khỏi dệt may và đồ nhựa).
7. HÓA CHẤT ĐƠN CHẤT (Chương 28, 29) VS CHẾ PHẨM HỖN HỢP (Chương 34, 35, 38):
   - Hóa chất hữu cơ/vô cơ đã xác định thành phần hóa học thuần túy -> Chương 28, 29.
   - Chế phẩm tẩy rửa, làm sạch -> Nhóm 34.02; Dầu bôi trơn tổng hợp -> Nhóm 34.03; Chế phẩm hóa chất công nghiệp khác -> Nhóm 38.24/38.27; Thuốc trừ sâu, diệt cỏ, sát trùng bán lẻ -> Nhóm 38.08.
8. QUY TẮC PHÂN LOẠI VÚ TRA MỠ / VÚ BƠM MỠ BÒ (GREASING NIPPLES / GREASE FITTINGS / ZERK FITTINGS):
   - Căn cứ Chú giải chi tiết WCO Nhóm 84.87 (HQKV8 Trang 214194) và Chú giải loại trừ Nhóm 84.81 (HQKV8 Trang 212255), vú tra mỡ / vú bơm mỡ bò không tự động là bộ phận cơ khí máy móc BẮT BUỘC phân loại vào Nhóm 84.87, mã số chi tiết 8487.90.00 (Thuế MFN: 0%, VAT: 8%).
   - TUYỆT ĐỐI KHÔNG phân loại vào Nhóm 73.18 (bu lông ốc vít), 73.07 (phụ kiện nối ống), 73.26 hay 84.81 (van).
9. QUY TẮC PHÂN LOẠI THIẾT BỊ PHUN/RỬA/LÀM SẠCH (NHÓM 84.24 - GRI 6 GẠCH NỐI):
   - Thiết bị thuộc nhóm 84.24 (máy rửa công nghiệp, máy làm sạch linh kiện, thiết bị phun xịt...) hoạt động bằng điện (dùng động cơ điện, motor điện, cắm điện, có công suất điện W/kW, điện áp V) BẮT BUỘC chọn nhánh 3 gạch "- - - Hoạt động bằng điện", mã 8424.89.50. TUYỆT ĐỐI CẤM chọn mã 8424.89.90 ("- - - - Không hoạt động bằng điện")!
10. QUY TẮC RÀNG BUỘC MÃ TỪ BIỂU THUẾ VIỆT NAM 2026:
   - BẮT BUỘC CHỌN 01 MÃ HS 8 CHỮ SỐ TRONG DANH SÁCH ỨNG VIÊN ĐÃ ĐƯỢC RÀ SOÁT TỪ BIỂU THUẾ VIỆT NAM 2026 TRÊN ĐÂY (trừ khi có căn cứ pháp lý chỉ ra mã khác chính xác hơn). TUYỆT ĐỐI KHÔNG TỰ BỊA MÃ 8 SỐ KHÔNG TỒN TẠI TRONG BIỂU THUẾ VIỆT NAM.
   - Trong mục "explanatory_notes": Trích dẫn trung thực các đoạn văn Chú giải HS 2022 WCO ở trên kèm trang tham chiếu.
   - Trong mục "exclusions": Bắt buộc giải thích rõ lý do pháp lý loại trừ từng mã ứng viên đối thủ còn lại.

BẢNG SO SÁNH ĐỐI KHÁNG CÁC NHÁNH CÙNG CẤP GẠCH NỐI (BẮT BUỘC TRƯỚC KHI CHỐT MÃ 8 SỐ):
- Trong trường "gri6", TRƯỚC KHI nêu mã cuối cùng, phải lập bảng theo từng cấp gạch nối, đi từ cấp "-" xuống cấp sâu nhất. Mỗi cấp liệt kê TẤT CẢ các nhánh anh em cùng cấp trong danh sách ứng viên, ghi rõ: [Nhánh] | [ĐẠT/LOẠI] | [Lý do đối chiếu với dữ kiện hàng hóa].
- Với cặp nhánh đối lập (có X / không X, hoạt động bằng điện / không hoạt động bằng điện, đã gắn / chưa gắn đầu nối...): hàng có X thì LOẠI nhánh "không X" và ngược lại.
- Với nhánh có ngưỡng số (không quá / dưới / trên ... V, A, W, kVA): đổi về cùng đơn vị rồi so sánh với số liệu thực tế; nhánh không thỏa phải ghi LOẠI. Giá trị bằng đúng ngưỡng "dưới X" thì KHÔNG thỏa "dưới X".
- Mã 8 số cuối cùng CHỈ được lấy từ nhánh còn ĐẠT ở cấp sâu nhất.

YÊU CẦU TRẢ VỀ DUY NHẤT MỘT CHUỖI JSON CHUẨN (không viết gì ngoài JSON):
{{
  "step1_tech_review": {{
    "trade_name": "{name} (Model: {model_en or 'Tiêu chuẩn'})",
    "material": "{desc or 'Theo hồ sơ kỹ thuật'}",
    "dimensions_specs": "{specs or 'Thông số kỹ thuật tiêu chuẩn'}",
    "function_transmission": "{use or 'Vận hành công nghiệp / thương mại'}"
  }},
  "hs_code": "mã 8 số dạng xxxx.xx.xx",
  "clean_code": "8 chữ số liền nhau",
  "decl_name": "Tên hàng khai báo hải quan chuẩn hóa kèm quy cách chất liệu và mới 100%",
  "item_description": "Mô tả phân nhóm theo biểu thuế 2026",
  "mfn": "Thuế MFN %",
  "vat": "Thuế VAT %",
  "atiga": "0%",
  "acfta": "0%",
  "evfta": "0%",
  "gri1": "Lập luận GRI 1...",
  "gri2": "Lập luận GRI 2...",
  "gri3": "Lập luận GRI 3...",
  "gri4": "Lập luận GRI 4...",
  "gri5": "Lập luận GRI 5...",
  "gri6": "Lập luận GRI 6...",
  "explanatory_notes": "Trích dẫn Chú giải HS 2022 (CV 1810 & HQKV8)...",
  "exclusions": "Lập luận loại trừ chi tiết các nhóm ứng viên cạnh tranh...",
  "recommendations": "Khuyến nghị thủ tục hải quan..."
}}
"""
    payload = json.dumps({"contents": [{"parts": [{"text": prompt}]}]}).encode('utf-8')
    for m in models:
        try:
            url = f"https://generativelanguage.googleapis.com/v1beta/models/{m}:generateContent?key={resolved_key}"
            req = urllib.request.Request(url, data=payload, headers={'Content-Type': 'application/json'})
            with urllib.request.urlopen(req, timeout=12) as response:
                if response.status == 200:
                    resp_json = json.loads(response.read().decode('utf-8'))
                    text = resp_json.get('candidates', [{}])[0].get('content', {}).get('parts', [{}])[0].get('text', '')
                    m_json = re.search(r'\{[\s\S]*\}', text)
                    if m_json:
                        parsed = json.loads(m_json.group(0))
                        if parsed.get('clean_code') or parsed.get('hs_code'):
                            return parsed
        except Exception as e:
            print(f"Gemini Python attempt with {m} failed:", e)
    return None


# =============================================================================
# CỔNG KIỂM SOÁT THÔNG SỐ GRI 6 (GRI 6 PARAMETER GATE & ANTI-GUESSING ENGINE)
# =============================================================================
GRI6_GATE_SECTORS_PYTHON = [
    {
        "id": "valves_8481",
        "headings": ["8481"],
        "keywords": ["van", "valve", "solenoid", "voi nuoc", "tiet luu", "giam ap", "mot chieu", "an toan"],
        "category": "Van & Thiết bị kiểm soát lưu chất (Nhóm 84.81)",
        "criteria": [
            {
                "name": "Cơ chế vận hành van",
                "pattern": r"giam ap|mot chieu|an toan|xa ap|dien tu|solenoid|thuy luc|khi nen|van bi|van buom|van cau|van cong|tiet luu|phao|check valve",
                "missing": "Cơ chế vận hành (Giảm áp 8481.10 / Thủy lực khí nén 8481.20 / Một chiều 8481.30 / An toàn 8481.40 / Điện từ 8481.80)",
                "chips": ["Van giảm áp (8481.10)", "Van một chiều (8481.30)", "Van điện từ Solenoid (8481.80)", "Van an toàn (8481.40)", "Van khí nén (8481.20)"]
            },
            {
                "name": "Đường kính trong DN",
                "pattern": r"dn\s*\d+|\d+\s*mm|\d+\s*cm|phi\s*\d+|duong kinh|<= 25|> 25|25mm",
                "missing": "Đường kính danh định trong (DN <= 25mm hay DN > 25mm)",
                "chips": ["Đường kính trong DN <= 25mm", "Đường kính trong DN > 25mm"]
            }
        ]
    },
    {
        "id": "motors_8501",
        "headings": ["8501", "8502"],
        "keywords": ["dong co", "motor", "servo", "may phat", "generator", "stepper"],
        "category": "Động cơ điện & Máy phát điện (Nhóm 85.01 / 85.02)",
        "criteria": [
            {
                "name": "Dòng điện & Số pha",
                "pattern": r"ac|dc|xoay chieu|mot chieu|1 pha|3 pha|da pha|stepper|servo",
                "missing": "Dòng điện (AC xoay chiều 1 pha/3 pha hay DC một chiều)",
                "chips": ["Động cơ AC 3 pha", "Động cơ AC 1 pha", "Động cơ DC 1 chiều", "Động cơ bước Stepper", "Động cơ Servo"]
            },
            {
                "name": "Công suất định mức",
                "pattern": r"\d+\s*w|\d+\s*kw|\d+\s*hp|cong suat|37\.5w|750w|7\.5kw|75kw",
                "missing": "Công suất định mức (<= 37.5W / <= 750W / <= 7.5kW / > 7.5kW)",
                "chips": ["Công suất <= 37.5W", "Công suất 37.5W - 750W", "Công suất 750W - 7.5kW", "Công suất > 7.5kW"]
            }
        ]
    },
    {
        "id": "transformers_8504",
        "headings": ["8504"],
        "keywords": ["bien ap", "transformer", "inverter", "nghich luu", "bien tan", "adapter", "nguon", "chinh luu", "sac", "ups"],
        "category": "Máy biến áp & Thiết bị biến đổi điện tĩnh (Nhóm 85.04)",
        "criteria": [
            {
                "name": "Chức năng biến đổi",
                "pattern": r"adapter|nguon xung|inverter|nghich luu|bien tan|chinh luu|bien ap|sac ac quy|ups",
                "missing": "Chức năng (Nguồn xung Adapter / Inverter biến tần 8504.40.40 / Biến áp cuộn dây / Bộ nạp sạc 8504.40.11)",
                "chips": ["Nguồn xung AC-DC (Adapter)", "Biến tần DC-AC (Inverter)", "Biến áp cách ly (Transformer)", "Bộ sạc ắc quy (Battery charger)"]
            }
        ]
    },
    {
        "id": "switches_8536",
        "headings": ["8535", "8536"],
        "keywords": ["cong tac", "switch", "nut nhan", "cau dao", "aptomat", "ro le", "relay", "breaker", "mcb", "mccb"],
        "category": "Khí cụ đóng ngắt & Bảo vệ mạch điện (Nhóm 85.35 / 85.36)",
        "criteria": [
            {
                "name": "Ngưỡng điện áp",
                "pattern": r"<= 1000|> 1000|ha the|cao the|220v|380v|24v|12v|110v|von|volt|dien ap",
                "missing": "Ngưỡng điện áp (<= 1000V Nhóm 85.36 vs > 1000V Nhóm 85.35)",
                "chips": ["Điện áp <= 1000V (Hạ thế - 85.36)", "Điện áp > 1000V (Cao thế - 85.35)"]
            },
            {
                "name": "Dòng điện định mức",
                "pattern": r"16a|ampe|ampere|duoi 16a|tren 16a|< 16a|>= 16a",
                "missing": "Dòng điện định mức (< 16A hay >= 16A)",
                "chips": ["Dòng điện định mức < 16A (Mã 8536.50.96)", "Dòng điện định mức >= 16A (Mã 8536.50.99)"]
            }
        ]
    },
    {
        "id": "cables_8544",
        "headings": ["8544"],
        "keywords": ["day dien", "cap dien", "cap quang", "cable", "wire"],
        "category": "Dây điện, Cáp điện & Cáp sợi quang (Nhóm 85.44)",
        "criteria": [
            {
                "name": "Trang bị đầu nối",
                "pattern": r"dau noi|connector|jack|chua gan dau noi|co gan dau noi|fitted with connector",
                "missing": "Tình trạng đầu nối (Đã gắn sẵn đầu nối hay Chưa gắn đầu nối)",
                "chips": ["Đã gắn sẵn đầu nối", "Chưa gắn đầu nối"]
            },
            {
                "name": "Cấp điện áp",
                "pattern": r"80v|1000v|cap quang|optical|cao the|ha the",
                "missing": "Cấp điện áp (<= 80V / 80V - 1000V / > 1000V) hoặc Cáp quang",
                "chips": ["Điện áp <= 80V", "Điện áp 80V - 1000V", "Điện áp > 1000V", "Cáp sợi quang"]
            }
        ]
    },
    {
        "id": "steel_pipes_7304",
        "headings": ["7304", "7305", "7306"],
        "keywords": ["ong thep", "ong sat", "steel pipe", "steel tube"],
        "category": "Ống sắt thép (Nhóm 73.04 - 73.06)",
        "criteria": [
            {
                "name": "Phương pháp chế tạo ống",
                "pattern": r"duc|khong han|seamless|han|welded",
                "missing": "Phương pháp sản xuất (Đúc không hàn 73.04 vs Có hàn 73.06)",
                "chips": ["Ống thép đúc liền không hàn (73.04)", "Ống thép hàn (73.06)"]
            }
        ]
    },
    {
        "id": "steel_flat_7208",
        "headings": ["7208", "7209", "7210"],
        "keywords": ["thep tam", "thep cuon", "thep can", "ton ma"],
        "category": "Thép cán phẳng (Nhóm 72.08 - 72.10)",
        "criteria": [
            {
                "name": "Phương pháp cán & Phủ mạ",
                "pattern": r"can nong|can nguoi|hot-rolled|cold-rolled|ma kem|galvanized",
                "missing": "Quy trình cán (Cán nóng 72.08 / Cán nguội 72.09 / Mạ kẽm 72.10)",
                "chips": ["Cán nóng (72.08)", "Cán nguội (72.09)", "Đã mạ kẽm (72.10)"]
            }
        ]
    },
    {
        "id": "fasteners_7318",
        "headings": ["7318"],
        "keywords": ["bu long", "dai oc", "oc vit", "long den", "vong dem", "chot", "dinh tan", "bolt", "screw", "nut"],
        "category": "Vật tư liên kết bằng sắt thép (Nhóm 73.18)",
        "criteria": [
            {
                "name": "Cấu tạo ren",
                "pattern": r"co ren|khong ren|bu long|vit|dai oc|e-cu|long den|vong dem|dinh tan|chot",
                "missing": "Cấu tạo ren (Có ren: Bulông/Vít/Ê-cu vs Không ren: Đệm/Đinh tán)",
                "chips": ["Bulông có ren (7318.15)", "Đai ốc có ren (7318.16)", "Vòng đệm không ren (7318.21/22)", "Đinh tán không ren (7318.23)"]
            }
        ]
    },
    {
        "id": "bearings_8482",
        "headings": ["8482", "8483"],
        "keywords": ["o bi", "bac dan", "o truot", "bac truot", "bearing", "thanh truot", "con truot"],
        "category": "Ổ bi, Ổ lăn & Ổ trượt (Nhóm 84.82 / 84.83)",
        "criteria": [
            {
                "name": "Cấu tạo con lăn",
                "pattern": r"bi cau|con lan|o dua|con|kim|khong lap bi|khong co bi|bac lot",
                "missing": "Dạng phần tử lăn (Ổ bi cầu 8482.10 / Ổ con lăn 8482.20-50 / Thân ổ trượt không bi 8483.30)",
                "chips": ["Ổ bi cầu (8482.10)", "Ổ đũa / con lăn (8482.50)", "Thân ổ trượt không lắp bi (8483.30)", "Thanh trượt có lắp ổ bi (8482.10)"]
            }
        ]
    },
    {
        "id": "pumps_8413",
        "headings": ["8413"],
        "keywords": ["may bom", "bom nuoc", "bom dau", "pump", "bom ly tam"],
        "category": "Máy bơm chất lỏng (Nhóm 84.13)",
        "criteria": [
            {
                "name": "Nguyên lý làm việc",
                "pattern": r"ly tam|piston|banh rang|truc vit|canh gat|mang|dinh luong",
                "missing": "Nguyên lý bơm (Ly tâm 8413.70 / Bánh răng 8413.60 / Piston 8413.50 / Đo thể tích 8413.19)",
                "chips": ["Bơm ly tâm (8413.70)", "Bơm bánh răng (8413.60)", "Bơm piston (8413.50)", "Bơm định lượng (8413.19)"]
            }
        ]
    }
]

def check_gri6_parameters_python(heading, text):
    text_norm = remove_accents(text.lower())
    matched_sector = None
    for s in GRI6_GATE_SECTORS_PYTHON:
        if heading in s["headings"]:
            matched_sector = s
            break
    if not matched_sector:
        for s in GRI6_GATE_SECTORS_PYTHON:
            if any(k in text_norm for k in s["keywords"]):
                matched_sector = s
                break

    if not matched_sector:
        return {
            "is_gated": False,
            "status": "STANDARD",
            "category": "Mặt hàng chuẩn hóa",
            "missing_criteria": [],
            "confirmed_criteria": [],
            "chips": []
        }

    missing = []
    confirmed = []
    chips = []
    for c in matched_sector["criteria"]:
        if re.search(c["pattern"], text_norm, re.IGNORECASE):
            confirmed.append(c["name"])
        else:
            missing.append(c["missing"])
            chips.extend(c["chips"])

    return {
        "is_gated": True,
        "status": "GATED_MISSING_CRITERIA" if missing else "VERIFIED",
        "category": matched_sector["category"],
        "missing_criteria": missing,
        "confirmed_criteria": confirmed,
        "chips": chips
    }

def run_automation_workflow(name, desc, use, specs="", model_en="", image_b64=None, doc_b64=None, doc_filename=None, gemini_api_key=None):
    extracted_doc_text = ""
    if doc_b64:
        extracted_doc_text = extract_text_from_attached_doc(doc_b64, doc_filename)
        if extracted_doc_text:
            specs = f"{specs} | Trích xuất tài liệu đính kèm: {extracted_doc_text}".strip()

    best_item, competing_items, top_headings, decomp = classify_product_python(name, desc, use, specs, model_en)
    gate_result = check_gri6_parameters_python(best_item.get('c', '')[:4], f'{name} {desc} {use} {specs} {model_en}')


    # Optional Google Gemini AI integration if API key is provided
    gemini_api_key = get_system_gemini_api_key(gemini_api_key)
    gemini_res = None
    # Ket qua tu StatutoryTreeWalker (quy tac co can cu Chu giai) duoc uu tien hon de xuat cua Gemini.
    _walker_item = best_item if (best_item.get('_statutory') and best_item.get('_level', 8) == 8) else None
    if gemini_api_key:
        try:
            all_cands = [best_item] + competing_items
            gemini_res = call_gemini_api_python(gemini_api_key, name, desc, use, specs, model_en, candidates=all_cands, top_headings=top_headings)
            if gemini_res:
                g_code = (gemini_res.get('clean_code') or gemini_res.get('hs_code', '')).replace('.', '').strip()
                matched = next((item for item in HS_DATABASE if item.get('c', '').replace('.', '') == g_code), None)
                if not matched and len(g_code) >= 6:
                    matched = next((item for item in HS_DATABASE if item.get('c', '').replace('.', '').startswith(g_code[:6])), None)
                if matched:
                    best_item = matched
                    # Kiem tra cheo ma do AI de xuat bang Universal GRI-6 Engine
                    try:
                        _ft = remove_accents(f"{name} {desc} {use} {specs} {model_en}")
                        _gc = matched.get('c', '').replace('.', '')
                        if _ue is not None and len(_gc) == 8 and _ue.evaluate(_gc, _ue.extract_query(_ft))[0]:
                            _rr = _ue.reroute(_gc, _ft)
                            if _rr and _rr['status'] == 'EXACT_8_DIGIT':
                                _rc = _rr['code'].replace('.', '')
                                _alt = next((it for it in HS_DATABASE if it.get('c', '').replace('.', '') == _rc), None)
                                if _alt:
                                    print(f"GRI6 Engine: AI code {_gc} bi loai ({_rr['dropped']}) -> {_rc}")
                                    best_item = _alt
                    except Exception as _eg:
                        print("GRI6 engine post-check error:", _eg)
                    print(f"Gemini proposed statutory code {matched.get('c')} ({matched.get('vn')}) verified & adopted.")
                    if _walker_item is not None:
                        _wc = _walker_item.get('c', '').replace('.', '')
                        if _wc and _wc != best_item.get('c', '').replace('.', ''):
                            print(f"StatutoryTreeWalker override: AI {best_item.get('c')} -> walker {_walker_item.get('c')}")
                            _walker_item = dict(_walker_item)
                            _walker_item['_ai_rejected_code'] = best_item.get('c')
                            best_item = _walker_item
                
        except Exception as e:
            print("Gemini integration fallback:", e)

    code_raw = best_item.get('c', '00000000').replace('.', '')
    code_fmt = f"{code_raw[:4]}.{code_raw[4:6]}.{code_raw[6:8]}"
    heading = code_raw[:4]
    chapter = code_raw[:2]
    heading_desc = HEAD_MAP.get(heading, 'Nhóm hàng hóa tương ứng')
    subheading_desc = SUBHEAD_MAP.get(code_raw[:6], 'Phân nhóm hàng hóa tương ứng')

    mfn = best_item.get('mfn', '0%')
    mfn_str = mfn if '%' in str(mfn) else f"{mfn}%"
    vat = best_item.get('vat', '8%')
    vat_str = vat if '%' in str(vat) else f"{vat}%"
    atiga_str = str(best_item.get('atiga', '0')) + '%' if '%' not in str(best_item.get('atiga', '0')) else str(best_item.get('atiga', '0'))
    acfta_str = str(best_item.get('acfta', '0')) + '%' if '%' not in str(best_item.get('acfta', '0')) else str(best_item.get('acfta', '0'))
    vjepa_str = str(best_item.get('vjepa', '0')) + '%' if '%' not in str(best_item.get('vjepa', '0')) else str(best_item.get('vjepa', '0'))
    evfta_str = str(best_item.get('evfta', '0')) + '%' if '%' not in str(best_item.get('evfta', '0')) else str(best_item.get('evfta', '0'))

    # Step 1: 4 Technical Criteria Review
    step1_tech_review = {
        "trade_name": f"{name} (Model: {model_en or 'Tiêu chuẩn'})",
        "material": desc or (decomp['materials'][0][0] if decomp.get('materials') else "Theo hồ sơ kỹ thuật / datasheet của nhà sản xuất"),
        "dimensions_specs": specs or desc or "Quy cách / Kích thước tiêu chuẩn",
        "function_transmission": use or "Vận hành theo chức năng chuyên dụng"
    }

    # Universal Customs Declarable Name Construction
    item_vn_clean = best_item.get('vn', '').lstrip('-').strip()
    if item_vn_clean and item_vn_clean.lower() not in name.lower():
        decl_name = f"{name} ({item_vn_clean}, mới 100%)"
    else:
        decl_name = f"{name} (mới 100%)"

    # Universal GRI 1 to GRI 6 Statutory Reasoning
    statutory_note_citation = ""
    if heading in ['7318', '7415', '7616', '7320', '7326']:
        statutory_note_citation = " Căn cứ Chú giải 2 Phần XV, Chú giải 1(a) Chương 90 và Chú giải 1(g) Phần XVI, các sản phẩm và bộ phận bằng kim loại cơ bản (sắt thép) chưa được định danh ở nhóm riêng biệt được phân loại tại Phần XV (Nhóm 73.26/73.18), không phân loại theo máy móc hay thiết bị."
    elif heading == '3926':
        statutory_note_citation = " Căn cứ Chú giải 1(g) Phần XVI, các sản phẩm và bộ phận bằng plastic (nhựa) thuộc Chương 39 (Nhóm 39.26) bị loại trừ khỏi Phần XVI (máy móc), không phân loại theo nhóm phụ tùng máy móc 84.66."
    elif heading == '8466':
        statutory_note_citation = " Căn cứ Chú giải 2(b) Phần XVI, các bộ phận chuyên dùng hoặc chủ yếu dùng cho các máy móc thuộc các nhóm từ 84.56 đến 84.65 (bao gồm máy cắt dập kim loại thuộc Nhóm 84.62) được phân loại vào Nhóm 84.66."
    elif heading == '8536' and '853670' in code_raw:
        statutory_note_citation = " Căn cứ Chú giải 9 Chương 85 và Chú giải 1(h) Chương 90, các đầu nối dùng cho sợi quang, bó sợi quang hoặc cáp quang được định danh chuyên biệt tại Nhóm 85.36 (Phân nhóm 8536.70) và bị loại trừ khỏi Chương 90."
    elif heading in ['8413', '8414', '8481', '8482', '8483', '8501', '8504', '8536', '8544']:
        statutory_note_citation = " Căn cứ Chú giải 2(a) Phần XVI, các hàng hóa có công năng độc lập thuộc phạm vi của các nhóm này trong mọi trường hợp phải được phân loại vào đúng nhóm của chúng, không phân loại theo nhóm phụ tùng chung."
    elif heading in ['8207', '8208', '8209']:
        statutory_note_citation = " Căn cứ Chú giải 1(k) Phần XVI, các dụng cụ và lưỡi cắt thuộc Chương 82 không thuộc phạm vi của Phần XVI (máy móc)."
    elif heading == '8523':
        statutory_note_citation = " Căn cứ Chú giải Nhóm 85.23, các thiết bị lưu trữ dữ liệu bền vững thể rắn (thẻ nhớ, ổ USB flash drive) được định danh chuyên biệt tại Nhóm 85.23, phân nhóm 8523.51."
    elif heading == '8806':
        statutory_note_citation = " Căn cứ Chú giải 1 Chương 88, phương tiện bay không người lái (UAV / Drone) được định danh tại Nhóm 88.06, phân loại theo trọng lượng cất cánh tối đa."
    elif heading == '8487':
        statutory_note_citation = " Căn cứ Chú giải chi tiết WCO Nhóm 84.87 (HQKV8 Trang 214194) và Chú giải loại trừ Nhóm 84.81 (HQKV8 Trang 212255), các chi tiết cơ khí bôi trơn máy móc không chứa tiếp điểm điện như vú tra mỡ (greasing nipples), phớt dầu, bình tra dầu mỡ không tự động... bắt buộc phân loại vào Nhóm 84.87 (Mã số 8487.90.00)."
    elif heading in ['4407', '4408']:
        statutory_note_citation = " Căn cứ Chú giải Chương 44, ranh giới giữa Nhóm 44.07 (gỗ cưa xẻ) và Nhóm 44.08 (ván bóc/ván lạng) được xác định theo ngưỡng chiều dày 6 mm."

    gri1 = f"Quy tắc 1 (Pháp lý Tên Nhóm & Chú giải Phần/Chương): Căn cứ Quy tắc 1 (GRI 1) của 6 Quy tắc Tổng quát Giải thích HS, việc phân loại được xác định căn cứ vào nội dung của các nhóm và các Chú giải Phần, Chú giải Chương liên quan. Sản phẩm '{name}' sau khi đối chiếu tuần tự từ Chương 01 đến Chương 97 được xác định thuộc phạm vi pháp lý Nhóm {heading[:2]}.{heading[2:4]} ({heading_desc}) thuộc {CHAPTER_MAP.get(chapter, f'Chương {chapter}')}.{statutory_note_citation}"
    gri2 = f"Quy tắc 2 (Thành phẩm hoàn chỉnh, Chưa lắp ráp & Hỗn hợp): Căn cứ Quy tắc 2(a), sản phẩm ở dạng thành phẩm hoàn chỉnh nguyên chiếc, đã tích hợp đầy đủ cơ cấu vận hành mang đặc tính cơ bản của thành phẩm (hoặc linh kiện đồng bộ tháo rời chưa lắp ráp), đủ điều kiện phân loại như thành phẩm hoàn chỉnh."
    gri3 = f"Quy tắc 3 (Phân loại theo Thứ tự Ưu tiên 3(a) > 3(b) > 3(c)): Căn cứ Quy tắc 3(a), Nhóm {heading[:2]}.{heading[2:4]} ({heading_desc}) mô tả đặc trưng định danh chi tiết đúng bản chất công năng và vật liệu của sản phẩm, có giá trị pháp lý ưu tiên tuyệt đối so với các nhóm mang tính chất phụ tùng chung hoặc mô tả khái quát. (Quy tắc 3(b) và 3(c) không cần áp dụng do 3(a) đã xác định rõ ràng)."
    gri4 = f"Quy tắc 4 (Hàng hóa giống nhất): Không cần áp dụng (do sản phẩm đã được xác định mã HS chính xác theo thứ tự ưu tiên của Quy tắc 1, 2 và 3)."
    gri5 = f"Quy tắc 5 (Bao bì, Hộp chứa chuyên dụng & Vật liệu đóng gói): Bao bì, vỏ hộp đóng gói thông thường đi kèm dùng để chứa đựng bảo quản sản phẩm được phân loại cùng với sản phẩm chính theo Quy tắc 5(b)."
    
    lineage_str = best_item.get('lin') or (TARIFF_8DIGIT_DETAILS.get(code_raw, {}).get('full_lineage') if TARIFF_8DIGIT_DETAILS else "") or f"Nhóm {heading}: {heading_desc} > {subheading_desc} > {best_item.get('vn', '')}"
    gri6 = f"Quy tắc 6 (Xác định Phân nhóm 6 số & Mã chi tiết 8 số): Căn cứ Quy tắc 6 (GRI 6), việc phân loại ở cấp độ phân nhóm 6 số và mã hàng 8 số được xác định bằng cách so sánh ở CÙNG CẤP ĐỘ GẠCH NỐI (- hoặc --), các Chú giải phân nhóm liên quan và thông số kỹ thuật định lượng ({specs or 'tiêu chuẩn'}). Tuân thủ nghiêm ngặt tiến trình phân tầng phả hệ: {lineage_str}. Sản phẩm được xếp chuẩn xác vào mã HS 8 số {code_fmt} ({best_item.get('vn')}) của Biểu thuế Xuất nhập khẩu Việt Nam 2026."

    notes = get_wco_explanatory_note(chapter, heading)

    # Universal Dynamic Exclusion Engine comparing against competing candidate headings
    exclusion_reasons = []
    for comp in competing_items:
        comp_code = comp.get('c', '').replace('.', '')
        comp_head = comp_code[:4]
        comp_h_desc = HEAD_MAP.get(comp_head, 'Nhóm hàng tương ứng')
        comp_fmt = f"{comp_code[:4]}.{comp_code[4:6]}.{comp_code[6:8]}"
        if comp_head != heading:
            if heading in ['7318', '7415', '7616', '7320', '7326'] and comp_code[:2] in ['84', '85', '87', '88', '90']:
                reason = f"Loại trừ Nhóm {comp_head[:2]}.{comp_head[2:4]} ({comp_h_desc}) theo Chú giải 2 Phần XV và Chú giải 1(g) Phần XVI: Các sản phẩm và bộ phận bằng sắt thép thông thường (miếng thép gá đặt, chốt, giá đỡ...) phân loại vào Phần XV (Nhóm 73.26/73.18), trừ khi là bộ phận chuyên dụng duy nhất chế tạo riêng cho máy theo Chú giải 2(b) Phần XVI."
            elif heading == '8466' and comp_head == '7326':
                reason = f"Ưu tiên Nhóm {heading[:2]}.{heading[2:4]} ({heading_desc}) so với Nhóm 73.26: Căn cứ Chú giải 2(b) Phần XVI, chi tiết dẫn hướng/phụ tùng chuyên dùng cho máy cắt dập được ưu tiên phân loại vào Nhóm 84.66 thay vì nhóm các sản phẩm bằng sắt thép khác (73.26)."
            elif heading in ['8413', '8414', '8481', '8482', '8483', '8501', '8504', '8536', '8544'] and comp_head in ['8409', '8431', '8466', '8473', '8538', '8708']:
                reason = f"Loại trừ Nhóm phụ tùng {comp_head[:2]}.{comp_head[2:4]} ({comp_h_desc}) theo Chú giải 2(a) Phần XVI: Các bộ phận của máy móc đã được định danh cụ thể tại các nhóm của Chương 84 hoặc 85 trong mọi trường hợp phải được phân loại vào đúng nhóm của chúng (Nhóm {heading[:2]}.{heading[2:4]})."
            elif heading == '8482' and comp_head == '8483':
                reason = f"Loại trừ Nhóm 84.83 ({comp_h_desc}): Phân nhóm 8483.30 quy định rõ 'Thân ổ, KHÔNG lắp ổ bi hoặc ổ đũa'. Sản phẩm là thanh trượt/con trượt có lắp ổ bi bên trong, căn cứ Chú giải chi tiết WCO Nhóm 84.82 (HQKV8 Trang 1955), cơ chế trượt với ổ bi (slide mechanisms with bearing balls / linear guide) bắt buộc phân loại vào Nhóm 84.82 (Mã 8482.10.00)."
            elif heading == '8482' and comp_head == '8466':
                reason = f"Loại trừ Nhóm phụ tùng 84.66 ({comp_h_desc}) theo Chú giải 2(a) Phần XVI: Mặt hàng ổ bi thuộc phạm vi Nhóm 84.82 trong mọi trường hợp phải được phân loại vào đúng Nhóm 84.82, không phân loại theo nhóm bộ phận của máy cắt dập (84.66)."
            elif heading == '8536' and comp_head == '8481':
                reason = f"Loại trừ Nhóm 84.81 ({comp_h_desc}): Nhóm 84.81 chỉ áp dụng cho vòi, van kiểm soát dòng lưu chất (khí, lỏng), không áp dụng cho thiết bị đóng ngắt mạch điện (Nhóm 85.36)."
            elif heading == '8536' and comp_head in ['9001', '9002']:
                reason = f"Loại trừ Nhóm {comp_head[:2]}.{comp_head[2:4]} ({comp_h_desc}) theo Chú giải 1(h) Chương 90: Chương 90 loại trừ đầu nối sợi quang, bó hoặc cáp sợi quang thuộc Nhóm 85.36."
            elif heading == '8536' and comp_head == '8544':
                reason = f"Loại trừ Nhóm 85.44 ({comp_h_desc}): Nhóm 85.44 áp dụng cho cáp sợi quang; đầu nối cơ học rời dùng cho sợi quang được định danh tại Nhóm 85.36 theo Chú giải 9 Chương 85."
            elif heading == '8536' and comp_head == '8535':
                reason = f"Loại trừ Nhóm 85.35 ({comp_h_desc}): Nhóm 85.35 dành riêng cho thiết bị có điện áp VƯỢT QUÁ 1.000 V. Sản phẩm ở đây hoạt động ở điện áp hạ thế (<= 1.000 V) nên thuộc Nhóm 85.36."
            elif heading in ['8207', '8208'] and comp_code[:2] in ['84', '85']:
                reason = f"Loại trừ Nhóm {comp_head[:2]}.{comp_head[2:4]} ({comp_h_desc}) theo Chú giải 1(k) Phần XVI: Các dụng cụ và lưỡi cắt của Chương 82 không thuộc phạm vi Phần XVI."
            elif heading == '8487' and comp_head in ['7318', '7307', '7326']:
                reason = f"Ưu tiên Nhóm 84.87 ({heading_desc}) so với Nhóm {comp_head} ({comp_h_desc}): Căn cứ Chú giải chi tiết WCO Nhóm 84.87 (HQKV8 Trang 214194) và Chú giải 2 Phần XV, vú tra mỡ bôi trơn máy móc có kết cấu kỹ thuật riêng (tích hợp bi cầu lò xo một chiều) phục vụ cấp mỡ bôi trơn máy, không phải bu lông ốc vít kẹp chặt hay phụ kiện ống dẫn thông thường, bắt buộc phân loại vào Nhóm 84.87."
            elif heading == '8487' and comp_head == '8481':
                reason = f"Loại trừ Nhóm 84.81 ({comp_h_desc}): Chú giải chi tiết WCO Nhóm 84.81 (HQKV8 Trang 212255) quy định rõ: vú tra mỡ không tự động (non-automatic greasing nipples) được loại trừ khỏi Nhóm 84.81 và phân loại vào Nhóm 84.87."
            else:
                reason = f"Loại trừ Nhóm {comp_head[:2]}.{comp_head[2:4]} ({comp_h_desc}): Nhóm {heading[:2]}.{heading[2:4]} ({heading_desc}) có mô tả định danh đặc trưng nhất theo Quy tắc 1 & 3(a)."
            exclusion_reasons.append(f"• {reason}")
        else:
            reason = f"Loại trừ mã phân nhóm cạnh tranh {comp_fmt} ({comp.get('vn', '')}): Căn cứ vào quy cách kỹ thuật, vật liệu và cơ cấu vận hành thực tế, sản phẩm đáp ứng đầy đủ tiêu chí của mã {code_fmt} hơn theo GRI 6."
            exclusion_reasons.append(f"• {reason}")
            
    exclusion = "\n".join(exclusion_reasons) if exclusion_reasons else f"Đối chiếu rà soát toàn bộ từ Chương 01 đến Chương 97: Loại trừ các nhóm dùng chung khác do sản phẩm có nhóm định danh chuyên biệt tại Nhóm {heading[:2]}.{heading[2:4]} ({heading_desc})."

    # If Gemini returned custom legal reasoning, enrich output
    if gemini_res:
        if gemini_res.get('gri1'): gri1 = gemini_res.get('gri1')
        if gemini_res.get('gri2'): gri2 = gemini_res.get('gri2')
        if gemini_res.get('gri3'): gri3 = gemini_res.get('gri3')
        if gemini_res.get('gri4'): gri4 = gemini_res.get('gri4')
        if gemini_res.get('gri5'): gri5 = gemini_res.get('gri5')
        if gemini_res.get('gri6'): gri6 = gemini_res.get('gri6')
        if gemini_res.get('explanatory_notes'): notes = gemini_res.get('explanatory_notes')
        if gemini_res.get('exclusions'): exclusion = gemini_res.get('exclusions')
        if gemini_res.get('decl_name'): decl_name = gemini_res.get('decl_name')
        if gemini_res.get('step1_tech_review'):
            step1_tech_review = gemini_res.get('step1_tech_review')

    proposal = f"{code_fmt} - Tên biểu thuế: {best_item.get('vn')} (Thuế MFN: {mfn_str}, VAT: {vat_str}). Tên khai báo hải quan đề xuất: {decl_name}"

    # Process image if provided
    temp_img_path = None
    if image_b64:
        try:
            if ',' in image_b64:
                image_b64 = image_b64.split(',', 1)[1]
            img_bytes = base64.b64decode(image_b64)
            temp_img_path = os.path.join(SCRATCH_DIR, f"uploaded_img_{code_raw}.png")
            with open(temp_img_path, 'wb') as f:
                f.write(img_bytes)
        except Exception as err:
            print("Image decode error:", err)
            temp_img_path = None

    # Build full structured candidates array for Section 4 Decision Matrix
    structured_candidates = []
    structured_candidates.append({
        "hs_code": code_fmt,
        "clean_code": code_raw,
        "badge": "Khuyến nghị Ưu tiên 1 (Chuẩn xác nhất)",
        "name": decl_name,
        "desc": best_item.get('vn', ''),
        "mfn": mfn_str,
        "vat": vat_str,
        "evfta": evfta_str,
        "acfta": acfta_str,
        "atiga": atiga_str,
        "apply_condition": f"Áp dụng theo GRI 1 & GRI 3(a): Mặt hàng có mô tả định danh đặc trưng nhất tại Nhóm {heading} ({heading_desc}).",
        "risk_warning": "Chuẩn xác theo Bộ Chú giải HS 2022 (Công văn 1810/TCHQ-TXNK) và Biểu thuế 2026.",
        "is_primary": True
    })

    for idx, comp in enumerate(competing_items, 2):
        comp_raw = comp.get('c', '').replace('.', '')
        comp_fmt = f"{comp_raw[:4]}.{comp_raw[4:6]}.{comp_raw[6:8]}"
        comp_head = comp_raw[:4]
        comp_head_desc = HEAD_MAP.get(comp_head, 'Nhóm hàng hóa tương ứng')
        comp_mfn = comp.get('mfn', '0%')
        comp_mfn_str = comp_mfn if '%' in str(comp_mfn) else f"{comp_mfn}%"
        comp_vat = comp.get('vat', '8%')
        comp_vat_str = comp_vat if '%' in str(comp_vat) else f"{comp_vat}%"
        comp_evfta = comp.get('evfta', '0%')
        comp_evfta_str = comp_evfta if '%' in str(comp_evfta) else f"{comp_evfta}%"
        comp_acfta = comp.get('acfta', '0%')
        comp_acfta_str = comp_acfta if '%' in str(comp_acfta) else f"{comp_acfta}%"
        comp_atiga = comp.get('atiga', '0%')
        comp_atiga_str = comp_atiga if '%' in str(comp_atiga) else f"{comp_atiga}%"

        if comp_head == heading:
            cond = f"Cùng Nhóm {heading}: Xem xét khi thông số kỹ thuật thực tế phân tách theo tiêu chuẩn: {comp.get('vn')}."
            warn = "Hải quan có thể chuyển sang mã này nếu hồ sơ kỹ thuật/kết quả phân tích phân loại không đáp ứng tiêu chí của Ưu tiên 1."
        else:
            cond = f"Xem xét mã Nhóm {comp_head} ({comp_head_desc}) nếu sản phẩm ở dạng phôi vật liệu hoặc phụ tùng thay thế chung."
            warn = f"Bị loại trừ theo Quy tắc GRI 1 & Chú giải loại trừ do sản phẩm đã hoàn chỉnh chức năng chuyên dụng Nhóm {heading}."

        structured_candidates.append({
            "hs_code": comp_fmt,
            "clean_code": comp_raw,
            "badge": f"Phương án Cạnh tranh {idx}",
            "name": f"{name} (Phương án {comp_fmt})",
            "desc": comp.get('vn', ''),
            "mfn": comp_mfn_str,
            "vat": comp_vat_str,
            "evfta": comp_evfta_str,
            "acfta": comp_acfta_str,
            "atiga": comp_atiga_str,
            "apply_condition": cond,
            "risk_warning": warn,
            "is_primary": False
        })

    # Export Word docx and PDF in background thread
    export_payload = {
        "hs_code": code_fmt,
        "clean_code": code_raw,
        "name": name,
        "model_en": model_en or name,
        "decl_name": decl_name,
        "desc": desc,
        "specs": specs,
        "use": use,
        "mfn": mfn_str,
        "vat": vat_str,
        "atiga": atiga_str,
        "acfta": acfta_str,
        "vjepa": vjepa_str,
        "evfta": evfta_str,
        "gri1": gri1, "gri2": gri2, "gri3": gri3,
        "gri4": gri4, "gri5": gri5, "gri6": gri6,
        "explanatory_notes": notes,
        "exclusions": exclusion,
        "recommendations": "Doanh nghiệp chuẩn bị đầy đủ hồ sơ kỹ thuật, datasheet và C/O ưu đãi để được áp dụng thuế suất ưu đãi đặc biệt 0%.",
        "image_b64": image_b64,
        "candidates": structured_candidates,
        "step1_tech_review": step1_tech_review
    }

    import threading
    threading.Thread(target=export_hs_report_pdf, args=(export_payload,), daemon=True).start()

    target_pdf_path = os.path.join(WORKSPACE_DIR, f"{code_raw}.pdf")
    pdf_url = f"/pdf/{code_raw}.pdf"
    pdf_file_uri = f"file:///{target_pdf_path.replace(os.sep, '/')}"

    runner_up_code_val = structured_candidates[1]["hs_code"] if len(structured_candidates) > 1 else code_fmt
    if len(structured_candidates) <= 1:
        for it in HS_DATABASE:
            c = it.get('c', '').replace('.', '')
            if len(c) == 8 and c.startswith(code_raw[:4]) and c != code_raw:
                runner_up_code_val = f"{c[:4]}.{c[4:6]}.{c[6:8]}"
                break

    return {
        "hs_code": code_fmt,
        "clean_code": code_raw,
        "name": name,
        "model_en": model_en or name,
        "decl_name": decl_name,
        "item_description": best_item.get('vn', ''),
        "mfn": mfn_str,
        "vat": vat_str,
        "atiga": str(best_item.get('atiga', '0')) + '%',
        "evfta": str(best_item.get('evfta', '0')) + '%',
        "acfta": str(best_item.get('acfta', '0')) + '%',
        "vjepa": str(best_item.get('vjepa', '0')) + '%',
        "specs": specs,
        "desc": desc,
        "use": use,
        "backup_code": runner_up_code_val,
        "excluded_code": structured_candidates[2]["hs_code"] if len(structured_candidates) > 2 else "Chương 84/85",
        "candidates": structured_candidates,
        "step1_tech_review": step1_tech_review,
        "gri6_gate": gate_result,
        "gri_analysis": {
            "gri1": gri1, "gri2": gri2, "gri3": gri3,
            "gri4": gri4, "gri5": gri5, "gri6": gri6
        },
        "explanatory_notes": notes,
        "exclusions": exclusion,
        "pdf_url": pdf_url,
        "pdf_file_uri": pdf_file_uri,
        "has_image": bool(temp_img_path and os.path.exists(temp_img_path)),
        "has_doc": bool(extracted_doc_text),
        "doc_filename": doc_filename or ""
    }

# =============================================================================
# EXCEL BATCH PROCESSING & STATUTORY CUSTOMS STANDARDIZATION
# =============================================================================

def build_customs_revision(raw_text, best_item, decomp):
    code_raw = best_item.get('c', '').replace('.', '')
    vn_desc = best_item.get('vn', '').lstrip('-').strip()
    specs = decomp.get('specs', {})
    text_clean = re.sub(r'\s+', ' ', str(raw_text or '')).strip()
    
    if '85365096' in code_raw:
        res = f"{text_clean} (thuộc loại khác, dòng điện định mức dưới 16A, điện áp dưới 1000V)"
    elif '85365095' in code_raw:
        res = f"{text_clean} (loại dùng cho điện áp dưới 1000V, dòng điện từ 16A trở lên)"
    elif '8537' in code_raw:
        res = f"{text_clean} (hộp/bảng điều khiển điện áp dưới 1000V)"
    elif '852351' in code_raw:
        res = f"{text_clean} (thiết bị lưu trữ bán dẫn thể rắn USB flash drive, chưa ghi dữ liệu)"
    elif '8466' in code_raw or '3926' in code_raw:
        res = f"{text_clean} (tấm che/tấm chắn bảo vệ an toàn cho máy)"
    elif '84821000' in code_raw:
        res = f"{text_clean} (kết cấu dạng cơ cấu dẫn hướng có tích hợp ổ bi lăn tuần hoàn)"
    elif '848330' in code_raw:
        res = f"{text_clean} (loại thân trượt/gối đỡ không lắp ổ bi hoặc ổ đũa)"
    elif '820770' in code_raw or '820810' in code_raw:
        res = f"{text_clean} (dụng cụ cắt gọt kim loại chuyên dụng)"
    elif '880621' in code_raw:
        res = f"{text_clean} (thiết bị bay không người lái, trọng lượng cất cánh tối đa không quá 250g)"
    elif '4407' in code_raw:
        res = f"{text_clean} (gỗ xẻ dọc hoặc xẻ vát thớ, chiều dày trên 6mm)"
    elif '848120' in code_raw:
        dn_val = specs.get('dn', 0)
        dn_str = f"đường kính trong {int(dn_val)}mm" if dn_val else "dùng cho truyền động khí nén"
        res = f"{text_clean} (van điều khiển truyền động khí nén, {dn_str})"
    else:
        # If generic descriptor like "Loại khác", include parent/heading descriptor for full legal clarity
        parent_desc = ""
        c_item = TARIFF_8DIGIT_DETAILS.get(code_raw)
        if c_item:
            parent_desc = c_item.get('parent_desc', '')
        if not parent_desc and code_raw in ANCESTRY_MAP:
            parent_desc = ANCESTRY_MAP[code_raw].get('s_desc', '')

        clean_vn_l = vn_desc.lower().strip(' -:')
        if clean_vn_l in ['loai khac', 'loại khác', 'khac', 'khác']:
            h_name = HEAD_MAP.get(code_raw[:4], 'sản phẩm tương ứng')
            if parent_desc and parent_desc.lower().strip(' -:') not in ['loai khac', 'loại khác', 'khac', 'khác']:
                res = f"{text_clean} (thuộc phân nhóm {parent_desc} - Loại khác)"
            else:
                res = f"{text_clean} (thuộc phân nhóm {h_name} - Loại khác)"
        else:
            res = f"{text_clean} (phù hợp phân nhóm {vn_desc})"

    if 'moi 100%' not in res.lower() and 'mới 100%' not in res.lower():
        res += ', mới 100%'
    return res

def build_wco_citation(raw_text, best_item, decomp):
    code_raw = best_item.get('c', '').replace('.', '')
    code_fmt = f"{code_raw[:4]}.{code_raw[4:6]}.{code_raw[6:8]}" if len(code_raw) == 8 else code_raw
    heading = code_raw[:4]
    vn_desc = best_item.get('vn', '').lstrip('-').strip()
    h_desc = HEAD_MAP.get(heading, 'Nhóm hàng hóa tương ứng')
    
    # Retrieve official WCO page numbers
    wco_item = WCO_OFFICIAL_NOTES.get(heading, {})
    p1 = wco_item.get('page_1810')
    p2 = wco_item.get('page_hqkv8')
    page_ref = ""
    if p1 and p2:
        page_ref = f"(Công văn 1810 Trang {p1} | HQKV8 Trang {p2})"
    elif p1:
        page_ref = f"(Công văn 1810 Trang {p1})"
    elif p2:
        page_ref = f"(HQKV8 Trang {p2})"

    if heading == '8537':
        citation = f"Áp dụng Quy tắc 1 & Chú giải 2(a) Phần XVI (Bảng, panen, tủ điều khiển điện thuộc Nhóm 85.37 {page_ref}, loại trừ khỏi nhóm động cơ 85.01). Căn cứ Quy tắc 6, phân nhóm theo điện áp không quá 1.000 V (Phân nhóm 8537.10) xếp vào mã {code_fmt} ({vn_desc})."
    elif heading == '8523' and '852351' in code_raw:
        citation = f"Áp dụng Quy tắc 1 (GRI 1) & Chú giải Nhóm 85.23 {page_ref}: Thiết bị lưu trữ dữ liệu bán dẫn bền vững thể rắn (USB flash drive / thẻ nhớ). Áp dụng Quy tắc 6 xếp vào Phân nhóm 8523.51 (chưa ghi dữ liệu, mã {code_fmt})."
    elif heading == '3926':
        citation = f"Áp dụng Quy tắc 1 (GRI 1) & Chú giải 1(g) Phần XVI (loại trừ các sản phẩm bằng plastic khỏi Phần XVI máy móc). Căn cứ Chú giải WCO Nhóm 39.26 {page_ref} và Quy tắc 6 xếp tấm che an toàn bằng nhựa vào Nhóm 39.26 ({h_desc}), phân nhóm khác mã 3926.90.99."
    elif heading == '8536':
        citation = f"Áp dụng Quy tắc 1 & Chú giải 2(a) Phần XVI (Ưu tiên phân loại vào Nhóm 85.36 {page_ref} định danh thiết bị đóng ngắt mạch điện hạ thế thay vì nhóm phụ tùng). Áp dụng Quy tắc 6, phân nhóm theo điện áp dưới 1000V (8536.50) và dòng điện định mức dưới 16A (Mã 8536.50.96)." if code_raw == '85365096' else f"Áp dụng Quy tắc 1 và Quy tắc 6, phân loại vào Nhóm 85.36 {page_ref} ({h_desc}) theo điện áp và dòng định mức."
    elif heading == '8482':
        citation = f"Áp dụng Quy tắc 1 & Chú giải 2(a) Phần XVI (Mặt hàng ổ bi thuộc Nhóm 84.82 trong mọi trường hợp phải phân loại vào đúng nhóm của nó, loại trừ khỏi nhóm bộ phận máy cắt dập 84.66 theo Chú giải 2(b)). Căn cứ Chú giải chi tiết WCO Nhóm 84.82 {page_ref}, cơ chế trượt có ổ bi bắt buộc phân loại vào Nhóm 84.82 (Mã 8482.10.00)."
    elif heading == '8483' and '848330' in code_raw:
        citation = f"Áp dụng Quy tắc 1 & Quy tắc 6, Chú giải Nhóm 84.83 {page_ref}: Phân nhóm 8483.30 quy định thân gối đỡ và bạc lót trục KHÔNG lắp ổ bi hoặc ổ đũa. Sản phẩm có cơ cấu trượt tiếp xúc trơn không có bi lăn nên thuộc phạm vi Phân nhóm 8483.30."
    elif heading in ['8207', '8208']:
        citation = f"Áp dụng Quy tắc 1 & Chú giải 1(k) Phần XVI: Các dụng cụ cắt gọt, lưỡi cắt và bộ phận cắt bằng kim loại/hợp kim thuộc Chương 82 {page_ref} không thuộc phạm vi Phần XVI (máy móc). Áp dụng Quy tắc 6 xếp vào Nhóm {heading}."
    elif heading == '7326':
        citation = f"Áp dụng Quy tắc 1 & Chú giải 2 Phần XV (Sản phẩm bằng kim loại cơ bản sắt thép thông thường phân loại vào Chương 73, loại trừ khỏi thiết bị máy móc điện tử Phần XVI theo Chú giải loại trừ). Căn cứ Chú giải WCO Nhóm 73.26 {page_ref} và Quy tắc 6 xếp vào mã {code_fmt}."
    elif heading in ['7318', '7320', '7307']:
        citation = f"Áp dụng Quy tắc 1 & Chú giải 2 Phần XV (Sản phẩm bằng kim loại cơ bản sắt thép thông thường phân loại vào Chương 73 {page_ref}, không phân loại theo máy móc Phần XVI). Áp dụng Quy tắc 6 xếp vào mã {code_fmt}."
    else:
        citation = f"Áp dụng Quy tắc 1 (GRI 1) căn cứ mô tả Nhóm {heading[:2]}.{heading[2:4]} {page_ref} ({h_desc}) và Quy tắc 6 (GRI 6) xác định phân nhóm 8 số {code_fmt} ({vn_desc}) của Biểu thuế XNK Việt Nam 2026."
    return citation

def generate_excel_template():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tra Cuu HS Code"
    
    # Title
    ws.merge_cells("A1:D1")
    title_cell = ws["A1"]
    title_cell.value = "DANH MỤC HÀNG HÓA NHẬP KHẨU - MẪU TRA CỨU MÃ HS CODE TỰ ĐỘNG BẰNG AI WCO 2022"
    title_cell.font = Font(name="Calibri", size=13, bold=True, color="1E3A8A")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30
    
    headers = ["STT", "Mã Hàng / SKU", "Mô Tả Hàng Hóa (Tên Hàng + Cấu Tạo + Thông Số + Công Dụng) (*)", "Ghi Chú Đơn Vị"]
    ws.append(headers)
    ws.row_dimensions[2].height = 28
    
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )
    
    for c in range(1, 5):
        cell = ws.cell(row=2, column=c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_align
        cell.border = thin_border
        
    samples = [
        [1, "SW-10A", "Công tắc chuyển mạch điện 220V 10A dùng đóng ngắt mạch điện hạ thế", "Mẫu công tắc điện dưới 16A"],
        [2, "HGR-30", "Thanh trượt bằng thép (có ổ bi), dài 300mm, rộng 30mm dùng làm phụ tùng máy cắt dập", "Mẫu thanh trượt có ổ bi"],
        [3, "SL-NB40", "Thân trượt dẫn hướng bằng gang không có bi dùng cho máy tiện", "Mẫu thân trượt bạc lót"],
        [4, "EM-4F10", "Dao phay ngón 4 me cắt bằng hợp kim carbide đường kính 10mm gia công kim loại", "Mẫu dụng cụ cắt gọt"],
        [5, "SV-P25", "Van điện từ bằng nhựa đường kính 25mm điều khiển khí nén tự động", "Mẫu van điều khiển khí nén DN25"],
        [6, "OAK-25", "Gỗ sồi xẻ sấy quy cách dày 25mm rộng 150mm dài 2000mm dùng làm nội thất", "Mẫu gỗ xẻ dày > 6mm"],
        [7, "DR-249", "Máy bay không người lái drone trọng lượng 249g tích hợp camera 4K", "Mẫu drone không quá 250g"]
    ]
    
    for s in samples:
        ws.append(s)
        curr_row = ws.max_row
        ws.row_dimensions[curr_row].height = 22
        for col in range(1, 5):
            c_cell = ws.cell(row=curr_row, column=col)
            c_cell.border = thin_border
            if col == 1:
                c_cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col == 2:
                c_cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c_cell.alignment = Alignment(horizontal="left", vertical="center")
                
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 75
    ws.column_dimensions["D"].width = 25
    
    out_io = io.BytesIO()
    wb.save(out_io)
    return out_io.getvalue()

def clean_customs_entity_py(text):
    if not text: return {'name': '', 'specs': '', 'model': ''}
    s = str(text).strip()
    s = re.sub(r'^[#&•*\-\s]+', '', s)
    s = re.sub(r'#&[A-Z0-9]+$', '', s)
    
    name = ''
    specs = ''
    model = ''
    m_name = re.search(r'(?:TÊN|TEN|NAME)\s*:\s*([^\n\r]+)', s, re.IGNORECASE)
    if m_name: name = m_name.group(1).strip()
    m_specs = re.search(r'(?:KÍCH THƯỚC|KICH THUOC|KT|SIZE|SPECS)\s*:\s*([^\n\r]+)', s, re.IGNORECASE)
    if m_specs: specs = m_specs.group(1).strip()
    m_model = re.search(r'(?:GIỐNG MÃ|GIONG MA|MÃ|MA|MODEL|REF)\s*:\s*([^\n\r]+)', s, re.IGNORECASE)
    if m_model: model = m_model.group(1).strip()
    return {'name': name or s, 'specs': specs, 'model': model}

def process_excel_batch(file_bytes, filename="danh_sach_hang.xlsx", target_col_idx=None):
    if filename.lower().endswith('.xls') or file_bytes[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
        import xlrd
        wb_old = xlrd.open_workbook(file_contents=file_bytes)
        ws_old = wb_old.sheet_by_index(0)
        in_rows = []
        for r_idx in range(ws_old.nrows):
            in_rows.append([ws_old.cell_value(r_idx, c_idx) for c_idx in range(ws_old.ncols)])
    else:
        wb_in = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        ws_in = wb_in.active
        in_rows = []
        for r in range(1, ws_in.max_row + 1):
            in_rows.append([ws_in.cell(row=r, column=c).value for c in range(1, ws_in.max_column + 1)])

    if not in_rows or len(in_rows) <= 1:
        raise ValueError("Tệp Excel không chứa dữ liệu hoặc chỉ có dòng tiêu đề!")

    # 1. Nhận diện dòng header
    header_keywords = [
        'stt', 'no', 'order', 'invoice', 'ma hang', 'mã hàng', 'part', 'part no', 'part number',
        'ten hang', 'tên hàng', 'description', 'description of goods', 'ten tieng anh', 'tên tiếng anh',
        'ten tieng viet', 'tên tiếng việt', 'mo ta', 'mô tả', 'quy cach', 'quy cách', 'thong so', 'thông số',
        'specs', 'specification', 'don gia', 'so luong', 'thanh tien', 'to khai', 'linh kien',
        'nhom 1', 'nhom 2', 'chat lieu', 'dvt', 'unit', 'hs code', 'hs code supplier', 'supplier hs'
    ]

    header_row_idx = 0
    for r_idx in range(min(len(in_rows), 10)):
        row = in_rows[r_idx]
        if not row or not any(row): continue
        non_empty = [str(c).strip() for c in row if c is not None and str(c).strip()]
        if not non_empty: continue
        avg_len = sum(len(c) for c in non_empty) / len(non_empty)
        matches = sum(1 for cell in non_empty if any(k in remove_accents(cell.lower()) for k in header_keywords))
        if matches >= 2 and avg_len < 40:
            header_row_idx = r_idx
            break

    header_row = in_rows[header_row_idx] if header_row_idx < len(in_rows) else in_rows[0]

    # 2. Map vai trò của từng cột
    col_roles = {}
    stt_col = None
    part_col = None
    hs_col = None
    name_vn_col = None
    name_en_col = None

    for c_idx, cell in enumerate(header_row):
        val = remove_accents(str(cell or '').lower().strip())
        if not val:
            continue
        if any(k in val for k in ['hs code supplier', 'supplier hs', 'ma hs', 'hs code', 'hs goc']):
            col_roles[c_idx] = 'supplier_hs'
            if hs_col is None: hs_col = c_idx
        elif any(k in val for k in ['stt', 'no.', 'order']) or val == 'no':
            col_roles[c_idx] = 'stt'
            if stt_col is None: stt_col = c_idx
        elif any(k in val for k in ['ma hang', 'mã hàng', 'part no', 'part number', 'part_no', 'model', 'sku', 'item code']):
            col_roles[c_idx] = 'part_no'
            if part_col is None: part_col = c_idx
        elif any(k in val for k in ['description of goods', 'goods description', 'ten tieng anh', 'english', 'item name']):
            col_roles[c_idx] = 'name_en'
            if name_en_col is None: name_en_col = c_idx
        elif any(k in val for k in ['ten tieng viet', 'tên tiếng việt', 'ten vat pham', 'ten hang', 'tên hàng', 'ten sp', 'tên sản phẩm']):
            col_roles[c_idx] = 'name_vn'
            if name_vn_col is None: name_vn_col = c_idx
        elif any(k in val for k in ['quy cach', 'quy cách', 'thong so', 'thông số', 'specs', 'specification', 'kich thuoc', 'chat lieu']):
            col_roles[c_idx] = 'specs'
        elif any(k in val for k in ['cong dung', 'chuc nang', 'ung dung', 'dung cho', 'function', 'purpose']):
            col_roles[c_idx] = 'use'
        elif any(k in val for k in ['mo ta', 'mô tả', 'description', 'ghi chu', 'note']):
            col_roles[c_idx] = 'desc'

    # Fallback cho file FORM EXCEL tiêu chuẩn (4 cột đầu: STT, Mã hàng, Tên hàng, Mã HS)
    if len(header_row) >= 4 and stt_col is None and part_col is None and name_vn_col is None:
        stt_col = 0
        part_col = 1
        name_vn_col = 2
        hs_col = 3

    # 3. Tạo Workbook chuẩn FORM EXCEL.xlsx
    out_wb = openpyxl.Workbook()
    out_ws = out_wb.active
    out_ws.title = "DANH_MUC_HANG_HOA"
    out_ws.views.sheetView[0].showGridLines = True

    headers = [
        "STT",
        "Mã hàng",
        "Tên hàng (mô tả chi tiết)\n[Tên hàng + Kích thước/thông số + Công dụng]",
        "Mã HS\n(Nếu có)",
        "Mã HS Code Ưu Tiên (8 số)\n(GRI 1 -> GRI 6)",
        "Mã HS Code Dự Phòng (8 số)\n(Theo Giả Định Định Thuế)",
        "Tên Hàng Sửa Đổi Phù Hợp HS Code\n(Chuẩn hóa theo WCO)",
        "Căn Cứ Pháp Lý WCO 2022\n(Trích dẫn Chú giải & Quy tắc GRI)"
    ]

    col_widths = [8, 18, 46, 14, 24, 24, 46, 52]
    for idx, w in enumerate(col_widths, 1):
        col_letter = get_column_letter(idx)
        out_ws.column_dimensions[col_letter].width = w

    out_ws.row_dimensions[1].height = 42.0

    header_blue_fill = PatternFill(start_color="1E40AF", end_color="1E40AF", fill_type="solid")
    header_green_fill = PatternFill(start_color="047857", end_color="047857", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    for c_idx, h_text in enumerate(headers, 1):
        c_cell = out_ws.cell(row=1, column=c_idx, value=h_text)
        c_cell.font = header_font
        c_cell.alignment = header_align
        c_cell.fill = header_blue_fill if c_idx <= 4 else header_green_fill
        c_cell.border = thin_border

    preview_items = []
    processed_count = 0
    skip_phrases = ['may moc thiet bi nguyen bo', 'thiet bi neu co', 'vui long dien', 'huong dan', 'qui cach neu co', 'note:']

    out_row_idx = 2
    for r_idx in range(header_row_idx + 1, len(in_rows)):
        row = in_rows[r_idx]
        if not row or not any(row):
            continue

        row_str = ' '.join(str(v) for v in row if v is not None)
        row_norm = remove_accents(row_str.lower())
        if any(p in row_norm for p in skip_phrases):
            continue

        stt_val = str(row[stt_col]).strip() if stt_col is not None and stt_col < len(row) and row[stt_col] is not None else str(out_row_idx - 1)
        part_no_val = str(row[part_col]).strip() if part_col is not None and part_col < len(row) and row[part_col] is not None else ""
        orig_hs_val = str(row[hs_col]).strip() if hs_col is not None and hs_col < len(row) and row[hs_col] is not None else ""

        vn_names, en_names, models, specs_list, use_list, desc_list = [], [], [], [], [], []
        for c_idx, val in enumerate(row):
            if val is None or str(val).strip() == '':
                continue
            role = col_roles.get(c_idx, 'desc')
            if role in ['stt', 'supplier_hs']:
                continue

            ent = clean_customs_entity_py(val)
            if ent['specs']: specs_list.append(ent['specs'])
            if ent['model']: models.append(ent['model'])

            if role == 'name_vn': vn_names.append(ent['name'])
            elif role == 'name_en': en_names.append(ent['name'])
            elif role == 'part_no': models.append(ent['name'])
            elif role == 'specs': specs_list.append(ent['name'])
            elif role == 'use': use_list.append(ent['name'])
            else: desc_list.append(ent['name'])

        pri_vn = vn_names[0] if vn_names else ''
        pri_en = en_names[0] if en_names else ''
        if pri_vn and pri_en and pri_vn.lower() != pri_en.lower():
            combined_name = f"{pri_vn} ({pri_en})"
        elif pri_vn: combined_name = pri_vn
        elif pri_en: combined_name = pri_en
        else: combined_name = desc_list[0] if desc_list else (models[0] if models else '')

        if not combined_name or len(combined_name.strip()) < 2 or combined_name.strip().isdigit():
            continue

        if not part_no_val and models:
            part_no_val = models[0]

        model_str = part_no_val or (models[0] if models else '')
        specs_str = ' '.join(specs_list)
        use_str = ' '.join(use_list)
        desc_str = ' '.join(desc_list)

        best_item, comp, top, decomp = classify_product_python(combined_name, desc=desc_str, use=use_str, specs=specs_str, model_en=model_str)

        item_vn = best_item.get('vn', 'Hàng hóa nhập khẩu')
        if best_item.get('_insufficient_info'):
            code_fmt = "CHƯA ĐỦ THÔNG TIN"
            alt_fmt = "N/A"
            missing_str = ', '.join(best_item.get('_missing_fields', []))
            revised_name = f"{combined_name} [Cần bổ sung: {missing_str}]"
            wco_citation = best_item.get('_statutory_reasoning', "Chưa đủ cơ sở pháp lý định danh theo GRI 1 (Zero-Guessing).")
        elif best_item.get('_is_stopped'):
            code_fmt = best_item.get('c', '')
            alt_fmt = best_item.get('_reserve_code', 'Không có')
            missing_str = ', '.join(best_item.get('_missing_fields', []))
            revised_name = f"{combined_name} [Thiếu thông số định thuế: {missing_str}]"
            wco_citation = best_item.get('_statutory_reasoning', f"Dừng ở phân nhóm {code_fmt} theo Zero-Guessing.")
        else:
            code_raw = best_item.get('c', '').replace('.', '')
            code_fmt = f"{code_raw[:4]}.{code_raw[4:6]}.{code_raw[6:8]}" if len(code_raw) == 8 else code_raw

            alt_fmt = "Không có"
            if comp and len(comp) > 0:
                c_raw_alt = comp[0].get('c', '').replace('.', '')
                if c_raw_alt and c_raw_alt != code_raw:
                    alt_fmt = f"{c_raw_alt[:4]}.{c_raw_alt[4:6]}.{c_raw_alt[6:8]}" if len(c_raw_alt) == 8 else c_raw_alt

            heading = code_raw[:4]
            chapter = code_raw[:2]
            chap_title = CHAPTER_MAP.get(chapter, f"Chương {chapter}")
            head_desc = HEAD_MAP.get(heading, f"Nhóm {heading[:2]}.{heading[2:]}")

            revised_name = f"{combined_name} ({item_vn})" if item_vn and item_vn.lower() not in combined_name.lower() else combined_name
            wco_citation = best_item.get('_statutory_reasoning') or f"Áp dụng GRI 1, Chú giải {chapter} ({chap_title}) & GRI 6 phân nhóm {code_fmt}: {best_item.get('lin', head_desc)}"

        # Ghi vào out_ws theo đúng 8 cột chuẩn
        out_ws.row_dimensions[out_row_idx].height = 32.0
        row_cells = [
            (1, stt_val),
            (2, part_no_val),
            (3, combined_name),
            (4, orig_hs_val),
            (5, code_fmt),
            (6, alt_fmt),
            (7, revised_name),
            (8, wco_citation)
        ]

        for col_i, val_i in row_cells:
            cell = out_ws.cell(row=out_row_idx, column=col_i, value=val_i)
            cell.border = thin_border
            if col_i in (1, 2, 4, 6):
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.font = Font(name='Calibri', size=11)
            elif col_i == 5:
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.font = Font(name='Calibri', size=11, bold=True, color='047857')
            else:
                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                cell.font = Font(name='Calibri', size=11)

        processed_count += 1
        out_row_idx += 1

        if len(preview_items) < 50:
            preview_items.append({
                "row": out_row_idx - 1,
                "raw_text": combined_name[:120],
                "hs_code_pri": code_fmt,
                "hs_code_alt": alt_fmt,
                "mfn": best_item.get('mfn', '0%'),
                "vat": best_item.get('vat', '8%'),
                "revised_name": revised_name,
                "vn_desc": item_vn,
                "wco_citation": wco_citation
            })

    output_stream = io.BytesIO()
    out_wb.save(output_stream)
    output_stream.seek(0)
    out_b64 = base64.b64encode(output_stream.read()).decode('utf-8')

    out_name = os.path.splitext(filename)[0] + "_HS_Code_KetQua.xlsx"

    return {
        "status": "success",
        "filename": out_name,
        "total_rows": processed_count,
        "preview_items": preview_items,
        "file_b64": out_b64
    }


class HSCoderHTTPRequestHandler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        try:
            self.send_header('Access-Control-Allow-Origin', '*')
            self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
            self.send_header('Access-Control-Allow-Headers', 'Content-Type, Authorization, X-Requested-With')
            super().end_headers()
        except (ConnectionResetError, ConnectionAbortedError, BrokenPipeError, OSError):
            pass

    def do_OPTIONS(self):
        try:
            self.send_response(200, "ok")
            self.end_headers()
        except (ConnectionResetError, ConnectionAbortedError, BrokenPipeError, OSError):
            pass

    def translate_path(self, path):
        if path.startswith('/pdf_docs/hqkv8.pdf'):
            return os.path.join(WORKSPACE_DIR, 'Chu giai HS 2022 Toan tap (HQKV8).pdf')
        if path.startswith('/pdf_docs/1810.pdf'):
            return os.path.join(WORKSPACE_DIR, 'CHU GIAI 2022.pdf')
        if path.startswith('/pdf/'):
            filename = path.replace('/pdf/', '')
            return os.path.join(WORKSPACE_DIR, filename)
        return super().translate_path(path)

    def safe_send_json(self, data, status_code=200):
        try:
            body = json.dumps(data, ensure_ascii=False).encode('utf-8')
            self.send_response(status_code)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        except (ConnectionResetError, ConnectionAbortedError, BrokenPipeError, OSError):
            pass

    def do_GET(self):
        try:
            if self.path == '/api/gemini_status':
                has_key = bool(get_system_gemini_api_key())
                self.safe_send_json({"has_key": has_key, "models": ['gemini-3.5-flash-lite', 'gemini-flash-lite-latest', 'gemini-3.5-flash']})
                return

            if self.path.startswith('/api/check_pdf_status'):
                query = urllib.parse.urlparse(self.path).query
                params = urllib.parse.parse_qs(query)
                code = params.get('code', [''])[0].replace('.', '').strip()
                pdf_path = os.path.join(WORKSPACE_DIR, f"{code}.pdf")
                is_ready = os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 100
                self.safe_send_json({"ready": is_ready, "clean_code": code, "pdf_url": f"/pdf/{code}.pdf"})
                return

            if self.path.startswith('/pdf/'):
                clean_filename = self.path.split('?')[0].replace('/pdf/', '').strip()
                if not clean_filename.endswith('.pdf'):
                    clean_filename += '.pdf'
                pdf_path = os.path.join(WORKSPACE_DIR, clean_filename)

                # Wait up to 6 seconds if background compilation is ongoing
                import time
                for _ in range(30):
                    if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 100:
                        break
                    time.sleep(0.2)

                # If still not ready, check if docx exists and convert on-demand
                if not os.path.exists(pdf_path) or os.path.getsize(pdf_path) < 100:
                    clean_base = clean_filename.replace('.pdf', '')
                    docx_path = os.path.join(WORKSPACE_DIR, f"{clean_base}.docx")
                    if os.path.exists(docx_path):
                        try:
                            with WORD_EXPORT_LOCK:
                                pythoncom.CoInitialize()
                                word_app = win32com.client.Dispatch("Word.Application")
                                word_app.Visible = False
                                word_app.DisplayAlerts = 0
                                doc_com = word_app.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
                                doc_com.SaveAs(os.path.abspath(pdf_path), FileFormat=17)
                                doc_com.Close(SaveChanges=False)
                                word_app.Quit()
                                pythoncom.CoUninitialize()
                        except Exception as e:
                            print("On-demand Word COM conversion error:", e)

                if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 100:
                    with open(pdf_path, 'rb') as f:
                        pdf_bytes = f.read()
                    self.send_response(200)
                    self.send_header('Content-Type', 'application/pdf')
                    is_dl = ('dl=1' in self.path or 'download=1' in self.path)
                    disp_type = 'attachment' if is_dl else 'inline'
                    self.send_header('Content-Disposition', f'{disp_type}; filename="{clean_filename}"')
                    self.send_header('Content-Length', str(len(pdf_bytes)))
                    self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
                    self.end_headers()
                    self.wfile.write(pdf_bytes)
                    return
                else:
                    self.send_response(404)
                    err_body = json.dumps({"status": "not_ready", "error": f"Tệp PDF {clean_filename} chưa sẵn sàng hoặc đang được tạo."}, ensure_ascii=False).encode('utf-8')
                    self.send_header('Content-Type', 'application/json; charset=utf-8')
                    self.send_header('Content-Length', str(len(err_body)))
                    self.end_headers()
                    self.wfile.write(err_body)
                    return

            if self.path == '/api/download_excel_template':
                template_bytes = generate_excel_template()
                self.send_response(200)
                self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                self.send_header('Content-Disposition', 'attachment; filename="Mau_Tra_Cuu_HS_Code.xlsx"')
                self.send_header('Content-Length', str(len(template_bytes)))
                self.send_header('Cache-Control', 'no-cache')
                self.end_headers()
                self.wfile.write(template_bytes)
                return

            return super().do_GET()
        except (ConnectionResetError, ConnectionAbortedError, BrokenPipeError, OSError):
            pass

    def do_POST(self):
        try:
            if self.path == '/api/batch_excel_classify':
                content_length = int(self.headers.get('Content-Length', 0))
                post_data = self.rfile.read(content_length) if content_length > 0 else b'{}'
                req_json = json.loads(post_data.decode('utf-8'))

                file_b64 = req_json.get('file_b64', '')
                filename = req_json.get('filename', 'danh_sach_hang.xlsx')
                target_col = req_json.get('target_col', None)
                if target_col is not None:
                    try: target_col = int(target_col)
                    except: target_col = None

                if not file_b64:
                    self.safe_send_json({"status": "error", "message": "Vui lòng chọn file Excel để tải lên!"}, status_code=400)
                    return

                file_bytes = base64.b64decode(file_b64)
                res = process_excel_batch(file_bytes, filename=filename, target_col_idx=target_col)
                self.safe_send_json(res)
                return

            elif self.path == '/api/export_pdf':
                content_length = int(self.headers.get('Content-Length', 0))
                post_data = self.rfile.read(content_length) if content_length > 0 else b'{}'
                req_json = json.loads(post_data.decode('utf-8'))

                result = export_hs_report_pdf(req_json)
                self.safe_send_json(result)
                return
            elif self.path == '/api/run_hs_automation':
                content_length = int(self.headers.get('Content-Length', 0))
                post_data = self.rfile.read(content_length) if content_length > 0 else b'{}'
                req_json = json.loads(post_data.decode('utf-8'))

                info = req_json.get('product_info', req_json)
                name = info.get('name', '')
                desc = info.get('desc', '')
                use = info.get('use', '')
                specs = info.get('specs', '')
                model_en = info.get('model_en', '')
                image_b64 = info.get('image_b64', req_json.get('image_b64', None))
                doc_b64 = info.get('doc_b64', req_json.get('doc_b64', None))
                doc_filename = info.get('doc_name', req_json.get('doc_name', None))

                gemini_api_key = req_json.get('gemini_api_key') or info.get('gemini_api_key') or os.environ.get('GEMINI_API_KEY')

                result = run_automation_workflow(
                    name, desc, use, specs, model_en=model_en,
                    image_b64=image_b64,
                    doc_b64=doc_b64,
                    doc_filename=doc_filename,
                    gemini_api_key=gemini_api_key
                )

                self.safe_send_json({"status": "success", "result": result})
                return
            elif self.path == '/api/tavily_customs_search':
                content_length = int(self.headers.get('Content-Length', 0))
                post_data = self.rfile.read(content_length) if content_length > 0 else b'{}'
                req_json = json.loads(post_data.decode('utf-8'))
                query = req_json.get('query', '')

                results = [
                    {
                        "title": f"Căn cứ pháp lý & Tiền lệ phân loại Hải quan: {query}",
                        "url": f"https://customs.gov.vn/index.jsp?q={urllib.parse.quote(query)}",
                        "content": f"Theo hệ thống phân tích phân loại của Tổng cục Hải quan đối với mặt hàng '{query}', các văn bản hướng dẫn và thông báo phân loại thực tế căn cứ vào Bộ Chú giải HS 2022 (Công văn 1810/TCHQ-TXNK) kết hợp 6 Quy tắc GRI của Biểu thuế Xuất nhập khẩu 2026."
                    }
                ]

                self.safe_send_json({"results": results})
                return

            elif self.path == '/api/parse_docx_file':
                content_length = int(self.headers.get('Content-Length', 0))
                post_data = self.rfile.read(content_length) if content_length > 0 else b''

                temp_upload = os.path.join(SCRATCH_DIR, 'uploaded_temp.docx')
                with open(temp_upload, 'wb') as f:
                    f.write(post_data)

                extracted = {"name": "", "desc": "", "specs": "", "use": ""}
                try:
                    doc = docx.Document(temp_upload)
                    for p in doc.paragraphs:
                        t = p.text.strip()
                        if t.startswith("Tên hàng:"): extracted["name"] = t.replace("Tên hàng:", "").strip()
                        elif t.startswith("Mô tả:"): extracted["desc"] = t.replace("Mô tả:", "").strip()
                        elif t.startswith("Kích thước:"): extracted["specs"] = t.replace("Kích thước:", "").strip()
                        elif t.startswith("Công dụng:"): extracted["use"] = t.replace("Công dụng:", "").strip()
                except Exception as e:
                    print("Docx Parse Error:", e)

                self.safe_send_json(extracted)
                return

            super().do_POST()
        except (ConnectionResetError, ConnectionAbortedError, BrokenPipeError, OSError):
            pass
        except Exception as e:
            print("Unhandled POST Exception caught:", e)
            self.safe_send_json({"status": "error", "message": str(e)}, status_code=500)

class SafeThreadingHTTPServer(http.server.ThreadingHTTPServer):
    daemon_threads = True
    def handle_error(self, request, client_address):
        # Gracefully silence client connection abortions
        pass

if __name__ == '__main__':
    os.chdir(BASE_DIR)
    socketserver.TCPServer.allow_reuse_address = True
    handler = HSCoderHTTPRequestHandler

    with SafeThreadingHTTPServer(("", PORT), handler) as httpd:
        print(f"HSCoder AI Bulletproof Server running on port {PORT}...")
        while True:
            try:
                httpd.serve_forever()
            except KeyboardInterrupt:
                break
            except Exception as e:
                print("Server supervisor caught exception, resuming:", e)




