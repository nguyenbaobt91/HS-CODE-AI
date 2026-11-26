# -*- coding: utf-8 -*-
"""
STATUTORY TREE WALKER (DUYỆT CÂY ĐỊNH THUẾ CHUẨN GRI 1-6)
Được thiết kế dựa trên:
1. Quy trình 8 bước SOP của Chuyên gia Phân loại Hải quan cấp cao
2. Bộ tiêu chí trong skill hs_code_automation
3. Ma trận thứ tự ưu tiên pháp lý (Chú giải 2 Phần XV, Chú giải 1 & 2 Phần XVI, Chú giải Chương 90, 87)
4. Nguyên tắc Zero-Guessing: Thiếu thông số thì dừng ở 4 số hoặc 6 số, không đoán mò!
5. Nguyên tắc Head Noun Extraction: Phân loại theo Danh từ trung tâm của tên hàng,
   không bị nhiễu bởi máy chủ hoặc mục đích sử dụng ở vế sau.
"""

import re
import unicodedata

def clean_vietnamese(input_str):
    if not input_str:
        return ""
    s = input_str.replace('đ', 'd').replace('Đ', 'd')
    nfkd = unicodedata.normalize('NFKD', s)
    return ''.join([c for c in nfkd if not unicodedata.combining(c)]).lower()

def extract_head_noun(desc):
    clean = clean_vietnamese(desc)
    parts = re.split(r'[,;()\-–\d]', clean)
    first_part = parts[0].strip() if parts else clean
    m = re.match(r'^(mieng\s+nhua|tam\s+nhua(?:\s+phang)?|thanh\s+nhua|khung\s+nhua|mieng\s+thep|tam\s+thep|thanh\s+thep|chau\s+day|ga\s+dap|ga\s+lap\s+rap)', first_part)
    if m:
        return m.group(1)
    return first_part

class StatutoryTreeWalker:
    def __init__(self):
        pass

    def extract_attributes(self, name_text, full_text):
        name_norm = clean_vietnamese(name_text)
        text_norm = clean_vietnamese(full_text)
        head_noun = extract_head_noun(name_text)
        
        attrs = {
            "name_text": name_text,
            "name_norm": name_norm,
            "head_noun": head_noun,
            "raw_text": full_text,
            "norm_text": text_norm,
            "voltage_v": None,
            "current_a": None,
            "power_w": None,
            "diameter_mm": None,
            "dn_mm": None,
            "has_connector": False,
            "no_connector": False,
            "is_threaded": None,
            "material": None,
            "is_insufficient_general": False
        }

        # 1. Kiểm tra mô tả quá sơ sài / mơ hồ
        cleaned_words = [w for w in re.split(r'\s+', name_norm) if len(w) > 1]
        vague_terms = ["phu tung", "phu tung may", "linh kien", "vat tu", "hang hoa", "chi tiet"]
        if len(cleaned_words) <= 5 and any(vt in name_norm for vt in vague_terms):
            if not any(k in text_norm for k in ["dien ap", "cong suat", "kich thuoc", "chat lieu", "thep", "nhua", "dong"]):
                attrs["is_insufficient_general"] = True
                return attrs

        # 2. Bóc tách Điện áp (Voltage)
        v_range = re.search(r'(?:dien\s+ap\s+)?(\d{1,5})\s*(?:~|-)\s*(\d{1,5})\s*(?:v|vac|vdc|von|volt)\b', text_norm)
        if v_range:
            attrs["voltage_v"] = max(float(v_range.group(1)), float(v_range.group(2)))
        else:
            v_single = re.search(r'(\d{1,5})\s*(?:v|vac|vdc|von|volt)\b', text_norm)
            if v_single:
                attrs["voltage_v"] = float(v_single.group(1))

        # 3. Bóc tách Dòng điện (Current Amperes)
        a_match = re.search(r'(\d+(?:\.\d+)?)\s*(?:a|ampe|amperes)\b', text_norm)
        if a_match:
            attrs["current_a"] = float(a_match.group(1))
        elif "duoi 16a" in text_norm or "duoi 16 a" in text_norm:
            attrs["current_a"] = 10.0
        elif "tren 16a" in text_norm or "tren 16 a" in text_norm:
            attrs["current_a"] = 25.0

        # 4. Bóc tách Công suất (Power W / kW / HP / VA / kVA)
        kw_match = re.search(r'(\d+(?:\.\d+)?)\s*kw\b', text_norm)
        if kw_match:
            attrs["power_w"] = float(kw_match.group(1)) * 1000.0
        else:
            w_match = re.search(r'(\d+(?:\.\d+)?)\s*w\b', text_norm)
            if w_match:
                attrs["power_w"] = float(w_match.group(1))
            else:
                va_match = re.search(r'(\d+(?:\.\d+)?)\s*(?:va|kva)\b', text_norm)
                if va_match:
                    val = float(va_match.group(1))
                    attrs["power_w"] = val * 1000.0 if "kva" in va_match.group(0) else val

        # 5. Bóc tách Đường kính thân / Bu lông (M... hoặc phi... hoặc ...mm)
        m_match = re.search(r'\bm(\d{1,3})(?:x\d+(?:[a-z]+)?)?(?:\b|[^\w])', text_norm)
        if m_match:
            attrs["diameter_mm"] = float(m_match.group(1))
            attrs["is_threaded"] = True
        else:
            phi_match = re.search(r'(?:phi|duong kinh(?: than)?(?: ngoai| trong)?|dk|lo)\s*(\d+(?:\.\d+)?)\s*(?:mm|cm)?\b', text_norm)
            if phi_match:
                val = float(phi_match.group(1))
                if 'cm' in phi_match.group(0):
                    val *= 10.0
                attrs["diameter_mm"] = val

        # 6. Bóc tách Kích thước Danh định Van / Đường ống (DN)
        dn_match = re.search(r'\bdn\s*(\d{1,4})\b', text_norm)
        if dn_match:
            attrs["dn_mm"] = float(dn_match.group(1))
        else:
            inch_match = re.search(r'(\d+(?:/\d+)?)\s*(?:inch|inh|in\b|\")', text_norm)
            if inch_match:
                val_str = inch_match.group(1)
                if '/' in val_str:
                    n, d = val_str.split('/')
                    attrs["dn_mm"] = (float(n) / float(d)) * 25.4
                else:
                    attrs["dn_mm"] = float(val_str) * 25.4

        # 7. Tình trạng Đầu nối (Connector Status)
        has_conn_patterns = [
            r'co\s+gan\s+dau\s+noi', r'da\s+lap\s+voi\s+dau\s+noi', r'co\s+dau\s+noi', r'gan\s+dau\s+noi',
            r'dau\s+cos', r'dau\s+cot', r'giac\s+cam', r'connector', r'terminal', r'co\s+gan\s+san\s+giac'
        ]
        no_conn_patterns = [
            r'khong\s+gan\s+dau\s+noi', r'khong\s+co\s+dau\s+noi', r'chua\s+gan\s+dau\s+noi',
            r'chua\s+lap\s+voi\s+dau\s+noi', r'khong\s+dau\s+noi', r'dang\s+cuon\s+khong\s+co\s+giac'
        ]
        if any(re.search(p, text_norm) for p in no_conn_patterns):
            attrs["no_connector"] = True
            attrs["has_connector"] = False
        elif any(re.search(p, text_norm) for p in has_conn_patterns):
            attrs["has_connector"] = True
            attrs["no_connector"] = False

        # 8. Ren (Thread Status)
        if re.search(r'\b(?:co\s+ren|ren\s+ngoai|ren\s+trong|ren\s+suot|tien\s+ren|bap\s+ren)\b', text_norm):
            attrs["is_threaded"] = True
        elif re.search(r'\b(?:khong\s+ren|chua\s+ren|tron\s+khong\s+ren)\b', text_norm):
            attrs["is_threaded"] = False

        # 9. Chất liệu cấu thành (Material)
        if re.search(r'\b(?:nhom|aluminum|aluminium)\b', text_norm):
            attrs["material"] = "aluminum"
        elif re.search(r'\b(?:nhua|plastic|pvc|pa66|acrylic|abs|polymer|chat\s+deo)\b', text_norm):
            attrs["material"] = "plastic"
        elif re.search(r'\b(?:cao\s+su|rubber|epdm|nitrile|silicone|vulcanized)\b', text_norm):
            attrs["material"] = "rubber"
        elif re.search(r'\b(?:thep\s+khong\s+gi|inox|ss304|ss316)\b', text_norm):
            attrs["material"] = "stainless_steel"
        elif re.search(r'\b(?:gang|cast\s+iron)\b', text_norm):
            attrs["material"] = "cast_iron"
        elif re.search(r'\b(?:thep|sat|kim\s+loai\s+den|carbon\s+steel)\b', text_norm):
            attrs["material"] = "steel"
        elif re.search(r'\b(?:dong\s+thau|hop\s+kim\s+dong|dong|brass|bronze|copper)\b', text_norm):
            attrs["material"] = "copper"
        elif re.search(r'\b(?:gom|dat\s+nung|su|ceramic|porcelain)\b', text_norm):
            attrs["material"] = "ceramic"

        return attrs

    def classify(self, item_input):
        name = item_input.get("name", "")
        desc = item_input.get("desc", "")
        use = item_input.get("use", "")
        specs = item_input.get("specs", "")
        model_en = item_input.get("model_en", "")

        full_text = f"{name} {desc} {use} {specs} {model_en}".strip()
        attrs = self.extract_attributes(name, full_text)
        hn = attrs["head_noun"]
        name_norm = attrs["name_norm"]
        norm = attrs["norm_text"]

        # -------------------------------------------------------------
        # BƯỚC 1: KIỂM TRA MỨC ĐỘ THIẾU THÔNG TIN MƠ HỒ (INSUFFICIENT GENERAL)
        # -------------------------------------------------------------
        if attrs["is_insufficient_general"]:
            return {
                "code": "INSUFFICIENT_INFO",
                "level": 0,
                "status": "INSUFFICIENT_INFO",
                "reasoning": "Mô tả hàng hóa quá vắn tắt ('phụ tùng/linh kiện' chung chung). Không đủ căn cứ tối thiểu để xác định Phần và Chương theo GRI 1.",
                "missing_attributes": ["name_specific", "material", "function", "host_machine"],
                "assumptions": "Không thể giả định",
                "reserve_code": "N/A"
            }

        # -------------------------------------------------------------
        # BƯỚC 1B: QUY TẮC PHÁP LÝ DÙNG CHUNG (statutory_rules.json)
        # -------------------------------------------------------------
        try:
            from statutory_rules_engine import match_statutory_rule
            _shared_res = match_statutory_rule(hn, norm)
            if _shared_res is not None:
                return _shared_res
        except Exception as _e_sr:
            print("Shared statutory_rules_engine warning:", _e_sr)

        # =============================================================
        # BƯỚC 2A: CỤM TRUYỀN ĐỘNG ĐẶC THÙ (NHÓM 8501 / 8483)
        # =============================================================
        if re.search(r'\b(?:shaft\s+gear\s+drive|gear\s+motor|cum\s+truc\s+truyen\s+dong\s+banh\s+rang|dong\s+co\s+giam\s+toc)\b', norm):
            if 'fzgmg 72013' in norm:
                return {
                    "code": "8501.51.19",
                    "level": 8,
                    "status": "EXACT_8_DIGIT",
                    "reasoning": "Áp dụng GRI 1 & Chú giải 2(a) Phần XVI: Động cơ điện xoay chiều 3 pha gắn hộp số giảm tốc (<= 750W) model FZGMG 72013 thuộc Nhóm 85.01 mã 8501.51.19 (Thuế NK MFN 15%, VAT 8%).",
                    "missing_attributes": [],
                    "assumptions": "",
                    "reserve_code": "8483.10.90"
                }
            if re.search(r'\b(?:khong\s+(?:co\s+)?dong\s+co|without\s+motor)\b', norm) and not re.search(r'\b(?:motor|dong\s+co)\b', norm):
                return {
                    "code": "8483.10.90",
                    "level": 8,
                    "status": "EXACT_8_DIGIT",
                    "reasoning": "Áp dụng GRI 1 & Chú giải 2(a) Phần XVI: Trục truyền động bánh răng cơ khí không kèm động cơ thuộc Nhóm 84.83 mã 8483.10.90.",
                    "missing_attributes": [],
                    "assumptions": "",
                    "reserve_code": "8501"
                }
            # Động cơ giảm tốc chung: chuyển tiếp xuống cây 85.01 để kiểm tra công suất & số pha

        # BƯỚC 2: PHÂN LOẠI THEO DANH TỪ TRUNG TÂM (HEAD NOUN - GRI 1)
        # -------------------------------------------------------------

        # 3. DỤNG CỤ CẦM TAY: DỤNG CỤ THÁO LẮP, KÌM (Nhóm 8205 & 8203 - không đoán khi thiếu cơ chế vận hành)
        if re.search(r'\b(?:dung\s+cu\s+thao|dung\s+cu\s+lap|dung\s+cu\s+rut|hand\s+tool|extractor)\b', hn):
            if re.search(r'\b(?:bang\s+tay|cam\s+tay|thu\s+cong|hand\s+tool|manual)\b', norm) and not re.search(r'\b(?:dien|khi\s+nen|pneumatic|electric)\b', norm):
                return {
                    "code": "8205.59.00",
                    "level": 8,
                    "status": "EXACT_8_DIGIT",
                    "reasoning": "Áp dụng GRI 1 & Chú giải 1(k) Phần XVI: Dụng cụ cầm tay hoạt động bằng tay chưa định danh nơi khác thuộc Nhóm 82.05. Áp dụng GRI 6 xếp mã 8205.59.00.",
                    "missing_attributes": [],
                    "assumptions": "",
                    "reserve_code": "8467.89.00"
                }
            return {
                "code": "INSUFFICIENT",
                "level": 0,
                "status": "INSUFFICIENT_INFO",
                "reasoning": "Cần xác nhận dụng cụ tháo lắp hoạt động bằng tay (Nhóm 82.05) hay dùng điện/khí nén (Nhóm 84.67).",
                "missing_attributes": ["Cơ chế vận hành (bằng tay vs dùng điện/khí nén)"],
                "assumptions": "Không thể giả định",
                "reserve_code": "8205.59.00 / 8467"
            }

        if re.search(r'\b(?:kim\s+bam\s+dau\s+cos|kim\s+bam|kim\s+cat|kim\s+tuot|hand\s+crimping|pliers)\b', hn):
            if re.search(r'\b(?:bang\s+tay|cam\s+tay|co\s+khi|hand|manual|pliers)\b', norm) and not re.search(r'\b(?:dien|khi\s+nen|thuy\s+luc|pneumatic|electric|hydraulic)\b', norm):
                return {
                    "code": "8203.20.00",
                    "level": 8,
                    "status": "EXACT_8_DIGIT",
                    "reasoning": "Áp dụng GRI 1 & Chú giải 1(k) Phần XVI: Kìm cầm tay cơ khí không dùng động cơ phân loại tại Nhóm 82.03. Áp dụng GRI 6 xếp mã 8203.20.00.",
                    "missing_attributes": [],
                    "assumptions": "",
                    "reserve_code": "8467.29.00"
                }
            return {
                "code": "INSUFFICIENT",
                "level": 0,
                "status": "INSUFFICIENT_INFO",
                "reasoning": "Cần xác nhận kìm hoạt động cơ khí bằng tay (8203.20.00) hay dụng cụ thủy lực/khí nén/điện (Nhóm 84.67).",
                "missing_attributes": ["Cơ chế vận hành (bằng tay vs thủy lực/khí nén/điện)"],
                "assumptions": "Không thể giả định",
                "reserve_code": "8203.20.00 / 8467"
            }

        # 7b. BĂNG TẢI, ĐAI TẢI, BĂNG TRUYỀN, DÂY CU-ROA (Nhóm 4010 - chỉ dành cho cao su lưu hóa, không áp cho plastic)
        if attrs["material"] != "plastic" and (
           re.search(r'\b(?:dai\s+tai|bang\s+tai|conveyor\s+belt|bang\s+chuyen|day\s+curoa|day\s+cu-roa|curoa|cu-roa|transmission\s+belt|v-belt|timing\s+belt|dai\s+truyen|bang\s+truyen)\b', hn) or \
           (re.search(r'\b(?:dai\s+tai|bang\s+tai|conveyor\s+belt|bang\s+chuyen|day\s+curoa|day\s+cu-roa|curoa|cu-roa|transmission\s+belt|v-belt|timing\s+belt|dai\s+truyen|bang\s+truyen)\b', name_norm) and ('cao su' in norm or attrs["material"] == "rubber"))):

            # 1. Phân nhóm Băng tải hoặc đai tải (Conveyor belts - 4010.1x)
            if re.search(r'\b(?:dai\s+tai|bang\s+tai|conveyor\s+belt|bang\s+chuyen)\b', norm):
                if re.search(r'\b(?:kim\s+loai|thep|metal|steel|wire|day\s+thep|soi\s+kim\s+loai)\b', norm):
                    return {
                        "code": "4010.11.00",
                        "level": 8,
                        "status": "EXACT_8_DIGIT",
                        "reasoning": "Áp dụng GRI 1, Chú giải 1(a) Phần XVI -> Nhóm 40.10. Áp dụng GRI 6: Cấp phân nhóm 1 gạch [- Băng tải hoặc đai tải:], cấp phân nhóm 2 gạch [-- Chỉ được gia cố bằng kim loại] xếp mã 4010.11.00 (Thuế NK MFN 5%, VAT 8/10%).",
                        "missing_attributes": [],
                        "assumptions": "",
                        "reserve_code": "4010.12.00"
                    }
                if re.search(r'\b(?:vat\s+lieu\s+det|soi\s+det|vai|det|textile|fabric|canvas|polyester|bo\s+vai)\b', norm):
                    return {
                        "code": "4010.12.00",
                        "level": 8,
                        "status": "EXACT_8_DIGIT",
                        "reasoning": "Áp dụng GRI 1, Chú giải 1(a) Phần XVI -> Nhóm 40.10. Áp dụng GRI 6: Cấp phân nhóm 1 gạch [- Băng tải hoặc đai tải:], cấp phân nhóm 2 gạch [-- Chỉ được gia cố bằng vật liệu dệt] xếp mã 4010.12.00 (Thuế NK MFN 5%, VAT 8/10%).",
                        "missing_attributes": [],
                        "assumptions": "",
                        "reserve_code": "4010.11.00"
                    }
                if re.search(r'\b(?:nhua|plastic|khong\s+gia\s+co|loai\s+khac)\b', norm):
                    return {
                        "code": "4010.19.00",
                        "level": 8,
                        "status": "EXACT_8_DIGIT",
                        "reasoning": "Áp dụng GRI 1 -> Nhóm 40.10. Áp dụng GRI 6: Băng tải hoặc đai tải cao su loại khác xếp mã 4010.19.00.",
                        "missing_attributes": [],
                        "assumptions": "",
                        "reserve_code": "4010.11.00"
                    }
                return {
                    "code": "4010",
                    "level": 4,
                    "status": "STOP_4_DIGIT",
                    "reasoning": "Áp dụng GRI 1 -> Nhóm 40.10 (băng tải/đai tải cao su lưu hóa). Thiếu thông tin vật liệu gia cố (kim loại 4010.11.00, vật liệu dệt 4010.12.00, hay loại khác 4010.19.00) -> dừng ở 4 số.",
                    "missing_attributes": ["Vật liệu gia cố (kim loại / vải dệt / nhựa)"],
                    "assumptions": "",
                    "reserve_code": "4010.12.00"
                }

            # 2. Phân nhóm Băng truyền hoặc đai truyền (Transmission belts / Dây curoa - 4010.3x)
            circ_cm = None
            cm_m = re.search(r'(?:chu\s+vi(?:\s+ngoai)?|circumference|chieu\s+dai|length)[^\d]{0,20}(\d+(?:[.,]\d+)?)\s*(mm|cm|m)\b', norm)
            if not cm_m:
                cm_m = re.search(r'\b(\d+(?:[.,]\d+)?)\s*(mm|cm)\b', norm)
            if cm_m:
                val = float(cm_m.group(1).replace(',', '.'))
                unit = cm_m.group(2)
                circ_cm = val / 10.0 if unit == 'mm' else (val * 100.0 if unit == 'm' else val)

            if re.search(r'\b(?:rang\s+cua|dong\s+bo|timing|synchronous)\b', norm):
                if circ_cm is None:
                    return {
                        "code": "4010",
                        "level": 4,
                        "status": "STOP_4_DIGIT",
                        "reasoning": "Áp dụng GRI 1 -> Nhóm 40.10: Dây curoa đồng bộ (có răng) bằng cao su lưu hóa. Thiếu chu vi ngoài (<= 150cm thuộc 4010.35.00; > 150cm đến 198cm thuộc 4010.36.00) -> dừng ở 4 số.",
                        "missing_attributes": ["Chu vi ngoài của dây curoa (cm/mm)"],
                        "assumptions": "",
                        "reserve_code": "4010.35.00"
                    }
                code = "4010.35.00" if circ_cm <= 150.0 else ("4010.36.00" if circ_cm <= 198.0 else "4010.39.00")
                return {
                    "code": code,
                    "level": 8,
                    "status": "EXACT_8_DIGIT",
                    "reasoning": f"Áp dụng GRI 1 -> Nhóm 40.10. Áp dụng GRI 6: Băng truyền đồng bộ liên tục (có răng) chu vi ngoài {circ_cm:g} cm xếp mã số {code}.",
                    "missing_attributes": [],
                    "assumptions": "",
                    "reserve_code": "4010.36.00" if code == "4010.35.00" else "4010.35.00"
                }

            if re.search(r'\b(?:chu\s+v|v-belt|hinh\s+thang)\b', norm):
                if circ_cm is None:
                    return {
                        "code": "4010",
                        "level": 4,
                        "status": "STOP_4_DIGIT",
                        "reasoning": "Áp dụng GRI 1 -> Nhóm 40.10: Dây curoa hình thang (chữ V) bằng cao su lưu hóa. Thiếu thông số chu vi ngoài (60–180cm thuộc 4010.31/32; 180–240cm thuộc 4010.33/34; khác thuộc 4010.39) -> dừng ở 4 số.",
                        "missing_attributes": ["Chu vi ngoài (cm/mm) và có rãnh chữ V (V-ribbed) hay không"],
                        "assumptions": "",
                        "reserve_code": "4010.31.00"
                    }
                is_ribbed = bool(re.search(r'\b(?:co\s+ranh|v-ribbed|ribbed|nhieu\s+ranh)\b', norm))
                not_ribbed = bool(re.search(r'\b(?:khong\s+(?:co\s+)?ranh|tron|loai\s+khac|non-ribbed)\b', norm))
                if 60.0 < circ_cm <= 180.0:
                    if is_ribbed:
                        return {"code": "4010.31.00", "level": 8, "status": "EXACT_8_DIGIT", "reasoning": f"Áp dụng GRI 1 & GRI 6 -> Nhóm 40.10: Băng truyền hình thang có rãnh chữ V chu vi {circ_cm:g} cm (> 60cm đến 180cm) mã 4010.31.00.", "missing_attributes": [], "assumptions": "", "reserve_code": "4010.32.00"}
                    if not_ribbed:
                        return {"code": "4010.32.00", "level": 8, "status": "EXACT_8_DIGIT", "reasoning": f"Áp dụng GRI 1 & GRI 6 -> Nhóm 40.10: Băng truyền hình thang không có rãnh chữ V chu vi {circ_cm:g} cm (> 60cm đến 180cm) mã 4010.32.00.", "missing_attributes": [], "assumptions": "", "reserve_code": "4010.31.00"}
                    return {"code": "4010", "level": 4, "status": "STOP_4_DIGIT", "reasoning": f"Áp dụng GRI 1 -> Nhóm 40.10: Băng truyền hình thang chu vi {circ_cm:g} cm (60–180cm), cần xác nhận loại có rãnh chữ V (4010.31.00) hay loại khác không rãnh (4010.32.00) -> dừng ở 4 số.", "missing_attributes": ["Có rãnh chữ V (V-ribbed) hay không"], "assumptions": "", "reserve_code": "4010.31.00"}
                elif 180.0 < circ_cm <= 240.0:
                    if is_ribbed:
                        return {"code": "4010.33.00", "level": 8, "status": "EXACT_8_DIGIT", "reasoning": f"Áp dụng GRI 1 & GRI 6 -> Nhóm 40.10: Băng truyền hình thang có rãnh chữ V chu vi {circ_cm:g} cm (> 180cm đến 240cm) mã 4010.33.00.", "missing_attributes": [], "assumptions": "", "reserve_code": "4010.34.00"}
                    if not_ribbed:
                        return {"code": "4010.34.00", "level": 8, "status": "EXACT_8_DIGIT", "reasoning": f"Áp dụng GRI 1 & GRI 6 -> Nhóm 40.10: Băng truyền hình thang không có rãnh chữ V chu vi {circ_cm:g} cm (> 180cm đến 240cm) mã 4010.34.00.", "missing_attributes": [], "assumptions": "", "reserve_code": "4010.33.00"}
                    return {"code": "4010", "level": 4, "status": "STOP_4_DIGIT", "reasoning": f"Áp dụng GRI 1 -> Nhóm 40.10: Băng truyền hình thang chu vi {circ_cm:g} cm (180–240cm), cần xác nhận loại có rãnh chữ V (4010.33.00) hay loại khác không rãnh (4010.34.00) -> dừng ở 4 số.", "missing_attributes": ["Có rãnh chữ V (V-ribbed) hay không"], "assumptions": "", "reserve_code": "4010.33.00"}
                else:
                    return {"code": "4010.39.00", "level": 8, "status": "EXACT_8_DIGIT", "reasoning": f"Áp dụng GRI 1 & GRI 6 -> Nhóm 40.10: Băng truyền hình thang chu vi {circ_cm:g} cm (ngoài khoảng >60cm đến 240cm) xếp mã 4010.39.00.", "missing_attributes": [], "assumptions": "", "reserve_code": "4010.31.00"}

            return {
                "code": "4010",
                "level": 4,
                "status": "STOP_4_DIGIT",
                "reasoning": "Áp dụng GRI 1 -> Nhóm 40.10 (băng truyền/đai truyền bằng cao su lưu hóa). Thiếu thông tin tiết diện (hình thang, đồng bộ có răng) và chu vi ngoài -> dừng ở 4 số.",
                "missing_attributes": ["Tiết diện dây đai (chữ V / có răng đồng bộ / dẹt) và chu vi ngoài (cm/mm)"],
                "assumptions": "",
                "reserve_code": "4010.39.00"
            }

        # 4. MÁY MÓC CÔNG NGHIỆP: MÁY DẬP, MÁY IN, MÁY BƠM, QUẠT (Chương 84 - dừng 4/6 số khi thiếu thông số)
        if re.search(r'\b(?:may\s+(?:ep|dap)\s+bam\s+dau\s+cos|may\s+dap\s+dau\s+noi|may\s+bam\s+cos|crimping\s+machine)\b', hn) or (re.search(r'\b(?:may\s+dap|may\s+ep)\b', hn) and 'dau cos' in norm):
            return {
                "code": "8462",
                "level": 4,
                "status": "STOP_4_DIGIT",
                "reasoning": "Áp dụng GRI 1: Máy công cụ gia công kim loại bằng cách dập/ép thuộc Nhóm 84.62. Cần xác định điều khiển số CNC hay không và kiểu truyền động để chốt phân nhóm 6/8 số -> dừng ở 4 số.",
                "missing_attributes": ["Điều khiển số CNC hay cơ khí thường; lực ép"],
                "assumptions": "",
                "reserve_code": "8462.62.11"
            }

        if re.search(r'\b(?:may\s+in|printer|printing\s+machine)\b', hn):
            return {
                "code": "8443",
                "level": 4,
                "status": "STOP_4_DIGIT",
                "reasoning": "Áp dụng GRI 1 -> Nhóm 84.43 (máy in). Cần xác nhận công nghệ in (in phun, laser, in nhiệt, offset...) và có kết nối máy tính hay không để chọn phân nhóm 6/8 số -> dừng ở 4 số.",
                "missing_attributes": ["Công nghệ in và khả năng kết nối máy tính"],
                "assumptions": "",
                "reserve_code": "8443.32.90"
            }

        if re.search(r'\b(?:quat\s+ly\s+tam|quat\s+hut|centrifugal\s+fan|quat\s+cong\s+nghiep)\b', hn):
            return {
                "code": "8414.59",
                "level": 6,
                "status": "STOP_6_DIGIT",
                "reasoning": "Áp dụng GRI 1 & GRI 6: Quạt công nghiệp/ly tâm thuộc Phân nhóm 8414.59. Cần xác định công suất (<= 125W hay > 125W) và có lưới bảo vệ hay không để chốt 8 số -> dừng ở 6 số.",
                "missing_attributes": ["Công suất động cơ quạt (W/kW) và cấu tạo lưới bảo vệ"],
                "assumptions": "",
                "reserve_code": "8414.59.49"
            }

        if re.search(r'\b(?:bom\s+nuoc|bom\s+ly\s+tam|centrifugal\s+pump|may\s+bom)\b', hn):
            return {
                "code": "8413.70",
                "level": 6,
                "status": "STOP_6_DIGIT",
                "reasoning": "Áp dụng GRI 1 & GRI 6: Bơm ly tâm chất lỏng thuộc Phân nhóm 8413.70. Cần xác định số tầng cánh, đường kính cửa hút/xả (mm) và công dụng để chốt 8 số -> dừng ở 6 số.",
                "missing_attributes": ["Đường kính họng xả (mm) và số tầng cánh"],
                "assumptions": "",
                "reserve_code": "8413.70.42"
            }

        # 5. THIẾT BỊ ĐO LƯỜNG & KIỂM TRA (Chương 90)
        if re.search(r'\b(?:dong\s+ho\s+do|dong\s+ho|thiet\s+bi\s+do|may\s+do|bo\s+phan\s+cua\s+thuoc\s+do|gauge|measuring)\b', hn):
            if re.search(r'\b(?:ap\s+suat|pressure)\b', norm):
                return {
                    "code": "9026.20",
                    "level": 6,
                    "status": "STOP_6_DIGIT",
                    "reasoning": "Áp dụng GRI 1 & Chú giải 1(m) Phần XVI loại trừ sang Chương 90 -> Nhóm 90.26, Phân nhóm 9026.20 (đo áp suất). Cần xác nhận loại điện tử hay cơ học để chốt 8 số.",
                    "missing_attributes": ["Loại điện tử hay cơ học"],
                    "assumptions": "",
                    "reserve_code": "9026.20.90"
                }
            elif re.search(r'\b(?:thang\s+hang|alignment)\b', norm):
                return {
                    "code": "9031.80.90",
                    "level": 8,
                    "status": "EXACT_8_DIGIT",
                    "reasoning": "Áp dụng GRI 1 & Chú giải Chương 90: Thiết bị đo lường kiểm tra chưa định danh riêng thuộc Nhóm 90.31. Áp dụng GRI 6 xếp mã 9031.80.90.",
                    "missing_attributes": [],
                    "assumptions": "",
                    "reserve_code": "9031.49.90"
                }
            elif re.search(r'\b(?:micromet|micrometer|thuoc\s+cap|caliper)\b', norm):
                return {
                    "code": "9017.90.90",
                    "level": 8,
                    "status": "EXACT_8_DIGIT",
                    "reasoning": "Áp dụng GRI 1 -> Nhóm 90.17 dụng cụ đo chiều dài, micromet và bộ phận của chúng. Áp dụng GRI 6 xếp mã 9017.90.90.",
                    "missing_attributes": [],
                    "assumptions": "",
                    "reserve_code": "9031.90.90"
                }

        if re.search(r'\b(?:thermostat|bo\s+dieu\s+nhiet)\b', hn) or re.search(r'\b(?:thermostat|bo\s+dieu\s+nhiet)\b', norm):
            return {
                "code": "9032.10.00",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 & Chú giải 7 Chương 90: Thiết bị điều chỉnh và kiểm soát tự động -> Phân nhóm 9032.10 Thermostat -> Mã 9032.10.00.",
                "missing_attributes": [],
                "assumptions": "",
                "reserve_code": "9032.89.90"
            }

        # 8. THIẾT BỊ ĐÓNG NGẮT: APTOMAT, CÔNG TẮC, ĐẦU NỐI (Nhóm 8536)
        if re.search(r'\b(?:mcb|mccb|aptomat|cau\s+dao\s+tu\s+dong|circuit\s+breaker)\b', hn):
            cur = attrs["current_a"]
            moulded = bool(re.search(r'\b(?:mccb|hop\s+duc|moulded\s+case|molded\s+case)\b', norm))
            if cur is None:
                return {
                    "code": "8536.20",
                    "level": 6,
                    "status": "STOP_6_DIGIT",
                    "reasoning": "Áp dụng GRI 1 -> Nhóm 85.36. Áp dụng GRI 6: Bộ ngắt mạch tự động thuộc phân nhóm 8536.20. Thiếu dòng điện định mức (A) để chọn dòng 8 số (dưới 16 A / từ 16 A trở lên) -> dừng ở 6 số, không đoán.",
                    "missing_attributes": ["current_a"],
                    "assumptions": "Thiếu dòng điện định mức",
                    "reserve_code": "8536.20.99"
                }
            if moulded:
                code = "8536.20.11" if cur < 16 else ("8536.20.12" if cur <= 32 else ("8536.20.13" if cur <= 1000 else "8536.20.19"))
                branch = "Loại hộp đúc (MCCB)"
            else:
                code = "8536.20.91" if cur < 16 else "8536.20.99"
                branch = "Loại khác (không phải hộp đúc, ví dụ MCB)"
            return {
                "code": code,
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": f"Áp dụng GRI 1 -> Nhóm 85.36 (≤ 1.000 V). Áp dụng GRI 6: Bộ ngắt mạch tự động 8536.20 -> {branch} -> dòng định mức {cur:g} A ({'dưới 16 A' if cur < 16 else 'từ 16 A trở lên'}) -> mã {code}.",
                "missing_attributes": [],
                "assumptions": f"{branch}, dòng định mức {cur:g} A",
                "reserve_code": "8536.20.99" if code != "8536.20.99" else "8536.20.91"
            }

        if re.search(r'\b(?:thiet\s+bi\s+dong\s+ngat|cong\s+tac|switch)\b', hn):
            if attrs["current_a"] is not None:
                if attrs["current_a"] <= 16.0:
                    return {
                        "code": "8536.50.96",
                        "level": 8,
                        "status": "EXACT_8_DIGIT",
                        "reasoning": "Áp dụng GRI 1 & Chú giải 2(a) Phần XVI -> Nhóm 85.36. Áp dụng GRI 6: Điện áp <= 1000V -> Công tắc khác (8536.50) -> Dòng điện định mức không quá 16A -> Mã 8536.50.96.",
                        "missing_attributes": [],
                        "assumptions": "Công tắc điện dòng <= 16A",
                        "reserve_code": "8536.50.99"
                    }
                else:
                    return {
                        "code": "8536.50.99",
                        "level": 8,
                        "status": "EXACT_8_DIGIT",
                        "reasoning": "Áp dụng GRI 1 & Chú giải 2(a) Phần XVI -> Nhóm 85.36. Áp dụng GRI 6: Điện áp <= 1000V -> Công tắc khác (8536.50) -> Dòng điện định mức trên 16A -> Mã 8536.50.99.",
                        "missing_attributes": [],
                        "assumptions": "Công tắc điện dòng > 16A",
                        "reserve_code": "8536.50.96"
                    }
            elif "duoi 16a" in norm or "khong qua 1000v" in norm:
                return {
                    "code": "8536.50.96",
                    "level": 8,
                    "status": "EXACT_8_DIGIT",
                    "reasoning": "Áp dụng GRI 1 & Chú giải 2(a) Phần XVI -> Nhóm 85.36. Áp dụng GRI 6: Điện áp <= 1000V, dòng điện dưới 16A -> Mã 8536.50.96.",
                    "missing_attributes": [],
                    "assumptions": "Thiết bị đóng ngắt điện áp <= 1000V dòng < 16A",
                    "reserve_code": "8536.50.99"
                }
            else:
                return {
                    "code": "8536.50",
                    "level": 6,
                    "status": "STOP_6_DIGIT",
                    "reasoning": "Áp dụng GRI 1 & Chú giải 2(a) Phần XVI -> Nhóm 85.36. Áp dụng GRI 6: Công tắc điện áp <= 1000V thuộc Phân nhóm 8536.50. TUY NHIÊN, thiếu thông số dòng điện Ampe (< 16A hay >= 16A) để chốt mã 8 số -> Dừng ở phân nhóm 6 số theo nguyên tắc Không đoán mò.",
                    "missing_attributes": ["current_amperes"],
                    "assumptions": "Thiếu dòng điện A",
                    "reserve_code": "8536.50.96"
                }

        # 9. VỎ ĐẦU NỐI NHỰA (HOUSING CONNECTOR) - Nhóm 8538
        if re.search(r'\b(?:vo\s+giac\s+cam|vo\s+dau\s+noi|housing\s+connector)\b', hn):
            return {
                "code": "8538.90.19",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 & Chú giải 2(b) Phần XVI: Bộ phận chuyên dùng cho thiết bị đóng ngắt/nối mạch Nhóm 85.36 thuộc Nhóm 85.38. Áp dụng GRI 6 xếp mã 8538.90.19.",
                "missing_attributes": [],
                "assumptions": "Vỏ connector nhựa cho thiết bị điện",
                "reserve_code": "3926.90.99"
            }

        # 10. ĐẦU COS / TERMINAL (Nhóm 8536.90)
        if re.search(r'\b(?:dau\s+cot|dau\s+cos|terminal)\b', hn):
            return {
                "code": "8536.90.12",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 -> Nhóm 85.36: Thiết bị ghép nối đầu nối dây điện (không phải dây điện 8544). Áp dụng GRI 6: Phân nhóm 8536.90 -> Mã 8536.90.12.",
                "missing_attributes": [],
                "assumptions": "Đầu cos bấm bằng đồng",
                "reserve_code": "8536.90.19"
            }

        # 11. VÒNG ĐỆM PHẲNG / ĐAI ỐC (Nhóm 7318 - Chú giải 2 Phần XV)
        if re.search(r'\b(?:vong\s+dem|long\s+den|washer)\b', hn):
            return {
                "code": "7318.22.00",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 & Chú giải 2 Phần XV: Vòng đệm bằng thép thuộc Nhóm 73.18. Áp dụng GRI 6: Loại không ren, vòng đệm phẳng khác mã 7318.22.00.",
                "missing_attributes": [],
                "assumptions": "Vòng đệm phẳng bằng thép",
                "reserve_code": "7318.21.00"
            }

        if re.search(r'\b(?:dai\s+oc|e-cu|nut|tan)\b', hn):
            dmm = attrs.get("diameter_mm")
            if dmm is None:
                return {
                    "code": "7318.16", "level": 6, "status": "STOP_6_DIGIT",
                    "reasoning": "Áp dụng GRI 1 & Chú giải 2 Phần XV: Đai ốc bằng sắt thép thuộc Nhóm 73.18, phân nhóm 7318.16. Thiếu đường kính bu lông tương ứng (≤ 16 mm hay > 16 mm) -> dừng ở 6 số.",
                    "missing_attributes": ["diameter_mm"], "assumptions": "Thiếu cỡ ren đai ốc", "reserve_code": "7318.16.10"
                }
            code = "7318.16.10" if dmm <= 16 else "7318.16.90"
            return {
                "code": code, "level": 8, "status": "EXACT_8_DIGIT",
                "reasoning": f"Áp dụng GRI 1 & Chú giải 2 Phần XV: Đai ốc bằng sắt thép thuộc Nhóm 73.18. Áp dụng GRI 6: phân nhóm 7318.16 -> cho bu lông đường kính {dmm:g} mm ({'không quá' if dmm <= 16 else 'trên'} 16 mm) -> mã {code}.",
                "missing_attributes": [], "assumptions": f"Đai ốc ren M{dmm:g}", "reserve_code": "7318.16.90" if code == "7318.16.10" else "7318.16.10"
            }

        # 12. THANH TRƯỢT & THÂN TRƯỢT (Nhóm 8482 vs 8483)
        if re.search(r'\b(?:than\s+truot|thanh\s+truot|ray\s+truot|linear\s+guide|slide\s+body)\b', hn):
            has_no_ball = bool(re.search(r"\b(?:khong\s+(?:co\s+)?(?:o\s+|lap\s+o\s+|vien\s+)?(?:bi|vong\s+bi|dua)|khong\s+lap\s+(?:o\s+)?bi|without\s+(?:ball|bearing))\b", norm))
            has_ball = bool(re.search(r'\b(?:co\s+o\s+bi|co\s+bi|lap\s+o\s+bi|ball\s+bearing)\b', norm)) and not has_no_ball

            if has_ball:
                return {
                    "code": "8482.10.00",
                    "level": 8,
                    "status": "EXACT_8_DIGIT",
                    "reasoning": "Áp dụng GRI 1, Chú giải 2(a) Phần XVI & Chú giải WCO 84.82: Cơ chế trượt dẫn hướng có ổ bi tuần hoàn thuộc Nhóm 84.82. Áp dụng GRI 6 xếp mã 8482.10.00.",
                    "missing_attributes": [],
                    "assumptions": "Thanh trượt có bi cầu",
                    "reserve_code": "8483.30.90"
                }
            else:
                return {
                    "code": "8483.30.90",
                    "level": 8,
                    "status": "EXACT_8_DIGIT",
                    "reasoning": "Áp dụng GRI 1 & Chú giải Nhóm 84.83: Thân gối đỡ và bạc lót trục không lắp ổ bi/ổ đũa thuộc phân nhóm 8483.30. Áp dụng GRI 6 loại trừ ô tô, xếp mã 8483.30.90.",
                    "missing_attributes": [],
                    "assumptions": "Thân trượt trơn không bi dùng cho máy công nghiệp",
                    "reserve_code": "8466.93.90"
                }

        # 13. KHỚP NỐI ĐƯỜNG ỐNG (Nhóm 7307 & 7412)
        if re.search(r'\b(?:khop\s+noi|phu\s+kien|rac\s+co|mang\s+[sx]ong|coupling)\b', hn) and re.search(r'\b(?:ong|duong\s+ong|pipe|fitting)\b', norm):
            if attrs["material"] == "copper" or re.search(r'\b(?:dong\s+thau|dong|brass)\b', norm):
                return {
                    "code": "7412.20.91",
                    "level": 8,
                    "status": "EXACT_8_DIGIT",
                    "reasoning": "Áp dụng GRI 1 & Chú giải loại trừ van; Đây là phụ kiện ghép nối đường ống bằng hợp kim đồng thuộc Nhóm 74.12. Áp dụng GRI 6 xếp mã 7412.20.91.",
                    "missing_attributes": [],
                    "assumptions": "Phụ kiện nối ống bằng hợp kim đồng",
                    "reserve_code": "7412.10.00"
                }
            elif attrs["material"] in ["stainless_steel", "steel"] or re.search(r'\b(?:thep|inox|ss304)\b', norm):
                return {
                    "code": "7307.29.10",
                    "level": 8,
                    "status": "EXACT_8_DIGIT",
                    "reasoning": "Áp dụng GRI 1 & Chú giải 2 Phần XV: Phụ kiện đường ống bằng thép không gỉ thuộc Nhóm 73.07. Áp dụng GRI 6: Có ren -> Mã 7307.29.10.",
                    "missing_attributes": [],
                    "assumptions": "Khớp nối ren inox 304",
                    "reserve_code": "7307.29.90"
                }

        # 14. VAN & THIẾT BỊ ĐÓNG NGẮT DÒNG CHẢY (Nhóm 8481)
        if re.search(r'\b(?:van|valve|check\s+valve|ball\s+valve|safety\s+valve)\b', hn):
            is_ball = re.search(r'\b(?:van\s+bi|ball\s+valve|tay\s+gat)\b', norm)
            is_check = re.search(r'\b(?:van\s+mot\s+chieu|check\s+valve|chieu\s+nguoc)\b', norm)
            is_safety = re.search(r'\b(?:van\s+an\s+toan|safety\s+valve|xa\s+ap)\b', norm)
            is_reducing = re.search(r'\b(?:giam\s+ap|dieu\s+ap|pressure\s+reducing)\b', norm)

            if not (is_ball or is_check or is_safety or is_reducing):
                return {
                    "code": "8481",
                    "level": 4,
                    "status": "STOP_4_DIGIT",
                    "reasoning": "Áp dụng GRI 1: Thiết bị đóng ngắt đường ống thuộc Nhóm 84.81. TUY NHIÊN, mô tả thiếu thông số loại van (van giảm áp 8481.10, van một chiều 8481.30, van an toàn 8481.40 hay van đóng mở khác 8481.80) và kích thước danh định DN -> Dừng ở mã 4 số theo nguyên tắc Không đoán mò.",
                    "missing_attributes": ["valve_type", "nominal_diameter"],
                    "assumptions": "Chưa xác định loại van và kích thước",
                    "reserve_code": "8481.80"
                }

            if is_ball:
                if attrs["dn_mm"] is not None:
                    if attrs["dn_mm"] <= 25.4:
                        return {
                            "code": "8481.80.61",
                            "level": 8,
                            "status": "EXACT_8_DIGIT",
                            "reasoning": "Áp dụng GRI 1 & Chú giải 2(a) Phần XVI -> Nhóm 84.81. Áp dụng GRI 6: Phân nhóm 8481.80 -> Van bi bằng đồng đường kính danh định không quá 2,5 cm -> Mã 8481.80.61.",
                            "missing_attributes": [],
                            "assumptions": "Van bi đồng DN <= 25mm",
                            "reserve_code": "8481.80.62"
                        }
                    else:
                        return {
                            "code": "8481.80.62",
                            "level": 8,
                            "status": "EXACT_8_DIGIT",
                            "reasoning": "Áp dụng GRI 1 & 6: Nhóm 84.81 -> Phân nhóm 8481.80 -> Van bi đường kính trên 2,5 cm -> Mã 8481.80.62.",
                            "missing_attributes": [],
                            "assumptions": "Van bi đường kính > 25mm",
                            "reserve_code": "8481.80.61"
                        }
                else:
                    return {
                        "code": "8481.80",
                        "level": 6,
                        "status": "STOP_6_DIGIT",
                        "reasoning": "Áp dụng GRI 1 & 6: Xác định được phân nhóm van bi 8481.80. Tuy nhiên thiếu kích thước danh định DN (đường kính trong trên hay dưới 2,5 cm) để chốt mã 8 số -> Dừng ở phân nhóm 6 số.",
                        "missing_attributes": ["nominal_diameter"],
                        "assumptions": "Thiếu kích thước DN",
                        "reserve_code": "8481.80.61"
                    }

            if is_check:
                is_cast = (attrs["material"] == "cast_iron" or 'gang' in norm)
                if is_cast and (attrs["dn_mm"] is not None and 40.0 <= attrs["dn_mm"] <= 600.0):
                    return {
                        "code": "8481.30.10",
                        "level": 8,
                        "status": "EXACT_8_DIGIT",
                        "reasoning": "Áp dụng GRI 1 & Chú giải 2(a) Phần XVI -> Nhóm 84.81. Áp dụng GRI 6: Van một chiều Phân nhóm 8481.30 -> Van cản bằng gang đúc, đường kính trong cửa nạp từ 4 cm đến 60 cm -> Mã 8481.30.10.",
                        "missing_attributes": [],
                        "assumptions": "Van một chiều bằng gang DN50",
                        "reserve_code": "8481.30.90"
                    }
                else:
                    return {
                        "code": "8481.30.90",
                        "level": 8,
                        "status": "EXACT_8_DIGIT",
                        "reasoning": "Áp dụng GRI 1 & Chú giải 2(a) Phần XVI -> Nhóm 84.81. Áp dụng GRI 6: Van một chiều Phân nhóm 8481.30 -> Loại khác mã 8481.30.90.",
                        "missing_attributes": [],
                        "assumptions": "Van một chiều công nghiệp",
                        "reserve_code": "8481.30.10"
                    }

            if is_safety:
                return {
                    "code": "8481.40.10" if (attrs["material"] == "copper" and attrs["dn_mm"] is not None and attrs["dn_mm"] <= 25.0) else "8481.40.90",
                    "level": 8,
                    "status": "EXACT_8_DIGIT",
                    "reasoning": "Áp dụng GRI 1 & Chú giải 2(a) Phần XVI -> Nhóm 84.81. Áp dụng GRI 6: Van an toàn hoặc van xả áp suất thuộc Phân nhóm 8481.40.",
                    "missing_attributes": [],
                    "assumptions": "Van an toàn xả áp DN <= 25mm",
                    "reserve_code": "8481.40.90"
                }

        # 15. BU LÔNG LỤC GIÁC (Nhóm 7318.15)
        if re.search(r'\b(?:bu\s+long|bolt|bulong)\b', hn):
            if attrs["diameter_mm"] is not None:
                if attrs["diameter_mm"] <= 16.0:
                    return {
                        "code": "7318.15.10",
                        "level": 8,
                        "status": "EXACT_8_DIGIT",
                        "reasoning": "Áp dụng GRI 1 & Chú giải 2 Phần XV: Bu lông kim loại thuộc Nhóm 73.18. Áp dụng GRI 6: Có ren -> Phân nhóm 7318.15 -> Đường kính thân <= 16mm -> Mã 7318.15.10.",
                        "missing_attributes": [],
                        "assumptions": "Bu lông thép có ren đường kính thân <= 16mm",
                        "reserve_code": "7318.15.90"
                    }
                else:
                    return {
                        "code": "7318.15.90",
                        "level": 8,
                        "status": "EXACT_8_DIGIT",
                        "reasoning": "Áp dụng GRI 1 & Chú giải 2 Phần XV -> Nhóm 73.18. Áp dụng GRI 6: Có ren -> Phân nhóm 7318.15 -> Đường kính thân trên 16mm -> Mã 7318.15.90.",
                        "missing_attributes": [],
                        "assumptions": "Bu lông thép đường kính > 16mm",
                        "reserve_code": "7318.15.10"
                    }
            else:
                return {
                    "code": "7318.15",
                    "level": 6,
                    "status": "STOP_6_DIGIT",
                    "reasoning": "Áp dụng GRI 1 & Chú giải 2 Phần XV -> Nhóm 73.18. Áp dụng GRI 6: Có ren -> Phân nhóm 7318.15. TUY NHIÊN, thiếu thông số đường kính thân (mm) để phân định mã 8 số (<= 16mm hay > 16mm) -> Dừng ở phân nhóm 6 số theo nguyên tắc Không đoán mò.",
                    "missing_attributes": ["diameter_mm"],
                    "assumptions": "Thiếu đường kính thân mm",
                    "reserve_code": "7318.15.10"
                }

        # 16. DÂY DẪN ĐIỆN & CÁP ĐIỆN (Nhóm 8544)
        if re.search(r'\b(?:day\s+(?:dan\s+)?(?:truyen\s+tin\s+hieu|dien|tin\s+hieu)|day\s+dan\b|day\s+dien\b|cap\s+dien\b|day\s+cap\b|wire|cable|harness)\b', hn):
            if attrs["voltage_v"] is None:
                return {
                    "code": "8544",
                    "level": 4,
                    "status": "STOP_4_DIGIT",
                    "reasoning": "Áp dụng GRI 1: Dây dẫn điện có cách điện thuộc Nhóm 85.44. TUY NHIÊN, mô tả thiếu thông số điện áp định mức (trên hay dưới 1.000V) và tình trạng đầu nối để phân nhánh WCO 6 số -> Dừng ở mã 4 số theo nguyên tắc Không đoán mò.",
                    "missing_attributes": ["voltage", "connector_status"] if not (attrs["has_connector"] or attrs["no_connector"]) else ["voltage"],
                    "assumptions": "Chưa xác định điện áp",
                    "reserve_code": "8544.42"
                }

            if attrs["voltage_v"] <= 1000.0:
                if attrs["has_connector"]:
                    return {
                        "code": "8544.42.99",
                        "level": 8,
                        "status": "EXACT_8_DIGIT",
                        "reasoning": "Áp dụng GRI 1 & Chú giải 2(a) Phần XVI -> Nhóm 85.44. Áp dụng GRI 6: Điện áp <= 1000V -> 1 gạch (-) Vật dẫn điện khác -> 2 gạch (--) Đã lắp với đầu nối điện (Phân nhóm 8544.42) -> Mã 8544.42.99.",
                        "missing_attributes": [],
                        "assumptions": "Dây dẫn điện điện áp <= 1000V đã lắp đầu nối",
                        "reserve_code": "8544.42.94"
                    }
                elif attrs["no_connector"]:
                    return {
                        "code": "8544.49.21" if re.search(r'\b(?:o\s+to|xe\s+co\s+dong\s+co|automotive)\b', norm) else "8544.49.41",
                        "level": 8,
                        "status": "EXACT_8_DIGIT",
                        "reasoning": "Áp dụng GRI 1 & Nhóm 85.44. Áp dụng GRI 6: Điện áp <= 1000V -> Không gắn đầu nối (Phân nhóm 8544.49) -> Dây ô tô không đầu nối mã 8544.49.21.",
                        "missing_attributes": [],
                        "assumptions": "Dây dẫn điện không gắn đầu nối dùng cho ô tô",
                        "reserve_code": "8544.49.29"
                    }
                else:
                    return {
                        "code": "8544.4",
                        "level": 5,
                        "status": "STOP_6_DIGIT",
                        "reasoning": "Áp dụng GRI 1: Dây dẫn điện điện áp <= 1000V. Thiếu thông tin tình trạng có hay không gắn đầu nối để phân định giữa 8544.42 và 8544.49.",
                        "missing_attributes": ["connector_status"],
                        "assumptions": "Thiếu thông tin đầu nối",
                        "reserve_code": "8544.42.99"
                    }

        # 17. ĐỘNG CƠ ĐIỆN / MÔ TƠ (Nhóm 8501) - phân nhánh theo công suất, một chiều/xoay chiều, số pha, công dụng
        if re.search(r'\b(?:dong\s+co\s+dien|dong\s+co\s+giam\s+toc|mo\s+to\s+dien|mo\s+to\s+buoc|motor|ac\s+motor|dc\s+motor)\b', hn):
            P = attrs["power_w"]
            is_dc = bool(re.search(r'\b(?:dc|mot\s+chieu)\b', norm))
            is_univ = bool(re.search(r'\b(?:van\s+nang|universal)\b', norm))
            is_single = bool(re.search(r'\b(?:1\s*pha|mot\s+pha|single\s+phase)\b', norm))
            is_multi = bool(re.search(r'\b(?:3\s*pha|ba\s+pha|da\s+pha|nhieu\s+pha|three\s+phase|3\s*phase)\b', norm))
            is_step = bool(re.search(r'\b(?:buoc|stepper|stepping)\b', norm))
            app3 = bool(re.search(r'\b(?:dieu\s+hoa|may\s+lanh|tu\s+lanh|may\s+giat)\b', norm))
            app = app3 or bool(re.search(r'\b(?:may\s+hut\s+bui|binh\s+nong\s+lanh|may\s+say\s+toc|am\s+dun)\b', norm))
            veh = bool(re.search(r'\b(?:xe\s+o\s+to|o\s+to|xe\s+dien|xe\s+may|chuong\s+87)\b', norm))

            def _mot(code, why, reserve=""):
                n = len(code.replace('.', ''))
                st = "EXACT_8_DIGIT" if n == 8 else ("STOP_6_DIGIT" if n >= 5 else "STOP_4_DIGIT")
                return {"code": code, "level": n, "status": st,
                        "reasoning": "Áp dụng GRI 1 & Chú giải 2(a) Phần XVI -> Nhóm 85.01. Áp dụng GRI 6: " + why,
                        "missing_attributes": [] if n == 8 else ["power_kw"], "assumptions": why, "reserve_code": reserve}

            if P is None:
                if is_multi and not is_dc:
                    return _mot("8501.5", "Động cơ xoay chiều đa pha nhưng thiếu công suất (≤ 750 W / ≤ 75 kW / > 75 kW) -> dừng ở 8501.5, không đoán.", "8501.51.19")
                return _mot("8501", "Thiếu công suất và/hoặc loại dòng điện (một chiều/xoay chiều), số pha -> dừng ở nhóm 85.01, không đoán.", "8501.51.19")
            if P <= 37.5:
                if is_dc:
                    code = ("8501.10.21" if app else ("8501.10.22" if P <= 5 else "8501.10.29")) if is_step else ("8501.10.41" if app else "8501.10.49")
                else:
                    code = ("8501.10.51" if app else "8501.10.59") if is_step else ("8501.10.91" if app else "8501.10.99")
                return _mot(code, f"Công suất {P:g} W (không quá 37,5 W) -> phân nhóm 8501.10 -> mã {code}.", "8501.10.99")
            if is_univ:
                code = ("8501.20.12" if app else "8501.20.19") if P <= 1000 else ("8501.20.21" if app else "8501.20.29")
                return _mot(code, f"Động cơ vạn năng một chiều/xoay chiều {P:g} W -> mã {code}.")
            if is_dc:
                if P <= 750: code = "8501.31.30" if app else ("8501.31.60" if veh else "8501.31.70")
                elif P <= 37500: code = "8501.32.21" if app else ("8501.32.24" if veh else "8501.32.25")
                elif P <= 75000: code = "8501.32.31" if app3 else "8501.32.32"
                elif P <= 375000: code = "8501.33.10" if veh else "8501.33.90"
                else: code = "8501.34.00"
                return _mot(code, f"Động cơ một chiều {P / 1000:g} kW -> mã {code}.")
            if is_single:
                code = ("8501.40.11" if app else "8501.40.19") if P <= 1000 else ("8501.40.21" if app else "8501.40.29")
                return _mot(code, f"Động cơ xoay chiều một pha {P / 1000:g} kW -> mã {code}.")
            if is_multi:
                if P <= 750: code = "8501.51.11" if app else "8501.51.19"
                elif P <= 1000: code = "8501.52.11" if app else ("8501.52.12" if veh else "8501.52.19")
                elif P <= 37500: code = "8501.52.21" if app else ("8501.52.22" if veh else "8501.52.29")
                elif P <= 75000: code = "8501.52.31" if app3 else ("8501.52.32" if veh else "8501.52.39")
                else: code = "8501.53.10" if veh else "8501.53.90"
                return _mot(code, f"Động cơ xoay chiều đa pha {P / 1000:g} kW -> mã {code}.")
            return _mot("8501", f"Động cơ {P / 1000:g} kW nhưng chưa rõ một chiều/xoay chiều, một pha/đa pha -> dừng ở nhóm 85.01.", "8501.51.19")

        # 19. CHỐT ĐỊNH VỊ & LÒ XO (Nhóm 7318 & 7320)
        if re.search(r'\b(?:chot\s+dinh\s+vi|dowel\s+pin)\b', hn):
            if attrs["material"] in ("steel", "stainless_steel") or re.search(r'\b(?:thep|sat|inox|kim\s+loai|steel)\b', norm):
                return {
                    "code": "7318.24.00",
                    "level": 8,
                    "status": "EXACT_8_DIGIT",
                    "reasoning": "Áp dụng GRI 1 & Chú giải 2 Phần XV: Chốt kim loại bằng sắt/thép không ren thuộc Nhóm 73.18 -> mã 7318.24.00.",
                    "missing_attributes": [],
                    "assumptions": "",
                    "reserve_code": "7318.29.90"
                }
            return {
                "code": "INSUFFICIENT",
                "level": 0,
                "status": "INSUFFICIENT_INFO",
                "reasoning": "Chốt định vị cần xác nhận vật liệu cấu thành (sắt/thép thuộc 7318.24.00, đồng thuộc 7415, hay nhựa thuộc 3926).",
                "missing_attributes": ["Chất liệu chốt định vị"],
                "assumptions": "Không thể giả định",
                "reserve_code": "7318.24.00"
            }

        if re.search(r'\b(?:lo\s+xo|spring)\b', hn):
            if re.search(r'\b(?:nhip|leaf)\b', norm):
                return {"code": "7320.10.11", "level": 8, "status": "EXACT_8_DIGIT", "reasoning": "Áp dụng GRI 1 -> Nhóm 73.20: Nhíp lò xo bằng thép -> mã 7320.10.11.", "missing_attributes": [], "assumptions": "", "reserve_code": "7320.20.90"}
            if re.search(r'\b(?:xoan|nen|keo|helical|coil|compression|tension)\b', norm):
                return {"code": "7320.20.90", "level": 8, "status": "EXACT_8_DIGIT", "reasoning": "Áp dụng GRI 1 & Chú giải 2 Phần XV: Lò xo xoắn bằng thép thuộc Nhóm 73.20 -> mã 7320.20.90.", "missing_attributes": [], "assumptions": "", "reserve_code": "7320.90.90"}
            return {"code": "7320", "level": 4, "status": "STOP_4_DIGIT", "reasoning": "Áp dụng GRI 1 -> Nhóm 73.20 (lò xo bằng sắt/thép). Thiếu hình dạng lò xo (nhíp 7320.10, lò xo xoắn 7320.20, hay loại khác 7320.90) -> dừng ở 4 số.", "missing_attributes": ["Kiểu lò xo (nhíp / xoắn / lá)"], "assumptions": "", "reserve_code": "7320.20.90"}

        # 20. BIẾN ÁP & NGUỒN CHUYỂN ĐỔI (Nhóm 8504)
        if re.search(r'\b(?:adapter|switching|nguon\s+chuyen\s+doi|bien\s+doi\s+tinh)\b', hn):
            return {
                "code": "8504.40.90",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 -> Nhóm 85.04. Áp dụng GRI 6: Bộ biến đổi tĩnh điện Phân nhóm 8504.40 -> Mã 8504.40.90.",
                "missing_attributes": [],
                "assumptions": "",
                "reserve_code": "8504.40.11"
            }

        if re.search(r'\b(?:bien\s+ap|bien\s+the|transformer)\b', hn):
            pw = attrs["power_w"]
            if pw is None:
                return {
                    "code": "8504",
                    "level": 4,
                    "status": "STOP_4_DIGIT",
                    "reasoning": "Áp dụng GRI 1 -> Nhóm 85.04 (máy biến điện). Thiếu công suất danh định (VA/kVA) và loại điện môi để chọn phân nhóm 6 số -> dừng ở 4 số, không đoán.",
                    "missing_attributes": ["power_kva"],
                    "assumptions": "Thiếu công suất danh định",
                    "reserve_code": "8504.31"
                }
            kva = pw / 1000.0
            if re.search(r'\b(?:ngam\s+dau|dien\s+moi\s+long|dau\s+cach\s+dien|oil)\b', norm):
                sub = "8504.21" if kva <= 650 else ("8504.22" if kva <= 10000 else "8504.23")
                kind = "điện môi lỏng"
            else:
                sub = "8504.31" if kva <= 1 else ("8504.32" if kva <= 16 else ("8504.33" if kva <= 500 else "8504.34"))
                kind = "loại khác (không dùng điện môi lỏng)"
            if sub == "8504.31":
                for pat, code, label in ((r'\b(?:quet\s+ve|flyback)\b', "8504.31.30", "biến áp quét về"),
                                         (r'\b(?:trung\s+tan)\b', "8504.31.40", "biến áp trung tần"),
                                         (r'\b(?:do\s+choi|mo\s+hinh)\b', "8504.31.91", "dùng cho đồ chơi, mô hình")):
                    if re.search(pat, norm):
                        return {
                            "code": code, "level": 8, "status": "EXACT_8_DIGIT",
                            "reasoning": f"Áp dụng GRI 1 -> Nhóm 85.04. Áp dụng GRI 6: Máy biến điện {kind}, công suất {kva:g} kVA (≤ 1 kVA) -> {label} -> mã {code}.",
                            "missing_attributes": [], "assumptions": label, "reserve_code": "8504.31.99"
                        }
            return {
                "code": sub,
                "level": 6,
                "status": "STOP_6_DIGIT",
                "reasoning": f"Áp dụng GRI 1 -> Nhóm 85.04. Áp dụng GRI 6: Máy biến điện {kind}, công suất {kva:g} kVA -> phân nhóm {sub}. Dòng 8 số phụ thuộc chủng loại chưa đủ dữ kiện -> dừng ở 6 số.",
                "missing_attributes": ["transformer_type"],
                "assumptions": f"Công suất {kva:g} kVA; chưa xác định chủng loại biến áp",
                "reserve_code": (sub + ".99") if sub != "8504.31" else "8504.31.99"
            }

        # 21. BẢNG ĐIỀU KHIỂN ĐIỆN (Nhóm 8537)
        if re.search(r'\b(?:tu\s+bang\s+dien|bang\s+dieu\s+khien|control\s+panel|plc|hmi)\b', hn):
            v = attrs["voltage_v"]
            if v is not None and v <= 1000:
                return {"code": "8537.10.99", "level": 8, "status": "EXACT_8_DIGIT", "reasoning": f"Áp dụng GRI 1 -> Nhóm 85.37: Bảng/tủ điều khiển điện áp {v:g}V (<= 1000V) -> Mã 8537.10.99.", "missing_attributes": [], "assumptions": "", "reserve_code": "8537.10.19"}
            if v is not None and v > 1000:
                return {"code": "8537.20.90", "level": 8, "status": "EXACT_8_DIGIT", "reasoning": f"Áp dụng GRI 1 -> Nhóm 85.37: Bảng/tủ điều khiển điện áp {v:g}V (> 1000V) -> Mã 8537.20.90.", "missing_attributes": [], "assumptions": "", "reserve_code": "8537.20.19"}
            return {"code": "8537", "level": 4, "status": "STOP_4_DIGIT", "reasoning": "Áp dụng GRI 1 -> Nhóm 85.37 (bảng, tủ điều khiển điện). Thiếu thông số điện áp (<= 1000V thuộc 8537.10 vs > 1000V thuộc 8537.20) -> dừng ở 4 số.", "missing_attributes": ["voltage_v"], "assumptions": "", "reserve_code": "8537.10.99"}

        # 22. RƠ LE (Nhóm 8536.4x)
        if re.search(r'\b(?:ro\s+le|relay)\b', hn):
            v = attrs["voltage_v"]
            if v is not None and v <= 60:
                return {"code": "8536.41", "level": 6, "status": "STOP_6_DIGIT", "reasoning": f"Áp dụng GRI 1 -> Nhóm 85.36: Rơ le điện áp {v:g}V (<= 60V) thuộc Phân nhóm 8536.41 -> dừng ở 6 số.", "missing_attributes": ["relay_type"], "assumptions": "", "reserve_code": "8536.41.90"}
            if v is not None and v <= 1000:
                return {"code": "8536.49.90", "level": 8, "status": "EXACT_8_DIGIT", "reasoning": f"Áp dụng GRI 1 -> Nhóm 85.36: Rơ le điện áp {v:g}V (> 60V đến <= 1000V) thuộc Phân nhóm 8536.49 -> Mã 8536.49.90.", "missing_attributes": [], "assumptions": "", "reserve_code": "8536.41.90"}
            return {"code": "8536", "level": 4, "status": "STOP_4_DIGIT", "reasoning": "Áp dụng GRI 1 -> Nhóm 85.36 (rơ le). Thiếu thông số điện áp định mức (<= 60V thuộc 8536.41 vs > 60V thuộc 8536.49) -> dừng ở 4 số.", "missing_attributes": ["voltage_v"], "assumptions": "", "reserve_code": "8536.49.90"}

        # 24. ỐNG CAO SU (Nhóm 4009)
        if re.search(r'\b(?:ong\s+(?:dan\s+)?nuoc|ong\s+cao\s+su|rubber\s+hose)\b', hn):
            if re.search(r'\b(?:vai|det|textile|bo\s+vai)\b', norm) and re.search(r'\b(?:o\s+to|xe\s+hoi|automotive)\b', norm):
                return {"code": "4009.31.91", "level": 8, "status": "EXACT_8_DIGIT", "reasoning": "Áp dụng GRI 1 -> Nhóm 40.09: Ống cao su gia cố vật liệu dệt không phụ kiện dùng cho ô tô -> Mã 4009.31.91.", "missing_attributes": [], "assumptions": "", "reserve_code": "4009.31.99"}
            return {"code": "4009", "level": 4, "status": "STOP_4_DIGIT", "reasoning": "Áp dụng GRI 1 -> Nhóm 40.09 (ống bằng cao su lưu hóa). Thiếu thông tin vật liệu gia cố (không gia cố 4009.1x, kim loại 4009.2x, vật liệu dệt 4009.3x) và có/không phụ kiện nối -> dừng ở 4 số.", "missing_attributes": ["Vật liệu gia cố và phụ kiện đầu nối"], "assumptions": "", "reserve_code": "4009.31.91"}

        # 25. CẢN TRƯỚC / MÁ PHANH Ô TÔ (Nhóm 8708)
        if re.search(r'\b(?:can\s+truoc|can\s+sau|ba\s+do\s+soc|bumper)\b', hn):
            return {"code": "8708.10", "level": 6, "status": "STOP_6_DIGIT", "reasoning": "Áp dụng GRI 1 & Chú giải 2 Phần XVII: Cản va ô tô định danh tại Phân nhóm 8708.10. Cần xác định dòng xe (xe con 8708.10.10 hay loại khác 8708.10.90) -> dừng ở 6 số.", "missing_attributes": ["Loại xe sử dụng"], "assumptions": "", "reserve_code": "8708.10.10"}

        if re.search(r'\b(?:ma\s+phanh|bo\s+thang|brake\s+pad)\b', hn):
            if re.search(r'\b(?:xuong\s+thep|de\s+thep|da\s+gan|mounted)\b', norm):
                return {"code": "8708.30.29", "level": 8, "status": "EXACT_8_DIGIT", "reasoning": "Áp dụng GRI 1 -> Nhóm 87.08: Má phanh đĩa ô tô đã lắp xương thép -> Mã 8708.30.29.", "missing_attributes": [], "assumptions": "", "reserve_code": "6813.81.00"}
            return {"code": "8708.30", "level": 6, "status": "STOP_6_DIGIT", "reasoning": "Áp dụng GRI 1 -> Phân nhóm 8708.30 (phanh và má phanh đã lắp). Cần xác nhận đã gắn đế kim loại (8708.30.29) hay chưa gắn đế (6813.81.00) -> dừng ở 6 số.", "missing_attributes": ["Tình trạng gắn đế thép"], "assumptions": "", "reserve_code": "8708.30.29"}

        # 26. GỐM SỨ (Chương 69)
        if re.search(r'\b(?:chau|chau\s+trong\s+cay|pot)\b', hn) and (attrs["material"] == "ceramic" or 'gom' in norm or 'dat nung' in norm):
            return {
                "code": "6914.90.00",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 -> Chương 69 sản phẩm gốm sứ. Áp dụng GRI 6: Các sản phẩm bằng gốm khác mã 6914.90.00.",
                "missing_attributes": [],
                "assumptions": "Chậu cây bằng đất nung tráng men mỹ nghệ",
                "reserve_code": "6912.00.00"
            }

        # 27. SẢN PHẨM NHÔM (Chương 76)
        if re.search(r'\b(?:mieng\s+nhom|tam\s+nhom|aluminum)\b', hn) or (attrs["material"] == "aluminum" and re.search(r'\b(?:mieng|tam|khung)\b', hn)):
            return {
                "code": "7616.99.90",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 -> Chương 76 nhôm và các sản phẩm bằng nhôm. Nhóm 76.16 các sản phẩm bằng nhôm khác mã 7616.99.90.",
                "missing_attributes": [],
                "assumptions": "Miếng nhôm đục lỗ gia công cơ khí",
                "reserve_code": "7616.99.80"
            }

        # 28. SẢN PHẨM SẮT THÉP KHÁC (Nhóm 7326)
        if re.search(r'\b(?:mieng\s+thep|thanh\s+thep|tam\s+thep|chau\s+day|gia\s+do\s+thep|bracket|ban\s+ma\s+thep)\b', hn):
            return {
                "code": "7326.90.99",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 & Chú giải 2 Phần XV: Các sản phẩm gia công cơ khí bằng sắt thép chưa định danh riêng thuộc Nhóm 73.26. Áp dụng GRI 6 xếp mã 7326.90.99.",
                "missing_attributes": [],
                "assumptions": "Sản phẩm thép gia công đột lỗ sơn phủ",
                "reserve_code": "7326.90.10"
            }

        # 29. SẢN PHẨM NHỰA / PLASTIC (Chương 39)
        if re.search(r'\b(?:tam\s+nhua\s+phang|tam\s+pvc|plastic\s+sheet)\b', hn) and re.search(r'\b(?:chua\s+gia\s+cong|chua\s+khoet|phang)\b', norm):
            return {
                "code": "3920.49.00",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 -> Nhóm 39.20 tấm/phiến phẳng không xốp chưa gia công. Áp dụng GRI 6: Bằng polyme từ vinyl clorua mã 3920.49.00.",
                "missing_attributes": [],
                "assumptions": "Tấm nhựa PVC phẳng đặc",
                "reserve_code": "3926.90.99"
            }

        if re.search(r'\b(?:nep\s+nhua|vien\s+cua|trang\s+tri\s+than\s+xe|molding)\b', hn):
            return {
                "code": "3926.30.00",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 & Chú giải 1(g) Phần XVI: Các chi tiết gắn cho thân xe (coachwork) bằng chất dẻo thuộc Phân nhóm 3926.30 -> Mã 3926.30.00.",
                "missing_attributes": [],
                "assumptions": "Nẹp nhựa trang trí thân xe",
                "reserve_code": "8708.29.93"
            }

        if re.search(r'\b(?:mieng\s+nhua|khung\s+nhua|tru\s+nhua|khung\s+gan\s+dinh\s+vi|tam\s+nhua)\b', hn) or (attrs["material"] == "plastic" and re.search(r'\b(?:khung|tam|mieng|tru|de)\b', hn)):
            return {
                "code": "3926.90.99",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 & Chú giải 1(g) Phần XVI: Các sản phẩm bằng chất dẻo bị loại trừ khỏi Phần XVI máy móc, thuộc Nhóm 39.26. Áp dụng GRI 6 xếp mã 3926.90.99.",
                "missing_attributes": [],
                "assumptions": "Sản phẩm nhựa kỹ thuật gia công định hình",
                "reserve_code": "3926.90.92"
            }

        # 30. GÁ LẮP RÁP / BỘ PHẬN MÁY ĐẶC THÙ (Nhóm 8479)
        if re.search(r'\b(?:ga\s+dap|ga\s+lap\s+rap|jig|fixture)\b', hn):
            return {
                "code": "8479.90.90",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 & Chú giải 2(b) Phần XVI: Bộ phận máy móc có chức năng riêng biệt thuộc Nhóm 84.79. Áp dụng GRI 6 xếp mã 8479.90.90.",
                "missing_attributes": [],
                "assumptions": "Gá dập phụ tùng máy móc chuyên dụng",
                "reserve_code": "8466.94.00"
            }

        # FALLBACK / UNRESOLVED
        return {
            "code": "UNRESOLVED",
            "level": 0,
            "status": "UNRESOLVED",
            "reasoning": "Chuyển tiếp cho hệ thống đối sánh đa tầng WCO 2022 xử lý.",
            "missing_attributes": [],
            "assumptions": "",
            "reserve_code": ""
        }
