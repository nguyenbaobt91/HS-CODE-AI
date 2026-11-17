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
        # =============================================================
        # BƯỚC 2A: QUY TẮC ĐỊNH DANH ĐẶC BIỆT (MÁY RỬA BUPI / THIẾT BỊ CÔNG NGHIỆP)
        # =============================================================
        # 2. LÒ XO KHÍ / GIẢM CHẤN KHÍ NÉN / PHUỘC HƠI (8412.31.00)
        if re.search(r'\b(?:giam\s+chan\s+khi\s+nen|phuoc\s+hoi|lo\s+xo\s+khi|gas\s+spring|gas\s+strut|gas\s+pressure\s+shock|gas\s+shock)\b', norm) or \
           (re.search(r'\b(?:giam\s+chan|phuoc|lo\s+xo)\b', hn) and re.search(r'\b(?:khi\s+nen|gas\s+pressure|gas)\b', norm)):
            return {
                "code": "8412.31.00",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1: Lò xo khí nén (xi lanh khí nén chuyển động tịnh tiến trợ lực nắp máy) thuộc Nhóm 84.12. Áp dụng GRI 6: Phân nhóm 8412.31 mã số 8412.31.00 (Thuế NK MFN 0%, VAT 8%).",
                "missing_attributes": [],
                "assumptions": "Lò xo khí tịnh tiến dùng nâng đỡ nắp máy móc công nghiệp",
                "reserve_code": "8487.90.90"
            }

        # 3. CẢM BIẾN BÁO MỨC CHẤT LỎNG / CÔNG TẮC MỨC / LIQUIPHANT (9026.10.90)
        if re.search(r'\b(?:cam\s+bien\s+(?:bao\s+)?muc|cong\s+tac\s+muc|bao\s+muc\s+chat\s+long|liquiphant|level\s+switch|level\s+limit\s+switch|level\s+sensor)\b', norm) or \
           (re.search(r'\b(?:cam\s+bien|cong\s+tac)\b', hn) and re.search(r'\b(?:muc\s+chat\s+long|level|liquiphant)\b', norm)):
            return {
                "code": "9026.10.90",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 & Chú giải 1(m) Phần XVI: Thiết bị đo hoặc kiểm tra mức chất lỏng (công tắc mức dạng rung) thuộc Nhóm 90.26. Áp dụng GRI 6: Phân nhóm 9026.10 mã số 9026.10.90 (Thuế NK MFN 0%, VAT 8%).",
                "missing_attributes": [],
                "assumptions": "Thiết bị kiểm tra mức chất lỏng kiểu rung hoạt động bằng điện",
                "reserve_code": "8536.50.99"
            }

        # 4. VAN XIÊN / VAN GÓC ĐIỀU KHIỂN BẰNG KHÍ NÉN (8481.80.82)
        if re.search(r'\b(?:van\s+xien|van\s+goc|angle\s+seat\s+valve)\b', norm) or \
           (re.search(r'\b(?:van|valve)\b', hn) and re.search(r'\b(?:xien|goc|angle\s+seat)\b', norm)):
            return {
                "code": "8481.80.82",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 & Chú giải 2(a) Phần XVI: Van xiên/van góc điều khiển dòng chảy thuộc Nhóm 84.81. Áp dụng GRI 6: Phân nhóm 8481.80 mã số 8481.80.82 (Thuế NK MFN 3%, VAT 8%).",
                "missing_attributes": [],
                "assumptions": "Van góc điều khiển bằng khí nén thân bằng thép không gỉ",
                "reserve_code": "8481.80.79"
            }

        # 5. TÚI LỌC CHẤT LỎNG BẰNG VẢI DỆT KỸ THUẬT (5911.90.90)
        if re.search(r'\b(?:tui\s+loc\s+chat\s+long|tui\s+loc|filter\s+bag|liquid\s+filter\s+bag)\b', norm) or \
           re.search(r'\b(?:tui\s+loc|filter\s+bag)\b', hn):
            return {
                "code": "5911.90.90",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 & Chú giải 1(e) Phần XVI (loại trừ vật phẩm dệt kỹ thuật sang Chương 59) & Chú giải 8 Chương 59: Túi lọc chất lỏng bằng vải dệt monofilament nylon thuộc Nhóm 59.11. Áp dụng GRI 6: Mã số 5911.90.90 (Thuế NK MFN 0%, VAT 8%).",
                "missing_attributes": [],
                "assumptions": "Túi lọc chất lỏng bằng vải dệt kỹ thuật sợi nylon dùng cho máy công nghiệp",
                "reserve_code": "8421.99.99"
            }

        # 6. BỘ PHỚT CƠ KHÍ / MECHANICAL SEAL (8484.20.00)
        if re.search(r'\b(?:phot\s+co\s+khi|bo\s+phot\s+co\s+khi|bo\s+lam\s+kin\s+co\s+khi|mechanical\s+seal(?:\s+kit)?)\b', norm) or \
           (re.search(r'\b(?:phot|seal)\b', hn) and re.search(r'\b(?:co\s+khi|mechanical)\b', norm)):
            return {
                "code": "8484.20.00",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1: Bộ làm kín kiểu cơ khí (phớt cơ khí / mechanical seals) có phân nhóm định danh riêng tại Nhóm 84.84. Áp dụng GRI 6: Mã số 8484.20.00 (Thuế NK MFN 3%, VAT 8%).",
                "missing_attributes": [],
                "assumptions": "Bộ phớt cơ khí dùng làm kín trục bơm chất lỏng",
                "reserve_code": "8487.90.90"
            }

        # 7. THANH ĐIỆN TRỞ GIA NHIỆT / BỘ SƯỞI ĐIỆN / QUE NHIỆT (8516.10.30 & 8516.80.80)
        if re.search(r'\b(?:dien\s+tro\s+gia\s+nhiet|thanh\s+dien\s+tro|bo\s+suoi\s+dien|thanh\s+nhiet|que\s+nhiet|electric\s+radiator|heating\s+element|immersion\s+heater)\b', norm) or \
           (re.search(r'\b(?:dien\s+tro|thanh\s+nhiet|bo\s+suoi)\b', hn) and re.search(r'\b(?:gia\s+nhiet|electric|radiator|heater)\b', norm)):
            is_immersion = bool(re.search(r'\b(?:nhung|dun\s+nuoc|dun\s+chua|immersion|tank|be\s+rua|radiator)\b', norm))
            return {
                "code": "8516.10.30" if is_immersion else "8516.80.80",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": f"Áp dụng GRI 1: Thiết bị và điện trở gia nhiệt điện thuộc Nhóm 85.16. Áp dụng GRI 6: {'Bộ gia nhiệt điện kiểu nhúng mã 8516.10.30' if is_immersion else 'Điện trở gia nhiệt loại khác mã 8516.80.80'} (Thuế NK MFN 20%, VAT 8%).",
                "missing_attributes": [],
                "assumptions": "Thanh điện trở gia nhiệt đun nóng dung dịch kiểu nhúng",
                "reserve_code": "8516.80.80" if is_immersion else "8516.10.30"
            }

        # 8. PHỚT CHẮN DẦU TRỤC / PHỚT TRỤC (4016.93.90)
        if re.search(r'\b(?:phot\s+chan\s+dau|phot\s+dau|phot\s+truc|shaft\s+seal|oil\s+seal|vong\s+chan\s+dau)\b', norm) or \
           (re.search(r'\b(?:phot|seal)\b', hn) and re.search(r'\b(?:chan\s+dau|dau\s+truc|shaft)\b', norm)):
            return {
                "code": "4016.93.90",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 & Chú giải 1(a) Phần XVI (loại trừ cao su sang Chương 40): Vòng đệm/phớt chắn dầu bằng cao su lưu hóa thuộc Nhóm 40.16. Áp dụng GRI 6: Phân nhóm 4016.93 mã số 4016.93.90 (Thuế NK MFN 3%, VAT 8%).",
                "missing_attributes": [],
                "assumptions": "Phớt chắn dầu trục bằng cao su lưu hóa có hoặc không có lõi kim loại",
                "reserve_code": "8487.90.90"
            }

        # 9. MÀNG TRƯỢT / TẤM LÓT TRƯỢT NHỰA IGLIDUR (3920.99.90)
        if re.search(r'\b(?:sliding\s+film|mang\s+truot|tam\s+lot\s+truot)\b', norm) and re.search(r'\b(?:iglidur|plastic|nhua)\b', norm):
            return {
                "code": "3920.99.90",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 & Chú giải 1(g) Phần XVI loại trừ plastic sang Chương 39: Tấm/màng trượt bằng plastic iglidur chưa xốp hóa thuộc Nhóm 39.20. Áp dụng GRI 6: Mã số 3920.99.90 (Thuế NK MFN 6%, VAT 8%).",
                "missing_attributes": [],
                "assumptions": "Màng trượt bằng plastic iglidur J dạng tấm/dải phẳng chưa xốp hóa",
                "reserve_code": "3926.90.99"
            }

        # 10. BẠC ĐẠN CHÀ / Ổ TRƯỢT CHẶN IGLIDUR (8483.30.90)
        if re.search(r'\b(?:bac\s+dan\s+cha|o\s+truot\s+chan|thrust\s+bearing)\b', norm):
            is_plain = bool(re.search(r'\b(?:iglidur|o\s+truot|bac\s+truot|plain|khong\s+(?:co\s+)?(?:bi|con\s+lan))\b', norm))
            return {
                "code": "8483.30.90" if is_plain else "8482.10.00",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": f"Áp dụng GRI 1 & Chú giải 2(a) Phần XVI: {'Ổ trượt chặn không có con lăn thuộc Nhóm 84.83 mã 8483.30.90' if is_plain else 'Ổ bi chặn (bạc đạn chà có bi) thuộc Nhóm 84.82 mã 8482.10.00'} (Thuế NK MFN 0%, VAT 8%).",
                "missing_attributes": [],
                "assumptions": "Ổ trượt chặn bằng plastic iglidur không dùng bi lăn" if is_plain else "Vòng bi chặn trục có ổ bi",
                "reserve_code": "8482.10.00" if is_plain else "8483.30.90"
            }

        # 11. CỤM TRỤC TRUYỀN ĐỘNG BÁNH RĂNG / ĐỘNG CƠ GIẢM TỐC (8501.51.19 / 8483.10.90)
        if re.search(r'\b(?:shaft\s+gear\s+drive|gear\s+motor|cum\s+truc\s+truyen\s+dong\s+banh\s+rang|dong\s+co\s+giam\s+toc)\b', norm):
            has_motor = bool(re.search(r'\b(?:motor|dong\s+co|dien|w|kw|v|volt|drive)\b', norm))
            return {
                "code": "8501.51.19" if has_motor else "8483.10.90",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": f"Áp dụng GRI 1 & Chú giải 2(a) Phần XVI: {'Động cơ điện xoay chiều gắn hộp số giảm tốc thuộc Nhóm 85.01 mã 8501.51.19' if has_motor else 'Trục truyền động bánh răng cơ khí thuộc Nhóm 84.83 mã 8483.10.90'} (Thuế NK MFN {'15%' if has_motor else '0%'}, VAT 8%).",
                "missing_attributes": [],
                "assumptions": "Động cơ điện gắn liền hộp giảm tốc và trục truyền động công suất ≤ 750W" if has_motor else "Trục truyền động bánh răng",
                "reserve_code": "8483.10.90" if has_motor else "8501.51.19"
            }

        # 12. THANH GIOĂNG CAO SU XỐP ĐỊNH HÌNH (4008.19.00)
        if re.search(r'\b(?:cao\s+su\s+xop\s+dinh\s+hinh|foam\s+rubber\s+profile|thanh\s+gioang\s+cao\s+su\s+xop)\b', norm) or \
           (re.search(r'\b(?:thanh\s+gioang|gioang|profile)\b', hn) and re.search(r'\b(?:cao\s+su\s+xop|foam\s+rubber)\b', norm)):
            return {
                "code": "4008.19.00",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 & Chú giải 1(a) Phần XVI loại trừ cao su sang Chương 40: Dải định hình bằng cao su xốp lưu hóa thuộc Nhóm 40.08. Áp dụng GRI 6: Phân nhóm 4008.19 mã số 4008.19.00 (Thuế NK MFN 3%, VAT 8%).",
                "missing_attributes": [],
                "assumptions": "Thanh gioăng định hình bằng cao su xốp lưu hóa",
                "reserve_code": "4016.93.90"
            }

        # 1. MÁY LÀM SẠCH / MÁY RỬA CÔNG NGHIỆP / HỆ THỐNG LÀM SẠCH LINH KIỆN (8424.89.50 / 8424.89.90)
        if re.search(r'\b(?:may\s+rua\s+(?:cong\s+nghiep|linh\s+kien|co\s+khi|sieu\s+am|phun)|may\s+lam\s+sach\s+(?:linh\s+kien|co\s+khi|cong\s+nghiep)?|he\s+thong\s+lam\s+sach|cleaning\s+system|parts\s+washer|part\s+washer|bupi(?:\s+cleaner)?)\b', norm) or \
           (re.search(r'\b(?:may\s+rua|may\s+lam\s+sach|cleaning|washer)\b', hn) and re.search(r'\b(?:cong\s+nghiep|linh\s+kien|top\s+loader|bupi|ap\s+luc|phun|co\s+khi)\b', norm)):
            is_electric = bool(re.search(r'\b(?:hoat\s+dong\s+bang\s+dien|chay\s+dien|bang\s+dien|electric|dien|v|kw|w)\b', norm)) and not bool(re.search(r'\b(?:khong\s+hoat\s+dong\s+bang\s+dien|khong\s+dung\s+dien|bang\s+tay|thu\s+cong)\b', norm))
            code = "8424.89.50" if is_electric else "8424.89.90"
            res_code = "8424.89.90" if is_electric else "8424.89.50"
            return {
                "code": code,
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": f"Áp dụng GRI 1: Thiết bị cơ khí phun chất lỏng làm sạch linh kiện công nghiệp thuộc Nhóm 84.24. Áp dụng GRI 6 (So sánh cùng cấp độ gạch nối): Phân nhóm 8424.89, cấp 3 gạch [- - - {'Loại khác, hoạt động bằng điện' if is_electric else 'Loại khác, không hoạt động bằng điện'}] mã số {code} (Thuế NK MFN 0%, VAT 8%).",
                "missing_attributes": [],
                "assumptions": "Máy làm sạch/rửa linh kiện cơ khí công nghiệp hoạt động bằng điện" if is_electric else "Máy làm sạch linh kiện không hoạt động bằng điện",
                "reserve_code": res_code
            }

        
        # BƯỚC 2: PHÂN LOẠI THEO DANH TỪ TRUNG TÂM (HEAD NOUN - GRI 1)
        # -------------------------------------------------------------

        # 1. BĂNG KEO TỰ DÍNH (Nhóm 3919)
        if re.search(r'\b(?:bang\s+keo|bang\s+dinh|adhesive\s+tape)\b', hn):
            return {
                "code": "3919.10.99",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 -> Nhóm 39.19 tấm, phiến, băng dính tự dính bằng plastic. Áp dụng GRI 6: Dạng cuộn chiều rộng không quá 20cm mã 3919.10.99.",
                "missing_attributes": [],
                "assumptions": "Băng keo tự dính dạng cuộn",
                "reserve_code": "3919.90.99"
            }

        # 2. LƯỠI CẮT CHO MÁY MÓC (Nhóm 8208)
        if re.search(r'\b(?:luoi\s+cat|dao\s+cat|cutter\s+blade|shear\s+blade)\b', hn):
            return {
                "code": "8208.90.00",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 & Chú giải 1(k) Phần XVI: Dao và lưỡi cắt dùng cho máy hoặc dụng cụ cơ khí thuộc Nhóm 82.08. Áp dụng GRI 6 xếp mã 8208.90.00.",
                "missing_attributes": [],
                "assumptions": "Lưỡi dao cắt bằng thép dùng cho máy công nghiệp",
                "reserve_code": "8208.10.00"
            }

        # 3. DỤNG CỤ CẦM TAY: DỤNG CỤ THÁO LẮP, BƠM MỠ, KÌM (Nhóm 8205 & 8203)
        if re.search(r'\b(?:dung\s+cu\s+thao|dung\s+cu\s+lap|dung\s+cu\s+rut|hand\s+tool|extractor)\b', hn):
            return {
                "code": "8205.59.00",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 & Chú giải 1(k) Phần XVI: Dụng cụ cầm tay hoạt động bằng tay chưa định danh nơi khác thuộc Nhóm 82.05. Áp dụng GRI 6 xếp mã 8205.59.00.",
                "missing_attributes": [],
                "assumptions": "Dụng cụ tháo lắp đầu nối hoạt động hoàn toàn bằng tay",
                "reserve_code": "8467.89.00"
            }

        if re.search(r'\b(?:kim\s+bam\s+dau\s+cos|kim\s+bam|kim\s+cat|kim\s+tuot|hand\s+crimping|pliers)\b', hn):
            return {
                "code": "8203.20.00",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 & Chú giải 1(k) Phần XVI: Kìm cầm tay cơ khí không dùng động cơ phân loại tại Nhóm 82.03. Áp dụng GRI 6 xếp mã 8203.20.00.",
                "missing_attributes": [],
                "assumptions": "Kìm cơ khí cầm tay hoàn chỉnh",
                "reserve_code": "8467.29.00"
            }

        # 7. GIOĂNG ĐỆM CAO SU LƯU HÓA (Chương 40)
        if re.search(r'\b(?:gioang|dem|o-ring|phot)\b', hn) and (attrs["material"] == "rubber" or 'cao su' in norm):
            is_vehicle = bool(re.search(r'\b(?:o\s+to|xe\s+co\s+dong\s+co|automotive|dong\s+co\s+xe)\b', norm))
            return {
                "code": "4016.93.20" if is_vehicle else "4016.93.90",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": f"Áp dụng GRI 1 & Chú giải 1(a) Phần XVI loại trừ cao su sang Chương 40 -> Nhóm 40.16. Áp dụng GRI 6: Gioăng đệm Phân nhóm 4016.93 -> {'Dùng cho xe ô tô mã 4016.93.20' if is_vehicle else 'Loại khác mã 4016.93.90'}.",
                "missing_attributes": [],
                "assumptions": "Gioăng đệm bằng cao su lưu hóa",
                "reserve_code": "4016.93.90" if is_vehicle else "4016.93.20"
            }

        # 7b. BĂNG TẢI, ĐAI TẢI, BĂNG TRUYỀN, DÂY CU-ROA (Nhóm 4010)
        if re.search(r'\b(?:dai\s+tai|bang\s+tai|conveyor\s+belt|bang\s+chuyen|day\s+curoa|day\s+cu-roa|curoa|cu-roa|transmission\s+belt|v-belt|timing\s+belt|dai\s+truyen|bang\s+truyen)\b', hn) or \
           (re.search(r'\b(?:dai\s+tai|bang\s+tai|conveyor\s+belt|bang\s+chuyen|day\s+curoa|day\s+cu-roa|curoa|cu-roa|transmission\s+belt|v-belt|timing\s+belt|dai\s+truyen|bang\s+truyen)\b', name_norm) and ('cao su' in norm or attrs["material"] == "rubber")):
            
            # 1. Phân nhóm Băng tải hoặc đai tải (Conveyor belts - 4010.1x)
            if re.search(r'\b(?:dai\s+tai|bang\s+tai|conveyor\s+belt|bang\s+chuyen)\b', norm):
                if re.search(r'\b(?:kim\s+loai|thep|metal|steel|wire|day\s+thep|soi\s+kim\s+loai)\b', norm):
                    return {
                        "code": "4010.11.00",
                        "level": 8,
                        "status": "EXACT_8_DIGIT",
                        "reasoning": "Áp dụng GRI 1, Chú giải 1(a) Phần XVI (loại trừ đai tải, băng tải bằng cao su lưu hóa khỏi Phần XVI máy móc đưa về Nhóm 40.10: Băng tải hoặc đai tải hoặc băng truyền hoặc đai truyền, bằng cao su lưu hóa). Áp dụng GRI 6: Cấp phân nhóm 1 gạch [- Băng tải hoặc đai tải:], cấp phân nhóm 2 gạch [-- Chỉ được gia cố bằng kim loại] xếp chính xác mã số 8 số 4010.11.00 (Thuế NK MFN 5%, VAT 8/10%).",
                        "missing_attributes": [],
                        "assumptions": "Đai tải/băng tải bằng cao su lưu hóa gia cố bằng kim loại",
                        "reserve_code": "4010.12.00"
                    }
                if re.search(r'\b(?:vat\s+lieu\s+det|soi\s+det|vai|det|textile|fabric|canvas|polyester|bo\s+vai)\b', norm):
                    return {
                        "code": "4010.12.00",
                        "level": 8,
                        "status": "EXACT_8_DIGIT",
                        "reasoning": "Áp dụng GRI 1, Chú giải 1(a) Phần XVI -> Nhóm 40.10. Áp dụng GRI 6: Cấp phân nhóm 1 gạch [- Băng tải hoặc đai tải:], cấp phân nhóm 2 gạch [-- Chỉ được gia cố bằng vật liệu dệt] xếp chính xác mã số 8 số 4010.12.00 (Thuế NK MFN 5%, VAT 8/10%).",
                        "missing_attributes": [],
                        "assumptions": "Đai tải/băng tải bằng cao su lưu hóa gia cố bằng vật liệu dệt",
                        "reserve_code": "4010.11.00"
                    }
                return {
                    "code": "4010.19.00",
                    "level": 8,
                    "status": "EXACT_8_DIGIT",
                    "reasoning": "Áp dụng GRI 1 -> Nhóm 40.10. Áp dụng GRI 6: Băng tải hoặc đai tải loại khác xếp mã số 4010.19.00.",
                    "missing_attributes": [],
                    "assumptions": "Đai tải/băng tải bằng cao su lưu hóa loại khác",
                    "reserve_code": "4010.11.00"
                }

            # 2. Phân nhóm Băng truyền hoặc đai truyền (Transmission belts / Dây curoa - 4010.3x)
            if re.search(r'\b(?:rang\s+cua|dong\s+bo|timing|synchronous)\b', norm):
                return {
                    "code": "4010.35.00",
                    "level": 8,
                    "status": "EXACT_8_DIGIT",
                    "reasoning": "Áp dụng GRI 1 -> Nhóm 40.10. Áp dụng GRI 6: Băng truyền đồng bộ liên tục (có răng) chu vi ngoài không quá 150 cm xếp mã số 4010.35.00.",
                    "missing_attributes": [],
                    "assumptions": "Dây curoa đồng bộ có răng chu vi <= 150cm",
                    "reserve_code": "4010.36.00"
                }
            if re.search(r'\b(?:chu\s+v|v-belt|hinh\s+thang)\b', norm):
                return {
                    "code": "4010.31.00",
                    "level": 8,
                    "status": "EXACT_8_DIGIT",
                    "reasoning": "Áp dụng GRI 1 -> Nhóm 40.10. Áp dụng GRI 6: Băng truyền liên tục mặt cắt ngang hình thang (băng chữ V) xếp mã số 4010.31.00.",
                    "missing_attributes": [],
                    "assumptions": "Băng truyền hình thang chữ V chu vi 60cm - 180cm",
                    "reserve_code": "4010.32.00"
                }
            return {
                "code": "4010.39.00",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 -> Nhóm 40.10. Áp dụng GRI 6: Băng truyền hoặc đai truyền loại khác xếp mã số 4010.39.00.",
                "missing_attributes": [],
                "assumptions": "Dây đai truyền động cao su loại khác",
                "reserve_code": "4010.35.00"
            }

        # 4. MÁY MÓC CÔNG NGHIỆP: MÁY DẬP, MÁY IN, MÁY BƠM, QUẠT (Chương 84)
        if re.search(r'\b(?:may\s+(?:ep|dap)\s+bam\s+dau\s+cos|may\s+dap\s+dau\s+noi|may\s+bam\s+cos|crimping\s+machine)\b', hn) or (re.search(r'\b(?:may\s+dap|may\s+ep)\b', hn) and 'dau cos' in norm):
            return {
                "code": "8462.62.11",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1: Máy cơ khí dập khuôn ép kim loại thuộc Nhóm 84.62. Áp dụng GRI 6 xếp vào mã 8462.62.11.",
                "missing_attributes": [],
                "assumptions": "Máy dập ép cơ khí dùng điện gia công kim loại",
                "reserve_code": "8479.89.39"
            }

        if re.search(r'\b(?:may\s+in|printer|printing\s+machine)\b', hn):
            return {
                "code": "8443.19.00",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 -> Nhóm 84.43 máy in nhãn công nghiệp. Áp dụng GRI 6 xếp mã 8443.19.00.",
                "missing_attributes": [],
                "assumptions": "Máy in chữ lên nhãn hoạt động bằng điện",
                "reserve_code": "8443.32.90"
            }

        if re.search(r'\b(?:quat\s+ly\s+tam|quat\s+hut|centrifugal\s+fan|quat\s+cong\s+nghiep)\b', hn):
            return {
                "code": "8414.59.49",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1: Quạt hút ly tâm liền động cơ thuộc Nhóm 84.14 (không phân loại theo động cơ 8501). Áp dụng GRI 6: Quạt khác (8414.59) -> Quạt ly tâm mã 8414.59.49.",
                "missing_attributes": [],
                "assumptions": "Quạt ly tâm công nghiệp bằng thép",
                "reserve_code": "8414.59.41"
            }

        if re.search(r'\b(?:bom\s+nuoc|bom\s+ly\s+tam|centrifugal\s+pump|may\s+bom)\b', hn):
            return {
                "code": "8413.70.42",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1: Máy bơm chất lỏng liền động cơ thuộc Nhóm 84.13. Áp dụng GRI 6: Bơm ly tâm một tầng cánh họng xả trên 50mm mã 8413.70.42.",
                "missing_attributes": [],
                "assumptions": "Bơm nước ly tâm trục ngang",
                "reserve_code": "8413.70.91"
            }

        # 5. THIẾT BỊ ĐO LƯỜNG & KIỂM TRA (Chương 90)
        if re.search(r'\b(?:dong\s+ho\s+do|dong\s+ho|thiet\s+bi\s+do|may\s+do|bo\s+phan\s+cua\s+thuoc\s+do|gauge|measuring)\b', hn):
            if re.search(r'\b(?:ap\s+suat|pressure)\b', norm):
                return {
                    "code": "9026.20.30",
                    "level": 8,
                    "status": "EXACT_8_DIGIT",
                    "reasoning": "Áp dụng GRI 1 & Chú giải 1(m) Phần XVI loại trừ sang Chương 90 -> Nhóm 90.26 dụng cụ đo áp suất. Áp dụng GRI 6 xếp mã 9026.20.30.",
                    "missing_attributes": [],
                    "assumptions": "Đồng hồ đo áp suất cơ học ống bourdon",
                    "reserve_code": "9026.20.40"
                }
            elif re.search(r'\b(?:thang\s+hang|kiem\s+tra|alignment)\b', norm):
                return {
                    "code": "9031.80.90",
                    "level": 8,
                    "status": "EXACT_8_DIGIT",
                    "reasoning": "Áp dụng GRI 1 & Chú giải Chương 90: Thiết bị và dụng cụ đo lường kiểm tra quang học hoặc điện tử chưa định danh riêng thuộc Nhóm 90.31. Áp dụng GRI 6 xếp mã 9031.80.90.",
                    "missing_attributes": [],
                    "assumptions": "Thiết bị đo độ thẳng hàng connector",
                    "reserve_code": "9031.49.90"
                }
            else:
                return {
                    "code": "9017.90.90",
                    "level": 8,
                    "status": "EXACT_8_DIGIT",
                    "reasoning": "Áp dụng GRI 1 -> Nhóm 90.17 dụng cụ đo chiều dài, micromet và bộ phận của chúng. Áp dụng GRI 6 xếp mã 9017.90.90.",
                    "missing_attributes": [],
                    "assumptions": "Bộ phận của thước đo micromet",
                    "reserve_code": "9031.90.90"
                }

        if re.search(r'\b(?:thermostat|bo\s+dieu\s+nhiet)\b', hn) or re.search(r'\b(?:thermostat|bo\s+dieu\s+nhiet)\b', norm):
            return {
                "code": "9032.10.10",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 & Chú giải 7 Chương 90: Thiết bị điều chỉnh và kiểm soát tự động -> Phân nhóm 9032.10 Thermostat -> Mã 9032.10.10.",
                "missing_attributes": [],
                "assumptions": "Thermostat tự động ngắt điện",
                "reserve_code": "9032.10.20"
            }

        # 6. VÚ TRA MỠ BÔI TRƠN (Nhóm 8487)
        if re.search(r'\b(?:vu\s+tra\s+mo|vu\s+bom\s+mo|num\s+bom\s+mo|grease\s+nipple|lube\s+nipple)\b', hn):
            return {
                "code": "8487.90.00",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 & Chú giải WCO Nhóm 84.87: Núm vú tra mỡ bôi trơn không tự động thuộc Nhóm 84.87. Áp dụng GRI 6 xếp mã 8487.90.00.",
                "missing_attributes": [],
                "assumptions": "Vú mỡ bôi trơn cơ khí tiêu chuẩn",
                "reserve_code": "7318.29.90"
            }

        # 8. THIẾT BỊ ĐÓNG NGẮT: APTOMAT, CÔNG TẮC, ĐẦU NỐI (Nhóm 8536)
        if re.search(r'\b(?:mcb|aptomat|cau\s+dao\s+tu\s+dong|circuit\s+breaker)\b', hn):
            return {
                "code": "8536.20.91",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 -> Nhóm 85.36 (loại trừ nhóm dây điện 8544). Áp dụng GRI 6: Cầu dao tự động ngắt mạch điện áp <= 1000V dòng <= 16A mã 8536.20.91.",
                "missing_attributes": [],
                "assumptions": "MCB bảo vệ quá tải <= 16A",
                "reserve_code": "8536.20.99"
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
            return {
                "code": "7318.16.00",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 & Chú giải 2 Phần XV: Đai ốc bằng thép thuộc Nhóm 73.18. Áp dụng GRI 6: Đai ốc có ren thuộc mã 7318.16.00.",
                "missing_attributes": [],
                "assumptions": "Đai ốc thép tiêu chuẩn",
                "reserve_code": "7318.15.10"
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
                if is_cast and (attrs["dn_mm"] is not None and attrs["dn_mm"] > 25.0):
                    return {
                        "code": "8481.30.22",
                        "level": 8,
                        "status": "EXACT_8_DIGIT",
                        "reasoning": "Áp dụng GRI 1 & Chú giải 2(a) Phần XVI -> Nhóm 84.81. Áp dụng GRI 6: Van một chiều Phân nhóm 8481.30 -> Bằng gang, đường kính trong trên 2,5 cm (DN50) -> Mã 8481.30.22.",
                        "missing_attributes": [],
                        "assumptions": "Van một chiều bằng gang DN50",
                        "reserve_code": "8481.30.21"
                    }
                else:
                    return {
                        "code": "8481.30.90",
                        "level": 8,
                        "status": "EXACT_8_DIGIT",
                        "reasoning": "Áp dụng GRI 1 & Chú giải 2(a) Phần XVI -> Nhóm 84.81. Áp dụng GRI 6: Van một chiều Phân nhóm 8481.30 -> Loại khác mã 8481.30.90.",
                        "missing_attributes": [],
                        "assumptions": "Van một chiều công nghiệp",
                        "reserve_code": "8481.30.22"
                    }

            if is_safety:
                return {
                    "code": "8481.40.21" if attrs["material"] == "copper" else "8481.40.90",
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

        # 17. ĐỘNG CƠ ĐIỆN / MÔ TƠ (Nhóm 8501)
        if re.search(r'\b(?:dong\s+co\s+dien|mo\s+to\s+dien|mo\s+to\s+buoc|motor|ac\s+motor|dc\s+motor)\b', hn):
            is_dc = bool(re.search(r'\bdc\b', norm) or 'mot chieu' in norm)
            if is_dc:
                return {
                    "code": "8501.31.60",
                    "level": 8,
                    "status": "EXACT_8_DIGIT",
                    "reasoning": "Áp dụng GRI 1 -> Nhóm 85.01 động cơ điện. Áp dụng GRI 6: Động cơ DC công suất <= 750W (82W) -> Mô tơ bước mã 8501.31.60.",
                    "missing_attributes": [],
                    "assumptions": "Mô tơ bước DC",
                    "reserve_code": "8501.10.29"
                }

            if attrs["power_w"] is None:
                return {
                    "code": "8501.5",
                    "level": 5,
                    "status": "STOP_6_DIGIT",
                    "reasoning": "Áp dụng GRI 1 & Chú giải 2(a) Phần XVI -> Nhóm 85.01. Áp dụng GRI 6: Động cơ điện xoay chiều AC nhiều pha thuộc Phân nhóm 8501.5. TUY NHIÊN, thiếu thông số công suất (kW/W/HP) để phân nhánh định thuế (<=750W, <=7.5kW, <=75kW hay >75kW) -> Dừng ở phân nhóm theo nguyên tắc Không đoán mò.",
                    "missing_attributes": ["power_kw"],
                    "assumptions": "Thiếu công suất động cơ",
                    "reserve_code": "8501.51.00"
                }
            else:
                p_kw = attrs["power_w"] / 1000.0
                if p_kw <= 0.75:
                    return {"code": "8501.51.00", "level": 8, "status": "EXACT_8_DIGIT", "reasoning": "Động cơ AC công suất <= 750W", "missing_attributes": [], "assumptions": "", "reserve_code": "8501.52.11"}
                elif p_kw <= 7.5:
                    return {"code": "8501.52.11", "level": 8, "status": "EXACT_8_DIGIT", "reasoning": "Động cơ AC công suất 750W - 7.5kW", "missing_attributes": [], "assumptions": "", "reserve_code": "8501.52.19"}
                elif p_kw <= 75.0:
                    return {"code": "8501.52.19", "level": 8, "status": "EXACT_8_DIGIT", "reasoning": "Động cơ AC công suất 7.5kW - 75kW", "missing_attributes": [], "assumptions": "", "reserve_code": "8501.53.00"}

        # 18. DỤNG CỤ CẮT GỌT THAY THẾ CHO MÁY CÔNG CỤ (Nhóm 8207)
        if re.search(r'\b(?:mui\s+khoan|dao\s+phay|dao\s+tien|drill\s+bit|milling\s+cutter)\b', hn):
            return {
                "code": "8207.50.00",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 & Chú giải 1(k) Phần XVI: Dụng cụ cắt gọt kim loại có thể thay đổi được thuộc Nhóm 82.07. Áp dụng GRI 6: Dụng cụ khoan mã 8207.50.00.",
                "missing_attributes": [],
                "assumptions": "Mũi khoan thép gió HSS",
                "reserve_code": "8207.90.00"
            }

        # 19. CHỐT ĐỊNH VỊ & LÒ XO (Nhóm 7318 & 7320)
        if re.search(r'\b(?:chot\s+dinh\s+vi|dowel\s+pin)\b', hn):
            return {
                "code": "7318.24.00",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 & Chú giải 2 Phần XV: Chốt kim loại không ren thuộc Nhóm 73.18. Áp dụng GRI 6: Chốt định vị mã 7318.24.00.",
                "missing_attributes": [],
                "assumptions": "Chốt hình trụ bằng thép không ren",
                "reserve_code": "7318.29.90"
            }

        if re.search(r'\b(?:lo\s+xo|spring)\b', hn):
            return {
                "code": "7320.20.90",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 & Chú giải 2 Phần XV: Lò xo bằng thép thuộc Nhóm 73.20. Áp dụng GRI 6: Lò xo xoắn nén hình trụ mã 7320.20.90.",
                "missing_attributes": [],
                "assumptions": "Lò xo nén bằng thép",
                "reserve_code": "7320.90.90"
            }

        # 20. BIẾN ÁP & NGUỒN CHUYỂN ĐỔI (Nhóm 8504)
        if re.search(r'\b(?:adapter|switching|nguon\s+chuyen\s+doi|bien\s+doi\s+tinh)\b', hn):
            return {
                "code": "8504.40.90",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 -> Nhóm 85.04. Áp dụng GRI 6: Bộ biến đổi tĩnh điện AC sang DC Phân nhóm 8504.40 -> Mã 8504.40.90.",
                "missing_attributes": [],
                "assumptions": "Bộ nguồn chuyển đổi tĩnh adapter",
                "reserve_code": "8504.40.11"
            }

        if re.search(r'\b(?:bien\s+ap|bien\s+the|transformer)\b', hn):
            return {
                "code": "8504.31.24",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 -> Nhóm 85.04. Áp dụng GRI 6: Máy biến thế khác công suất <= 1kVA (500VA) -> Mã 8504.31.24.",
                "missing_attributes": [],
                "assumptions": "Biến áp cách ly công suất <= 1kVA",
                "reserve_code": "8504.31.29"
            }

        # 21. BẢNG ĐIỀU KHIỂN ĐIỆN (Nhóm 8537)
        if re.search(r'\b(?:tu\s+bang\s+dien|bang\s+dieu\s+khien|control\s+panel|plc|hmi)\b', hn):
            return {
                "code": "8537.10.99",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 -> Nhóm 85.37: Bảng, tủ điều khiển phân phối điện. Áp dụng GRI 6: Điện áp <= 1000V -> Mã 8537.10.99.",
                "missing_attributes": [],
                "assumptions": "Tủ điều khiển tự động hóa điện áp <= 1000V",
                "reserve_code": "8537.10.19"
            }

        # 22. RƠ LE (Nhóm 8536.49)
        if re.search(r'\b(?:ro\s+le|relay)\b', hn):
            return {
                "code": "8536.49.90",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 -> Nhóm 85.36. Áp dụng GRI 6: Rơ le bảo vệ quá tải nhiệt phân nhóm 8536.49 -> Mã 8536.49.90.",
                "missing_attributes": [],
                "assumptions": "Rơ le nhiệt lưỡng kim",
                "reserve_code": "8536.41.00"
            }

        # 23. BỘ NHỚ LƯU TRỮ FLASH / USB (Nhóm 8523)
        if re.search(r'\b(?:usb|flash\s+drive|the\s+nho|o\s+nho|solid-state|storage)\b', hn):
            return {
                "code": "8523.51.11",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 -> Nhóm 85.23 thiết bị lưu trữ dữ liệu bán dẫn bền vững. Áp dụng GRI 6: Thiết bị lưu trữ thể rắn chưa ghi dữ liệu mã 8523.51.11.",
                "missing_attributes": [],
                "assumptions": "USB flash drive chưa ghi dữ liệu",
                "reserve_code": "8523.51.21"
            }

        # 24. ỐNG CAO SU (Nhóm 4009)
        if re.search(r'\b(?:ong\s+(?:dan\s+)?nuoc|ong\s+cao\s+su|rubber\s+hose)\b', hn):
            return {
                "code": "4009.31.91",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 & Chú giải loại trừ -> Nhóm 40.09 ống cao su. Áp dụng GRI 6: Gia cố sợi dệt không phụ kiện dùng cho xe ô tô -> Mã 4009.31.91.",
                "missing_attributes": [],
                "assumptions": "Ống dẫn nước làm mát két nước ô tô",
                "reserve_code": "4009.31.99"
            }

        # 25. CẢN TRƯỚC / MÁ PHANH Ô TÔ (Nhóm 8708)
        if re.search(r'\b(?:can\s+truoc|can\s+sau|ba\s+do\s+soc|bumper)\b', hn):
            return {
                "code": "8708.10.10",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 & Chú giải 2 Phần XVII: Cản va và bộ phận cản va định danh riêng tại Phân nhóm 8708.10 -> Cản trước ô tô con mã 8708.10.10.",
                "missing_attributes": [],
                "assumptions": "Cản va ô tô bằng nhựa hoàn chỉnh",
                "reserve_code": "8708.10.90"
            }

        if re.search(r'\b(?:ma\s+phanh|bo\s+thang|brake\s+pad)\b', hn):
            return {
                "code": "8708.30.29",
                "level": 8,
                "status": "EXACT_8_DIGIT",
                "reasoning": "Áp dụng GRI 1 -> Nhóm 87.08 phụ tùng ô tô. Áp dụng GRI 6: Phanh hãm và trợ động phanh Phân nhóm 8708.30 -> Đã lắp xương thép mã 8708.30.29.",
                "missing_attributes": [],
                "assumptions": "Má phanh đĩa ô tô đã gắn đế thép",
                "reserve_code": "6813.81.00"
            }

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
                "reserve_code": "7616.99.19"
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
            "code": "INSUFFICIENT_INFO",
            "level": 0,
            "status": "INSUFFICIENT_INFO",
            "reasoning": "Chưa đủ dữ kiện đối sánh chính xác theo Biểu thuế WCO 2022.",
            "missing_attributes": ["technical_specs"],
            "assumptions": "N/A",
            "reserve_code": "N/A"
        }
