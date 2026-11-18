# -*- coding: utf-8 -*-
"""Golden test cho Universal GRI-6 Engine (muc 2.1 / 2.2 / 2.3)."""
import sys, json
sys.stdout.reconfigure(encoding='utf-8')
import universal_engine as u

# (mo ta, query, ma phai SONG SOT, ma phai BI LOAI)
ENGINE_CASES = [
    ('Cong tac 220V 10A',            'Công tắc chuyển mạch điện 220V 10A',                     ['85365096'], ['85365099', '85351000']),
    ('Cong tac 25A',                 'Công tắc điện công nghiệp 380V 25A',                      ['85365099'], ['85365096']),
    ('Bien 15,9A',                   'Công tắc điện 220V 15.9A',                                ['85365096'], ['85365099']),
    ('Bien 16A (khong duoi 16)',     'Công tắc điện 220V 16A',                                  ['85365099'], ['85365096']),
    ('Bien 16,1A',                   'Công tắc điện 220V 16.1A',                                ['85365099'], ['85365096']),
    ('Nguoi dung viet "duoi 16A"',   'Công tắc điện 220V, dòng điện dưới 16A',                  ['85365096'], ['85365099']),
    ('Nguoi dung viet "tren 16A"',   'Công tắc điện 220V, dòng điện trên 16A',                  ['85365099'], ['85365096']),
    ('Ha the 1000V -> 8536',         'Cầu dao đóng ngắt 1000V',                                 ['85365099'], ['85352190', '85351000']),
    ('Trung the 1001V -> 8535',      'Cầu dao cách ly 1001V',                                   ['85352190'], ['85365099']),
    ('Cau dao 24kV 630A -> 8535',    'Cầu dao cách ly trung thế 24kV 630A',                     ['85353019'], ['85365099', '85365096', '85353020']),
    ('Hoat dong bang dien',          'Máy làm sạch linh kiện hoạt động bằng điện',               ['84248950'], ['84248990']),
    ('Khong hoat dong bang dien',    'Thiết bị phun rửa không hoạt động bằng điện, thao tác thủ công', ['84248990'], ['84248950']),
    ('Cap mang Cat6 UTP (vien thong, <=80V)', 'Cáp mạng Cat6 UTP 4 đôi vỏ PVC chưa gắn đầu nối, điện áp tín hiệu 5V, dùng cho mạng LAN', ['85444913'], ['85444921', '85444922', '85444923', '85444924', '85444929']),
    ('Cap vien thong 220V (>80V)', 'Cáp viễn thông điện thoại vỏ plastic điện áp 220V', ['85444932'], ['85444913', '85444922']),
    ('Day o to 12V khong vien thong', 'Dây đơn có vỏ bọc chống nhiễu dùng sản xuất bó dây ô tô, 12V, chưa gắn đầu nối', ['85444921'], ['85444932']),
    ('Bang tai gia co kim loai', 'Đai tải băng tải cao su gia cố bằng lõi thép', ['40101100'], ['40101200', '40101900']),
    ('Bang tai gia co vat lieu det', 'Băng tải cao su gia cố bằng sợi vải dệt', ['40101200'], ['40101100', '40101900']),
    ('Co o bi',                      'Thanh trượt thép có ổ bi',                                 ['84821000'], []),
]

# Ca cuoi cung: chay qua toan bo pipeline server (Python) -> ma 8 so
PIPELINE_CASES = [
    ('Cong tac 10A',   {'name': 'Công tắc chuyển mạch điện 220V 10A', 'desc': 'Vỏ nhựa PC/ABS, tiếp điểm đồng', 'use': 'Đóng ngắt mạch điện hạ thế', 'specs': 'Điện áp 220V, dòng điện định mức 10A'}, '85365096'),
    ('Cong tac 25A',   {'name': 'Công tắc điện công nghiệp 380V 25A', 'desc': 'Vỏ nhựa', 'use': 'Đóng ngắt mạch điện', 'specs': 'Điện áp 380V, dòng điện 25A'}, '85365099'),
    ('May rua dien',   {'name': 'Máy làm sạch linh kiện cơ khí (máy rửa công nghiệp) hoạt động bằng điện', 'desc': 'Sử dụng thiết bị bơm ly tâm với cánh bơm dạng mở', 'use': 'dùng để rửa linh kiện cơ khí', 'specs': ''}, '84248950'),
    ('Cau dao 24kV',   {'name': 'Cầu dao cách ly trung thế 24kV 630A', 'desc': 'Cầu dao đóng ngắt mạch điện cao thế ngoài trời', 'use': 'Đóng cắt và bảo vệ lưới điện truyền tải', 'specs': 'Điện áp định mức 24kV, dòng định mức 630A'}, '853530'),
]


# 13 hang da duoc chuyen gia tham dinh (nguon: KQ_HS_CODE_PHOTO_BUPI_TL80.xlsx / test_final_py.py)
EXPERT = [
    ('Cleaning system Top loader TL 80 MV2A (BUPI-GOLSER, Austria, 3x400V, 6,8 kW, 2026)', 'Hệ thống làm sạch (máy rửa công nghiệp) nạp liệu từ trên xuống, có giỏ quay, công suất 6.8 kW', '84248950'),
    ('FZGMG 72013 Shaft gear drive TL complete', 'Cụm trục truyền động bánh răng hoàn chỉnh cho dòng TL, động cơ giảm tốc', '85015119'),
    ('FZGMK 59553 Gas pressure shock absorber 250N', 'Giảm chấn khí nén (phuộc hơi) lực 250N, mã FZGMK 59553', '84123100'),
    ('FZGMK 59552 Gas pressure shock absorber 500N', 'Giảm chấn khí nén (phuộc hơi) lực 500N dùng cho máy TL80', '84123100'),
    ('FZGEL 85348 Liquiphant FTL31-13E0/0 24V DC', 'Cảm biến báo mức chất lỏng (công tắc mức dạng rung) Liquiphant FTL31', '90261090'),
    ('FZGMK 85149 Angle Seat Valve (Schubert & Salzer)', 'Van xiên (van góc) DN015-1/2 inch kèm van điều khiển 1/8 inch', '84818082'),
    ('FZEZB 76488 Filter bag NMO 50 P04S 50S', 'Túi lọc chất lỏng (lưới nylon NMO 50 micron), model P04S 50S', '59119090'),
    ('FZGDT 80230 Mechanical Seal Kit (Grundfos)', 'Bộ phớt cơ khí và bộ gioăng làm kín dùng cho bơm Grundfos', '84842000'),
    ('FZGEL 52424 Electric Radiator', 'Thanh điện trở gia nhiệt (bộ sưởi điện) dùng cho máy TL80', '85161030'),
    ('HZGGK 80687 Foam rubber profile 12 x 12 mm', 'Thanh gioăng cao su xốp định hình kích thước 12 x 12 mm', '40081900'),
    ('FZGLG 72947 Thrust bearing TL80/100 (iglidur)', 'Bạc đạn chà (vòng bi chặn trục / ổ trượt chặn) dùng cho máy TL80/100', '84833090'),
    ('FZGLG 74907 iglidur J sliding film', 'Tấm lót trượt (bạc lót / màng trượt) vật liệu nhựa Iglidur J, dài 57mm', '84833090'),
    ('FZGLG 72794 Shaft seal WDR 25/35/7 SL', 'Phớt chắn dầu trục (phớt trục) kích thước 25/35/7 mm', '84879000'),
]
PIPELINE_CASES += [(f'[CHUYEN GIA] {n[:40]}', {'name': n, 'desc': d, 'use': '', 'specs': ''}, c) for n, d, c in EXPERT]


PIPELINE_CASES += [
    ('Bang tai kim loai', {'name': 'Băng tải (đai tải) cao su gia cố bằng lõi thép', 'desc': 'Cao su lưu hóa, lõi thép', 'use': 'Dùng cho băng chuyền công nghiệp', 'specs': ''}, '40101100'),
    ('Bang tai vai det', {'name': 'Băng tải (đai tải) cao su gia cố bằng vải dệt', 'desc': 'Cao su lưu hóa, lớp vải dệt', 'use': 'Dùng cho băng chuyền công nghiệp', 'specs': ''}, '40101200'),
]

# Bo sung 18/11: phan dinh theo dac tinh (Chu giai chi tiet 84.87, 84.83, Bieu thue 2026)
PIPELINE_CASES += [
    ('Phot WDR web (cot kim loai)', {'name': 'Vòng làm kín dầu kiểu WDR SL', 'desc': 'bằng cao su lưu hóa liền cốt kim loại, có môi chắn bụi, kích thước 25x35x7 mm', 'use': 'phụ tùng máy rửa công nghiệp BUPI CLEANER', 'specs': ''}, '84879000'),
    ('Phot toan cao su', {'name': 'Phớt chắn dầu trục hoàn toàn bằng cao su', 'desc': 'Cao su NBR, không có cốt kim loại, kích thước 20x30x5 mm', 'use': 'Làm kín trục bơm', 'specs': ''}, '40169390'),
    ('Mang truot iglidur phang', {'name': 'Màng trượt iglidur J dạng tấm phẳng', 'desc': 'Tấm nhựa iglidur J dày 0,5 mm, khổ 500 x 1000 mm', 'use': 'Cắt làm bề mặt trượt', 'specs': ''}, '39209990'),
    ('Dien tro khong nhung', {'name': 'Điện trở gia nhiệt dạng ống dùng cho máy ép nhựa', 'desc': 'Heating element gắn khuôn', 'use': 'Gia nhiệt khuôn', 'specs': '220V 500W'}, '85168090'),
]

# Bo sung 18/11 (2): MCB theo dong dien, bien ap theo cong suat (dung 6 so khi thieu chung loai)
PIPELINE_CASES += [
    ('MCB 16A', {'name': 'Aptomat cầu dao tự động tép MCB 1 cực 16A', 'desc': 'Thiết bị đóng ngắt bảo vệ quá tải và ngắn mạch', 'use': 'Bảo vệ đường dây điện chiếu sáng', 'specs': 'Điện áp 230V, dòng định mức 16A, cắt 4.5kA'}, '85362099'),
    ('MCB 10A', {'name': 'Aptomat MCB 1 cực 10A', 'desc': 'Cầu dao tự động', 'use': 'Bảo vệ quá tải', 'specs': '230V, 10A'}, '85362091'),
    ('MCCB 100A', {'name': 'Aptomat MCCB loại hộp đúc 3 cực 100A', 'desc': 'Moulded case circuit breaker', 'use': 'Bảo vệ tủ điện', 'specs': '400V, 100A'}, '85362013'),
    ('Bien ap 500VA', {'name': 'Máy biến áp cách ly 1 pha hạ áp 220V xuống 110V', 'desc': 'Lõi thép silic quấn cuộn sơ cấp và thứ cấp cách điện', 'use': 'Đổi điện dùng cho thiết bị nội địa Nhật', 'specs': 'Công suất 500VA (0.5 kVA)'}, '850431'),
    ('Bien ap 10kVA', {'name': 'Máy biến áp khô 3 pha', 'desc': 'Biến áp cách ly', 'use': 'Cấp điện cho máy CNC', 'specs': 'Công suất 10 kVA, 380V/220V'}, '850432'),
]

# Bo sung 18/11 (3): cay 85.01, dai oc, van mot chieu
PIPELINE_CASES += [
    ('Dong co 3 pha 5.5kW', {'name': 'Động cơ điện xoay chiều 3 pha', 'desc': 'Vỏ gang', 'use': 'Truyền động băng tải', 'specs': '380V, 5.5 kW'}, '85015229'),
    ('Dong co 3 pha 0.37kW', {'name': 'Động cơ điện 3 pha', 'desc': 'Vỏ nhôm', 'use': 'Quay giỏ máy rửa', 'specs': '400V, 370 W'}, '85015119'),
    ('Dong co 1 pha 0.75kW', {'name': 'Động cơ điện 1 pha', 'desc': 'Vỏ nhôm', 'use': 'Máy bơm nước', 'specs': '220V, 750 W'}, '85014019'),
    ('Mo to buoc DC 82W', {'name': 'Mô tơ bước DC', 'desc': 'Động cơ một chiều', 'use': 'Máy CNC', 'specs': '24VDC, 82 W'}, '85013170'),
    ('Dong co thieu cong suat', {'name': 'Động cơ điện xoay chiều', 'desc': 'Vỏ nhôm', 'use': 'Truyền động máy', 'specs': '220V'}, '8501'),
    ('Dai oc M12', {'name': 'Đai ốc lục giác thép M12', 'desc': 'Thép mạ kẽm, có ren', 'use': 'Lắp ghép kết cấu', 'specs': 'M12'}, '73181610'),
    ('Mica mo ho -> hoi chat lieu', {'name': 'Tấm che an toàn chất liệu từ mica', 'desc': 'kích thước 220x230x2mm (Phụ tùng thay thế máy gắn cục chống thấm nước)', 'use': '', 'specs': ''}, 'INSUFFICIENT'),
    ('Nhua mica PMMA tam che', {'name': 'Tấm che an toàn bằng nhựa mica PMMA', 'desc': 'Gia công khoan lỗ gắn máy', 'use': 'Che chắn an toàn cho máy', 'specs': '220x230x2mm'}, '39269099'),
    ('Mica khoang chiu nhiet', {'name': 'Tấm che bằng mica khoáng tự nhiên chịu nhiệt', 'desc': 'Cách điện chịu nhiệt cao', 'use': 'Lót cách nhiệt lò', 'specs': '220x230x2mm'}, '68149000'),
    ('Phot mo ho -> hoi cot kim loai', {'name': 'Phớt chắn dầu trục', 'desc': 'kích thước 25x35x7 mm', 'use': 'Làm kín trục bơm', 'specs': ''}, 'INSUFFICIENT'),
    ('MCB thieu Ampe -> dung 6 so', {'name': 'Cầu dao tự động MCB 1 cực', 'desc': 'Thiết bị đóng ngắt hạ áp', 'use': 'Bảo vệ mạch điện', 'specs': '230V'}, '853620'),
]

# Ma do AI/walker de xuat SAI -> engine phai sua dung (thay Gatekeeper 1 va 4)
CORRECTION_CASES = [
    ('AI sai 8544.49.21 cho cap Cat6', '85444921', 'Cáp mạng Cat6 UTP 4 đôi vỏ PVC, chưa gắn đầu nối, không có lớp bọc chống nhiễu, điện áp tín hiệu 5V, mạng LAN viễn thông', '85444913'),
    ('AI sai 4010.12 cho bang tai loi thep', '40101200', 'Băng tải (đai tải) cao su gia cố bằng lõi thép dùng cho băng chuyền công nghiệp', '40101100'),
    ('AI sai 4010.11 cho bang tai vai det', '40101100', 'Băng tải (đai tải) cao su gia cố bằng vải dệt dùng cho băng chuyền công nghiệp', '40101200'),
    ('AI sai 8536.50.96 cho cong tac 25A', '85365096', 'Công tắc điện 380V 25A', '853650'),
    ('AI sai 8424.89.90 cho may rua dien', '84248990', 'Máy rửa công nghiệp hoạt động bằng điện 3x400V', '84248950'),
]


def run_correction_cases():
    ok, rows = 0, []
    for name, wrong, text, expected in CORRECTION_CASES:
        r = u.reroute(wrong, u.norm(text))
        got = (r['code'].replace('.', '') if r else 'NONE')
        good = got.startswith(expected)
        ok += good
        rows.append((name, good, f'got={got} expected={expected}', {}))
    return ok, rows


def run_engine_cases():
    ok = 0
    rows = []
    for name, query, keep, drop in ENGINE_CASES:
        qi = u.extract_query(query)
        res = []
        good = True
        for c in keep:
            dq, why, _ = u.evaluate(c, qi)
            if dq:
                good = False
                res.append(f'{c} BI LOAI SAI ({why})')
        for c in drop:
            dq, why, _ = u.evaluate(c, qi)
            if not dq:
                good = False
                res.append(f'{c} PHAI BI LOAI nhung con song')
        ok += good
        rows.append((name, good, '; '.join(res), qi['q']))
    return ok, rows


def run_pipeline_cases():
    import server
    server.get_system_gemini_api_key = lambda explicit_key=None: None  # test engine cuc bo, khong goi Gemini
    ok = 0
    rows = []
    for name, info, expected in PIPELINE_CASES:
        r = server.run_automation_workflow(info['name'], info['desc'], info['use'], info.get('specs', ''))
        got = 'INSUFFICIENT' if r.get('_insufficient_info') else (r.get('clean_code') or r.get('hs_code') or '').replace('.', '')
        good = got.startswith(expected)
        ok += good
        rows.append((name, good, f'got={got} expected={expected}', {}))
    return ok, rows


def export_cases(path='golden_cases.json'):
    data = {
        'engine': [{'name': n, 'query': q, 'keep': k, 'drop': d} for n, q, k, d in ENGINE_CASES],
        'pipeline': [{'name': n, 'info': i, 'expected': e} for n, i, e in PIPELINE_CASES],
        'correction': [{'name': n, 'wrong': w, 'text': t, 'expected': e} for n, w, t, e in CORRECTION_CASES],
    }
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    print('exported', len(data['engine']), 'engine +', len(data['pipeline']), 'pipeline cases')


if __name__ == '__main__':
    if '--export' in sys.argv:
        export_cases()
        sys.exit(0)
    total_ok = total = 0
    ok, rows = run_engine_cases()
    print('=== ENGINE ===')
    for n, g, d, q in rows:
        print(('PASS ' if g else 'FAIL ') + n, d, q)
    total_ok += ok; total += len(rows)
    ok, rows = run_correction_cases()
    print('=== CORRECTION ===')
    for n, g, d, q in rows:
        print(('PASS ' if g else 'FAIL ') + n, d)
    total_ok += ok; total += len(rows)
    if '--pipeline' in sys.argv:
        ok, rows = run_pipeline_cases()
        print('=== PIPELINE ===')
        for n, g, d, q in rows:
            print(('PASS ' if g else 'FAIL ') + n, d)
        total_ok += ok; total += len(rows)
    print(f'TONG: {total_ok}/{total}')
    if total_ok != total:
        sys.exit(1)
